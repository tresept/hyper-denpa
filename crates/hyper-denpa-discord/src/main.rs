mod timetable;

use anyhow::Context as _;
use chrono::{Local, NaiveDate, NaiveDateTime};
use knct_sharepoint::config::{DATA_DIR, ENV_FILE};
use knct_sharepoint::env::{load_dotenv, optional_env};
use knct_sharepoint::fs_utils::run_prefix;
use knct_sharepoint::models::OutputLayout;
use knct_sharepoint::pipeline::{FetchRequest, fetch_and_store};
use knct_sharepoint::sharepoint::resolve_default_timetable_target;
use log::{error, info, warn};
use serde::{Deserialize, Serialize};
use serenity::all::{
    ChannelId, Command, CommandInteraction, CommandOptionType, Context, CreateAttachment,
    CreateCommand, CreateCommandOption, CreateEmbed, CreateEmbedAuthor, CreateEmbedFooter,
    CreateInteractionResponse, CreateInteractionResponseMessage, CreateMessage,
    EditInteractionResponse, EventHandler, GatewayIntents, Interaction, Ready, ResolvedValue,
};
use sha2::{Digest, Sha256};
use std::cmp::Ordering as CmpOrdering;
use std::collections::{BTreeMap, BTreeSet};
use std::path::PathBuf;
use std::sync::{
    Arc,
    atomic::{AtomicBool, Ordering},
};
use timetable::{TimetableEntry, convert_xlsx_to_csvs, resolve_csv_path, resolve_show_result};
use tokio::sync::Mutex;
use tokio::time::{self, Duration};

const MAX_NOTIFY_CHANNELS: usize = 3;
const STATE_FILE: &str = "discord_state.json";
const POLL_INTERVAL_SECS: u64 = 60 * 60;
const MAX_VISIBLE_ENTRIES: usize = 30;
const MAX_MESSAGE_CONTENT_LEN: usize = 2_000;
const MAX_EMBED_DESCRIPTION_LEN: usize = 4_000;
const MAX_ATTACHMENT_BYTES: u64 = 8 * 1024 * 1024;

#[derive(Debug, Clone)]
struct Snapshot {
    date: String,
    fetched_at: String,
    entries: Vec<TimetableEntry>,
    csv_path: PathBuf,
    csv_hash: String,
}

#[derive(Debug, Default, Clone, Serialize, Deserialize)]
#[serde(default)]
struct BotState {
    notify_channels: Vec<u64>,
    last_notified_hash: Option<String>,
    last_periodic_error: Option<String>,
    last_snapshot: Option<StoredSnapshot>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct StoredSnapshot {
    date: String,
    csv_hash: String,
    entries: Vec<TimetableEntry>,
}

impl StoredSnapshot {
    fn from_snapshot(snapshot: &Snapshot) -> Self {
        Self {
            date: snapshot.date.clone(),
            csv_hash: snapshot.csv_hash.clone(),
            entries: visible_entries(&snapshot.entries),
        }
    }
}

#[derive(Debug, Default)]
struct EntryDiff {
    added: Vec<TimetableEntry>,
    removed: Vec<TimetableEntry>,
}

impl EntryDiff {
    fn is_empty(&self) -> bool {
        self.added.is_empty() && self.removed.is_empty()
    }
}

struct StateStore {
    path: PathBuf,
    inner: Mutex<BotState>,
}

impl StateStore {
    async fn load(path: PathBuf) -> anyhow::Result<Self> {
        let state = match tokio::fs::read_to_string(&path).await {
            Ok(body) => serde_json::from_str(&body)
                .with_context(|| format!("failed to parse {}", path.display()))?,
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => BotState::default(),
            Err(error) => {
                return Err(error).with_context(|| format!("failed to read {}", path.display()));
            }
        };

        info!("loaded bot state from {}", path.display());
        Ok(Self {
            path,
            inner: Mutex::new(state),
        })
    }

    async fn save_locked(&self, state: &BotState) -> anyhow::Result<()> {
        if let Some(parent) = self.path.parent() {
            tokio::fs::create_dir_all(parent)
                .await
                .with_context(|| format!("failed to create {}", parent.display()))?;
        }

        tokio::fs::write(&self.path, serde_json::to_vec_pretty(state)?)
            .await
            .with_context(|| format!("failed to write {}", self.path.display()))
    }

    async fn add_channel(&self, channel_id: u64) -> anyhow::Result<AddChannelResult> {
        let mut state = self.inner.lock().await;
        if state.notify_channels.contains(&channel_id) {
            info!(
                "channel {} is already registered for notifications",
                channel_id
            );
            return Ok(AddChannelResult::AlreadyRegistered);
        }
        if state.notify_channels.len() >= MAX_NOTIFY_CHANNELS {
            warn!(
                "notification channel limit reached while adding {}",
                channel_id
            );
            return Ok(AddChannelResult::LimitReached);
        }

        state.notify_channels.push(channel_id);
        self.save_locked(&state).await?;
        info!("registered channel {} for notifications", channel_id);
        Ok(AddChannelResult::Added {
            count: state.notify_channels.len(),
        })
    }

    async fn remove_channel(&self, channel_id: u64) -> anyhow::Result<bool> {
        let mut state = self.inner.lock().await;
        let before = state.notify_channels.len();
        state.notify_channels.retain(|id| *id != channel_id);
        let removed = before != state.notify_channels.len();
        if removed {
            self.save_locked(&state).await?;
            info!("unregistered channel {} from notifications", channel_id);
        }
        Ok(removed)
    }

    async fn snapshot(&self) -> BotState {
        self.inner.lock().await.clone()
    }

    async fn remember_snapshot(&self, snapshot: StoredSnapshot) -> anyhow::Result<()> {
        let mut state = self.inner.lock().await;
        state.last_notified_hash = Some(snapshot.csv_hash.clone());
        state.last_periodic_error = None;
        state.last_snapshot = Some(snapshot);
        self.save_locked(&state).await
    }

    async fn clear_error(&self) -> anyhow::Result<()> {
        let mut state = self.inner.lock().await;
        if state.last_periodic_error.is_none() {
            return Ok(());
        }
        state.last_periodic_error = None;
        self.save_locked(&state).await
    }

    async fn mark_periodic_error(&self, error_message: String) -> anyhow::Result<(bool, Vec<u64>)> {
        let mut state = self.inner.lock().await;
        let changed = state.last_periodic_error.as_deref() != Some(error_message.as_str());
        state.last_periodic_error = Some(error_message);
        if changed {
            self.save_locked(&state).await?;
        }
        Ok((changed, state.notify_channels.clone()))
    }
}

enum AddChannelResult {
    Added { count: usize },
    AlreadyRegistered,
    LimitReached,
}

struct Handler {
    state: Arc<StateStore>,
    scheduler_started: AtomicBool,
}

#[serenity::async_trait]
impl EventHandler for Handler {
    async fn ready(&self, ctx: Context, ready: Ready) {
        if let Err(error) = register_commands(&ctx).await {
            error!("failed to register commands: {error:#}");
        }

        if !self.scheduler_started.swap(true, Ordering::SeqCst) {
            let http = ctx.http.clone();
            let state = self.state.clone();
            tokio::spawn(async move {
                if let Err(error) = run_scheduler(http, state).await {
                    error!("scheduler stopped: {error:#}");
                }
            });
        }

        info!("{} is connected", ready.user.name);
    }

    async fn interaction_create(&self, ctx: Context, interaction: Interaction) {
        let Interaction::Command(command) = interaction else {
            return;
        };

        let result = match command.data.name.as_str() {
            "reload" => handle_reload(&ctx, &command).await,
            "show" => handle_show(&ctx, &command).await,
            "grep" => handle_grep(&ctx, &command).await,
            "law-csv" => handle_law_csv(&ctx, &command).await,
            "set-notify" => handle_set_notify(&ctx, &command, &self.state).await,
            "unset-notify" => handle_unset_notify(&ctx, &command, &self.state).await,
            other => respond_message(&ctx, &command, &format!("unknown command: {other}")).await,
        };

        if let Err(error) = result {
            error!("interaction failed: {error:#}");
            let _ = respond_message(&ctx, &command, &format!("エラー: {error:#}")).await;
        }
    }
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let env_values =
        load_dotenv().with_context(|| format!("bot 用の {} を読み取れませんでした", ENV_FILE))?;
    let log_filter =
        optional_env(&env_values, "RUST_LOG").unwrap_or_else(|| "info,serenity=warn".to_string());

    let mut logger = env_logger::Builder::new();
    logger.parse_filters(&log_filter);
    logger.format_timestamp_secs();
    logger.init();

    let token = optional_env(&env_values, "DISCORD_TOKEN")
        .context("DISCORD_TOKEN が .env に設定されていません")?;
    info!(
        "loaded discord configuration from {} with RUST_LOG={}",
        ENV_FILE, log_filter
    );

    let state_path = PathBuf::from(DATA_DIR).join(STATE_FILE);
    let state = Arc::new(StateStore::load(state_path).await?);
    let intents = GatewayIntents::GUILDS;

    let mut client = serenity::Client::builder(token, intents)
        .event_handler(Handler {
            state,
            scheduler_started: AtomicBool::new(false),
        })
        .await
        .context("failed to build discord client")?;

    client.start().await.context("discord client exited")
}

async fn register_commands(ctx: &Context) -> anyhow::Result<()> {
    let commands = vec![
        CreateCommand::new("reload").description("時間割変更を再取得します"),
        CreateCommand::new("show").description("最後に取得した時間割変更を日付順に表示します"),
        CreateCommand::new("grep")
            .description("指定した学年・学科クラスの時間割変更だけを表示します")
            .add_option(
                CreateCommandOption::new(CommandOptionType::Integer, "grade", "学年")
                    .add_int_choice("1", 1)
                    .add_int_choice("2", 2)
                    .add_int_choice("3", 3)
                    .add_int_choice("4", 4)
                    .add_int_choice("5", 5)
                    .required(true),
            )
            .add_option(
                CreateCommandOption::new(
                    CommandOptionType::String,
                    "class_name",
                    "学科・クラス。省略時は学年全体を表示",
                )
                .add_string_choice("CN", "CN")
                .add_string_choice("ES", "ES")
                .add_string_choice("IT", "IT")
                .add_string_choice("1組", "1")
                .add_string_choice("2組", "2")
                .add_string_choice("3組", "3")
                .required(false),
            ),
        CreateCommand::new("law-csv").description("最新の生 CSV を取得して添付します"),
        CreateCommand::new("set-notify").description("このチャンネルを定期通知先に追加します"),
        CreateCommand::new("unset-notify")
            .description("定期通知先からチャンネルを外します")
            .add_option(
                CreateCommandOption::new(
                    CommandOptionType::Channel,
                    "channel",
                    "解除するチャンネル。省略時はこのチャンネル",
                )
                .required(false),
            ),
    ];

    Command::set_global_commands(&ctx.http, commands)
        .await
        .context("failed to register global commands")?;
    Ok(())
}

async fn respond_message(
    ctx: &Context,
    command: &CommandInteraction,
    content: &str,
) -> anyhow::Result<()> {
    command
        .create_response(
            &ctx.http,
            CreateInteractionResponse::Message(
                CreateInteractionResponseMessage::new().content(content),
            ),
        )
        .await
        .context("failed to send interaction response")
}

async fn defer_response(ctx: &Context, command: &CommandInteraction) -> anyhow::Result<()> {
    command
        .create_response(
            &ctx.http,
            CreateInteractionResponse::Defer(CreateInteractionResponseMessage::new()),
        )
        .await
        .context("failed to defer interaction response")
}

async fn handle_reload(ctx: &Context, command: &CommandInteraction) -> anyhow::Result<()> {
    info!("received /reload from channel {}", command.channel_id);
    defer_response(ctx, command).await?;

    let snapshot = fetch_latest_snapshot().await?;
    let content = format!("再取得しました．最終取得 ：{}", snapshot.fetched_at);

    command
        .edit_response(&ctx.http, EditInteractionResponse::new().content(content))
        .await
        .context("failed to edit reload response")?;
    Ok(())
}

async fn handle_show(ctx: &Context, command: &CommandInteraction) -> anyhow::Result<()> {
    info!("received /show from channel {}", command.channel_id);
    defer_response(ctx, command).await?;

    let snapshot = match load_saved_snapshot(None).await {
        Ok(snapshot) => snapshot,
        Err(error) => {
            return command
                .edit_response(
                    &ctx.http,
                    EditInteractionResponse::new().content(no_saved_data_message(&error)),
                )
                .await
                .context("failed to edit show response")
                .map(|_| ());
        }
    };

    if sorted_visible_entries(&snapshot.entries).is_empty() {
        return command
            .edit_response(
                &ctx.http,
                EditInteractionResponse::new().content(no_timetable_changes_message()),
            )
            .await
            .context("failed to edit show empty response")
            .map(|_| ());
    }

    let response = create_manual_check_response(&snapshot);

    let builder = match response {
        CheckResponse::Embed(embed) => EditInteractionResponse::new().content("").embed(*embed),
        CheckResponse::Fallback(message) => EditInteractionResponse::new().content(message),
    };

    command
        .edit_response(&ctx.http, builder)
        .await
        .context("failed to edit show response")?;
    Ok(())
}

async fn handle_grep(ctx: &Context, command: &CommandInteraction) -> anyhow::Result<()> {
    info!("received /grep from channel {}", command.channel_id);
    defer_response(ctx, command).await?;

    let mut grade: Option<String> = None;
    let mut class_name: Option<String> = None;
    for option in command.data.options() {
        match option.name {
            "grade" => {
                if let ResolvedValue::Integer(value) = option.value {
                    grade = Some(value.to_string());
                }
            }
            "class_name" => {
                if let ResolvedValue::String(value) = option.value {
                    class_name = Some(value.to_string());
                }
            }
            _ => {}
        }
    }

    let Some(grade) = grade else {
        return command
            .edit_response(
                &ctx.http,
                EditInteractionResponse::new().content("学年を指定してください．"),
            )
            .await
            .context("failed to edit grep validation response")
            .map(|_| ());
    };

    let snapshot = match load_saved_snapshot(None).await {
        Ok(snapshot) => snapshot,
        Err(error) => {
            return command
                .edit_response(
                    &ctx.http,
                    EditInteractionResponse::new().content(no_saved_data_message(&error)),
                )
                .await
                .context("failed to edit grep response")
                .map(|_| ());
        }
    };

    let filtered_entries = filter_entries(&snapshot.entries, &grade, class_name.as_deref());
    if filtered_entries.is_empty() {
        return command
            .edit_response(
                &ctx.http,
                EditInteractionResponse::new().content(no_timetable_changes_message()),
            )
            .await
            .context("failed to edit grep empty response")
            .map(|_| ());
    }

    let mut filtered_snapshot = snapshot.clone();
    filtered_snapshot.entries = filtered_entries;
    let title = match class_name.as_deref() {
        Some(class_name) => format!(
            "{}年 {} の時間割変更一覧",
            normalize_grade_for_display(&grade),
            class_name.trim()
        ),
        None => format!("{}年の時間割変更一覧", normalize_grade_for_display(&grade)),
    };
    let response = create_custom_embed_response(&filtered_snapshot, &title, 1_035_687);

    let builder = match response {
        CheckResponse::Embed(embed) => EditInteractionResponse::new().content("").embed(*embed),
        CheckResponse::Fallback(message) => EditInteractionResponse::new().content(message),
    };

    command
        .edit_response(&ctx.http, builder)
        .await
        .context("failed to edit grep response")?;
    Ok(())
}

async fn handle_law_csv(ctx: &Context, command: &CommandInteraction) -> anyhow::Result<()> {
    info!("received /law-csv from channel {}", command.channel_id);
    defer_response(ctx, command).await?;

    let snapshot = fetch_latest_snapshot().await?;
    let csv_body = tokio::fs::read_to_string(&snapshot.csv_path)
        .await
        .with_context(|| format!("failed to read {}", snapshot.csv_path.display()))?;
    let response = build_law_csv_response(&snapshot, &csv_body).await?;

    let builder = match response {
        CsvResponse::Message(message) => EditInteractionResponse::new().content(message),
        CsvResponse::Attachment(message, attachment) => EditInteractionResponse::new()
            .content(message)
            .new_attachment(attachment),
    };

    command
        .edit_response(&ctx.http, builder)
        .await
        .context("failed to edit law-csv response")?;
    Ok(())
}

async fn handle_set_notify(
    ctx: &Context,
    command: &CommandInteraction,
    state: &Arc<StateStore>,
) -> anyhow::Result<()> {
    let channel_id = command.channel_id.get();
    info!("received /set-notify for channel {}", channel_id);
    let message = match state.add_channel(channel_id).await? {
        AddChannelResult::Added { count } => {
            format!(
                "このチャンネルを定期通知先に追加しました。現在 {count}/{MAX_NOTIFY_CHANNELS} 件です。"
            )
        }
        AddChannelResult::AlreadyRegistered => "このチャンネルはすでに通知先です。".to_string(),
        AddChannelResult::LimitReached => {
            format!("定期通知に設定できるチャンネルは最大 {MAX_NOTIFY_CHANNELS} 件です。")
        }
    };

    respond_message(ctx, command, &message).await
}

async fn handle_unset_notify(
    ctx: &Context,
    command: &CommandInteraction,
    state: &Arc<StateStore>,
) -> anyhow::Result<()> {
    info!("received /unset-notify from channel {}", command.channel_id);
    let channel_id = command
        .data
        .options()
        .into_iter()
        .find_map(|option| match option.value {
            ResolvedValue::Channel(channel) => Some(channel.id),
            _ => None,
        })
        .unwrap_or(command.channel_id)
        .get();

    let message = if state.remove_channel(channel_id).await? {
        "指定チャンネルの定期通知を解除しました。".to_string()
    } else {
        "指定チャンネルは通知先に登録されていません。".to_string()
    };

    respond_message(ctx, command, &message).await
}

async fn run_scheduler(
    http: Arc<serenity::http::Http>,
    state: Arc<StateStore>,
) -> anyhow::Result<()> {
    let mut interval = time::interval(Duration::from_secs(POLL_INTERVAL_SECS));

    loop {
        interval.tick().await;

        if let Err(error) = run_periodic_check(&http, &state).await {
            let error_message = format!("{error:#}");
            error!("periodic check failed: {error_message}");
            let (should_notify, channels) =
                state.mark_periodic_error(error_message.clone()).await?;
            if should_notify {
                broadcast_text(
                    &http,
                    &channels,
                    &format!("定期チェックでエラーが発生しました。\n```text\n{error_message}\n```"),
                )
                .await;
            }
        }
    }
}

async fn run_periodic_check(
    http: &Arc<serenity::http::Http>,
    state: &Arc<StateStore>,
) -> anyhow::Result<()> {
    info!("running periodic timetable check");
    let snapshot = fetch_latest_snapshot().await?;
    let state_snapshot = state.snapshot().await;

    let current_snapshot = StoredSnapshot::from_snapshot(&snapshot);
    let Some(previous_snapshot) = state_snapshot.last_snapshot.clone() else {
        state.remember_snapshot(current_snapshot).await?;
        return Ok(());
    };

    if previous_snapshot.csv_hash == snapshot.csv_hash {
        state.clear_error().await?;
        return Ok(());
    }

    let diff = diff_entries(&previous_snapshot.entries, &current_snapshot.entries);
    if diff.is_empty() {
        info!("csv changed, but visible timetable entries did not change");
        state.remember_snapshot(current_snapshot).await?;
        return Ok(());
    }

    if state_snapshot.notify_channels.is_empty() {
        state.remember_snapshot(current_snapshot).await?;
        return Ok(());
    }

    match create_periodic_update_response(&snapshot, &diff) {
        CheckResponse::Embed(embed) => {
            info!(
                "broadcasting periodic diff embed to {} channels",
                state_snapshot.notify_channels.len()
            );
            broadcast_embed(http, &state_snapshot.notify_channels, *embed).await;
        }
        CheckResponse::Fallback(message) => {
            warn!("periodic diff embed exceeded limit, sending fallback message");
            broadcast_text(http, &state_snapshot.notify_channels, &message).await;
        }
    }

    state.remember_snapshot(current_snapshot).await?;
    Ok(())
}

async fn broadcast_text(http: &Arc<serenity::http::Http>, channels: &[u64], content: &str) {
    for channel_id in channels {
        if let Err(error) = ChannelId::new(*channel_id)
            .send_message(http, CreateMessage::new().content(content))
            .await
        {
            error!("failed to send message to channel {}: {error}", channel_id);
        }
    }
}

async fn broadcast_embed(http: &Arc<serenity::http::Http>, channels: &[u64], embed: CreateEmbed) {
    for channel_id in channels {
        if let Err(error) = ChannelId::new(*channel_id)
            .send_message(http, CreateMessage::new().embed(embed.clone()))
            .await
        {
            error!("failed to send embed to channel {}: {error}", channel_id);
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
struct EntryKey {
    date: String,
    weekday: String,
    period: String,
    grade: String,
    class_name: String,
    change_type: String,
    subject: String,
}

fn diff_entries(previous: &[TimetableEntry], current: &[TimetableEntry]) -> EntryDiff {
    let previous_map = group_entries(previous);
    let current_map = group_entries(current);
    let mut keys = BTreeSet::new();
    keys.extend(previous_map.keys().cloned());
    keys.extend(current_map.keys().cloned());

    let mut added = Vec::new();
    let mut removed = Vec::new();

    for key in keys {
        let previous_entries = previous_map.get(&key).map(Vec::as_slice).unwrap_or(&[]);
        let current_entries = current_map.get(&key).map(Vec::as_slice).unwrap_or(&[]);

        match current_entries.len().cmp(&previous_entries.len()) {
            CmpOrdering::Greater => {
                added.extend(current_entries.iter().skip(previous_entries.len()).cloned())
            }
            CmpOrdering::Less => {
                removed.extend(previous_entries.iter().skip(current_entries.len()).cloned())
            }
            CmpOrdering::Equal => {}
        }
    }

    sort_entries(&mut added);
    sort_entries(&mut removed);
    EntryDiff { added, removed }
}

fn group_entries(entries: &[TimetableEntry]) -> BTreeMap<EntryKey, Vec<TimetableEntry>> {
    let mut grouped = BTreeMap::new();
    for entry in entries {
        grouped
            .entry(entry_key(entry))
            .or_insert_with(Vec::new)
            .push(entry.clone());
    }
    grouped
}

fn entry_key(entry: &TimetableEntry) -> EntryKey {
    EntryKey {
        date: normalize_entry_value(&entry.date),
        weekday: normalize_entry_value(&entry.weekday),
        period: normalize_entry_value(&entry.period),
        grade: normalize_entry_value(&entry.grade),
        class_name: normalize_entry_value(&entry.class_name),
        change_type: normalize_entry_value(&entry.change_type),
        subject: normalize_entry_value(&entry.subject),
    }
}

fn normalize_entry_value(value: &str) -> String {
    value.split_whitespace().collect::<Vec<_>>().join(" ")
}

fn compare_entries(left: &TimetableEntry, right: &TimetableEntry) -> CmpOrdering {
    parse_entry_date(&left.date)
        .cmp(&parse_entry_date(&right.date))
        .then(first_period_number(&left.period).cmp(&first_period_number(&right.period)))
        .then(left.grade.cmp(&right.grade))
        .then(left.class_name.cmp(&right.class_name))
        .then(left.change_type.cmp(&right.change_type))
        .then(left.subject.cmp(&right.subject))
}

fn sort_entries(entries: &mut [TimetableEntry]) {
    entries.sort_by(compare_entries);
}

fn visible_entries(entries: &[TimetableEntry]) -> Vec<TimetableEntry> {
    let today = Local::now().date_naive();
    let mut visible = entries
        .iter()
        .filter_map(|entry| {
            let date = parse_entry_date(&entry.date)?;
            if date < today {
                return None;
            }

            Some(entry.clone())
        })
        .collect::<Vec<_>>();

    sort_entries(&mut visible);
    visible
}

async fn fetch_latest_snapshot() -> anyhow::Result<Snapshot> {
    let date = run_prefix();
    info!("fetching latest snapshot for {}", date);
    let target = resolve_default_timetable_target()?;
    let report = fetch_and_store(FetchRequest {
        target,
        local_xlsx_path: None,
        date: Some(date.clone()),
    })
    .await?;
    let layout = OutputLayout::new(&date);
    convert_xlsx_to_csvs(
        std::path::Path::new(&report.file_path),
        &layout.csv_dir,
        &date,
    )?;
    let show_result = resolve_show_result(Some(&date))?;
    let (_, csv_path) = resolve_csv_path(Some(&date))?;
    let csv_body = tokio::fs::read(&csv_path)
        .await
        .with_context(|| format!("failed to read {}", csv_path.display()))?;

    Ok(Snapshot {
        fetched_at: format_run_key(&date),
        date,
        entries: show_result.entries,
        csv_path,
        csv_hash: format!("{:x}", Sha256::digest(&csv_body)),
    })
}

async fn load_saved_snapshot(date: Option<&str>) -> anyhow::Result<Snapshot> {
    let show_result = resolve_show_result(date)?;
    let (run_key, csv_path) = resolve_csv_path(date)?;
    let csv_body = tokio::fs::read(&csv_path)
        .await
        .with_context(|| format!("failed to read {}", csv_path.display()))?;

    Ok(Snapshot {
        date: run_key.clone(),
        fetched_at: format_run_key(&run_key),
        entries: show_result.entries,
        csv_path,
        csv_hash: format!("{:x}", Sha256::digest(&csv_body)),
    })
}

enum CheckResponse {
    Embed(Box<CreateEmbed>),
    Fallback(String),
}

enum CsvResponse {
    Message(String),
    Attachment(String, CreateAttachment),
}

fn create_manual_check_response(snapshot: &Snapshot) -> CheckResponse {
    create_embed_response(snapshot, "現在発表中の時間割変更一覧", 1_035_687)
}

fn create_custom_embed_response(snapshot: &Snapshot, title: &str, color: u32) -> CheckResponse {
    create_embed_response(snapshot, title, color)
}

fn create_periodic_update_response(snapshot: &Snapshot, diff: &EntryDiff) -> CheckResponse {
    create_diff_embed_response(
        snapshot,
        diff,
        "更新されました：時間割変更の差分",
        16_711_680,
    )
}

fn create_embed_response(snapshot: &Snapshot, title: &str, color: u32) -> CheckResponse {
    let description = format_manual_check_description(snapshot);
    if description.chars().count() > MAX_EMBED_DESCRIPTION_LEN {
        return CheckResponse::Fallback(limit_exceeded_message());
    }

    CheckResponse::Embed(Box::new(
        CreateEmbed::new()
            .author(CreateEmbedAuthor::new("時間割変更通知サービス"))
            .title(title)
            .description(description)
            .footer(CreateEmbedFooter::new(format!(
                "最終取得 ：{}",
                snapshot.fetched_at
            )))
            .field(
                "免責事項",
                "この情報には誤差がある可能性があります．正確な変更については，学内掲示またはデンパポータルをご覧ください．損害に対しての責任は負いかねます．",
                false,
            )
            .color(color),
    ))
}

fn create_diff_embed_response(
    snapshot: &Snapshot,
    diff: &EntryDiff,
    title: &str,
    color: u32,
) -> CheckResponse {
    let description = format_periodic_diff_description(diff);
    if description.chars().count() > MAX_EMBED_DESCRIPTION_LEN {
        return CheckResponse::Fallback(limit_exceeded_message());
    }

    CheckResponse::Embed(Box::new(
        CreateEmbed::new()
            .author(CreateEmbedAuthor::new("時間割変更通知サービス"))
            .title(title)
            .description(description)
            .footer(CreateEmbedFooter::new(format!(
                "最終取得 ：{}",
                snapshot.fetched_at
            )))
            .field(
                "免責事項",
                "この情報には誤差がある可能性があります．正確な変更については，学内掲示またはデンパポータルをご覧ください．損害に対しての責任は負いかねます．",
                false,
            )
            .color(color),
    ))
}

fn format_manual_check_description(snapshot: &Snapshot) -> String {
    let entries = sorted_visible_entries(&snapshot.entries);
    if entries.is_empty() {
        return "現在発表中の変更はありません。".to_string();
    }

    format_entry_sections(&entries)
}

fn format_periodic_diff_description(diff: &EntryDiff) -> String {
    let mut sections = Vec::new();

    if !diff.added.is_empty() {
        sections.push(render_diff_group("追加", &diff.added));
    }
    if !diff.removed.is_empty() {
        sections.push(render_diff_group("削除", &diff.removed));
    }

    sections.join("\n\n")
}

fn render_diff_group(title: &str, entries: &[TimetableEntry]) -> String {
    let visible_entries = entries
        .iter()
        .take(MAX_VISIBLE_ENTRIES)
        .cloned()
        .collect::<Vec<_>>();
    let mut section = format!("## {title}\n{}", format_entry_sections(&visible_entries));
    let hidden_count = entries.len().saturating_sub(visible_entries.len());
    if hidden_count > 0 {
        section.push_str(&format!("\n…ほか {hidden_count} 件"));
    }
    section
}

fn format_entry_sections(entries: &[TimetableEntry]) -> String {
    let mut sections = Vec::new();
    let mut current_key: Option<(&str, &str, &str)> = None;
    let mut current_lines: Vec<String> = Vec::new();

    for entry in entries {
        let key = (
            entry.date.as_str(),
            entry.weekday.as_str(),
            entry.period.as_str(),
        );

        if current_key != Some(key) {
            if let Some((date, weekday, period)) = current_key {
                sections.push(render_manual_check_section(
                    date,
                    weekday,
                    period,
                    &current_lines,
                ));
                current_lines.clear();
            }
            current_key = Some(key);
        }

        current_lines.push(format!(
            "{}年 {}: **{}** - {}",
            entry.grade, entry.class_name, entry.change_type, entry.subject
        ));
    }

    if let Some((date, weekday, period)) = current_key {
        sections.push(render_manual_check_section(
            date,
            weekday,
            period,
            &current_lines,
        ));
    }

    sections.join("\n")
}

fn render_manual_check_section(
    date: &str,
    weekday: &str,
    period: &str,
    lines: &[String],
) -> String {
    let mut section = format!("### {} {}曜日 {}限", date, weekday, period);
    for line in lines {
        section.push('\n');
        section.push_str(line);
    }
    section.push('\n');
    section
}

fn sorted_visible_entries(entries: &[TimetableEntry]) -> Vec<TimetableEntry> {
    visible_entries(entries)
        .into_iter()
        .take(MAX_VISIBLE_ENTRIES)
        .collect()
}

fn parse_entry_date(raw: &str) -> Option<NaiveDate> {
    let normalized = raw.trim().replace('-', "/");
    let mut parts = normalized.split('/');
    let year = parts.next()?.parse::<i32>().ok()?;
    let month = parts.next()?.parse::<u32>().ok()?;
    let day = parts.next()?.parse::<u32>().ok()?;
    if parts.next().is_some() {
        return None;
    }

    NaiveDate::from_ymd_opt(year, month, day)
}

fn first_period_number(raw: &str) -> u32 {
    raw.split(',')
        .next()
        .map(str::trim)
        .and_then(|value| value.parse::<u32>().ok())
        .unwrap_or(u32::MAX)
}

async fn build_law_csv_response(
    snapshot: &Snapshot,
    csv_body: &str,
) -> anyhow::Result<CsvResponse> {
    let trimmed = csv_body.trim_start_matches('\u{feff}');
    let code_block = format!("```csv\n{trimmed}\n```");
    if code_block.chars().count() <= MAX_MESSAGE_CONTENT_LEN {
        info!("sending /law-csv as code block");
        return Ok(CsvResponse::Message(code_block));
    }

    let metadata = tokio::fs::metadata(&snapshot.csv_path)
        .await
        .with_context(|| {
            format!(
                "failed to read metadata for {}",
                snapshot.csv_path.display()
            )
        })?;
    if metadata.len() > MAX_ATTACHMENT_BYTES {
        warn!(
            "/law-csv attachment exceeds byte limit: {} bytes",
            metadata.len()
        );
        return Ok(CsvResponse::Message(limit_exceeded_message()));
    }

    let attachment = CreateAttachment::path(&snapshot.csv_path)
        .await
        .with_context(|| format!("failed to attach {}", snapshot.csv_path.display()))?;
    info!("sending /law-csv as attachment fallback");
    Ok(CsvResponse::Attachment(
        format!(
            "{} の生 CSV は文字数制限のためファイルで送信します。\n最終取得 ：{}",
            snapshot.date, snapshot.fetched_at
        ),
        attachment,
    ))
}

fn limit_exceeded_message() -> String {
    match resolve_default_timetable_target()
        .ok()
        .and_then(|target| target.document_url)
    {
        Some(url) => {
            format!("文字数制限の超過のため送信できません．デンパポータルを参照してください：{url}")
        }
        None => {
            "文字数制限の超過のため送信できません．デンパポータルを参照してください。".to_string()
        }
    }
}

fn format_run_key(run_key: &str) -> String {
    NaiveDateTime::parse_from_str(run_key, "%Y-%m-%d_%H-%M-%S")
        .map(|datetime| datetime.format("%Y-%m-%d %H:%M:%S").to_string())
        .or_else(|_| {
            NaiveDate::parse_from_str(run_key, "%Y-%m-%d")
                .map(|date| date.format("%Y-%m-%d 00:00:00").to_string())
        })
        .unwrap_or_else(|_| run_key.to_string())
}

fn no_saved_data_message(error: &anyhow::Error) -> String {
    let body = format!("{error:#}");
    if body.contains("data ディレクトリがまだありません")
        || body.contains("data 配下に日付ディレクトリがありません")
    {
        "まだ取得済みデータがありません．先に /reload を実行してください．".to_string()
    } else {
        format!("エラー: {body}")
    }
}

fn no_timetable_changes_message() -> &'static str {
    "時間割変更は現段階でありません．"
}

fn filter_entries(
    entries: &[TimetableEntry],
    grade: &str,
    class_name: Option<&str>,
) -> Vec<TimetableEntry> {
    let normalized_grade = normalize_grade(grade);
    let normalized_class = class_name.map(normalize_class_filter);

    visible_entries(entries)
        .into_iter()
        .filter(|entry| {
            normalize_grade(&entry.grade) == normalized_grade
                && normalized_class
                    .as_ref()
                    .is_none_or(|class_name| class_name_matches(&entry.class_name, class_name))
        })
        .collect()
}

fn normalize_grade_for_display(grade: &str) -> String {
    normalize_grade(grade)
}

fn normalize_grade(grade: &str) -> String {
    grade.trim().trim_end_matches('年').trim().to_string()
}

fn normalize_class_name(class_name: &str) -> String {
    normalize_entry_value(class_name).to_ascii_uppercase()
}

fn normalize_class_filter(class_name: &str) -> String {
    normalize_class_name(class_name)
        .trim_end_matches('組')
        .to_string()
}

fn class_name_matches(entry_class_name: &str, selected_class_name: &str) -> bool {
    let normalized_entry = normalize_class_name(entry_class_name);
    let normalized_selected = normalize_class_filter(selected_class_name);

    normalized_entry == normalized_selected
        || normalized_entry == format!("{normalized_selected}組")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn entry(
        grade: &str,
        class_name: &str,
        date: &str,
        weekday: &str,
        period: &str,
        change_type: &str,
        subject: &str,
    ) -> TimetableEntry {
        TimetableEntry {
            grade: grade.to_string(),
            class_name: class_name.to_string(),
            date: date.to_string(),
            weekday: weekday.to_string(),
            period: period.to_string(),
            change_type: change_type.to_string(),
            subject: subject.to_string(),
        }
    }

    #[test]
    fn bot_state_loads_legacy_json_without_snapshot() {
        let body = r#"{
          "notify_channels": [123456789],
          "last_notified_hash": "abc123",
          "last_periodic_error": null
        }"#;

        let state: BotState = serde_json::from_str(body).expect("legacy state should deserialize");
        assert_eq!(state.notify_channels, vec![123456789]);
        assert_eq!(state.last_notified_hash.as_deref(), Some("abc123"));
        assert!(state.last_snapshot.is_none());
    }

    #[test]
    fn diff_entries_ignores_whitespace_noise() {
        let previous = vec![entry("2", "IT", "2026/4/8", "水", "3", "補講", "英語ⅡB")];
        let current = vec![entry(
            " 2 ",
            "IT",
            "2026/4/8",
            "水",
            "3",
            "補講",
            "英語ⅡB  ",
        )];

        let diff = diff_entries(&previous, &current);
        assert!(diff.is_empty());
    }

    #[test]
    fn diff_entries_reports_added_and_removed_rows() {
        let previous = vec![
            entry("2", "IT", "2026/4/8", "水", "3", "補講", "英語ⅡB"),
            entry("3", "ME", "2026/4/9", "木", "1", "休講", "数学B"),
        ];
        let current = vec![
            entry("2", "IT", "2026/4/8", "水", "3", "補講", "英語ⅡB"),
            entry("1", "CN", "2026/4/10", "金", "2", "教室変更", "化学"),
        ];

        let diff = diff_entries(&previous, &current);
        assert_eq!(
            diff.added,
            vec![entry("1", "CN", "2026/4/10", "金", "2", "教室変更", "化学")]
        );
        assert_eq!(
            diff.removed,
            vec![entry("3", "ME", "2026/4/9", "木", "1", "休講", "数学B")]
        );
    }

    #[test]
    fn filter_entries_matches_grade_only_and_class_name() {
        let entries = vec![
            entry("2", "IT", "2026/4/8", "水", "3", "補講", "英語ⅡB"),
            entry("2", "CN", "2026/4/8", "水", "4", "休講", "数学"),
            entry("3", "IT", "2026/4/8", "水", "2", "補講", "物理"),
        ];

        assert_eq!(filter_entries(&entries, "2年", None).len(), 2);
        assert_eq!(
            filter_entries(&entries, "2", Some("it")),
            vec![entry("2", "IT", "2026/4/8", "水", "3", "補講", "英語ⅡB")]
        );
    }

    #[test]
    fn filter_entries_matches_numbered_classes_with_or_without_kumi() {
        let entries = vec![
            entry("1", "1", "2026/4/8", "水", "1", "補講", "国語"),
            entry("1", "1組", "2026/4/8", "水", "2", "休講", "数学"),
            entry("1", "2組", "2026/4/8", "水", "3", "教室変更", "英語"),
        ];

        assert_eq!(filter_entries(&entries, "1", Some("1")).len(), 2);
        assert_eq!(filter_entries(&entries, "1", Some("2")).len(), 1);
    }

    #[test]
    fn format_run_key_supports_datetime_and_date() {
        assert_eq!(format_run_key("2026-04-07_13-45-00"), "2026-04-07 13:45:00");
        assert_eq!(format_run_key("2026-04-07"), "2026-04-07 00:00:00");
    }
}
