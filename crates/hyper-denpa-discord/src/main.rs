use anyhow::Context as _;
use chrono::{Local, NaiveDate};
use hyper_denpa_core::config::{DATA_DIR, ENV_FILE};
use hyper_denpa_core::env::{load_dotenv, optional_env};
use hyper_denpa_core::fs_utils::date_prefix;
use hyper_denpa_core::pipeline::{FetchRequest, fetch_and_store};
use hyper_denpa_core::sharepoint::resolve_default_timetable_target;
use hyper_denpa_core::show::{resolve_csv_path, resolve_show_result};
use log::{error, info, warn};
use serde::{Deserialize, Serialize};
use serenity::all::{
    ChannelId, Command, CommandInteraction, CommandOptionType, Context, CreateAttachment,
    CreateCommand, CreateCommandOption, CreateEmbed, CreateEmbedAuthor, CreateEmbedFooter,
    CreateInteractionResponse, CreateInteractionResponseMessage, CreateMessage,
    EditInteractionResponse, EventHandler, GatewayIntents, Interaction, Ready, ResolvedValue,
};
use sha2::{Digest, Sha256};
use std::path::PathBuf;
use std::sync::{
    Arc,
    atomic::{AtomicBool, Ordering},
};
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
    entries: Vec<hyper_denpa_core::models::TimetableEntry>,
    csv_path: PathBuf,
    csv_hash: String,
}

#[derive(Debug, Default, Clone, Serialize, Deserialize)]
struct BotState {
    notify_channels: Vec<u64>,
    last_notified_hash: Option<String>,
    last_periodic_error: Option<String>,
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

    async fn remember_notified_hash(&self, hash: String) -> anyhow::Result<()> {
        let mut state = self.inner.lock().await;
        state.last_notified_hash = Some(hash);
        state.last_periodic_error = None;
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
            "check" => handle_check(&ctx, &command).await,
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
        CreateCommand::new("check").description("最新の時間割変更を強制再取得して表示します"),
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

async fn handle_check(ctx: &Context, command: &CommandInteraction) -> anyhow::Result<()> {
    info!("received /check from channel {}", command.channel_id);
    defer_response(ctx, command).await?;

    let snapshot = fetch_latest_snapshot().await?;
    let response = create_manual_check_response(&snapshot);

    let builder = match response {
        CheckResponse::Embed(embed) => EditInteractionResponse::new().content("").embed(*embed),
        CheckResponse::Fallback(message) => EditInteractionResponse::new().content(message),
    };

    command
        .edit_response(&ctx.http, builder)
        .await
        .context("failed to edit check response")?;
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

    match state_snapshot.last_notified_hash.as_deref() {
        None => {
            state.remember_notified_hash(snapshot.csv_hash).await?;
        }
        Some(previous_hash) if previous_hash == snapshot.csv_hash => {
            state.clear_error().await?;
        }
        Some(_) => {
            if state_snapshot.notify_channels.is_empty() {
                state.remember_notified_hash(snapshot.csv_hash).await?;
                return Ok(());
            }

            match create_periodic_update_response(&snapshot) {
                CheckResponse::Embed(embed) => {
                    info!(
                        "broadcasting periodic update embed to {} channels",
                        state_snapshot.notify_channels.len()
                    );
                    broadcast_embed(http, &state_snapshot.notify_channels, *embed).await
                }
                CheckResponse::Fallback(message) => {
                    warn!("periodic embed exceeded limit, sending fallback message");
                    broadcast_text(http, &state_snapshot.notify_channels, &message).await
                }
            }
            state.remember_notified_hash(snapshot.csv_hash).await?;
        }
    }

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

async fn fetch_latest_snapshot() -> anyhow::Result<Snapshot> {
    let date = date_prefix();
    info!("fetching latest snapshot for {}", date);
    let target = resolve_default_timetable_target()?;
    fetch_and_store(FetchRequest {
        target,
        local_xlsx_path: None,
        date: Some(date.clone()),
    })
    .await?;
    let show_result = resolve_show_result(Some(&date))?;
    let (_, csv_path) = resolve_csv_path(Some(&date))?;
    let csv_body = tokio::fs::read(&csv_path)
        .await
        .with_context(|| format!("failed to read {}", csv_path.display()))?;

    Ok(Snapshot {
        date,
        fetched_at: Local::now().format("%Y-%m-%d %H:%M").to_string(),
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

fn create_periodic_update_response(snapshot: &Snapshot) -> CheckResponse {
    create_embed_response(
        snapshot,
        "更新されました：現在発表中の時間割変更一覧",
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

fn format_manual_check_description(snapshot: &Snapshot) -> String {
    let entries = sorted_visible_entries(&snapshot.entries);
    if entries.is_empty() {
        return "現在発表中の変更はありません。".to_string();
    }

    let mut sections = Vec::new();
    let mut current_key: Option<(&str, &str, &str)> = None;
    let mut current_lines: Vec<String> = Vec::new();

    for entry in &entries {
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

fn sorted_visible_entries(
    entries: &[hyper_denpa_core::models::TimetableEntry],
) -> Vec<hyper_denpa_core::models::TimetableEntry> {
    let today = Local::now().date_naive();
    let mut visible = entries
        .iter()
        .filter_map(|entry| {
            let date = parse_entry_date(&entry.date)?;
            if date < today {
                return None;
            }

            Some((date, first_period_number(&entry.period), entry.clone()))
        })
        .collect::<Vec<_>>();

    visible.sort_by(|left, right| {
        left.0
            .cmp(&right.0)
            .then(left.1.cmp(&right.1))
            .then(left.2.grade.cmp(&right.2.grade))
            .then(left.2.class_name.cmp(&right.2.class_name))
    });

    visible
        .into_iter()
        .take(MAX_VISIBLE_ENTRIES)
        .map(|(_, _, entry)| entry)
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
