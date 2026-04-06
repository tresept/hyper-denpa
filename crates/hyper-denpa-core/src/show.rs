use crate::config::{DATA_DIR, TIMETABLE_SHEET_NAME};
use crate::models::{OutputLayout, ShowResult, TimetableEntry};
use anyhow::{Context, bail};
use chrono::NaiveDate;
use csv::{ReaderBuilder, StringRecord};
use std::fs;
use std::path::PathBuf;

fn trim_bom(value: &str) -> &str {
    value.trim_start_matches('\u{feff}')
}

fn normalize_record(record: &StringRecord) -> Vec<String> {
    record
        .iter()
        .map(|value| trim_bom(value).trim().to_string())
        .collect()
}

fn parse_timetable_entries(csv_path: &std::path::Path) -> anyhow::Result<Vec<TimetableEntry>> {
    let mut reader = ReaderBuilder::new()
        .has_headers(false)
        .flexible(true)
        .from_path(csv_path)
        .with_context(|| format!("failed to open {}", csv_path.display()))?;

    let mut saw_header = false;
    let mut entries = Vec::new();

    for record in reader.records() {
        let record = record.with_context(|| format!("failed to read {}", csv_path.display()))?;
        let mut values = normalize_record(&record);

        if !saw_header {
            if values.len() >= 7
                && trim_bom(values[0].as_str()) == "学 年"
                && values[1] == "学科・クラス"
                && values[2] == "月日"
            {
                saw_header = true;
            }
            continue;
        }

        if values.len() < 7 {
            values.resize(7, String::new());
        }

        if values[..7].iter().all(|value| value.is_empty()) {
            continue;
        }

        entries.push(TimetableEntry {
            grade: values[0].clone(),
            class_name: values[1].clone(),
            date: values[2].clone(),
            weekday: values[3].clone(),
            period: values[4].clone(),
            change_type: values[5].clone(),
            subject: values[6].clone(),
        });
    }

    if !saw_header {
        bail!(
            "時間割変更 CSV のヘッダ行が見つかりませんでした: {}",
            csv_path.display()
        );
    }

    Ok(entries)
}

fn latest_data_date() -> anyhow::Result<String> {
    let data_dir = PathBuf::from(DATA_DIR);
    if !data_dir.exists() {
        bail!("data ディレクトリがまだありません。先に `hyper-denpa --get` を実行してください。");
    }

    let mut dates = Vec::new();
    for entry in
        fs::read_dir(&data_dir).with_context(|| format!("failed to read {}", data_dir.display()))?
    {
        let entry = entry?;
        if !entry.file_type()?.is_dir() {
            continue;
        }
        let name = entry.file_name();
        let Some(name) = name.to_str() else {
            continue;
        };
        if NaiveDate::parse_from_str(name, "%Y-%m-%d").is_ok() {
            dates.push(name.to_string());
        }
    }

    dates.sort();
    dates
        .pop()
        .context("data 配下に日付ディレクトリがありません")
}

fn resolve_date(date: Option<&str>) -> anyhow::Result<String> {
    match date {
        Some(date) => {
            NaiveDate::parse_from_str(date, "%Y-%m-%d")
                .with_context(|| format!("invalid date: {date}"))?;
            Ok(date.to_string())
        }
        None => latest_data_date(),
    }
}

pub fn resolve_show_result(date: Option<&str>) -> anyhow::Result<ShowResult> {
    let date = resolve_date(date)?;
    let layout = OutputLayout::new(&date);
    let csv_path = layout.timetable_csv_path();
    if !csv_path.exists() {
        bail!(
            "{} が見つかりません: {}",
            TIMETABLE_SHEET_NAME,
            csv_path.display()
        );
    }

    Ok(ShowResult {
        date,
        csv_path: csv_path.display().to_string(),
        entries: parse_timetable_entries(&csv_path)?,
    })
}

pub fn resolve_csv_path(date: Option<&str>) -> anyhow::Result<(String, PathBuf)> {
    let date = resolve_date(date)?;
    let layout = OutputLayout::new(&date);
    let csv_path = layout.timetable_csv_path();
    if !csv_path.exists() {
        bail!(
            "{} が見つかりません: {}",
            TIMETABLE_SHEET_NAME,
            csv_path.display()
        );
    }

    Ok((date, csv_path))
}

pub fn print_show_result(result: &ShowResult) {
    println!(
        "{} の時間割変更一覧 ({}件)",
        result.date,
        result.entries.len()
    );
    println!("source: {}", result.csv_path);
    println!();

    for (index, entry) in result.entries.iter().enumerate() {
        println!(
            "{:>2}. {} {} {} | {} {} | {} | {}",
            index + 1,
            entry.date,
            entry.weekday,
            entry.period,
            entry.grade,
            entry.class_name,
            entry.change_type,
            entry.subject
        );
    }
}

pub fn run_show(json: bool, date: Option<String>) -> anyhow::Result<()> {
    let result = resolve_show_result(date.as_deref())?;

    if json {
        println!("{}", serde_json::to_string_pretty(&result)?);
    } else {
        print_show_result(&result);
    }

    Ok(())
}

pub fn run_csv(date: Option<String>) -> anyhow::Result<()> {
    let (_, csv_path) = resolve_csv_path(date.as_deref())?;
    let body = fs::read_to_string(&csv_path)
        .with_context(|| format!("failed to read {}", csv_path.display()))?;
    print!("{}", trim_bom(&body));
    Ok(())
}
