use anyhow::{Context, bail};
use calamine::{Data, DataType, Reader, open_workbook_auto};
use chrono::{NaiveDate, NaiveDateTime};
use csv::{ReaderBuilder, StringRecord, WriterBuilder};
use hyper_denpa_core::config::DATA_DIR;
use hyper_denpa_core::models::OutputLayout;
use serde::{Deserialize, Serialize};
use std::fs;
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};

const TIMETABLE_SHEET_NAME: &str = "時間割変更";

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct TimetableEntry {
    pub grade: String,
    pub class_name: String,
    pub date: String,
    pub weekday: String,
    pub period: String,
    pub change_type: String,
    pub subject: String,
}

#[derive(Debug, Serialize)]
pub struct ShowResult {
    pub date: String,
    pub csv_path: String,
    pub entries: Vec<TimetableEntry>,
}

fn trim_bom(value: &str) -> &str {
    value.trim_start_matches('\u{feff}')
}

fn sanitize_sheet_name(name: &str) -> String {
    let sanitized: String = name
        .chars()
        .map(|ch| match ch {
            '/' | '\\' | ':' | '*' | '?' | '"' | '<' | '>' | '|' => '_',
            _ => ch,
        })
        .collect();

    if sanitized.is_empty() {
        "sheet".to_string()
    } else {
        sanitized
    }
}

fn format_date(date: NaiveDate) -> String {
    date.format("%Y/%-m/%-d").to_string()
}

fn parse_iso_date(value: &str) -> Option<String> {
    NaiveDate::parse_from_str(value, "%Y-%m-%d")
        .ok()
        .map(format_date)
        .or_else(|| {
            NaiveDateTime::parse_from_str(value, "%Y-%m-%dT%H:%M:%S%.f")
                .ok()
                .map(|datetime| format_date(datetime.date()))
        })
        .or_else(|| {
            NaiveDateTime::parse_from_str(value, "%Y-%m-%d %H:%M:%S%.f")
                .ok()
                .map(|datetime| format_date(datetime.date()))
        })
}

fn cell_to_string(cell: &Data) -> String {
    match cell {
        Data::Empty => String::new(),
        Data::String(value) => value.clone(),
        Data::Bool(value) => {
            if *value {
                "TRUE".to_string()
            } else {
                "FALSE".to_string()
            }
        }
        Data::Int(value) => value.to_string(),
        Data::Float(value) => value.to_string(),
        Data::DateTime(_) => cell
            .as_date()
            .map(format_date)
            .or_else(|| {
                cell.as_datetime()
                    .map(|datetime| format_date(datetime.date()))
            })
            .unwrap_or_else(|| cell.to_string()),
        Data::DateTimeIso(value) => parse_iso_date(value).unwrap_or_else(|| value.clone()),
        Data::DurationIso(value) => value.clone(),
        Data::Error(error) => error.to_string(),
    }
}

fn trim_trailing_empty_fields(record: &mut Vec<String>) {
    while matches!(record.last(), Some(value) if value.is_empty()) {
        record.pop();
    }
}

fn row_to_record(row: &[Data], leading_columns: usize) -> Vec<String> {
    let mut record = vec![String::new(); leading_columns];
    record.extend(row.iter().map(cell_to_string));
    trim_trailing_empty_fields(&mut record);
    record
}

fn write_blank_row<W: Write>(writer: &mut csv::Writer<W>, csv_path: &Path) -> anyhow::Result<()> {
    writer
        .write_record(std::iter::empty::<&str>())
        .with_context(|| format!("failed to write {}", csv_path.display()))
}

pub fn convert_xlsx_to_csvs(
    xlsx_path: &Path,
    output_dir: &Path,
    prefix: &str,
) -> anyhow::Result<Vec<PathBuf>> {
    fs::create_dir_all(output_dir)
        .with_context(|| format!("failed to create {}", output_dir.display()))?;

    let mut workbook = open_workbook_auto(xlsx_path)
        .with_context(|| format!("failed to open {}", xlsx_path.display()))?;
    let sheet_names = workbook.sheet_names().to_owned();
    let mut output_paths = Vec::new();

    for sheet_name in sheet_names {
        let range = workbook
            .worksheet_range(&sheet_name)
            .with_context(|| format!("failed to read worksheet {sheet_name}"))?;

        let csv_path =
            output_dir.join(format!("{prefix}-{}.csv", sanitize_sheet_name(&sheet_name)));
        let mut csv_file = File::create(&csv_path)
            .with_context(|| format!("failed to create {}", csv_path.display()))?;
        csv_file
            .write_all(b"\xEF\xBB\xBF")
            .with_context(|| format!("failed to write BOM to {}", csv_path.display()))?;

        let mut writer = WriterBuilder::new()
            .has_headers(false)
            .flexible(true)
            .terminator(csv::Terminator::CRLF)
            .from_writer(csv_file);

        if let Some((start_row, start_col)) = range.start() {
            for _ in 0..start_row {
                write_blank_row(&mut writer, &csv_path)?;
            }

            for row in range.rows() {
                let record = row_to_record(row, start_col as usize);
                writer
                    .write_record(record.iter())
                    .with_context(|| format!("failed to write {}", csv_path.display()))?;
            }
        }

        writer
            .flush()
            .with_context(|| format!("failed to flush {}", csv_path.display()))?;
        output_paths.push(csv_path);
    }

    if output_paths.is_empty() {
        bail!(
            "xlsx から CSV を生成できませんでした: {}",
            xlsx_path.display()
        );
    }

    Ok(output_paths)
}

fn normalize_record(record: &StringRecord) -> Vec<String> {
    record
        .iter()
        .map(|value| trim_bom(value).trim().to_string())
        .collect()
}

fn parse_timetable_entries(csv_path: &Path) -> anyhow::Result<Vec<TimetableEntry>> {
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

fn parse_run_key(value: &str) -> Option<NaiveDateTime> {
    NaiveDateTime::parse_from_str(value, "%Y-%m-%d_%H-%M-%S")
        .ok()
        .or_else(|| {
            NaiveDate::parse_from_str(value, "%Y-%m-%d")
                .ok()
                .and_then(|date| date.and_hms_opt(0, 0, 0))
        })
}

fn latest_data_date() -> anyhow::Result<String> {
    let data_dir = PathBuf::from(DATA_DIR);
    if !data_dir.exists() {
        bail!("data ディレクトリがまだありません。先に `hyper-denpa` を実行してください。");
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
        if let Some(parsed) = parse_run_key(name) {
            dates.push((parsed, name.to_string()));
        }
    }

    dates.sort_by(|left, right| left.0.cmp(&right.0));
    dates
        .pop()
        .map(|(_, name)| name)
        .context("data 配下に日付ディレクトリがありません")
}

fn resolve_date(date: Option<&str>) -> anyhow::Result<String> {
    match date {
        Some(date) => {
            parse_run_key(date).with_context(|| format!("invalid date: {date}"))?;
            Ok(date.to_string())
        }
        None => latest_data_date(),
    }
}

pub fn timetable_csv_path(layout: &OutputLayout) -> PathBuf {
    layout
        .csv_dir
        .join(format!("{}-{}.csv", layout.date, TIMETABLE_SHEET_NAME))
}

pub fn resolve_show_result(date: Option<&str>) -> anyhow::Result<ShowResult> {
    let date = resolve_date(date)?;
    let layout = OutputLayout::new(&date);
    let csv_path = timetable_csv_path(&layout);
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
    let csv_path = timetable_csv_path(&layout);
    if !csv_path.exists() {
        bail!(
            "{} が見つかりません: {}",
            TIMETABLE_SHEET_NAME,
            csv_path.display()
        );
    }

    Ok((date, csv_path))
}
