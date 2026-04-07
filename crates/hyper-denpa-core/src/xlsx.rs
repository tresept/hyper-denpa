use anyhow::{Context, bail};
use calamine::{Data, DataType, Reader, open_workbook_auto};
use chrono::{NaiveDate, NaiveDateTime};
use csv::WriterBuilder;
use std::fs;
use std::fs::File;
use std::io::Write;
use std::path::{Path, PathBuf};

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
            .or_else(|| cell.as_datetime().map(|datetime| format_date(datetime.date())))
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

#[cfg(test)]
mod tests {
    use super::convert_xlsx_to_csvs;
    use std::fs;
    use std::path::Path;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn trim_bom(value: &str) -> &str {
        value.trim_start_matches('\u{feff}')
    }

    fn convert_fixture(date: &str) -> std::path::PathBuf {
        let manifest_dir = Path::new(env!("CARGO_MANIFEST_DIR"));
        let xlsx_path = manifest_dir
            .join("../../data")
            .join(date)
            .join("xlsx")
            .join(format!("{date}-original.xlsx"));
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("system time should be after unix epoch")
            .as_nanos();
        let output_dir = std::env::temp_dir().join(format!("hyper-denpa-xlsx-test-{date}-{unique}"));

        convert_xlsx_to_csvs(&xlsx_path, &output_dir, date)
            .expect("fixture conversion should succeed");

        output_dir
    }

    #[test]
    fn converts_fixture_for_2026_04_07() {
        let output_dir = convert_fixture("2026-04-07");
        let timetable_path = output_dir.join("2026-04-07-時間割変更.csv");
        let notes_path = output_dir.join("2026-04-07-注意事項.csv");

        assert!(timetable_path.exists());
        assert!(notes_path.exists());

        let timetable = fs::read_to_string(&timetable_path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", timetable_path.display()));
        let notes = fs::read_to_string(&notes_path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", notes_path.display()));

        assert!(trim_bom(&timetable).contains("学 年,学科・クラス,月日"));
        assert!(trim_bom(&timetable).contains("2026/4/7"));
        assert!(trim_bom(&notes).contains("時間割変更"));
        assert!(trim_bom(&notes).lines().count() >= 3);

        fs::remove_dir_all(&output_dir)
            .unwrap_or_else(|error| panic!("failed to clean {}: {error}", output_dir.display()));
    }
}
