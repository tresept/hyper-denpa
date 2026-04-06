use anyhow::{Context, bail};
use chrono::{Duration, NaiveDate};
use csv::WriterBuilder;
use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::{HashMap, HashSet};
use std::fs;
use std::fs::File;
use std::io::{Read, Seek, Write};
use std::path::{Path, PathBuf};
use zip::ZipArchive;

fn read_zip_entry<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> anyhow::Result<String> {
    let mut file = archive
        .by_name(name)
        .with_context(|| format!("missing zip entry {name}"))?;
    let mut body = String::new();
    file.read_to_string(&mut body)
        .with_context(|| format!("failed to read zip entry {name}"))?;
    Ok(body)
}

fn read_optional_zip_entry<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> anyhow::Result<Option<String>> {
    match archive.by_name(name) {
        Ok(mut file) => {
            let mut body = String::new();
            file.read_to_string(&mut body)
                .with_context(|| format!("failed to read zip entry {name}"))?;
            Ok(Some(body))
        }
        Err(zip::result::ZipError::FileNotFound) => Ok(None),
        Err(error) => Err(error).with_context(|| format!("failed to open zip entry {name}")),
    }
}

fn attr_value(
    decoder: quick_xml::encoding::Decoder,
    start: &quick_xml::events::BytesStart<'_>,
    key: &[u8],
) -> anyhow::Result<Option<String>> {
    for attr in start.attributes() {
        let attr = attr.context("failed to parse xml attribute")?;
        if attr.key.as_ref() == key {
            return Ok(Some(
                attr.decode_and_unescape_value(decoder)
                    .context("failed to decode xml attribute")?
                    .into_owned(),
            ));
        }
    }

    Ok(None)
}

fn parse_shared_strings(xml: &str) -> anyhow::Result<Vec<String>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut values = Vec::new();
    let mut current = String::new();
    let mut in_si = false;
    let mut in_t = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref event)) => match event.name().as_ref() {
                b"si" => {
                    in_si = true;
                    current.clear();
                }
                b"t" if in_si => {
                    in_t = true;
                }
                _ => {}
            },
            Ok(Event::End(ref event)) => match event.name().as_ref() {
                b"si" => {
                    values.push(current.clone());
                    current.clear();
                    in_si = false;
                }
                b"t" => {
                    in_t = false;
                }
                _ => {}
            },
            Ok(Event::Text(text)) => {
                if in_si && in_t {
                    current.push_str(
                        &text
                            .decode()
                            .context("failed to decode shared string text")?,
                    );
                }
            }
            Ok(Event::CData(text)) => {
                if in_si && in_t {
                    current.push_str(
                        &text
                            .decode()
                            .context("failed to decode shared string cdata")?,
                    );
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error).context("failed to parse sharedStrings.xml"),
            _ => {}
        }
    }

    Ok(values)
}

fn is_date_format_code(format_code: &str) -> bool {
    let lowered = format_code.to_ascii_lowercase();
    let has_date_token = lowered.contains('y')
        || lowered.contains("m/")
        || lowered.contains("/d")
        || lowered.contains("d/")
        || lowered.contains('h')
        || lowered.contains('s');
    has_date_token && !lowered.contains("0_ ")
}

fn builtin_date_formats() -> HashSet<u32> {
    [14_u32, 15, 16, 17, 22, 27, 30, 36, 45, 46, 47, 50, 57]
        .into_iter()
        .collect()
}

fn parse_date_style_indexes(xml: &str) -> anyhow::Result<HashSet<usize>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let builtin = builtin_date_formats();
    let mut custom_date_formats = HashSet::new();
    let mut date_styles = HashSet::new();
    let mut in_cell_xfs = false;
    let mut xf_index = 0usize;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref event)) => match event.name().as_ref() {
                b"numFmt" => {
                    let Some(id) = attr_value(reader.decoder(), event, b"numFmtId")? else {
                        continue;
                    };
                    let Some(code) = attr_value(reader.decoder(), event, b"formatCode")? else {
                        continue;
                    };
                    let num_fmt_id = id.parse::<u32>().context("invalid numFmtId")?;
                    if is_date_format_code(&code) {
                        custom_date_formats.insert(num_fmt_id);
                    }
                }
                b"cellXfs" => {
                    in_cell_xfs = true;
                    xf_index = 0;
                }
                b"xf" if in_cell_xfs => {
                    let Some(id) = attr_value(reader.decoder(), event, b"numFmtId")? else {
                        xf_index += 1;
                        continue;
                    };
                    let num_fmt_id = id.parse::<u32>().context("invalid xf numFmtId")?;
                    if builtin.contains(&num_fmt_id) || custom_date_formats.contains(&num_fmt_id) {
                        date_styles.insert(xf_index);
                    }
                    xf_index += 1;
                }
                _ => {}
            },
            Ok(Event::Empty(ref event)) => match event.name().as_ref() {
                b"numFmt" => {
                    let Some(id) = attr_value(reader.decoder(), event, b"numFmtId")? else {
                        continue;
                    };
                    let Some(code) = attr_value(reader.decoder(), event, b"formatCode")? else {
                        continue;
                    };
                    let num_fmt_id = id.parse::<u32>().context("invalid numFmtId")?;
                    if is_date_format_code(&code) {
                        custom_date_formats.insert(num_fmt_id);
                    }
                }
                b"xf" if in_cell_xfs => {
                    let Some(id) = attr_value(reader.decoder(), event, b"numFmtId")? else {
                        xf_index += 1;
                        continue;
                    };
                    let num_fmt_id = id.parse::<u32>().context("invalid xf numFmtId")?;
                    if builtin.contains(&num_fmt_id) || custom_date_formats.contains(&num_fmt_id) {
                        date_styles.insert(xf_index);
                    }
                    xf_index += 1;
                }
                _ => {}
            },
            Ok(Event::End(ref event)) => {
                if event.name().as_ref() == b"cellXfs" {
                    in_cell_xfs = false;
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error).context("failed to parse styles.xml"),
            _ => {}
        }
    }

    Ok(date_styles)
}

fn parse_workbook_sheets(xml: &str) -> anyhow::Result<Vec<(String, String)>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut sheets = Vec::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref event)) | Ok(Event::Empty(ref event))
                if event.name().as_ref() == b"sheet" =>
            {
                let Some(name) = attr_value(reader.decoder(), event, b"name")? else {
                    continue;
                };
                let Some(rel_id) = attr_value(reader.decoder(), event, b"r:id")? else {
                    continue;
                };
                sheets.push((name, rel_id));
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error).context("failed to parse workbook.xml"),
            _ => {}
        }
    }

    Ok(sheets)
}

fn parse_relationships(xml: &str) -> anyhow::Result<HashMap<String, String>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut rels = HashMap::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref event)) | Ok(Event::Empty(ref event))
                if event.name().as_ref() == b"Relationship" =>
            {
                let Some(id) = attr_value(reader.decoder(), event, b"Id")? else {
                    continue;
                };
                let Some(target) = attr_value(reader.decoder(), event, b"Target")? else {
                    continue;
                };
                rels.insert(id, target);
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error).context("failed to parse workbook relationships"),
            _ => {}
        }
    }

    Ok(rels)
}

fn column_index_from_reference(reference: &str) -> usize {
    let mut index = 0usize;
    for ch in reference.chars() {
        if !ch.is_ascii_alphabetic() {
            break;
        }
        index = index * 26 + ((ch.to_ascii_uppercase() as u8 - b'A') as usize + 1);
    }

    index.saturating_sub(1)
}

fn normalize_sheet_target(target: &str) -> String {
    if let Some(stripped) = target.strip_prefix("../") {
        stripped.to_string()
    } else if target.starts_with("xl/") {
        target.to_string()
    } else {
        format!("xl/{target}")
    }
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

fn excel_serial_to_date(raw: &str) -> Option<String> {
    let serial = raw.parse::<f64>().ok()?;
    if !serial.is_finite() {
        return None;
    }
    let whole_days = serial.trunc() as i64;
    let date = NaiveDate::from_ymd_opt(1899, 12, 30)? + Duration::days(whole_days);
    Some(date.format("%Y/%-m/%-d").to_string())
}

fn parse_cell_value(
    raw_value: &str,
    inline_text: &str,
    cell_type: Option<&str>,
    style_index: Option<usize>,
    shared_strings: &[String],
    date_styles: &HashSet<usize>,
) -> String {
    match cell_type {
        Some("s") => raw_value
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index))
            .cloned()
            .unwrap_or_default(),
        Some("inlineStr") => inline_text.to_string(),
        Some("b") => {
            if raw_value == "1" {
                "TRUE".to_string()
            } else if raw_value == "0" {
                "FALSE".to_string()
            } else {
                raw_value.to_string()
            }
        }
        Some("str") => raw_value.to_string(),
        _ => {
            if let Some(style_index) = style_index {
                if date_styles.contains(&style_index) {
                    return excel_serial_to_date(raw_value)
                        .unwrap_or_else(|| raw_value.to_string());
                }
            }
            raw_value.to_string()
        }
    }
}

fn parse_sheet_rows(
    xml: &str,
    shared_strings: &[String],
    date_styles: &HashSet<usize>,
) -> anyhow::Result<Vec<Vec<String>>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);

    let mut rows = Vec::new();
    let mut current_row = Vec::new();
    let mut current_row_number = 0usize;

    let mut cell_reference = String::new();
    let mut cell_type: Option<String> = None;
    let mut cell_style: Option<usize> = None;
    let mut raw_value = String::new();
    let mut inline_text = String::new();
    let mut reading_value = false;
    let mut reading_text = false;
    let mut in_cell = false;

    loop {
        match reader.read_event() {
            Ok(Event::Start(ref event)) => match event.name().as_ref() {
                b"row" => {
                    let row_number = attr_value(reader.decoder(), event, b"r")?
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(rows.len() + 1);

                    while rows.len() + 1 < row_number {
                        rows.push(Vec::new());
                    }

                    current_row.clear();
                    current_row_number = row_number;
                }
                b"c" => {
                    in_cell = true;
                    cell_reference = attr_value(reader.decoder(), event, b"r")?.unwrap_or_default();
                    cell_type = attr_value(reader.decoder(), event, b"t")?;
                    cell_style = attr_value(reader.decoder(), event, b"s")?
                        .and_then(|value| value.parse::<usize>().ok());
                    raw_value.clear();
                    inline_text.clear();
                }
                b"v" if in_cell => {
                    reading_value = true;
                }
                b"t" if in_cell => {
                    reading_text = true;
                }
                _ => {}
            },
            Ok(Event::Empty(ref event)) => {
                if event.name().as_ref() == b"c" {
                    let cell_reference =
                        attr_value(reader.decoder(), event, b"r")?.unwrap_or_default();
                    let column_index = column_index_from_reference(&cell_reference);
                    while current_row.len() < column_index {
                        current_row.push(String::new());
                    }
                    current_row.push(String::new());
                }
            }
            Ok(Event::End(ref event)) => match event.name().as_ref() {
                b"v" => {
                    reading_value = false;
                }
                b"t" => {
                    reading_text = false;
                }
                b"c" => {
                    let column_index = column_index_from_reference(&cell_reference);
                    while current_row.len() < column_index {
                        current_row.push(String::new());
                    }

                    current_row.push(parse_cell_value(
                        &raw_value,
                        &inline_text,
                        cell_type.as_deref(),
                        cell_style,
                        shared_strings,
                        date_styles,
                    ));

                    in_cell = false;
                    cell_reference.clear();
                    cell_type = None;
                    cell_style = None;
                    raw_value.clear();
                    inline_text.clear();
                }
                b"row" => {
                    if current_row_number > 0 {
                        rows.push(current_row.clone());
                        current_row.clear();
                        current_row_number = 0;
                    }
                }
                _ => {}
            },
            Ok(Event::Text(text)) => {
                let decoded = text.decode().context("failed to decode worksheet text")?;
                if reading_value {
                    raw_value.push_str(&decoded);
                }
                if reading_text {
                    inline_text.push_str(&decoded);
                }
            }
            Ok(Event::CData(text)) => {
                let decoded = text.decode().context("failed to decode worksheet cdata")?;
                if reading_value {
                    raw_value.push_str(&decoded);
                }
                if reading_text {
                    inline_text.push_str(&decoded);
                }
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(error).context("failed to parse worksheet xml"),
            _ => {}
        }
    }

    Ok(rows)
}

pub fn convert_xlsx_to_csvs(
    xlsx_path: &Path,
    output_dir: &Path,
    prefix: &str,
) -> anyhow::Result<Vec<PathBuf>> {
    fs::create_dir_all(output_dir)
        .with_context(|| format!("failed to create {}", output_dir.display()))?;

    let file =
        File::open(xlsx_path).with_context(|| format!("failed to open {}", xlsx_path.display()))?;
    let mut archive = ZipArchive::new(file).context("failed to open xlsx as zip archive")?;

    let workbook_xml = read_zip_entry(&mut archive, "xl/workbook.xml")?;
    let workbook_rels_xml = read_zip_entry(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let shared_strings_xml = read_optional_zip_entry(&mut archive, "xl/sharedStrings.xml")?;
    let styles_xml = read_optional_zip_entry(&mut archive, "xl/styles.xml")?;

    let shared_strings = match shared_strings_xml {
        Some(xml) => parse_shared_strings(&xml)?,
        None => Vec::new(),
    };
    let date_styles = match styles_xml {
        Some(xml) => parse_date_style_indexes(&xml)?,
        None => HashSet::new(),
    };

    let sheets = parse_workbook_sheets(&workbook_xml)?;
    let relationships = parse_relationships(&workbook_rels_xml)?;

    let mut output_paths = Vec::new();

    for (sheet_name, rel_id) in sheets {
        let Some(target) = relationships.get(&rel_id) else {
            continue;
        };
        let worksheet_path = normalize_sheet_target(target);
        let worksheet_xml = read_zip_entry(&mut archive, &worksheet_path)?;
        let rows = parse_sheet_rows(&worksheet_xml, &shared_strings, &date_styles)?;

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

        for row in rows {
            writer
                .write_record(row.iter())
                .with_context(|| format!("failed to write {}", csv_path.display()))?;
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
