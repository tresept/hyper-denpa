use anyhow::{Context, bail};
use chrono::Local;
use knct_sharepoint::models::AuthMaterial;
use knct_sharepoint::pipeline::FetchRequest;
use knct_sharepoint::sharepoint::SharePointTarget;
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

pub const DATA_DIR: &str = "data";
pub const ENV_FILE: &str = ".env";
const DEFAULT_TIMETABLE_LABEL: &str = "時間割変更";

#[derive(Debug, Clone)]
pub struct OutputLayout {
    pub date: String,
    pub day_dir: PathBuf,
    pub xlsx_dir: PathBuf,
    pub csv_dir: PathBuf,
}

impl OutputLayout {
    pub fn new(date: impl Into<String>) -> Self {
        let date = date.into();
        let day_dir = PathBuf::from(DATA_DIR).join(&date);
        Self {
            date,
            xlsx_dir: day_dir.join("xlsx"),
            csv_dir: day_dir.join("csv"),
            day_dir,
        }
    }

    pub fn ensure(date: impl Into<String>) -> anyhow::Result<Self> {
        let layout = Self::new(date);
        fs::create_dir_all(&layout.day_dir)
            .with_context(|| format!("failed to create {}", layout.day_dir.display()))?;
        fs::create_dir_all(&layout.xlsx_dir)
            .with_context(|| format!("failed to create {}", layout.xlsx_dir.display()))?;
        fs::create_dir_all(&layout.csv_dir)
            .with_context(|| format!("failed to create {}", layout.csv_dir.display()))?;
        Ok(layout)
    }
}

pub fn run_prefix() -> String {
    Local::now().format("%Y-%m-%d_%H-%M-%S").to_string()
}

pub fn load_dotenv() -> anyhow::Result<HashMap<String, String>> {
    let body =
        fs::read_to_string(ENV_FILE).with_context(|| format!("failed to read {}", ENV_FILE))?;
    let mut values = HashMap::new();

    for line in body.lines() {
        let trimmed = line.trim();
        if trimmed.is_empty() || trimmed.starts_with('#') {
            continue;
        }

        let Some((key, value)) = trimmed.split_once('=') else {
            continue;
        };
        values.insert(key.trim().to_string(), parse_env_value(value.trim()));
    }

    values.extend(std::env::vars());
    Ok(values)
}

pub fn optional_env(values: &HashMap<String, String>, key: &str) -> Option<String> {
    values
        .get(key)
        .cloned()
        .filter(|value| !value.trim().is_empty())
}

pub fn resolve_document_url(values: &HashMap<String, String>) -> Option<String> {
    optional_env(values, "SHAREPOINT_DOC_URL")
}

pub fn resolve_fetch_request(
    values: &HashMap<String, String>,
    output_dir: PathBuf,
) -> anyhow::Result<FetchRequest> {
    let auth = resolve_auth_material(values)?;

    if let Some(document_url) = resolve_document_url(values) {
        return FetchRequest::from_document_url(document_url, auth, output_dir);
    }

    let site_url = required_env(values, "SHAREPOINT_SITE_URL")?;
    let item_id = required_env(values, "SHAREPOINT_ITEM_ID")?;

    Ok(FetchRequest {
        target: SharePointTarget {
            site_url,
            item_id,
            label: DEFAULT_TIMETABLE_LABEL.to_string(),
            document_url: None,
        },
        auth,
        output_dir,
    })
}

fn required_env(values: &HashMap<String, String>, key: &str) -> anyhow::Result<String> {
    optional_env(values, key).with_context(|| format!("{key} が .env に設定されていません"))
}

fn resolve_auth_material(values: &HashMap<String, String>) -> anyhow::Result<AuthMaterial> {
    if let Some(cookie_header) = optional_env(values, "SHAREPOINT_COOKIE_HEADER") {
        return Ok(AuthMaterial::CookieHeader(cookie_header));
    }

    let fed_auth = optional_env(values, "SHAREPOINT_FEDAUTH");
    let rt_fa = optional_env(values, "SHAREPOINT_RTFA");
    if let (Some(fed_auth), Some(rt_fa)) = (fed_auth, rt_fa) {
        return Ok(AuthMaterial::FedAuthRtFa { fed_auth, rt_fa });
    }

    if optional_env(values, "ESTSAUTHPERSISTENT").is_some() {
        bail!(
            "ESTSAUTHPERSISTENT は knct-sharepoint では未対応です。SHAREPOINT_COOKIE_HEADER か SHAREPOINT_FEDAUTH/SHAREPOINT_RTFA を設定してください"
        );
    }

    bail!(
        "SharePoint 認証情報が見つかりません。SHAREPOINT_COOKIE_HEADER か SHAREPOINT_FEDAUTH/SHAREPOINT_RTFA を設定してください"
    )
}

fn parse_env_value(raw: &str) -> String {
    let trimmed = raw.trim();
    let unquoted = if (trimmed.starts_with('"') && trimmed.ends_with('"'))
        || (trimmed.starts_with('\'') && trimmed.ends_with('\''))
    {
        &trimmed[1..trimmed.len() - 1]
    } else {
        trimmed
    };

    let mut out = String::with_capacity(unquoted.len());
    let mut chars = unquoted.chars();
    while let Some(ch) = chars.next() {
        if ch == '\\' {
            match chars.next() {
                Some('n') => out.push('\n'),
                Some('r') => out.push('\r'),
                Some('t') => out.push('\t'),
                Some('\\') => out.push('\\'),
                Some('"') => out.push('"'),
                Some('\'') => out.push('\''),
                Some(other) => {
                    out.push('\\');
                    out.push(other);
                }
                None => out.push('\\'),
            }
        } else {
            out.push(ch);
        }
    }

    out
}

#[cfg(test)]
mod tests {
    use super::parse_env_value;

    #[test]
    fn parse_env_value_supports_common_quotes() {
        assert_eq!(parse_env_value(" plain "), "plain");
        assert_eq!(parse_env_value("\"a b\""), "a b");
        assert_eq!(parse_env_value("'a b'"), "a b");
        assert_eq!(parse_env_value("\"line\\nend\""), "line\nend");
    }
}
