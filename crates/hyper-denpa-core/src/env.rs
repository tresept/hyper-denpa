use crate::config::ENV_FILE;
use anyhow::{Context, bail};
use std::collections::HashMap;
use std::fs;

fn parse_env_value(raw: &str) -> String {
    let trimmed = raw.trim();
    if trimmed.len() >= 2 {
        if trimmed.starts_with('"') && trimmed.ends_with('"') {
            return trimmed[1..trimmed.len() - 1]
                .replace("\\n", "\n")
                .replace("\\r", "\r")
                .replace("\\t", "\t")
                .replace("\\\"", "\"")
                .replace("\\\\", "\\");
        }
        if trimmed.starts_with('\'') && trimmed.ends_with('\'') {
            return trimmed[1..trimmed.len() - 1].to_string();
        }
    }

    trimmed.to_string()
}

pub fn load_dotenv() -> anyhow::Result<HashMap<String, String>> {
    let body =
        fs::read_to_string(ENV_FILE).with_context(|| format!("failed to read {}", ENV_FILE))?;
    let mut values = HashMap::new();

    for (index, raw_line) in body.lines().enumerate() {
        let line = raw_line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        let line = line.strip_prefix("export ").unwrap_or(line).trim();
        let Some((key, value)) = line.split_once('=') else {
            bail!(
                "failed to parse {}:{}: expected KEY=VALUE",
                ENV_FILE,
                index + 1
            );
        };
        let key = key.trim();
        if key.is_empty() {
            bail!("failed to parse {}:{}: empty key", ENV_FILE, index + 1);
        }

        values.insert(key.to_string(), parse_env_value(value));
    }

    Ok(values)
}

pub fn optional_env(values: &HashMap<String, String>, name: &str) -> Option<String> {
    values
        .get(name)
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty())
}

pub fn required_env(values: &HashMap<String, String>, name: &str) -> anyhow::Result<String> {
    optional_env(values, name).with_context(|| format!("{name} が {ENV_FILE} に設定されていません"))
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
