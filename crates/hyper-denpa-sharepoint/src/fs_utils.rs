use anyhow::Context;
use chrono::Local;
use serde::Serialize;
use std::fs;
use std::path::{Path, PathBuf};

pub fn ensure_output_dir(path: PathBuf) -> anyhow::Result<PathBuf> {
    fs::create_dir_all(&path).with_context(|| format!("failed to create {}", path.display()))?;
    Ok(path)
}

pub fn date_prefix() -> String {
    Local::now().format("%Y-%m-%d").to_string()
}

pub fn run_prefix() -> String {
    Local::now().format("%Y-%m-%d_%H-%M-%S").to_string()
}

pub fn write_text(path: &Path, body: &str) -> anyhow::Result<()> {
    fs::write(path, body).with_context(|| format!("failed to write {}", path.display()))
}

pub fn write_bytes(path: &Path, body: &[u8]) -> anyhow::Result<()> {
    fs::write(path, body).with_context(|| format!("failed to write {}", path.display()))
}

pub fn write_json<T: Serialize>(path: &Path, value: &T) -> anyhow::Result<()> {
    write_text(path, &serde_json::to_string_pretty(value)?)
}
