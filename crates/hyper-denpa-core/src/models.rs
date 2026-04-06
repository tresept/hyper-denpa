use crate::config::{DATA_DIR, TIMETABLE_SHEET_NAME};
use crate::fs_utils::ensure_output_dir;
use serde::{Deserialize, Serialize};
use std::path::PathBuf;

#[derive(Debug, Serialize, Deserialize)]
pub struct RunReport {
    pub metadata_status: u16,
    pub file_status: u16,
    pub file_path: String,
    pub auth_mode: String,
    pub csv_files: Vec<String>,
}

#[derive(Debug, Clone)]
pub struct OutputLayout {
    pub date: String,
    pub day_dir: PathBuf,
    pub xlsx_dir: PathBuf,
    pub csv_dir: PathBuf,
    pub metadata_dir: PathBuf,
}

impl OutputLayout {
    pub fn new(date: impl Into<String>) -> Self {
        let date = date.into();
        let day_dir = PathBuf::from(DATA_DIR).join(&date);
        Self {
            date,
            xlsx_dir: day_dir.join("xlsx"),
            csv_dir: day_dir.join("csv"),
            metadata_dir: day_dir.join("metadata"),
            day_dir,
        }
    }

    pub fn ensure(date: impl Into<String>) -> anyhow::Result<Self> {
        let layout = Self::new(date);
        ensure_output_dir(layout.day_dir.clone())?;
        ensure_output_dir(layout.xlsx_dir.clone())?;
        ensure_output_dir(layout.csv_dir.clone())?;
        ensure_output_dir(layout.metadata_dir.clone())?;
        Ok(layout)
    }

    pub fn xlsx_path(&self) -> PathBuf {
        self.xlsx_dir.join(format!("{}-original.xlsx", self.date))
    }

    pub fn timetable_csv_path(&self) -> PathBuf {
        self.csv_dir
            .join(format!("{}-{}.csv", self.date, TIMETABLE_SHEET_NAME))
    }
}

#[derive(Debug, Clone)]
pub enum AuthMaterial {
    CookieHeader(String),
    FedAuthRtFa { fed_auth: String, rt_fa: String },
    EstsAuthPersistent(String),
}

#[derive(Debug, Clone, Serialize)]
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
