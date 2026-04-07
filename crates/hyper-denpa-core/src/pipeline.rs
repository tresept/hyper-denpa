use crate::auth::{build_client, resolve_auth_material};
use crate::fs_utils::{date_prefix, write_bytes, write_json, write_text};
use crate::models::{OutputLayout, RunReport};
use crate::sharepoint::{SharePointTarget, download_file, fetch_file_metadata};
use anyhow::{Context, bail};
use log::info;
use std::fs;
use std::path::PathBuf;

#[derive(Debug, Clone)]
pub struct FetchRequest {
    pub target: SharePointTarget,
    pub local_xlsx_path: Option<PathBuf>,
    pub date: Option<String>,
}

fn write_report(layout: &OutputLayout, report: &RunReport) -> anyhow::Result<()> {
    write_json(&layout.metadata_dir.join("report.json"), report)
}

pub async fn fetch_and_store(request: FetchRequest) -> anyhow::Result<RunReport> {
    let today = request.date.unwrap_or_else(date_prefix);
    info!(
        "starting fetch_and_store for target={} date={}",
        request.target.label, today
    );
    let layout = OutputLayout::ensure(&today)?;

    if let Some(local_xlsx_path) = request.local_xlsx_path {
        info!(
            "using local xlsx path {} for target={}",
            local_xlsx_path.display(),
            request.target.label
        );
        let managed_xlsx_path = layout.xlsx_path();
        if local_xlsx_path != managed_xlsx_path {
            fs::copy(&local_xlsx_path, &managed_xlsx_path).with_context(|| {
                format!(
                    "failed to copy {} to {}",
                    local_xlsx_path.display(),
                    managed_xlsx_path.display()
                )
            })?;
        }

        let report = RunReport {
            metadata_status: 0,
            file_status: 0,
            file_path: managed_xlsx_path.display().to_string(),
            auth_mode: "local_xlsx_path".to_string(),
            csv_files: Vec::new(),
        };
        write_report(&layout, &report)?;
        info!("completed local xlsx import for {}", request.target.label);
        return Ok(report);
    }

    let (auth_material, auth_mode) = resolve_auth_material()?;
    info!("resolved auth mode {}", auth_mode);
    let client = build_client(&auth_material)?;

    if let Some(document_url) = &request.target.document_url {
        write_text(
            &layout.metadata_dir.join("sharepoint_doc_url.txt"),
            document_url,
        )?;
    }

    let metadata_res = fetch_file_metadata(&client, &request.target).await?;
    let metadata_status = metadata_res.status().as_u16();
    let metadata_final_url = metadata_res.url().to_string();
    let metadata_headers = format!("{:#?}", metadata_res.headers());
    let metadata_body = metadata_res.text().await.unwrap_or_default();
    write_text(
        &layout.metadata_dir.join("file_metadata_headers.txt"),
        &metadata_headers,
    )?;
    write_text(
        &layout.metadata_dir.join("file_metadata.json"),
        &metadata_body,
    )?;

    if metadata_status >= 400 {
        bail!(
            "metadata request failed with status {} at {}. saved body to {}",
            metadata_status,
            metadata_final_url,
            layout.metadata_dir.join("file_metadata.json").display()
        );
    }

    let file_res = download_file(&client, &request.target).await?;
    let file_status = file_res.status().as_u16();
    let file_headers = format!("{:#?}", file_res.headers());
    let file_bytes = file_res
        .bytes()
        .await
        .context("failed to read excel bytes")?;

    write_text(
        &layout.metadata_dir.join("file_download_headers.txt"),
        &file_headers,
    )?;

    let file_path = layout.xlsx_path();
    if file_status >= 400 {
        let error_path = layout.metadata_dir.join("file_download_error.bin");
        write_bytes(&error_path, &file_bytes)?;
        bail!(
            "file download failed with status {}. saved body to {}",
            file_status,
            error_path.display()
        );
    }

    write_bytes(&file_path, &file_bytes)?;
    info!("downloaded xlsx to {}", file_path.display());

    let report = RunReport {
        metadata_status,
        file_status,
        file_path: file_path.display().to_string(),
        auth_mode,
        csv_files: Vec::new(),
    };
    write_report(&layout, &report)?;
    info!(
        "completed fetch_and_store for {} with {} csv files",
        request.target.label,
        report.csv_files.len()
    );
    Ok(report)
}
