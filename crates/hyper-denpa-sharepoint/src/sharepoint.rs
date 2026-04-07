use crate::config::ENV_FILE;
use crate::env::{load_dotenv, optional_env, required_env};
use anyhow::Context;
use log::info;
use reqwest::header::ACCEPT;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SharePointTarget {
    pub site_url: String,
    pub item_id: String,
    pub label: String,
    pub document_url: Option<String>,
}

impl SharePointTarget {
    pub fn metadata_url(&self) -> String {
        format!(
            "{}/_api/web/GetFileById(guid'{}')?$select=Exists,Length,Name,ServerRelativeUrl,TimeLastModified,UniqueId",
            self.site_url, self.item_id
        )
    }

    pub fn download_url(&self) -> String {
        format!(
            "{}/_api/web/GetFileById(guid'{}')/$value",
            self.site_url, self.item_id
        )
    }
}

pub fn resolve_default_timetable_target() -> anyhow::Result<SharePointTarget> {
    let env_values = load_dotenv()
        .with_context(|| format!("SharePoint 設定用の {} を読み取れませんでした", ENV_FILE))?;

    Ok(SharePointTarget {
        site_url: required_env(&env_values, "SHAREPOINT_SITE_URL")?,
        item_id: required_env(&env_values, "SHAREPOINT_ITEM_ID")?,
        label: "時間割変更".to_string(),
        document_url: optional_env(&env_values, "SHAREPOINT_DOC_URL"),
    })
}

pub async fn fetch_file_metadata(
    client: &reqwest::Client,
    target: &SharePointTarget,
) -> anyhow::Result<reqwest::Response> {
    let metadata_url = target.metadata_url();
    info!(
        "fetching sharepoint metadata for {}: {}",
        target.label, metadata_url
    );

    client
        .get(metadata_url)
        .header(ACCEPT, "application/json;odata=nometadata")
        .send()
        .await
        .context("failed to fetch file metadata")
}

pub async fn download_file(
    client: &reqwest::Client,
    target: &SharePointTarget,
) -> anyhow::Result<reqwest::Response> {
    let download_url = target.download_url();
    info!(
        "downloading sharepoint file for {}: {}",
        target.label, download_url
    );

    client
        .get(download_url)
        .send()
        .await
        .context("failed to download excel file")
}

#[cfg(test)]
mod tests {
    use super::SharePointTarget;

    #[test]
    fn target_builds_expected_urls() {
        let target = SharePointTarget {
            site_url: "https://example.sharepoint.com/sites/demo".to_string(),
            item_id: "03D5B4F5-5F36-4158-BD00-297A14C1ABC2".to_string(),
            label: "時間割変更".to_string(),
            document_url: Some("https://example.sharepoint.com/doc".to_string()),
        };
        assert_eq!(
            target.metadata_url(),
            "https://example.sharepoint.com/sites/demo/_api/web/GetFileById(guid'03D5B4F5-5F36-4158-BD00-297A14C1ABC2')?$select=Exists,Length,Name,ServerRelativeUrl,TimeLastModified,UniqueId"
        );
        assert_eq!(
            target.download_url(),
            "https://example.sharepoint.com/sites/demo/_api/web/GetFileById(guid'03D5B4F5-5F36-4158-BD00-297A14C1ABC2')/$value"
        );
    }
}
