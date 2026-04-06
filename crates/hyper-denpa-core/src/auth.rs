use crate::config::{DEFAULT_USER_AGENT, ENV_FILE};
use crate::env::{load_dotenv, optional_env};
use crate::models::AuthMaterial;
use crate::sharepoint::resolve_default_timetable_target;
use anyhow::{Context, bail};
use log::{debug, info};
use reqwest::cookie::Jar;
use std::sync::Arc;

fn add_cookie_header_to_jar(
    jar: &Jar,
    url: &reqwest::Url,
    cookie_header: &str,
) -> anyhow::Result<()> {
    let mut inserted = 0usize;

    for segment in cookie_header.split(';') {
        let pair = segment.trim();
        if pair.is_empty() {
            continue;
        }

        let Some((name, value)) = pair.split_once('=') else {
            continue;
        };
        let name = name.trim();
        let value = value.trim();
        if name.is_empty() || value.is_empty() {
            continue;
        }

        jar.add_cookie_str(&format!("{name}={value}"), url);
        inserted += 1;
    }

    if inserted == 0 {
        bail!("cookie header did not contain any valid key=value pairs");
    }

    Ok(())
}

pub fn resolve_auth_material() -> anyhow::Result<(AuthMaterial, String)> {
    debug!("loading auth material from {}", ENV_FILE);
    let env_values =
        load_dotenv().with_context(|| format!("認証用の {} を読み取れませんでした", ENV_FILE))?;

    if let Some(cookie_header) = optional_env(&env_values, "SHAREPOINT_COOKIE_HEADER") {
        info!("resolved auth material from SHAREPOINT_COOKIE_HEADER");
        return Ok((
            AuthMaterial::CookieHeader(cookie_header),
            "dotenv_cookie_header".to_string(),
        ));
    }

    if let (Some(fed_auth), Some(rt_fa)) = (
        optional_env(&env_values, "SHAREPOINT_FEDAUTH"),
        optional_env(&env_values, "SHAREPOINT_RTFA"),
    ) {
        info!("resolved auth material from SHAREPOINT_FEDAUTH/SHAREPOINT_RTFA");
        return Ok((
            AuthMaterial::FedAuthRtFa { fed_auth, rt_fa },
            "dotenv_fedauth_rtfa".to_string(),
        ));
    }

    if let Some(ests_auth_persistent) = optional_env(&env_values, "ESTSAUTHPERSISTENT") {
        info!("resolved auth material from ESTSAUTHPERSISTENT");
        return Ok((
            AuthMaterial::EstsAuthPersistent(ests_auth_persistent),
            "dotenv_estsauthpersistent".to_string(),
        ));
    }

    bail!(
        "認証情報がありません。{} に SHAREPOINT_COOKIE_HEADER か、SHAREPOINT_FEDAUTH / SHAREPOINT_RTFA、または ESTSAUTHPERSISTENT を設定してください。",
        ENV_FILE
    )
}

pub fn build_client(auth: &AuthMaterial) -> anyhow::Result<reqwest::Client> {
    let jar = Arc::new(Jar::default());
    let target = resolve_default_timetable_target()?;
    debug!("building reqwest client for target {}", target.label);
    let sharepoint_url =
        reqwest::Url::parse(&target.site_url).context("failed to parse sharepoint url")?;
    let login_url = reqwest::Url::parse("https://login.microsoftonline.com/")
        .context("failed to parse login url")?;

    match auth {
        AuthMaterial::CookieHeader(cookie_header) => {
            add_cookie_header_to_jar(&jar, &sharepoint_url, cookie_header)?;
            if cookie_header.contains("ESTSAUTHPERSISTENT=") {
                add_cookie_header_to_jar(&jar, &login_url, cookie_header)?;
            }
        }
        AuthMaterial::FedAuthRtFa { fed_auth, rt_fa } => {
            jar.add_cookie_str(&format!("FedAuth={fed_auth}"), &sharepoint_url);
            jar.add_cookie_str(&format!("rtFa={rt_fa}"), &sharepoint_url);
        }
        AuthMaterial::EstsAuthPersistent(token) => {
            jar.add_cookie_str("AADSSO=NA|NoExtension", &login_url);
            jar.add_cookie_str(&format!("ESTSAUTHPERSISTENT={token}"), &login_url);
        }
    }

    reqwest::Client::builder()
        .cookie_provider(jar)
        .user_agent(DEFAULT_USER_AGENT)
        .build()
        .context("failed to build reqwest client")
}
