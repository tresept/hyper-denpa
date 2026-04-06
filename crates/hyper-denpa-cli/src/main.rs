mod cli;

use clap::Parser;
use cli::Cli;
use hyper_denpa_core::models::OutputLayout;
use hyper_denpa_core::pipeline::{FetchRequest, fetch_and_store};
use hyper_denpa_core::sharepoint::resolve_default_timetable_target;

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let cli = Cli::parse();

    if cli.get {
        let target = resolve_default_timetable_target()?;
        let date = cli
            .date
            .clone()
            .unwrap_or_else(hyper_denpa_core::fs_utils::date_prefix);
        let report = fetch_and_store(FetchRequest {
            target,
            local_xlsx_path: cli.local_xlsx_path,
            date: Some(date.clone()),
        })
        .await?;
        let layout = OutputLayout::new(date);
        println!(
            "完了: {} に保存しました (auth={})",
            layout.day_dir.display(),
            report.auth_mode
        );
        return Ok(());
    }

    if cli.show {
        hyper_denpa_core::show::run_show(cli.json, cli.date)?;
        return Ok(());
    }

    if cli.csv {
        hyper_denpa_core::show::run_csv(cli.date)?;
        return Ok(());
    }

    Ok(())
}
