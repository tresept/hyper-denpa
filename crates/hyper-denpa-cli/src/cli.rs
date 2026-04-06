use clap::{ArgAction, Parser};
use std::path::PathBuf;

#[derive(Parser, Debug)]
#[command(
    name = "hyper-denpa",
    version,
    about = "SharePoint時間割監視ツール",
    arg_required_else_help = true
)]
pub struct Cli {
    #[arg(long, action = ArgAction::SetTrue, group = "action", help = "SharePoint から XLSX/CSV を取得")]
    pub get: bool,
    #[arg(long, action = ArgAction::SetTrue, group = "action", help = "保存済み CSV から変更一覧を表示")]
    pub show: bool,
    #[arg(long, action = ArgAction::SetTrue, group = "action", help = "保存済み CSV をそのまま表示")]
    pub csv: bool,
    #[arg(
        long,
        requires = "show",
        conflicts_with_all = ["csv", "get"],
        help = "--show の結果を JSON で出力"
    )]
    pub json: bool,
    #[arg(long, requires = "action", help = "表示対象の日付 (yyyy-mm-dd)")]
    pub date: Option<String>,
    #[arg(long, hide = true, requires = "get")]
    pub local_xlsx_path: Option<PathBuf>,
}
