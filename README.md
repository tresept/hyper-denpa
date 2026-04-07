# hyper-denpa
[![GitHub](https://img.shields.io/badge/GitHub-tresept%2Fhyper--denpa-blue?logo=github)](https://github.com/tresept/hyper-denpa)

SharePointから時間割変更Excelを引っこ抜いて整形してDiscordに送ったりCLIとしてOpenClawのSkillとして使えるようにする試み

## Workspace

```text
crates/
  hyper-denpa-core
  hyper-denpa-cli
  hyper-denpa-discord
```

- `hyper-denpa-core`: SharePoint 取得と解析の共通ライブラリ
- `hyper-denpa-cli`: `cargo install --path crates/hyper-denpa-cli` で入れる CLI（実運用はまだ非推奨）
- `hyper-denpa-discord`: Discord Bot

## Auth

認証情報は `.env` から読み込みます．次のすべてを設定するか：

```env
RUST_LOG=warn,hyper_denpa_core=debug,hyper_denpa_discord=debug
DISCORD_TOKEN=...
SHAREPOINT_SITE_URL=https://example.sharepoint.com/sites/example
SHAREPOINT_ITEM_ID=00000000-0000-0000-0000-000000000000
SHAREPOINT_DOC_URL=https://example.sharepoint.com/:x:/r/sites/example/_layouts/15/Doc.aspx?sourcedoc=%7B00000000-0000-0000-0000-000000000000%7D&file=henkou.xlsx
SHAREPOINT_FEDAUTH=...
SHAREPOINT_RTFA=...
ESTSAUTHPERSISTENT=...
```

または `SHAREPOINT_COOKIE_HEADER="FedAuth=...; rtFa=..."` でも動きます．

`SHAREPOINT_SITE_URL` と `SHAREPOINT_ITEM_ID` は必須です．`SHAREPOINT_DOC_URL` は Discord の案内文にリンクを出したい場合だけ設定してください

## Commands

```bash
cargo run -p hyper-denpa-cli -- --help
cargo run -p hyper-denpa-cli -- --get
cargo run -p hyper-denpa-discord
```

Discord Bot の slash command:

- `/check`: 最新の時間割変更を強制再取得
- `/law-csv`: 最新の生 CSV を添付
- `/set-notify`: 実行したチャンネルを定期通知先に追加
- `/unset-notify`: 指定チャンネル、または実行したチャンネルを通知先から解除

Bot のログレベルは `.env` の `RUST_LOG` で切り替えられます。例:

```env
RUST_LOG=warn,hyper_denpa_core=debug,hyper_denpa_discord=debug
```
