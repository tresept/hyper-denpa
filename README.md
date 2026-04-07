# hyper-denpa
SharePointから時間割変更Excelを引っこ抜いて整形してDiscordに送ったりCLIとしてOpenClawのSkillとして使えるようにする試み

## Workspace

```text
crates/
  hyper-denpa-sharepoint
  hyper-denpa-discord
```

- `hyper-denpa-sharepoint`: SharePoint からファイルを取得・保存する共通ライブラリ
- `hyper-denpa-discord`: Discord Bot

## Auth

認証情報は `.env` から読み込みます．次のすべてを設定するか：

```env
RUST_LOG=warn,hyper_denpa_sharepoint=debug,hyper_denpa_discord=debug
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
cargo run -p hyper-denpa-discord
```

Discord Bot の slash command:

- `/reload`: 最新の時間割変更を再取得
- `/show`: 最後に取得した時間割変更を日付順に表示
- `/grep`: 学年を数値 choice で選び、`CN / ES / IT / 1組 / 2組 / 3組` から必要なら絞り込んで表示
- `/law-csv`: 最新の生 CSV を添付
- `/set-notify`: 実行したチャンネルを定期通知先に追加
- `/unset-notify`: 指定チャンネル、または実行したチャンネルを通知先から解除

取得データは `data/YYYY-MM-DD_HH-MM-SS/` 単位で保存され、`/show` と `/grep` は最新の取得結果を参照します。

Bot のログレベルは `.env` の `RUST_LOG` で切り替えられます。例:

```env
RUST_LOG=warn,hyper_denpa_sharepoint=debug,hyper_denpa_discord=debug
```
