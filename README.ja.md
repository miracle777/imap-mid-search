[![English](https://img.shields.io/badge/README-English-black)](README.md)
[![日本語](https://img.shields.io/badge/README-日本語-blue)](README.ja.md)

# 日本語版 README

## IMAP Message-ID Search (Python, 標準ライブラリのみ)

IMAP メールボックスを RFC822 **Message-ID** で検索し、結果を CSV にエクスポートするツールです。

- 外部ライブラリ不要（Python 標準ライブラリのみ）
- メッセージ ID の `< >` は自動処理
- `Message-ID` / `Message-Id` / `References` / `In-Reply-To` すべてに対応
- オプションで「深掘り検索」：同日時 ±1日 + 送信元ドメインから候補を抽出し、ヘッダ検証
- 特定フォルダや、全フォルダ（`--mailboxes *`）を対象に検索可能

---

## 使い方

### 単一の ID を検索する場合

```bash
python3 imap_mid_search.py \
  --host imap.example.com --user info@example.com --password 'YOUR_PASSWORD' \
  --mailboxes INBOX Sent Trash \
  --ids 20240213212126.4429A161827048B0@gmail.com
```

# ファイルから複数の ID を検索する場合
```bash
python3 imap_mid_search.py \
  --host imap.example.com --user info@example.com --password 'YOUR_PASSWORD' \
  --mailboxes * \
  --ids-file ids.txt --out result.csv
```

出力 CSV のカラム:
```
message_id, mailbox, seqnum, from, to, subject, date
```
# プロバイダー設定方法

任意のサーバー情報を辞書形式で登録しておくと便利です。
以下はサンプル（***** 部分を任意のホスト名に置き換えてください）:
```python
IMAP_CONFIGS = {
    "gmail": {
        "server": "imap.gmail.com",
        "port": 993,
    },
    "outlook": {
        "server": "outlook.office365.com",
        "port": 993,
    },
    "custom": {
        "server": "*****.your-mail-server.ne.jp",  # 任意のサーバーに差し替え
        "port": 993,
    },
}

```
# セキュリティ

- パスワードは環境変数または対話入力を推奨（シェル履歴に残さない）
- メールボックスは 読み取り専用 で選択されます
- データは IMAP サーバー以外に送信されず、ローカルの CSV にのみ保存されます

# Windows (PowerShell)
```
$env:IMAP_HOST="imap.example.com"
$env:IMAP_USER="info@example.com"
$env:IMAP_PASS="YOUR_PASSWORD"

python .\imap_mid_search.py --mailboxes * --ids-file .\ids.txt

```

# ライセンス

MIT License
