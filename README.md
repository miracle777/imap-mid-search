# IMAP Message-ID Search (Python, stdlib only)

Searches an IMAP mailbox by RFC822 **Message-ID** and exports matches to CSV.  
- No external deps (only Python stdlib)
- Angle brackets `< >` auto-handled
- Falls back across `Message-ID`/`Message-Id`, `References`, `In-Reply-To`
- Optional **deep search**: time-window ±1 day + sender-domain hint → header verification
- Search selected folders or **all selectable folders** (`--mailboxes *`)

---

## Usage

```bash
python3 imap_mid_search.py \
  --host imap.example.com --user info@example.com --password 'YOUR_PASSWORD' \
  --mailboxes INBOX Sent Trash \
  --ids 20240213212126.4429A161827048B0@gmail.com

## Security

Prefer environment variables or interactive prompt for passwords (avoid shell history).

This tool selects mailboxes readonly.

No data is sent anywhere except your IMAP server; results are written to local CSV.

## Windows (PowerShell)
$env:IMAP_HOST="imap.example.com"
$env:IMAP_USER="info@example.com"
$env:IMAP_PASS="YOUR_PASSWORD"

python .\imap_mid_search.py --mailboxes * --ids-file .\ids.txt
