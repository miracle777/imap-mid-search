[![English](https://img.shields.io/badge/README-English-blue)](README.md)
[![日本語](https://img.shields.io/badge/README-日本語-black)](README.ja.md)


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
```

### Provider dictionary override (providers.json)

You can keep customer-specific server names **outside the code**.

If a `providers.json` file exists next to `email_search.py`, its entries will **extend/override** the built-in `IMAP_CONFIGS`.

**providers.json example:**
```json
{
  "mycorp": { "server": "mail1234.mycorp.example.com", "port": 993 },
  "legacy": { "server": "mx.legacy-host.example.jp",   "port": 993 }
}
```
In the interactive menu, just type the provider key (e.g., mycorp).
Or choose manual to input host/port by hand.

If providers.json does not exist, the program still works with the built-in defaults.


## Security

Prefer environment variables or interactive prompt for passwords (avoid shell history).

This tool selects mailboxes readonly.

No data is sent anywhere except your IMAP server; results are written to local CSV.

## Windows (PowerShell)
```
$env:IMAP_HOST="imap.example.com"
$env:IMAP_USER="info@example.com"
$env:IMAP_PASS="YOUR_PASSWORD"

python .\imap_mid_search.py --mailboxes * --ids-file .\ids.txt
```


## MIT License

Copyright (c) 2025 miracle777

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
