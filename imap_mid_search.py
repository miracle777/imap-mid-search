#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
imap_mid_search.py
Search IMAP mailboxes by RFC822 Message-ID(s) and export matches to CSV.

- 標準ライブラリのみ（requests等の依存なし）
- 認証情報は引数 or 環境変数（IMAP_HOST/IMAP_PORT/IMAP_USER/IMAP_PASS）
- フォルダ指定 or 全フォルダ自動列挙（'*'）
- 角括弧(< >)自動付与、Message-ID 大小文字ゆらぎへフォールバック
- References / In-Reply-To 検索、日時±1日レンジの“深掘り検索”内蔵

Usage:
  python3 imap_mid_search.py --host imap.example.com --user info@example.com --password '***' \
    --mailboxes INBOX Sent Trash --ids 20240213212126.4429A161827048B0@gmail.com

  # 複数IDをファイルで
  python3 imap_mid_search.py --host imap.example.com --user info@example.com --password '***' \
    --mailboxes * --ids-file ids.txt

Security:
  - パスワードは引数ではなく環境変数 IMAP_PASS やプロンプト入力がおすすめ
  - --out で出力先CSVを指定可能（デフォルト: imap_messageid_matches.csv）
"""

import argparse
import csv
import imaplib
import os
import ssl
import sys
import time
import re
from datetime import datetime, timedelta
from typing import List, Tuple, Optional

DEFAULT_MAILBOXES = [
    "INBOX",
    "Trash", "Junk", "Spam",
    "Sent", "Drafts",
    "Archive",
    "[Gmail]/All Mail", "[Gmail]/Trash", "[Gmail]/Spam",
    "[Gmail]/Sent Mail", "[Gmail]/Drafts",
]

def parse_args():
    p = argparse.ArgumentParser(description="Search IMAP by Message-ID and output CSV results")
    p.add_argument("--host", default=os.getenv("IMAP_HOST"), help="IMAP hostname (env IMAP_HOST)")
    p.add_argument("--port", type=int, default=int(os.getenv("IMAP_PORT", "993")), help="IMAPS port (default 993)")
    p.add_argument("--user", default=os.getenv("IMAP_USER"), help="IMAP username (env IMAP_USER)")
    p.add_argument("--password", default=os.getenv("IMAP_PASS") or os.getenv("IMAP_PASSWORD"),
                   help="IMAP password (env IMAP_PASS or IMAP_PASSWORD). If omitted, prompt will appear.")
    p.add_argument("--ids", nargs="*", help="Message-IDs without angle brackets (space-separated)")
    p.add_argument("--ids-file", help="Path to a text file containing Message-IDs (one per line)")
    p.add_argument("--mailboxes", nargs="*", default=None,
                   help="Mailboxes to search. Use '*' to search all selectable mailboxes on server. If omitted, a common set is used.")
    p.add_argument("--timeout", type=int, default=60, help="IMAP socket timeout seconds (default 60)")
    p.add_argument("--out", default="imap_messageid_matches.csv", help="Output CSV path (default: imap_messageid_matches.csv)")
    p.add_argument("--debug", action="store_true", help="Enable imaplib debug output")
    return p.parse_args()

def load_ids(args) -> List[str]:
    ids: List[str] = []
    if args.ids:
        ids.extend(args.ids)
    if args.ids_file:
        with open(args.ids_file, "r", encoding="utf-8") as f:
            for line in f:
                s = line.strip()
                if s:
                    ids.append(s)
    # normalize: remove surrounding < >
    cleaned = []
    for mid in ids:
        s = mid.strip()
        if s.startswith("<") and s.endswith(">"):
            s = s[1:-1].strip()
        cleaned.append(s)
    # de-duplicate preserving order
    seen = set()
    uniq = []
    for x in cleaned:
        if x not in seen:
            uniq.append(x); seen.add(x)
    return uniq

def list_all_mailboxes(M: imaplib.IMAP4_SSL) -> List[str]:
    typ, data = M.list()
    boxes = []
    if typ == "OK" and data:
        for raw in data:
            s = raw.decode("utf-8", errors="ignore")
            # (\\HasNoChildren \\Noselect) "." "Archives"
            flags_part = s.split(")")[0] if ")" in s else ""
            if "\\Noselect" in flags_part or "\\NoSelect" in flags_part:
                continue  # skip non-selectable
            if '"' in s:
                name = s.split('"')[-2]
            else:
                name = s.split()[-1]
            boxes.append(name)
    # unique preserve order
    out, seen = [], set()
    for b in boxes:
        if b not in seen:
            out.append(b); seen.add(b)
    return out

def select_box(M: imaplib.IMAP4_SSL, mailbox: str) -> bool:
    typ, _ = M.select(mailbox, readonly=True)
    return typ == "OK"

def fetch_headers(M: imaplib.IMAP4_SSL, num: bytes) -> Tuple[str, str, str, str]:
    typ, msgdata = M.fetch(num, '(BODY[HEADER.FIELDS (FROM TO SUBJECT DATE MESSAGE-ID)])')
    if typ != "OK" or not msgdata or msgdata[0] is None:
        return "", "", "", ""
    raw = msgdata[0][1].decode("utf-8", errors="replace")
    def pick(prefix: str) -> str:
        for line in raw.splitlines():
            if line.lower().startswith(prefix.lower()):
                return line.split(":",1)[1].strip()
        return ""
    return pick("From"), pick("To"), pick("Subject"), pick("Date")

def normalize_mid(mid: str) -> Tuple[str, str]:
    bare = mid.strip().strip("<>").strip()
    return bare, f"<{bare}>"

def try_search_header(M: imaplib.IMAP4_SSL, key: str, value: str) -> List[bytes]:
    tries = [
        lambda: M.search(None, 'HEADER', key, value),
        lambda: M.search(None, f'(HEADER "{key}" "{value}")'),
    ]
    for t in tries:
        typ, data = t()
        if typ == "OK" and data and data[0]:
            return data[0].split()
    return []

def search_by_mid_in_selected(M: imaplib.IMAP4_SSL, mid: str) -> List[bytes]:
    bare, with_br = normalize_mid(mid)

    # 1) Message-ID（大小文字ゆらぎ含む）
    for key in ("Message-ID", "Message-Id"):
        for val in (with_br, bare):
            nums = try_search_header(M, key, val)
            if nums:
                return nums

    # 2) References / In-Reply-To に含まれていないか
    for key in ("References", "In-Reply-To"):
        for val in (with_br, bare):
            nums = try_search_header(M, key, val)
            if nums:
                return nums
    return []

def parse_mid_timestamp(mid: str) -> Optional[datetime]:
    s = mid.strip().strip("<>").strip()
    m = re.match(r"^(\d{14})[.\-@]", s)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y%m%d%H%M%S")
    except ValueError:
        return None

def imap_date(dt: datetime) -> str:
    return dt.strftime("%d-%b-%Y")

def deep_search_by_mid(M: imaplib.IMAP4_SSL, mid: str, from_domain_hint: Optional[str] = None) -> List[bytes]:
    # まず通常の検索
    nums = search_by_mid_in_selected(M, mid)
    if nums:
        return nums

    # 日時±1日で候補抽出→ヘッダ局所照合
    ts = parse_mid_timestamp(mid)
    if not ts:
        return []
    since, before = imap_date(ts - timedelta(days=1)), imap_date(ts + timedelta(days=1))
    typ, data = M.search(None, 'SINCE', since, 'BEFORE', before)
    if typ != "OK" or not data or not data[0]:
        return []
    candidates = data[0].split()

    pool = candidates
    if from_domain_hint:
        # 差出人ドメインで軽く絞り込み（ヘッダ簡易取得）
        filtered = []
        for n in candidates:
            f, _, _, _ = fetch_headers(M, n)
            if from_domain_hint.lower() in f.lower():
                filtered.append(n)
        if filtered:
            pool = filtered

    # Message-IDのヘッダを直接比較
    bare, with_br = normalize_mid(mid)
    for n in pool:
        typ, msgdata = M.fetch(n, '(BODY[HEADER.FIELDS (MESSAGE-ID)])')
        if typ == "OK" and msgdata and msgdata[0] is not None:
            hdr = msgdata[0][1].decode("utf-8", errors="ignore")
            if bare in hdr or with_br in hdr:
                return [n]
    return []

def main():
    args = parse_args()

    # 入力バリデーション
    missing = [k for k, v in {"--host": args.host, "--user": args.user}.items() if not v]
    if missing:
        sys.stderr.write("Missing required: " + ", ".join(missing) + "\n")
        sys.exit(2)
    if not args.password:
        # 遅延インポート（プロンプトはここでのみ）
        import getpass
        args.password = getpass.getpass("IMAP password: ")

    ids = load_ids(args)
    if not ids:
        sys.stderr.write("No Message-IDs provided via --ids or --ids-file\n")
        sys.exit(2)

    # 接続
    if args.debug:
        imaplib.Debug = 4
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = True
    ssl_ctx.verify_mode = ssl.CERT_REQUIRED
    imaplib._MAXLINE = 10000000
    # socket timeout
    try:
        import socket
        socket.setdefaulttimeout(args.timeout)
    except Exception:
        pass

    print(f"Connecting to {args.host}:{args.port} as {args.user} ...")
    M = imaplib.IMAP4_SSL(args.host, args.port, ssl_context=ssl_ctx)
    try:
        typ, _ = M.login(args.user, args.password)
        if typ != "OK":
            raise RuntimeError("Login failed")
    except imaplib.IMAP4.error as e:
        sys.stderr.write(f"Login error: {e}\n")
        sys.exit(1)

    # メールボックス決定
    if args.mailboxes is None:
        mailboxes = DEFAULT_MAILBOXES
    elif len(args.mailboxes) == 1 and args.mailboxes[0] == "*":
        mailboxes = list_all_mailboxes(M)
        if not mailboxes:
            mailboxes = DEFAULT_MAILBOXES
    else:
        mailboxes = args.mailboxes

    fieldnames = ["message_id", "mailbox", "seqnum", "from", "to", "subject", "date"]
    rows = []
    found_total = 0
    start = time.time()

    try:
        # 最初に INBOX を優先
        scan_order = list(dict.fromkeys(["INBOX"] + mailboxes))
        for mid in ids:
            print(f"\n== Searching Message-ID: <{mid}> ==")
            matched = False
            # ドメインヒント（任意）
            hint = "gmail.com" if "@gmail.com" in mid else ("cpanel.net" if "@cpanel.net" in mid else None)

            for mbox in scan_order:
                if not select_box(M, mbox):
                    continue
                # まず通常検索→なければ深掘り
                nums = search_by_mid_in_selected(M, mid)
                if not nums:
                    nums = deep_search_by_mid(M, mid, from_domain_hint=hint)
                if not nums:
                    continue

                for num in nums:
                    f, t, s, d = fetch_headers(M, num)
                    rows.append({
                        "message_id": mid,
                        "mailbox": mbox,
                        "seqnum": num.decode("ascii", errors="ignore"),
                        "from": f, "to": t, "subject": s, "date": d
                    })
                    print(f"  - {mbox}: seq {num.decode()} | {d} | {s}")
                    found_total += 1
                matched = True
                break  # 1通につき最初に見つかったボックスで確定

            if not matched:
                rows.append({
                    "message_id": mid,
                    "mailbox": "",
                    "seqnum": "",
                    "from": "",
                    "to": "",
                    "subject": "",
                    "date": "",
                })
                print("  (not found)")

    finally:
        try:
            M.logout()
        except Exception:
            pass

    # CSV 出力
    out_path = args.out
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, fieldnames=fieldnames)
        wr.writeheader()
        wr.writerows(rows)

    elapsed = time.time() - start
    print(f"\nDone. Found {found_total} matches in {elapsed:.1f}s. CSV -> {out_path}")

if __name__ == "__main__":
    main()
