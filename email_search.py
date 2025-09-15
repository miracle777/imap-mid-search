# -*- coding: utf-8 -*-
"""
email_search.py (interactive, provider dictionary + deep Message-ID search)
- 既存の対話メニューを維持
- プロバイダー辞書(IMAP_CONFIGS)を同梱（サーバーは ***** で伏せ字）
- providers.json があれば辞書を上書き（顧客固有情報はコード外に）
- 角括弧の自動付与 / 大小文字ゆらぎ / References / In-Reply-To / 日付±1日 の深掘り検索
- \Noselect を除外、select は read-only、安全に全フォルダ探索
"""

import imaplib
import email
from email.header import decode_header
import getpass
from typing import List, Optional, Tuple
from datetime import datetime, timedelta
import json
import os
import re

# ====== ここが「公開用のプレースホルダ辞書」 ======
IMAP_CONFIGS = {
    # 例: 共通プロバイダー（汎用）
    "gmail":   {"server": "imap.gmail.com",          "port": 993},
    "outlook": {"server": "outlook.office365.com",   "port": 993},
    "yahoo":   {"server": "imap.mail.yahoo.com",     "port": 993},

    # 例: カスタム（伏せ字） — 実運用では providers.json で上書きしてください
    "custom1": {"server": "*****.your-mail-server.ne.jp", "port": 993},
    "custom2": {"server": "mail.*****.example.net",       "port": 993},
}
PROVIDERS_JSON = os.path.join(os.path.dirname(__file__), "providers.json")
# providers.json の例：
# {
#   "mycorp": {"server": "mail1234.mycorp.example.com", "port": 993},
#   "legacy": {"server": "mx.legacy-host.example.jp",   "port": 993}
# }

def load_external_providers():
    """providers.json があれば IMAP_CONFIGS を上書き/追加"""
    if not os.path.isfile(PROVIDERS_JSON):
        return
    try:
        with open(PROVIDERS_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            IMAP_CONFIGS.update(data)
    except Exception:
        # 失敗しても致命的ではないので無視（必要ならログ出力に変更）
        pass


class IMAPEmailSearcher:
    def __init__(self, server: str, port: int = 993):
        self.server = server
        self.port = port
        self.connection: Optional[imaplib.IMAP4_SSL] = None
        self.current_mailbox: Optional[str] = None

    def connect(self, username: str, password: str) -> bool:
        try:
            self.connection = imaplib.IMAP4_SSL(self.server, self.port)
            self.connection.login(username, password)
            print(f"✅ {self.server} に正常に接続しました")
            return True
        except imaplib.IMAP4.error as e:
            print(f"❌ IMAP接続エラー: {e}")
            return False
        except Exception as e:
            print(f"❌ 接続エラー: {e}")
            return False

    def list_mailboxes(self) -> List[str]:
        """SELECT 可能なメールボックスのみ返す（\\Noselect を除外）"""
        if not self.connection:
            print("❌ 接続されていません")
            return []
        try:
            status, mailboxes = self.connection.list()
            mailbox_list: List[str] = []
            if status == "OK" and mailboxes:
                for mailbox in mailboxes:
                    line = mailbox.decode(errors="ignore")
                    # 例: (\\HasNoChildren \\Noselect) "." "Archives"
                    flags_part = line.split(")")[0] if ")" in line else ""
                    if "\\Noselect" in flags_part or "\\NoSelect" in flags_part:
                        continue  # 選択不可は除外
                    if '"' in line:
                        name = line.split('"')[-2]
                    else:
                        name = line.split()[-1]
                    mailbox_list.append(name)
            return mailbox_list
        except Exception as e:
            print(f"❌ メールボックス取得エラー: {e}")
            return []

    def select_mailbox(self, mailbox: str = "INBOX") -> bool:
        """read-only で SELECT。成功時のみ current_mailbox を更新"""
        if not self.connection:
            print("❌ 接続されていません")
            return False
        try:
            status, _ = self.connection.select(mailbox, readonly=True)
            if status == 'OK':
                self.current_mailbox = mailbox
                print(f"✅ メールボックス '{mailbox}' を選択しました")
                return True
            else:
                print(f"❌ メールボックス選択失敗: {status}")
                return False
        except Exception as e:
            print(f"❌ メールボックス選択エラー: {e}")
            return False

    def ensure_selected(self) -> bool:
        """SEARCH直前に必ず SELECTED 状態にしておく"""
        if not self.connection:
            print("❌ 接続されていません")
            return False
        if self.current_mailbox:
            return True
        return self.select_mailbox("INBOX")

    # ===== Message-ID検索の堅牢化 =====
    def _normalize_message_id(self, message_id: str) -> Tuple[str, str]:
        mid = message_id.strip().strip("<>").strip()
        return mid, f"<{mid}>"

    def _search_by_message_id_once(self, normalized_value: str) -> List[bytes]:
        """サーバー差吸収のため複数パターンで SEARCH"""
        if not self.ensure_selected():
            return []
        assert self.connection is not None
        c = self.connection
        value = normalized_value
        tries = [
            lambda: c.search(None, 'HEADER', 'Message-ID', value),
            lambda: c.search(None, f'(HEADER "Message-ID" "{value}")'),
            lambda: c.search(None, 'HEADER', 'Message-Id', value),
            lambda: c.search(None, f'(HEADER "Message-Id" "{value}")'),
        ]
        for t in tries:
            status, nums = t()
            if status == 'OK' and nums and nums[0]:
                return nums[0].split()
        return []

    def _search_header_generic(self, header_name: str, value: str) -> List[bytes]:
        if not self.ensure_selected():
            return []
        assert self.connection is not None
        c = self.connection
        tries = [
            lambda: c.search(None, 'HEADER', header_name, value),
            lambda: c.search(None, f'(HEADER "{header_name}" "{value}")')
        ]
        for t in tries:
            status, nums = t()
            if status == "OK" and nums and nums[0]:
                return nums[0].split()
        return []

    def _parse_mid_timestamp(self, message_id: str) -> Optional[datetime]:
        s = message_id.strip().strip("<>").strip()
        m = re.match(r"^(\d{14})[.\-@]", s)
        if not m:
            return None
        try:
            return datetime.strptime(m.group(1), "%Y%m%d%H%M%S")
        except ValueError:
            return None

    def _imap_date(self, dt: datetime) -> str:
        return dt.strftime("%d-%b-%Y")

    def search_by_message_id(self, message_id: str) -> Optional[dict]:
        if not self.connection:
            print("❌ 接続されていません")
            return None
        if not self.ensure_selected():
            print("❌ メールボックス未選択のため検索できませんでした")
            return None
        try:
            bare, with_br = self._normalize_message_id(message_id)
            nums = self._search_by_message_id_once(with_br)
            if not nums:
                nums = self._search_by_message_id_once(bare)
            if not nums:
                print(f"❌ メッセージID '{message_id}' が見つかりませんでした（{self.current_mailbox or '未選択'}）")
                return None
            return self._fetch_email_details(nums[0].decode())
        except Exception as e:
            print(f"❌ 検索エラー: {e}")
            return None

    def search_by_message_id_deep(self, message_id: str, from_domain_hint: Optional[str] = None) -> Optional[dict]:
        # 1) 通常検索
        r = self.search_by_message_id(message_id)
        if r:
            return r

        # 2) References / In-Reply-To
        bare = message_id.strip().strip("<>").strip()
        for h in ("References", "In-Reply-To"):
            for val in (f"<{bare}>", bare):
                nums = self._search_header_generic(h, val)
                if nums:
                    return self._fetch_email_details(nums[0].decode())

        # 3) 日時±1日 → 候補抽出 → ヘッダ直接比較
        ts = self._parse_mid_timestamp(message_id)
        if ts and self.ensure_selected():
            assert self.connection is not None
            since = self._imap_date(ts - timedelta(days=1))
            before = self._imap_date(ts + timedelta(days=1))
            status, nums_all = self.connection.search(None, 'SINCE', since, 'BEFORE', before)
            if status == "OK" and nums_all and nums_all[0]:
                candidates = nums_all[0].split()
                pool = candidates
                # 差出人ドメインで軽絞り
                if from_domain_hint:
                    filtered = []
                    for n in candidates:
                        info = self._fetch_email_details(n.decode())
                        if info and from_domain_hint.lower() in info['from'].lower():
                            filtered.append(n)
                    if filtered:
                        pool = filtered
                for n in pool:
                    status, msg_data = self.connection.fetch(n, '(BODY[HEADER.FIELDS (MESSAGE-ID)])')
                    if status == 'OK' and msg_data and msg_data[0] is not None:
                        hdr = msg_data[0][1].decode('utf-8', errors='ignore')
                        if bare in hdr or f"<{bare}>" in hdr:
                            return self._fetch_email_details(n.decode())

        print("❌ 深掘り検索でも見つかりませんでした")
        return None

    def _fetch_email_details(self, msg_num: str) -> Optional[dict]:
        try:
            assert self.connection is not None
            status, msg_data = self.connection.fetch(msg_num, '(RFC822)')
            if status != 'OK' or not msg_data or not msg_data[0]:
                return None
            msg = email.message_from_bytes(msg_data[0][1])
            subject = self._decode_header(msg.get('Subject', ''))
            from_addr = self._decode_header(msg.get('From', ''))
            to_addr = self._decode_header(msg.get('To', ''))
            date = msg.get('Date', '')
            message_id = msg.get('Message-ID', '')
            body = self._get_email_body(msg)
            return {
                'message_number': msg_num,
                'message_id': message_id,
                'subject': subject,
                'from': from_addr,
                'to': to_addr,
                'date': date,
                'body': body[:200] + '...' if len(body) > 200 else body
            }
        except Exception as e:
            print(f"❌ メール取得エラー: {e}")
            return None

    def _decode_header(self, header: str) -> str:
        if not header:
            return ""
        decoded_parts = decode_header(header)
        decoded_str = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                decoded_str += part.decode(encoding or 'utf-8', errors='ignore')
            else:
                decoded_str += part
        return decoded_str

    def _get_email_body(self, msg) -> str:
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        body = (part.get_payload(decode=True) or b"").decode(charset, errors='ignore')
                    except Exception:
                        body = (part.get_payload(decode=True) or b"").decode('utf-8', errors='ignore')
                    break
        else:
            charset = msg.get_content_charset() or 'utf-8'
            try:
                body = (msg.get_payload(decode=True) or b"").decode(charset, errors='ignore')
            except Exception:
                body = (msg.get_payload(decode=True) or b"").decode('utf-8', errors='ignore')
        return body

    def disconnect(self):
        if self.connection:
            try:
                try:
                    self.connection.close()
                except Exception:
                    pass
                self.connection.logout()
                print("✅ 接続を終了しました")
            except Exception:
                pass


def main():
    # 外部providers.jsonがあれば読み込み（顧客固有の接続先はここで管理）
    load_external_providers()

    print("=== IMAP メール検索プログラム ===\n")
    print("メールプロバイダーを選択してください（辞書から選択 or 'manual' で手入力）:")
    keys = list(IMAP_CONFIGS.keys())
    for k in keys:
        print(f"- {k}")
    print("- manual  （ホスト名/ポートを手入力）")

    provider_input = input("\nプロバイダー名を入力: ").strip()
    if provider_input.lower() == "manual":
        server = input("IMAPサーバー（例: imap.example.com）: ").strip()
        port_s = input("ポート（デフォルト 993）: ").strip() or "993"
        try:
            port = int(port_s)
        except ValueError:
            print("❌ ポート番号が不正です"); return
        config = {"server": server, "port": port}
    else:
        # 大文字小文字吸収
        normalized_map = {str(k).lower(): k for k in IMAP_CONFIGS.keys()}
        key = normalized_map.get(provider_input.lower())
        if key is None:
            print("❌ サポートされていないプロバイダーです")
            return
        config = IMAP_CONFIGS[key]

    username = input("メールアドレス: ").strip()
    password = getpass.getpass("パスワード: ")

    searcher = IMAPEmailSearcher(config['server'], int(config['port']))
    try:
        if not searcher.connect(username, password):
            return
        if not searcher.select_mailbox("INBOX"):
            return

        while True:
            print("\n=== 検索メニュー ===")
            print("1. メッセージIDで検索（見つからなければ深掘り+他フォルダ探索）")
            print("2. 送信者で検索")
            print("3. 件名で検索")
            print("4. メールボックス一覧表示")
            print("5. 終了")

            choice = input("\n選択してください (1-5): ").strip()

            if choice == '1':
                message_id = input("検索するメッセージID（< > なしでOK）: ").strip()
                # 現在のボックス（INBOX）
                result = searcher.search_by_message_id(message_id)

                # 見つからなければ、SELECT可能ボックスで総当たり
                if not result:
                    selectable = searcher.list_mailboxes()
                    others = [b for b in selectable if b != "INBOX"]
                    priority = ["Trash", "Junk", "Spam", "Sent", "Drafts", "Archive",
                                "[Gmail]/All Mail", "[Gmail]/Trash", "[Gmail]/Spam",
                                "[Gmail]/Sent Mail", "[Gmail]/Drafts"]
                    prioritized = [b for b in priority if b in others]
                    remaining   = [b for b in others if b not in prioritized]
                    scan_order  = prioritized + remaining
                    for m in scan_order:
                        if not searcher.select_mailbox(m):
                            continue
                        result = searcher.search_by_message_id(message_id)
                        if result:
                            result['mailbox'] = m
                            break

                # それでもダメなら深掘り（全フォルダ）
                if not result:
                    all_boxes = searcher.list_mailboxes()
                    if all_boxes:
                        # ドメインヒント（任意）
                        hint = None
                        if "@gmail.com" in message_id:
                            hint = "gmail.com"
                        elif "@cpanel.net" in message_id:
                            hint = "cpanel.net"
                        for m in all_boxes:
                            if not searcher.select_mailbox(m):
                                continue
                            result = searcher.search_by_message_id_deep(message_id, from_domain_hint=hint)
                            if result:
                                result['mailbox'] = m
                                break

                if result:
                    print(f"\n✅ メールが見つかりました（{result.get('mailbox', searcher.current_mailbox) or 'N/A'}）:")
                    print(f"件名: {result['subject']}")
                    print(f"送信者: {result['from']}")
                    print(f"日時: {result['date']}")
                    print(f"本文: {result['body']}")
                else:
                    print("\n❌ すべての手段で見つかりませんでした。")
                    print("  - Message-ID が書き換わっている可能性があります。")
                    print("  - 同日時±1日の候補も検索しましたが該当なし。")
                    print("  - 別アカウント/別サーバー側に存在しないかをご確認ください。")

            elif choice == '2':
                sender = input("送信者のメールアドレス: ").strip()
                results = searcher.search_emails(f'FROM "{sender}"')
                print(f"\n✅ {len(results)}件のメールが見つかりました:")
                for i, email_info in enumerate(results, 1):
                    print(f"\n{i}. 件名: {email_info['subject']}")
                    print(f"   日時: {email_info['date']}")

            elif choice == '3':
                subject = input("検索する件名: ").strip()
                results = searcher.search_emails(f'SUBJECT "{subject}"')
                print(f"\n✅ {len(results)}件のメールが見つかりました:")
                for i, email_info in enumerate(results, 1):
                    print(f"\n{i}. 件名: {email_info['subject']}")
                    print(f"   送信者: {email_info['from']}")
                    print(f"   日時: {email_info['date']}")

            elif choice == '4':
                mailboxes = searcher.list_mailboxes()
                print("\n✅ 利用可能なメールボックス:")
                for mailbox in mailboxes:
                    print(f"- {mailbox}")

            elif choice == '5':
                break
            else:
                print("❌ 無効な選択です")

    finally:
        searcher.disconnect()


if __name__ == "__main__":
    main()
