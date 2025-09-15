import imaplib
import email
from email.header import decode_header
import getpass
import os
import json
from typing import List, Optional, Dict


class IMAPEmailSearcher:
    def __init__(self, server: str, port: int = 993):
        """
        IMAPメール検索クラス
        Args:
            server: IMAPサーバーのホスト名
            port: ポート番号（デフォルト: 993 for SSL）
        """
        self.server = server
        self.port = port
        self.connection = None

    def connect(self, username: str, password: str) -> bool:
        """IMAPサーバーに接続"""
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
        """利用可能なメールボックス一覧を取得"""
        if not self.connection:
            print("❌ 接続されていません")
            return []
        try:
            status, mailboxes = self.connection.list()
            mailbox_list = []
            for mailbox in mailboxes:
                decoded = mailbox.decode()
                if "\\Noselect" in decoded:
                    continue
                name = decoded.split('"')[-2] if '"' in decoded else decoded.split()[-1]
                mailbox_list.append(name)
            return mailbox_list
        except Exception as e:
            print(f"❌ メールボックス取得エラー: {e}")
            return []

    def select_mailbox(self, mailbox: str = "INBOX") -> bool:
        """メールボックスを選択"""
        if not self.connection:
            print("❌ 接続されていません")
            return False
        try:
            status, _ = self.connection.select(mailbox, readonly=True)
            if status == "OK":
                print(f"✅ メールボックス '{mailbox}' を選択しました")
                return True
            else:
                print(f"❌ メールボックス選択失敗: {status}")
                return False
        except Exception as e:
            print(f"❌ メールボックス選択エラー: {e}")
            return False

    def search_by_message_id(self, message_id: str) -> Optional[dict]:
        """メッセージIDでメールを検索"""
        if not self.connection:
            print("❌ 接続されていません")
            return None
        try:
            search_criteria = f'HEADER Message-ID "{message_id}"'
            status, message_numbers = self.connection.search(None, search_criteria)
            if status != "OK":
                print(f"❌ 検索失敗: {status}")
                return None
            message_nums = message_numbers[0].split()
            if not message_nums:
                print(f"❌ メッセージID '{message_id}' が見つかりませんでした")
                return None
            msg_num = message_nums[0]
            return self._fetch_email_details(msg_num.decode())
        except Exception as e:
            print(f"❌ 検索エラー: {e}")
            return None

    def search_emails(self, criteria: str) -> List[dict]:
        """指定された条件でメールを検索"""
        if not self.connection:
            print("❌ 接続されていません")
            return []
        try:
            status, message_numbers = self.connection.search(None, criteria)
            if status != "OK":
                print(f"❌ 検索失敗: {status}")
                return []
            message_nums = message_numbers[0].split()
            emails = []
            for msg_num in message_nums[-10:]:  # 最新10件のみ取得
                email_info = self._fetch_email_details(msg_num.decode())
                if email_info:
                    emails.append(email_info)
            return emails
        except Exception as e:
            print(f"❌ 検索エラー: {e}")
            return []

    def _fetch_email_details(self, msg_num: str) -> Optional[dict]:
        """指定されたメール番号のメール詳細を取得"""
        try:
            status, msg_data = self.connection.fetch(msg_num, "(RFC822)")
            if status != "OK":
                return None
            msg = email.message_from_bytes(msg_data[0][1])
            subject = self._decode_header(msg.get("Subject", ""))
            from_addr = self._decode_header(msg.get("From", ""))
            to_addr = self._decode_header(msg.get("To", ""))
            date = msg.get("Date", "")
            message_id = msg.get("Message-ID", "")
            body = self._get_email_body(msg)
            return {
                "message_number": msg_num,
                "message_id": message_id,
                "subject": subject,
                "from": from_addr,
                "to": to_addr,
                "date": date,
                "body": body[:200] + "..." if len(body) > 200 else body,
            }
        except Exception as e:
            print(f"❌ メール取得エラー: {e}")
            return None

    def _decode_header(self, header: str) -> str:
        """メールヘッダーをデコード"""
        if not header:
            return ""
        decoded_parts = decode_header(header)
        decoded_str = ""
        for part, encoding in decoded_parts:
            if isinstance(part, bytes):
                decoded_str += part.decode(encoding or "utf-8", errors="ignore")
            else:
                decoded_str += part
        return decoded_str

    def _get_email_body(self, msg) -> str:
        """メール本文を取得"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition", ""))
                if content_type == "text/plain" and "attachment" not in content_disposition:
                    charset = part.get_content_charset() or "utf-8"
                    body = part.get_payload(decode=True).decode(charset, errors="ignore")
                    break
        else:
            charset = msg.get_content_charset() or "utf-8"
            body = msg.get_payload(decode=True).decode(charset, errors="ignore")
        return body

    def disconnect(self):
        """接続を終了"""
        if self.connection:
            try:
                self.connection.close()
                self.connection.logout()
                print("✅ 接続を終了しました")
            except:
                pass


def load_providers() -> Dict[str, dict]:
    """
    プロバイダー設定を読み込む
    providers.json が存在すれば読み込み、無ければデフォルトを返す
    """
    default = {
        "gmail": {"server": "imap.gmail.com", "port": 993},
        "outlook": {"server": "outlook.office365.com", "port": 993},
        "yahoo": {"server": "imap.mail.yahoo.com", "port": 993},
        "custom1": {"server": "*****.your-mail-server.ne.jp", "port": 993},
        "custom2": {"server": "*****.example.jp", "port": 993},
    }
    if os.path.exists("providers.json"):
        try:
            with open("providers.json", "r", encoding="utf-8") as f:
                custom = json.load(f)
            default.update(custom)
            print(f"✅ providers.json を読み込みました ({len(custom)} 件追加)")
        except Exception as e:
            print(f"⚠️ providers.json の読み込みに失敗しました: {e}")
    return default


def main():
    print("=== IMAP メール検索プログラム ===\n")

    IMAP_CONFIGS = load_providers()

    print("メールプロバイダーを選択してください:")
    for key in IMAP_CONFIGS.keys():
        print(f"- {key}")
    print("- manual")

    provider = input("\nプロバイダー名を入力: ").lower()

    if provider == "manual":
        host = input("IMAP サーバー: ")
        port = int(input("ポート番号 (既定 993): ") or 993)
        config = {"server": host, "port": port}
    elif provider in IMAP_CONFIGS:
        config = IMAP_CONFIGS[provider]
    else:
        print("❌ サポートされていないプロバイダーです")
        return

    username = os.getenv("IMAP_USER") or input("メールアドレス: ")
    password = os.getenv("IMAP_PASS") or getpass.getpass("パスワード: ")

    searcher = IMAPEmailSearcher(config["server"], config["port"])

    try:
        if not searcher.connect(username, password):
            return
        if not searcher.select_mailbox("INBOX"):
            return

        while True:
            print("\n=== 検索メニュー ===")
            print("1. メッセージIDで検索")
            print("2. 送信者で検索")
            print("3. 件名で検索")
            print("4. メールボックス一覧表示")
            print("5. 終了")

            choice = input("\n選択してください (1-5): ")

            if choice == "1":
                message_id = input("検索するメッセージID: ")
                result = searcher.search_by_message_id(message_id)
                if result:
                    print(f"\n✅ メールが見つかりました:")
                    print(f"件名: {result['subject']}")
                    print(f"送信者: {result['from']}")
                    print(f"日時: {result['date']}")
                    print(f"本文: {result['body']}")

            elif choice == "2":
                sender = input("送信者のメールアドレス: ")
                results = searcher.search_emails(f'FROM "{sender}"')
                print(f"\n✅ {len(results)}件のメールが見つかりました:")
                for i, email_info in enumerate(results, 1):
                    print(f"\n{i}. 件名: {email_info['subject']}")
                    print(f"   日時: {email_info['date']}")

            elif choice == "3":
                subject = input("検索する件名: ")
                results = searcher.search_emails(f'SUBJECT "{subject}"')
                print(f"\n✅ {len(results)}件のメールが見つかりました:")
                for i, email_info in enumerate(results, 1):
                    print(f"\n{i}. 件名: {email_info['subject']}")
                    print(f"   送信者: {email_info['from']}")
                    print(f"   日時: {email_info['date']}")

            elif choice == "4":
                mailboxes = searcher.list_mailboxes()
                print("\n✅ 利用可能なメールボックス:")
                for mailbox in mailboxes:
                    print(f"- {mailbox}")

            elif choice == "5":
                break
            else:
                print("❌ 無効な選択です")

    finally:
        searcher.disconnect()


if __name__ == "__main__":
    main()
