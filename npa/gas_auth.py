"""
Google OAuth2 認証モジュール

GAS Web App（組織内限定）にアクセスするため、
ブラウザでGoogleログインを行いアクセストークンを取得する。

初回セットアップ:
  1. Google Cloud Console → APIとサービス → 認証情報
  2. OAuth 2.0 クライアントID を作成（種類: デスクトップアプリ）
  3. JSONをダウンロード → npa/credentials.json として保存
  4. python npa/gas_auth.py を実行してブラウザでログイン
"""

import json
import os
from pathlib import Path

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# GAS Web Appにアクセスするために必要なスコープ
SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
]

_DIR = Path(__file__).parent
TOKEN_PATH = _DIR / "token.json"
CREDENTIALS_PATH = _DIR / "credentials.json"


def get_credentials() -> Credentials:
    """
    有効なGoogle認証情報を返す。

    - 保存済みトークンがあればそれを使う（期限切れなら自動更新）
    - なければブラウザでOAuth2ログインフローを実行
    """
    creds = None

    # 保存済みトークンを読み込む
    if TOKEN_PATH.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)

    # トークンが無効なら更新 or 再認証
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_PATH.exists():
                raise FileNotFoundError(
                    "OAuth認証情報ファイルが見つかりません。\n\n"
                    "セットアップ手順:\n"
                    "  1. Google Cloud Console → APIとサービス → 認証情報\n"
                    "  2. 「認証情報を作成」→「OAuth クライアント ID」\n"
                    "     アプリの種類: デスクトップ アプリ\n"
                    "  3. JSONをダウンロード\n"
                    f"  4. ファイルを {CREDENTIALS_PATH} として保存\n"
                    "  5. python npa/gas_auth.py を実行してログイン"
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(CREDENTIALS_PATH), SCOPES
            )
            creds = flow.run_local_server(port=0)

        # トークンを保存
        TOKEN_PATH.write_text(creds.to_json())

    return creds


def get_access_token() -> str:
    """有効なアクセストークン文字列を返す"""
    return get_credentials().token


# ── CLI: 初回認証用 ──
if __name__ == "__main__":
    print("Google OAuth2 認証を開始します...")
    print("ブラウザが開きます。Googleアカウントでログインしてください。\n")
    try:
        token = get_access_token()
        print(f"\n認証成功！トークンを保存しました: {TOKEN_PATH}")
        print("Streamlitアプリを起動できます: streamlit run npa/app.py")
    except FileNotFoundError as e:
        print(f"\nエラー: {e}")
    except Exception as e:
        print(f"\n認証エラー: {e}")
