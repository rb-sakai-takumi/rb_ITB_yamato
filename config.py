# ヤマトビジネスメンバーズ 設定ファイル
# このファイルはリポジトリに含めないこと（.gitignoreに追加推奨）

YAMATO_CONFIG = {
    # ログイン情報
    "login_id": "YOUR_LOGIN_ID",       # ← ここに変更
    "password": "YOUR_PASSWORD",        # ← ここに変更
    "customer_code": "0358093257",      # お客様コード（画面から確認）

    # スクレイピング設定
    "base_url": "https://webseikyu.kuronekoyamato.co.jp",
    "request_interval_sec": 2.0,        # サーバー負荷軽減のための待機秒数（推奨2秒以上）
    "max_retries": 3,                   # エラー時リトライ回数

    # 出力設定
    "output_dir": "./output",           # Excel出力先ディレクトリ
}
