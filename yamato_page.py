"""
ヤマト運賃情報取得ページ
既存のStreamlitアプリに pages/ として追加するか、
単体で `streamlit run yamato_page.py` で起動して使用する。
"""

import streamlit as st
from datetime import date, timedelta
import threading
import time as _time

from yamato_scraper import YamatoScraper


# ------------------------------------------------------------ #
# ページ設定
# ------------------------------------------------------------ #
st.set_page_config(
    page_title="ヤマト運賃取得",
    page_icon="🚚",
    layout="centered",
)

st.title("🚚 ヤマト運賃情報取得")
st.caption("Web請求書から部門名1・部門名2を含む全データをExcelで出力します")

# ------------------------------------------------------------ #
# 期間入力
# ------------------------------------------------------------ #
st.subheader("① 取得期間を指定")

col1, col2 = st.columns(2)
today = date.today()
default_from = today - timedelta(days=7)

with col1:
    date_from = st.date_input("開始日（受付日）", value=default_from)
with col2:
    date_to = st.date_input("終了日（受付日）", value=today)

# 31日制限チェック
delta_days = (date_to - date_from).days
if delta_days > 31:
    st.error("⚠️ 最大31日間までです。期間を短くしてください。")
    st.stop()
elif delta_days < 0:
    st.error("⚠️ 開始日 > 終了日になっています。")
    st.stop()

st.info(f"📅 取得対象：{date_from} ～ {date_to}（{delta_days + 1}日間）")

# ------------------------------------------------------------ #
# 推定時間の表示
# ------------------------------------------------------------ #
st.subheader("② 実行")

with st.expander("⏱️ 推定時間の目安"):
    st.markdown("""
| 件数 | 推定時間 |
|------|---------|
| 50件 | 約2〜3分 |
| 100件 | 約4〜5分 |
| 200件 | 約7〜10分 |
| 500件 | 約17〜25分 |

※ サーバー負荷軽減のため1件ごとに2秒待機しています
""")

# ------------------------------------------------------------ #
# 実行ボタン
# ------------------------------------------------------------ #
if st.button("▶ 取得開始", type="primary", use_container_width=True):

    date_from_str = date_from.strftime("%Y%m%d")
    date_to_str   = date_to.strftime("%Y%m%d")

    # プログレス表示の初期化
    progress_bar  = st.progress(0)
    status_text   = st.empty()
    log_container = st.empty()
    logs = []

    def on_progress(current, total, message):
        logs.append(message)
        log_container.text("\n".join(logs[-8:]))  # 直近8件を表示
        if total > 0:
            pct = int(current / total * 100)
            progress_bar.progress(pct)
            status_text.markdown(f"**{message}** （{current}/{total}件）")
        else:
            status_text.markdown(f"**{message}**")

    try:
        scraper = YamatoScraper(progress_callback=on_progress)

        # Step 1: ログイン
        status_text.markdown("**ログイン中...**")
        scraper.login()
        st.toast("✅ ログイン成功")

        # Step 2: 一覧取得
        status_text.markdown("**一覧を取得中...**")
        records = scraper.fetch_list(date_from_str, date_to_str)
        total_count = len(records)

        if total_count == 0:
            st.warning("該当データがありませんでした。期間を確認してください。")
            st.stop()

        st.toast(f"✅ 一覧取得完了：{total_count}件")

        # Step 3: 詳細取得（部門名）
        status_text.markdown(f"**詳細取得中（{total_count}件）...**")
        progress_bar.progress(0)
        records = scraper.enrich_with_details(records)

        # Step 4: Excel出力
        status_text.markdown("**Excelファイルを作成中...**")
        filepath = scraper.export_to_excel(records, date_from_str, date_to_str)

        progress_bar.progress(100)
        status_text.markdown("✅ **完了！**")

        # ダウンロードボタン
        st.success(f"🎉 {total_count}件のデータを取得しました")
        with open(filepath, "rb") as f:
            file_bytes = f.read()

        st.download_button(
            label="📥 Excelをダウンロード",
            data=file_bytes,
            file_name=filepath.split("/")[-1],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except RuntimeError as e:
        st.error(f"❌ エラー：{e}")
    except Exception as e:
        st.error(f"❌ 予期しないエラーが発生しました：{e}")
        st.exception(e)
