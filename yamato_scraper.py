"""
ヤマトビジネスメンバーズ Web請求書 スクレイパー
- 一覧ページを全件巡回
- 各原票の詳細画面から「部門名1・部門名2」を取得
- CSVと結合してExcelで出力
"""

import time
import os
import re
from datetime import datetime
from typing import Optional, Callable

import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from config import YAMATO_CONFIG


class YamatoScraper:
    def __init__(self, progress_callback: Optional[Callable] = None):
        """
        progress_callback: Streamlitのst.progressなどに渡すコールバック
                           callback(current, total, message) の形で呼ばれる
        """
        self.cfg = YAMATO_CONFIG
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        })
        self.progress_callback = progress_callback or (lambda c, t, m: None)
        self._logged_in = False

    # ------------------------------------------------------------------ #
    # ログイン
    # ------------------------------------------------------------------ #
    def login(self) -> bool:
        """ヤマトビジネスメンバーズにログインする"""
        base = self.cfg["base_url"]

        # 1. ログインページ取得（CSRFトークン等を拾う）
        login_page_url = f"{base}/seikyu/login"
        resp = self.session.get(login_page_url, timeout=30)
        resp.raise_for_status()

        soup = BeautifulSoup(resp.text, "html.parser")

        # hidden フィールドを全部 POST データに含める
        payload = {}
        for inp in soup.select("input[type=hidden]"):
            if inp.get("name"):
                payload[inp["name"]] = inp.get("value", "")

        payload["loginId"] = self.cfg["login_id"]
        payload["password"] = self.cfg["password"]

        # 2. ログイン POST
        login_post_url = f"{base}/seikyu/login"
        resp = self.session.post(login_post_url, data=payload, timeout=30)
        resp.raise_for_status()

        # ログイン成功判定（メニューページへのリダイレクト等）
        if "ログアウト" in resp.text or "メインメニュー" in resp.text:
            self._logged_in = True
            return True

        raise RuntimeError("ログインに失敗しました。ID・パスワードを確認してください。")

    # ------------------------------------------------------------------ #
    # 一覧取得（ページング）
    # ------------------------------------------------------------------ #
    def fetch_list(
        self,
        date_from: str,
        date_to: str,
        use_receipt_date: bool = True,
    ) -> list[dict]:
        """
        一覧ページを全件巡回してレコードを返す。

        Parameters
        ----------
        date_from : str  "20260501" 形式
        date_to   : str  "20260528" 形式
        use_receipt_date : bool  True=受付日, False=請求月日
        """
        if not self._logged_in:
            self.login()

        base = self.cfg["base_url"]
        service_url = f"{base}/seikyu/service"
        interval = self.cfg["request_interval_sec"]

        all_records = []
        page = 1

        while True:
            self.progress_callback(0, 0, f"一覧取得中 {page}ページ目...")

            # 検索実行（初回）または次ページ取得
            if page == 1:
                params = {
                    "uketukeFrom": date_from,
                    "uketsukeTo": date_to,
                    "searchType": "2" if use_receipt_date else "1",
                    "dispCount": "100",   # 1ページ最大表示件数
                    "page": "1",
                    "execBtn": "実行",
                }
                resp = self.session.post(service_url, data=params, timeout=30)
            else:
                params = {
                    "uketukeFrom": date_from,
                    "uketsukeTo": date_to,
                    "searchType": "2" if use_receipt_date else "1",
                    "dispCount": "100",
                    "page": str(page),
                }
                resp = self.session.post(service_url, data=params, timeout=30)

            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")

            # テーブル行を解析
            rows = self._parse_list_rows(soup)
            if not rows:
                break

            all_records.extend(rows)

            # 次ページ存在チェック
            if not self._has_next_page(soup):
                break

            page += 1
            time.sleep(interval)

        return all_records

    def _parse_list_rows(self, soup: BeautifulSoup) -> list[dict]:
        """一覧テーブルの行データを辞書リストで返す"""
        records = []
        table = soup.select_one("table.list, table#listTable, table")
        if not table:
            return records

        rows = table.select("tr")
        for row in rows:
            cols = row.select("td")
            if len(cols) < 9:
                continue

            # 原票Noリンクから detail_url を取得
            link = row.select_one("a[href]")
            detail_url = link["href"] if link else ""

            try:
                record = {
                    "No":           cols[0].get_text(strip=True),
                    "受付日":       cols[1].get_text(strip=True),
                    "原票No":       cols[2].get_text(strip=True),
                    "商品区分":     cols[3].get_text(strip=True),
                    "サイズ":       cols[4].get_text(strip=True),
                    "個数":         cols[5].get_text(strip=True),
                    "都道府県":     cols[6].get_text(strip=True) if len(cols) > 6 else "",
                    "市区町村":     cols[7].get_text(strip=True) if len(cols) > 7 else "",
                    "運賃":         cols[8].get_text(strip=True) if len(cols) > 8 else "",
                    "立替金":       cols[9].get_text(strip=True) if len(cols) > 9 else "",
                    "保険料":       cols[10].get_text(strip=True) if len(cols) > 10 else "",
                    "運賃等合計":   cols[11].get_text(strip=True) if len(cols) > 11 else "",
                    "_detail_url":  detail_url,
                    # 詳細取得後に埋める
                    "部門名1":      "",
                    "部門名2":      "",
                }
                if record["原票No"]:  # ヘッダー行を除く
                    records.append(record)
            except IndexError:
                continue

        return records

    def _has_next_page(self, soup: BeautifulSoup) -> bool:
        """「次へ」リンクが存在するか"""
        next_link = soup.find("a", string=re.compile(r"次[へ>]?"))
        return next_link is not None

    # ------------------------------------------------------------------ #
    # 詳細取得（部門名1・2）
    # ------------------------------------------------------------------ #
    def enrich_with_details(self, records: list[dict]) -> list[dict]:
        """
        各レコードの詳細画面を開いて部門名1・部門名2を補完する。
        """
        base = self.cfg["base_url"]
        interval = self.cfg["request_interval_sec"]
        total = len(records)

        for i, rec in enumerate(records):
            self.progress_callback(i + 1, total, f"詳細取得中 {i+1}/{total}件")

            detail_path = rec.get("_detail_url", "")
            if not detail_path:
                continue

            # 相対URLを絶対URLに
            if detail_path.startswith("http"):
                url = detail_path
            else:
                url = f"{base}{detail_path}"

            try:
                resp = self.session.get(url, timeout=30)
                resp.raise_for_status()
                soup = BeautifulSoup(resp.text, "html.parser")

                rec["部門名1"] = self._extract_field(soup, "部門名1")
                rec["部門名2"] = self._extract_field(soup, "部門名2")

            except Exception as e:
                rec["部門名1"] = f"[取得失敗: {e}]"
                rec["部門名2"] = ""

            time.sleep(interval)

        return records

    def _extract_field(self, soup: BeautifulSoup, label: str) -> str:
        """ラベル文字列に対応するtd値を返す"""
        for th in soup.find_all(["th", "td"]):
            if label in th.get_text():
                next_td = th.find_next_sibling("td")
                if next_td:
                    return next_td.get_text(strip=True)
        return ""

    # ------------------------------------------------------------------ #
    # Excel出力
    # ------------------------------------------------------------------ #
    def export_to_excel(self, records: list[dict], date_from: str, date_to: str) -> str:
        """
        レコードをExcelファイルに出力してパスを返す。
        """
        os.makedirs(self.cfg["output_dir"], exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"yamato_seikyu_{date_from}_{date_to}_{timestamp}.xlsx"
        filepath = os.path.join(self.cfg["output_dir"], filename)

        # DataFrameに変換（内部フィールド除外）
        df = pd.DataFrame(records)
        drop_cols = [c for c in df.columns if c.startswith("_")]
        df = df.drop(columns=drop_cols)

        # 列の並び順を整理
        ordered_cols = [
            "受付日", "原票No", "商品区分", "サイズ", "個数",
            "都道府県", "市区町村",
            "部門名1", "部門名2",
            "運賃", "立替金", "保険料", "運賃等合計",
        ]
        existing_cols = [c for c in ordered_cols if c in df.columns]
        other_cols = [c for c in df.columns if c not in ordered_cols and c != "No"]
        df = df[existing_cols + other_cols]

        # 数値列を整数に変換
        for col in ["運賃", "立替金", "保険料", "運賃等合計", "個数"]:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(",", ""), errors="coerce"
                ).fillna(0).astype(int)

        # Excel書き込み
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="運賃情報")
            ws = writer.sheets["運賃情報"]
            self._style_worksheet(ws, df)

        return filepath

    def _style_worksheet(self, ws, df: pd.DataFrame):
        """Excelシートにスタイルを適用"""
        header_fill = PatternFill("solid", fgColor="1A5276")
        header_font = Font(color="FFFFFF", bold=True, size=10)
        dept_fill   = PatternFill("solid", fgColor="D6EAF8")
        border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        # データ行
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center")
                # 部門名列を色付け
                col_name = df.columns[cell.column - 1] if cell.column <= len(df.columns) else ""
                if col_name in ("部門名1", "部門名2"):
                    cell.fill = dept_fill

        # 列幅の自動調整
        for col_idx, col_name in enumerate(df.columns, start=1):
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).str.len().max() if len(df) > 0 else 0,
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)

        # 先頭行を固定
        ws.freeze_panes = "A2"
