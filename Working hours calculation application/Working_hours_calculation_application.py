#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import csv
import json
import os
import sys
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, Tuple, List

from PySide6.QtCore import Qt, QDate, QTime
from PySide6.QtGui import QAction, QPixmap, QIcon
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QTabWidget, QLabel, QLineEdit, QTimeEdit, QPushButton, QTextEdit, QSpinBox,
    QDoubleSpinBox, QMessageBox, QCalendarWidget, QComboBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QCheckBox, QFileDialog
)
# シングルインスタンス用（既に起動中のアプリを検出/アクティベート）
from PySide6.QtNetwork import QLocalServer, QLocalSocket
from PySide6.QtCore import QThread, Signal

# ====== Slack SDK（任意） ======
try:
    from slack_sdk import WebClient
    from slack_sdk.errors import SlackApiError
except Exception:
    WebClient = None
    SlackApiError = Exception

# ====== Keyring（任意） ======
try:
    import keyring
except Exception:
    keyring = None

# ====== Matplotlib（任意：グラフPNG出力 & 日本語フォント対策） ======
try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager as _fm
    # 日本語フォント自動検出（Noto/IPA/Yu/Hiragino 等があれば採用）
    try:
        _candidates = [
            "Noto Sans CJK JP", "Noto Sans JP", "IPAexGothic", "IPAPGothic",
            "Yu Gothic", "Hiragino Sans", "Noto Serif CJK JP", "Source Han Sans JP"
        ]
        _found = None
        for f in _fm.findSystemFonts(fontpaths=None, fontext='ttf'):
            name = _fm.FontProperties(fname=f).get_name()
            if name in _candidates:
                _found = name
                break
        if _found:
            matplotlib.rcParams['font.sans-serif'] = [_found]
            matplotlib.rcParams['font.family'] = 'sans-serif'
        matplotlib.rcParams['axes.unicode_minus'] = False
    except Exception:
        pass
except Exception:
    plt = None

# ====== portalocker（任意：JSON保存の排他ロック強化） ======
try:
    import portalocker
except Exception:
    portalocker = None

APP_NAME = "Worktime"
CONFIG_DIR = Path.home() / ".worktime"
CONFIG_PATH = CONFIG_DIR / "config.json"
DATA_FILE = "work_data.json"  # 既存互換
KEYRING_SERVICE = "worktime"
KEYRING_SLACK_TOKEN_KEY = "slack_bot_token"
SINGLE_INSTANCE_KEY = "worktime.single.instance.guard"

# ====== モデル ======
@dataclass
class DayRecord:
    """1日分のレコード。保存は分単位（表示はHH:MM）。"""
    start: str
    break_start: str
    break_end: str
    end: str
    location: str
    worked_minutes: int
    project: str = ""
    memo: str = ""

    @property
    def worked_time(self) -> timedelta:
        return timedelta(minutes=self.worked_minutes)


# ====== ユーティリティ ======
def calc_duration(start: datetime, bstart: datetime, bend: datetime, end: datetime) -> timedelta:
    """既存ロジック：(休憩開始-開始)+(終了-休憩終了)。"""
    return (bstart - start) + (end - bend)


def dt_on(date_str: str, hhmm: str) -> datetime:
    """'YYYY-MM-DD' × 'HH:MM' を実際のdatetimeへ。"""
    y, m, d = map(int, date_str.split("-"))
    h, mi = map(int, hhmm.split(":"))
    return datetime(y, m, d, h, mi)


def normalize_monotonic(start: datetime, bstart: datetime, bend: datetime, end: datetime):
    """深夜跨ぎなどで時系列が逆転しても、翌日繰り上げで単調増加に補正。"""
    seq = [start, bstart, bend, end]
    out = [seq[0]]
    for t in seq[1:]:
        if t < out[-1]:
            t = t + timedelta(days=1)
        out.append(t)
    return tuple(out)

def resource_path(rel: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / rel

ICON_PATH = resource_path("work_time_icon.ico")

# ====== 設定の保存/読み込み ======
def load_config() -> dict:
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {
        "slack_channel_id": "",         # C… または G…
        "rounding_minutes": 1,
        "overtime_threshold_hours": 8.0,
        "send_with_preview_header": True,
        "fixed_break_minutes": 60,       # 打刻モード用の固定休憩
        "use_keyring": True if keyring else False,
        # keyring不使用時の平文保存（推奨はkeyring）
        "slack_bot_token": "",
        # 候補と自動選択
        "locations": [],
        "projects": [],
        "default_location": "",
        "last_location": "",
    }


def save_config(cfg: dict) -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


def load_token_from_store(cfg: dict) -> str:
    if cfg.get("use_keyring") and keyring is not None:
        try:
            token = keyring.get_password(KEYRING_SERVICE, KEYRING_SLACK_TOKEN_KEY)
            return token or ""
        except Exception:
            return ""
    return cfg.get("slack_bot_token", "")


def save_token_to_store(cfg: dict, token: str) -> None:
    if cfg.get("use_keyring") and keyring is not None:
        try:
            if token:
                keyring.set_password(KEYRING_SERVICE, KEYRING_SLACK_TOKEN_KEY, token)
            else:
                try:
                    keyring.delete_password(KEYRING_SERVICE, KEYRING_SLACK_TOKEN_KEY)
                except Exception:
                    pass
            return
        except Exception:
            pass
    # keyring使わない場合はconfigに保存
    cfg["slack_bot_token"] = token
    save_config(cfg)


# ====== データストア（JSON互換 + 原子書き込み/バックアップ/ロック） ======
class WorkStore:
    def __init__(self, path: str = DATA_FILE):
        self.path = Path(path)
        self.data: Dict[str, DayRecord] = self.load()

    def load(self) -> Dict[str, DayRecord]:
        def _parse_text(txt: str) -> Dict[str, DayRecord]:
            raw = json.loads(txt)
            parsed: Dict[str, DayRecord] = {}
            for date_str, record in raw.items():
                parsed[date_str] = DayRecord(
                    start=record.get("start", ""),
                    break_start=record.get("break_start", ""),
                    break_end=record.get("break_end", ""),
                    end=record.get("end", ""),
                    location=record.get("location", ""),
                    worked_minutes=int(record.get("worked_minutes", 0)),
                    project=record.get("project", ""),
                    memo=record.get("memo", ""),
                )
            return parsed

        # まず本体を試す
        if self.path.exists():
            try:
                return _parse_text(self.path.read_text(encoding="utf-8"))
            except Exception:
                pass

        # ダメなら .bak を試す
        bak = self.path.with_suffix(".bak")
        if bak.exists():
            try:
                return _parse_text(bak.read_text(encoding="utf-8"))
            except Exception:
                pass

        # どうしてもダメなら空
        return {}

    def save(self) -> None:
        """JSONを **一時ファイル→fsync→原子置換**。旧版は .bak に退避。portalockerがあれば排他ロック。"""
        serializable = {}
        for date_str, rec in self.data.items():
            serializable[date_str] = {
                "start": rec.start,
                "break_start": rec.break_start,
                "break_end": rec.break_end,
                "end": rec.end,
                "location": rec.location,
                "worked_minutes": rec.worked_minutes,
                "project": rec.project,
                "memo": rec.memo,
            }
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        bak = self.path.with_suffix(".bak")
        # 既存を退避
        if self.path.exists():
            try:
                if bak.exists():
                    bak.unlink()
                self.path.replace(bak)
            except Exception:
                pass
        # 原子書き込み
        with open(tmp, "w", encoding="utf-8") as f:
            try:
                if portalocker:
                    portalocker.lock(f, portalocker.LOCK_EX)
            except Exception:
                pass
            json.dump(serializable, f, ensure_ascii=False, indent=2)
            f.flush(); os.fsync(f.fileno())
            try:
                if portalocker:
                    portalocker.unlock(f)
            except Exception:
                pass
        os.replace(tmp, self.path)

    # 集計（全期間: 勤務先別 / 案件別）
    def totals_by_location(self) -> Dict[str, timedelta]:
        totals: Dict[str, timedelta] = {}
        for rec in self.data.values():
            if not rec.location:
                continue
            totals[rec.location] = totals.get(rec.location, timedelta()) + rec.worked_time
        return totals

    def totals_by_project(self) -> Dict[str, timedelta]:
        totals: Dict[str, timedelta] = {}
        for rec in self.data.values():
            key = rec.project or "(未設定)"
            totals[key] = totals.get(key, timedelta()) + rec.worked_time
        return totals

    # 月間集計（勤務先別 / 案件別）
    def monthly_totals_by_location(self, year: int, month: int) -> Dict[str, timedelta]:
        totals: Dict[str, timedelta] = {}
        prefix = f"{year}-{month:02d}"
        for date_str, rec in self.data.items():
            if date_str.startswith(prefix):
                if not rec.location:
                    continue
                totals[rec.location] = totals.get(rec.location, timedelta()) + rec.worked_time
        return totals

    def monthly_totals_by_project(self, year: int, month: int) -> Dict[str, timedelta]:
        totals: Dict[str, timedelta] = {}
        prefix = f"{year}-{month:02d}"
        for date_str, rec in self.data.items():
            if date_str.startswith(prefix):
                key = rec.project or "(未設定)"
                totals[key] = totals.get(key, timedelta()) + rec.worked_time
        return totals


# ====== アプリ本体 ======
class MainWindow(QMainWindow):
    """タブUI＋Slack連携のメインウィンドウ。"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — 個人向け勤務時間管理")
        self.setWindowIcon(QIcon(str(ICON_PATH)))
        self.resize(1120, 740)

        # 設定
        self.cfg = load_config()
        self.slack_token = load_token_from_store(self.cfg)
        self.slack_channel_id = self.cfg.get("slack_channel_id", "")

        # ストア
        self.store = WorkStore()

        # 既存データから勤務先/案件の候補を初期投入（初回のみ）
        self.seed_candidates_from_store()

        # タブ
        self.tabs = QTabWidget()
        self.tab_punch = QWidget()      # 0) 打刻
        self.tab_daily = QWidget()      # 1) 日次入力/修正
        self.tab_rules = QWidget()      # 2) ルール
        self.tab_preview = QWidget()    # 3) プレビュー&送信
        self.tab_reports = QWidget()    # 4) 集計
        self.tab_settings = QWidget()   # 5) 設定

        self.tabs.addTab(self.tab_punch,   "0) 打刻")
        self.tabs.addTab(self.tab_daily,   "1) 日次入力")
        self.tabs.addTab(self.tab_rules,   "2) ルール設定")
        self.tabs.addTab(self.tab_preview, "3) プレビュー & 送信")
        self.tabs.addTab(self.tab_reports, "4) 集計")
        self.tabs.addTab(self.tab_settings,"5) 設定")

        self._init_punch()
        self._init_daily()
        self._init_rules()
        self._init_preview()
        self._init_reports()
        self._init_settings()

        root = QWidget(); layout = QVBoxLayout(root)
        layout.addWidget(self.tabs)
        self.setCentralWidget(root)

        # メニュー
        menu = self.menuBar().addMenu("ファイル")
        act_save = QAction("保存", self); act_save.triggered.connect(self.on_save_clicked)
        menu.addAction(act_save)

    # ---------- 0) 打刻（前回勤務先の自動選択＋候補同期） ----------
    def _init_punch(self):
        v = QVBoxLayout(self.tab_punch)

        # 勤務先 / 案件（打刻用コンボ）
        form = QFormLayout()
        self.punch_location = QComboBox(); self.punch_location.setEditable(True)
        for loc in self.cfg.get("locations", []):
            self.punch_location.addItem(loc)
        self.punch_project = QComboBox(); self.punch_project.setEditable(True)
        for pj in self.cfg.get("projects", []):
            self.punch_project.addItem(pj)
        # 前回またはデフォルトの勤務先を自動選択
        if self.cfg.get("last_location"):
            self.punch_location.setEditText(self.cfg.get("last_location"))
        elif self.cfg.get("default_location"):
            self.punch_location.setEditText(self.cfg.get("default_location"))
        form.addRow("勤務先(打刻)", self.punch_location)
        form.addRow("案件(任意)", self.punch_project)

        self.lbl_punch_status = QLabel("状態: 未出勤")
        self.btn_punch_in = QPushButton("出勤（現在時刻）")
        self.btn_punch_out = QPushButton("退勤（現在時刻）")
        self.chk_auto_send = QCheckBox("退勤時に自動でSlack送信"); self.chk_auto_send.setChecked(True)

        hint = QLabel("※ 固定休憩は設定タブで変更可能（既定: 60分）。※ 出退勤は当日レコードに記録されます。")
        hint.setWordWrap(True)

        h = QHBoxLayout(); h.addWidget(self.btn_punch_in); h.addWidget(self.btn_punch_out); h.addStretch(1)
        v.addLayout(form)
        v.addWidget(self.lbl_punch_status)
        v.addLayout(h)
        v.addWidget(self.chk_auto_send)
        v.addWidget(hint)
        v.addStretch(1)

        self.btn_punch_in.clicked.connect(self.on_punch_in)
        self.btn_punch_out.clicked.connect(self.on_punch_out)

    # ---------- 1) 日次入力（深夜跨ぎOK） ----------
    def _init_daily(self):
        main = QHBoxLayout(self.tab_daily)

        # 左: カレンダー
        left = QVBoxLayout()
        self.calendar = QCalendarWidget(); self.calendar.setGridVisible(True)
        self.calendar.setSelectedDate(QDate.currentDate())
        self.calendar.selectionChanged.connect(self.on_calendar_change)
        left.addWidget(self.calendar)
        main.addLayout(left, 1)

        # 右: 入力フォーム
        right = QVBoxLayout(); form = QFormLayout()

        # 勤務先/案件はコンボボックス（履歴候補）
        self.input_location = QComboBox(); self.input_location.setEditable(True)
        for loc in self.cfg.get("locations", []):
            self.input_location.addItem(loc)
        self.input_project = QComboBox(); self.input_project.setEditable(True)
        for pj in self.cfg.get("projects", []):
            self.input_project.addItem(pj)
        # 起動時は前回値があれば自動選択
        if self.cfg.get("last_location"):
            self.input_location.setEditText(self.cfg.get("last_location"))
        elif self.cfg.get("default_location"):
            self.input_location.setEditText(self.cfg.get("default_location"))

        self.input_start = QTimeEdit(); self.input_start.setDisplayFormat("HH:mm"); self.input_start.setTime(QTime(9, 0))
        self.input_bstart = QTimeEdit(); self.input_bstart.setDisplayFormat("HH:mm"); self.input_bstart.setTime(QTime(12, 0))
        self.input_bend = QTimeEdit(); self.input_bend.setDisplayFormat("HH:mm"); self.input_bend.setTime(QTime(13, 0))
        self.input_end = QTimeEdit(); self.input_end.setDisplayFormat("HH:mm"); self.input_end.setTime(QTime(18, 0))
        self.input_memo = QTextEdit(); self.input_memo.setPlaceholderText("任意: タスクや所感など")

        form.addRow("勤務先名", self.input_location)
        form.addRow("案件（任意）", self.input_project)
        form.addRow("開始 (HH:MM)", self.input_start)
        form.addRow("休憩開始 (HH:MM)", self.input_bstart)
        form.addRow("休憩終了 (HH:MM)", self.input_bend)
        form.addRow("終了 (HH:MM)", self.input_end)
        form.addRow("メモ", self.input_memo)

        btns = QHBoxLayout(); self.btn_register = QPushButton("登録/更新"); self.btn_register.clicked.connect(self.on_register)
        self.btn_to_preview = QPushButton("→ プレビューへ"); self.btn_to_preview.clicked.connect(lambda: (self.update_preview(), self.tabs.setCurrentIndex(3)))
        btns.addWidget(self.btn_register); btns.addStretch(1); btns.addWidget(self.btn_to_preview)

        right.addLayout(form); right.addLayout(btns)
        main.addLayout(right, 1)

        self.fill_day_from_store(self.selected_date_str())

    # ---------- 2) ルール（＋指定月の再計算） ----------
    def _init_rules(self):
        lay = QFormLayout(self.tab_rules)
        self.round_min = QSpinBox(); self.round_min.setRange(1, 60); self.round_min.setValue(int(self.cfg.get("rounding_minutes", 1)))
        self.overtime_th = QDoubleSpinBox(); self.overtime_th.setRange(0.0, 24.0); self.overtime_th.setDecimals(2); self.overtime_th.setSingleStep(0.25); self.overtime_th.setValue(float(self.cfg.get("overtime_threshold_hours", 8.0)))
        self.chk_header = QCheckBox("送信メッセージに見出し（日付/区切り）を付ける"); self.chk_header.setChecked(bool(self.cfg.get("send_with_preview_header", True)))

        lay.addRow("端数丸め（分）", self.round_min)
        lay.addRow("残業しきい値（時間）", self.overtime_th)
        lay.addRow("メッセージ装飾", self.chk_header)

        # 指定月の遡及再計算（現在の丸め/固定休憩/残業閾値で再計算）
        self.btn_recalc = QPushButton("指定月を再計算…")
        self.btn_recalc.clicked.connect(lambda: self._ask_recalc_month())
        lay.addRow(self.btn_recalc)

    # ---------- 3) プレビュー & 送信（Block Kit日次） ----------
    def _init_preview(self):
        v = QVBoxLayout(self.tab_preview)
        self.lbl_preview = QLabel("ここに送信プレビューが表示されます。")
        self.lbl_preview.setTextInteractionFlags(Qt.TextSelectableByMouse)
        btns = QHBoxLayout(); self.btn_send = QPushButton("Slackに送信（専用チャンネル）"); self.btn_send.clicked.connect(self.on_send_slack)
        btns.addStretch(1); btns.addWidget(self.btn_send)
        v.addWidget(self.lbl_preview); v.addStretch(1); v.addLayout(btns)

    # ---------- 4) 集計（全期間・月次／CSV／グラフPNG） ----------
    def _init_reports(self):
        v = QVBoxLayout(self.tab_reports)

        v.addWidget(QLabel("勤務先別 合計（全期間）"))
        self.table_total = QTableWidget(0, 2); self.table_total.setHorizontalHeaderLabels(["勤務先", "実働(時:分)"])
        self.table_total.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        v.addWidget(self.table_total)

        v.addWidget(QLabel("案件別 合計（全期間）"))
        self.table_proj_total = QTableWidget(0, 2); self.table_proj_total.setHorizontalHeaderLabels(["案件", "実働(時:分)"])
        self.table_proj_total.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        v.addWidget(self.table_proj_total)

        controls = QHBoxLayout()
        v.addWidget(QLabel("指定月の集計"))
        controls.addWidget(QLabel("年")); self.cmb_year = QComboBox(); cur_year = datetime.now().year
        for y in range(cur_year - 2, cur_year + 3): self.cmb_year.addItem(str(y))
        self.cmb_year.setCurrentText(str(cur_year)); controls.addWidget(self.cmb_year)

        controls.addWidget(QLabel("月")); self.cmb_month = QComboBox();
        for m in range(1, 13): self.cmb_month.addItem(f"{m:02d}")
        self.cmb_month.setCurrentText(f"{datetime.now().month:02d}"); controls.addWidget(self.cmb_month)

        self.btn_show_month = QPushButton("月間集計を表示"); self.btn_show_month.clicked.connect(self.refresh_monthly); controls.addWidget(self.btn_show_month)

        controls.addWidget(QLabel("案件")); self.cmb_project_month = QComboBox(); self.cmb_project_month.setEditable(True)
        for pj in self.cfg.get("projects", []): self.cmb_project_month.addItem(pj)
        self.cmb_project_month.addItem("(未設定)"); controls.addWidget(self.cmb_project_month)

        self.btn_export_month_csv = QPushButton("CSV出力（案件別・指定月）"); self.btn_export_month_csv.clicked.connect(self.export_month_project_csv); controls.addWidget(self.btn_export_month_csv)
        self.btn_share_month_project = QPushButton("Slack共有（この案件・指定月）"); self.btn_share_month_project.clicked.connect(self.share_month_project_to_slack); controls.addWidget(self.btn_share_month_project)
        self.btn_graph = QPushButton("棒グラフPNG生成（案件別・指定月）"); self.btn_graph.clicked.connect(self.generate_month_project_bar_png); controls.addWidget(self.btn_graph)
        self.btn_graph_daily = QPushButton("日別棒グラフPNG（この案件・指定月）"); self.btn_graph_daily.clicked.connect(self.generate_month_project_daily_bar_png); controls.addWidget(self.btn_graph_daily)
        controls.addStretch(1); v.addLayout(controls)

        self.table_month_loc = QTableWidget(0, 2); self.table_month_loc.setHorizontalHeaderLabels(["勤務先", "実働(時:分)"])
        self.table_month_loc.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); v.addWidget(self.table_month_loc)
        self.table_month_proj = QTableWidget(0, 2); self.table_month_proj.setHorizontalHeaderLabels(["案件", "実働(時:分)"])
        self.table_month_proj.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch); v.addWidget(self.table_month_proj)

        self.lbl_chart = QLabel("（グラフPNGをここに表示します）"); self.lbl_chart.setAlignment(Qt.AlignCenter); v.addWidget(self.lbl_chart)

        self.refresh_totals(); self.refresh_monthly()

    # ---------- 5) 設定（Slack/keyring/固定休憩 等） ----------
    def _init_settings(self):
        lay = QFormLayout(self.tab_settings)
        self.edit_token = QLineEdit(load_token_from_store(self.cfg)); self.edit_token.setEchoMode(QLineEdit.Password)
        self.edit_channel = QLineEdit(self.slack_channel_id)
        self.default_break_min = QSpinBox(); self.default_break_min.setRange(0, 180); self.default_break_min.setValue(int(self.cfg.get("fixed_break_minutes", 60)))
        self.chk_use_keyring = QCheckBox("SlackトークンをOSのキーチェーン(keyring)に保存（推奨）"); self.chk_use_keyring.setChecked(bool(self.cfg.get("use_keyring", keyring is not None)))

        self.btn_test = QPushButton("接続テスト（専用チャンネルへ）"); self.btn_test.clicked.connect(self.on_test_slack)
        self.btn_save = QPushButton("保存（設定 & データ & 候補）"); self.btn_save.clicked.connect(self.on_save_clicked)

        lay.addRow("Slack Bot Token", self.edit_token)
        lay.addRow("送信先チャンネルID", self.edit_channel)
        lay.addRow("固定休憩（分）", self.default_break_min)
        lay.addRow(self.chk_use_keyring)
        lay.addRow(self.btn_test)
        lay.addRow(self.btn_save)

    # ====== 日付・入力ハンドリング ======
    def selected_date_str(self) -> str:
        d = self.calendar.selectedDate().toPython(); return d.strftime("%Y-%m-%d")

    def fill_day_from_store(self, date_key: str):
        rec = self.store.data.get(date_key)
        if not rec:
            self.input_location.setEditText(""); self.input_project.setEditText(""); self.input_memo.clear(); return
        self.input_location.setEditText(rec.location); self.input_project.setEditText(rec.project); self.input_memo.setText(rec.memo)
        def set_time(widget: QTimeEdit, hhmm: str):
            try:
                t = datetime.strptime(hhmm, "%H:%M").time(); widget.setTime(QTime(t.hour, t.minute))
            except Exception:
                pass
        set_time(self.input_start, rec.start); set_time(self.input_bstart, rec.break_start); set_time(self.input_bend, rec.break_end); set_time(self.input_end, rec.end)

    def on_calendar_change(self):
        self.fill_day_from_store(self.selected_date_str())

    def _update_candidates(self, value: str, key: str):
        value = (value or "").strip(); 
        if not value: 
            return
        lst: List[str] = list(self.cfg.get(key, []))
        if value not in lst:
            lst.append(value); self.cfg[key] = lst; save_config(self.cfg)
            # UIにも反映（打刻/日次/月次）
            if key == "locations":
                self.input_location.addItem(value); self.punch_location.addItem(value)
            elif key == "projects":
                self.input_project.addItem(value); self.punch_project.addItem(value)
                self.cmb_project_month.addItem(value)

    def on_register(self):
        """日次フォームの登録。深夜跨ぎも単調増加補正して計算。"""
        date_key = self.selected_date_str()
        location = self.input_location.currentText().strip()
        project = self.input_project.currentText().strip()
        memo = self.input_memo.toPlainText().strip()
        if not location:
            QMessageBox.warning(self, "入力エラー", "勤務先名を入力してください。"); return
        def to_dt(qt: QTime) -> datetime:
            return datetime(2000,1,1, qt.hour(), qt.minute())
        t_start  = to_dt(self.input_start.time())
        t_bstart = to_dt(self.input_bstart.time())
        t_bend   = to_dt(self.input_bend.time())
        t_end    = to_dt(self.input_end.time())
        # 深夜跨ぎを許容
        t_start, t_bstart, t_bend, t_end = normalize_monotonic(t_start, t_bstart, t_bend, t_end)
        duration = calc_duration(t_start, t_bstart, t_bend, t_end)
        duration = self.round_duration(duration, self.round_min.value())
        rec = DayRecord(
            start=t_start.strftime("%H:%M"), break_start=t_bstart.strftime("%H:%M"),
            break_end=t_bend.strftime("%H:%M"), end=t_end.strftime("%H:%M"),
            location=location, worked_minutes=int(duration.total_seconds() // 60),
            project=project, memo=memo,
        )
        self.store.data[date_key] = rec; self.store.save()
        # 候補更新 & last_location
        self._update_candidates(location, "locations");
        if project: self._update_candidates(project, "projects")
        self.cfg["last_location"] = location; save_config(self.cfg)
        self.punch_location.setEditText(location); self.punch_project.setEditText(project)
        h, m = self.hhmm_from_td(duration)
        QMessageBox.information(self, "登録完了", f"{date_key} を登録/更新しました。実働: {h}時間 {m}分")
        self.refresh_totals(); self.update_preview()

    # ====== 打刻処理（深夜跨ぎOK / 前回勤務先を保持） ======
    def _find_open_record_date(self) -> str | None:
        """endが未入力の最新日付を返す（前日の出勤→翌日退勤などに対応）。"""
        for dk in sorted(self.store.data.keys(), reverse=True):
            rec = self.store.data[dk]
            if rec.start and not rec.end:
                return dk
        return None

    def on_punch_in(self):
        now = datetime.now(); date_key = now.strftime("%Y-%m-%d")
        location = self.punch_location.currentText().strip() or self.cfg.get("last_location") or self.cfg.get("default_location", "")
        if not location and self.cfg.get("locations"):
            location = self.cfg["locations"][0]; self.punch_location.setEditText(location)
        project = self.punch_project.currentText().strip(); memo = self.input_memo.toPlainText().strip()
        rec = self.store.data.get(date_key)
        if rec is None:
            rec = DayRecord(start=now.strftime('%H:%M'), break_start="", break_end="", end="", location=location, worked_minutes=0, project=project, memo=memo)
            self.store.data[date_key] = rec
        else:
            rec.start = now.strftime('%H:%M'); rec.end = ""; rec.break_start = ""; rec.break_end = ""; rec.worked_minutes = 0
            rec.location = location; rec.project = project; rec.memo = memo
        self.store.save()
        # 候補 & last_location
        self._update_candidates(location, "locations");
        if project: self._update_candidates(project, "projects")
        self.cfg["last_location"] = location; save_config(self.cfg)
        self.lbl_punch_status.setText(f"状態: 出勤中（{rec.start} 開始）")
        self.calendar.setSelectedDate(QDate.currentDate()); self.fill_day_from_store(date_key)

    def on_punch_out(self):
        now = datetime.now()
        open_date = self._find_open_record_date()
        if not open_date:
            QMessageBox.warning(self, "未出勤", "先に『出勤』を押してください。"); return
        rec = self.store.data.get(open_date)
        if rec is None or not rec.start:
            QMessageBox.warning(self, "未出勤", "先に『出勤』を押してください。"); return
        t_start = dt_on(open_date, rec.start)
        t_end   = now
        if t_end < t_start:  # システム時計変化などの保険
            t_end = t_start
        duration = t_end - t_start
        duration -= timedelta(minutes=int(self.cfg.get("fixed_break_minutes", 60)))  # 固定休憩控除
        if duration.total_seconds() < 0:
            duration = timedelta(0)
        duration = self.round_duration(duration, int(self.cfg.get("rounding_minutes", 1)))
        rec.end = now.strftime('%H:%M'); rec.worked_minutes = int(duration.total_seconds() // 60)
        self.store.save(); self.lbl_punch_status.setText("状態: 退勤済み")
        self.refresh_totals(); self.update_preview()
        if self.chk_auto_send.isChecked():
            self.on_send_slack()

    # ====== 計算/整形 ======
    @staticmethod
    def round_duration(td: timedelta, minutes: int) -> timedelta:
        if minutes <= 1: return td
        total = int(td.total_seconds() // 60); r = minutes
        rounded = (total + r // 2) // r * r  # 四捨五入
        return timedelta(minutes=rounded)

    @staticmethod
    def hhmm_from_td(td: timedelta) -> Tuple[int, int]:
        mins = int(td.total_seconds() // 60); h, m = divmod(mins, 60); return h, m

    def build_preview(self) -> Tuple[str, dict]:
        """テキスト版プレビュー（Block Kit送信時のフォールバック用）。"""
        date_key = self.selected_date_str(); rec = self.store.data.get(date_key)
        if not rec: return ("本日（または選択日）のデータが未登録です。", {})
        rounded = timedelta(minutes=rec.worked_minutes)
        ot_th = float(self.overtime_th.value()); ot_td = max(timedelta(0), rounded - timedelta(hours=ot_th)); nor_td = rounded - ot_td
        def fmt(td: timedelta) -> str:
            h, m = self.hhmm_from_td(td); return f"{h:02d}:{m:02d}"
        parts = []
        if self.chk_header.isChecked(): parts.append(f"――――――――――――――――――{date_key} 勤務記録")
        parts.append(f"勤務先: {rec.location}")
        if rec.project: parts.append(f"案件: {rec.project}")
        bstart = ("(固定)" if (rec.break_start == "" and rec.break_end == "") else rec.break_start)
        parts.append(f"開始: {rec.start} / 休憩: {bstart}-{rec.break_end} / 終了: {rec.end}")
        parts.append(f"所定: {fmt(nor_td)} / 残業: {fmt(ot_td)} / 合計: {fmt(rounded)}")
        if rec.memo: parts.append(f"メモ: {rec.memo}")
        text = "\n".join(parts)
        meta = {
            "rounded_minutes": int(rounded.total_seconds() // 60),
            "normal_minutes": int(nor_td.total_seconds() // 60),
            "overtime_minutes": int(ot_td.total_seconds() // 60),
        }
        return text, meta

    # ====== Slackメッセージ（Block Kit） ======
    def build_blocks_daily(self, date_key: str):
        rec = self.store.data.get(date_key)
        if not rec: return None
        rounded = timedelta(minutes=rec.worked_minutes)
        ot_th = float(self.overtime_th.value()); ot_td = max(timedelta(0), rounded - timedelta(hours=ot_th)); nor_td = rounded - ot_td
        def fmt(td: timedelta) -> str:
            h, m = self.hhmm_from_td(td); return f"{h:02d}:{m:02d}"
        lines = [
            f"開始: {rec.start} / 休憩: {'(固定)' if (rec.break_start=='' and rec.break_end=='') else rec.break_start}-{rec.break_end} / 終了: {rec.end}",
            f"所定: {fmt(nor_td)} / 残業: {fmt(ot_td)} / 合計: {fmt(rounded)}",
        ]
        if rec.memo: lines.append(f"メモ: {rec.memo}")
        blocks = [
            {"type": "header", "text": {"type": "plain_text", "text": f"{date_key} 勤務記録", "emoji": True}},
            {"type": "section", "fields": [
                {"type": "mrkdwn", "text": f"*勤務先*: {rec.location}"},
                {"type": "mrkdwn", "text": f"*案件*: {rec.project or '(未設定)'}"},
            ]},
            {"type": "section", "text": {"type": "mrkdwn", "text": "\n".join(lines)}},
        ]
        return blocks

    def build_blocks_month_project(self, year: int, month: int, project: str):
        rows = self._month_project_rows(year, month, project)
        total_mins = sum(m for _, m in rows); days = len(rows); avg = (total_mins / days) if days else 0
        h_total, m_total = divmod(total_mins, 60)
        table_lines = ["日付   実働(時間)  分"]
        for d, mins in rows:
            hh, mm = divmod(mins, 60); table_lines.append(f"{d}  {hh:02d}:{mm:02d}     {mins:>4}")
        blocks = [
            {"type": "header", "text": {"type": "plain_text", "text": f"{year}-{month:02d} 案件レポート", "emoji": True}},
            {"type": "section", "fields": [
                {"type": "mrkdwn", "text": f"*案件*: {project or '(未設定)'}"},
                {"type": "mrkdwn", "text": f"*合計*: {h_total:02d}:{m_total:02d}（{total_mins}分）"},
                {"type": "mrkdwn", "text": f"*日数*: {days} 日"},
                {"type": "mrkdwn", "text": f"*平均*: {avg/60:.1f} h/日"},
            ]},
            {"type": "divider"},
            {"type": "section", "text": {"type": "mrkdwn", "text": "```" + "\n".join(table_lines) + "```"}},
        ]
        return blocks

    # ====== 集計表示 ======
    @staticmethod
    def _set_table_from_totals(table: QTableWidget, totals: Dict[str, timedelta]):
        table.setRowCount(0)
        if not totals:
            table.setRowCount(1); table.setItem(0, 0, QTableWidgetItem("(データなし)")); table.setItem(0, 1, QTableWidgetItem("00:00")); return
        for i, (key, td) in enumerate(totals.items()):
            table.insertRow(i)
            h, m = divmod(int(td.total_seconds() // 60), 60)
            table.setItem(i, 0, QTableWidgetItem(key)); table.setItem(i, 1, QTableWidgetItem(f"{h:02d}:{m:02d}"))

    def refresh_totals(self):
        self._set_table_from_totals(self.table_total, self.store.totals_by_location())
        self._set_table_from_totals(self.table_proj_total, self.store.totals_by_project())

    def refresh_monthly(self):
        y = int(self.cmb_year.currentText()); m = int(self.cmb_month.currentText())
        self._set_table_from_totals(self.table_month_loc, self.store.monthly_totals_by_location(y, m))
        self._set_table_from_totals(self.table_month_proj, self.store.monthly_totals_by_project(y, m))

    # ====== CSV/グラフ/Slack 共有（案件×指定月） ======
    def _month_project_rows(self, year: int, month: int, project: str) -> List[Tuple[str, int]]:
        prefix = f"{year}-{month:02d}"; rows: List[Tuple[str, int]] = []
        for date_str, rec in sorted(self.store.data.items()):
            if date_str.startswith(prefix) and (rec.project or "(未設定)") == (project or "(未設定)"):
                rows.append((date_str, rec.worked_minutes))
        return rows

    def export_month_project_csv(self):
        y = int(self.cmb_year.currentText()); m = int(self.cmb_month.currentText()); project = self.cmb_project_month.currentText().strip(); rows = self._month_project_rows(y, m, project)
        if not rows: QMessageBox.information(self, "CSV出力", "該当データがありません。"); return
        fn, _ = QFileDialog.getSaveFileName(self, "CSVを保存", f"worktime_{y}{m:02d}_{project or '未設定'}.csv", "CSV Files (*.csv)")
        if not fn: return
        with open(fn, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f); w.writerow(["日付", "実働(分)"])
            for d, mins in rows: w.writerow([d, mins])
        QMessageBox.information(self, "CSV出力", f"保存しました:{fn}")

    def share_month_project_to_slack(self):
        self.on_save_clicked()
        if WebClient is None:
            QMessageBox.critical(self, "依存関係", "slack_sdk が見つかりません。 pip install slack_sdk を実行してください。"); return
        if not self.slack_token or not self.slack_channel_id:
            QMessageBox.warning(self, "未入力", "設定タブで Slack トークンとチャンネルIDを入力・保存してください。"); return

        y = int(self.cmb_year.currentText()); m = int(self.cmb_month.currentText())
        project = self.cmb_project_month.currentText().strip()
        blocks = self.build_blocks_month_project(y, m, project)

        # 任意の短いテキストを同送
        text = "案件レポート"
        self.btn_share_month_project.setEnabled(False)
        self._worker = SlackWorker(self.slack_token, self.slack_channel_id, text=text, blocks=blocks)
        self._worker.success.connect(lambda m: (QMessageBox.information(self, "送信", m), self.btn_share_month_project.setEnabled(True)))
        self._worker.failure.connect(lambda m: (QMessageBox.critical(self, "エラー", m), self.btn_share_month_project.setEnabled(True)))
        self._worker.start()

    def generate_month_project_bar_png(self):
        if plt is None: QMessageBox.information(self, "グラフ", "matplotlib が見つかりません。任意で 'pip install matplotlib' を実行してください。"); return
        y = int(self.cmb_year.currentText()); m = int(self.cmb_month.currentText()); totals = self.store.monthly_totals_by_project(y, m)
        if not totals: QMessageBox.information(self, "グラフ", "該当データがありません。"); return
        labels = list(totals.keys()); values_hours = [int(td.total_seconds() // 60) / 60.0 for td in totals.values()]
        plt.figure(); bars = plt.bar(labels, values_hours); plt.ylabel("実働(時間)"); plt.title(f"案件別 実働 {y}-{m:02d}"); plt.xticks(rotation=20, ha='right')
        for rect, v in zip(bars, values_hours): plt.text(rect.get_x() + rect.get_width()/2, rect.get_height(), f"{v:.1f}h", ha="center", va="bottom")
        out = CONFIG_DIR / f"chart_{y}{m:02d}_project.png"; CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        plt.tight_layout(); plt.savefig(out); plt.close()
        pix = QPixmap(str(out));
        if not pix.isNull(): self.lbl_chart.setPixmap(pix.scaledToWidth(720, Qt.SmoothTransformation))
        QMessageBox.information(self, "グラフ", f"保存しました:{out}")

    def generate_month_project_daily_bar_png(self):
        if plt is None: QMessageBox.information(self, "グラフ", "matplotlib が見つかりません。任意で 'pip install matplotlib' を実行してください。"); return
        y = int(self.cmb_year.currentText()); m = int(self.cmb_month.currentText()); project = self.cmb_project_month.currentText().strip(); rows = self._month_project_rows(y, m, project)
        if not rows: QMessageBox.information(self, "グラフ", "該当データがありません。"); return
        dates = [d[8:] for d, _ in rows]; values_hours = [mins/60.0 for _, mins in rows]
        plt.figure(); bars = plt.bar(dates, values_hours); plt.ylabel("実働(時間)"); plt.title(f"{project or '(未設定)'} 日別 実働 {y}-{m:02d}")
        for rect, v in zip(bars, values_hours): plt.text(rect.get_x() + rect.get_width()/2, rect.get_height(), f"{v:.1f}h", ha="center", va="bottom")
        out = CONFIG_DIR / f"chart_{y}{m:02d}_{(project or '未設定')}_daily.png"; CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        plt.tight_layout(); plt.savefig(out); plt.close()
        pix = QPixmap(str(out));
        if not pix.isNull(): self.lbl_chart.setPixmap(pix.scaledToWidth(720, Qt.SmoothTransformation))
        QMessageBox.information(self, "グラフ", f"保存しました:{out}")

    # ====== 保存/Slack ======
    def on_save_clicked(self):
        # UI → 設定へ反映
        self.cfg["slack_channel_id"] = self.edit_channel.text().strip()
        self.cfg["rounding_minutes"] = int(self.round_min.value())
        self.cfg["overtime_threshold_hours"] = float(self.overtime_th.value())
        self.cfg["send_with_preview_header"] = bool(self.chk_header.isChecked())
        self.cfg["fixed_break_minutes"] = int(self.default_break_min.value())
        self.cfg["use_keyring"] = bool(self.chk_use_keyring.isChecked() and keyring is not None)
        # Token 保存
        token = self.edit_token.text().strip(); save_token_to_store(self.cfg, token)
        self.slack_token = load_token_from_store(self.cfg); self.slack_channel_id = self.cfg.get("slack_channel_id", "")
        save_config(self.cfg); self.store.save()
        QMessageBox.information(self, "保存", f"設定とデータを保存しました。{os.path.abspath(DATA_FILE)}設定ファイル: {CONFIG_PATH}")

    def on_test_slack(self):
        self.on_save_clicked()
        if not self.slack_token or not self.slack_channel_id:
            QMessageBox.warning(self, "未入力", "SlackトークンとチャンネルIDを入力・保存してください。"); return
        if WebClient is None:
            QMessageBox.critical(self, "依存関係", "slack_sdk が見つかりません。 pip install slack_sdk を実行してください。"); return

        text = ":white_check_mark: 接続テスト（このチャンネルに打刻/レポートを送ります）"
        self.btn_test.setEnabled(False)
        self._worker = SlackWorker(self.slack_token, self.slack_channel_id, text=text, blocks=None)
        self._worker.success.connect(lambda m: (QMessageBox.information(self, "成功", m), self.btn_test.setEnabled(True)))
        self._worker.failure.connect(lambda m: (QMessageBox.critical(self, "Slackエラー", m), self.btn_test.setEnabled(True)))
        self._worker.start()


    def on_send_slack(self):
        self.on_save_clicked()
        if WebClient is None:
            QMessageBox.critical(self, "依存関係", "slack_sdk が見つかりません。 pip install slack_sdk を実行してください。"); return
        if not self.slack_token or not self.slack_channel_id:
            QMessageBox.warning(self, "未入力", "設定タブで Slack トークンとチャンネルIDを入力・保存してください。"); return

        date_key = self.selected_date_str()
        blocks = self.build_blocks_daily(date_key)
        text, _ = self.build_preview()
        if not blocks:
            QMessageBox.warning(self, "未登録", "選択日のデータを登録してください。"); return

        self.btn_send.setEnabled(False)
        self._worker = SlackWorker(self.slack_token, self.slack_channel_id, text=text, blocks=blocks)
        self._worker.success.connect(lambda m: (QMessageBox.information(self, "送信", m), self.btn_send.setEnabled(True)))
        self._worker.failure.connect(lambda m: (QMessageBox.critical(self, "エラー", m), self.btn_send.setEnabled(True)))
        self._worker.start()

    class SlackWorker(QThread):
        success = Signal(str)   # メッセージ
        failure = Signal(str)   # エラーメッセージ

        def __init__(self, token: str, channel: str, text: str = "", blocks: dict | list | None = None):
            super().__init__()
            self.token = token
            self.channel = channel
            self.text = text
            self.blocks = blocks

        def run(self):
            try:
                client = WebClient(token=self.token)
                if self.blocks:
                    client.chat_postMessage(channel=self.channel, text=self.text, blocks=self.blocks)
                else:
                    client.chat_postMessage(channel=self.channel, text=self.text)
                self.success.emit("Slack送信が完了しました。")
            except SlackApiError as e:
                msg = getattr(e, "response", {}).get("error", str(e))
                self.failure.emit(f"Slackエラー: {msg}")
            except Exception as e:
                self.failure.emit(f"送信に失敗しました: {e}")

    # ====== 遡及再計算（指定月） ======
    def _ask_recalc_month(self):
        y = int(self.cmb_year.currentText()) if hasattr(self, "cmb_year") else int(QDate.currentDate().toString("yyyy"))
        m = int(self.cmb_month.currentText()) if hasattr(self, "cmb_month") else int(QDate.currentDate().toString("MM"))
        ret = QMessageBox.question(self, "再計算", f"{y}-{m:02d} を現在のルールで再計算します。よろしいですか？")
        if ret == QMessageBox.Yes:
            self.recalc_month(y, m)

    def recalc_month(self, year: int, month: int):
        prefix = f"{year}-{month:02d}"; changed = 0; skipped = 0
        for date_key, rec in sorted(self.store.data.items()):
            if not date_key.startswith(prefix):
                continue
            try:
                if rec.break_start and rec.break_end:
                    # 手入力レコード
                    def to_dt(hhmm: str) -> datetime:
                        h, mi = map(int, hhmm.split(":"))
                        return datetime(2000, 1, 1, h, mi)
                    t_start = to_dt(rec.start); t_bstart = to_dt(rec.break_start)
                    t_bend  = to_dt(rec.break_end); t_end = to_dt(rec.end)
                    t_start, t_bstart, t_bend, t_end = normalize_monotonic(t_start, t_bstart, t_bend, t_end)
                    dur = calc_duration(t_start, t_bstart, t_bend, t_end)
                else:
                    # 打刻レコード（固定休憩）
                    if not rec.start or not rec.end:
                        skipped += 1
                        continue
                    t_start_real = dt_on(date_key, rec.start)
                    t_end_real   = dt_on(date_key, rec.end)
                    if t_end_real < t_start_real:
                        t_end_real += timedelta(days=1)
                    dur = t_end_real - t_start_real
                    dur -= timedelta(minutes=int(self.cfg.get("fixed_break_minutes", 60)))
                    if dur.total_seconds() < 0:
                        dur = timedelta(0)
                dur = self.round_duration(dur, int(self.cfg.get("rounding_minutes", 1)))
                new_minutes = int(dur.total_seconds() // 60)
                if new_minutes != rec.worked_minutes:
                    rec.worked_minutes = new_minutes; changed += 1
            except Exception:
                skipped += 1
                continue

        if changed:
            self.store.save()
        QMessageBox.information(self, "再計算",
            f"{year}-{month:02d} の {changed} 件を再計算しました（スキップ: {skipped} 件）。")
        self.refresh_totals(); self.refresh_monthly(); self.update_preview()


    # ====== その他 ======
    def seed_candidates_from_store(self):
        """初回起動時、既存JSONから勤務先/案件の候補を生成。"""
        need_save = False
        if not self.cfg.get("locations"):
            locs = sorted({rec.location for rec in self.store.data.values() if rec.location})
            if locs: self.cfg["locations"] = locs; need_save = True
        if not self.cfg.get("projects"):
            projs = sorted({rec.project for rec in self.store.data.values() if rec.project})
            if projs: self.cfg["projects"] = projs; need_save = True
        if need_save: save_config(self.cfg)

    def closeEvent(self, event):
        """終了時の念のため保存。"""
        try:
            save_config(self.cfg); self.store.save()
        except Exception:
            pass
        super().closeEvent(event)


# ====== メイン（シングルインスタンスガード + 既存起動の前面化） ======
def main():
    import sys
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(str(ICON_PATH)))

    # 既存プロセスに接続できたら、アクティベート要求を送り自身は終了
    probe = QLocalSocket(); probe.connectToServer(SINGLE_INSTANCE_KEY)
    if probe.waitForConnected(150):
        try:
            probe.write(b"ACTIVATE"); probe.flush(); probe.waitForBytesWritten(100)
        except Exception:
            pass
        probe.disconnectFromServer()
        QMessageBox.information(None, APP_NAME, "すでにWorktimeが起動しています。画面を前面に表示します。")
        sys.exit(0)

    # 本インスタンスが受け口（サーバ）になる
    QLocalServer.removeServer(SINGLE_INSTANCE_KEY)  # ステイルなソケットを掃除
    server = QLocalServer(); server.listen(SINGLE_INSTANCE_KEY)

    w = MainWindow(); w.show()

    def on_new_connection():
        # 他プロセスからの接続（=二重起動試行）→ 画面を前面化
        sock = server.nextPendingConnection()
        if sock is not None:
            sock.close()
        w.show(); w.raise_(); w.activateWindow()
    server.newConnection.connect(on_new_connection)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
