# app.py — 單一檔 PyQt6 應用程式
# --------------------------------
# 環境：Windows、Python 3.12.5
# 相依：PyQt6、pandas、openpyxl
# 本程式維持原有 UI/功能，外加兩段式 Download→Upload 與 D:/<Project>/logs 輸出。

from __future__ import annotations
import sys
import os
import logging
from logging.handlers import RotatingFileHandler
import time as time_mod
from datetime import datetime, timedelta
from datetime import datetime as _dt
from typing import Optional, List, Tuple, Any, Dict
import pandas as pd
from PyQt6.QtCore import Qt, QTimer, QAbstractTableModel, QModelIndex, QUrl, QSettings
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QGroupBox, QLabel, QLineEdit, QPushButton, QCheckBox, QMessageBox,
    QTableView, QSplitter, QFileDialog, QSpinBox, QFormLayout, QInputDialog,
 QDialog,
 QTextEdit)
import subprocess
import re
import configparser
import locale
from dataclasses import dataclass
import shlex

import csv  # NEW
import re   # NEW
from PyQt6.QtGui import QDesktopServices  # NEW
from PyQt6.QtGui import QDesktopServices, QIcon  # UPDATED

# 應用版本（發版時手動同步 /VERSION）
__version__ = "2.0.0"  # NEW

# --- 常數與 QSettings 範圍 ---  # NEW
LOG_ROOT = "D:/"
ORG_NAME = "UploadDashboard"
APP_NAME = "UploadDashboard"

# 設定鍵值
SETK_AUTO_ENABLED   = "Cleanup/AutoEnabled"
SETK_RETENTION_DAYS = "Cleanup/RetentionDays"
SETK_SKIP_PREVIEW   = "Cleanup/SkipPreview"
SETK_ROOT_PATH      = "Cleanup/RootPath"  # 預留（目前固定用 LOG_ROOT）
DEFAULT_RETENTION   = 30


DEFAULT_TIMEOUT_SEC: int = 300
SUCCESS_RETURN_CODES = {0}

EXCEL_PATH: str = "upload_stats.xlsx"
SHEET_PROJECTS: str = "Projects"
SHEET_EVENTS: str = "Events"
LOG_DIR: str = "logs"
LOG_FILE: str = os.path.join(LOG_DIR, "app.log")

def _project_logs_dir_on_D(project_name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", str(project_name)).strip() or "unnamed"
    return os.path.normpath(os.path.join("D:/", cleaned, "logs"))

def _setup_logging() -> logging.Logger:
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger("app")
    logger.setLevel(logging.INFO)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch_fmt = logging.Formatter("[%(asctime)s] %(levelname)s - %(message)s")
    ch.setFormatter(ch_fmt)
    if not logger.handlers:
        logger.addHandler(ch)
    return logger

LOGGER = _setup_logging()

def now_iso() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")

def parse_dt(s: str) -> datetime:
    return datetime.fromisoformat(s)

def safe_int(x: Any, default: int = 0) -> int:
    try:
        return int(x)
    except Exception:
        return default

def _decode_best_effort(b: bytes) -> tuple[str, str]:
    if not b:
        return "", "none"
    if b.startswith(b"\xef\xbb\xbf"):
        return b.decode("utf-8-sig", errors="replace"), "utf-8-sig"
    candidates: List[Optional[str]] = []
    pref = locale.getpreferredencoding(False) or ""
    candidates.extend(["utf-8", pref, "cp950", "big5", "cp936", "gbk", "cp932", "cp437"])
    seen = set()
    for enc in candidates:
        enc_lc = (enc or "").lower()
        if not enc_lc or enc_lc in seen:
            continue
        seen.add(enc_lc)
        try:
            return b.decode(enc_lc), enc_lc
        except Exception:
            continue
    return b.decode("utf-8", errors="replace"), "utf-8*"

class FileLock:
    def __init__(self, target_path: str, timeout_sec: float = 5.0, poll_sec: float = 0.1) -> None:
        self.lock_path: str = target_path + ".lock"
        self.timeout_sec: float = timeout_sec
        self.poll_sec: float = poll_sec
        self.fd: Optional[int] = None

    def acquire(self) -> None:
        start = time_mod.time()
        while True:
            try:
                self.fd = os.open(self.lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
                os.write(self.fd, str(os.getpid()).encode("utf-8"))
                return
            except FileExistsError:
                if time_mod.time() - start > self.timeout_sec:
                    raise TimeoutError(f"Could not acquire lock: {self.lock_path}")
                time_mod.sleep(self.poll_sec)

    def release(self) -> None:
        try:
            if self.fd is not None:
                os.close(self.fd)
                self.fd = None
            if os.path.exists(self.lock_path):
                os.remove(self.lock_path)
        except Exception as e:
            LOGGER.warning("FileLock release error: %s", e)

    def __enter__(self) -> "FileLock":
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # type: ignore[override]
        self.release()

class ExcelStore:
    """Excel 資料庫：
       Projects: id, name, interval_minutes, enabled, last_run_at, download_bat_path, upload_bat_path
       Events:   id, project_id, project, status, created_at, message
    """
    def __init__(self, path: str = EXCEL_PATH) -> None:
        self.path: str = path
        self.lock = FileLock(self.path, timeout_sec=8.0)
        self._ensure_file()

    @staticmethod
    def _sanitize_for_path(name: str) -> str:
        cleaned = re.sub(r'[<>:"/\\|?*]+', "_", str(name)).strip()
        return cleaned or "unnamed"

    def _ensure_file(self) -> None:
        with self.lock:
            if not os.path.exists(self.path):
                LOGGER.info("Creating new Excel store at %s", self.path)
                projects = pd.DataFrame(columns=[
                    "id", "name", "interval_minutes", "enabled", "last_run_at",
                    "download_bat_path", "upload_bat_path"
                ])
                events = pd.DataFrame(columns=["id", "project_id", "project", "status", "created_at", "message"])
                self._write_all(projects, events)
                return

            projects, events = self._read_all(allow_missing=True)

            proj_cols = [
                "id", "name", "interval_minutes", "enabled", "last_run_at",
                "download_bat_path", "upload_bat_path"
            ]
            for c in proj_cols:
                if c not in projects.columns:
                    if c == "interval_minutes":
                        projects[c] = 60
                    elif c == "enabled":
                        projects[c] = 0
                    else:
                        projects[c] = ""
            projects = projects.reindex(columns=proj_cols, fill_value="")

            evt_cols = ["id", "project_id", "project", "status", "created_at", "message"]
            for c in evt_cols:
                if c not in events.columns:
                    events[c] = ""
            events = events[evt_cols]

            projects["id"] = pd.to_numeric(projects["id"], errors="coerce").fillna(0).astype(int)
            projects["interval_minutes"] = pd.to_numeric(projects["interval_minutes"], errors="coerce").fillna(60).astype(int)
            projects["enabled"] = pd.to_numeric(projects["enabled"], errors="coerce").fillna(0).astype(int)
            events["id"] = pd.to_numeric(events["id"], errors="coerce").fillna(0).astype(int)
            events["project_id"] = pd.to_numeric(events["project_id"], errors="coerce").fillna(0).astype(int)

            self._write_all(projects, events)

    def _read_all(self, allow_missing: bool = False) -> Tuple[pd.DataFrame, pd.DataFrame]:
        try:
            xls = pd.ExcelFile(self.path, engine="openpyxl")
            sheets = xls.sheet_names
        except Exception:
            if allow_missing:
                empty_p = pd.DataFrame(columns=[
                    "id", "name", "interval_minutes", "enabled", "last_run_at",
                    "download_bat_path", "upload_bat_path"
                ])
                empty_e = pd.DataFrame(columns=["id", "project_id", "project", "status", "created_at", "message"])
                return empty_p, empty_e
            raise

        if SHEET_PROJECTS in sheets:
            projects = pd.read_excel(self.path, sheet_name=SHEET_PROJECTS, engine="openpyxl")
        else:
            projects = pd.DataFrame(columns=[
                "id", "name", "interval_minutes", "enabled", "last_run_at",
                "download_bat_path", "upload_bat_path"
            ])

        if SHEET_EVENTS in sheets:
            events = pd.read_excel(self.path, sheet_name=SHEET_EVENTS, engine="openpyxl")
        else:
            events = pd.DataFrame(columns=["id", "project_id", "project", "status", "created_at", "message"])

        for c in ["download_bat_path", "upload_bat_path"]:
            if c not in projects.columns:
                projects[c] = ""
        projects = projects[["id", "name", "interval_minutes", "enabled", "last_run_at", "download_bat_path", "upload_bat_path"]]

        return projects, events

    def _write_all(self, projects: pd.DataFrame, events: pd.DataFrame) -> None:
        with pd.ExcelWriter(self.path, engine="openpyxl", mode="w") as writer:
            projects.to_excel(writer, sheet_name=SHEET_PROJECTS, index=False)
            events.to_excel(writer, sheet_name=SHEET_EVENTS, index=False)

    def _next_id(self, df: pd.DataFrame) -> int:
        if df.empty or "id" not in df.columns:
            return 1
        mx = pd.to_numeric(df["id"], errors="coerce").fillna(0).max()
        return int(mx) + 1

    def list_projects(self) -> List[Dict[str, Any]]:
        with self.lock:
            projects, _ = self._read_all()
            if projects.empty:
                return []
            projects = projects.sort_values(by="name", kind="stable")
            return projects.to_dict("records")

    def get_project_by_id(self, project_id: int) -> Optional[Dict[str, Any]]:
        with self.lock:
            projects, _ = self._read_all()
            row = projects[projects["id"] == project_id]
            if row.empty:
                return None
            return row.iloc[0].to_dict()

    def add_project(self, name: str, interval_minutes: int = 60, enabled: int = 0) -> None:
        name = name.strip()
        with self.lock:
            projects, events = self._read_all()
            if (projects["name"].astype(str).str.strip() == name).any():
                raise ValueError("DUPLICATE_NAME")
            pid = self._next_id(projects)
            new_row = {
                "id": pid,
                "name": name,
                "interval_minutes": int(interval_minutes),
                "enabled": int(enabled),
                "last_run_at": "",
                "download_bat_path": "",
                "upload_bat_path": "",
            }
            projects = pd.concat([projects, pd.DataFrame([new_row])], ignore_index=True)
            self._write_all(projects, events)
            LOGGER.info("Project added: id=%s name=%s", pid, name)

    def delete_project(self, project_id: int) -> None:
        with self.lock:
            projects, events = self._read_all()
            projects = projects[projects["id"] != project_id].reset_index(drop=True)
            events = events[events["project_id"] != project_id].reset_index(drop=True)
            self._write_all(projects, events)
            LOGGER.info("Project deleted: id=%s", project_id)

    def update_project(
        self,
        project_id: int,
        name: str,
        interval_minutes: int,
        enabled: int,
        download_bat_path: Optional[str] = None,
        upload_bat_path: Optional[str] = None,
    ) -> None:
        name = name.strip()
        with self.lock:
            projects, events = self._read_all()
            dup = projects[(projects["name"].astype(str).str.strip() == name) & (projects["id"] != project_id)]
            if not dup.empty:
                raise ValueError("DUPLICATE_NAME")
            idx = projects.index[projects["id"] == project_id]
            if len(idx) == 0:
                return
            i = idx[0]
            projects.at[i, "name"] = name
            projects.at[i, "interval_minutes"] = int(interval_minutes)
            projects.at[i, "enabled"] = int(enabled)
            if download_bat_path is not None:
                projects.at[i, "download_bat_path"] = download_bat_path
            if upload_bat_path is not None:
                projects.at[i, "upload_bat_path"] = upload_bat_path
            if not events.empty:
                events.loc[events["project_id"] == project_id, "project"] = name
            self._write_all(projects, events)
            LOGGER.info(
                "Project updated: id=%s name=%s interval=%s enabled=%s download_bat=%s upload_bat=%s",
                project_id, name, interval_minutes, enabled, download_bat_path, upload_bat_path
            )

    def update_last_run_at(self, project_id: int, dt_iso: Optional[str]) -> None:
        with self.lock:
            projects, events = self._read_all()
            idx = projects.index[projects["id"] == project_id]
            if len(idx) == 0:
                return
            i = idx[0]
            projects.at[i, "last_run_at"] = dt_iso if dt_iso else ""
            self._write_all(projects, events)

    # ---- 在 ExcelStore 類別內，整段替換 add_event(...) ----  # UPDATED
    def add_event(self, project_id: int, status: str, created_at: str, message: str = "", error_code: str = "") -> None:
        with self.lock:
            projects, events = self._read_all()
            pr = projects[projects["id"] == project_id]
            project_name = pr.iloc[0]["name"] if not pr.empty else f"#{project_id}"
            eid = self._next_id(events)
            new_row = {
                "id": eid,
                "project_id": int(project_id),
                "project": str(project_name),
                "status": status,
                "created_at": created_at,
                "message": message or "",
            }
            events = pd.concat([events, pd.DataFrame([new_row])], ignore_index=True)
            self._write_all(projects, events)

            # --- 專案日誌 CSV：改新表頭與欄位順序 ---
            try:
                proj_dir = _project_logs_dir_on_D(project_name)
                os.makedirs(proj_dir, exist_ok=True)
                dt = pd.to_datetime(created_at, errors="coerce")
                date_str = (dt.strftime("%Y%m%d") if pd.notna(dt) else datetime.now().strftime("%Y%m%d"))
                csv_path = os.path.join(proj_dir, f"{project_name}_{date_str}.csv")
                csv_path = os.path.normpath(csv_path)
                write_header = not os.path.exists(csv_path)

                # 若外部未傳入，這裡仍做一次保底解析（不觸發 UI）
                _ec = (error_code or "")
                if not _ec:
                    _ec = ERROR_RESOLVER.resolve(status, message)

                with open(csv_path, "a", encoding="utf-8-sig") as f:
                    if write_header:
                        f.write("event_id,created_at,status,trigger,message,error_code\n")

                    # 由 message 前綴萃取 trigger（與原寫法一致邏輯）
                    low_msg = (message or "").strip().lower()
                    if low_msg.startswith("manual:"):
                        trigger = "manual"
                    elif low_msg.startswith("scheduler:"):
                        trigger = "scheduler"
                    else:
                        trigger = "unknown"

                    safe_msg = (message or "").replace("\n", " ").replace("\r", " ")
                    f.write(f"{eid},{created_at},{status},{trigger},{safe_msg},{_ec}\n")
            except Exception:
                LOGGER.debug("project CSV log append failed", exc_info=True)

            LOGGER.info("Event added: id=%s project_id=%s status=%s", eid, project_id, status)

    def delete_event(self, event_id: int) -> None:
        with self.lock:
            projects, events = self._read_all()
            events = events[events["id"] != event_id].reset_index(drop=True)
            self._write_all(projects, events)
            LOGGER.info("Event deleted: id=%s", event_id)

    # -------------------- KPI/彙總（過去 7 天） --------------------
    def summary_last_7_days(self) -> List[Tuple[str, int, int, int]]:
        """回傳各專案過去 7 天的 (name, success_7d, fail_7d, total_7d)
        # UPDATED: 成功只計「兩段皆成功」→ 等價於「message 含 upload: 的 success」
                  兼容舊資料：沒有 'download:' / 'upload:' 前綴但為 success 也視為整體成功
                  失敗：計所有 fail（包含 download 失敗或 upload 失敗）
        """
        with self.lock:
            projects, events = self._read_all()
            if projects.empty:
                return []
            since = (datetime.now() - timedelta(days=7)).replace(microsecond=0)
            if events.empty:
                result: List[Tuple[str, int, int, int]] = []
                for _, p in projects.sort_values(by="name", kind="stable").iterrows():
                    result.append((p["name"], 0, 0, 0))
                return result

            ev = events.copy()
            ev["created_at_dt"] = pd.to_datetime(ev["created_at"], errors="coerce")
            ev = ev[ev["created_at_dt"].notna()]
            ev = ev[ev["created_at_dt"] >= pd.Timestamp(since)]
            if ev.empty:
                result: List[Tuple[str, int, int, int]] = []
                for _, p in projects.sort_values(by="name", kind="stable").iterrows():
                    result.append((p["name"], 0, 0, 0))
                return result

            msg_lc = ev["message"].astype(str).str.lower()
            is_success_upload = (ev["status"] == "success") & (msg_lc.str.contains("upload:", na=False))
            is_success_legacy = (ev["status"] == "success") & (~msg_lc.str.contains("upload:", na=False)) & (~msg_lc.str.contains("download:", na=False))
            is_fail_any = (ev["status"] == "fail")

            succ = ev[is_success_upload | is_success_legacy].groupby("project_id")["id"].count()
            fail = ev[is_fail_any].groupby("project_id")["id"].count()

            merged = projects[["id", "name"]].copy()
            merged = merged.merge(succ.rename("success"), left_on="id", right_index=True, how="left")
            merged = merged.merge(fail.rename("fail"), left_on="id", right_index=True, how="left")
            merged[["success", "fail"]] = merged[["success", "fail"]].fillna(0).astype(int)
            merged["total"] = merged["success"] + merged["fail"]

            merged = merged.sort_values(by=["total", "success", "name"], ascending=[False, False, True], kind="stable")
            return [(r["name"], int(r["success"]), int(r["fail"]), int(r["total"])) for _, r in merged.iterrows()]
    # ----------------------------------------------------------------

    # -------------------- KPI/彙總（今日） --------------------
    def summary_today(self) -> List[Tuple[str, int, int, int]]:
        """回傳各專案『今天』(本機時區) 的 (name, success_today, fail_today, total_today)
        # UPDATED: 成功只計 upload 成功（或舊資料無前綴之 success），失敗計所有 fail
        """
        with self.lock:
            projects, events = self._read_all()
            if projects.empty:
                return []
            today_date = datetime.now().date()
            if events.empty:
                result: List[Tuple[str, int, int, int]] = []
                for _, p in projects.sort_values(by="name", kind="stable").iterrows():
                    result.append((p["name"], 0, 0, 0))
                return result

            ev = events.copy()
            ev["created_at_dt"] = pd.to_datetime(ev["created_at"], errors="coerce")
            ev = ev[ev["created_at_dt"].notna()]
            ev = ev[ev["created_at_dt"].dt.date == today_date]
            if ev.empty:
                result: List[Tuple[str, int, int, int]] = []
                for _, p in projects.sort_values(by="name", kind="stable").iterrows():
                    result.append((p["name"], 0, 0, 0))
                return result

            msg_lc = ev["message"].astype(str).str.lower()
            is_success_upload = (ev["status"] == "success") & (msg_lc.str.contains("upload:", na=False))
            is_success_legacy = (ev["status"] == "success") & (~msg_lc.str.contains("upload:", na=False)) & (~msg_lc.str.contains("download:", na=False))
            is_fail_any = (ev["status"] == "fail")

            succ = ev[is_success_upload | is_success_legacy].groupby("project_id")["id"].count()
            fail = ev[is_fail_any].groupby("project_id")["id"].count()

            merged = projects[["id", "name"]].copy()
            merged = merged.merge(succ.rename("success"), left_on="id", right_index=True, how="left")
            merged = merged.merge(fail.rename("fail"), left_on="id", right_index=True, how="left")
            merged[["success", "fail"]] = merged[["success", "fail"]].fillna(0).astype(int)
            merged["total"] = merged["success"] + merged["fail"]

            # 不在此排序（排序在 UI 側依 7d 規則）
            return [(r["name"], int(r["success"]), int(r["fail"]), int(r["total"])) for _, r in merged.iterrows()]
    # ----------------------------------------------------------------

@dataclass
class ActionSpec:
    command: str
    args: List[str]
    cwd: Optional[str]

class SettingsLoader:
    def __init__(self, ini_path: str = "Setting.ini") -> None:
        self.ini_path = ini_path
    def _read_cfg(self) -> Dict[str, ActionSpec]:
        if not os.path.exists(self.ini_path):
            return {}
        cfg = configparser.ConfigParser(interpolation=None)
        try:
            cfg.read(self.ini_path, encoding="utf-8")
        except Exception:
            enc = locale.getpreferredencoding(False) or "utf-8"
            cfg.read(self.ini_path, encoding=enc)
        mapping: Dict[str, ActionSpec] = {}
        for section in cfg.sections():
            if section.lower() == "global":
                continue
            raw_cmd = cfg.get(section, "command", fallback="").strip()
            raw_cwd = cfg.get(section, "cwd", fallback="").strip()
            raw_args = cfg.get(section, "args", fallback="").strip()
            def _strip_quotes(s: str) -> str:
                return s[1:-1] if (len(s) >= 2 and ((s[0] == s[-1] == '"') or (s[0] == s[-1] == "'"))) else s
            cmd = _strip_quotes(raw_cmd)
            cwd = _strip_quotes(raw_cwd) or None
            args_list: List[str] = shlex.split(raw_args) if raw_args else []
            if cmd:
                mapping[section] = ActionSpec(command=cmd, args=args_list, cwd=cwd)
        return mapping
    def get_action(self, project_name: str) -> Optional[ActionSpec]:
        cfg = self._read_cfg()
        return cfg.get(project_name)

class SimpleTableModel(QAbstractTableModel):
    """極簡 TableModel：只讀、置中對齊（第一欄靠左）。"""
    def __init__(self, headers: List[str], rows: Optional[List[List[Any]]] = None) -> None:
        super().__init__()
        self.headers: List[str] = headers
        self.rows: List[List[Any]] = rows or []
    def set_rows(self, rows: List[List[Any]]) -> None:
        self.beginResetModel()
        self.rows = rows
        self.endResetModel()
    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # type: ignore[override]
        return len(self.rows)
    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # type: ignore[override]
        return len(self.headers)
    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:  # type: ignore[override]
        if not index.isValid():
            return None
        r, c = index.row(), index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            val = self.rows[r][c]
            return "" if val is None else str(val)
        if role == Qt.ItemDataRole.TextAlignmentRole:
            if c == 0:
                return int(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft)
            return int(Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignCenter)
        return None
    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:  # type: ignore[override]
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return self.headers[section]
        return str(section + 1)

class BatFileDropLineEdit(QLineEdit):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()
    def dropEvent(self, event) -> None:
        urls = [u for u in event.mimeData().urls()] if event.mimeData().hasUrls() else []
        if not urls:
            event.ignore()
            return
        url: QUrl = urls[0]
        path = url.toLocalFile() or url.path()
        if not path:
            event.ignore()
            return
        try:
            path = os.path.normpath(path)
            if os.path.isdir(path):
                self._notify_status("僅接受拖入 .bat 檔；資料夾已忽略。")
                event.ignore()
                return
            _, ext = os.path.splitext(path)
            if ext.lower() == ".bat":
                self.setText(path)
                self._notify_status(f"已載入 .bat：{path}")
                event.acceptProposedAction()
                return
            self._notify_status("僅接受拖入 .bat 檔。")
            event.ignore()
        except Exception as e:
            self._notify_status(f"拖放處理失敗：{e}")
            event.ignore()
    def _notify_status(self, msg: str) -> None:
        w = self.window()
        if isinstance(w, QMainWindow):
            w.statusBar().showMessage(msg, 5000)


# ====================== LogCleaner ======================  # NEW
# ====================== LogCleaner ======================  # UPDATED
class LogCleaner:
    r"""
    掃描 LOG_ROOT 下所有 <Project>\logs\ ，清理：
    1) CSV：<Project>_YYYYMMDD.csv（僅此規格）
    2) RUN-LOG：
       - 全部 *.log
       - 無副檔名的 ..._stdout / ..._stderr
    日期優先取自檔名（_YYYYMMDD_），否則退回檔案 mtime。
    """
    FILE_PATTERN_CSV = re.compile(r"^(?P<proj>.+)_(?P<date>\d{8})\.csv$", re.IGNORECASE)
    FILE_PATTERN_DATE = re.compile(r"_(?P<date>\d{8})(?:_|\.|$)")  # 用於 *.log / 無副檔名 stdout/stderr
    FILE_PATTERN_STDSTREAM_BARE = re.compile(r"_(stdout|stderr)$", re.IGNORECASE)

    def __init__(self, root: str | None = None):
        # 避免類別定義期引用常數名稱順序問題
        self.root = os.path.normpath(root or LOG_ROOT)

    def _iter_project_log_files(self):
        """
        yield (project, full_path, date_obj, kind)
        kind in {"csv", "runlog"}；異常/不符命名者以 (None, path, None, None) 回報。
        """
        if not os.path.isdir(self.root):
            return

        for entry in os.scandir(self.root):
            if not entry.is_dir():
                continue
            proj = entry.name
            log_dir = os.path.join(self.root, proj, "logs")
            if not os.path.isdir(log_dir):
                continue

            for f in os.scandir(log_dir):
                if not f.is_file():
                    continue
                path = f.path
                name = f.name

                # 1) CSV：只吃規格 <Project>_YYYYMMDD.csv
                if name.lower().endswith(".csv"):
                    m = self.FILE_PATTERN_CSV.match(name)
                    if not m:
                        yield (None, path, None, None)  # 非規格 CSV → 忽略
                        continue
                    p2 = m.group("proj")
                    d8 = m.group("date")
                    if p2 != proj:
                        yield (None, path, None, None)  # 前綴不一致 → 忽略
                        continue
                    try:
                        dt = _dt.strptime(d8, "%Y%m%d").date()
                    except Exception:
                        yield (None, path, None, None)
                        continue
                    yield (proj, path, dt, "csv")
                    continue

                low = name.lower()

                # 2) RUN-LOG：全部 *.log
                if low.endswith(".log"):
                    # 檔名擷取日期；若無 → 用 mtime
                    m = self.FILE_PATTERN_DATE.search(name)
                    if m:
                        try:
                            dt = _dt.strptime(m.group("date"), "%Y%m%d").date()
                        except Exception:
                            dt = _dt.fromtimestamp(os.path.getmtime(path)).date()
                    else:
                        dt = _dt.fromtimestamp(os.path.getmtime(path)).date()
                    yield (proj, path, dt, "runlog")
                    continue

                # 3) 無副檔名 stdout/stderr（…_stdout / …_stderr）
                if self.FILE_PATTERN_STDSTREAM_BARE.search(low):
                    m = self.FILE_PATTERN_DATE.search(name)
                    if m:
                        try:
                            dt = _dt.strptime(m.group("date"), "%Y%m%d").date()
                        except Exception:
                            dt = _dt.fromtimestamp(os.path.getmtime(path)).date()
                    else:
                        dt = _dt.fromtimestamp(os.path.getmtime(path)).date()
                    yield (proj, path, dt, "runlog")
                    continue

                # 其他檔案 → 忽略
                yield (None, path, None, None)

    @staticmethod
    def _fmt_size(n: int) -> str:
        units = ["B", "KB", "MB", "GB", "TB"]
        i = 0
        x = float(n)
        while x >= 1024 and i < len(units) - 1:
            x /= 1024.0
            i += 1
        return f"{x:.2f} {units[i]}"

    def preview(self, retention_days: int):
        """
        回傳 dict：
        {
          "candidates": [ (proj, path, date_str, size) ... ],
          "ignored":    [ path, ... ],
          "total_size": int(bytes),
          "count":      int
        }
        """
        today = _dt.today().date()
        keep_after = today - timedelta(days=retention_days - 1)
        cands, ig, total = [], [], 0

        for proj, path, dt, kind in self._iter_project_log_files():
            if proj is None or dt is None:
                ig.append(path)
                continue
            # 小於 keep_after → 刪除；等於/大於 → 保留
            if dt < keep_after:
                try:
                    size = os.path.getsize(path)
                except Exception:
                    size = 0
                cands.append((proj, path, dt.strftime("%Y-%m-%d"), size))
                total += size

        return {
            "candidates": cands,
            "ignored": ig,
            "total_size": total,
            "count": len(cands),
        }

    def delete(self, file_list):
        """實際刪除：file_list 來自 preview()['candidates'] 的 path 欄位"""
        ok, fail, bytes_freed = 0, 0, 0
        for item in file_list:
            path = item[1] if isinstance(item, tuple) else item
            try:
                sz = os.path.getsize(path)
            except Exception:
                sz = 0
            try:
                os.remove(path)
                ok += 1
                bytes_freed += sz
            except Exception:
                fail += 1
        return {"deleted": ok, "failed": fail, "bytes_freed": bytes_freed}
# ==================== /LogCleaner =======================

# ====================== LogCleanupDialog ======================  # NEW
class LogCleanupDialog(QDialog):
    def __init__(self, parent=None, settings: QSettings | None = None):
        super().__init__(parent)
        self.setWindowTitle("日誌清理")
        self.setModal(True)
        self.settings = settings or QSettings(ORG_NAME, APP_NAME)
        self.cleaner = LogCleaner(LOG_ROOT)

        # 介面
        v = QVBoxLayout(self)

        # 根路徑與說明
        v.addWidget(QLabel(f"根路徑：{LOG_ROOT}"))
        v.addWidget(QLabel("說明：會掃描 D:\\<Project>\\logs\\ 內符合 <Project>_YYYYMMDD.csv 的日誌檔；"
                           "不符合命名規則者將顯示為「忽略」。"))

        # 手動清理區
        hv1 = QHBoxLayout()
        hv1.addWidget(QLabel("保留最近"))
        self.sb_days = QSpinBox()
        self.sb_days.setRange(1, 3650)
        self.sb_days.setValue(self.settings.value(SETK_RETENTION_DAYS, DEFAULT_RETENTION, int))
        hv1.addWidget(self.sb_days)
        hv1.addWidget(QLabel("天"))
        self.btn_preview = QPushButton("預覽")
        self.btn_execute = QPushButton("執行清理")
        self.btn_open = QPushButton("開啟日誌根目錄")
        hv1.addWidget(self.btn_preview)
        hv1.addWidget(self.btn_execute)
        hv1.addWidget(self.btn_open)
        hv1.addStretch()
        v.addLayout(hv1)

        # 自動清理區
        self.chk_auto = QCheckBox("啟用自動清理（程式啟動時執行）")
        self.chk_auto.setChecked(self.settings.value(SETK_AUTO_ENABLED, False, bool))
        v.addWidget(self.chk_auto)

        # 記住選擇
        self.chk_skip_preview = QCheckBox("刪除前不再顯示預覽（記住我的選擇）")
        self.chk_skip_preview.setChecked(self.settings.value(SETK_SKIP_PREVIEW, False, bool))
        v.addWidget(self.chk_skip_preview)

        # 結果輸出
        self.txt = QTextEdit()
        self.txt.setReadOnly(True)
        self.txt.setMinimumHeight(240)
        v.addWidget(self.txt)

        # 底部
        hb = QHBoxLayout()
        self.btn_close = QPushButton("關閉")
        hb.addStretch()
        hb.addWidget(self.btn_close)
        v.addLayout(hb)

        # 事件
        self.btn_open.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(LOG_ROOT)))
        self.btn_preview.clicked.connect(self.on_preview)
        self.btn_execute.clicked.connect(self.on_execute)
        self.btn_close.clicked.connect(self.accept)
        self.chk_auto.stateChanged.connect(self.on_settings_changed)
        self.sb_days.valueChanged.connect(self.on_settings_changed)
        self.chk_skip_preview.stateChanged.connect(self.on_settings_changed)

        self._last_preview = None

    def on_settings_changed(self, *args):
        self.settings.setValue(SETK_AUTO_ENABLED, bool(self.chk_auto.isChecked()))
        self.settings.setValue(SETK_RETENTION_DAYS, int(self.sb_days.value()))
        self.settings.setValue(SETK_SKIP_PREVIEW, bool(self.chk_skip_preview.isChecked()))
        self.settings.sync()

    def on_preview(self):
        rdays = int(self.sb_days.value())
        rep = self.cleaner.preview(rdays)
        self._last_preview = rep
        size_str = self.cleaner._fmt_size(rep["total_size"])
        lines = [f"將刪除檔案：{rep['count']} 個，總計 {size_str}"]
        if rep["count"] > 0:
            lines.append("— 列表 —")
            for proj, path, d, sz in rep["candidates"]:
                lines.append(f"[{proj}] {os.path.basename(path)}  ({d}, {self.cleaner._fmt_size(sz)})")
        if rep["ignored"]:
            lines.append("")
            lines.append(f"忽略（檔名非規格或日期解析失敗）：{len(rep['ignored'])} 個")
        self.txt.setPlainText("\n".join(lines))

    def on_execute(self):
        rdays = int(self.sb_days.value())
        # 若未預覽或使用者勾了「不再顯示預覽」
        need_preview = not bool(self.settings.value(SETK_SKIP_PREVIEW, False, bool))
        if (self._last_preview is None or need_preview) and not self._confirm_preview_then_cache(rdays):
            return
        rep = self._last_preview or self.cleaner.preview(rdays)
        paths = rep["candidates"]
        if not paths:
            QMessageBox.information(self, "日誌清理", "沒有可刪除的檔案。")
            return
        # 執行刪除
        stat = self.cleaner.delete(paths)
        freed = self.cleaner._fmt_size(stat["bytes_freed"])
        self.txt.append("\n[執行結果]")
        self.txt.append(f"刪除：{stat['deleted']}，失敗：{stat['failed']}，釋放空間：{freed}")
        # 使預覽失效（避免重用舊清單）
        self._last_preview = None

    def _confirm_preview_then_cache(self, rdays: int) -> bool:
        """在未關閉預覽的情況下，先跑一次預覽，讓使用者看過再刪。"""
        self.on_preview()
        rep = self._last_preview
        size_str = self.cleaner._fmt_size(rep["total_size"])
        msg = f"將刪除 {rep['count']} 個檔案，總計 {size_str}。\n\n是否繼續？"
        r = QMessageBox.question(self, "日誌清理確認", msg,
                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                 QMessageBox.StandardButton.No)
        return r == QMessageBox.StandardButton.Yes
# ==================== /LogCleanupDialog =======================

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Upload Dashboard + Interval Scheduler (Excel) - PyQt6")
        self.resize(1200, 780)
        self.store = ExcelStore(EXCEL_PATH)
        self._build_menu()
        self._build_ui()
        self.timer = QTimer(self)
        self.timer.setInterval(15_000)
        self.timer.timeout.connect(self._safe_check_intervals)
        self.timer.start()
        self.refresh_all()
        self._unknown_warned_keys: set[str] = set()  # NEW：避免 E000 對話框洗版
        self.settings = QSettings(ORG_NAME, APP_NAME)
        QTimer.singleShot(1500, self._maybe_auto_cleanup)


    def _build_menu(self) -> None:
        menubar = self.menuBar()
        file_menu = menubar.addMenu("File")
        act_export = QAction("Export Summary CSV...", self)
        act_export.triggered.connect(self.export_summary_csv)
        file_menu.addAction(act_export)
        act_refresh = QAction("Refresh", self)
        act_refresh.triggered.connect(self.refresh_all)
        file_menu.addAction(act_refresh)
        file_menu.addSeparator()
        act_exit = QAction("Exit", self)
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)
        # --- 在 _build_menu() 中新增一個 Rules / 規則 選單 ---  # NEW
        rules_menu = menubar.addMenu("Rules")  # NEW

        act_open_rules = QAction("開啟關鍵字規則檔…", self)  # NEW
        # Tools / 工具 選單
        tools_menu = None
        for a in menubar.actions():
            if a.text().replace("&","") in ("工具","Tools"):
                tools_menu = a.menu()
                break
        if tools_menu is None:
            tools_menu = menubar.addMenu("工具")

        act_cleanup = QAction("日誌清理…", self)
        def _open_cleanup():
            dlg = LogCleanupDialog(self, settings=self.settings)
            dlg.exec()
        act_cleanup.triggered.connect(_open_cleanup)
        tools_menu.addAction(act_cleanup)

        def _open_rules():  # NEW
            path = os.path.abspath(ERROR_KEYWORDS_CSV)
            # 若檔案不存在先建立（含表頭）
            if not os.path.exists(path):
                try:
                    with open(path, "w", encoding="utf-8-sig", newline="") as f:
                        w = csv.writer(f)
                        w.writerow(["error_code", "keyword"])
                except Exception:
                    LOGGER.debug("create empty error_keywords.csv failed", exc_info=True)
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))
        act_open_rules.triggered.connect(_open_rules)
        rules_menu.addAction(act_open_rules)

        act_reload_rules = QAction("重新載入規則", self)  # NEW
        def _reload_rules():  # NEW
            n = ERROR_RESOLVER.load()
            self.statusBar().showMessage(f"已重新載入錯誤關鍵字規則：{n} 筆", 5000)
        act_reload_rules.triggered.connect(_reload_rules)
        rules_menu.addAction(act_reload_rules)

        act_help_rules = QAction("規則說明", self)  # NEW
        def _help_rules():  # NEW
            txt = (
                "【error_keywords.csv】\n"
                "• 位置：程式同層；編碼 UTF-8（含 BOM）。\n"
                "• 欄位（兩欄）：error_code, keyword。\n"
                "• 比對：不分大小寫；訊息凡包含 keyword 即命中；由上而下第一個命中即用。\n"
                "• 若完全無命中，會套用內建保底規則；仍無則指派 E000，並跳提醒。\n\n"
                "常見範例：\n"
                "  E001, 逾時（>\n"
                "  E003, 指定的 .bat 不存在\n"
                "  E005, return_code=\n"
            )
            QMessageBox.information(self, "規則說明", txt)
        act_help_rules.triggered.connect(_help_rules)
        rules_menu.addAction(act_help_rules)
        
        # --- 在 _build_menu() 末端附近加入 ---  # NEW
        help_menu = menubar.addMenu("說明")  # NEW

        act_about = QAction("關於…", self)  # NEW
        def _about_dialog():  # NEW
            # 解析專案根目錄（exe onedir 同層 / 原始碼模式兩者皆可）
            if getattr(sys, "frozen", False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))

            changelog_path = os.path.join(base_dir, "CHANGELOG.md")
            changelog_exists = os.path.isfile(changelog_path)

            msg = (
                f"UploadDashboard\n"
                f"版本：{__version__}\n\n"
                f"• 專案首頁：本機目錄\n"
                f"• 變更記錄：{'已檢測到' if changelog_exists else '未找到'} CHANGELOG.md"
            )
            btns = QMessageBox.StandardButton.Ok
            if changelog_exists:
                # 追加「開啟 CHANGELOG」按鈕
                btns |= QMessageBox.StandardButton.Open

            r = QMessageBox.information(
                self, "關於 UploadDashboard", msg, btns, QMessageBox.StandardButton.Ok
            )
            if changelog_exists and r == QMessageBox.StandardButton.Open:
                QDesktopServices.openUrl(QUrl.fromLocalFile(changelog_path))

        act_about.triggered.connect(_about_dialog)
        help_menu.addAction(act_about)

    def _build_ui(self) -> None:
        tabs = QTabWidget()
        self.setCentralWidget(tabs)

        tab_dashboard = QWidget()
        tabs.addTab(tab_dashboard, "Dashboard")
        dash_layout = QVBoxLayout(tab_dashboard)

        kpi = QGroupBox("KPI（過去 7 天）")
        kpi_layout = QHBoxLayout(kpi)
        self.lbl_success = QLabel("總成功：0")
        self.lbl_fail = QLabel("總失敗：0")
        self.lbl_total = QLabel("總上傳：0")
        self.lbl_fail_rate = QLabel("失敗率：0.0%")
        for lab in (self.lbl_success, self.lbl_fail, self.lbl_total, self.lbl_fail_rate):
            lab.setStyleSheet("font-size: 14px;")
        kpi_layout.addWidget(self.lbl_success)
        kpi_layout.addWidget(self.lbl_fail)
        kpi_layout.addWidget(self.lbl_total)
        kpi_layout.addWidget(self.lbl_fail_rate)
        kpi_layout.addStretch(1)
        dash_layout.addWidget(kpi)

        kpi_today = QGroupBox("KPI（今日）")
        kpi_today_layout = QHBoxLayout(kpi_today)
        self.lbl_today_success = QLabel("今日成功：0")
        self.lbl_today_fail = QLabel("今日失敗：0")
        self.lbl_today_total = QLabel("今日上傳：0")
        self.lbl_today_fail_rate = QLabel("今日失敗率：0.0%")
        for lab in (self.lbl_today_success, self.lbl_today_fail, self.lbl_today_total, self.lbl_today_fail_rate):
            lab.setStyleSheet("font-size: 14px;")
        kpi_today_layout.addWidget(self.lbl_today_success)
        kpi_today_layout.addWidget(self.lbl_today_fail)
        kpi_today_layout.addWidget(self.lbl_today_total)
        kpi_today_layout.addWidget(self.lbl_today_fail_rate)
        kpi_today_layout.addStretch(1)
        dash_layout.addWidget(kpi_today)

        splitter = QSplitter(Qt.Orientation.Vertical)
        dash_layout.addWidget(splitter, 1)

        sum_group = QGroupBox("各專案統計（過去 7 天）")
        sum_layout = QVBoxLayout(sum_group)
        self.summary_model = SimpleTableModel(
            ["Project", "Success(7d)", "Fail(7d)", "Total(7d)", "Success(Today)", "Fail(Today)", "Total(Today)"], [])
        self.tbl_summary = QTableView()
        self.tbl_summary.setModel(self.summary_model)
        self.tbl_summary.setAlternatingRowColors(True)
        self.tbl_summary.horizontalHeader().setStretchLastSection(True)
        sum_layout.addWidget(self.tbl_summary)

        btn_row = QHBoxLayout()
        self.btn_refresh = QPushButton("Refresh")
        self.btn_refresh.clicked.connect(self.refresh_all)
        self.btn_export = QPushButton("Export Summary CSV")
        self.btn_export.clicked.connect(self.export_summary_csv)
        btn_row.addWidget(self.btn_refresh)
        btn_row.addWidget(self.btn_export)
        btn_row.addStretch(1)
        sum_layout.addLayout(btn_row)
        splitter.addWidget(sum_group)

        evt_group = QGroupBox("最近事件（可刪除）")
        evt_layout = QVBoxLayout(evt_group)
        limit_row = QHBoxLayout()
        limit_row.addWidget(QLabel("顯示筆數："))
        self.spin_limit = QSpinBox()
        self.spin_limit.setRange(10, 500)
        self.spin_limit.setValue(50)
        self.spin_limit.valueChanged.connect(self.refresh_events)
        limit_row.addWidget(self.spin_limit)
        limit_row.addStretch(1)
        self.btn_delete_event = QPushButton("刪除選取事件")
        self.btn_delete_event.clicked.connect(self.delete_selected_event)
        limit_row.addWidget(self.btn_delete_event)
        evt_layout.addLayout(limit_row)

        self.events_model = SimpleTableModel(["ID", "Project", "Status", "Created At", "Message"], [])
        self.tbl_events = QTableView()
        self.tbl_events.setModel(self.events_model)
        self.tbl_events.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.tbl_events.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.tbl_events.setAlternatingRowColors(True)
        self.tbl_events.horizontalHeader().setStretchLastSection(True)
        evt_layout.addWidget(self.tbl_events)
        splitter.addWidget(evt_group)
        splitter.setSizes([420, 300])

        tab_projects = QWidget()
        tabs.addTab(tab_projects, "Projects & Scheduler")
        proj_layout = QHBoxLayout(tab_projects)

        left = QGroupBox("專案列表（可新增/刪除）")
        left_layout = QVBoxLayout(left)
        self.projects_model = SimpleTableModel(
            ["ID", "Name", "Interval(min)", "Enabled", "Last Run At", "Next Run At"], []
        )
        self.tbl_projects = QTableView()
        self.tbl_projects.setModel(self.projects_model)
        self.tbl_projects.setSelectionBehavior(QTableView.SelectionBehavior.SelectRows)
        self.tbl_projects.setSelectionMode(QTableView.SelectionMode.SingleSelection)
        self.tbl_projects.setAlternatingRowColors(True)
        self.tbl_projects.horizontalHeader().setStretchLastSection(True)
        self.tbl_projects.clicked.connect(self.on_project_selected)
        left_layout.addWidget(self.tbl_projects)

        proj_btns = QHBoxLayout()
        self.btn_add_project = QPushButton("新增專案")
        self.btn_add_project.clicked.connect(self.add_project)
        self.btn_delete_project = QPushButton("刪除專案")
        self.btn_delete_project.clicked.connect(self.delete_project)
        self.btn_reload_projects = QPushButton("重新整理")
        self.btn_reload_projects.clicked.connect(self.refresh_projects)
        proj_btns.addWidget(self.btn_add_project)
        proj_btns.addWidget(self.btn_delete_project)
        proj_btns.addWidget(self.btn_reload_projects)
        proj_btns.addStretch(1)
        left_layout.addLayout(proj_btns)

        right = QGroupBox("專案設定（Interval / 啟用 / 手動觸發）")
        right_layout = QVBoxLayout(right)

        form = QFormLayout()
        self.lbl_project_id = QLabel("-")
        self.edit_name = QLineEdit()
        self.spin_interval = QSpinBox()
        self.spin_interval.setRange(1, 100000)
        self.spin_interval.setValue(60)
        self.chk_enabled = QCheckBox("啟用定時上傳")
        self.lbl_last_run = QLabel("-")
        self.lbl_next_run = QLabel("-")

        form.addRow("Project ID:", self.lbl_project_id)
        form.addRow("Name:", self.edit_name)
        form.addRow("Interval Minutes:", self.spin_interval)
        form.addRow("", self.chk_enabled)
        form.addRow("Last Run At:", self.lbl_last_run)
        form.addRow("Next Run At:", self.lbl_next_run)

        dl_row = QHBoxLayout()
        self.edit_download_bat = BatFileDropLineEdit(self)
        self.btn_browse_download_bat = QPushButton("瀏覽")
        self.btn_browse_download_bat.clicked.connect(lambda: self._browse_bat(self.edit_download_bat))
        dl_row.addWidget(self.edit_download_bat)
        dl_row.addWidget(self.btn_browse_download_bat)
        form.addRow("Download BAT", dl_row)

        ul_row = QHBoxLayout()
        self.edit_upload_bat = BatFileDropLineEdit(self)
        self.btn_browse_upload_bat = QPushButton("瀏覽")
        self.btn_browse_upload_bat.clicked.connect(lambda: self._browse_bat(self.edit_upload_bat))
        ul_row.addWidget(self.edit_upload_bat)
        ul_row.addWidget(self.btn_browse_upload_bat)
        form.addRow("Upload BAT", ul_row)

        right_layout.addLayout(form)

        action_row = QHBoxLayout()
        self.btn_save_project = QPushButton("儲存設定")
        self.btn_save_project.clicked.connect(self.save_project_settings)
        self.btn_run_now = QPushButton("立即上傳（手動觸發）")
        self.btn_run_now.clicked.connect(self.run_selected_now)  # 兩段式：Download -> Upload
        self.btn_reset_last_run = QPushButton("重置 last_run（下次檢查立刻跑）")
        self.btn_reset_last_run.clicked.connect(self.reset_last_run)
        action_row.addWidget(self.btn_save_project)
        action_row.addWidget(self.btn_run_now)
        action_row.addWidget(self.btn_reset_last_run)
        action_row.addStretch(1)
        right_layout.addLayout(action_row)

       # UPDATED: 說明文字加入 nan 跳過機制，並啟用自動換行
        note = QLabel(
            "Excel 模式說明：\n"
            "• 資料存於 upload_stats.xlsx（Projects / Events 兩張表）。\n"
            "• Interval 模式：每 N 分鐘執行一次（依 last_run_at 計算）。\n"
            "• Excel 寫入會重寫整個檔案，頻率太高會變慢；建議事件量不要爆量。\n"
            "\n"
            "nan 跳過機制（Download/Upload BAT）：\n"
            "• 在 Download BAT 或 Upload BAT 欄位輸入 'nan'（不含括號；大小寫不分；可有前後空白），該階段會「跳過驗證與執行」。\n"
            "• 跳過時會直接記錄一筆 success 事件：\n"
            "  - download: skipped (nan)\n"
            "  - upload: skipped (nan)（此筆會被 KPI 視為成功）\n"
            "• 僅一段為 nan：整體成功/失敗依另一段的實際執行結果而定。\n"
            "• 兩段皆為 nan：整體直接成功（兩段皆記一筆 skipped 事件）。\n"
            "• 非 'nan' 且路徑無效或檔案不存在：仍視為失敗（不會自動跳過）。\n"
        )
        note.setWordWrap(True)              # NEW：確保多行說明可換行
        note.setStyleSheet("color:#555;")
        right_layout.addWidget(note)

        proj_layout.addWidget(left, 2)
        proj_layout.addWidget(right, 1)
        # 設定視窗圖示（相對路徑 assets/monitor_dashboard_icon_136391.ico）  # NEW
        try:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            ico_path = os.path.join(base_dir, "assets", "monitor_dashboard_icon_136391.ico")
            if os.path.isfile(ico_path):
                self.setWindowIcon(QIcon(ico_path))
        except Exception:
            pass
        self.statusBar().showMessage("Ready")

    def _maybe_auto_cleanup(self):
        try:
            if not self.settings.value(SETK_AUTO_ENABLED, False, bool):
                return
            rdays = self.settings.value(SETK_RETENTION_DAYS, DEFAULT_RETENTION, int)
            cleaner = LogCleaner(LOG_ROOT)
            rep = cleaner.preview(rdays)
            if rep["count"] > 0:
                stat = cleaner.delete(rep["candidates"])
                freed = cleaner._fmt_size(stat["bytes_freed"])
                self.statusBar().showMessage(
                    f"自動清理完成：刪除 {stat['deleted']} 個，釋放 {freed}", 7000
                )
        except Exception:
            # 靜默處理，避免影響啟動
            pass

    def _browse_bat(self, target_edit: QLineEdit) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "選取 .bat", "", "Batch Files (*.bat)")
        if not path:
            return
        path = os.path.normpath(path)
        target_edit.setText(path)
        self.statusBar().showMessage(f"已選取：{path}", 5000)

    def refresh_all(self) -> None:
        self.refresh_projects()
        self.refresh_summary()
        self.refresh_events()

    def refresh_projects(self) -> None:
        rows = self.store.list_projects()
        view_rows: List[List[Any]] = []
        now_dt = datetime.now().replace(microsecond=0)
        for r in rows:
            pid = safe_int(r.get("id"), 0)
            name = str(r.get("name", ""))
            interval = safe_int(r.get("interval_minutes"), 60)
            enabled = bool(safe_int(r.get("enabled"), 0))
            last_run_at = str(r.get("last_run_at", "") or "")
            if last_run_at.strip():
                try:
                    nxt = parse_dt(last_run_at) + timedelta(minutes=interval)
                    next_run_at = nxt.replace(microsecond=0).isoformat(sep=" ")
                except Exception:
                    next_run_at = ""
            else:
                next_run_at = now_dt.isoformat(sep=" ")
            view_rows.append([pid, name, interval, "ON" if enabled else "OFF", last_run_at, next_run_at])
        self.projects_model.set_rows(view_rows)

    def refresh_summary(self) -> None:
        summary_7d = self.store.summary_last_7_days()
        summary_today = self.store.summary_today()
        map_7d = {name: (succ7, fail7, total7) for name, succ7, fail7, total7 in summary_7d}
        map_today = {name: (succT, failT, totalT) for name, succT, failT, totalT in summary_today}
        rows: list[list[object]] = []
        for name, succ7, fail7, total7 in summary_7d:
            succT, failT, totalT = map_today.get(name, (0, 0, 0))
            rows.append([name, succ7, fail7, total7, succT, failT, totalT])
        self.summary_model.set_rows(rows)

        total_success_7d = sum(r[1] for r in summary_7d)
        total_fail_7d = sum(r[2] for r in summary_7d)
        total_all_7d = total_success_7d + total_fail_7d
        fail_rate_7d = (total_fail_7d / total_all_7d * 100.0) if total_all_7d else 0.0
        self.lbl_success.setText(f"總成功：{total_success_7d}")
        self.lbl_fail.setText(f"總失敗：{total_fail_7d}")
        self.lbl_total.setText(f"總上傳：{total_all_7d}")
        self.lbl_fail_rate.setText(f"失敗率：{fail_rate_7d:.1f}%")

        total_success_T = sum(r[1] for r in summary_today)
        total_fail_T = sum(r[2] for r in summary_today)
        total_all_T = total_success_T + total_fail_T
        fail_rate_T = (total_fail_T / total_all_T * 100.0) if total_all_T else 0.0
        self.lbl_today_success.setText(f"今日成功：{total_success_T}")
        self.lbl_today_fail.setText(f"今日失敗：{total_fail_T}")
        self.lbl_today_total.setText(f"今日上傳：{total_all_T}")
        self.lbl_today_fail_rate.setText(f"今日失敗率：{fail_rate_T:.1f}%")

    def refresh_events(self) -> None:
        limit = int(self.spin_limit.value())
        rows = self.store.recent_events(limit=limit) if hasattr(self.store, "recent_events") else []
        view_rows: List[List[Any]] = []
        for r in rows:
            status_text = "成功" if str(r.get("status")) == "success" else "失敗"
            view_rows.append([
                safe_int(r.get("id"), 0),
                r.get("project", ""),
                status_text,
                r.get("created_at", ""),
                r.get("message", ""),
            ])
        self.events_model.set_rows(view_rows)

    def _selected_project_id(self) -> Optional[int]:
        sel = self.tbl_projects.selectionModel().selectedRows()
        if not sel:
            return None
        row = sel[0].row()
        try:
            return int(self.projects_model.rows[row][0])
        except Exception:
            return None

    def on_project_selected(self) -> None:
        pid = self._selected_project_id()
        if pid is None:
            return
        p = self.store.get_project_by_id(pid)
        if not p:
            return
        self.lbl_project_id.setText(str(p.get("id")))
        self.edit_name.setText(str(p.get("name", "")))
        self.spin_interval.setValue(safe_int(p.get("interval_minutes"), 60))
        self.chk_enabled.setChecked(bool(safe_int(p.get("enabled"), 0)))
        last_run = str(p.get("last_run_at", "") or "")
        self.lbl_last_run.setText(last_run if last_run.strip() else "-")
        if last_run.strip():
            try:
                nxt = parse_dt(last_run) + timedelta(minutes=safe_int(p.get("interval_minutes"), 60))
                self.lbl_next_run.setText(nxt.replace(microsecond=0).isoformat(sep=" "))
            except Exception:
                self.lbl_next_run.setText("-")
        else:
            self.lbl_next_run.setText("(啟用後將立即符合條件)")
        self.edit_download_bat.setText(os.path.normpath(str(p.get("download_bat_path", "") or "")))
        self.edit_upload_bat.setText(os.path.normpath(str(p.get("upload_bat_path", "") or "")))

    def add_project(self) -> None:
        name, ok = QInputDialog.getText(self, "新增專案", "請輸入專案名稱：")
        if not ok:
            return
        name = name.strip()
        if not name:
            QMessageBox.warning(self, "提示", "專案名稱不可空白。")
            return
        try:
            self.store.add_project(name=name, interval_minutes=60, enabled=0)
            self.statusBar().showMessage(f"新增專案：{name}", 5000)
            self.refresh_all()
        except ValueError as ve:
            if str(ve) == "DUPLICATE_NAME":
                QMessageBox.warning(self, "提示", "專案名稱已存在（需唯一）。")
            else:
                QMessageBox.critical(self, "Error", f"新增失敗：{ve}")
            LOGGER.warning("Add project failed: %s", ve)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"新增失敗：{e}")
            LOGGER.exception("Add project unexpected error")

    def delete_project(self) -> None:
        pid = self._selected_project_id()
        if pid is None:
            QMessageBox.information(self, "提示", "請先選取要刪除的專案。")
            return
        p = self.store.get_project_by_id(pid)
        if not p:
            return
        msg = (
            f"確定要刪除專案？\n\n"
            f"ID={p.get('id')}\nName={p.get('name')}\n\n"
            f"注意：該專案的所有上傳事件也會一併刪除。"
        )
        if QMessageBox.question(self, "確認刪除", msg) == QMessageBox.StandardButton.Yes:
            try:
                self.store.delete_project(pid)
                self.statusBar().showMessage(f"刪除專案：{p.get('name')}", 5000)
                self.lbl_project_id.setText("-")
                self.edit_name.clear()
                self.chk_enabled.setChecked(False)
                self.lbl_last_run.setText("-")
                self.lbl_next_run.setText("-")
                self.edit_download_bat.clear()
                self.edit_upload_bat.clear()
                self.refresh_all()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"刪除失敗：{e}")
                LOGGER.exception("Delete project error")

    def save_project_settings(self) -> None:
        pid_text = self.lbl_project_id.text().strip()
        if pid_text == "-" or not pid_text:
            QMessageBox.information(self, "提示", "請先選取一個專案。")
            return
        pid = int(pid_text)
        name = self.edit_name.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "專案名稱不可空白。")
            return
        interval = int(self.spin_interval.value())
        enabled = 1 if self.chk_enabled.isChecked() else 0
        raw_dl = self.edit_download_bat.text().strip()
        raw_ul = self.edit_upload_bat.text().strip()
        dl_path = os.path.normpath(raw_dl) if raw_dl else ""
        ul_path = os.path.normpath(raw_ul) if raw_ul else ""
        try:
            self.store.update_project(
                pid, name, interval, enabled,
                download_bat_path=dl_path, upload_bat_path=ul_path
            )
            self.statusBar().showMessage(
                f"已儲存：{name}（every {interval} min, {'ON' if enabled else 'OFF'}）", 5000
            )
            self.refresh_projects()
            self.on_project_selected()
        except ValueError as ve:
            if str(ve) == "DUPLICATE_NAME":
                QMessageBox.warning(self, "提示", "專案名稱已存在（需唯一）。")
            else:
                QMessageBox.critical(self, "Error", f"儲存失敗：{ve}")
            LOGGER.warning("Save project failed: %s", ve)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"儲存失敗：{e}")
            LOGGER.exception("Save project unexpected error")

    def reset_last_run(self) -> None:
        pid_text = self.lbl_project_id.text().strip()
        if pid_text == "-" or not pid_text:
            QMessageBox.information(self, "提示", "請先選取一個專案。")
            return
        pid = int(pid_text)
        self.store.update_last_run_at(pid, None)
        self.statusBar().showMessage("已重置 last_run_at（下次檢查會立刻符合條件）", 5000)
        self.refresh_projects()
        self.on_project_selected()

    def _safe_check_intervals(self) -> None:
        try:
            self.check_intervals()
        except Exception:
            LOGGER.exception("check_intervals failed")

    def check_intervals(self) -> None:
        projects = self.store.list_projects()
        now_dt = datetime.now().replace(microsecond=0)
        any_ran = False
        for p in projects:
            if not bool(safe_int(p.get("enabled"), 0)):
                continue
            pid = int(p.get("id"))
            name = str(p.get("name", ""))
            interval = safe_int(p.get("interval_minutes"), 60)
            last_run_at = str(p.get("last_run_at", "") or "")
            should_run = False
            if not last_run_at.strip():
                should_run = True
            else:
                try:
                    last_dt = parse_dt(last_run_at)
                    if now_dt - last_dt >= timedelta(minutes=interval):
                        should_run = True
                except Exception:
                    should_run = True
            if not should_run:
                continue
            overall_status = self.perform_two_stage(pid, name, trigger="scheduler")
            self.store.update_last_run_at(pid, now_dt.isoformat(sep=" "))
            self.statusBar().showMessage(
                f"[Scheduler] {name} every {interval} min => {overall_status}", 7000
            )
            any_ran = True
        if any_ran:
            self.refresh_summary()
            self.refresh_events()
            self.refresh_projects()

    def _run_bat_with_logging(self, project_name: str, stage: str, bat_path: str) -> Tuple[str, str]:
        bat_path = os.path.normpath(bat_path)
        if not os.path.isfile(bat_path):
            return "fail", f"{stage}: 指定的 .bat 不存在：{bat_path}"
        cmd_list: List[str] = ["cmd.exe", "/C", bat_path]
        cwd = os.path.dirname(bat_path) or None
        try:
            completed = subprocess.run(
                cmd_list,
                capture_output=True,
                text=False,
                cwd=cwd,
                timeout=DEFAULT_TIMEOUT_SEC,
            )
            stdout_txt, stdout_enc = _decode_best_effort(completed.stdout or b"")
            stderr_txt, stderr_enc = _decode_best_effort(completed.stderr or b"")
            def _last_line(s: str) -> str:
                for line in reversed(s.splitlines()):
                    t = line.strip()
                    if t:
                        return t
                return ""
            tail = _last_line(stdout_txt) or _last_line(stderr_txt) or f"return_code={completed.returncode}"
            proj_dir = _project_logs_dir_on_D(project_name)
            os.makedirs(proj_dir, exist_ok=True)
            date_str = datetime.now().strftime("%Y%m%d")
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            stdout_path = os.path.join(proj_dir, f"{project_name}_{date_str}_{stage}_stdout.log")
            stderr_path = os.path.join(proj_dir, f"{project_name}_{date_str}_{stage}_stderr.log")
            stdout_path = os.path.normpath(stdout_path)
            stderr_path = os.path.normpath(stderr_path)
            with open(stdout_path, "a", encoding="utf-8-sig", errors="replace") as f_out:
                f_out.write(f"[{ts}] rc={completed.returncode} bat={bat_path}\n")
                f_out.write(f"cwd={cwd or ''}\n")
                f_out.write(f"decoded_with={stdout_enc}\n")
                f_out.write(stdout_txt)
                if not stdout_txt.endswith("\n"):
                    f_out.write("\n")
                f_out.write("-" * 60 + "\n")
            with open(stderr_path, "a", encoding="utf-8-sig", errors="replace") as f_err:
                f_err.write(f"[{ts}] rc={completed.returncode} bat={bat_path}\n")
                f_err.write(f"cwd={cwd or ''}\n")
                f_err.write(f"decoded_with={stderr_enc}\n")
                f_err.write(stderr_txt)
                if not stderr_txt.endswith("\n"):
                    f_err.write("\n")
                f_err.write("-" * 60 + "\n")
            status = "success" if completed.returncode in SUCCESS_RETURN_CODES else "fail"
            return status, f"{stage}: {os.path.basename(bat_path)}: {tail}"
        except subprocess.TimeoutExpired:
            return "fail", f"{stage}: {os.path.basename(bat_path)} 逾時（>{DEFAULT_TIMEOUT_SEC}s）"
        except Exception as e:
            return "fail", f"{stage}: {os.path.basename(bat_path)} 執行例外：{e!s}"

    # NEW: nan 哨兵判斷（大小寫不分；容忍前後空白；不含中括號）
    @staticmethod
    def _is_skip_bat(text: str) -> bool:  # NEW
        return (text or "").strip().lower() == "nan"  # NEW

    # ---- 在 MainWindow 類別內，整段替換 perform_two_stage(...) ----  # UPDATED
    def perform_two_stage(self, project_id: int, project_name: str, trigger: str) -> str:
        p = self.store.get_project_by_id(project_id)
        dl_raw = str((p or {}).get("download_bat_path", "") or "")
        ul_raw = str((p or {}).get("upload_bat_path", "") or "")
        dl = os.path.normpath(dl_raw) if dl_raw else ""
        ul = os.path.normpath(ul_raw) if ul_raw else ""

        def _warn_unknown_once(msg: str) -> None:
            key = (msg or "").strip().lower()[:160]
            if key and key not in self._unknown_warned_keys:
                self._unknown_warned_keys.add(key)
                r = QMessageBox.question(
                    self, "未定義錯誤碼",
                    f"偵測到未定義錯誤碼，已暫以 E000 記錄。\n\n訊息：\n{msg}\n\n要立刻開啟 error_keywords.csv 新增規則嗎？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes,
                )
                if r == QMessageBox.StandardButton.Yes:
                    QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(ERROR_KEYWORDS_CSV)))
                self.statusBar().showMessage("未定義錯誤碼：已以 E000 記錄（可在規則檔新增關鍵字）", 7000)

        # nan 跳過規則
        skip_dl = self._is_skip_bat(dl_raw)
        skip_ul = self._is_skip_bat(ul_raw)
        overall = "success"

        # --- Download ---
        if skip_dl:
            msg = f"{trigger}: download: skipped (nan)"
            self.store.add_event(project_id, "success", now_iso(), msg, error_code="")  # success 不給 code
        else:
            if not dl:
                msg = f"{trigger}: download: 未設定 download_bat_path"
                ec = ERROR_RESOLVER.resolve("fail", msg)
                if ec == "E000": _warn_unknown_once(msg)
                self.store.add_event(project_id, "fail", now_iso(), msg, error_code=ec)
                return "fail"
            if not os.path.isfile(dl):
                msg = f"{trigger}: download: 指定的 .bat 不存在：{dl}"
                ec = ERROR_RESOLVER.resolve("fail", msg)
                if ec == "E000": _warn_unknown_once(msg)
                self.store.add_event(project_id, "fail", now_iso(), msg, error_code=ec)
                return "fail"
            st_dl, msg_dl = self._run_bat_with_logging(project_name, "download", dl)
            msg_full = f"{trigger}: {msg_dl}"
            if st_dl != "success":
                ec = ERROR_RESOLVER.resolve("fail", msg_full)
                if ec == "E000": _warn_unknown_once(msg_full)
                self.store.add_event(project_id, st_dl, now_iso(), msg_full, error_code=ec)
                return "fail"
            else:
                self.store.add_event(project_id, st_dl, now_iso(), msg_full, error_code="")

        # --- Upload ---
        if skip_ul:
            msg = f"{trigger}: upload: skipped (nan)"
            self.store.add_event(project_id, "success", now_iso(), msg, error_code="")
            return overall
        else:
            if not ul:
                msg = f"{trigger}: upload: 未設定 upload_bat_path"
                ec = ERROR_RESOLVER.resolve("fail", msg)
                if ec == "E000": _warn_unknown_once(msg)
                self.store.add_event(project_id, "fail", now_iso(), msg, error_code=ec)
                return "fail"
            if not os.path.isfile(ul):
                msg = f"{trigger}: upload: 指定的 .bat 不存在：{ul}"
                ec = ERROR_RESOLVER.resolve("fail", msg)
                if ec == "E000": _warn_unknown_once(msg)
                self.store.add_event(project_id, "fail", now_iso(), msg, error_code=ec)
                return "fail"
            st_ul, msg_ul = self._run_bat_with_logging(project_name, "upload", ul)
            msg_full = f"{trigger}: {msg_ul}"
            if st_ul != "success":
                ec = ERROR_RESOLVER.resolve("fail", msg_full)
                if ec == "E000": _warn_unknown_once(msg_full)
                self.store.add_event(project_id, st_ul, now_iso(), msg_full, error_code=ec)
                return "fail"
            else:
                self.store.add_event(project_id, st_ul, now_iso(), msg_full, error_code="")

        return overall

    def run_selected_now(self) -> None:
        pid_text = self.lbl_project_id.text().strip()
        if pid_text == "-" or not pid_text:
            QMessageBox.information(self, "提示", "請先選取一個專案。")
            return
        pid = int(pid_text)
        p = self.store.get_project_by_id(pid)
        if not p:
            return
        overall_status = self.perform_two_stage(pid, str(p.get("name", "")), trigger="manual")
        self.store.update_last_run_at(pid, now_iso())
        self.statusBar().showMessage(f"[Manual] {p.get('name')} => {overall_status}", 5000)
        self.refresh_all()

    def delete_selected_event(self) -> None:
        sel = self.tbl_events.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "提示", "請先選取一筆事件。")
            return
        row = sel[0].row()
        try:
            event_id = int(self.events_model.rows[row][0])
        except Exception:
            return
        if QMessageBox.question(self, "確認刪除", f"確定刪除事件 ID={event_id}？") == QMessageBox.StandardButton.Yes:
            try:
                self.store.delete_event(event_id)
                self.refresh_events()
                self.refresh_summary()
                self.statusBar().showMessage(f"Deleted event ID={event_id}", 5000)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"刪除失敗：{e}")
                LOGGER.exception("Delete event error")

    def export_summary_csv(self) -> None:
        summary = self.store.summary_last_7_days()
        path, _ = QFileDialog.getSaveFileName(self, "Save CSV", "upload_summary_last7days.csv", "CSV Files (*.csv)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig") as f:
                f.write("project,success_7d,fail_7d,total_7d\n")
                for project, succ, fail, total in summary:
                    f.write(f"{project},{succ},{fail},{total}\n")
            QMessageBox.information(self, "完成", f"已匯出：\n{path}")
            LOGGER.info("CSV exported: %s", path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"匯出失敗：{e}")
            LOGGER.exception("Export CSV error")

# -----------（補齊 recent_events：若你的現檔已有可保留，此段可忽略）-----------
def _store_recent_events_patch(self: ExcelStore, limit: int = 50) -> List[Dict[str, Any]]:
    with self.lock:
        projects, events = self._read_all()
        if events.empty:
            return []
        ev = events.copy()
        ev["created_at_dt"] = pd.to_datetime(ev["created_at"], errors="coerce")
        ev = ev.sort_values(by=["created_at_dt", "id"], ascending=[False, False], kind="stable").head(limit)
        return ev[["id", "project", "status", "created_at", "message"]].to_dict("records")

if not hasattr(ExcelStore, "recent_events"):
    ExcelStore.recent_events = _store_recent_events_patch  # type: ignore[attr-defined]

# ====================== Error Code Resolver ======================  # NEW
ERROR_KEYWORDS_CSV = "error_keywords.csv"  # NEW

class ErrorCodeResolver:  # NEW
    """以最小複雜度處理錯誤碼對應：
       - 使用者自訂：error_keywords.csv（兩欄：error_code,keyword；大小寫不敏感；由上而下第一個命中）
       - 內建保底規則：只在 CSV 無命中時才判定
    """
    def __init__(self, path: str = ERROR_KEYWORDS_CSV) -> None:
        self.path = path
        self.rules: list[tuple[str, str]] = []  # (code_upper, keyword_lower)
        self._ensure_file_exists()
        self.load()

    def _ensure_file_exists(self) -> None:
        if not os.path.exists(self.path):
            try:
                with open(self.path, "w", encoding="utf-8-sig", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(["error_code", "keyword"])
                    # 範例（你可刪除或自行追加）
                    w.writerow(["E001", "逾時（>"])
                    w.writerow(["E003", "指定的 .bat 不存在"])
                    w.writerow(["E005", "return_code="])
            except Exception:
                LOGGER.debug("create error_keywords.csv failed", exc_info=True)

    def load(self) -> int:
        rules: list[tuple[str, str]] = []
        try:
            with open(self.path, "r", encoding="utf-8-sig", newline="") as f:
                r = csv.DictReader(f)
                for row in r:
                    code = (row.get("error_code") or "").strip().upper()
                    kw = (row.get("keyword") or "").strip().lower()
                    if not code or not kw:
                        continue
                    rules.append((code, kw))
        except FileNotFoundError:
            self._ensure_file_exists()
        except Exception:
            LOGGER.debug("load error_keywords.csv failed", exc_info=True)

        self.rules = rules
        return len(self.rules)

    # 內建保底規則（僅在 CSV 無命中時才用）
    @staticmethod
    def _builtin_rule(message: str) -> str:
        msg = (message or "").lower()

        # 共用
        if "逾時（>" in message:
            return "E001"
        if "未設定" in message:
            return "E002"
        if "指定的 .bat 不存在" in message:
            return "E003"
        if "執行例外" in message:
            return "E004"
        m = re.search(r"return_code=(\d+)", message, flags=re.IGNORECASE)
        if m:
            try:
                if int(m.group(1)) != 0:
                    return "E005"
            except Exception:
                pass
        if "is not recognized as an internal or external command" in msg:
            return "E006"
        if ("the system cannot find the file" in msg) or ("could not find" in msg):
            return "E007"
        if "access is denied" in msg:
            return "E008"

        # download / upload 泛化（訊息通常包含前綴）
        if "download:" in msg:
            return "E101"
        if "upload:" in msg:
            return "E201"

        return "E000"  # 仍未命中

    def resolve(self, status: str, message: str) -> str:
        if (status or "").strip().lower() != "fail":
            return ""  # success 不給 code
        msg_lc = (message or "").lower()

        # 先比使用者自訂（兩欄關鍵字：包含即命中；第一筆優先）
        for code, kw in self.rules:
            if kw in msg_lc:
                return code

        # 再用內建保底規則
        return self._builtin_rule(message)

# 全域單例（MainWindow 與 ExcelStore 共用）
ERROR_RESOLVER = ErrorCodeResolver()  # NEW
# ==================== /Error Code Resolver =======================

# ---------------- main entry ----------------
def main() -> int:
    try:
        from PyQt6.QtCore import Qt
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
    except Exception:
        pass
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    return app.exec()

if __name__ == "__main__":
    sys.exit(main())