import json
import sys
import time
import os
import re
import ctypes
import pickle
from datetime import datetime
from pathlib import Path

# --- 外部ライブラリ ---
import google.generativeai as genai
from PyQt6.QtWidgets import (
    QApplication, QComboBox, QLabel, QListWidget, QMainWindow, QMenu,
    QMessageBox, QPushButton, QStyle, QSystemTrayIcon, QVBoxLayout,
    QHBoxLayout, QWidget, QLineEdit, QListWidgetItem, QProgressDialog, QInputDialog
)
from PyQt6.QtGui import QAction, QColor, QIcon
from PyQt6.QtCore import Qt, QSize
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from PIL import Image

# --- Google Drive API 関連 ---
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# --- パス・環境設定 ---
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
else:
    BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "data"
HISTORY_PATH = DATA_DIR / "history.jsonl"
CONFIG_PATH = DATA_DIR / "config.json"
ICON_PATH = BASE_DIR / "icon.ico"
CRED_PATH = BASE_DIR / "credentials.json"  # Google Cloudから取得したファイル
TOKEN_PATH = DATA_DIR / "token.pickle"      # ログイン情報を保存するファイル
WINDOWS_FORBIDDEN_CHARS = re.compile(r'[\\/:*?"<>|]')

# ドライブの権限スコープ（アプリが作成したファイルのみを操作）
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# --- Google Drive 認証・アップロードエンジン ---

def get_drive_service(parent):
    """Google Drive APIの認証を行い、サービスオブジェクトを返す"""
    creds = None
    if TOKEN_PATH.exists():
        with open(TOKEN_PATH, 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CRED_PATH.exists():
                QMessageBox.critical(parent, "認証エラー", "credentials.json が見つかりません。\nGoogle Cloud Consoleから取得してアプリと同じ場所に配置してください。")
                return None
            flow = InstalledAppFlow.from_client_secrets_file(str(CRED_PATH), SCOPES)
            creds = flow.run_local_server(port=0)
        
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(TOKEN_PATH, 'wb') as token:
            pickle.dump(creds, token)
            
    return build('drive', 'v3', credentials=creds)

def upload_to_drive(file_path, mime_type, parent):
    """ファイルをGoogleドライブへアップロードする"""
    service = get_drive_service(parent)
    if not service: return None
    
    file_metadata = {'name': file_path.name}
    media = MediaFileUpload(str(file_path), mimetype=mime_type)
    
    try:
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        return file.get('id')
    except Exception as e:
        print(f"Drive Upload Error: {e}")
        return None

# --- 現代風 QSS ---
MODERN_STYLE = """
QMainWindow { background-color: #1e1e1e; }
QWidget { color: #e0e0e0; font-family: "Segoe UI", "Meiryo"; font-size: 14px; }
QLabel { font-weight: bold; color: #aaaaaa; margin-top: 8px; }
QLineEdit, QComboBox, QListWidget {
    background-color: #2d2d2d; border: 1px solid #3f3f3f; border-radius: 6px; padding: 8px;
}
QPushButton {
    background-color: #3e3e3e; border: none; border-radius: 8px; padding: 8px 12px; font-weight: bold;
}
QPushButton#primary { background-color: #007acc; color: white; }
QPushButton#primary:hover { background-color: #0098ff; }
QPushButton#danger { color: #ff6b6b; }
QPushButton#danger:hover { background-color: #3d2b2b; }
QListWidget::item { background-color: #353535; border-radius: 8px; margin: 4px 8px; border: 1px solid #3f3f3f; }
"""

# --- ユーティリティ関数 ---

def append_history_entry(entry):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with HISTORY_PATH.open("a", encoding="utf-8", newline="\n") as h:
        h.write(json.dumps(entry, ensure_ascii=False) + "\n")

def load_run_records(limit=10):
    runs_dir = DATA_DIR / "runs"
    if not runs_dir.exists(): return []
    records = []
    for d in runs_dir.iterdir():
        p = d / "run.json"
        if p.exists():
            try:
                with p.open("r", encoding="utf-8") as h: records.append(json.load(h))
            except: continue
    records.sort(key=lambda x: x.get("created_at", ""), reverse=True)
    return records[:limit]

def save_run_record(record):
    run_id = record.get("run_id")
    if not run_id: return False
    run_dir = DATA_DIR / "runs" / run_id
    run_dir.mkdir(parents=True, exist_ok=True)
    try:
        with (run_dir / "run.json").open("w", encoding="utf-8") as h:
            json.dump(record, h, ensure_ascii=False, indent=2)
        return True
    except: return False

def sanitize_output_basename(raw):
    return WINDOWS_FORBIDDEN_CHARS.sub("_", (raw or "").strip()) or "result"

def get_api_key(parent=None):
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            c = json.load(f)
            if "api_key" in c: return c["api_key"]
    k, ok = QInputDialog.getText(parent, "Gemini APIキー", "キーを入力:", QLineEdit.EchoMode.Password)
    if ok and k:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f: json.dump({"api_key": k}, f)
        return k
    return None

def execute_gemini_process(record, owner):
    api_key = get_api_key(owner)
    if not api_key: return
    
    p = QProgressDialog("AI解析中 ＆ ドライブ同期中...", "キャンセル", 0, 0, owner)
    p.setWindowModality(Qt.WindowModality.WindowModal)
    p.show(); QApplication.processEvents()
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('models/gemini-2.5-flash')
        run_id, doc_format, purpose = record["run_id"], record.get("doc_format"), record["purpose"]
        
        prompt = f"目的: {purpose}\n"
        if doc_format == "Word": prompt += "指示: 構造化レポート形式で出力。"
        elif doc_format == "Excel": prompt += "指示: タブ区切りの表形式で出力。"
        elif doc_format == "PowerPoint": prompt += "指示: スライド構成案を出力。"
        else: prompt += "指示: 装飾記号を使わないプレーンテキストで出力。"
        
        prompt_parts = [prompt]
        for img in record.get("captures", []):
            img_p = DATA_DIR / "captures" / run_id / img
            if img_p.exists(): prompt_parts.append(Image.open(img_p))
        
        res = model.generate_content(prompt_parts)
        ai_text = res.text
        
        base = record["output_basename"]
        out_d = DATA_DIR / "outputs" / run_id; out_d.mkdir(parents=True, exist_ok=True)
        
        m_type = "text/plain" # デフォルトMIME

        if doc_format == "Word":
            target = out_d / f"{base}.docx"
            m_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            doc = Document(); doc.add_heading(purpose, 0); doc.add_paragraph(ai_text); doc.save(str(target))
        elif doc_format == "Excel":
            target = out_d / f"{base}.xlsx"
            m_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            wb = Workbook(); wb.active["A1"] = ai_text; wb.save(str(target))
        elif doc_format == "PowerPoint":
            target = out_d / f"{base}.pptx"
            m_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            prs = Presentation(); s = prs.slides.add_slide(prs.slide_layouts[1]); s.shapes.title.text = purpose; s.placeholders[1].text = ai_text; prs.save(str(target))
        else:
            target = out_d / f"{base}.txt"
            with target.open("w", encoding="utf-8") as h: h.write(ai_text.replace("**", "").replace("#", "").strip())
        
        # 履歴とステータスの更新
        record["status"] = "done"; record["output_file"] = target.name
        save_run_record(record); os.startfile(str(target))

        # --- Googleドライブへアップロード ---
        drive_id = upload_to_drive(target, m_type, owner)
        
        if drive_id:
            QMessageBox.information(owner, "成功", f"解析完了！Googleドライブにも保存しました。\n(File ID: {drive_id})")

    except Exception as e:
        QMessageBox.critical(owner, "エラー", str(e))
    finally:
        p.close()

# --- ウィンドウクラス ---

class NewRunWindow(QMainWindow):
    def __init__(self, run_id):
        super().__init__(); self.run_id = run_id; self.capture_index = 1
        self.created_at = datetime.now().isoformat(timespec="seconds")
        self.setWindowTitle("新規実行 - SnipAI")
        if ICON_PATH.exists(): self.setWindowIcon(QIcon(str(ICON_PATH)))
        self._build_ui()
        
    def _build_ui(self):
        c = QWidget(); l = QVBoxLayout(c); l.setContentsMargins(20, 20, 20, 20); l.setSpacing(12)
        self.output_name_input = QLineEdit(); self.output_name_input.setPlaceholderText("ファイル名を入力")
        self.purpose_combo = QComboBox(); self.purpose_combo.addItems(["まとめる", "解説する", "ドキュメント化"])
        self.doc_format_combo = QComboBox(); self.doc_format_combo.addItems(["Word", "Excel", "PowerPoint"])
        self.doc_format_combo.setEnabled(False)
        self.purpose_combo.currentTextChanged.connect(lambda t: self.doc_format_combo.setEnabled(t == "ドキュメント化"))
        self.capture_list = QListWidget()
        bl = QHBoxLayout(); cb = QPushButton("📸 スクショ追加"); db = QPushButton("🗑 選択削除")
        db.setObjectName("danger"); bl.addWidget(cb); bl.addWidget(db)
        rb = QPushButton("🚀 AI解析を実行"); rb.setObjectName("primary"); sb = QPushButton("💾 保存")
        l.addWidget(QLabel("出力ファイル名")); l.addWidget(self.output_name_input)
        row = QHBoxLayout(); v1, v2 = QVBoxLayout(), QVBoxLayout()
        v1.addWidget(QLabel("目的")); v1.addWidget(self.purpose_combo)
        v2.addWidget(QLabel("形式")); v2.addWidget(self.doc_format_combo)
        row.addLayout(v1); row.addLayout(v2); l.addLayout(row)
        l.addWidget(QLabel("スクリーンショット")); l.addWidget(self.capture_list)
        l.addLayout(bl); l.addSpacing(15); l.addWidget(rb); l.addWidget(sb)
        cb.clicked.connect(self.capture_screenshot); db.clicked.connect(self.remove_selected)
        rb.clicked.connect(self.run_now); sb.clicked.connect(self.save_only)
        self.setCentralWidget(c); self.resize(450, 600)
        
    def _collect(self):
        base = sanitize_output_basename(self.output_name_input.text())
        fmt = self.doc_format_combo.currentText() if self.doc_format_combo.isEnabled() else "Text"
        ext = {"Word": ".docx", "Excel": ".xlsx", "PowerPoint": ".pptx", "Text": ".txt"}[fmt]
        return {
            "run_id": self.run_id, "created_at": self.created_at, "purpose": self.purpose_combo.currentText(),
            "doc_format": fmt if fmt != "Text" else None, "save_target": "local", 
            "captures": [self.capture_list.item(i).text() for i in range(self.capture_list.count())],
            "status": "ready", "output_basename": base, "output_file": f"{base}{ext}"
        }
        
    def capture_screenshot(self):
        self.hide(); time.sleep(0.3)
        try:
            import mss, mss.tools, win32gui
            d = DATA_DIR / "captures" / self.run_id; d.mkdir(parents=True, exist_ok=True)
            f = f"{self.capture_index:03d}.png"; fp = d / f
            h = win32gui.GetForegroundWindow(); l, t, r, b = win32gui.GetWindowRect(h)
            with mss.mss() as sct:
                img = sct.grab({"left": l, "top": t, "width": r-l, "height": b-t})
                mss.tools.to_png(img.rgb, img.size, output=str(fp))
            self.capture_list.addItem(f); self.capture_index += 1
        finally: self.show()
        
    def remove_selected(self):
        for i in self.capture_list.selectedItems(): self.capture_list.takeItem(self.capture_list.row(i))
        
    def run_now(self):
        if self.capture_list.count() == 0: return QMessageBox.warning(self, "SnipAI", "スクショを撮ってください")
        d = self._collect(); save_run_record(d); append_history_entry({**d, "id": self.run_id})
        execute_gemini_process(d, self); self.close()
        
    def save_only(self):
        d = self._collect(); save_run_record(d); append_history_entry({**d, "id": self.run_id}); self.close()

class HistoryWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("実行履歴")
        self.list_widget = QListWidget(); self.list_widget.setSpacing(6)
        if ICON_PATH.exists(): self.setWindowIcon(QIcon(str(ICON_PATH)))
        c = QWidget(); l = QVBoxLayout(c); l.addWidget(self.list_widget); self.setCentralWidget(c); self.resize(500, 450)
        
    def refresh(self):
        self.list_widget.clear()
        for r in load_run_records():
            item = QListWidgetItem(); w = QWidget(); lay = QHBoxLayout(w)
            lay.setContentsMargins(10, 2, 10, 2)
            lbl = QLabel(f"{r['created_at'][5:16].replace('T', ' ')} | {r['purpose']}")
            btn = QPushButton("📄 開く" if r.get('status') == "done" else "▶ 実行")
            btn.setFixedWidth(110); btn.setFixedHeight(34)
            btn.clicked.connect(lambda _, rec=r: self._handle(rec))
            lay.addWidget(lbl, 1); lay.addWidget(btn); item.setSizeHint(QSize(100, 58))
            self.list_widget.addItem(item); self.list_widget.setItemWidget(item, w)
            
    def _handle(self, r):
        if r['status'] == "done":
            p = DATA_DIR / "outputs" / r["run_id"] / r["output_file"]
            if p.exists(): os.startfile(str(p))
            else: QMessageBox.warning(self, "エラー", "ファイルなし")
        else: execute_gemini_process(r, self); self.refresh()

class HomeWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle("SnipAI"); self.setStyleSheet(MODERN_STYLE)
        if ICON_PATH.exists(): self.setWindowIcon(QIcon(str(ICON_PATH)))
        c = QWidget(); l = QVBoxLayout(c); l.setContentsMargins(30, 30, 30, 30); l.setSpacing(15)
        b1 = QPushButton("🆕 新規実行"); b1.setObjectName("primary"); b1.setFixedHeight(50)
        b2 = QPushButton("📜 履歴を表示"); b2.setFixedHeight(50)
        b1.clicked.connect(self.open_new); b2.clicked.connect(self.open_hist)
        l.addWidget(b1); l.addWidget(b2); self.setCentralWidget(c); self.resize(350, 250)
        
    def open_new(self):
        rid = datetime.now().strftime("%Y%m%d_%H%M%S"); self.nw = NewRunWindow(rid); self.nw.setStyleSheet(MODERN_STYLE); self.nw.show()
        
    def open_hist(self):
        if not hasattr(self, 'hw'): self.hw = HistoryWindow(); self.hw.setStyleSheet(MODERN_STYLE)
        self.hw.refresh(); self.hw.show()
        
    def closeEvent(self, e): self.hide(); e.ignore()

def main():
    try:
        myappid = 'mickyyya.snipai.v1' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except: pass

    app = QApplication(sys.argv); app.setQuitOnLastWindowClosed(False)
    icon = QIcon(str(ICON_PATH)) if ICON_PATH.exists() else app.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
    home = HomeWindow(); home.setWindowIcon(icon)
    tray = QSystemTrayIcon(icon, app); menu = QMenu()
    menu.addAction("ホーム").triggered.connect(home.show); menu.addAction("終了").triggered.connect(app.quit)
    tray.setContextMenu(menu); tray.show(); home.show()
    return app.exec()

if __name__ == "__main__": main()