import sys
import os
import re
import zipfile
import shutil
import subprocess
import tempfile
import webbrowser
import requests
import time
import random
import string
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QPushButton, QLabel, QTextEdit, QFileDialog, 
                             QMessageBox, QProgressBar, QHBoxLayout, QFrame, QDialog)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QFont, QColor, QPalette, QIcon

# --- UTILITY FOR RESOURCE PATHS ---
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- WORKER FOR DOWNLOADING CALIBRE PORTABLE ---
class DownloadWorker(QThread):
    progress_val = pyqtSignal(int)
    progress_msg = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def run(self):
        random_suffix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))
        target_dir = os.path.join(os.environ.get('LOCALAPPDATA', os.getcwd()), f"Raafat_Engine_{random_suffix}")
        save_path = os.path.join(tempfile.gettempdir(), f"setup_{random_suffix}.exe")
        
        try:
            url = "https://download.calibre-ebook.com/portable/calibre-portable-7.3.0.exe"
            self.progress_msg.emit("🌍 Connecting to Masterpiece Servers...")
            
            response = requests.get(url, stream=True, timeout=60)
            if response.status_code != 200:
                self.finished.emit(False, f"Download failed: HTTP {response.status_code}")
                return

            total_size = int(response.headers.get('content-length', 125 * 1024 * 1024))
            downloaded = 0
            
            with open(save_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=512 * 1024):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        percent = int((downloaded / total_size) * 100)
                        self.progress_val.emit(min(percent, 99))
                        if downloaded % (5 * 1024 * 1024) == 0:
                            self.progress_msg.emit(f"📥 Downloading: {percent}% ({downloaded//1024//1024}MB / {total_size//1024//1024}MB)")

            self.progress_msg.emit("⌛ Finalizing download...")
            time.sleep(4)
            
            self.progress_msg.emit("📦 Extracting Engine...")
            os.makedirs(target_dir, exist_ok=True)
            
            cmd = f'"{save_path}" "{target_dir}"'
            process = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            stdout, stderr = process.communicate()
            
            if process.returncode != 0:
                raise Exception(f"Extractor exited with code {process.returncode}")

            engine_exe = None
            for root, dirs, files in os.walk(target_dir):
                if "ebook-convert.exe" in files:
                    engine_exe = os.path.join(root, "ebook-convert.exe")
                    break
            
            if engine_exe and os.path.exists(engine_exe):
                try: os.remove(save_path)
                except: pass
                self.finished.emit(True, engine_exe)
            else:
                self.finished.emit(False, "Setup finished but 'ebook-convert.exe' was not found.")

        except Exception as e:
            self.finished.emit(False, str(e))

class RaafatWorker(QThread):
    progress_msg = pyqtSignal(str)
    progress_val = pyqtSignal(int)
    finished = pyqtSignal(bool, str)

    def __init__(self, source_path, target_path, calibre_exe="ebook-convert"):
        super().__init__()
        self.source_path = source_path
        self.target_path = target_path
        self.calibre_exe = calibre_exe
        self._is_cancelled = False
        self.process = None

    def run(self):
        try:
            self.progress_val.emit(5)
            with tempfile.TemporaryDirectory() as temp_dir:
                extract_path = os.path.join(temp_dir, "extracted")
                with zipfile.ZipFile(self.source_path, 'r') as zf: zf.extractall(extract_path)
                epub_root = None
                for root, dirs, files in os.walk(extract_path):
                    if "mimetype" in files: epub_root = root; break
                
                if not epub_root:
                    self.finished.emit(False, "Valid ebook not found.")
                    return

                epub_path = os.path.join(temp_dir, "temp.epub")
                with zipfile.ZipFile(epub_path, "w") as zf:
                    mimetype_src = os.path.join(epub_root, "mimetype")
                    zf.write(mimetype_src, "mimetype", compress_type=zipfile.ZIP_STORED)
                    for r, d, files in os.walk(epub_root):
                        for f in files:
                            if self._is_cancelled: return
                            zf.write(os.path.join(r, f), os.path.relpath(os.path.join(r, f), epub_root).replace("\\", "/"), compress_type=zipfile.ZIP_DEFLATED)
                
                self.progress_val.emit(40)
                cmd = [self.calibre_exe, epub_path, self.target_path, "--paper-size", "a4", "--pdf-page-numbers"]
                self.process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1, creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                p_pattern = re.compile(r"(\d+)%")
                while True:
                    if self._is_cancelled: self.process.terminate(); return
                    line = self.process.stdout.readline()
                    if not line and self.process.poll() is not None: break
                    if line:
                        m = p_pattern.search(line)
                        if m: self.progress_val.emit(40 + int(int(m.group(1)) * 0.6))

                if self.process.returncode == 0:
                    self.finished.emit(True, self.target_path)
                else:
                    self.finished.emit(False, "Conversion engine failed.")
        except Exception as e:
            self.finished.emit(False, str(e))

    def cancel(self): self._is_cancelled = True

class ContactDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("صلي علي النبي")
        self.setFixedSize(380, 320)
        self.setStyleSheet(self.get_theme())
        self.setWindowIcon(QIcon(resource_path("icon (1).ico")))
        layout = QVBoxLayout(self)
        h = QLabel("صلي علي النبي"); h.setAlignment(Qt.AlignmentFlag.AlignCenter); h.setStyleSheet("font-size: 32px; color: #4CAF50; font-family: 'Amiri';")
        layout.addWidget(h)
        for t, u, c in [("🌐 Facebook", "https://fb.com/raafat.reda.366930/", "#1877F2"), ("🐙 GitHub", "https://github.com/RaafatReda1", "#FFFFFF"), ("💬 WhatsApp", "https://wa.me/201022779263", "#25D366")]:
            b = QPushButton(t); b.setStyleSheet(f"text-align: left; padding: 12px; color: {c};"); b.clicked.connect(lambda _, u=u: webbrowser.open(u)); layout.addWidget(b)
        cbtn = QPushButton("Done"); cbtn.clicked.connect(self.close); layout.addWidget(cbtn)
    def get_theme(self): return "QDialog { background-color: #121212; } QWidget { color: #E0E0E0; } QPushButton { background-color: #333; border-radius: 6px; padding: 10px; }"

class RaafatConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EXE to PDF by Raafat")
        self.setWindowIcon(QIcon(resource_path("icon (1).ico")))
        self.resize(650, 650)
        self.setStyleSheet(self.get_main_theme())
        self.selected_path = None
        self.calibre_path = "ebook-convert"
        self.init_ui()
        self.detect_calibre()

    def get_main_theme(self):
        return """
        QMainWindow { background-color: #121212; }
        QWidget { color: #E0E0E0; font-family: 'Segoe UI', Arial; }
        QFrame#MainContainer { background-color: #1E1E1E; border-radius: 15px; }
        QLabel#Title { color: #4CAF50; font-size: 24px; font-weight: bold; }
        QLabel#SubTitle { color: #888; font-size: 12px; }
        QLabel#SelectedFileLabel { color: #4CAF50; font-weight: bold; background-color: #252525; padding: 10px; border-radius: 8px; border: 1px solid #333; font-size: 13px; }
        QLabel#DropZone { border: 2px dashed #444; border-radius: 12px; background-color: #252525; color: #AAA; font-size: 16px; padding: 30px; }
        QLabel#DropZone:hover { border: 2px dashed #4CAF50; background-color: #2A2A2A; color: #4CAF50; }
        QPushButton { background-color: #333; color: white; border: none; padding: 12px 24px; border-radius: 6px; font-size: 14px; font-weight: 500; }
        QPushButton:hover { background-color: #444; }
        QPushButton#Primary { background-color: #2E7D32; }
        QPushButton#Primary:hover { background-color: #388E3C; }
        QPushButton#Cancel { background-color: #C62828; }
        QPushButton#Cancel:hover { background-color: #D32F2F; }
        QPushButton#ContactBtn { background-color: transparent; color: #888; font-size: 12px; text-decoration: underline; }
        QPushButton#ContactBtn:hover { color: #4CAF50; }
        QProgressBar { border: none; background-color: #333; height: 10px; border-radius: 5px; text-align: center; }
        QProgressBar::chunk { background-color: #4CAF50; border-radius: 5px; }
        QTextEdit { background-color: #1A1A1A; border: 1px solid #333; border-radius: 8px; color: #00FF41; font-family: 'Consolas'; font-size: 11px; padding: 10px; }
        """

    def init_ui(self):
        container = QFrame(); container.setObjectName("MainContainer"); self.setCentralWidget(container)
        layout = QVBoxLayout(container); layout.setContentsMargins(25, 25, 25, 25); layout.setSpacing(14)
        header = QHBoxLayout()
        title_box = QVBoxLayout()
        t = QLabel("EXE to PDF by Raafat"); t.setObjectName("Title")
        s = QLabel("Premium Calibre Mastery | صلي علي النبي"); s.setObjectName("SubTitle")
        title_box.addWidget(t); title_box.addWidget(s); header.addLayout(title_box)
        c_btn = QPushButton("Support & Contact"); c_btn.setObjectName("ContactBtn"); c_btn.clicked.connect(lambda: ContactDialog(self).exec())
        header.addWidget(c_btn, 0, Qt.AlignmentFlag.AlignTop|Qt.AlignmentFlag.AlignRight); layout.addLayout(header)
        self.drop_zone = QLabel("Drag and Drop your EXE here")
        self.drop_zone.setObjectName("DropZone"); self.drop_zone.setAlignment(Qt.AlignmentFlag.AlignCenter); self.drop_zone.setAcceptDrops(True); self.drop_zone.setMinimumHeight(120); layout.addWidget(self.drop_zone)
        self.btn_browse = QPushButton("📁 Browse Archive..."); self.btn_browse.clicked.connect(self.browse_file); layout.addWidget(self.btn_browse)
        self.file_display = QLabel("No file selected"); self.file_display.setObjectName("SelectedFileLabel"); self.file_display.setAlignment(Qt.AlignmentFlag.AlignCenter); layout.addWidget(self.file_display)
        ctrl = QHBoxLayout()
        self.btn_run = QPushButton("🚀 Run Conversion"); self.btn_run.setObjectName("Primary"); self.btn_run.setEnabled(False); self.btn_run.clicked.connect(self.run_conversion)
        self.btn_cancel = QPushButton("❌ Cancel"); self.btn_cancel.setObjectName("Cancel"); self.btn_cancel.setEnabled(False); self.btn_cancel.clicked.connect(self.cancel_conversion)
        ctrl.addWidget(self.btn_run); ctrl.addWidget(self.btn_cancel); layout.addLayout(ctrl)
        self.progress_bar = QProgressBar(); layout.addWidget(self.progress_bar)
        self.status = QLabel("Ready"); layout.addWidget(self.status)
        self.logs = QTextEdit(); self.logs.setReadOnly(True); layout.addWidget(self.logs)

    def log(self, m): self.logs.append(m); self.logs.verticalScrollBar().setValue(self.logs.verticalScrollBar().maximum())

    def detect_calibre(self):
        try:
            subprocess.run(["ebook-convert", "--version"], capture_output=True, check=True)
            self.calibre_path = "ebook-convert"
            self.log("✅ Calibre Detect: System Ready.")
            return
        except: pass
        base_dir = os.environ.get('LOCALAPPDATA', os.getcwd())
        if os.path.exists(base_dir):
            for folder in os.listdir(base_dir):
                if "Raafat_Engine" in folder:
                    p = os.path.join(base_dir, folder, "calibre", "ebook-convert.exe")
                    if os.path.exists(p):
                        self.calibre_path = p; self.log(f"✅ Calibre Detect: Masterpiece Mode ({folder})"); return

        self.log("⚠️ Calibre Engine Not Found.")
        msg = "محتاجين ننزل حاجه من النت حجمها 120 ميجا\n\n" + \
              "الخطوه ده بتحصل مره واحده بس ومش هتحتاج تكررها\n" + \
              "متنساش تصلي علي النبي"
        
        if QMessageBox.question(self, "صلي علي النبي", msg, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
            self.start_download()

    def start_download(self):
        self.set_ui_lock(True); self.status.setText("Status: Downloading..."); self.logs.clear()
        self.w_dl = DownloadWorker(); self.w_dl.progress_val.connect(self.progress_bar.setValue)
        self.w_dl.progress_msg.connect(self.log); self.w_dl.finished.connect(self.on_dl_fin); self.w_dl.start()

    def on_dl_fin(self, ok, res):
        self.set_ui_lock(False)
        if ok:
            self.calibre_path = res; self.log("✨ Setup Complete!"); QMessageBox.information(self, "Success", "Ready to convert!")
        else:
            self.status.setText("Setup Failed")
            err_msg = "معلش معرفتش انزل الباكدج اللي هتساعدنا بشكل اوتوماتيك جرب تشغلني تاني كمسؤول الاول ولو منفعش نزل البرنامج اللي هيساعدنا من هنا وبعدين اعمله setup علي الجهاز وتعالي هنا تاني"
            
            mbox = QMessageBox(self)
            mbox.setWindowTitle("صلي علي النبي")
            mbox.setText(err_msg)
            mbox.setIcon(QMessageBox.Icon.Critical)
            yes_btn = mbox.addButton("التوجه لصفحة التحميل", QMessageBox.ButtonRole.YesRole)
            no_btn = mbox.addButton("اغلاق البرنامج", QMessageBox.ButtonRole.NoRole)
            mbox.exec()
            
            if mbox.clickedButton() == yes_btn:
                webbrowser.open("https://calibre-ebook.com/download")
            elif mbox.clickedButton() == no_btn:
                self.close()

    def browse_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "Open", "", "Files (*.exe *.zip)"); self.select_file(p) if p else None
    def select_file(self, p): self.selected_path = p; self.file_display.setText(f"📄 File: {os.path.basename(p)}"); self.btn_run.setEnabled(True)
    def dragEnterEvent(self, e): e.accept() if e.mimeData().hasUrls() else e.ignore()
    def dropEvent(self, e): 
        p = e.mimeData().urls()[0].toLocalFile()
        if p.lower().endswith(('.exe', '.zip')): self.select_file(p)

    def run_conversion(self):
        target, _ = QFileDialog.getSaveFileName(self, "Save", os.path.splitext(os.path.basename(self.selected_path))[0]+".pdf", "PDF (*.pdf)")
        if not target: return
        self.set_ui_lock(True); self.worker = RaafatWorker(self.selected_path, target, self.calibre_path)
        self.worker.progress_msg.connect(self.log); self.worker.progress_val.connect(self.progress_bar.setValue)
        self.worker.finished.connect(lambda s, r: (self.set_ui_lock(False), QMessageBox.information(self, "Success", "Done!") if s else QMessageBox.warning(self, "Err", r)))
        self.worker.start()

    def cancel_conversion(self):
        if hasattr(self, 'worker') and self.worker.isRunning(): self.worker.cancel()

    def set_ui_lock(self, l):
        for b in [self.btn_browse, self.btn_run]: b.setEnabled(not l)
        self.btn_cancel.setEnabled(l)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = RaafatConverterApp(); window.show(); sys.exit(app.exec())
