import sys
import os
import sqlite3
import csv
from datetime import datetime
import random
import time

# Only lightweight imports at top
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QPushButton,
    QHBoxLayout, QFrame, QMainWindow, QTextEdit, QSlider,
    QFileDialog, QDialog, QListWidget, QListWidgetItem,
    QAction, QMenu, QComboBox, QMessageBox, QSizePolicy,
    QGraphicsOpacityEffect, QGraphicsDropShadowEffect
)
from PyQt5.QtGui import QPixmap, QFont, QImage, QIcon, QColor
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve

# Optional lightweight modules
spell = None
DOCX_AVAILABLE = False
PDF_AVAILABLE = False

try:
    from spellchecker import SpellChecker
    spell = SpellChecker()
except:
    pass

try:
    import docx
    DOCX_AVAILABLE = True
except:
    pass

try:
    from reportlab.pdfgen import canvas
    PDF_AVAILABLE = True
except:
    pass

# ---------- Helper for PyInstaller ----------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ---------- DB ----------
DB_PATH = resource_path("notes.db")
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()
cur.execute("""
CREATE TABLE IF NOT EXISTS notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    created_at TEXT NOT NULL,
    lang TEXT NOT NULL,
    content TEXT NOT NULL
)
""")
conn.commit()

# ---------- ANIMATED FRONT PAGE (YELLOW THEME) ----------
class AnimatedFrontPage(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Handwriting Recognition System")
        self.setFixedSize(950, 700)
        self.setStyleSheet("background-color: #fff8dc;")
        self.setWindowIcon(QIcon(resource_path("app_icon.ico")))

        self.init_ui()
        QTimer.singleShot(200, self.animate_startup)

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setAlignment(Qt.AlignTop)
        main_layout.setContentsMargins(30, 20, 30, 20)

        top_row_layout = QHBoxLayout()
        top_row_layout.setSpacing(20)

        left_frame = QFrame()
        left_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 3px solid #b30000;
                border-radius: 12px;
            }
        """)
        left_layout = QVBoxLayout(left_frame)
        poli_logo = QLabel()
        poli_logo_path = resource_path("poli logo.jpg")
        if os.path.exists(poli_logo_path):
            poli_pix = QPixmap(poli_logo_path)
            poli_logo.setPixmap(poli_pix.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            poli_logo.setText("POLI\nLOGO")
            poli_logo.setFont(QFont("Poppins", 10))
            poli_logo.setAlignment(Qt.AlignCenter)
        poli_logo.setAlignment(Qt.AlignCenter)
        left_layout.addWidget(poli_logo)

        left_shadow = QGraphicsDropShadowEffect()
        left_shadow.setBlurRadius(25)
        left_shadow.setColor(QColor(0, 0, 0, 80))
        left_frame.setGraphicsEffect(left_shadow)
        top_row_layout.addWidget(left_frame, alignment=Qt.AlignLeft)

        title = QLabel("POLITEKNIK UNGKU OMAR")
        title.setFont(QFont("Poppins", 24, QFont.Bold))
        title.setStyleSheet("color: #b30000; letter-spacing: 1px;")
        title.setAlignment(Qt.AlignCenter)
        top_row_layout.addStretch(1)
        top_row_layout.addWidget(title, alignment=Qt.AlignCenter)
        top_row_layout.addStretch(1)

        right_frame = QFrame()
        right_frame.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 3px solid #b30000;
                border-radius: 12px;
            }
        """)
        right_layout = QVBoxLayout(right_frame)
        app_logo = QLabel()
        app_logo_path = resource_path("logo.png")
        if os.path.exists(app_logo_path):
            app_pix = QPixmap(app_logo_path)
            app_logo.setPixmap(app_pix.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        else:
            app_logo.setText("APP\nLOGO")
            app_logo.setFont(QFont("Poppins", 10))
            app_logo.setAlignment(Qt.AlignCenter)
        app_logo.setAlignment(Qt.AlignCenter)
        right_layout.addWidget(app_logo)

        right_shadow = QGraphicsDropShadowEffect()
        right_shadow.setBlurRadius(25)
        right_shadow.setColor(QColor(0, 0, 0, 80))
        right_frame.setGraphicsEffect(right_shadow)
        top_row_layout.addWidget(right_frame, alignment=Qt.AlignRight)

        main_layout.addLayout(top_row_layout)

        subtitle = QLabel("JABATAN KEJURUTERAAN ELEKTRIK")
        subtitle.setFont(QFont("Poppins", 14))
        subtitle.setStyleSheet("color: #444;")
        subtitle.setAlignment(Qt.AlignCenter)

        main_title = QLabel("HANDWRITING RECOGNITION SYSTEM")
        main_title.setFont(QFont("Poppins", 20, QFont.Bold))
        main_title.setStyleSheet("color: #1a1a1a; margin-top: 6px;")
        main_title.setAlignment(Qt.AlignCenter)

        main_layout.addWidget(subtitle)
        main_layout.addWidget(main_title)

        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: white;
                border-radius: 18px;
                border: 2px solid #b30000;
                padding: 20px;
            }
        """)
        info = QLabel(
            "📘 FINAL YEAR PROJECT 2025/2026\n\n"
            "👨‍💻 Developed by:\nSHAARANGIN A/L K JAYA SANGKAR (01DTK23F1076)\n\n"
            "🧑‍🏫 Supervised by:\nMR SHAIFOL IFRAD BIN IBRAHIM"
        )
        info.setFont(QFont("Segoe UI", 13))
        info.setStyleSheet("color: #222;")
        info.setAlignment(Qt.AlignCenter)
        info.setWordWrap(True)
        info.setMaximumWidth(600)

        card_layout = QVBoxLayout(card)
        card_layout.addWidget(info, alignment=Qt.AlignCenter)

        start_btn = QPushButton("START SYSTEM ▶️")
        start_btn.setFont(QFont("Poppins", 14, QFont.Bold))
        start_btn.setFixedSize(250, 58)
        start_btn.setStyleSheet("""
            QPushButton {
                background-color: qlineargradient(
                    spread:pad, x1:0, y1:0, x2:1, y2:0,
                    stop:0 #b30000, stop:1 #e6b800
                );
                color: white;
                border-radius: 29px;
                border: none;
            }
            QPushButton:hover {
                background-color: #e63900;
            }
        """)
        start_btn.clicked.connect(self.start_app)

        main_layout.addSpacing(12)
        main_layout.addWidget(card, alignment=Qt.AlignCenter)
        main_layout.addSpacing(20)
        main_layout.addWidget(start_btn, alignment=Qt.AlignCenter)
        self.setLayout(main_layout)

        self.widgets = [title, subtitle, main_title, card, start_btn, left_frame, right_frame]

    def animate_startup(self):
        delay = 0
        for widget in self.widgets:
            opacity = QGraphicsOpacityEffect()
            widget.setGraphicsEffect(opacity)

            anim = QPropertyAnimation(opacity, b"opacity")
            anim.setDuration(800)
            anim.setStartValue(0)
            anim.setEndValue(1)
            anim.setEasingCurve(QEasingCurve.InOutQuad)

            QTimer.singleShot(delay, anim.start)
            delay += 200

    def start_app(self):
        self.ocr_window = OCRWindow()
        self.ocr_window.show()
        self.hide()


# ---------- OCR Window ----------
class OCRWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Handwriting Recognition System - OCR")
        self.setMinimumSize(1100, 780)
        self.resize(1100, 780)
        self.is_dark_mode = False

        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            png_icon = resource_path("logo.png")
            if os.path.exists(png_icon):
                self.setWindowIcon(QIcon(png_icon))

        self.captured_image = None
        self.captured_frame = None
        self.font_size = 14
        self.current_lang = "en"
        self.ratios = {"16:9": 16/9, "4:3": 4/3, "1:1": 1}
        self.current_ratio = "16:9"
        self.available_cams = self.detect_cameras()
        self.camera_index = -1
        self.cap = None
        self.easyocr_reader = None
        self.easyocr_reader_lang = None

        self.FUNNY_MESSAGES = [
            "🧠 Aiyo… teaching robot to read your writing… even my nenek write better than this lah!",
            "📝 Translating ‘mystery letters’… even doctor prescription easier to read lah!",
            "☕ You go make kopi first — this one take time to understand your style!"
        ]
        self.current_tip_index = 0

        self.init_ui()
        self.update_status("Ready")

        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.timer.start(33)

    def set_light_mode_base(self):
        self.setStyleSheet("""
            QMainWindow, QFrame { background-color: #ffffff; }
            QLabel { color: #222; }
            QTextEdit { background: #fff0f5; color: #111; border-radius: 10px; padding: 10px; }
            QPushButton { background: #2196F3; color: white; border-radius: 10px; padding: 10px 18px; font-weight: 600; }
            QPushButton:hover { background: #1976D2; }
            QComboBox {
                background-color: #f0f0f0;
                color: #000;
                border-radius: 5px;
                padding: 4px;
                min-width: 90px;
            }
            QSlider::groove:horizontal {
                height: 6px;
                background: #ccc;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                background: #2196F3;
                border: 1px solid #1976D2;
                width: 16px;
                margin: -5px 0;
                border-radius: 8px;
            }
            QStatusBar {
                background: #f5f5f5;
                color: #111;
                border-top: 1px solid #ccc;
            }
        """)
        self.is_dark_mode = False
        self.dos_label.setStyleSheet("color: #222;")
        self.donts_label.setStyleSheet("color: #222;")

    def set_dark_mode_base(self):
        self.setStyleSheet("""
            QMainWindow, QFrame { background-color: #1e1e1e; }
            QLabel { color: #ffffff; }
            QTextEdit { background: #2a2a2a; color: #ffffff; border-radius: 10px; padding: 10px; }
            QPushButton { background-color: #FF0000; color: white; border-radius: 8px; padding: 10px 18px; font-weight: 600; }
            QPushButton:hover { background-color: #cc0000; }
            QComboBox {
                background-color: #2a2a2a;
                color: #ffffff;
                border: 1px solid #555;
                border-radius: 5px;
                padding: 4px;
                min-width: 90px;
            }
            QComboBox QAbstractItemView {
                background-color: #2a2a2a;
                color: #ffffff;
                selection-background-color: #444;
                selection-color: #fff;
            }
            QSlider::groove:horizontal {
                height: 6px;
                background: #555;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                background: #FF0000;
                border: 1px solid #cc0000;
                width: 16px;
                margin: -5px 0;
                border-radius: 8px;
            }
            QStatusBar {
                background: #1e1e1e;
                color: #fff;
                border-top: 1px solid #444;
            }
        """)
        self.is_dark_mode = True
        self.dos_label.setStyleSheet("color: white;")
        self.donts_label.setStyleSheet("color: white;")

    def detect_cameras(self, max_check=5):
        # ✅ LAZY IMPORT
        import cv2
        cams = []
        for i in range(max_check):
            try:
                cap = cv2.VideoCapture(i, cv2.CAP_DSHOW) if os.name == 'nt' else cv2.VideoCapture(i)
                if cap is None or not cap.isOpened():
                    if cap:
                        cap.release()
                    continue
                ret, _ = cap.read()
                if ret:
                    cams.append(i)
                cap.release()
            except:
                continue
        return cams

    def init_ui(self):
        central = QFrame()
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(14, 10, 14, 10)
        main_layout.setSpacing(14)

        top_section = QHBoxLayout()
        top_section.setSpacing(14)

        # === LEFT PANEL: Do's & Don'ts ===
        left_frame = QFrame()
        left_layout = QVBoxLayout(left_frame)
        left_layout.setSpacing(12)

        # DO's Box (Green)
        dos_box = QFrame()
        dos_box.setStyleSheet("""
            QFrame {
                background-color: #e8f5e9;
                border: 2px solid #4caf50;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        self.dos_label = QLabel(
            "✅ <b>DO’s</b><br>"
            "1. ✍️ Write in <b>BLOCK CAPITALS</b><br>"
            "2. ⏹️ Leave space: <b>H E L L O</b><br>"
            "3. 📄 Use dark ink on white paper"
        )
        self.dos_label.setFont(QFont("Segoe UI", 9))
        self.dos_label.setWordWrap(True)
        dos_layout = QVBoxLayout(dos_box)
        dos_layout.addWidget(self.dos_label)
        left_layout.addWidget(dos_box)

        # DON'Ts Box (Red)
        donts_box = QFrame()
        donts_box.setStyleSheet("""
            QFrame {
                background-color: #ffebee;
                border: 2px solid #f44336;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        self.donts_label = QLabel(
            "❌ <b>DON’Ts</b><br>"
            "1. 🚫 Avoid cursive or connected letters<br>"
            "2. 🚫 Don’t smudge or press lightly<br>"
            "3. 🚫 No overlapping words or doodles"
        )
        self.donts_label.setFont(QFont("Segoe UI", 9))
        self.donts_label.setWordWrap(True)
        donts_layout = QVBoxLayout(donts_box)
        donts_layout.addWidget(self.donts_label)
        left_layout.addWidget(donts_box)

        # Polite note
        note_label = QLabel("🙏 We’re still improving accuracy — sorry for the inconvenience!")
        note_label.setFont(QFont("Segoe UI", 8))
        note_label.setWordWrap(True)
        note_label.setStyleSheet("color: #555;")
        left_layout.addWidget(note_label)
        left_layout.addStretch()

        # === CENTER: Camera Feed ===
        center_layout = QVBoxLayout()
        center_layout.setSpacing(10)

        self.video_label = QLabel("Webcam Feed")
        self.video_label.setAlignment(Qt.AlignCenter)
        self.video_label.setStyleSheet("QLabel{background:#111;color:#eee;border-radius:10px;}")
        self.video_label.setFixedSize(960, 540)
        center_layout.addWidget(self.video_label, alignment=Qt.AlignCenter)

        self.capture_btn = QPushButton("Capture Image")
        self.capture_btn.setFixedHeight(35)
        self.capture_btn.clicked.connect(self.capture_image)
        center_layout.addWidget(self.capture_btn, alignment=Qt.AlignCenter)

        # === RIGHT PANEL: Grey Control Panel ===
        self.right_frame = QFrame()
        self.right_frame.setFixedWidth(240)
        self.right_frame.setStyleSheet("""
            QFrame {
                background-color: #f0f0f0;
                border-radius: 10px;
                padding: 12px;
                border: 1px solid #ccc;
            }
        """)
        right_layout = QVBoxLayout(self.right_frame)
        right_layout.setSpacing(12)

        # Font slider
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("Font:"))
        self.font_slider = QSlider(Qt.Horizontal)
        self.font_slider.setRange(8, 28)
        self.font_slider.setValue(14)
        self.font_slider.setFixedHeight(24)
        self.font_slider.valueChanged.connect(self.change_font_size)
        font_layout.addWidget(self.font_slider)
        right_layout.addLayout(font_layout)

        # Batman Mode button
        self.dark_btn = QPushButton("🌙 Batman Mode")
        self.dark_btn.setFixedHeight(36)
        self.dark_btn.clicked.connect(self.toggle_dark_mode)
        right_layout.addWidget(self.dark_btn)

        # Camera selection
        cam_layout = QHBoxLayout()
        cam_layout.addWidget(QLabel("Camera:"))
        self.camera_dropdown = QComboBox()
        self.camera_dropdown.addItem("Off", -1)
        for i in self.available_cams:
            self.camera_dropdown.addItem(f"Cam {i}", i)
        self.camera_dropdown.currentIndexChanged.connect(self.change_camera)
        cam_layout.addWidget(self.camera_dropdown)
        right_layout.addLayout(cam_layout)

        # Aspect ratio
        ratio_layout = QHBoxLayout()
        ratio_layout.addWidget(QLabel("Ratio:"))
        self.ratio_box = QComboBox()
        self.ratio_box.addItems(self.ratios.keys())
        self.ratio_box.setCurrentText(self.current_ratio)
        self.ratio_box.currentTextChanged.connect(self.change_ratio)
        ratio_layout.addWidget(self.ratio_box)
        right_layout.addLayout(ratio_layout)

        # Tips
        self.tips_label = QLabel("💡 Write in block capitals with spaces: EXAMPLE: H E L L O")
        self.tips_label.setWordWrap(True)
        self.tips_label.setFont(QFont("Segoe UI", 9))
        right_layout.addWidget(self.tips_label)

        right_layout.addStretch()

        top_section.addWidget(left_frame)
        top_section.addLayout(center_layout)
        top_section.addWidget(self.right_frame)

        # === STATUS BAR (GRAY BACKGROUND, BLACK TEXT ALWAYS) ===
        self.status_frame = QFrame()
        self.status_frame.setFixedHeight(50)
        status_layout = QVBoxLayout(self.status_frame)
        self.custom_status_label = QLabel("Ready")
        self.custom_status_label.setContentsMargins(10, 0, 0, 0)
        self.custom_status_label.setFont(QFont("Segoe UI", 10))
        self.custom_status_label.setStyleSheet("color: black;")
        status_layout.addWidget(self.custom_status_label)
        self.status_frame.setStyleSheet("QFrame { background: #d9d9d9; border: 1px solid #aaa; }")

        self.statusBar().setStyleSheet("""
            background: #d9d9d9;
            color: black;
            border-top: 1px solid #aaa;
        """)

        # Result Display
        self.result_display = QTextEdit()
        self.result_display.setReadOnly(False)
        self.result_display.setFont(QFont("Segoe UI", self.font_size))
        self.result_display.setStyleSheet("""
            QTextEdit { 
                background: #fff0f5; 
                color: #111; 
                border: 3px solid #FFA500; 
                border-radius: 10px; 
                padding: 10px; 
            }
        """)

        main_layout.addLayout(top_section)
        main_layout.addWidget(self.status_frame)
        main_layout.addWidget(self.result_display)

        self.setCentralWidget(central)

        # Menu
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("☰ Menu")

        save_action = QAction("Save As (.txt/.docx)", self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)

        export_pdf_action = QAction("Export PDF", self)
        export_pdf_action.triggered.connect(self.export_pdf)
        file_menu.addAction(export_pdf_action)

        export_csv_action = QAction("Export CSV", self)
        export_csv_action.triggered.connect(self.export_csv)
        file_menu.addAction(export_csv_action)

        copy_action = QAction("Copy to Clipboard", self)
        copy_action.triggered.connect(self.copy_clipboard)
        file_menu.addAction(copy_action)

        spell_action = QAction("Spell Check (EN)", self)
        spell_action.triggered.connect(self.spell_check)
        file_menu.addAction(spell_action)

        clear_action = QAction("Clear Text", self)
        clear_action.triggered.connect(self.clear_text)
        file_menu.addAction(clear_action)

        history_action = QAction("Show History", self)
        history_action.triggered.connect(self.show_history)
        file_menu.addAction(history_action)

        lang_menu = QMenu("Language", self)
        en_action = QAction("English", self)
        en_action.triggered.connect(lambda: self.set_language("en"))
        lang_menu.addAction(en_action)
        file_menu.addMenu(lang_menu)

    def update_status(self, message):
        self.custom_status_label.setText(message)

    def change_font_size(self, value):
        self.result_display.setFont(QFont("Segoe UI", value))

    def toggle_dark_mode(self):
        current_font = self.result_display.font()

        if self.is_dark_mode:
            self.set_light_mode_base()
            self.status_frame.setStyleSheet("QFrame { background: #d9d9d9; border: 1px solid #aaa; }")
            self.result_display.setStyleSheet("""
                QTextEdit { 
                    background: #fff0f5; 
                    color: #111; 
                    border: 3px solid #FFA500; 
                    border-radius: 10px; 
                    padding: 10px; 
                }
            """)
            self.tips_label.setStyleSheet("color: #000;")

            for widget in self.findChildren(QFrame):
                ss = widget.styleSheet()
                if "border: 2px solid #4caf50" in ss:
                    widget.setStyleSheet("""
                        QFrame {
                            background-color: #e8f5e9;
                            border: 2px solid #4caf50;
                            border-radius: 8px;
                            padding: 10px;
                        }
                    """)
                    label = widget.findChild(QLabel)
                    if label:
                        label.setStyleSheet("color: #222;")
                elif "border: 2px solid #f44336" in ss:
                    widget.setStyleSheet("""
                        QFrame {
                            background-color: #ffebee;
                            border: 2px solid #f44336;
                            border-radius: 8px;
                            padding: 10px;
                        }
                    """)
                    label = widget.findChild(QLabel)
                    if label:
                        label.setStyleSheet("color: #222;")

            combo_light = """
                QComboBox {
                    background-color: #f0f0f0;
                    color: #000;
                    border-radius: 5px;
                    padding: 4px;
                    min-width: 90px;
                }
            """
            self.camera_dropdown.setStyleSheet(combo_light)
            self.ratio_box.setStyleSheet(combo_light)

            self.right_frame.setStyleSheet("""
                QFrame {
                    background-color: #f0f0f0;
                    border-radius: 10px;
                    padding: 12px;
                    border: 1px solid #ccc;
                }
            """)

        else:
            self.set_dark_mode_base()
            self.status_frame.setStyleSheet("QFrame { background: #d9d9d9; border: 1px solid #aaa; }")
            self.result_display.setStyleSheet("""
                QTextEdit { 
                    background: #2a2a2a; 
                    color: #ffffff; 
                    border: 3px solid #FFA500; 
                    border-radius: 10px; 
                    padding: 10px; 
                }
            """)
            self.tips_label.setStyleSheet("color: white;")

            for widget in self.findChildren(QFrame):
                ss = widget.styleSheet()
                if "border: 2px solid #4caf50" in ss:
                    widget.setStyleSheet("""
                        QFrame {
                            background-color: #1b5e20;
                            border: 2px solid #4caf50;
                            border-radius: 8px;
                            padding: 10px;
                        }
                    """)
                    label = widget.findChild(QLabel)
                    if label:
                        label.setStyleSheet("color: white;")
                elif "border: 2px solid #f44336" in ss:
                    widget.setStyleSheet("""
                        QFrame {
                            background-color: #b71c1c;
                            border: 2px solid #f44336;
                            border-radius: 8px;
                            padding: 10px;
                        }
                    """)
                    label = widget.findChild(QLabel)
                    if label:
                        label.setStyleSheet("color: white;")

            combo_dark = """
                QComboBox {
                    background-color: #2a2a2a;
                    color: white;
                    border: 1px solid #555;
                    border-radius: 5px;
                    padding: 4px;
                    min-width: 90px;
                }
                QComboBox QAbstractItemView {
                    background-color: #2a2a2a;
                    color: white;
                    selection-background-color: #444;
                    selection-color: #fff;
                }
            """
            self.camera_dropdown.setStyleSheet(combo_dark)
            self.ratio_box.setStyleSheet(combo_dark)

            self.right_frame.setStyleSheet("""
                QFrame {
                    background-color: #2a2a2a;
                    border-radius: 10px;
                    padding: 12px;
                    border: 1px solid #555;
                }
            """)

        self.result_display.setFont(current_font)
        self.dark_btn.setText("☀️ Light Mode" if self.is_dark_mode else "🌙 Batman Mode")

    def change_camera(self, combo_index):
        # ✅ LAZY IMPORT
        import cv2
        cam_index = self.camera_dropdown.itemData(combo_index)
        if self.cap is not None and self.cap.isOpened():
            self.cap.release()
            self.cap = None

        if cam_index == -1:
            self.update_status("Camera turned off")
            self.video_label.setText("Camera Off")
            return

        self.cap = cv2.VideoCapture(cam_index, cv2.CAP_DSHOW)
        if not self.cap.isOpened():
            QMessageBox.critical(self, "Camera Error", f"Cannot open camera {cam_index}")
            self.camera_dropdown.setCurrentIndex(0)
            self.video_label.setText("Camera Error")
            return

        self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        time.sleep(0.3)
        self.camera_index = cam_index
        self.update_status(f"Camera {cam_index} activated (720p)")

    def change_ratio(self, text):
        self.current_ratio = text

    def set_language(self, lang):
        self.current_lang = lang
        self.update_status(f"Language: {lang}")

    def update_frame(self):
        if self.camera_index == -1 or self.cap is None or not self.cap.isOpened():
            return

        # ✅ LAZY IMPORT
        import cv2
        ret, frame = self.cap.read()
        if not ret or frame is None:
            return

        h, w = frame.shape[:2]
        target_ratio = self.ratios[self.current_ratio]
        current_ratio = w / h
        if current_ratio > target_ratio:
            new_w = int(h * target_ratio)
            crop_x = (w - new_w) // 2
            disp_frame = frame[:, crop_x:crop_x + new_w]
        else:
            new_h = int(w / target_ratio)
            crop_y = (h - new_h) // 2
            disp_frame = frame[crop_y:crop_y + new_h, :]

        self.captured_frame = frame.copy()

        rgb_image = cv2.cvtColor(disp_frame, cv2.COLOR_BGR2RGB)
        h, w, ch = rgb_image.shape
        bytes_per_line = ch * w
        qt_image = QImage(rgb_image.data, w, h, bytes_per_line, QImage.Format_RGB888)
        pixmap = QPixmap.fromImage(qt_image)
        self.video_label.setPixmap(pixmap.scaled(960, 540, Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def capture_image(self):
        if self.captured_frame is not None:
            ocr_image = self.captured_frame.copy()

            try:
                # ✅ LAZY IMPORT
                import cv2
                import numpy as np
                gray = cv2.cvtColor(ocr_image, cv2.COLOR_BGR2GRAY)
                gray = cv2.resize(gray, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
                denoised = cv2.fastNlMeansDenoising(gray, h=10)
                kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
                sharpened = cv2.filter2D(denoised, -1, kernel)
                _, thresh = cv2.threshold(sharpened, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                ocr_image = cv2.cvtColor(thresh, cv2.COLOR_GRAY2BGR)
            except Exception as e:
                self.update_status(f"Preprocessing failed: {e}")
                pass

            self.captured_image = ocr_image
            self.perform_ocr()
        else:
            self.update_status("No frame to capture!")

    def perform_ocr(self):
        if self.captured_image is None:
            self.result_display.setPlainText("No image captured!")
            return

        funny_msg = self.FUNNY_MESSAGES[self.current_tip_index % len(self.FUNNY_MESSAGES)]
        self.update_status(f"🧠 {funny_msg}")
        self.current_tip_index += 1
        QApplication.processEvents()

        # ✅ LAZY IMPORT + FIRST-TIME LOAD MESSAGE
        if self.current_lang != self.easyocr_reader_lang or self.easyocr_reader is None:
            self.update_status("🧠 Loading handwriting AI model... (one-time, ~5-10 sec)")
            QApplication.processEvents()
            try:
                import easyocr
                self.easyocr_reader = easyocr.Reader([self.current_lang], gpu=False)
                self.easyocr_reader_lang = self.current_lang
            except Exception as e:
                self.result_display.setPlainText(f"[EasyOCR Load Failed] {str(e)}")
                return

        try:
            import cv2
            rgb_image = cv2.cvtColor(self.captured_image, cv2.COLOR_BGR2RGB)
            results = self.easyocr_reader.readtext(rgb_image, detail=0, paragraph=True)
            text = "\n".join(results) if results else "[No text detected]"

            self.result_display.setPlainText(text)
            self.update_status("✅ AI OCR Success")

            if text.strip() and text != "[No text detected]":
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cur.execute("INSERT INTO notes (created_at, lang, content) VALUES (?, ?, ?)",
                            (now, self.current_lang, text))
                conn.commit()

        except Exception as e:
            self.result_display.setPlainText(f"[AI OCR FAILED] {str(e)}")
            self.update_status("⚠️ Falling back to Tesseract")

            try:
                import cv2
                import pytesseract
                gray = cv2.cvtColor(self.captured_image, cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                text = pytesseract.image_to_string(thresh, lang='eng', config='--psm 6 --oem 3')
                self.result_display.setPlainText(f"[FALLBACK]\n{text}")
            except Exception as te:
                self.result_display.setPlainText(f"[TESSERACT FAILED] {str(te)}")

    # --- File Operations ---
    def save_file(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Text Files (*.txt);;Word Documents (*.docx)")
        if not filename:
            return
        content = self.result_display.toPlainText()
        if filename.endswith(".txt"):
            with open(filename, "w", encoding="utf-8") as f:
                f.write(content)
            self.update_status("Saved as .txt")
        elif filename.endswith(".docx") and DOCX_AVAILABLE:
            import docx
            doc = docx.Document()
            doc.add_paragraph(content)
            doc.save(filename)
            self.update_status("Saved as .docx")
        elif filename.endswith(".docx"):
            QMessageBox.warning(self, "Save Error", "Install python-docx")

    def export_pdf(self):
        if not PDF_AVAILABLE:
            QMessageBox.warning(self, "PDF Export", "Install reportlab")
            return
        filename, _ = QFileDialog.getSaveFileName(self, "Export PDF", "", "PDF Files (*.pdf)")
        if not filename:
            return
        from reportlab.pdfgen import canvas
        c = canvas.Canvas(filename)
        text = self.result_display.toPlainText()
        y = 800
        for line in text.splitlines():
            c.drawString(20, y, line[:100])
            y -= 15
            if y < 50:
                c.showPage()
                y = 800
        c.save()
        self.update_status("Exported to PDF")

    def export_csv(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Export CSV", "", "CSV Files (*.csv)")
        if not filename:
            return
        content = self.result_display.toPlainText()
        with open(filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            for line in content.splitlines():
                writer.writerow([line])
        self.update_status("Exported to CSV")

    def copy_clipboard(self):
        QApplication.clipboard().setText(self.result_display.toPlainText())
        self.update_status("Copied to clipboard")

    def spell_check(self):
        if not spell:
            QMessageBox.warning(self, "Spell Check", "Install pyspellchecker")
            return
        text = self.result_display.toPlainText()
        words = text.split()
        corrected = [spell.correction(w) or w for w in words]
        self.result_display.setPlainText(" ".join(corrected))
        self.update_status("Spell check applied")

    def clear_text(self):
        self.result_display.clear()
        self.update_status("Text cleared")

    def show_history(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Recent Notes")
        dlg.setFixedSize(600, 450)
        layout = QVBoxLayout(dlg)
        list_widget = QListWidget()
        cur.execute("SELECT id, created_at, lang, content FROM notes ORDER BY id DESC LIMIT 20")
        rows = cur.fetchall()
        for row in rows:
            item = QListWidgetItem(f"{row[1]} | {row[2]}:\n{row[3]}")
            list_widget.addItem(item)
        layout.addWidget(list_widget)
        dlg.exec_()

    def closeEvent(self, event):
        if self.cap is not None and self.cap.isOpened():
            self.cap.release()
        conn.close()
        event.accept()


# ---------- Run App ----------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AnimatedFrontPage()
    window.show()
    sys.exit(app.exec_())