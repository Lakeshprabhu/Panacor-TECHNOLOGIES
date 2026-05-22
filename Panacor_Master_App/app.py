import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QStackedWidget, QGridLayout, QGraphicsDropShadowEffect,
    QSizePolicy
)
from PyQt6.QtCore import Qt, QPropertyAnimation, QEasingCurve
from PyQt6.QtGui import QPixmap, QFont

current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

from pages.page_quotation import PageQuotation
from pages.page_legends import PageLegends
from pages.page_pdf2excel import PagePdf2Excel
from pages.page_om_manual import PageOMManual
from pages.page_settings import PageSettings
from core.config import config_manager

NAV_ITEMS = [
    ("🏠 Dashboard", "Dashboard Overview"),
    ("📝 Quotation AI", "Quotation Automation Engine"),
    ("📐 Legends Maker", "Legends Template Maker"),
    ("📊 PDF to Excel", "PDF to Excel Extractor"),
    ("📚 O&M Manuals", "O&M Manual Preparation Builder"),
]

DASHBOARD_TILES = [
    ("📝", "Quotation Automation", "Extract BOQ data from PDFs and automatically populate Excel templates using AI.", 1),
    ("📐", "Legends Template Maker", "Extract schematic symbols and descriptions visually into Excel tables.", 2),
    ("📊", "PDF to Excel Extractor", "Line-by-line smart text extraction from any document directly to spreadsheets.", 3),
    ("📚", "O&M Manuals Builder", "Auto-search, download, and merge technical datasheets into complete Manuals.", 4),
]

THEMES = {
    "Purple": {"bg": "#F5F3FA", "primary": "#5A3286", "sidebar": "#FFFFFF", "text_main": "#2D3748"},
    "Black": {"bg": "#121212", "primary": "#2D2D2D", "sidebar": "#1F1F1F", "text_main": "#E0E0E0"},
    "Blue": {"bg": "#EBF8FF", "primary": "#2B6CB0", "sidebar": "#FFFFFF", "text_main": "#2A4365"},
    "Emerald": {"bg": "#F0FFF4", "primary": "#22543D", "sidebar": "#FFFFFF", "text_main": "#276749"},
    "Crimson": {"bg": "#FFF5F5", "primary": "#9B2C2C", "sidebar": "#FFFFFF", "text_main": "#742A2A"},
}
        
class PanacorMasterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Panacor Technologies - Master Intelligence Hub")
        self.setMinimumSize(1100, 800)
        
        style_path = os.path.join(current_dir, "styles.qss")
        self.base_stylesheet = ""
        if os.path.exists(style_path):
            with open(style_path, "r", encoding="utf-8") as f:
                self.base_stylesheet = f.read()
                
        self.init_ui()
        self.apply_theme()
        
        if config_manager.get("startup_animation", True):
            self.run_startup_animation()
        
    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_lay = QHBoxLayout(central)
        main_lay.setContentsMargins(0, 0, 0, 0)
        main_lay.setSpacing(0)

        # Sidebar
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(250)
        self.sidebar_layout = QVBoxLayout(sidebar)
        self.sidebar_layout.setContentsMargins(10, 20, 10, 20)
        self.sidebar_layout.setSpacing(5)

        logo = QLabel()
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_path = os.path.join(current_dir, "assets", "logo.png")
        if os.path.exists(logo_path):
            logo.setPixmap(QPixmap(logo_path).scaled(
                180, 80, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            ))
        else:
            logo.setText("PANACOR")
            logo.setObjectName("brand_title")
        self.sidebar_layout.addWidget(logo)
        self.sidebar_layout.addSpacing(30)

        self.nav_buttons = []
        for text, _ in NAV_ITEMS:
            self._add_nav(text)
        self.sidebar_layout.addStretch()
        self._add_nav("⚙️ Settings")

        ver = QLabel("v1.0.0 Master")
        ver.setObjectName("brand_subtitle")
        ver.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.sidebar_layout.addWidget(ver)
        main_lay.addWidget(sidebar)

        # Content area
        content = QWidget()
        content.setObjectName("content_area")
        content_lay = QVBoxLayout(content)
        content_lay.setContentsMargins(0, 0, 0, 0)
        content_lay.setSpacing(0)

        header = QWidget()
        header.setObjectName("header")
        header.setFixedHeight(70)
        h_lay = QHBoxLayout(header)
        h_lay.setContentsMargins(30, 0, 30, 0)
        self.header_title = QLabel("Dashboard Overview")
        self.header_title.setObjectName("brand_title")
        h_lay.addWidget(self.header_title)
        content_lay.addWidget(header)

        self.stacked = QStackedWidget()
        for p in [self._dashboard(), PageQuotation(), PageLegends(),
                  PagePdf2Excel(), PageOMManual(), PageSettings(self)]:
            self.stacked.addWidget(p)
        content_lay.addWidget(self.stacked)
        main_lay.addWidget(content)
        self.nav_buttons[0].setChecked(True)

    def _add_nav(self, text):
        btn = QPushButton(text)
        btn.setObjectName("SidebarButton")
        btn.setProperty("class", "SidebarButton")
        btn.setCheckable(True)
        idx = len(self.nav_buttons)
        btn.clicked.connect(lambda _, i=idx, b=btn: self._switch(i, b))
        self.sidebar_layout.addWidget(btn)
        self.nav_buttons.append(btn)

    def _switch(self, index, button):
        for b in self.nav_buttons:
            b.setChecked(b is button)
        self.stacked.setCurrentIndex(index)
        titles = [t for _, t in NAV_ITEMS] + ["Settings & Configuration"]
        self.header_title.setText(titles[index])
        
    def _dashboard(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)

        welcome = QLabel("Welcome to the Panacor Intelligence Hub")
        welcome.setObjectName("brand_title")
        layout.addWidget(welcome)
        sub = QLabel("Select a tool below or from the sidebar menu to begin your workflow.")
        sub.setObjectName("brand_subtitle")
        layout.addWidget(sub)
        layout.addSpacing(20)

        grid = QGridLayout()
        grid.setSpacing(30)
        for i, (emoji, title, desc, target) in enumerate(DASHBOARD_TILES):
            grid.addWidget(self._tile(emoji, title, desc, target), i // 2, i % 2)
        grid.setRowStretch(2, 1)
        layout.addLayout(grid)
        return page

    def _tile(self, emoji, title, desc, target_index):
        btn = QPushButton()
        btn.setProperty("class", "HeroTile")
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        btn.setMinimumHeight(200)

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setXOffset(0)
        shadow.setYOffset(5)
        shadow.setColor(Qt.GlobalColor.lightGray)
        btn.setGraphicsEffect(shadow)

        lay = QVBoxLayout(btn)
        lay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        icon_lbl = QLabel(emoji)
        icon_lbl.setFont(QFont("Segoe UI", 48))
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(icon_lbl)

        title_lbl = QLabel(title)
        title_lbl.setProperty("class", "CardTitle")
        title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(title_lbl)

        desc_lbl = QLabel(desc)
        desc_lbl.setProperty("class", "BodyText")
        desc_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc_lbl.setWordWrap(True)
        desc_lbl.setFixedWidth(280)
        lay.addWidget(desc_lbl)

        btn.clicked.connect(lambda: self._switch(target_index, self.nav_buttons[target_index]))
        return btn

    def apply_theme(self):
        theme = config_manager.get("theme", "Default")
        qss = "QWidget { color: #2D3748; }\n" + self.base_stylesheet

        if theme in THEMES:
            c = THEMES[theme]
            qss = qss.replace("#F0F4F8", c["bg"])

            if theme == "Black":
                for old, new in {
                    "color: #28235D;": "color: #E0E0E0;",
                    "background-color: #28235D;": f"background-color: {c['primary']};",
                    "border: 1px solid #28235D;": f"border: 1px solid {c['primary']};",
                    "border-left: 4px solid #E84545;": "border-left: 4px solid #FF4500;",
                    "#FFFFFF": c["sidebar"], "#2D3748": c["text_main"],
                    "#718096": "#A0AEC0", "#E2E8F0": "#333333",
                    "#F7FAFC": "#242424", "#FAFBFC": "#2A2A2A",
                    "#EBF4FF": "#333333", "#2B6CB0": "#E0E0E0",
                    "#CBD5E0": "#444444",
                    "background-color: #FFF5F5;": "background-color: #3A1C1C;",
                    "border: 1px solid #FEB2B2;": "border: 1px solid #E53E3E;",
                    "background-color: #FED7D7;": "background-color: #4A2222;",
                }.items():
                    qss = qss.replace(old, new)
            qss = qss.replace("#28235D", c["primary"])

        self.setStyleSheet(qss)

    def run_startup_animation(self):
        self.setWindowOpacity(0.0)
        self.anim = QPropertyAnimation(self, b"windowOpacity")
        self.anim.setDuration(1200)
        self.anim.setStartValue(0.0)
        self.anim.setEndValue(1.0)
        self.anim.setEasingCurve(QEasingCurve.Type.InOutQuad)
        self.anim.start()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PanacorMasterApp()
    window.show()
    sys.exit(app.exec())
