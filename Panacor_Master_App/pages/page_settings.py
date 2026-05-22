import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QLineEdit, QComboBox, QCheckBox, QMessageBox
)
from PyQt6.QtCore import Qt
from core.config import config_manager


class PageSettings(QWidget):
    def __init__(self, main_app):
        super().__init__()
        self.main_app = main_app
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 20, 40, 20)
        layout.setSpacing(20)

        # Appearance
        app_card = QWidget()
        app_card.setProperty("class", "Card")
        app_lay = QVBoxLayout(app_card)
        app_lay.setContentsMargins(20, 20, 20, 20)
        app_lay.setSpacing(15)

        lbl = QLabel("🎨 Appearance")
        lbl.setProperty("class", "CardTitle")
        app_lay.addWidget(lbl)

        theme_row = QHBoxLayout()
        t_lbl = QLabel("Color Scheme:")
        t_lbl.setProperty("class", "BodyText")
        theme_row.addWidget(t_lbl)
        self.combo_theme = QComboBox()
        self.combo_theme.addItems(["Default", "Purple", "Black", "Blue", "Emerald", "Crimson"])
        self.combo_theme.currentTextChanged.connect(self._on_theme)
        theme_row.addWidget(self.combo_theme)
        theme_row.addStretch()
        app_lay.addLayout(theme_row)

        anim_row = QHBoxLayout()
        self.chk_animation = QCheckBox("Enable Startup Animation")
        self.chk_animation.setProperty("class", "BodyText")
        self.chk_animation.stateChanged.connect(
            lambda s: config_manager.set("startup_animation", s == Qt.CheckState.Checked.value))
        anim_row.addWidget(self.chk_animation)
        anim_row.addStretch()
        app_lay.addLayout(anim_row)
        layout.addWidget(app_card)

        # API Keys
        api_card = QWidget()
        api_card.setProperty("class", "Card")
        api_lay = QVBoxLayout(api_card)
        api_lay.setContentsMargins(20, 20, 20, 20)
        api_lay.setSpacing(15)

        lbl2 = QLabel("🔑 API Configuration")
        lbl2.setProperty("class", "CardTitle")
        api_lay.addWidget(lbl2)
        desc = QLabel("Enter your API keys for AI processing. These are stored locally and used when needed.")
        desc.setProperty("class", "MutedText")
        desc.setWordWrap(True)
        api_lay.addWidget(desc)

        self.txt_gemini = self._api_row(api_lay, "Gemini API Key:", "Enter Gemini API Key")
        self.txt_claude = self._api_row(api_lay, "Claude API Key:", "Enter Claude API Key")

        btn_save = QPushButton("💾 Save API Keys")
        btn_save.setProperty("class", "PrimaryButton")
        btn_save.clicked.connect(self._save_keys)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(btn_save)
        api_lay.addLayout(btn_row)
        layout.addWidget(api_card)
        layout.addStretch()

        self._load()

    @staticmethod
    def _api_row(parent_lay, label, placeholder):
        row = QHBoxLayout()
        lbl = QLabel(label)
        lbl.setProperty("class", "BodyText")
        lbl.setFixedWidth(100)
        row.addWidget(lbl)
        txt = QLineEdit()
        txt.setEchoMode(QLineEdit.EchoMode.Password)
        txt.setPlaceholderText(placeholder)
        row.addWidget(txt)
        parent_lay.addLayout(row)
        return txt

    def _load(self):
        self.combo_theme.blockSignals(True)
        self.combo_theme.setCurrentText(config_manager.get("theme", "Default"))
        self.combo_theme.blockSignals(False)
        self.chk_animation.blockSignals(True)
        self.chk_animation.setChecked(config_manager.get("startup_animation", True))
        self.chk_animation.blockSignals(False)
        self.txt_gemini.setText(config_manager.get("gemini_api_key", ""))
        self.txt_claude.setText(config_manager.get("claude_api_key", ""))

    def _on_theme(self, new_theme):
        config_manager.set("theme", new_theme)
        if self.main_app:
            self.main_app.apply_theme()

    def _save_keys(self):
        config_manager.set("gemini_api_key", self.txt_gemini.text().strip())
        config_manager.set("claude_api_key", self.txt_claude.text().strip())
        QMessageBox.information(self, "Settings Saved", "API Keys saved successfully.")
