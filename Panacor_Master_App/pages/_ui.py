from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTextEdit, QMessageBox
)
from PyQt6.QtCore import QThread, pyqtSignal


def card(title=None, margins=None):
    w = QWidget()
    w.setProperty("class", "Card")
    lay = QVBoxLayout(w)
    if margins:
        lay.setContentsMargins(*margins)
    if title:
        lbl = QLabel(title)
        lbl.setProperty("class", "CardTitle")
        lay.addWidget(lbl)
    return w, lay


def file_row(btn_text, default_lbl="No file selected"):
    row = QHBoxLayout()
    btn = QPushButton(btn_text)
    btn.setProperty("class", "SecondaryButton")
    lbl = QLabel(default_lbl)
    lbl.setProperty("class", "MutedText")
    row.addWidget(btn)
    row.addWidget(lbl)
    row.addStretch()
    return row, btn, lbl


def action_btn(text):
    btn = QPushButton(text)
    btn.setProperty("class", "PrimaryButton")
    return btn


def log_console():
    c = QTextEdit()
    c.setProperty("class", "TerminalConsole")
    c.setReadOnly(True)
    return c


class BaseWorker(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str, bool)


def handle_finished(page, btn, result, success):
    btn.setEnabled(True)
    if success:
        page.console.append(f"\n[SUCCESS] {result}")
        QMessageBox.information(page, "Success", result)
    else:
        page.console.append(f"\n[ERROR] {result}")
        QMessageBox.critical(page, "Error", result)
