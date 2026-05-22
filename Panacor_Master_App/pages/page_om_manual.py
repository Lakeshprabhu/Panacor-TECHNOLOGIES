import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QTextEdit, QMessageBox, QComboBox, QListWidget, QProgressBar
)
from PyQt6.QtCore import pyqtSignal
from pages._ui import card, file_row, action_btn, log_console, BaseWorker
import core.om_manual_engine as engine

INPUT_TYPES = ["PDF File", "Excel File", "Image (OCR)", "Text File", "Manual Entry"]
FILE_FILTERS = {
    "PDF File": "PDF Files (*.pdf)", "Excel File": "Excel Files (*.xlsx *.xls)",
    "Image (OCR)": "Images (*.png *.jpg *.jpeg *.bmp)", "Text File": "Text Files (*.txt *.csv)",
}
PARSERS = {
    "PDF File": engine.parse_items_from_pdf, "Excel File": engine.parse_items_from_excel,
    "Image (OCR)": engine.parse_items_from_image, "Text File": engine.parse_items_from_text,
}


class OMWorker(BaseWorker):
    progress_signal = pyqtSignal(int, int, str, str)

    def __init__(self, items, output_path):
        super().__init__()
        self.items, self.output_path = items, output_path

    def _progress(self, current, total, name, status):
        labels = {"searching": "Searching", "downloading": "Downloading",
                  "cached": "Cached", "done": "Done", "not_found": "Not found"}
        self.log_signal.emit(f"[{current+1}/{total}] {labels.get(status, status)}: {name}")
        self.progress_signal.emit(current, total, name, status)

    def run(self):
        try:
            path = engine.process_om_manual(self.items, self.output_path, progress_callback=self._progress)
            if path:
                self.finished_signal.emit(f"O&M Manual created successfully!\n\nSaved to: {path}", True)
            else:
                self.finished_signal.emit("No datasheets found. Check log for details.", False)
        except Exception as e:
            self.finished_signal.emit(str(e), False)


class PageOMManual(QWidget):
    def __init__(self):
        super().__init__()
        self.items, self.input_file, self.output_file = [], "", ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 20, 40, 20)
        layout.setSpacing(15)

        # Input section
        top, top_lay = card(margins=(15, 15, 15, 15))
        row = QHBoxLayout()
        row.addWidget(QLabel("Input Type:"))
        self.combo_type = QComboBox()
        self.combo_type.addItems(INPUT_TYPES)
        self.combo_type.currentIndexChanged.connect(self._on_type_change)
        row.addWidget(self.combo_type)
        self.btn_browse = QPushButton("Browse File...")
        self.btn_browse.setProperty("class", "SecondaryButton")
        self.btn_browse.clicked.connect(self._browse)
        row.addWidget(self.btn_browse)
        self.lbl_file = QLabel("No file selected")
        self.lbl_file.setProperty("class", "MutedText")
        row.addWidget(self.lbl_file)
        row.addStretch()
        top_lay.addLayout(row)

        self.txt_manual = QTextEdit()
        self.txt_manual.setPlaceholderText("Enter items, one per line...")
        self.txt_manual.setFixedHeight(80)
        self.txt_manual.setVisible(False)
        top_lay.addWidget(self.txt_manual)

        btn_parse = QPushButton("Parse Items")
        btn_parse.setProperty("class", "SecondaryButton")
        btn_parse.clicked.connect(self._parse)
        top_lay.addWidget(btn_parse)
        layout.addWidget(top)

        # Items list
        mid, mid_lay = card(margins=(15, 15, 15, 15))
        hdr = QHBoxLayout()
        lbl_list = QLabel("Parsed Items List")
        lbl_list.setProperty("class", "CardTitle")
        hdr.addWidget(lbl_list)
        hdr.addStretch()
        btn_clear = QPushButton("Clear All")
        btn_clear.setProperty("class", "DangerButton")
        btn_clear.clicked.connect(self._clear)
        hdr.addWidget(btn_clear)
        btn_add = QPushButton("Add Item")
        btn_add.setProperty("class", "SecondaryButton")
        hdr.addWidget(btn_add)
        mid_lay.addLayout(hdr)
        self.list_widget = QListWidget()
        self.list_widget.setFixedHeight(120)
        mid_lay.addWidget(self.list_widget)
        layout.addWidget(mid)

        # Run section
        run_w, run_lay = card()
        out_row, self.btn_out, self.lbl_out = file_row("Select Output PDF (Optional)", "(Auto-generated in Master App Folder)")
        self.btn_out.clicked.connect(self._select_output)
        run_lay.addLayout(out_row)
        self.btn_process = action_btn("▶ SEARCH, DOWNLOAD & MERGE")
        self.btn_process.clicked.connect(self._run)
        run_lay.addWidget(self.btn_process)
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        run_lay.addWidget(self.progress_bar)
        layout.addWidget(run_w)

        self.console = log_console()
        layout.addWidget(self.console)

    def _on_type_change(self):
        manual = self.combo_type.currentText() == "Manual Entry"
        self.btn_browse.setVisible(not manual)
        self.lbl_file.setVisible(not manual)
        self.txt_manual.setVisible(manual)

    def _browse(self):
        t = self.combo_type.currentText()
        path, _ = QFileDialog.getOpenFileName(self, f"Select {t}", "", FILE_FILTERS.get(t, "All Files (*.*)"))
        if path:
            self.input_file = path
            self.lbl_file.setText(os.path.basename(path))

    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Manual As", "", "PDF Files (*.pdf)")
        if path:
            self.output_file = path
            self.lbl_out.setText(os.path.basename(path))

    def _parse(self):
        sel = self.combo_type.currentText()
        try:
            if sel == "Manual Entry":
                new_items = engine.parse_items_from_string(self.txt_manual.toPlainText())
            else:
                if not self.input_file:
                    QMessageBox.warning(self, "Error", "Select a file first.")
                    return
                new_items = PARSERS[sel](self.input_file)
            existing = {i.lower() for i in self.items}
            for i in new_items:
                if i.lower() not in existing:
                    self.items.append(i)
                    existing.add(i.lower())
            self._refresh_list()
            self.console.append(f"Parsed and added items. Total: {len(self.items)}")
        except Exception as e:
            QMessageBox.critical(self, "Parse Error", str(e))

    def _refresh_list(self):
        self.list_widget.clear()
        for i, item in enumerate(self.items, 1):
            self.list_widget.addItem(f"{i}. {item}")

    def _clear(self):
        self.items.clear()
        self._refresh_list()

    def _run(self):
        if not self.items:
            QMessageBox.warning(self, "Error", "Please parse or add items first.")
            return
        self.btn_process.setEnabled(False)
        self.console.clear()
        self.progress_bar.setValue(0)
        self.console.append("Starting O&M Manual Search, Download, and Merge Phase...")
        self.worker = OMWorker(list(self.items), self.output_file or None)
        self.worker.log_signal.connect(self.console.append)
        self.worker.progress_signal.connect(
            lambda c, t, n, s: self.progress_bar.setValue(int(((c + 1) / t) * 100) if t > 0 else 0))
        self.worker.finished_signal.connect(self._on_done)
        self.worker.start()

    def _on_done(self, result, success):
        self.btn_process.setEnabled(True)
        self.progress_bar.setValue(100 if success else 0)
        if success:
            QMessageBox.information(self, "Success", result)
        else:
            QMessageBox.warning(self, "Process Issue", result)
