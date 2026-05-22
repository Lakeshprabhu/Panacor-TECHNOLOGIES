import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QLabel, QTextEdit, QFileDialog, QMessageBox
from pages._ui import card, file_row, action_btn, log_console, BaseWorker
import core.quotation_engine as engine
from core.config import config_manager


class QuotationWorker(BaseWorker):
    def __init__(self, excel, pdfs, prompt, key, provider):
        super().__init__()
        self.excel, self.pdfs, self.prompt, self.key, self.provider = excel, pdfs, prompt, key, provider

    def run(self):
        try:
            self.log_signal.emit(f"Processing {len(self.pdfs)} PDF(s) via {self.provider}...")
            output_file = engine.process_quotation(self.excel, self.pdfs, self.prompt, self.key, self.provider)
            self.finished_signal.emit(output_file, True)
        except Exception as e:
            self.finished_signal.emit(str(e), False)


class PageQuotation(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_paths = []
        self.excel_path = ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 20, 40, 20)
        layout.setSpacing(20)

        fc, fc_lay = card("📁 Input Files", (20, 20, 20, 20))
        fc_lay.addSpacing(10)
        row_ex, self.btn_excel, self.lbl_excel = file_row("📄 Select Quotation Template (Excel)")
        self.btn_excel.clicked.connect(self._select_excel)
        fc_lay.addLayout(row_ex)
        row_pdf, self.btn_pdf, self.lbl_pdfs = file_row("📑 Select BOQ & Tech Details (PDFs)", "No files selected")
        self.btn_pdf.clicked.connect(self._select_pdfs)
        fc_lay.addLayout(row_pdf)
        layout.addWidget(fc)

        cfg, cfg_lay = card("🤖 AI Configuration", (20, 20, 20, 20))
        cfg_lay.addSpacing(10)
        lbl = QLabel("Custom Instructions (Optional):")
        lbl.setProperty("class", "BodyText")
        cfg_lay.addWidget(lbl)
        self.txt_prompt = QTextEdit()
        self.txt_prompt.setFixedHeight(60)
        cfg_lay.addWidget(self.txt_prompt)
        layout.addWidget(cfg)

        self.btn_process = action_btn("▶ PROCESS QUOTATION")
        self.btn_process.clicked.connect(self._run)
        layout.addWidget(self.btn_process)

        self.console = log_console()
        layout.addWidget(self.console)

    def _select_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_path = path
            self.lbl_excel.setText(os.path.basename(path))

    def _select_pdfs(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select PDFs", "", "PDF Files (*.pdf)")
        if paths:
            self.pdf_paths = paths
            self.lbl_pdfs.setText(f"{len(paths)} file(s) selected")

    def _run(self):
        if not self.excel_path:
            QMessageBox.warning(self, "Error", "Select an Excel template.")
            return
        if not self.pdf_paths:
            QMessageBox.warning(self, "Error", "Select at least one PDF.")
            return
        key = config_manager.get("gemini_api_key", "").strip()
        if not key:
            QMessageBox.warning(self, "Error", "API Key is required. Please set it in Settings.")
            return
        self.btn_process.setEnabled(False)
        self.console.clear()
        self.console.append("Initializing Quotation Automation Engine...")
        self.worker = QuotationWorker(
            self.excel_path, self.pdf_paths, self.txt_prompt.toPlainText().strip(), key, "Gemini")
        self.worker.log_signal.connect(self.console.append)
        self.worker.finished_signal.connect(self._on_done)
        self.worker.start()

    def _on_done(self, result, success):
        self.btn_process.setEnabled(True)
        if success:
            self.console.append(f"\n[SUCCESS] Saved to:\n{result}")
            QMessageBox.information(self, "Success", f"Quotation processed!\n{result}")
        else:
            self.console.append(f"\n[ERROR] {result}")
            QMessageBox.critical(self, "Error", result)
