import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QFileDialog, QMessageBox
from pages._ui import card, file_row, action_btn, log_console, BaseWorker, handle_finished
import core.pdf_excel_engine as engine


class PdfExcelWorker(BaseWorker):
    def __init__(self, pdf_path, excel_path, extra_cols):
        super().__init__()
        self.pdf_path, self.excel_path, self.extra_cols = pdf_path, excel_path, extra_cols

    def run(self):
        try:
            self.log_signal.emit(f"Starting extraction for: {os.path.basename(self.pdf_path)}")
            lines = engine.process_pdf(self.pdf_path, self.excel_path, self.extra_cols)
            count = len(lines) if lines else 0
            bold_count = sum(1 for l in lines if l.get("is_bold")) if lines else 0
            self.finished_signal.emit(
                f"Processed successfully!\n• {count} line(s) extracted\n• {bold_count} heading(s) detected.", True)
        except Exception as e:
            self.finished_signal.emit(str(e), False)


class PagePdf2Excel(QWidget):
    def __init__(self):
        super().__init__()
        self.pdf_path = ""
        self.excel_path = ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 20, 40, 20)
        layout.setSpacing(20)

        fc, fc_lay = card("📄 Input Source", (20, 20, 20, 20))
        row, self.btn_pdf, self.lbl_pdf = file_row("Select PDF File", "No PDF selected")
        self.btn_pdf.clicked.connect(self._select_pdf)
        fc_lay.addLayout(row)
        layout.addWidget(fc)

        cfg, cfg_lay = card("🛠️ Extraction Configuration", (20, 20, 20, 20))
        cfg_row = QHBoxLayout()
        cfg_row.addWidget(QLabel("Extra Columns (comma-sep):"))
        self.txt_cols = QLineEdit()
        self.txt_cols.setPlaceholderText("e.g. Unit, Qty, Rate, Amount")
        cfg_row.addWidget(self.txt_cols)
        cfg_lay.addLayout(cfg_row)
        out_row, self.btn_out, self.lbl_out = file_row("Select Output Excel", "(Auto-generated)")
        self.btn_out.clicked.connect(self._select_output)
        cfg_lay.addLayout(out_row)
        layout.addWidget(cfg)

        self.btn_process = action_btn("▶ EXTRACT TO EXCEL")
        self.btn_process.clicked.connect(self._run)
        layout.addWidget(self.btn_process)

        self.console = log_console()
        layout.addWidget(self.console)

    def _select_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF Files (*.pdf)")
        if path:
            self.pdf_path = path
            self.lbl_pdf.setText(os.path.basename(path))

    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel As", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_path = path
            self.lbl_out.setText(os.path.basename(path))

    def _run(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "Error", "Select a PDF first.")
            return
        self.btn_process.setEnabled(False)
        self.console.clear()
        self.console.append("Initializing PDF Extraction Engine...")
        cols_text = self.txt_cols.text().strip()
        extra = [c.strip() for c in cols_text.split(",")] if cols_text else []
        self.worker = PdfExcelWorker(self.pdf_path, self.excel_path or None, extra)
        self.worker.log_signal.connect(self.console.append)
        self.worker.finished_signal.connect(lambda r, s: handle_finished(self, self.btn_process, r, s))
        self.worker.start()
