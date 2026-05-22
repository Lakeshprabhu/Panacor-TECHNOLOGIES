import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QFileDialog, QMessageBox
from pages._ui import card, file_row, action_btn, log_console, BaseWorker, handle_finished
import core.legends_engine as engine


class LegendsWorker(BaseWorker):
    def __init__(self, image_path, output_path):
        super().__init__()
        self.image_path, self.output_path = image_path, output_path

    def run(self):
        try:
            self.log_signal.emit(f"Processing Legend Table: {self.image_path}...")
            results = engine.process_legend(self.image_path, None, self.output_path)
            count = len(results) if results else 0
            self.finished_signal.emit(f"Processed successfully! {count} icons extracted.", True)
        except Exception as e:
            self.finished_signal.emit(str(e), False)


class PageLegends(QWidget):
    def __init__(self):
        super().__init__()
        self.image_path = ""
        self.excel_path = ""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 20, 40, 20)
        layout.setSpacing(20)

        fc, fc_lay = card("📐 Source Image", (20, 20, 20, 20))
        row, self.btn_img, self.lbl_img = file_row("🖼️ Select Drawing (Image)", "No image selected")
        self.btn_img.clicked.connect(self._select_image)
        fc_lay.addLayout(row)
        layout.addWidget(fc)

        oc, oc_lay = card("💾 Output Options", (20, 20, 20, 20))
        row2, self.btn_out, self.lbl_out = file_row("Select Output Excel (Optional)", "(Auto-generated)")
        self.btn_out.clicked.connect(self._select_output)
        oc_lay.addLayout(row2)
        layout.addWidget(oc)

        self.btn_process = action_btn("▶ GENERATE EXCEL AND ICONS")
        self.btn_process.clicked.connect(self._run)
        layout.addWidget(self.btn_process)

        self.console = log_console()
        layout.addWidget(self.console)

    def _select_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Legend Image", "", "Images (*.png *.jpg *.jpeg *.bmp)")
        if path:
            self.image_path = path
            self.lbl_img.setText(os.path.basename(path))

    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel As", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_path = path
            self.lbl_out.setText(os.path.basename(path))

    def _run(self):
        if not self.image_path:
            QMessageBox.warning(self, "Error", "Select an image first.")
            return
        self.btn_process.setEnabled(False)
        self.console.clear()
        self.console.append("Initializing Legends Processing Engine...")
        self.worker = LegendsWorker(self.image_path, self.excel_path or None)
        self.worker.log_signal.connect(self.console.append)
        self.worker.finished_signal.connect(lambda r, s: handle_finished(self, self.btn_process, r, s))
        self.worker.start()
