#!/usr/bin/env python3
"""
O&M Manual Preparation — GUI
==============================
Tkinter-based GUI for building O&M manuals.
Supports 5 input methods: PDF, Excel, Image, Text file, Manual entry.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from om_processor import (
    parse_items_from_pdf,
    parse_items_from_excel,
    parse_items_from_image,
    parse_items_from_text,
    parse_items_from_string,
    process_om_manual,
    _SCRIPT_DIR,
)


class RedirectStdout:
    """Thread-safe stdout → Tkinter Text widget redirect."""

    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.original_stdout = sys.stdout

    def write(self, string):
        self.text_widget.after(0, self._insert, string)

    def _insert(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


class OMManualGUI:
    INPUT_TYPES = ["PDF File", "Excel File", "Image (OCR)", "Text File", "Manual Entry"]
    FILE_FILTERS = {
        "PDF File":   [("PDF Files", "*.pdf"), ("All Files", "*.*")],
        "Excel File": [("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        "Image (OCR)":[("Image Files", "*.png *.jpg *.jpeg *.bmp *.tiff"), ("All Files", "*.*")],
        "Text File":  [("Text Files", "*.txt *.csv"), ("All Files", "*.*")],
    }

    def __init__(self, root):
        self.root = root
        self.root.title("O&M Manual Preparation")
        self.root.geometry("780x700")
        self.root.minsize(700, 620)

        self.items: list[str] = []
        self.create_widgets()

    # ======================================================================
    # Widget construction
    # ======================================================================

    def create_widgets(self):
        main = ttk.Frame(self.root, padding="10")
        main.pack(fill=tk.BOTH, expand=True)

        # ── Row 0: Input type selector ─────────────────────────────────
        ttk.Label(main, text="Input Type:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_type_var = tk.StringVar(value=self.INPUT_TYPES[0])
        type_combo = ttk.Combobox(
            main, textvariable=self.input_type_var,
            values=self.INPUT_TYPES, state="readonly", width=20,
        )
        type_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        type_combo.bind("<<ComboboxSelected>>", self._on_type_change)

        # ── Row 1: File path / manual entry ────────────────────────────
        self.file_frame = ttk.Frame(main)
        self.file_frame.grid(row=1, column=0, columnspan=4, sticky=tk.EW, pady=2)

        ttk.Label(self.file_frame, text="File:").pack(side=tk.LEFT)
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=55)
        self.file_entry.pack(side=tk.LEFT, padx=5)
        self.browse_btn = ttk.Button(self.file_frame, text="Browse…", command=self.browse_file)
        self.browse_btn.pack(side=tk.LEFT)

        # Manual entry text area (hidden by default)
        self.manual_frame = ttk.Frame(main)
        ttk.Label(self.manual_frame, text="Enter items (one per line):").pack(anchor=tk.W)
        self.manual_text = tk.Text(self.manual_frame, width=60, height=6)
        self.manual_text.pack(fill=tk.X, expand=True)

        # ── Row 2: Parse + action buttons ──────────────────────────────
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=2, column=0, columnspan=4, pady=8)

        self.parse_btn = ttk.Button(btn_frame, text="Parse Items", command=self.parse_items)
        self.parse_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = ttk.Button(btn_frame, text="Clear List", command=self.clear_items)
        self.clear_btn.pack(side=tk.LEFT, padx=5)

        self.add_btn = ttk.Button(btn_frame, text="Add Item…", command=self.add_item_dialog)
        self.add_btn.pack(side=tk.LEFT, padx=5)

        self.remove_btn = ttk.Button(btn_frame, text="Remove Selected", command=self.remove_selected)
        self.remove_btn.pack(side=tk.LEFT, padx=5)

        # ── Row 3: Item listbox ────────────────────────────────────────
        ttk.Label(main, text="Items to process:").grid(row=3, column=0, sticky=tk.NW, pady=2)

        list_frame = ttk.Frame(main)
        list_frame.grid(row=3, column=1, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=2)

        self.item_listbox = tk.Listbox(list_frame, width=60, height=8, selectmode=tk.EXTENDED)
        self.item_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        list_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.item_listbox.yview)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.item_listbox["yscrollcommand"] = list_scroll.set

        # ── Row 4: Output path ─────────────────────────────────────────
        ttk.Label(main, text="Output PDF:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.output_path_var = tk.StringVar(value="(auto-generated)")
        ttk.Entry(main, textvariable=self.output_path_var, width=55).grid(
            row=4, column=1, columnspan=2, padx=5, pady=5, sticky=tk.EW
        )
        ttk.Button(main, text="Browse…", command=self.browse_output).grid(
            row=4, column=3, pady=5
        )

        # ── Row 5: Process button ──────────────────────────────────────
        self.process_btn = ttk.Button(
            main, text="▶  Search, Download & Merge", command=self.start_processing
        )
        self.process_btn.grid(row=5, column=1, columnspan=2, pady=8)

        # ── Row 6: Progress bar ────────────────────────────────────────
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            main, variable=self.progress_var, maximum=100, length=400
        )
        self.progress_bar.grid(row=6, column=1, columnspan=2, pady=2, sticky=tk.EW)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var).grid(
            row=7, column=1, columnspan=2, sticky=tk.W
        )

        # ── Row 8: Log area ────────────────────────────────────────────
        ttk.Label(main, text="Logs:").grid(row=8, column=0, sticky=tk.NW, pady=5)
        self.log_text = tk.Text(main, width=60, height=12, state=tk.NORMAL)
        self.log_text.grid(
            row=8, column=1, columnspan=3,
            sticky=(tk.W, tk.E, tk.N, tk.S), pady=5,
        )

        log_scroll = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.grid(row=8, column=4, sticky=(tk.N, tk.S), pady=5)
        self.log_text["yscrollcommand"] = log_scroll.set

        main.rowconfigure(8, weight=1)
        main.columnconfigure(1, weight=1)

        # Redirect stdout
        sys.stdout = RedirectStdout(self.log_text)

        print("Welcome to O&M Manual Preparation Tool.")
        print("1. Select an input type and load your item list.")
        print("2. Review / edit the parsed items.")
        print("3. Click 'Search, Download & Merge' to build the manual.")
        print()

    # ======================================================================
    # Input-type switching
    # ======================================================================

    def _on_type_change(self, event=None):
        selected = self.input_type_var.get()
        if selected == "Manual Entry":
            self.file_frame.grid_remove()
            self.manual_frame.grid(row=1, column=0, columnspan=4, sticky=tk.EW, pady=2)
        else:
            self.manual_frame.grid_remove()
            self.file_frame.grid(row=1, column=0, columnspan=4, sticky=tk.EW, pady=2)

    # ======================================================================
    # File browsers
    # ======================================================================

    def browse_file(self):
        input_type = self.input_type_var.get()
        filters = self.FILE_FILTERS.get(input_type, [("All Files", "*.*")])
        filepath = filedialog.askopenfilename(title=f"Select {input_type}", filetypes=filters)
        if filepath:
            self.file_path_var.set(filepath)

    def browse_output(self):
        filepath = filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        if filepath:
            self.output_path_var.set(filepath)

    # ======================================================================
    # Item list management
    # ======================================================================

    def parse_items(self):
        input_type = self.input_type_var.get()

        try:
            if input_type == "Manual Entry":
                text = self.manual_text.get("1.0", tk.END)
                new_items = parse_items_from_string(text)
            else:
                filepath = self.file_path_var.get().strip()
                if not filepath:
                    messagebox.showerror("Error", "Please select a file first.")
                    return

                parsers = {
                    "PDF File": parse_items_from_pdf,
                    "Excel File": parse_items_from_excel,
                    "Image (OCR)": parse_items_from_image,
                    "Text File": parse_items_from_text,
                }
                parser = parsers[input_type]
                print(f"Parsing {input_type}: {filepath}")
                new_items = parser(filepath)

            # Append to existing list (de-dup)
            existing_lower = {i.lower() for i in self.items}
            added = 0
            for item in new_items:
                if item.lower() not in existing_lower:
                    self.items.append(item)
                    existing_lower.add(item.lower())
                    added += 1

            self._refresh_listbox()
            print(f"Parsed {len(new_items)} item(s), added {added} new. Total: {len(self.items)}")

        except Exception as e:
            messagebox.showerror("Parse Error", str(e))
            print(f"[ERROR] {e}")

    def clear_items(self):
        self.items.clear()
        self._refresh_listbox()
        print("Item list cleared.")

    def add_item_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Item")
        dialog.geometry("400x120")
        dialog.transient(self.root)
        dialog.grab_set()

        ttk.Label(dialog, text="Item name:").pack(pady=(10, 2))
        entry_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=entry_var, width=50)
        entry.pack(padx=10)
        entry.focus()

        def do_add():
            name = entry_var.get().strip()
            if name:
                self.items.append(name)
                self._refresh_listbox()
                print(f"Added: {name}")
            dialog.destroy()

        ttk.Button(dialog, text="Add", command=do_add).pack(pady=10)
        entry.bind("<Return>", lambda e: do_add())

    def remove_selected(self):
        selected = self.item_listbox.curselection()
        if not selected:
            return
        # Remove in reverse order to keep indices stable
        for idx in reversed(selected):
            removed = self.items.pop(idx)
            print(f"Removed: {removed}")
        self._refresh_listbox()

    def _refresh_listbox(self):
        self.item_listbox.delete(0, tk.END)
        for i, item in enumerate(self.items, 1):
            self.item_listbox.insert(tk.END, f"{i}. {item}")

    # ======================================================================
    # Processing
    # ======================================================================

    def start_processing(self):
        if not self.items:
            messagebox.showerror("Error", "No items to process. Parse or add items first.")
            return

        output_path = self.output_path_var.get().strip()
        if output_path == "(auto-generated)" or not output_path:
            output_path = None

        self.process_btn.state(["disabled"])
        self.parse_btn.state(["disabled"])
        self.progress_var.set(0)
        self.log_text.insert(tk.END, "\n" + "=" * 60 + "\n")

        thread = threading.Thread(
            target=self._run_process,
            args=(list(self.items), output_path),
        )
        thread.daemon = True
        thread.start()

    def _progress_callback(self, current, total, item_name, status):
        """Called from the worker thread to update the progress bar."""
        pct = ((current + 1) / total) * 100 if total > 0 else 0
        self.root.after(0, self.progress_var.set, pct)

        status_text = {
            "searching": f"Searching: {item_name}",
            "downloading": f"Downloading: {item_name}",
            "cached": f"Cached: {item_name}",
            "done": f"Done: {item_name}",
            "not_found": f"Not found: {item_name}",
        }.get(status, item_name)

        self.root.after(0, self.status_var.set, f"[{current+1}/{total}] {status_text}")

    def _run_process(self, items, output_path):
        try:
            result_path = process_om_manual(
                items, output_path,
                progress_callback=self._progress_callback,
            )

            if result_path:
                self.root.after(0, self.status_var.set, "Complete!")
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Success",
                        f"O&M Manual created successfully!\n\n"
                        f"Saved to:\n{result_path}",
                    ),
                )
            else:
                self.root.after(0, self.status_var.set, "No datasheets found.")
                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "Warning",
                        "No datasheets could be downloaded.\n"
                        "Check the log for details.",
                    ),
                )
        except Exception as e:
            print(f"[CRITICAL ERROR] {e}")
            self.root.after(
                0,
                lambda: messagebox.showerror("Error", f"An error occurred:\n{e}"),
            )
        finally:
            self.root.after(0, lambda: self.process_btn.state(["!disabled"]))
            self.root.after(0, lambda: self.parse_btn.state(["!disabled"]))


# ======================================================================
# Entry point
# ======================================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = OMManualGUI(root)

    original_stdout = sys.stdout

    def on_closing():
        sys.stdout = original_stdout
        root.destroy()
        sys.exit(0)

    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
