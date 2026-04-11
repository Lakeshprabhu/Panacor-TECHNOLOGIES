import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Import logic from the main script
from main import process_legend, TESSERACT_CMD, _SCRIPT_DIR

class RedirectStdout:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.original_stdout = sys.stdout

    def write(self, string):
        # We use root.after to safely insert text into the Tkinter widget from a thread
        self.text_widget.after(0, self._insert, string)

    def _insert(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


class LegendProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Legends Template Maker")
        self.root.geometry("650x520")
        
        self.create_widgets()
        
    def create_widgets(self):
        # Frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Image selection
        ttk.Label(main_frame, text="Image Path:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.image_path_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.image_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_image).grid(row=0, column=2, pady=5)
        
        # Tesseract path
        ttk.Label(main_frame, text="Tesseract Path:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.tesseract_path_var = tk.StringVar(value=TESSERACT_CMD)
        ttk.Entry(main_frame, textvariable=self.tesseract_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_tesseract).grid(row=1, column=2, pady=5)

        # Excel output path
        ttk.Label(main_frame, text="Excel Output:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.excel_path_var = tk.StringVar(value="(auto-generated)")
        ttk.Entry(main_frame, textvariable=self.excel_path_var, width=50).grid(row=2, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Browse...", command=self.browse_excel_output).grid(row=2, column=2, pady=5)
        
        # Log Text Area
        ttk.Label(main_frame, text="Logs:").grid(row=3, column=0, sticky=tk.NW, pady=5)
        self.log_text = tk.Text(main_frame, width=50, height=15, state=tk.NORMAL)
        self.log_text.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # Make the log text box resizable
        main_frame.rowconfigure(3, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Scrollbar for log
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=3, column=3, sticky=(tk.N, tk.S), pady=5)
        self.log_text['yscrollcommand'] = scrollbar.set
        
        # Redirect stdout
        sys.stdout = RedirectStdout(self.log_text)
        
        # Default text
        print("Welcome to Legends Template Maker GUI.")
        print("Select an image and click Process.")
        print("Icons will be saved to 'output_icons/' folder.")
        print("An Excel file with icons + names will also be generated.")
        
        # Process button
        self.process_button = ttk.Button(main_frame, text="Process Image", command=self.start_processing)
        self.process_button.grid(row=4, column=1, pady=10)
        
    def browse_image(self):
        filepath = filedialog.askopenfilename(
            title="Select Image File",
            filetypes=[("Image Files", "*.jpeg *.jpg *.png *.bmp"), ("All Files", "*.*")]
        )
        if filepath:
            self.image_path_var.set(filepath)
            # Auto-generate Excel output path based on image name
            base_name = os.path.splitext(os.path.basename(filepath))[0]
            excel_path = os.path.join(_SCRIPT_DIR, f"{base_name}_legend_icons.xlsx")
            self.excel_path_var.set(excel_path)
            
    def browse_tesseract(self):
        filepath = filedialog.askopenfilename(
            title="Select Tesseract Executable",
            filetypes=[("Executable Files", "*.exe"), ("All Files", "*.*")]
        )
        if filepath:
            self.tesseract_path_var.set(filepath)

    def browse_excel_output(self):
        filepath = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if filepath:
            self.excel_path_var.set(filepath)
            
    def start_processing(self):
        image_path = self.image_path_var.get().strip()
        tess_path = self.tesseract_path_var.get().strip()
        excel_path = self.excel_path_var.get().strip()
        
        if not image_path:
            messagebox.showerror("Error", "Please select an image first.")
            return

        # If excel path is auto-generated placeholder, let main.py handle it
        if excel_path == "(auto-generated)" or not excel_path:
            excel_path = None
            
        # Disable button to prevent multiple clicks
        self.process_button.state(['disabled'])
        self.log_text.insert(tk.END, "-"*40 + "\n")
        
        # Run in separate thread
        thread = threading.Thread(target=self.run_process, args=(image_path, tess_path, excel_path))
        thread.daemon = True
        thread.start()
        
    def run_process(self, image_path, tess_path, excel_path):
        try:
            print(f"Starting process for: {image_path}")
            results = process_legend(image_path, tess_path, excel_path)
            count = len(results) if results else 0
            print(f"Processing completed successfully. {count} icon(s) extracted.")
            self.root.after(0, lambda: messagebox.showinfo(
                "Success",
                f"Image processed successfully!\n\n"
                f"• {count} icon(s) saved to output_icons/\n"
                f"• Excel file generated with icons and names"
            ))
        except Exception as e:
            print(f"[CRITICAL ERROR] {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"An error occurred:\n{e}"))
        finally:
            # Re-enable the button
            self.root.after(0, lambda: self.process_button.state(['!disabled']))


if __name__ == "__main__":
    root = tk.Tk()
    app = LegendProcessorGUI(root)
    # Ensure stdout is restored before Tkinter exits, to avoid issues
    original_stdout = sys.stdout
    def on_closing():
        sys.stdout = original_stdout
        root.destroy()
        sys.exit(0)
        
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()
