#!/usr/bin/env python3
"""Point-and-click window for prep_invoices.py (no terminal needed).

Run from source:  python app.py
Build an app:     see build_app.py
"""
import contextlib
import datetime as dt
import io
import os
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import prep_invoices


def parse_date(text, label, required=False):
    """Accept YYYY-MM-DD or YYYYMMDD. Blank gives None unless required."""
    text = text.strip().replace("-", "")
    if not text:
        if required:
            raise ValueError(f"{label} is required.")
        return None
    try:
        return prep_invoices.yyyymmdd(text)
    except Exception:
        raise ValueError(f"{label} is not a valid date; use YYYY-MM-DD, e.g. 2026-08-15.")


def reveal(path):
    """Show the file in Finder / Explorer."""
    if sys.platform == "darwin":
        subprocess.run(["open", "-R", path])
    elif sys.platform == "win32":
        subprocess.run(["explorer", "/select,", os.path.normpath(path)])
    else:
        subprocess.run(["xdg-open", os.path.dirname(path)])


class App(ttk.Frame):
    def __init__(self, root):
        super().__init__(root, padding=14)
        self.pack(fill="both", expand=True)
        self.output_path = None
        self.vars = {
            "workbook": tk.StringVar(),
            "output": tk.StringVar(),
            "invoice_date": tk.StringVar(value=dt.date.today().isoformat()),
            "start": tk.StringVar(),
            "end": tk.StringVar(),
            "terms": tk.StringVar(value="30"),
            "first_no": tk.StringVar(value="1001"),
        }
        self.columnconfigure(1, weight=1)

        self.row = 0
        self.file_row("Timesheet (.xlsx)", "workbook", self.pick_workbook)
        self.file_row("Save CSV as", "output", self.pick_output)
        self.field("Invoice date", "invoice_date", "YYYY-MM-DD")
        self.field("First day to include", "start", "YYYY-MM-DD, blank = whole month")
        self.field("Last day to include", "end", "YYYY-MM-DD, blank = whole month")
        self.field("Payment terms (days)", "terms", "0 = due on receipt")
        self.field("First invoice number", "first_no", "")

        buttons = ttk.Frame(self)
        buttons.grid(row=self.row, column=0, columnspan=3, pady=(12, 6), sticky="w")
        self.run_button = ttk.Button(buttons, text="Create invoice CSV", command=self.run)
        self.run_button.pack(side="left")
        self.show_button = ttk.Button(buttons, text="Show file", command=self.show, state="disabled")
        self.show_button.pack(side="left", padx=8)
        self.row += 1

        self.log = tk.Text(self, height=10, width=70, state="disabled", wrap="word")
        self.log.grid(row=self.row, column=0, columnspan=3, sticky="nsew")
        self.rowconfigure(self.row, weight=1)
        self.log.tag_configure("error", foreground="#b00020")

    def field(self, label, key, hint):
        ttk.Label(self, text=label).grid(row=self.row, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.vars[key], width=18).grid(row=self.row, column=1, sticky="w", padx=8)
        ttk.Label(self, text=hint, foreground="gray").grid(row=self.row, column=2, sticky="w")
        self.row += 1

    def file_row(self, label, key, command):
        ttk.Label(self, text=label).grid(row=self.row, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.vars[key], width=50).grid(row=self.row, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="Browse...", command=command).grid(row=self.row, column=2)
        self.row += 1

    def pick_workbook(self):
        path = filedialog.askopenfilename(title="Choose the timesheet",
                                          filetypes=[("Excel workbook", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.vars["workbook"].set(path)
            stem = os.path.splitext(path)[0]
            self.vars["output"].set(stem + "_invoices.csv")

    def pick_output(self):
        path = filedialog.asksaveasfilename(title="Save invoice CSV", defaultextension=".csv",
                                            initialfile=os.path.basename(self.vars["output"].get() or "invoices.csv"),
                                            filetypes=[("CSV", "*.csv")])
        if path:
            self.vars["output"].set(path)

    def write(self, text, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip() + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def run(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.show_button.configure(state="disabled")
        v = {k: var.get().strip() for k, var in self.vars.items()}
        try:
            if not v["workbook"]:
                raise ValueError("Choose the timesheet first.")
            if not os.path.isfile(v["workbook"]):
                raise ValueError("That timesheet file was not found.")
            if not v["output"]:
                raise ValueError("Choose where to save the CSV.")
            invoice_date = parse_date(v["invoice_date"], "Invoice date", required=True)
            start, end = parse_date(v["start"], "First day"), parse_date(v["end"], "Last day")
            if start and end and start > end:
                raise ValueError("The first day is after the last day.")
            terms, first_no = int(v["terms"]), int(v["first_no"])
        except ValueError as e:
            msg = str(e)
            if msg.startswith("invalid literal"):
                msg = "Payment terms and invoice number must be whole numbers."
            self.write(msg, "error")
            return

        # The converter reports progress on stderr and stops with sys.exit("message") on bad data.
        captured = io.StringIO()
        try:
            with contextlib.redirect_stderr(captured):
                prep_invoices.write_invoices(v["workbook"], v["output"], invoice_date, terms, first_no, start, end)
        except SystemExit as e:
            self.write(captured.getvalue())
            self.write(f"Stopped: {e.code}", "error")
            return
        except Exception as e:
            self.write(f"Could not create the CSV: {e}", "error")
            return
        self.write(captured.getvalue())
        self.output_path = v["output"]
        self.show_button.configure(state="normal")

    def show(self):
        if self.output_path and os.path.exists(self.output_path):
            reveal(self.output_path)
        else:
            messagebox.showinfo("Not found", "The CSV file is no longer there.")


def main():
    root = tk.Tk()
    root.title("Prep Invoices")
    root.minsize(640, 400)
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
