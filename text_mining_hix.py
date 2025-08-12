"""
HiX Complication Miner – Improved Tkinter App (Python)

Key improvements over your original script:
- Robust tokenization (regex, lowercasing, punctuation/whitespace handling, NaN-safe)
- Exact whole-word matching (case-insensitive) with optional substring matching toggle
- Dynamic Treeview columns to match the result DataFrame
- Non-blocking UI feel with a progress/status label (still single-threaded for simplicity)
- Sheet name dropdown auto-populated after selecting a file
- Safer image handling for word cloud via Pillow (ImageTk)
- Clear function names (avoid name shadowing like `save_button` function vs widget)
- Better error messages and input validation
- Extra Dutch stopwords and configurable list

Dependencies: pandas, openpyxl (for .xlsx), wordcloud, pillow, matplotlib
"""

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import re
from io import BytesIO
from wordcloud import WordCloud
import matplotlib.pyplot as plt
from PIL import Image, ImageTk

# -----------------------------
# Configuration / Defaults
# -----------------------------
DEFAULT_SHEET_GUESS = "VPK Rapportage"
DEFAULT_KEYWORDS = ["bradycard", "onrust", "apneu", "pijn", "hoofdpijn"]
DEFAULT_PATIENT_ID_COL = "patient_id"
DEFAULT_TEXT_COL = "Report"

DUTCH_STOPWORDS = set(
    [
        "de", "het", "en", "een", "in", "van", "met", "op", "te", "dat", "die",
        "is", "was", "bij", "als", "maar", "ook", "niet", "wel", "om", "voor",
        "naar", "uit", "aan", "door", "tot", "over", "onder", "hij", "zij", "ze",
        "hun", "zijn", "haar", "we", "wij", "jij", "je", "u", "ik"
    ]
)

# -----------------------------
# Text Processing
# -----------------------------

def tokenize(text: str):
    """Return a list of lowercase word tokens using regex; NaN-safe."""
    if not isinstance(text, str):
        return []
    # Find sequences of letters/numbers (handles accents due to \w with UNICODE in Python3)
    tokens = re.findall(r"\w+", text.lower(), flags=re.UNICODE)
    return tokens


def extract_keywords_from_text(text: str, medical_keywords, use_substring=False):
    """Extract keywords present in `text`.

    If use_substring is False (default): only whole-word matches are considered.
    If use_substring is True: keywords are matched as substrings of tokens.
    """
    tokens = tokenize(text)
    if not tokens:
        return []

    # Precompute for quick membership
    token_set = set(tokens)
    kws = [kw.strip().lower() for kw in medical_keywords if kw and kw.strip()]

    if not use_substring:
        # whole-word presence only
        return [kw for kw in kws if kw in token_set]
    else:
        # substring presence: any token containing kw
        matched = []
        for kw in kws:
            for t in token_set:
                if kw in t:
                    matched.append(kw)
                    break
        return matched


def process_sheet(file_path, sheet_name, medical_keywords, patient_id_column, text_column, use_substring=False):
    # Load data
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        raise ValueError(f"Failed to read sheet '{sheet_name}' from file. Details: {e}")

    missing = [c for c in [patient_id_column, text_column] if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}.\nAvailable columns: {', '.join(df.columns.astype(str))}")

    kws = [kw.strip() for kw in medical_keywords if kw and kw.strip()]
    # Build results rows
    rows = []
    all_words = []

    for _, r in df.iterrows():
        pid = r[patient_id_column]
        text = r[text_column]
        text = text if isinstance(text, str) else ""

        matched = extract_keywords_from_text(text, kws, use_substring=use_substring)

        row = {patient_id_column: pid}
        for kw in kws:
            row[kw] = int(kw in matched)
        rows.append(row)

        # Words for wordcloud (stopwords removed)
        tokens = [t for t in tokenize(text) if t not in DUTCH_STOPWORDS]
        all_words.extend(tokens)

    results = pd.DataFrame(rows)
    if results.empty:
        # Ensure at least the patient_id column exists
        results = pd.DataFrame(columns=[patient_id_column] + kws)

    return results, all_words


def build_wordcloud_image(all_words):
    text = " ".join(all_words)
    if not text.strip():
        raise ValueError("No words available to generate a word cloud.")

    wc = WordCloud(width=800, height=800, background_color="white", colormap="Dark2", max_words=500)
    image = wc.generate(text).to_image()
    return image


# -----------------------------
# GUI App
# -----------------------------

class HiXComplicationMinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("HiX Free-Text Complication Miner")

        self._build_ui()
        self._set_defaults()
        self.data_results = None  # pandas DataFrame
        self.wordcloud_imgtk = None  # keep a reference

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # File row
        ttk.Label(main, text="Excel file:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.ent_file = ttk.Entry(main, width=50)
        self.ent_file.grid(row=0, column=1, sticky="we", padx=4, pady=4)
        ttk.Button(main, text="Browse…", command=self.on_browse_file).grid(row=0, column=2, padx=4, pady=4)

        # Sheet row
        ttk.Label(main, text="Sheet:").grid(row=1, column=0, sticky="w", padx=4, pady=4)
        self.cmb_sheet = ttk.Combobox(main, values=[], state="readonly", width=47)
        self.cmb_sheet.grid(row=1, column=1, sticky="we", padx=4, pady=4)
        self.cmb_sheet.set(DEFAULT_SHEET_GUESS)
        ttk.Button(main, text="Reload sheets", command=self.on_reload_sheets).grid(row=1, column=2, padx=4, pady=4)

        # Columns row
        ttk.Label(main, text="Patient ID column:").grid(row=2, column=0, sticky="w", padx=4, pady=4)
        self.ent_pid = ttk.Entry(main, width=50)
        self.ent_pid.grid(row=2, column=1, sticky="we", padx=4, pady=4)

        ttk.Label(main, text="Text column:").grid(row=3, column=0, sticky="w", padx=4, pady=4)
        self.ent_textcol = ttk.Entry(main, width=50)
        self.ent_textcol.grid(row=3, column=1, sticky="we", padx=4, pady=4)

        # Keywords row
        ttk.Label(main, text="Keywords (comma-separated):").grid(row=4, column=0, sticky="w", padx=4, pady=4)
        self.ent_keywords = ttk.Entry(main, width=50)
        self.ent_keywords.grid(row=4, column=1, sticky="we", padx=4, pady=4)
        self.var_substring = tk.BooleanVar(value=False)
        ttk.Checkbutton(main, text="Allow substring matches (e.g., 'pijn' in 'pijnstilling')", variable=self.var_substring).grid(row=4, column=2, sticky="w", padx=4, pady=4)

        # Output row
        ttk.Label(main, text="Output .xlsx:").grid(row=5, column=0, sticky="w", padx=4, pady=4)
        self.ent_out = ttk.Entry(main, width=50)
        self.ent_out.grid(row=5, column=1, sticky="we", padx=4, pady=4)
        ttk.Button(main, text="Choose…", command=self.on_choose_output).grid(row=5, column=2, padx=4, pady=4)

        # Action buttons
        ttk.Button(main, text="Extract Keywords", command=self.on_extract).grid(row=6, column=0, padx=4, pady=8, sticky="we")
        ttk.Button(main, text="Generate Word Cloud", command=self.on_wordcloud).grid(row=6, column=1, padx=4, pady=8, sticky="we")

        # Status
        self.lbl_status = ttk.Label(main, text="Ready", foreground="#444")
        self.lbl_status.grid(row=6, column=2, sticky="e")

        # Results table (Treeview)
        self.tree = ttk.Treeview(main, columns=("placeholder",), show="headings", height=12)
        self.tree.grid(row=7, column=0, columnspan=3, sticky="nsew", padx=4, pady=8)
        self.scroll_y = ttk.Scrollbar(main, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.scroll_y.set)
        self.scroll_y.grid(row=7, column=3, sticky="ns")

        # Word cloud image
        self.lbl_wc = ttk.Label(main)
        self.lbl_wc.grid(row=8, column=0, columnspan=3, pady=8)

        # grid weights
        for c in range(3):
            main.columnconfigure(c, weight=1)
        main.rowconfigure(7, weight=1)

    def _set_defaults(self):
        self.ent_pid.insert(0, DEFAULT_PATIENT_ID_COL)
        self.ent_textcol.insert(0, DEFAULT_TEXT_COL)
        self.ent_keywords.insert(0, ",".join(DEFAULT_KEYWORDS))

    # -----------------------------
    # Handlers
    # -----------------------------
    def on_browse_file(self):
        path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if path:
            self.ent_file.delete(0, tk.END)
            self.ent_file.insert(0, path)
            self._populate_sheets(path)

    def on_reload_sheets(self):
        path = self.ent_file.get().strip()
        if not path:
            messagebox.showerror("Missing file", "Please choose an Excel file first.")
            return
        self._populate_sheets(path)

    def _populate_sheets(self, path):
        try:
            xl = pd.ExcelFile(path)
            names = xl.sheet_names
            self.cmb_sheet["values"] = names
            # Keep current selection if available, else pick first or default guess
            guess = DEFAULT_SHEET_GUESS if DEFAULT_SHEET_GUESS in names else (names[0] if names else "")
            self.cmb_sheet.set(guess)
            self._set_status(f"Loaded {len(names)} sheet(s)")
        except Exception as e:
            messagebox.showerror("Error reading workbook", str(e))

    def on_choose_output(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.ent_out.delete(0, tk.END)
            self.ent_out.insert(0, path)

    def _read_inputs(self):
        file_path = self.ent_file.get().strip()
        sheet_name = self.cmb_sheet.get().strip()
        pid_col = self.ent_pid.get().strip()
        txt_col = self.ent_textcol.get().strip()
        keywords = [k.strip() for k in self.ent_keywords.get().split(",") if k.strip()]
        use_substring = self.var_substring.get()
        return file_path, sheet_name, pid_col, txt_col, keywords, use_substring

    def on_extract(self):
        file_path, sheet, pid_col, txt_col, kws, use_sub = self._read_inputs()
        if not file_path:
            messagebox.showerror("Missing file", "Please choose an Excel file.")
            return
        if not sheet:
            messagebox.showerror("Missing sheet", "Please select a sheet.")
            return
        if not pid_col or not txt_col:
            messagebox.showerror("Missing columns", "Please specify both Patient ID and Text columns.")
            return
        if not kws:
            if not messagebox.askyesno("No keywords", "You didn't provide any keywords. Continue anyway?"):
                return

        try:
            self._set_status("Processing…")
            results, _ = process_sheet(file_path, sheet, kws, pid_col, txt_col, use_substring=use_sub)
            self.data_results = results
            self._populate_tree(results)

            # Save if output path provided
            out = self.ent_out.get().strip()
            if out:
                results.to_excel(out, index=False)
                self._set_status(f"Done. Saved to {out}")
                messagebox.showinfo("Success", f"Keywords extracted and saved to:\n{out}")
            else:
                self._set_status("Done.")
                messagebox.showinfo("Success", "Keywords extracted.")
        except Exception as e:
            self._set_status("Error")
            messagebox.showerror("Error", str(e))

    def on_wordcloud(self):
        file_path, sheet, pid_col, txt_col, kws, use_sub = self._read_inputs()
        if not file_path:
            messagebox.showerror("Missing file", "Please choose an Excel file.")
            return
        if not sheet:
            messagebox.showerror("Missing sheet", "Please select a sheet.")
            return
        try:
            self._set_status("Building word cloud…")
            _, all_words = process_sheet(file_path, sheet, kws, pid_col, txt_col, use_substring=use_sub)
            img = build_wordcloud_image(all_words)
            self.wordcloud_imgtk = ImageTk.PhotoImage(img)
            self.lbl_wc.configure(image=self.wordcloud_imgtk)
            self._set_status("Word cloud ready.")
        except Exception as e:
            self._set_status("Error")
            messagebox.showerror("Error", str(e))

    def _populate_tree(self, df: pd.DataFrame):
        # Clear old columns/items
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())

        cols = list(df.columns.astype(str))
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120, anchor="center")

        # Insert rows
        for _, row in df.iterrows():
            values = [row[c] for c in cols]
            self.tree.insert("", "end", values=values)

    def _set_status(self, text: str):
        self.lbl_status.configure(text=text)


def main():
    root = tk.Tk()
    # Improve default theming if available
    try:
        from tkinter import ttk
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass

    app = HiXComplicationMinerApp(root)
    root.minsize(900, 700)
    root.mainloop()


if __name__ == "__main__":
    main()
