import pandas as pd
import os
import json
from typing import List, Optional
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def list_excel_files(folder_path: str) -> List[str]:
    files = []
    try:
        for fname in os.listdir(folder_path):
            if fname.lower().endswith(".xlsx") and not fname.startswith("~$"):
                files.append(os.path.join(folder_path, fname))
    except Exception:
        return []
    files.sort()
    return files


def read_excel_normalized(path: str) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_excel(path)
    except Exception:
        return None
    # Map from new structure: "Card name", "Card number", "Card variant"
    columns_lower = {c.lower(): c for c in df.columns}
    rename: dict[str, str] = {}
    
    # Check for required column names (Card name, Card number, Card variant)
    if "card name" not in columns_lower:
        return None
    rename[columns_lower["card name"]] = "Name"
    
    if "card number" not in columns_lower:
        return None
    rename[columns_lower["card number"]] = "Number"
    
    if "card variant" not in columns_lower:
        return None
    rename[columns_lower["card variant"]] = "Variant"
    
    df = df.rename(columns=rename)
    needed = ["Name", "Number", "Variant"]
    if not all(c in df.columns for c in needed):
        return None
    return df[needed].copy()


def read_csv_normalized(path: str) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_csv(path)
    except Exception:
        return None
    # Map from new structure: "Card name", "Card number", "Card variant"
    columns_lower = {c.lower(): c for c in df.columns}
    rename: dict[str, str] = {}
    
    # Check for required column names (Card name, Card number, Card variant)
    if "card name" not in columns_lower:
        return None
    rename[columns_lower["card name"]] = "Name"
    
    if "card number" not in columns_lower:
        return None
    rename[columns_lower["card number"]] = "Number"
    
    if "card variant" not in columns_lower:
        return None
    rename[columns_lower["card variant"]] = "Variant"
    
    df = df.rename(columns=rename)
    needed = ["Name", "Number", "Variant"]
    if not all(c in df.columns for c in needed):
        return None
    return df[needed].copy()


def export_pdf_from_dataframe(df: pd.DataFrame, output_path: str) -> None:
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20,
    )
    styles = getSampleStyleSheet()
    style_center = styles["Normal"]
    style_center.alignment = 1

    cards = []
    for _, row in df.iterrows():
        content = [
            Paragraph(f"<b>{row['Name']}</b>", style_center),
            Paragraph(str(row['Number']), style_center),
            Paragraph(str(row['Variant']), style_center),
        ]
        card_table = Table([[content]], colWidths=100, rowHeights=60)
        card_table.setStyle(
            TableStyle([
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ])
        )
        cards.append(card_table)

    rows = []
    row_cells = []
    for i, card in enumerate(cards, start=1):
        row_cells.append(card)
        if i % 4 == 0:
            rows.append(row_cells)
            row_cells = []
    if row_cells:
        rows.append(row_cells)

    main_table = Table(rows, hAlign="CENTER", colWidths=[120] * 4)
    main_table.setStyle(TableStyle([( "VALIGN", (0, 0), (-1, -1), "MIDDLE" )]))
    doc.build([main_table])


def run_gui() -> None:
    root = tk.Tk()
    root.title("Marker Maker")
    root.geometry("700x520")

    # Determine base directory: when frozen (PyInstaller), use the EXE's directory;
    # otherwise use the script's directory. Avoid temporary _MEI* paths.
    try:
        if getattr(__import__('sys'), 'frozen', False):
            script_dir = os.path.dirname(__import__('sys').executable)
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        script_dir = os.getcwd()
    # Always open at project root by default, unless a saved folder exists in JSON
    config_path = os.path.join(script_dir, "marker_maker_config.json")
    initial_folder = script_dir
    initial_csv = ""
    if os.path.isfile(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                candidate = data.get("last_folder")
                if isinstance(candidate, str) and os.path.isdir(candidate):
                    initial_folder = candidate
                candidate_csv = data.get("last_csv")
                if isinstance(candidate_csv, str) and os.path.isfile(candidate_csv):
                    initial_csv = candidate_csv
        except Exception:
            pass
    folder_var = tk.StringVar(value=initial_folder)
    csv_file_var = tk.StringVar(value=initial_csv)

    # Save selected folder and CSV file to config
    def save_config() -> None:
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump({
                    "last_folder": folder_var.get(),
                    "last_csv": csv_file_var.get()
                }, f)
        except Exception:
            pass

    # Container for per-file row widgets
    file_row_widgets: list[tk.Widget] = []

    def refresh_file_list() -> None:
        for w in file_row_widgets:
            try:
                w.destroy()
            except Exception:
                pass
        file_row_widgets.clear()
        folder = folder_var.get().strip() or script_dir
        if not os.path.isdir(folder):
            return
        files = list_excel_files(folder)
        for p in files:
            display = os.path.splitext(os.path.basename(p))[0]
            row = ttk.Frame(files_list_frame)
            row.pack(fill="x", anchor="w")
            file_row_widgets.append(row)

            row.columnconfigure(0, weight=1)

            lbl = ttk.Label(row, text=display)
            lbl.grid(row=0, column=0, sticky="w", padx=(0, 8))

            def handle_full(path: str = p) -> None:
                df = read_excel_normalized(path)
                if df is None:
                    messagebox.showwarning("Missing Columns", f"Skipping file (missing required columns):\n{os.path.basename(path)}")
                    return
                base = os.path.splitext(os.path.basename(path))[0]
                out_pdf = os.path.join(os.path.dirname(path), f"{base} - Full.pdf")
                try:
                    export_pdf_from_dataframe(df, out_pdf)
                    messagebox.showinfo("Exported", f"PDF saved as:\n{out_pdf}")
                except Exception as exc:
                    messagebox.showerror("Error", f"Error processing {os.path.basename(path)}:\n{exc}")

            def handle_select(path: str = p) -> None:
                df = read_excel_normalized(path)
                if df is None:
                    messagebox.showwarning("Missing Columns", f"Skipping file (missing required columns):\n{os.path.basename(path)}")
                    return
                rows_win = tk.Toplevel(root)
                rows_win.title(f"Select Cards - {os.path.basename(path)}")
                tk.Label(rows_win, text="Check cards to print:").pack(anchor="w", padx=10, pady=(10, 6))

                # Scrollable area with checkboxes
                container = ttk.Frame(rows_win)
                container.pack(fill="both", expand=True, padx=10)
                canvas = tk.Canvas(container, height=360)
                vscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
                inner = ttk.Frame(canvas)
                inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
                canvas.create_window((0, 0), window=inner, anchor="nw")
                canvas.configure(yscrollcommand=vscroll.set)
                canvas.pack(side="left", fill="both", expand=True)
                vscroll.pack(side="right", fill="y")

                # Enable mouse wheel scrolling on all platforms
                def _on_mousewheel(event):
                    try:
                        delta = int(-1 * (event.delta / 120))
                        canvas.yview_scroll(delta, "units")
                    except Exception:
                        pass
                    return "break"

                # Windows and macOS
                canvas.bind_all("<MouseWheel>", _on_mousewheel)
                # Linux
                canvas.bind_all("<Button-4>", lambda e: (canvas.yview_scroll(-1, "units"), "break"))
                canvas.bind_all("<Button-5>", lambda e: (canvas.yview_scroll(1, "units"), "break"))

                vars_list: List[tk.BooleanVar] = []
                for _, r in df.iterrows():
                    var = tk.BooleanVar(value=False)
                    text = f"{r['Number']} - {r['Name']} - {r['Variant']}"
                    cb = tk.Checkbutton(inner, text=text, variable=var, anchor="w", justify="left")
                    cb.pack(fill="x", anchor="w")
                    vars_list.append(var)

                selected_idx: List[int] = []

                def on_rows_ok() -> None:
                    idxs = [i for i, v in enumerate(vars_list) if v.get()]
                    if not idxs:
                        messagebox.showinfo("No selection", "Please check at least one card or Cancel.")
                        return
                    selected_idx.extend(idxs)
                    rows_win.destroy()

                def on_rows_cancel() -> None:
                    selected_idx.clear()
                    rows_win.destroy()

                btns2 = tk.Frame(rows_win)
                btns2.pack(fill="x", padx=10, pady=10)
                tk.Button(btns2, text="Cancel", command=on_rows_cancel).pack(side="right")
                tk.Button(btns2, text="Print Selected", command=on_rows_ok).pack(side="right", padx=(0, 8))

                rows_win.transient(root)
                rows_win.grab_set()
                root.wait_window(rows_win)

                if not selected_idx:
                    return
                export_subset = df.iloc[selected_idx].reset_index(drop=True)
                base = os.path.splitext(os.path.basename(path))[0]
                out_pdf = os.path.join(os.path.dirname(path), f"{base} - Selected.pdf")
                try:
                    export_pdf_from_dataframe(export_subset, out_pdf)
                    messagebox.showinfo("Exported", f"PDF saved as:\n{out_pdf}")
                except Exception as exc:
                    messagebox.showerror("Error", f"Error processing {os.path.basename(path)}:\n{exc}")

            def handle_missing(path: str = p) -> None:
                # Read master set file (contains all cards from that set)
                master_df = read_excel_normalized(path)
                if master_df is None:
                    messagebox.showwarning("Missing Columns", f"Skipping file (missing required columns):\n{os.path.basename(path)}")
                    return
                
                # Get TCGCollector.csv file path (contains cards already in collection)
                csv_path = csv_file_var.get().strip()
                if not csv_path or not os.path.isfile(csv_path):
                    messagebox.showwarning("CSV File Required", "Please select a TCGCollector.csv file first.")
                    return
                
                # Read TCGCollector.csv file
                csv_df = read_csv_normalized(csv_path)
                if csv_df is None:
                    messagebox.showwarning("Invalid CSV", f"CSV file is missing required columns:\n{os.path.basename(csv_path)}")
                    return
                
                # Create sets of tuples (Name, Number, Variant) for comparison
                # Master set contains all cards from the set
                master_set = set(zip(master_df['Name'], master_df['Number'], master_df['Variant']))
                # CSV contains cards already in collection (may have duplicates, but set handles that)
                csv_set = set(zip(csv_df['Name'], csv_df['Number'], csv_df['Variant']))
                
                # Find cards in master set that are NOT in collection (missing cards to complete the set)
                missing_cards = master_set - csv_set
                
                if not missing_cards:
                    messagebox.showinfo("No Missing Cards", "All cards from the master set are present in your TCGCollector collection!")
                    return
                
                # Convert back to DataFrame for display
                missing_list = [{"Name": n, "Number": num, "Variant": v} for n, num, v in missing_cards]
                missing_df = pd.DataFrame(missing_list)
                missing_df = missing_df.sort_values(by=["Number", "Name", "Variant"]).reset_index(drop=True)
                
                # Show missing cards in a selection window
                rows_win = tk.Toplevel(root)
                rows_win.title(f"Missing Cards - {os.path.basename(path)}")
                tk.Label(rows_win, text=f"Found {len(missing_df)} missing card(s):").pack(anchor="w", padx=10, pady=(10, 6))

                # Scrollable area with checkboxes
                container = ttk.Frame(rows_win)
                container.pack(fill="both", expand=True, padx=10)
                canvas = tk.Canvas(container, height=360)
                vscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
                inner = ttk.Frame(canvas)
                inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
                canvas.create_window((0, 0), window=inner, anchor="nw")
                canvas.configure(yscrollcommand=vscroll.set)
                canvas.pack(side="left", fill="both", expand=True)
                vscroll.pack(side="right", fill="y")

                # Enable mouse wheel scrolling on all platforms
                def _on_mousewheel(event):
                    try:
                        delta = int(-1 * (event.delta / 120))
                        canvas.yview_scroll(delta, "units")
                    except Exception:
                        pass
                    return "break"

                # Windows and macOS
                canvas.bind_all("<MouseWheel>", _on_mousewheel)
                # Linux
                canvas.bind_all("<Button-4>", lambda e: (canvas.yview_scroll(-1, "units"), "break"))
                canvas.bind_all("<Button-5>", lambda e: (canvas.yview_scroll(1, "units"), "break"))

                vars_list: List[tk.BooleanVar] = []
                for _, r in missing_df.iterrows():
                    var = tk.BooleanVar(value=True)  # All selected by default
                    text = f"{r['Number']} - {r['Name']} - {r['Variant']}"
                    cb = tk.Checkbutton(inner, text=text, variable=var, anchor="w", justify="left")
                    cb.pack(fill="x", anchor="w")
                    vars_list.append(var)

                selected_idx: List[int] = []

                def on_rows_ok() -> None:
                    idxs = [i for i, v in enumerate(vars_list) if v.get()]
                    if not idxs:
                        messagebox.showinfo("No selection", "Please check at least one card or Cancel.")
                        return
                    selected_idx.extend(idxs)
                    rows_win.destroy()

                def on_rows_cancel() -> None:
                    selected_idx.clear()
                    rows_win.destroy()

                btns2 = tk.Frame(rows_win)
                btns2.pack(fill="x", padx=10, pady=10)
                tk.Button(btns2, text="Cancel", command=on_rows_cancel).pack(side="right")
                tk.Button(btns2, text="Print Selected", command=on_rows_ok).pack(side="right", padx=(0, 8))

                rows_win.transient(root)
                rows_win.grab_set()
                root.wait_window(rows_win)

                if not selected_idx:
                    return
                export_subset = missing_df.iloc[selected_idx].reset_index(drop=True)
                base = os.path.splitext(os.path.basename(path))[0]
                out_pdf = os.path.join(os.path.dirname(path), f"{base} - Missing Cards.pdf")
                try:
                    export_pdf_from_dataframe(export_subset, out_pdf)
                    messagebox.showinfo("Exported", f"PDF saved as:\n{out_pdf}")
                except Exception as exc:
                    messagebox.showerror("Error", f"Error processing {os.path.basename(path)}:\n{exc}")

            btn_full = ttk.Button(row, text="Full List", command=handle_full)
            btn_full.grid(row=0, column=1, sticky="e")
            btn_sel = ttk.Button(row, text="Card Selection", command=handle_select)
            btn_sel.grid(row=0, column=2, sticky="e", padx=(8, 0))
            btn_missing = ttk.Button(row, text="Missing Cards", command=handle_missing)
            btn_missing.grid(row=0, column=3, sticky="e", padx=(8, 0))

    main_frame = ttk.Frame(root, padding=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="Master Sets.xlsx files:").grid(row=0, column=0, sticky="w")
    entry = ttk.Entry(main_frame, textvariable=folder_var, width=52)
    entry.grid(row=1, column=0, padx=(0, 8), sticky="we")

    def browse_folder() -> None:
        selected = filedialog.askdirectory(initialdir=folder_var.get() or script_dir)
        if selected:
            folder_var.set(selected)
            refresh_file_list()
            save_config()

    def browse_csv() -> None:
        selected = filedialog.askopenfilename(
            title="Select TCGCollector.csv file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=os.path.dirname(csv_file_var.get()) if csv_file_var.get() else script_dir
        )
        if selected:
            csv_file_var.set(selected)
            save_config()

    ttk.Button(main_frame, text="Browse...", command=browse_folder).grid(row=1, column=1)

    ttk.Label(main_frame, text="TCGCollector.csv:").grid(row=2, column=0, sticky="w", pady=(10, 0))
    csv_entry = ttk.Entry(main_frame, textvariable=csv_file_var, width=52)
    csv_entry.grid(row=3, column=0, padx=(0, 8), sticky="we")
    ttk.Button(main_frame, text="Browse...", command=browse_csv).grid(row=3, column=1)

    ttk.Label(main_frame, text="Files found in folder:").grid(row=4, column=0, columnspan=2, sticky="w", pady=(10, 4))
    files_container = ttk.Frame(main_frame)
    files_container.grid(row=5, column=0, columnspan=2, sticky="nsew")

    canvas = tk.Canvas(files_container, height=300)
    scrollbar = ttk.Scrollbar(files_container, orient="vertical", command=canvas.yview)
    files_list_frame = ttk.Frame(canvas)
    files_list_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=files_list_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Initial load: last-folder if available, else script_dir\Master Sets
    try:
        refresh_file_list()
    except Exception:
        pass

    main_frame.columnconfigure(0, weight=1)
    main_frame.rowconfigure(5, weight=1)

    root.mainloop()


if __name__ == "__main__":
    run_gui()

