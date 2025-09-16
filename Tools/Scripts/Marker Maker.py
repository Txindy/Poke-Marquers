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


REQUIRED_COLUMNS: List[str] = ["Name", "Number", "Variant Type"]


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
    # Accept either "Variant Type" or "Variant"; normalize to "Variant" for PDF rendering
    columns_lower = {c.lower(): c for c in df.columns}
    rename: dict[str, str] = {}
    if "name" not in columns_lower or "number" not in columns_lower:
        return None
    if "variant type" in columns_lower:
        rename[columns_lower["variant type"]] = "Variant"
    elif "variant" in columns_lower:
        pass
    else:
        return None
    if columns_lower.get("name") and columns_lower["name"] != "Name":
        rename[columns_lower["name"]] = "Name"
    if columns_lower.get("number") and columns_lower["number"] != "Number":
        rename[columns_lower["number"]] = "Number"
    if rename:
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
    if os.path.isfile(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                candidate = data.get("last_folder")
                if isinstance(candidate, str) and os.path.isdir(candidate):
                    initial_folder = candidate
        except Exception:
            pass
    folder_var = tk.StringVar(value=initial_folder)

    # Save selected folder to config
    def save_last_folder(path: str) -> None:
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump({"last_folder": path}, f)
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
                tk.Label(rows_win, text="Select cards to print (Ctrl/Cmd or Shift for multi-select):").pack(anchor="w", padx=10, pady=(10, 6))
                listbox = tk.Listbox(rows_win, selectmode=tk.EXTENDED, width=80, height=20)
                listbox.pack(fill="both", expand=True, padx=10)
                for _, r in df.iterrows():
                    listbox.insert(tk.END, f"{r['Number']} - {r['Name']} - {r['Variant']}")

                selected_idx: List[int] = []

                def on_rows_ok() -> None:
                    sel = listbox.curselection()
                    if not sel:
                        messagebox.showinfo("No selection", "Please select at least one card or Cancel.")
                        return
                    selected_idx.extend(sel)
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

            btn_full = ttk.Button(row, text="Full List", command=handle_full)
            btn_full.grid(row=0, column=1, sticky="e")
            btn_sel = ttk.Button(row, text="Card Selection", command=handle_select)
            btn_sel.grid(row=0, column=2, sticky="e", padx=(8, 0))

    main_frame = ttk.Frame(root, padding=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="Folder with .xlsx files:").grid(row=0, column=0, sticky="w")
    entry = ttk.Entry(main_frame, textvariable=folder_var, width=52)
    entry.grid(row=1, column=0, padx=(0, 8), sticky="we")

    def browse_folder() -> None:
        selected = filedialog.askdirectory(initialdir=folder_var.get() or script_dir)
        if selected:
            folder_var.set(selected)
            refresh_file_list()
            save_last_folder(selected)

    ttk.Button(main_frame, text="Browse...", command=browse_folder).grid(row=1, column=1)

    ttk.Label(main_frame, text="Files found in folder:").grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 4))
    files_container = ttk.Frame(main_frame)
    files_container.grid(row=3, column=0, columnspan=2, sticky="nsew")

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
    main_frame.rowconfigure(3, weight=1)

    root.mainloop()


if __name__ == "__main__":
    run_gui()

