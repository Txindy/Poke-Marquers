import pandas as pd
import os
import glob
import sys
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def generate_pdfs(input_directory: str) -> tuple[int, list[str]]:
    """Generate PDFs for all .xlsx files in the given directory.

    Returns a tuple of (processed_count, messages).
    """
    original_cwd = os.getcwd()
    os.chdir(input_directory)
    try:
        excel_files = glob.glob("*.xlsx")
        messages: list[str] = []
        processed_count = 0

        if not excel_files:
            return 0, ["No .xlsx files found in the selected folder."]

        for excel_path in excel_files:
            output_path = excel_path.replace('.xlsx', '.pdf')
            try:
                data = pd.read_excel(excel_path, sheet_name="Sheet1")
                data.columns = ["Name", "Number", "Variant"]

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
                for _, row in data.iterrows():
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
                row = []
                for i, card in enumerate(cards, start=1):
                    row.append(card)
                    if i % 4 == 0:
                        rows.append(row)
                        row = []
                if row:
                    rows.append(row)

                main_table = Table(rows, hAlign="CENTER", colWidths=[120] * 4)
                main_table.setStyle(
                    TableStyle([( "VALIGN", (0, 0), (-1, -1), "MIDDLE" )])
                )
                doc.build([main_table])

                messages.append(f"✓ PDF saved as {output_path}")
                processed_count += 1
            except Exception as exc:  # noqa: BLE001
                messages.append(f"✗ Error processing {excel_path}: {str(exc)}")

        return processed_count, messages
    finally:
        os.chdir(original_cwd)


def run_gui() -> None:
    root = tk.Tk()
    root.title("Marker Maker")
    root.geometry("520x160")

    folder_var = tk.StringVar(value=os.getcwd())

    def browse_folder() -> None:
        selected = filedialog.askdirectory(initialdir=folder_var.get() or os.getcwd())
        if selected:
            folder_var.set(selected)

    def do_generate() -> None:
        folder = folder_var.get().strip() or os.getcwd()
        if not os.path.isdir(folder):
            messagebox.showerror("Invalid Folder", "Please choose a valid folder.")
            return
        count, messages = generate_pdfs(folder)
        summary = "\n".join(messages)
        if count == 0 and messages and messages[0].startswith("No .xlsx"):
            messagebox.showwarning("No Files", summary)
        else:
            messagebox.showinfo("Done", f"Processed {count} Excel file(s)\n\n{summary}")

    main_frame = ttk.Frame(root, padding=12)
    main_frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main_frame, text="Folder with .xlsx files:").grid(row=0, column=0, sticky="w")
    entry = ttk.Entry(main_frame, textvariable=folder_var, width=52)
    entry.grid(row=1, column=0, padx=(0, 8), sticky="we")
    ttk.Button(main_frame, text="Browse...", command=browse_folder).grid(row=1, column=1)

    ttk.Separator(main_frame).grid(row=2, column=0, columnspan=2, pady=10, sticky="we")
    ttk.Button(main_frame, text="Generate PDFs", command=do_generate).grid(row=3, column=0, columnspan=2)

    main_frame.columnconfigure(0, weight=1)

    root.mainloop()


if __name__ == "__main__":
    run_gui()
