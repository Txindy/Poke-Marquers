import pandas as pd
import os
import glob
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

# === SETTINGS ===
# Find all Excel files in the current directory
excel_files = glob.glob("*.xlsx")

if not excel_files:
    print("No .xlsx files found in the current directory!")
    exit()

# Process each Excel file
for excel_path in excel_files:
    # Generate output PDF name (same name, different extension)
    output_path = excel_path.replace('.xlsx', '.pdf')
    
    print(f"Processing: {excel_path}")
    
    try:
        # Load Excel
        data = pd.read_excel(excel_path, sheet_name="Sheet1")
        data.columns = ["Name", "Number", "Variant"]

        # Create PDF document
        doc = SimpleDocTemplate(output_path, pagesize=A4,
                                rightMargin=20, leftMargin=20,
                                topMargin=20, bottomMargin=20)

        styles = getSampleStyleSheet()
        style_center = styles["Normal"]
        style_center.alignment = 1  # center text

        # Build marker cards
        cards = []
        for _, row in data.iterrows():
            content = [
                Paragraph(f"<b>{row['Name']}</b>", style_center),
                Paragraph(str(row['Number']), style_center),
                Paragraph(str(row['Variant']), style_center)
            ]
            card_table = Table([[content]], colWidths=100, rowHeights=60)
            card_table.setStyle(TableStyle([
                ("BOX", (0, 0), (-1, -1), 1, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            cards.append(card_table)

        # Arrange into rows (4 per row)
        rows = []
        row = []
        for i, card in enumerate(cards, start=1):
            row.append(card)
            if i % 4 == 0:
                rows.append(row)
                row = []
        if row:
            rows.append(row)

        # Build final layout
        main_table = Table(rows, hAlign="CENTER", colWidths=[120]*4)
        main_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
        doc.build([main_table])

        print(f"✓ PDF saved as {output_path}")
        
    except Exception as e:
        print(f"✗ Error processing {excel_path}: {str(e)}")

print(f"\nProcessed {len(excel_files)} Excel file(s)")
