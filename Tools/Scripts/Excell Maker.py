#!/usr/bin/env python3
import re
import sys
import os
from collections import defaultdict
from typing import List, Dict, Tuple, Optional
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("This script requires pandas. Install with: pip install pandas openpyxl")
    sys.exit(1)

# Fix stdin for PyInstaller onefile executables
if hasattr(sys, '_MEIPASS'):
    # Running as PyInstaller executable
    if sys.stdin is None or not sys.stdin.isatty():
        # Reopen stdin if it's been lost
        try:
            sys.stdin = open('CONIN$', 'r')
        except Exception:
            pass

def read_blocked_input(prompt: str) -> List[str]:
    print(prompt)
    print("(Paste the list, then enter a single '.' on a new line to finish)")
    lines = []
    while True:
        try:
            line = input()
        except (EOFError, RuntimeError):
            break
        if line.strip() == '.':
            break
        lines.append(line.rstrip('\n'))
    return lines

def parse_list(lines: List[str], variant_override: str = None) -> List[Dict[str, str]]:
    # We assume blocks that look like:
    # qty
    # Card Name
    # 001/182
    # Set Name
    # Set Name (duplicate)
    # PAR
    # Type
    # Rarity
    # $price
    items = []
    i = 0
    n = len(lines)

    def is_qty(s: str) -> bool:
        return bool(re.fullmatch(r'\d+', s.strip()))

    def looks_like_number(s: str) -> bool:
        return '/' in s

    def looks_like_price(s: str) -> bool:
        return s.strip().startswith('$')

    def looks_like_set_code(s: str) -> bool:
        # 2–5 uppercase letters/numbers, often 3; allow hyphen (e.g., SWSH)
        return bool(re.fullmatch(r'[A-Z0-9-]{2,5}', s.strip()))

    while i < n:
        # Skip empties
        while i < n and not lines[i].strip():
            i += 1
        if i >= n:
            break

        # Expect qty line; if not, try to sync by advancing
        if not is_qty(lines[i]):
            i += 1
            continue

        i += 1
        # Card name
        if i >= n:
            break
        name = lines[i].strip()
        i += 1

        # Number (e.g., 001/182)
        number = ""
        if i < n and looks_like_number(lines[i]):
            number = lines[i].strip()
            i += 1

        # Parse the structure: Set Name, Set Name (duplicate), Set Code, Type, Rarity, Price
        set_name = ""
        set_code = ""
        card_type = ""
        rarity = ""
        
        # First set name (line after number) - capture first non-empty line that's not a set code
        if i < n:
            potential_set_name = lines[i].strip()
            if potential_set_name and not looks_like_set_code(potential_set_name) and not looks_like_price(potential_set_name):
                set_name = potential_set_name
                i += 1
        
        # Skip duplicate set name if present
        if i < n and lines[i].strip() == set_name:
            i += 1
        
        # Set code (e.g., SVI) - look for 2-5 uppercase letters/numbers
        if i < n and looks_like_set_code(lines[i]):
            set_code = lines[i].strip()
            i += 1
        
        # Type (e.g., Grass, Fire, etc.) - any non-empty line that's not price, qty, or set code
        if i < n and not looks_like_price(lines[i]) and not is_qty(lines[i]) and not looks_like_set_code(lines[i]):
            card_type = lines[i].strip()
            i += 1
        
        # Rarity (e.g., Common, Uncommon, Rare) - this is what we'll use as Card variant
        # It's the last non-price, non-qty line before the price
        if i < n and not looks_like_price(lines[i]) and not is_qty(lines[i]):
            rarity = lines[i].strip()
            i += 1
        
        # Skip price line if present
        if i < n and looks_like_price(lines[i]):
            i += 1

        if name and number:
            # Card variant uses Normal or override, not rarity
            if variant_override:
                final_variant = variant_override
            else:
                final_variant = "Normal"
            
            # Extract card number sorting order (the numeric part before the slash)
            card_number_sort = ""
            if '/' in number:
                try:
                    card_number_sort = int(number.split('/')[0])
                except ValueError:
                    card_number_sort = 0
            
            # Create item with all tcgcollector.csv columns
            items.append({
                "TCG region": "International",
                "Expansion": set_name if set_name else "",
                "Card number": number,
                "Card number sorting order": card_number_sort,
                "Card name": name,
                "Rarity": rarity if rarity else "",  # Store the rarity value from input
                "Card variant": final_variant,  # Use Normal or override
                "Card language": "English",
                "Card condition": "Unspecified",
                "Note": "",
                "Quantity": ""  # Remains empty
            })

    return items

def merge_items_with_duplicates(items_a: List[Dict[str, str]], items_b: List[Dict[str, str]]) -> List[Dict[str, str]]:
    # Start with first list
    result = items_a.copy()
    
    # Track which cards from first list we've seen
    seen_cards = set()
    for item in items_a:
        key = (item["Card name"], item["Card number"])
        seen_cards.add(key)
    
    # Add items from second list
    for item in items_b:
        key = (item["Card name"], item["Card number"])
        if key in seen_cards:
            # This is a duplicate - add as Reverse Holo variant
            reverse_item = item.copy()
            reverse_item["Card variant"] = "Reverse Holo"
            result.append(reverse_item)
        else:
            # New card - add as is
            result.append(item)
    
    return result

def sort_excel_file(file_path: str) -> None:
    """Sort the Excel file by Card number sorting order and Card variant"""
    try:
        df = pd.read_excel(file_path)
        
        required = {"Card number", "Card variant", "Card number sorting order"}
        missing = required - set(df.columns)
        if missing:
            print(f"Error: missing columns: {', '.join(sorted(missing))}")
            return
        
        # Create sorting order for variant types
        variant_order = {"Normal": 0, "Reverse Holo": 1, "Normal Holo": 2}
        df["VariantSort"] = df["Card variant"].map(variant_order).fillna(99)
        
        # Ensure Card number sorting order is numeric
        df["Card number sorting order"] = pd.to_numeric(df["Card number sorting order"], errors='coerce').fillna(0)
        
        # Sort by Card number sorting order first, then by Variant Type
        df.sort_values(by=["Card number sorting order", "VariantSort"], inplace=True)
        df.drop(columns=["VariantSort"], inplace=True)
        
        # Save the sorted data back to the same file
        df.to_excel(file_path, index=False)
        print("✅ File has been automatically sorted and saved!")
        
    except Exception as e:
        print(f"Error sorting file: {e}")

def main():
    print("=== Excel Maker and Sorter ===")
    print("This script will create an Excel file and automatically sort it.\n")
    
    # Get set information from user
    try:
        set_name = input("Enter Set Name: ").strip()
        set_number = input("Enter Set Number: ").strip()
        set_era = input("Enter Set Era (e.g., SV, SWSH): ").strip()
    except (EOFError, RuntimeError) as e:
        print(f"Error reading input: {e}")
        print("Please run this program from a command prompt/terminal.")
        input("Press Enter to exit...")
        return
    
    if not set_name or not set_number or not set_era:
        print("Error: Set Name, Set Number, and Set Era are required.")
        return
    
    print()  # Empty line for spacing
    
    # Get input from user
    lines_a = read_blocked_input("Paste Standard list:")
    lines_b = read_blocked_input("Paste Parallel list:")
    
    # Parse the lists - standard list uses "Normal" variant, parallel list uses its own variant
    items_a = parse_list(lines_a, variant_override="Normal")
    items_b = parse_list(lines_b, variant_override=None)
    merged = merge_items_with_duplicates(items_a, items_b)

    if not merged:
        print("No cards parsed. Please check the input format.")
        return

    # Define all columns in the order matching tcgcollector.csv format
    tcgcollector_columns = [
        "TCG region",
        "Expansion",
        "Card number",
        "Card number sorting order",
        "Card name",
        "Rarity",
        "Card variant",
        "Card language",
        "Card condition",
        "Note",
        "Quantity"
    ]
    
    # Create DataFrame with all required columns
    df = pd.DataFrame(merged, columns=tcgcollector_columns)
    
    # Create output filename: "Master Set {set era} {set number} {set name}.xlsx"
    out_file = f"Master Set {set_era} {set_number} - {set_name}.xlsx"
    df.to_excel(out_file, index=False)
    print(f"📝 Created Excel file with {len(df)} rows: {out_file}")
    
    # Automatically sort the created file
    print("🔄 Automatically sorting the Excel file...")
    sort_excel_file(out_file)
    
    print(f"\n✅ Complete! The sorted Excel file '{out_file}' is ready to use.")

if __name__ == "__main__":
    main()
