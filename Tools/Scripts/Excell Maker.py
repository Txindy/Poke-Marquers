import re
import sys
from collections import defaultdict
from typing import List, Dict, Tuple

try:
    import pandas as pd
except ImportError:
    print("This script requires pandas. Install with: pip install pandas openpyxl")
    sys.exit(1)

def read_blocked_input(prompt: str) -> List[str]:
    print(prompt)
    print("(Paste the list, then enter a single '.' on a new line to finish)")
    lines = []
    while True:
        try:
            line = input()
        except EOFError:
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

        # Consume a few lines to find set code; also capture set name as fallback
        set_name = ""
        variant = ""
        lookahead = 0
        while i < n and lookahead < 6 and not looks_like_price(lines[i]):
            s = lines[i].strip()
            if not set_name and s and not looks_like_set_code(s) and not looks_like_number(s):
                set_name = s
            if looks_like_set_code(s) and not variant:
                variant = s
            lookahead += 1
            i += 1

        # Skip until price line end of block (optional)
        while i < n and not looks_like_price(lines[i]):
            # prevent infinite loops; break if we encounter another qty which suggests next block
            if is_qty(lines[i]):
                break
            i += 1
        # Skip price line if present
        if i < n and looks_like_price(lines[i]):
            i += 1

        if name and number:
            # Always use override if provided, otherwise use parsed variant or set_name
            if variant_override:
                final_variant = variant_override
            else:
                final_variant = variant if variant else set_name
            
            items.append({
                "Name": name,
                "Number": number,
                "Variant Type": final_variant
            })

    return items

def merge_items_with_duplicates(items_a: List[Dict[str, str]], items_b: List[Dict[str, str]]) -> List[Dict[str, str]]:
    # Start with first list
    result = items_a.copy()
    
    # Track which cards from first list we've seen
    seen_cards = set()
    for item in items_a:
        key = (item["Name"], item["Number"])
        seen_cards.add(key)
    
    # Add items from second list
    for item in items_b:
        key = (item["Name"], item["Number"])
        if key in seen_cards:
            # This is a duplicate - add as Reverse Holo after the original
            reverse_item = {
                "Name": item["Name"],
                "Number": item["Number"],
                "Variant Type": "Reverse Holo"
            }
            result.append(reverse_item)
        else:
            # New card - add as is
            result.append(item)
    
    return result

def main():
    lines_a = read_blocked_input("Paste FIRST list:")
    lines_b = read_blocked_input("Paste SECOND list:")
    items_a = parse_list(lines_a, variant_override="Normal")
    items_b = parse_list(lines_b)
    merged = merge_items_with_duplicates(items_a, items_b)

    if not merged:
        print("No cards parsed. Please check the input format.")
        return

    df = pd.DataFrame(merged, columns=["Name", "Number", "Variant Type"])
    out_file = "cards.xlsx"
    df.to_excel(out_file, index=False)
    print(f"Wrote {len(df)} rows to {out_file}")

if __name__ == "__main__":
    main()