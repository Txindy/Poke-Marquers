# Poke-Marquers

A comprehensive Python toolkit for managing Pokémon card collections and creating custom markers. This project provides three essential tools for Pokémon card enthusiasts to organize, process, and create visual markers for their collections.

## Supported Sets

### Completed Sets
- **Master Set SV 1** - Scarlet and Violet (Base Set) ✅
- **Master Set SV 2** - Paldea Evolved ✅
- **Master Set SV 3** - Obsidian Flames ✅
- **Master Set SV 3.5** - Scarlet & Violet 151 ✅

### In Progress
- **Master Set SV 4** - Paradox Rift 

### Coming Soon
- Master Set SV 4.5 - Paldean Fates
- Master Set SV 5 - Temporal Forces
- Master Set SV 6 - Twilight Masquerade
- Master Set SV 6.5 - Shrouded Fable
- Master Set SV 7 - Stellar Crown
- Master Set SV 8 - Surging Sparks
- Master Set SV 8.5 - Prismatic Evolutions
- Master Set SV 9 - Journey Together
- Master Set SV 10 - Destined Rivals
- Master Set SV 10.5.1 - Black Bolt
- Master Set SV 10.5.2 - White Flare

## Features

- **Excel Maker**: Convert card lists into organized Excel spreadsheets
- **Excel Sorter**: Automatically sort Excel files by card number and variant type
- **Marker Maker**: Generate custom markers for your Pokémon card collection

## Tools Overview

### 1. Excel Maker (`Tools/Excell Maker`)

Converts raw card data into organized Excel spreadsheets with proper formatting.

**Features:**
- Parses card lists from pasted text input
- Handles duplicate cards by creating "Reverse Holo" variants
- Exports to Excel with standardized columns: "Name", "Number", "Variant Type"
- Supports multiple input formats


**How to use:**
1. Run the script
2. Paste your first card list when prompted(Copy Standart list from https://www.tcgcollector.com/)
3. Enter a single '.' on a new line to finish the first list
4. Paste your second card list when prompted(Copy Parallel list from https://www.tcgcollector.com/)
5. Enter a single '.' on a new line to finish the second list
6. The script will generate a `cards.xlsx` file with merged and organized data

**Input Format:**
The script expects card data in this format:
```
1
Card Name
001/182
Set Name
PAR
Type
Rarity
$price
```

### 2. Excel Sorter (`Tools/Excell Sorter`)

Automatically sorts Excel files by card number and variant type for better organization.

**Features:**
- Sorts by card number (ascending)
- Secondary sort by variant type (Normal, then Reverse Holo)
- Works with any Excel file in the script directory
- Command-line interface for flexibility
- Reorder Excel file by card number mantaining Reverso Holo below normal variant

**Requirements:**
- Excel file must contain "Number" and "Variant Type" columns
- File must be in .xlsx or .xlsm format

**How to use:**
- Excel file must be located on same folder as Excell sorter
- Run the tool

### 3. Marker Maker

Creates custom visual markers for your Pokémon card collection.

**Features:**
- Simple UI for easy operation
- Generates markers for completed sets
- Supports all Scarlet & Violet series sets


## Workflow

1. **Collect Data**: Use Excel Maker to convert your card lists into organized spreadsheets
2. **Organize**: Use Excel Sorter to arrange your data by card number and variant
3. **Create Markers**: Use Marker Maker to generate visual markers for your collection

## Prerequisites

- Python 3.7 or higher
- Required Python packages:
  - `pandas`
  - `openpyxl`

Install dependencies with:
```bash
pip install pandas openpyxl
```

## File Structure

```
Poke-Marquers/
├── README.md
├── Tools/
│   └── Scripts/
│       ├── Excell Maker.py
│       └── Excell Sorter.py
└── marker_maker.py
```

## Contributing

This project is designed for Pokémon card collectors. Feel free to submit issues or feature requests for additional set support or functionality improvements.

## License

This project is for personal use by Pokémon card collectors.

