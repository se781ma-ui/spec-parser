# Specification Parser (BOM Tree)
Web application that parses engineering specifications from SolidWorks (exported as Excel or Word) and builds a full BOM tree — from top-level assembly down to individual parts — to automate the production request process.

## The Problem
When working with SolidWorks, specifications are exported as Excel or Word files. Equipment is structured as a hierarchical tree:

To create a production request for parts and assemblies, engineers had to manually trace this tree — which takes **1–2 hours** per request.

## The Solution

This parser automatically:

- Reads specifications from Excel and Word files
- Expands the tree from the top-level assembly down to individual parts
- Builds a production request with the ability to edit or add items
- Auto-fills manufacturing dates
- Auto-generates headers (customizable via config files)
- Exports the final request to Excel

## Features

- Upload Excel / Word specification files
- Parse hierarchical BOM tree (assemblies → sub-assemblies → parts)
- Auto-generate production request
- Edit / add items before export
- Auto-fill dates
- Custom header templates (via config)
- Export to Excel
- Web UI

## Technical Challenge: Encrypted Specifications

SolidWorks exports specifications via macro and **encrypts** the output files, which prevents standard parsing. 

To solve this, I reverse-engineered the export format and implemented a **decryption layer** that:
- Detects the encryption signature
- Extracts the original XML/data structure
- Bypasses the macro protection without modifying SolidWorks itself

This allows the parser to read specification files that standard tools cannot open.

## Tech Stack

- **Backend:** Python
- **Frontend:** HTML
- **Data:** Excel / Word parsing (openpyxl, python-docx)
- **Output:** Excel

## How It Works

1. User uploads a specification file (Excel or Word).
2. Parser extracts the hierarchical BOM tree.
3. System expands the tree from top-level assembly to parts.
4. Auto-generated production request is displayed.
5. User can edit or add items.
6. Request is exported to Excel with headers and dates.

## Use Case

**Before:** 1–2 hours of manual work per production request.
**After:** a few clicks and auto-generated Excel file.

## License

MIT
