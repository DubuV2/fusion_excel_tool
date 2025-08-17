# Fusion Excel Tool
> A lightweight tool to merge multiple CSV and Excel files into a single output.

## Overview
Fusion Excel Tool is a simple utility designed to:
- Merge multiple `.csv` and `.xlsx` files into one output file.
- Handle different CSV encodings (`UTF-8`, `Latin-1`, `CP1252`).
- Provide a progress bar and status updates during execution.
- Preview the merged data before saving.
- Automatically saves user settings for future use

This project is aimed at making data consolidation easier for non-technical users as well as developers.

---

## Installing / Getting Started

### Option 1 – Windows Executable
The easiest way to use the tool is to download the precompiled `.exe` file from the [Releases](https://github.com/DubuV2/fusion_excel_tool/releases) page.  
Simply run it — no installation required.

### Option 2 – Run with Python
Clone the repository and install the dependencies:

```bash
git clone https://github.com/DubuV2/fusion_excel_tool.git
cd fusion_excel_tool
pip install -r requirements.txt
python fusion_excel.py
