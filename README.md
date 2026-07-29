# HOP Checklist Platform

A client-side checklist generator for HOP (Handover of Plant) site acceptance testing. Designed for telecom field engineers working on IBC/DAS installations.

**Live site:** https://maopeseth.github.io/skyaus/

## Features

- **Template Designer** — Create and edit scope templates with fixed, equipList, connList, and equipment sections
- **Checklist Generator** — Fill site details, auto-populate connection sections from a cable schedule, and generate a formatted `.xlsx` checklist
- **Excel Upload** — Upload a `.xlsx` cable schedule (columns: `cable number`, `cable label`) to auto-fill connection sections by prefix matching

## Converting PDF Cable Schedules

The live site accepts `.xlsx` only. If you have a PDF cable schedule, convert it locally:

```bash
pip install -r requirements.txt
python pdf_to_cable_schedule.py path/to/cable_schedule.pdf
```

Output is saved to a `converted/` folder; upload the `.xlsx` file to the platform.

## Tech

- Pure HTML/CSS/JS — no build step, no framework
- SheetJS / ExcelJS for `.xlsx` reading and generation
- All data stored in browser `localStorage`
- PDF extraction via `pdfplumber` (Python, local use only)
