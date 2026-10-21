# QA Toolkit

A collection of practical QA/PM utilities for documentation conversion, bug extraction, and test planning.

## Contents

- scripts/converter_para_md.py — Batch convert .docx, .pptx, .pdf, .xlsx to Markdown (uses pandoc, PyMuPDF, python-pptx, openpyxl).
- scripts/extrair_bugs.py — Extract BUG-YYYYMMDD-NN entries from a Markdown bug report and generate a Taiga-compatible JSON import file.
- scripts/gerar-plano-testes-generico-xlsx.py — Generate a generic smoke-test plan Excel sheet with validation dropdowns and a defect template.
- scripts/pandoc_tabelas_simples.lua — Pandoc Lua filter for cleaner simple tables.
- 	emplates/ — Ready-to-use templates (PRD, QA/PM guides).
- ules/ — AI-assistant rules for QA/PM workflows.

## Requirements

- Python 3.9+
- pandoc (for document conversion)
- See individual scripts for extra deps

## Quick start

`ash
# Convert documents
python scripts/converter_para_md.py

# Extract bugs to Taiga JSON
python scripts/extrair_bugs.py Relatorio-Testes-Manuais.md taiga-import.json

# Generate a test plan
python scripts/gerar-plano-testes-generico-xlsx.py
`

## License

[MIT](LICENSE) © Gabriel Vidal