# PaleoN

**PaleoN** is a local Streamlit application for conservative multilingual extraction, review, editing, and export of palaeontological taxonomic records from PDF, DOCX, and TXT documents.

- **Author:** Martin Valent, Národní Muzeum Praha — Oddělení Paleontologie
- **Project:** DKRVO 2024–2028/2.I.c (NM 00023272)
- **License:** MIT

## Key features

- Local Streamlit interface with SQLite backend.
- Multi-user data folders and per-user dictionaries.
- Conservative taxon candidate detection with manual review before export.
- Taxon Dossier search across the whole local library.
- Morphology and Stratigraphy dossier based on editable TSV dictionaries.
- Input formats: PDF, DOCX, TXT.
- Optional OCR support through Tesseract/easyOCR/PyMuPDF, depending on local installation.
- Optional LM Studio integration for assisted validation, field mapping, and translation.
- Exports for structured tabular data and human-readable reports.

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run paleon_st.py
```

The application creates its working data directory automatically as `paleon_data/`.

## Optional OCR setup

For OCR, install the Python packages listed in `requirements-optional.txt`. Tesseract OCR may also require a system installation and language data files, depending on your operating system.

## LM Studio integration

LLM assistance is optional and disabled by default. If used, run LM Studio locally and configure the base URL and model in the application settings. PaleoN is designed so that LLM output is advisory only; users should review and approve taxonomic records manually.

## Repository contents

- `paleon_st.py` — main Streamlit application.
- `requirements.txt` — core Python dependencies.
- `requirements-optional.txt` — optional OCR/extra export dependencies.
- `README.md` — English project documentation.
- `README.cs.md` — Czech project documentation.
- `LICENSE` — MIT license.
- `CITATION.cff` — citation metadata.
- `CONTRIBUTING.md` — contribution guidelines.
- `SECURITY.md` — security policy.
- `.gitignore` — ignores local runtime data and generated files.

## Data and privacy

PaleoN is intended as a local research tool. Uploaded documents, generated databases, settings, exports, and user dictionaries are stored locally under `paleon_data/` unless you deliberately move or publish them. Do **not** commit `paleon_data/` to GitHub.

## Acknowledgement

This work was supported by the Ministry of Culture of the Czech Republic through project DKRVO 2024–2028/2.I.c (NM 00023272).

## License

This project is licensed under the MIT License. See [`LICENSE`](LICENSE).
