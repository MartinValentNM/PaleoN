# PaleoN — exceptionally detailed user and administrator manual

**Application:** PaleoN  
**Documentation version:** 2026-07-08  
**Application version:** 2.3  
**Author:** Martin Valent, Národní muzeum Praha — Oddělení paleontologie  
**Project:** DKRVO 2024–2028/2.I.c (NM 00023272)  
**Recommended repository license:** MIT  

> This work was supported by the Ministry of Culture of the Czech Republic through project DKRVO 2024–2028/2.I.c (NM 00023272).

---

## 1. Purpose of the application

PaleoN is a local multi-user application for conservative extraction, review, editing, and export of palaeontological taxonomic records from scientific documents. It is intended for cases where OCR or simple full-text search is not sufficient: the goal is to transform difficult, historically variable taxonomic prose into controlled records while preserving a link to the source document, source page, context, and verbatim taxonomic block.

The application is designed for palaeontologists, taxonomists, collection curators, and digitisation staff. It imports PDF, DOCX, and TXT files, stores page text in SQLite, detects taxon candidates, supports manual review before approval, maps taxonomic text to canonical fields, creates Taxon Dossiers and Morphology/Stratigraphy Dossiers, and exports results for scientific or museum use.

The key philosophy of PaleoN is **review-first**. A detected candidate is not a final record. It is a proposal that must be approved, rejected, or marked for additional review by a human user.

---

## 2. What PaleoN can and cannot do

### 2.1 PaleoN can

- import PDF, DOCX, and TXT documents into a local library;
- extract text and optionally use OCR depending on installed components;
- handle two-column PDF layouts and filter common boilerplate;
- detect taxonomic candidates in Latin-script, Cyrillic, and Chinese systematic headings;
- manage candidate states: `pending`, `approved`, `rejected`, `needs_review`;
- preserve the original taxonomic block as verbatim source text;
- map sections such as Diagnosis, Description, Synonymy, Type material, Locality, Stratigraphy, Remarks, and others to canonical fields;
- maintain editable dictionaries: section schema, taxonomic gazetteer, morphology terms, stratigraphy terms, and systematic-section headings;
- create a Taxon Dossier across the entire local library;
- create a Morphology/Stratigraphy Dossier based on controlled terms;
- export data to human-readable and tabular formats;
- optionally use local LLM assistance via LM Studio;
- separate data, settings, and dictionaries by user.

### 2.2 PaleoN should not be used as

- an automatic producer of taxonomic claims without expert review;
- a replacement for taxonomic revision;
- a guarantee of perfect OCR from poor scans;
- an online cloud workflow unless the user explicitly creates such a workflow outside the application;
- a general PDF viewer or full bibliographic manager.

---

## 3. Technical architecture

PaleoN is a local Streamlit application. The usual command is:

```bash
streamlit run paleon_st.py
```

Main components:

| Area | Technology / file | Purpose |
|---|---|---|
| User interface | Streamlit | Local browser-based interface |
| Database | SQLite | Documents, pages, candidates, fields, term matches |
| Text extraction | PyMuPDF, pdfplumber, python-docx | PDF/DOCX/TXT text |
| OCR | Tesseract, easyOCR, PyMuPDF OCR if available | Scanned pages or weak text layer |
| Tables | pandas, openpyxl | Previews and XLSX export |
| Documents | python-docx | DOCX exports and reports |
| LLM | LM Studio API | Optional local assistance |
| Application data | `paleon_data/` | Databases, uploads, dictionaries, exports, settings |

---

## 4. Installation

### 4.1 Recommended environment

Use Python 3.10+ or 3.11+, a dedicated virtual environment, and a local working directory outside system folders.

### 4.2 Create a virtual environment

Linux/macOS:

```bash
python -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
```

Windows PowerShell:

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
```

### 4.3 Install core dependencies

```bash
pip install streamlit pandas openpyxl python-docx pymupdf pdfplumber Pillow
```

### 4.4 Optional OCR dependencies

```bash
pip install pytesseract easyocr
```

`pytesseract` is only a Python wrapper. Full OCR usually requires a system Tesseract installation and language data. A typical language setting for multilingual palaeontological work is:

```text
eng+ces+rus+deu+fra+chi_sim
```

### 4.5 Start the application

```bash
streamlit run paleon_st.py
```

Open the URL printed by Streamlit, usually `http://localhost:8501`.

---

## 5. First launch and data directory

On first launch, PaleoN creates `paleon_data/`. This directory contains runtime data and should not be committed to a public repository.

Important paths:

| Path | Meaning |
|---|---|
| `paleon_data/paleon.db` | SQLite database in legacy/default mode |
| `paleon_data/users/` | Per-user data folders |
| `paleon_data/paleon_users.db` | User registry |
| `paleon_data/uploads/` | Uploaded files |
| `paleon_data/exports/` | Exported outputs |
| `paleon_data/settings.json` | Application settings |
| `paleon_data/prompts.json` | LLM prompts |
| `paleon_data/section_schema.tsv` | Section-label-to-field mapping |
| `paleon_data/taxons.txt` | Optional taxon gazetteer |
| `paleon_data/paleon_morphology_terms.tsv` | Morphology terms |
| `paleon_data/stratigraphy_terms.tsv` | Stratigraphy terms |
| `paleon_data/systematic_sections.tsv` | Systematic section headings |

---

## 6. Users and administration

PaleoN supports separated user profiles. If no user registry exists, a default `trial` account is created. Each user can have a separate database, dictionaries, and settings.

### 6.1 Switching users

Select the user in the sidebar or user-management area. Switching users invalidates schema and term caches so the correct per-user dictionaries are loaded.

### 6.2 Creating a new user

The administration area is protected by an admin password. The default admin password is:

```text
PaleoN
```

After unlocking, you can create a user, optionally assign a password, and grant admin rights. New user folders receive copies of the system dictionaries.

### 6.3 User-management recommendations

- Do not use `trial` for sensitive or long-term research projects.
- Create separate profiles for individual researchers or projects.
- Back up the database and TSV dictionaries before major changes.
- Never publish `paleon_data/` to GitHub.

---

## 7. Main interface panels

| Panel | Purpose |
|---|---|
| `Library` | Import documents, re-index, inspect pages, metadata, statistics |
| `Review` | Candidate review, approval, rejection, batch actions |
| `Editor` | Record, block, and field editing; manual record creation |
| `Taxon Dossier` | Search a taxon across the whole library |
| `Morpho/Strat.` | Search by morphology and stratigraphy terms |
| `Export` | Export approved data and reports |
| `Settings` | Schema, dictionaries, LLM, data, quality checks, users |

---

## 8. Recommended basic workflow

1. Prepare documents and verify that they can be processed legally and practically.
2. Start PaleoN.
3. Select the correct user/project.
4. Upload PDF/DOCX/TXT files in Library.
5. Set the document language.
6. Upload and index.
7. Inspect extracted page text and OCR quality.
8. Open Review.
9. Approve only real taxonomic treatments.
10. Mark uncertain items as `needs_review` or reject false positives.
11. Open Editor and inspect the RAW block and mapped fields.
12. Correct block boundaries and field mapping if needed.
13. Recompute Morphology/Stratigraphy terms.
14. Use Taxon Dossier to check names across the library.
15. Export DOCX/XLSX and perform expert review.

---

## 9. Document Library

### 9.1 Uploading documents

Use `Upload new documents` in the Library panel. Supported formats are PDF, DOCX, and TXT. Multiple files can be selected at once. Always set the document language; it helps with section labels, optional translation, and later interpretation.

### 9.2 Indexing

When you click `Upload and index`, PaleoN:

1. stores the file in the local library;
2. creates a document record in the database;
3. extracts page text;
4. uses OCR if enabled and the text layer is insufficient;
5. normalises text;
6. runs candidate detection;
7. stores candidates with context, score, and diagnostics.

### 9.3 Re-index

Use `Re-index` when OCR settings changed, dictionaries or systematic-section patterns changed, extraction quality was poor, or you want to regenerate pages and candidates using new rules. Re-indexing updates page text and increments an internal `pages_version` so cached text units are invalidated.

### 9.4 Reset detection

`Reset det.` removes and regenerates candidate detection for a document. Use it carefully if you already have approved or manually edited records.

### 9.5 Page viewer

Use the page viewer to verify whether extracted text matches the PDF, columns are not mixed, OCR has not skipped relevant lines, taxon headings are not broken, and headers/footers/captions have not polluted the taxonomic block.

---

## 10. OCR and text quality

### 10.1 When to enable OCR

Enable OCR if a PDF has no text layer or page text is extremely short. The threshold is controlled by `pdf_min_chars`.

### 10.2 OCR languages

Use only installed languages. Too many OCR languages may slow processing and sometimes reduce accuracy.

### 10.3 OCR verification

Always verify taxon names, authors, years, catalogue numbers, measurement units, Holotype/Paratype/Type species lines, stratigraphic units, and localities.

---

## 11. Candidate detection

The detector combines:

- taxon-name shape: genus, species, subspecies, family, order, class;
- rank labels such as Genus, Species, Family, Rod, Druh, Вид, 属;
- systematic-section context;
- following text context;
- strong section labels;
- stopword penalties;
- boilerplate filters;
- gazetteer bonus;
- penalty outside systematic regions;
- reference/index/table-of-contents zone detection.

The output is a scored candidate. The score is a review priority, not a truth value.

### 11.1 Detection thresholds

| Setting | Meaning |
|---|---|
| `taxon_min_confidence` | minimum score for normal candidates |
| `taxon_low_confidence` | low-confidence boundary |
| `outside_systematic_penalty` | penalty outside systematic sections |
| `gazetteer_bonus` | soft bonus for gazetteer match |

Lower thresholds increase sensitivity but create more false positives.

---

## 12. Candidate Review

Review is the key quality-control step.

### 12.1 Filters

Filter by document, status, rank, name, and whether low-confidence or rejected records should be shown.

### 12.2 Quick review

For every candidate, inspect:

1. taxon name;
2. rank;
3. page;
4. heading text;
5. context before and after;
6. RAW block;
7. score and debug information.

### 12.3 Candidate states

| State | Use when |
|---|---|
| `approved` | The item is a real taxonomic record and should enter further processing. |
| `rejected` | The item is a false positive, caption, reference, sentence, index entry, or OCR noise. |
| `needs_review` | The item may be correct but needs closer checking. |
| `pending` | Default state after detection. |

### 12.4 Batch actions

Use batch actions only after filtering. Never approve a whole document blindly. Batch approve is appropriate only for clear high-confidence groups; batch reject is useful for obvious boilerplate or index sections; batch LLM validation is advisory only.

---

## 13. Record Editor

The Editor is used for detailed correction of approved or draft records.

### 13.1 Selecting a record

Search by taxon and document. The Editor shows RAW block, context, fields, rank/name, diagnostics, and automap/LLM tools.

### 13.2 RAW block

The RAW block should contain the complete taxonomic treatment and not include the next taxon, bibliography, or unrelated captions. Keep it as verbatim as possible; edit mainly OCR errors or wrong boundaries.

### 13.3 Category annotation

Category annotation prepends markers such as `[DIAGNOSIS]` or `[DESCRIPTION]` to recognised paragraphs. This helps verify the mapping without summarising the source text.

### 13.4 Saving a manual block

Save a manual block when the parser ended too early, included the next taxon, OCR merged headings, or page layout broke reading order.

### 13.5 Automap fields

`Auto-map` maps section labels to canonical fields. Typical mappings include Diagnosis to DIAGNOSIS, Description to DESCRIPTION, Synonymy to SYNONYMY, Type material to TYPE MATERIAL, Holotype/Paratype to TYPE SPECIMENS, Type species to TYPE TAXON, Locality to LOCALITY, Stratigraphy to STRATIGRAPHY, Remarks to REMARKS, and Occurrence to OCCURRENCE.

### 13.6 Manual field editing

Preserve Latin names verbatim, verify authors and years, preserve catalogue-number formatting, avoid unsupported interpretation, and record uncertainty in notes rather than silently changing the taxon name.

### 13.7 Renaming a taxon

Rename only when detection or OCR clearly captured the name incorrectly. After renaming, check rank and synonym relationships.

---

## 14. Block Editor and boundaries

Block Editor addresses one of the hardest tasks: determining where a taxonomic treatment starts and ends.

### 14.1 Recommended process

1. Select document and taxon.
2. Inspect the automatic block.
3. If too short, add following text units.
4. If too long, remove trailing paragraphs or set end before the next heading.
5. Use boundary suggestions if available.
6. Save the manual block.
7. Re-extract fields.
8. Recompute Morphology/Stratigraphy terms.

### 14.2 Common boundary errors

- the block ends immediately after the taxon name;
- the block includes synonymy from the previous taxon;
- a figure caption is treated as a taxonomic treatment;
- a two-column PDF merges unrelated columns;
- OCR inserted page breaks inside paragraphs.

---

## 15. Manual record creation

If the parser does not find a taxon, create a record manually:

1. Open Editor.
2. Expand manual taxon-record creation.
3. Enter taxon name.
4. Select rank.
5. Paste RAW block if available.
6. Fill base fields.
7. Save.
8. Run automap or post-approval processing.
9. Verify the export row.

If no document is selected, the record is stored under the virtual document `Manual records`.

---

## 16. Taxon Dossier

Taxon Dossier searches for a taxon across the entire library.

### 16.1 How to use it

1. Open Taxon Dossier.
2. Enter part of a name.
3. Decide whether to include low-confidence/rejected items.
4. Inspect occurrences.
5. Open record details.
6. Check fields, context, and full block.
7. Export TXT for a quick dossier if needed.

### 16.2 Including rejected or low-confidence items

Enable this during exploratory research to see whether a taxon was detected but rejected. For final datasets, rely primarily on approved records.

---

## 17. Morphology/Stratigraphy Dossier

This module searches by terms rather than taxon names.

### 17.1 Morphology

Morphological terms are primarily searched in DESCRIPTION and DIAGNOSIS. If those fields are not filled, the RAW block may act as fallback.

### 17.2 Stratigraphy

Stratigraphic terms are primarily searched in STRATIGRAPHY. The application also recognises named units such as Formation, Member, Stage, souvrství, svita, horizon, 组, 段, etc.

### 17.3 Recomputing terms

After editing dictionaries or fields, use `Recompute terms`. This refreshes stored `term_matches` for approved records.

---

## 18. Export

Exports support further checking and downstream work. Typical formats:

- DOCX for readable reports;
- XLSX for table review and import;
- TXT for quick text inspection;
- JSON for machine processing where available;
- PDF/HTML print output depending on configuration.

Before exporting, check approved candidates, RAW blocks, key fields, type material, authors/years, locality, stratigraphy, duplicates, and synonyms.

---

## 19. Settings — schema, gazetteer, and dictionaries

### 19.1 Section schema

`section_schema.tsv` defines how source labels map to canonical fields. Typical columns are `enabled`, `label`, `label_regex`, `canonical_label`, `target_field`, `is_strong`, `can_start_treatment`, `type_species_standalone`, `priority`, and `language`.

Recommendations: add precise rather than overly general labels, use high priority for unambiguous labels, specify language, and re-index or re-map fields after schema changes.

### 19.2 Gazetteer

`taxons.txt` is an optional list of known taxa. Absence from the gazetteer is not evidence that a candidate is wrong. It provides a soft bonus, not a hard filter.

### 19.3 Morphology terms

`paleon_morphology_terms.tsv` supports morphology searches. Recommended columns: `term`, `canonical`, `category`, `language`, and `enabled`.

### 19.4 Stratigraphy terms

`stratigraphy_terms.tsv` supports stratigraphic search. Keep canonical names and categories consistent across documents.

### 19.5 Systematic sections

`systematic_sections.tsv` stores headings such as Systematic palaeontology, Taxonomy, Systematika, Систематика, 系统古生物学. Good systematic-section settings reduce false positives.

---

## 20. LLM / LM Studio mode

LLM support is optional and disabled by default.

### 20.1 Enabling it

In Settings, configure `llm_enabled`, `lmstudio_base_url` (usually `http://localhost:1234/v1`), model, timeout, and temperature. A temperature of `0.0` is recommended.

### 20.2 LLM use cases

- candidate validation;
- block-boundary assessment;
- JSON field mapping;
- translation of field values into English;
- assistance with Chinese or Russian blocks when rules are insufficient.

### 20.3 Safe LLM rules

LLM output is advisory, must not invent missing data, should work only from verbatim text, must return valid JSON, and must be checked by the user. For sensitive documents, use only local models and a local server.

---

## 21. Data quality and validation

### 21.1 Completeness score

Completeness depends on rank and filled fields. A low score is not automatically an error if the source lacks the information, but it signals a need for review.

### 21.2 Fuzzy duplicates

Fuzzy duplicate checks help find OCR variants, transliteration differences, missing diacritics, author spelling variants, and repeated taxa across documents.

### 21.3 Author OCR typo checks

Author-name checks help identify variants such as Sysoev/Sysoiev or OCR mistakes. Always verify against the source.

### 21.4 Gold set evaluator

Use a gold set to compare parser output against reference records: prepare a reference DOCX, link it to a document, run the benchmark, and evaluate field-mapping accuracy.

---

## 22. Backup and restore

Back up per-user `paleon.db`, `paleon_users.db`, user TSV dictionaries, `settings.json`, and final exports. Before major changes, close the application, copy the whole `paleon_data/` folder to a dated archive, perform changes, restart, check the library, and restore from backup if needed.

---

## 23. Recommended professional standard

For institutional or publishable outputs:

- each document should have bibliographic metadata;
- each candidate should be manually approved;
- RAW blocks should be checked against the PDF;
- taxon name and rank should be verified;
- type material should be checked for species/subspecies;
- TYPE TAXON should be checked for higher taxa where relevant;
- locality and stratigraphy should be verified;
- XLSX should be reviewed in a spreadsheet;
- DOCX should serve as a readable appendix;
- local data should not be published without rights and privacy review.

---

## 24. Troubleshooting

### 24.1 Application does not start

Check that the virtual environment is active, Streamlit is installed, the file is `paleon_st.py`, and the terminal is in the application directory.

### 24.2 PDF has no text

Enable OCR, verify Tesseract installation, set correct languages, improve scan quality if possible, and inspect page text in the page viewer.

### 24.3 Too many false positives

Raise `taxon_min_confidence`, check systematic-section settings, add stopwords or schema rules, hide low-confidence items, and batch-reject only after filtering.

### 24.4 A taxon was not detected

Verify that OCR contains the heading, temporarily lower thresholds, add systematic headings or rank-label variants, or create the record manually.

### 24.5 Fields do not map correctly

Check section labels in the RAW block, add missing labels to `section_schema.tsv`, run Auto-map again, manually remap, or add language-specific label variants.

### 24.6 LM Studio does not respond

Check that LM Studio is running, base URL is correct, a model is loaded, timeout is sufficient, or disable LLM if not needed.

---

## 25. Recommended repository contents

Recommended GitHub contents:

- `paleon_st.py`;
- `README.md`;
- `README.cs.md`;
- `LICENSE`;
- `NOTICE`;
- `CITATION.cff`;
- `requirements.txt`;
- `requirements-optional.txt`;
- `.gitignore`;
- `CONTRIBUTING.md`;
- `SECURITY.md`;
- documentation and this manual.

Do not commit `paleon_data/`, databases, uploaded copyrighted PDFs/DOCX, internal exports with sensitive data, passwords, tokens, or private configuration.

---

## 26. Pre-export checklist

- [ ] Document imported correctly.
- [ ] OCR/page text visually checked.
- [ ] Candidates reviewed.
- [ ] False positives rejected.
- [ ] RAW blocks for important taxa complete.
- [ ] Fields automapped or manually filled.
- [ ] Type material verified.
- [ ] Locality and stratigraphy verified.
- [ ] Morpho/Strat matches recomputed.
- [ ] XLSX/DOCX export opened and checked.
- [ ] Results backed up.

---

## 27. Glossary

| Term | Meaning |
|---|---|
| Candidate | Automatically detected possible taxonomic record |
| Approved | Reviewed candidate used in exports |
| RAW block | Original taxonomic-treatment text block |
| Occurrence fields | Canonical fields extracted from a record |
| Gazetteer | Auxiliary list of known taxa |
| Schema | TSV mapping from section labels to fields |
| Morpho/Strat terms | Dictionaries of morphology and stratigraphy terms |
| LLM | Local language model via LM Studio |
| Gold set | Reference set for accuracy evaluation |

---

## 28. Conclusion

PaleoN is a practical, auditable, and conservative research tool. It performs best when automatic detection is combined with expert review, careful RAW-block checking, dictionary maintenance, and regular export to both readable and tabular formats.
