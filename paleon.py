#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import hashlib
import hmac
import io, json, logging, os, pathlib, re, sqlite3, threading, zipfile
import difflib, shutil, unicodedata
import urllib.request, urllib.error
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

# ── optional deps ──────────────────────────────────────────────────────────────
try:
    import streamlit as st
except ImportError:
    st = None  # allows syntax-checking outside Streamlit

try:
    from docx import Document as _DocxDoc
    from docx.shared import Pt
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

try:
    import fitz          # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import pytesseract
    from PIL import Image as _PILImage
    HAS_TESSERACT = True
except ImportError:
    HAS_TESSERACT = False

try:
    import easyocr as _easyocr
    HAS_EASYOCR = True
except ImportError:
    HAS_EASYOCR = False

try:
    from PIL import Image as _PILImage   # muže být dostupné i bez pytesseract
except ImportError:
    pass

# Souhrnný flag: máme alespoň jeden OCR engine?
# fitz_ocr funguje bez pytesseract binárky jen v PyMuPDF >= 1.19 se systémovým tesseractem,
# ale easyocr funguje čistě pythonovsky (pip install easyocr).
def _HAS_ANY_OCR() -> bool:
    """Returns True if at least one OCR engine is available."""
    return HAS_EASYOCR or HAS_TESSERACT or HAS_FITZ

try:
    from fpdf import FPDF as _FPDF
    HAS_FPDF = True
except ImportError:
    HAS_FPDF = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ══════════════════════════════════════════════════════════════════════════════
# KONSTANTY
# ══════════════════════════════════════════════════════════════════════════════
APP_NAME    = "PaleoN - Taxon Extraktor"
APP_VERSION = "1.0"
NOT_PROVIDED = "Not provided"
ILLEGIBLE    = "[illegible]"
MORPHO_STRAT_MATCHER_VERSION = "legacy_regex_from_claude-morfph"

# Kanonický seznam klíčových slov pro typové exempláře (holotyp/paratyp/…)
# ve allch podporovaných jazycích. Používá se pro detekci, zda text obsahuje
# zmínku o typovém materiálu (TYPE SPECIMENS ≡ druh/poddruh). Centralizováno,
# aby se stejný seznam neduplikoval na 5 místech s rizikem nekonzistence.
TYPE_SPECIMEN_KEYWORDS: Tuple[str, ...] = (
    # EN
    "holotype", "paratype", "lectotype", "paralectotype", "neotype",
    "syntype", "topotype", "allotype",
    # DE (-typus)
    "holotypus", "paratypus", "lectotypus", "syntypus", "neotypus",
    # CS (-typ / -typy)
    "holotyp", "paratyp", "lektotyp", "syntyp", "neotyp",
    "holotypy", "paratypy", "syntypy",
    # RU
    "голотип", "паратип", "лектотип", "неотип", "синтип", "паралектотип",
    # ZH
    "正模", "副模", "选模", "新模", "地模", "模式标本",
)

# ══════════════════════════════════════════════════════════════════════════════
# ENGLISH-ONLY UI TEXT
# ══════════════════════════════════════════════════════════════════════════════

# One-language build: only English UI labels are kept. The multilingual parser
# dictionaries below remain intact because they are used to read source documents.
TRANSLATIONS: Dict[str, str] = {'tab_library': '📚 Library',
 'tab_review': '🔍 Review',
 'tab_editor': '✏️ Editor',
 'tab_dossier': '🧬 Taxon Dossier',
 'tab_morpho': '⏳🔬 Morpho/Strat.',
 'tab_export': '📤 Export',
 'tab_settings': '⚙️ Settings',
 'dark_mode': '🌙 Dark mode',
 'docs': 'Documents',
 'approved': 'Approved',
 'pending': 'Pending',
 'low_conf': 'Low-conf',
 'save_settings': '💾 Save settings',
 'close_app': '🛑 Close app',
 'confirm_close': 'Really close the application?',
 'yes': 'Yes, close',
 'cancel': 'Cancel',
 'library_header': '📚 Document Library',
 'upload_new': '➕ Upload new documents',
 'upload_hint': 'PDF / DOCX / TXT — multiple files supported',
 'upload_lang': 'Document language',
 'upload_btn': '📥 Upload and index',
 'no_docs': 'Library is empty. Upload your first document above.',
 'doc_actions': 'Document actions',
 'reindex': '🔄 Re-index',
 'reset_det': '♻️ Reset det.',
 'export_txt': '📄 TXT pages',
 'delete_doc': '🗑️ Delete',
 'page_viewer': '🔍 Page viewer',
 'stats_header': '📊 Library-wide statistics',
 'ranks_chart': 'Taxa by rank',
 'strat_chart': 'Taxa by stratigraphic period',
 'timeline_chart': 'Timeline — document uploads and candidate detection',
 'review_header': '🔍 Candidate Review',
 'all_docs': 'All documents',
 'filter_status': 'Status',
 'filter_rank': 'Rank',
 'filter_doc': 'Document',
 'filter_name': 'Search name',
 'quick_review': '🎯 Quick review',
 'approve': '✅ Approve (A)',
 'reject': '❌ Reject (R)',
 'needs_review': '🔍 Needs review',
 'prev': '⬅️ Prev.',
 'next': 'Noxt ➡️',
 'batch_approve': '✅ Approve',
 'batch_reject': '❌ Reject',
 'batch_automap': '🔄 Auto-map',
 'batch_llm': '🤖 LLM',
 'no_candidates': 'No candidates to display.',
 'show_rejected': 'Show rejected',
 'editor_header': '✏️ Record Editor',
 'no_approved': 'No approved records. Go to Review first.',
 'search_taxon': '🔍 Search taxon',
 'search_placeholder': 'Gracilitheca, astronauta…',
 'all_docs_opt': '— All —',
 'raw_block': '📜 RAW block (verbatim)',
 'annotate_toggle': '🏷️ Categories',
 'save_block': '💾 Save block',
 'automap': '🔄 Auto-map',
 'llm_fields': '🤖 LLM fields',
 'context_label': '📝 Context',
 'fields_header': '📋 Record fields',
 'rename_header': '✏️ Rename taxon',
 'new_name': 'Now name',
 'save_name': '💾 Save name/rank',
 'scoring_debug': '🔬 Scoring debug',
 'filled_fields': 'Filled: {f}/{t} fields ({p} %)',
 'block_saved': 'Block saved.',
 'name_saved': 'Saved.',
 'no_labels': 'No labels found.',
 'dossier_header': '🧬 Taxon Dossier',
 'dossier_caption': 'Find a taxon name across the ENTIRE library.',
 'dossier_query': '🔍 Taxon name',
 'dossier_placeholder': 'Gracilitheca, astronauta, Nophrotheca sophia…',
 'include_low': '+ low-conf/rej.',
 'annotate_cats': '🏷️ Categories',
 'dossier_tips': '💡 Hints: ',
 'dossier_hint': 'Enter at least 2 characters.',
 'dossier_not_found': 'No occurrences found.',
 'occurrences': 'Occurrences',
 'distinct_names': 'Distinct names',
 'documents': 'Documents',
 'full_record': '↕\nDetail',
 'full_record_title': 'full record',
 'no_fields': 'No filled fields.',
 'full_block': 'Full text block:',
 'ctx_before': 'Context before:',
 'ctx_after': 'Context after:',
 'export_txt_btn': '📥 Export as TXT',
 'morpho_header': '⏳ Morphology & Stratigraphy Dossier',
 'morpho_caption': 'Filter taxa by morphological or stratigraphic terms.',
 'stratigraphy': '⏳ Stratigraphy',
 'morphology': '🐚 Morphology',
 'category': 'Category',
 'term': '🔍 Term',
 'filter_names': '🔍 Filter names',
 'recompute': '🔄 Recompute terms',
 'no_terms': 'No terms computed yet.',
 'export_header': '📤 Data Export',
 'settings_header': '⚙️ Settings',
 'schema_tab': '📋 Schema',
 'gazetteer_tab': '🧬 Gazetteer',
 'terms_tab': '⏳🔬 Terms',
 'llm_tab': '🤖 LLM prompts',
 'data_tab': '💾 Data',
 'quality_tab': '🧹 Data quality',
 'empty_library': 'Library is empty.',
 'record': 'record',
 'records': 'records',
 'page': 'page',
 'confidence': 'Score',
 'status': 'Status',
 'rank': 'Rank',
 'document': 'Document',
 'needs_review_btn': '🔍 Needs review',
 'show_ctx': 'Context',
 'auto_ocr': 'Auto OCR',
 're_index': '🔄 Re-index',
 'save_fields': '💾 Save fields',
 'field_value': 'Field value',
 'type_taxon': 'Type taxon',
 'included_taxa': 'Included taxa',
 'type_specimens': 'Type specimens',
 'material_examined': 'Examined material',
 'stratigraphy_field': 'Stratigraphy',
 'locality_field': 'Locality',
 'occurrence_field': 'Occurrence',
 'sort_order': 'Sort order',
 'batch_threshold': 'Batch approval threshold',
 'show_context': 'Context',
 'copy_block': '📋 Copy block',
 'auto_map_all': '🔄 Auto-map all',
 'detail_title': 'Candidate detail',
 'pdf_preview_requires_pymupdf': 'ℹ️ PDF preview requires `pip install pymupdf`.',
 'file_path_not_in_db': 'ℹ️ File path not stored in DB.',
 'preview_pdf_only': 'ℹ️ Page preview is only available for PDF.',
 'pdf_page_requires_pymupdf': 'ℹ️ PDF page preview requires PyMuPDF (`pip install pymupdf`).',
 'preview_pdf_docs_only': 'ℹ️ Page preview is only available for PDF documents.',
 'pdf_page_real': '📄 Actual PDF page (for verifying OCR errors)',
 'page_preview_failed': 'Page preview could not be rendered.',
 'page_text_label': 'Page text',
 'redetect_lmstudio': '🔁 Re-detect LM Studio',
 'llm_disabled_caption': 'LLM assistant is disabled.',
 'indexing_settings': '⚙️ Indexing settings',
 'ocr_caption': '**📄 OCR**',
 'candidate_detection_caption': '**🔧 Candidate detection**',
 'save_caption': '**💾 Save**',
 'settings_saved_ok': '✓ Settings saved',
 'ocr_no_engine': '❌ no engine',
 'doc_notes_save': '💾 Save note',
 'biblio_metadata': '📚 Bibliographic metadata',
 'biblio_caption': 'Optional — facilitates citation and DwC-A export.',
 'save_metadata_btn': '💾 Save metadata',
 'metadata_saved': 'Metadata saved.',
 'ocr_running': 'OCR running — may take a while…',
 'no_pages_reindex': 'No pages — please re-index.',
 'no_pages': 'No pages.',
 'lib_stats': '📊 Library-wide statistics',
 'no_candidates_yet': 'No candidates yet.',
 'taxa_by_strat': '**Taxa by stratigraphic period**',
 'timeline_docs_candidates': '**Timeline — document uploads and candidate detection**',
 'docs_per_day': 'Documents uploaded / day',
 'candidates_per_day': 'Candidates detected / day',
 'no_timeline_data': 'Not enough data for timeline yet.',
 'upload_doc_first': 'Upload a document first in the Library tab.',
 'filters_expander': '🔽 Filters',
 'no_candidates_filter': 'No candidates for current filter.',
 'batch_translate_expander': '🌐 Batch document translation',
 'no_non_en_docs': 'No non-English documents with approved records.',
 'no_approved_in_filter': 'No approved records in filter.',
 'no_changes_to_save': 'No changes to save.',
 'no_approved_in_sel': 'No approved records in selection.',
 'candidate_detail_expander': '🔎 Candidate detail',
 'context_before_caption': 'Context before:',
 'block_text_caption': 'Block text:',
 'context_after_caption': 'Context after:',
 'extracted_fields_caption': 'Extracted fields:',
 'block_editor_help': '❓ How Block Editor works — click for help',
 'select_at_least_one_status': 'Select at least one status.',
 'candidate_block_label': 'Candidate / taxonomic block',
 'candidate_not_found': 'Candidate not found.',
 'records_in_doc': '① Records in document',
 'all_taxa_found': 'All found taxa. Select record above.',
 'suggest_boundaries': '🧠 Suggest boundaries',
 'apply_boundary_suggestion': '✅ Apply boundary suggestion',
 'boundary_suggestion_applied': 'Boundary suggestion applied.',
 'text_boundaries': '② Text boundaries in document',
 'end_before_start': 'End unit is before start unit; adjust the range.',
 'boundaries_updated': 'Boundaries and parser block updated according to units.',
 'block_text_fields': '③ Block text and extracted fields',
 'switch_to_source': '✅ Switch to selected source',
 'auto_block_expander': '📄 Auto block (parser output)',
 'manual_block_saved': 'Manual block saved and activated.',
 'field_extraction_done': 'Field extraction complete.',
 'detected_sections_debug': '🔍 Detected sections in block (debug)',
 'extracted_fields_md': '**📋 Extracted fields**',
 'extracted_fields_hint': 'Field values (Description, Occurrence, Type specimens etc.) extracted from block.',
 'no_fields_yet': 'No fields yet — press 🔄 Re-extract fields from block.',
 'create_record_expander': '➕ Create new taxon record manually',
 'create_record_caption': 'Creates a new taxon in DB including RAW block, fields, FTS and Morpho/Strat terms.',
 'raw_block_placeholder': 'Paste the full taxonomic block here. Empty = minimal block created from fields.',
 'optional_base_fields': '**Optional base fields**',
 'no_record_in_filter': 'No record matches the filter.',
 'record_not_found': 'Record not found.',
 'edit_name_assign': '✏️ Edit name / assign to correct taxon',
 'save_name_rank_btn': '💾 Save name and rank',
 'assign_to_taxon_caption': '**Assign to existing taxon** (records will be marked as synonyms):',
 'prev_unit_added': 'Previous unit added.',
 'no_prev_unit': 'No previous unit.',
 'first_para_removed': 'First paragraph removed.',
 'only_one_para': 'Block has only one paragraph.',
 'last_para_removed': 'Last paragraph removed.',
 'next_unit_added': 'Noxt unit added.',
 'no_next_unit': 'No next unit.',
 'block_saved_ok2': 'Block saved.',
 'no_labels_found2': 'No labels found.',
 'llm_assigning_fields': 'LLM assigning fields…',
 'llm_no_valid_json': 'LLM did not return valid JSON with fields.',
 'translate_fields_expander': '🌐 Translate fields to English',
 'doc_is_english': 'Document is in English — translation probably not needed.',
 'lmstudio_translating': 'LM Studio translating…',
 'field_mapper_expander': '🗺️ Field mapper — detected sections',
 'save_remap_btn': '💾 Save remapping',
 'no_changes_remap': 'No changes.',
 'no_sections_in_block': 'No sections detected in current block. ',
 'load_block_first': 'Extract/load block text first.',
 'context_subheader': '📝 Context',
 'context_before_after': 'Context before / after',
 'context_before_ta': 'Context before',
 'context_after_ta': 'Context after',
 'no_docs_dossier': 'No documents.',
 'no_doc_matches': 'No document matches the search.',
 'check_at_least_one': 'Check at least one document.',
 'no_records_sel_filter': 'No records for current selection/filter.',
 'no_refs_match': 'No references match the search.',
 'formats_label': '**Formats**',
 'open_print_save': 'Open in browser → Ctrl+P → Save as PDF.',
 'show_table_preview': '📊 Show table preview',
 'show_field_stats': '📈 Field completeness statistics',
 'no_placement': 'Placement not filled — complete the TAXONOMIC PLACEMENT field.',
 'no_terms_computed2': 'No terms computed yet. Click 🔄 Recompute.',
 'no_term_matches': 'No term matches the filter.',
 'no_match_for_term': 'No matches for this term.',
 'label_dist_expander': '📊 Label distribution',
 'labels_by_field': 'Label count by field',
 'labels_by_lang': 'Label count by language',
 'label_editor_sub': '✏️ Label editor',
 'add_label_expander': '➕ Add new label',
 'add_label_btn': '➕ Add label',
 'label_text_required': 'Enter label text.',
 'gaz_save_btn': '💾 Save and apply new gazetteer',
 'gaz_preview': 'Gazetteer preview (first 60 genera)',
 'morpho_sub': '🐚 Morphology',
 'save_morpho_btn': '💾 Save morphology terms',
 'terms_preview': 'Terms preview',
 'strat_sub': '⏳ Stratigraphy',
 'save_strat_btn': '💾 Save stratigraphy terms',
 'sys_sections_sub': '📋 Systematic sections',
 'sys_sections_edit': 'Editable headings table',
 'save_sys_sections_btn': '💾 Save systematic sections',
 'upload_file_btn': '💾 Upload file',
 'roundtrip_sub': '📋 Round-trip XLSX editor',
 'roundtrip_caption': 'Export records to XLSX, edit in Excel, re-import to update the database. Gray columns (id, document, page) are read-only.',
 'export_edit_xlsx_btn': '⬇️ Export for editing (XLSX)',
 'import_xlsx_label': '📂 Upload modified XLSX',
 'import_xlsx_btn': '📥 Import and update records',
 'import_ok': '✅ Updated {n} records, {f} fields changed.',
 'import_skipped': 'Skipped (no changes): {s}.',
 'import_errors': 'Warnings: {e}',
 'import_nothing': 'No changes detected in the file.',
 'import_no_sheet': 'Sheet "Records" not found — upload a file exported by this application.',
 'fts_label': '**🔍 FTS5 index**',
 'indexing_records': 'Indexing records…',
 'db_sub': '🗄️ Database',
 'backup_sub': '🧰 Backup / restore database',
 'restore_db_btn': '♻️ Restore database from backup',
 'restore_confirm_warn': 'Really overwrite database with uploaded file? This cannot be undone.',
 'db_restored': 'Database restored from backup.',
 'db_not_exist': 'Database does not exist yet.',
 'fuzzy_dup_sub': '🧬 Fuzzy duplicate taxa across documents',
 'searching_similar': 'Searching similar names…',
 'no_dup_pairs': 'No suspicious pairs at selected threshold.',
 'ocr_typos_sub': '✒️ OCR typos in author names',
 'comparing_authors': 'Comparing author names…',
 'no_author_typos': 'No suspicious name pairs at selected threshold.',
 'gold_set_sub': '🏆 Gold Set Evaluator',
 'upload_gold_docx': '**1) Upload gold set DOCX**',
 'pair_with_lib': '**2) Match with document in library**',
 'save_evaluate': '**3) Save and evaluate**',
 'gold_sets_saved_hdr': '**Saved gold sets**',
 'run_benchmark': 'Running benchmark…',
 'mapping_accuracy': '📊 Field mapping accuracy',
 'record_detail_exp': '📋 Record detail',
 'no_gold_sets': 'No gold sets in DB. Upload a DOCX above.',
 'user_mgmt_sub': '👥 User management',
 'admin_required': 'User management requires admin password.',
 'wrong_password_admin': 'Wrong password. (Hint: PaleoN)',
 'admin_unlocked': '✅ Admin access unlocked',
 'save_password_btn': 'Save password',
 'dicts_expander': '📚 Dictionaries',
 'add_user_sub': '➕ Add user',
 'admin_rights': 'Admin rights',
 'create_btn2': '✅ Create',
 'name_required': 'Enter a name.',
 'username_invalid': 'Name must be 2–40 characters: letters, digits, _ or -',
 'update_sys_dicts': '🔄 Update system dictionaries',
 'publish_dicts_btn': '📤 Publish current dictionaries to system/',
 'similar_taxa_sub': '🔀 Similar taxa across documents',
 'comparing_taxa': 'Comparing taxa…',
 'select_user_md': '### Select user',
 'new_user_expander': '➕ Now user',
 'unique_username': 'Choose a unique username.',
 'create_account_btn': '✅ Create account',
 'wrong_password_simple': 'Wrong password.',
 'name_letters_only': 'Name must contain only letters, digits, _ or - (2–40 characters).',
 'preset_load': '📋 Load preset',
 'enter_name_warning': 'Enter a name.'}

_FIELD_LABEL_UI: Dict[str, str] = {'TAXONOMIC PLACEMENT': 'Taxonomic placement',
 'TAXON': 'Taxon',
 'NOMENCLATURAL ACTS': 'Nomenclatural acts',
 'TYPE TAXON': 'Type taxon',
 'INCLUDED TAXONS': 'Included taxa',
 'TYPE SPECIMENS': 'Type specimens',
 'AUTHOR': 'Author',
 'FIGURES': 'Figures',
 'SYNONYMY': 'Synonymy',
 'TYPE MATERIAL': 'Type material',
 'MATERIAL EXAMINED': 'Material examined',
 'ETYMOLOGY': 'Etymology',
 'LOCALITY': 'Locality',
 'STRATIGRAPHY': 'Stratigraphy',
 'DESCRIPTION': 'Description',
 'DIAGNOSIS': 'Diagnosis',
 'SIZE': 'Size',
 'REMARKS': 'Remarks',
 'OCCURRENCE': 'Occurrence',
 'REFERENCE': 'Reference'}


def t(key: str, **kwargs) -> str:
    """Return an English UI string. Missing keys fall back to the key itself."""
    text = TRANSLATIONS.get(key, key)
    return text.format(**kwargs) if kwargs else text


def tt(cs: str, en: str) -> str:
    """Compatibility helper for old bilingual calls; always returns English."""
    return en


def _current_lang() -> str:
    return "en"


def _ensure_i18n_completeness() -> None:
    return None


def _tr_auto(text_value: Any) -> Any:
    """English-only build: no runtime language switching/translation patching."""
    return text_value


def _field_label_for_ui(label: str) -> str:
    if not isinstance(label, str):
        return label
    return _FIELD_LABEL_UI.get(label, label)


def _localize_format_func(existing_func=None):
    def _fmt(x):
        return existing_func(x) if existing_func else x
    return _fmt


def _localize_dataframe_for_display(data: Any) -> Any:
    return data


def _install_streamlit_i18n_patch() -> None:
    """No-op in the English-only build."""
    return None

BASE_DIR    = pathlib.Path("paleon_data")
SCHEMA_FILE = BASE_DIR / "section_schema.tsv"
DB_FILE     = BASE_DIR / "paleon.db"
EXPORTS_DIR = BASE_DIR / "exports"
UPLOADS_DIR  = BASE_DIR / "uploads"
SETTINGS_FILE = BASE_DIR / "settings.json"
PROMPTS_FILE  = BASE_DIR / "prompts.json"
GAZETTEER_FILE = BASE_DIR / "taxons.txt"
MORPHOLOGY_FILE = BASE_DIR / "paleon_morphology_terms.tsv"
STRATIGRAPHY_FILE = BASE_DIR / "stratigraphy_terms.tsv"
SYSTEMATIC_SECTIONS_FILE = BASE_DIR / "systematic_sections.tsv"

# ── Multi-user paths ──────────────────────────────────────────────────────
USERS_DIR      = BASE_DIR / "users"          # paleon_data/users/<username>/
SYSTEM_DIR     = BASE_DIR / "system"         # systémové slovníky (šablona)
USERS_REGISTRY = BASE_DIR / "paleon_users.db"  # registr uživatelu
ADMIN_PASS_HASH = hashlib.sha256("PaleoN".encode()).hexdigest()


def _sanitize_username(name: str) -> str:
    """Safe directory name derived from a username."""
    return re.sub(r"[^a-zA-Z0-9_\-]", "_", (name or "").strip())[:40].lower()


def _user_dir(username: str = None) -> pathlib.Path:
    """Returns the data directory of the current (or specified) user."""
    u = username or (st.session_state.get("pn_user") if st is not None else None)
    if not u:
        return BASE_DIR  # fallback na legacy cesty
    return USERS_DIR / _sanitize_username(u)


def _upath(filename: str, username: str = None) -> pathlib.Path:
    """Shortcut: path to a file in the given user's directory."""
    return _user_dir(username) / filename


# ── User registry DB ──────────────────────────────────────────────────────

def _users_db() -> sqlite3.Connection:
    """Connection to the user registry (global, not per-user)."""
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(USERS_REGISTRY), check_same_thread=False)
    con.row_factory = sqlite3.Row
    return con


def _init_user_registry() -> None:
    """Creates the users table and trial account (if they do not exist)."""
    con = _users_db()
    con.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL COLLATE NOCASE,
            display_name TEXT DEFAULT '',
            password_hash TEXT DEFAULT NULL,
            is_admin INTEGER DEFAULT 0,
            created_at TEXT,
            last_seen TEXT
        )
    """)
    # Pokud neexists žádný uživatel, vytvoř trial (z legacy dat)
    n = con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    if n == 0:
        con.execute(
            "INSERT OR IGNORE INTO users (username,display_name,is_admin,created_at) "
            "VALUES (?,?,?,?)",
            ("trial", "Trial", 0, datetime.now().isoformat()))
    con.commit()
    con.close()
    # Zajistit adresáře
    USERS_DIR.mkdir(parents=True, exist_ok=True)
    SYSTEM_DIR.mkdir(parents=True, exist_ok=True)
    # Inicializovat trial uživatele z legacy dat
    trial_dir = USERS_DIR / "trial"
    trial_dir.mkdir(exist_ok=True)
    _migrate_legacy_to_trial(trial_dir)
    # Zkopírovat systémové slovníky do system/
    _init_system_dicts()


def _migrate_legacy_to_trial(trial_dir: pathlib.Path) -> None:
    """Moves existing data to the trial user (one-time migration)."""
    import shutil
    for fname, legacy_path in [
        ("paleon.db",                    DB_FILE),
        ("section_schema.tsv",           SCHEMA_FILE),
        ("paleon_morphology_terms.tsv",  MORPHOLOGY_FILE),
        ("stratigraphy_terms.tsv",       STRATIGRAPHY_FILE),
        ("systematic_sections.tsv",      SYSTEMATIC_SECTIONS_FILE),
        ("taxons.txt",                   GAZETTEER_FILE),
        ("settings.json",                SETTINGS_FILE),
    ]:
        dst = trial_dir / fname
        if not dst.exists() and legacy_path.exists():
            shutil.copy2(str(legacy_path), str(dst))


def _init_system_dicts() -> None:
    """Copies the current dictionaries to system/ as templates for new users."""
    import shutil
    for fname, legacy_path in [
        ("section_schema.tsv",           SCHEMA_FILE),
        ("paleon_morphology_terms.tsv",  MORPHOLOGY_FILE),
        ("stratigraphy_terms.tsv",       STRATIGRAPHY_FILE),
        ("systematic_sections.tsv",      SYSTEMATIC_SECTIONS_FILE),
        ("taxons.txt",                   GAZETTEER_FILE),
    ]:
        dst = SYSTEM_DIR / fname
        if not dst.exists() and legacy_path.exists():
            shutil.copy2(str(legacy_path), str(dst))


def _create_user(username: str, display_name: str = "",
                 password: str = None, is_admin: bool = False) -> bool:
    """
    Creates a new user: registry entry + directory + copies of system dictionaries.
    Returns True on success, False if the user already exists.
    """
    import shutil
    # Guard against reserved names
    _RESERVED = {"admin", "root", "system", "anonymous", "guest", "test"}
    if username.strip().lower() in _RESERVED:
        return False  # reserved name
    pw_hash = hashlib.sha256(password.encode()).hexdigest() if password else None
    try:
        con = _users_db()
        con.execute(
            "INSERT INTO users (username,display_name,password_hash,is_admin,created_at) "
            "VALUES (?,?,?,?,?)",
            (username.strip().lower(), display_name.strip() or username,
             pw_hash, 1 if is_admin else 0, datetime.now().isoformat()))
        con.commit()
        con.close()
    except Exception:
        return False  # UNIQUE constraint → uživatel exists
    # Vytvořit adresář a zkopírovat systémové slovníky
    user_dir = USERS_DIR / _sanitize_username(username)
    user_dir.mkdir(parents=True, exist_ok=True)
    for fname in ["section_schema.tsv","paleon_morphology_terms.tsv",
                  "stratigraphy_terms.tsv","systematic_sections.tsv","taxons.txt"]:
        src = SYSTEM_DIR / fname
        dst = user_dir / fname
        if src.exists() and not dst.exists():
            shutil.copy2(str(src), str(dst))
    return True


def _delete_user(username: str) -> None:
    """Deletes a user and their data (admin only)."""
    import shutil
    con = _users_db()
    con.execute("DELETE FROM users WHERE username=? COLLATE NOCASE", (username,))
    con.commit()
    con.close()
    user_dir = USERS_DIR / _sanitize_username(username)
    if user_dir.exists():
        shutil.rmtree(str(user_dir), ignore_errors=True)


def _get_all_users() -> List[sqlite3.Row]:
    con = _users_db()
    rows = con.execute(
        "SELECT * FROM users ORDER BY is_admin DESC, username"
    ).fetchall()
    con.close()
    return rows


def _switch_user(username: str) -> None:
    """Switches the active user and invalidates all caches."""
    global _SCHEMA_DF, _MORPH_DF, _STRAT_DF
    st.session_state["pn_user"] = username
    # Aktualizovat last_seen
    try:
        con = _users_db()
        con.execute("UPDATE users SET last_seen=? WHERE username=? COLLATE NOCASE",
                    (datetime.now().isoformat(), username))
        con.commit(); con.close()
    except Exception:
        pass
    # Invalidovat allchny per-session cache
    with _SCHEMA_LOCK:
        _SCHEMA_DF = None
    with _TERM_LOCK:
        _MORPH_DF = None
        _STRAT_DF = None
    _invalidate_schema_derived_regex_cache()
    st.session_state.pop("paleon_settings", None)
    st.session_state.pop("__gaz_cache", None)


def _check_admin_password(entered: str) -> bool:
    return hmac.compare_digest(
        hashlib.sha256(entered.encode()).hexdigest(), ADMIN_PASS_HASH)


def _copy_system_dict_to_user(username: str, dict_name: str) -> bool:
    """Copies the system dictionary to the user profile (overwrites)."""
    import shutil
    src = SYSTEM_DIR / dict_name
    dst = USERS_DIR / _sanitize_username(username) / dict_name
    if src.exists():
        (USERS_DIR / _sanitize_username(username)).mkdir(parents=True, exist_ok=True)
        shutil.copy2(str(src), str(dst))
        return True
    return False

OUTPUT_FIELDS = [
    "RECORD_ID", "SOURCE_DOCUMENT", "SOURCE_DOCUMENT_SHORT_NAME",
    "SOURCE_LANGUAGE", "TAXON_OCCURRENCE_NUMBER", "IS_REPEATED_TAXON_OCCURRENCE",
    "RELATED_RECORD_ID", "TAXON_RANK_AS_WRITTEN", "TAXON_NAME_VERBATIM",
    "TAXON_NAME_ENGLISH_TRANSLATION_IF_APPLICABLE",
    "RAW_TAXONOMIC_BLOCK", "BLOCK_START_LOCATOR", "BLOCK_END_LOCATOR",
    "TAXONOMIC PLACEMENT", "TAXONOMIC PLACEMENT_SOURCE_PAGE(S)",
    "TAXON", "TAXON_SOURCE_PAGE(S)",
    "NOMENCLATURAL ACTS", "NOMENCLATURAL ACTS_SOURCE_PAGE(S)",
    "TYPE TAXON", "TYPE TAXON_SOURCE_PAGE(S)",
    "INCLUDED TAXONS", "INCLUDED TAXONS_SOURCE_PAGE(S)",
    "TYPE SPECIMENS", "TYPE SPECIMENS_SOURCE_PAGE(S)",
    "AUTHOR", "AUTHOR_SOURCE_PAGE(S)",
    "FIGURES", "FIGURES_SOURCE_PAGE(S)",
    "SYNONYMY", "SYNONYMY_SOURCE_PAGE(S)",
    "TYPE MATERIAL", "TYPE MATERIAL_SOURCE_PAGE(S)",
    "MATERIAL EXAMINED", "MATERIAL EXAMINED_SOURCE_PAGE(S)",
    "ETYMOLOGY", "ETYMOLOGY_SOURCE_PAGE(S)",
    "LOCALITY", "LOCALITY_SOURCE_PAGE(S)",
    "STRATIGRAPHY", "STRATIGRAPHY_SOURCE_PAGE(S)",
    "DESCRIPTION", "DESCRIPTION_SOURCE_PAGE(S)",
    "DIAGNOSIS", "DIAGNOSIS_SOURCE_PAGE(S)",
    "SIZE", "SIZE_SOURCE_PAGE(S)",
    "REMARKS", "REMARKS_SOURCE_PAGE(S)",
    "OCCURRENCE", "OCCURRENCE_SOURCE_PAGE(S)",
    "OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS",
    "OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS_SOURCE_PAGE(S)",
    "REFERENCE", "REFERENCE_SOURCE_PAGE(S)",
    "SOURCE PAGES", "FIELD_LANGUAGE_STATUS",
    "OCR_QUALITY_NOTE", "AMBIGUITY_FLAG",
    "EXTRACTION_CONFIDENCE", "EXTRACTION_NOTES",
    # Feature 1: TYPE SPECIMENS strukturované
    "TYPE_SPECIMEN_KIND", "INSTITUTION_CODE", "CATALOG_NUMBER",
    # Feature 2: SIZE strukturované
    "SIZE_PARSED",
    # Feature 3: parent-child hierarchie
    "PARENT_TAXON_NAME", "PARENT_RANK",
    # Feature 4: completeness
    "COMPLETENESS_SCORE", "MISSING_REQUIRED_FIELDS",
]

SECTION_FIELDS = [f for f in OUTPUT_FIELDS
                  if not f.endswith("_SOURCE_PAGE(S)") and f not in {
                      "RECORD_ID","SOURCE_DOCUMENT","SOURCE_DOCUMENT_SHORT_NAME",
                      "SOURCE_LANGUAGE","TAXON_OCCURRENCE_NUMBER",
                      "IS_REPEATED_TAXON_OCCURRENCE","RELATED_RECORD_ID",
                      "TAXON_RANK_AS_WRITTEN","TAXON_NAME_VERBATIM",
                      "TAXON_NAME_ENGLISH_TRANSLATION_IF_APPLICABLE",
                      "RAW_TAXONOMIC_BLOCK","BLOCK_START_LOCATOR","BLOCK_END_LOCATOR",
                      "SOURCE PAGES","FIELD_LANGUAGE_STATUS","OCR_QUALITY_NOTE",
                      "AMBIGUITY_FLAG","EXTRACTION_CONFIDENCE","EXTRACTION_NOTES",
                  }]

# LLM prompts
LLM_VALIDATION_PROMPT = """\
You are the PaleoN validation assistant. You will receive a candidate taxon heading and context.
Return ONLY valid JSON without any other text:
{"action":"keep|reject|needs_review","reason":"<briefly>","suggested_rank":"<or empty>","confidence":"High|Medium|Low"}
NEVER create taxon records. NEVER fabricate data."""

LLM_BOUNDARY_PROMPT = """\
You are the PaleoN block-boundary assistant. You will receive an approved taxon block and the text after it.
Return ONLY valid JSON:
{"assessment":"correct|too_short|too_long","reason":"<briefly>","suggested_end_marker":"<or empty>"}
NEVER fabricate text or infer missing data."""

LLM_FIELD_PROMPT = """\
You are the PaleoN field-assignment assistant. You will receive a verbatim taxonomic block.
Assign text to canonical fields. Return ONLY valid JSON:
{"fields":{"FIELD":"verbatim text..."},"reasons":{"FIELD":"why"},"confidence":"High|Medium|Low"}
Use ONLY text from the input. Omit missing fields (do not write "Not provided").
Nover summarise or modify the text."""

LLM_TRANSLATION_PROMPT = """\
You are a scientific translation assistant for palaeontological taxonomy.
Translate the provided field values from their source language into English.
Return ONLY valid JSON in this exact format:
{"fields":{"FIELD_NAME":"English translation"}}

Rules:
- Translate ONLY the text values, never the field names (they are already canonical English).
- Latin taxon names, author names, and year numbers must be kept verbatim — do not translate them.
- Stratigraphic unit names (Formation, Member, Stage, Biozone) keep their original proper name but translate the generic part (e.g. "Buchavské souvrství" → "Buchava Formation").
- Geographic place names: keep the original local name but add English equivalent if widely known.
- If a field value is already in English or is a proper name only, return it unchanged.
- Nover summarise, shorten, or infer content that is not in the original text.
- Omit fields you cannot translate (do not include them in the output JSON at all)."""

LLM_CJK_RU_EXTRACTION_PROMPT = """\
You are a palaeontological taxonomy extraction assistant for ClaudePaleoN.
You receive a raw text block from a Chinese or Russian palaeontological monograph
containing one taxonomic treatment (phylum / class / order / family / genus / species).

Your task: extract field values and return ONLY valid JSON in this exact format:
{"fields": {"FIELD_NAME": "value in English"}, "confidence": "High|Medium|Low"}

FIELD NAMES to extract (use only these canonical names):
  DIAGNOSIS        — morphological characteristics / 特征 / 鉴别特征 / Диагноз
  DESCRIPTION      — detailed description / 描述 / Описание
  REMARKS          — discussion, comparison / 讨论 / 比较 / 讨论与比较 / Замечания / Сравнение
  OCCURRENCE       — age and distribution / 时代和分布 / 分布与时代 / Геологическое распространение
  STRATIGRAPHY     — stratigraphic unit / 地层 / Стратиграфия
  LOCALITY         — geographic locality / 产地 / Местонахождение
  TYPE TAXON       — type species or type genus / 模式种 / 模式属 / Типовой вид
  TYPE SPECIMENS   — holotype, paratype etc. / 正模 / 副模 / Голотип
  SYNONYMY         — synonymy list / 同物异名 / Синонимика
  ETYMOLOGY        — name derivation / 种名来源 / 属名来源 / Этимология
  SIZE             — measurements / 壳体度量 / Размеры
  INCLUDED TAXONS  — list of included species or genera / Состав / 包含种

STRICT RULES:
- Latin taxon names (e.g. Circotheca Sysoiev, 1958), author names, and years MUST be kept
  verbatim — never translate or alter them.
- Translate ALL other content into English (from Chinese or Russian).
- Stratigraphic terms: translate generic part, keep formation name (e.g. 寒武纪 → Cambrian).
- Do NOT invent or infer data not present in the text.
- Do NOT include fields that are absent from the block (omit them from JSON).
- Do NOT include the raw Chinese/Russian text in values — provide English translation only.
- If a section label appears but the content is empty or illegible, omit that field.
- Return ONLY the JSON object — no markdown, no preamble, no explanation.\
"""

DEFAULT_SETTINGS: Dict[str, Any] = {
    "llm_enabled": False,
    "lmstudio_base_url": "http://localhost:1234/v1",
    "lmstudio_model": "",
    "llm_temperature": 0.0,
    "llm_timeout": 180,
    "llm_validation_prompt": LLM_VALIDATION_PROMPT,
    "llm_boundary_prompt": LLM_BOUNDARY_PROMPT,
    "llm_field_prompt": LLM_FIELD_PROMPT,
    "llm_translation_prompt": LLM_TRANSLATION_PROMPT,
    "llm_auto_translate": False,   # automatický překlad při indexaci (ne-EN Documenty)
    "ocr_enabled": True,
    "ocr_languages": "eng+ces+rus+deu+fra+chi_sim",
    "pdf_min_chars": 80,
    "use_column_detection": True,
    "taxon_min_confidence": 0.60,
    "taxon_low_confidence": 0.45,
    "outside_systematic_penalty": 0.20,
    "gazetteer_bonus": 0.10,
    "show_rejected": False,
    "theme": "dark",
    "lang": "en",
}

# ══════════════════════════════════════════════════════════════════════════════
# NAČTENÍ SCHÉMATU  (TSV → engine)
# ══════════════════════════════════════════════════════════════════════════════

_SCHEMA_DF: Optional[pd.DataFrame] = None
_SCHEMA_LOCK = threading.Lock()

# Minimální fallback schema (zahrnuto inline – app funguje i bez TSV fileu)
_MINIMAL_SCHEMA_TSV = """\
enabled\tlabel\tlabel_regex\tcanonical_label\ttarget_field\tis_strong\tcan_start_treatment\ttype_species_standalone\tpriority\tlanguage
1\tDiagnosis\tDiagnosis\tDiagnosis\tDIAGNOSIS\t1\t0\t0\t100\ten/mixed
1\tDiagnóza\tDiagnóza\tDiagnóza\tDIAGNOSIS\t1\t0\t0\t100\tcs
1\tДиагноз\tДиагноз\tДиагноз\tDIAGNOSIS\t1\t0\t0\t100\tru
1\tDiagnose\tDiagnose\tDiagnose\tDIAGNOSIS\t1\t0\t0\t100\tde
1\tDiagnose\tDiagnose\tDiagnose\tDIAGNOSIS\t1\t0\t0\t100\tfr
1\t鉴别特征\t鉴别特征\t鉴别特征\tDIAGNOSIS\t1\t0\t0\t100\tzh
1\tEmended diagnosis\tEmended diagnosis\tEmended diagnosis\tDIAGNOSIS\t1\t0\t0\t95\ten/mixed
1\tDifferential diagnosis\tDifferential diagnosis\tDifferential diagnosis\tDIAGNOSIS\t1\t0\t0\t90\ten/mixed
1\tDescription\tDescription\tDescription\tDESCRIPTION\t1\t1\t0\t100\ten/mixed
1\tPopis\tPopis\tPopis\tDESCRIPTION\t1\t1\t0\t100\tcs
1\tОписание\tОписание\tОписание\tDESCRIPTION\t1\t1\t0\t100\tru
1\tBeschreibung\tBeschreibung\tBeschreibung\tDESCRIPTION\t1\t1\t0\t100\tde
1\t描述\t描述\t描述\tDESCRIPTION\t1\t1\t0\t100\tzh
1\tRedescription\tRedescription\tRedescription\tDESCRIPTION\t1\t1\t0\t95\ten/mixed
1\tMorphology\tMorphology\tMorphology\tDESCRIPTION\t1\t0\t0\t80\ten/mixed
1\tSynonymy\tSynonymy\tSynonymy\tSYNONYMY\t1\t0\t0\t100\ten/mixed
1\tSynonymika\tSynonymika\tSynonymika\tSYNONYMY\t1\t0\t0\t100\tcs
1\tСинонимика\tСинонимика\tСинонимика\tSYNONYMY\t1\t0\t0\t100\tru
1\tSynonymie\tSynonymie\tSynonymie\tSYNONYMY\t1\t0\t0\t100\tde
1\t同物异名\t同物异名\t同物异名\tSYNONYMY\t1\t0\t0\t100\tzh
1\tMaterial examined\tMaterial examined\tMaterial examined\tMATERIAL EXAMINED\t1\t0\t0\t100\ten/mixed
1\tStudovaný materiál\tStudovaný materiál\tStudovaný materiál\tMATERIAL EXAMINED\t1\t0\t0\t100\tcs
1\tИзученный материал\tИзученный материал\tИзученный materiál\tMATERIAL EXAMINED\t1\t0\t0\t100\tru
1\tUntersuchtes Material\tUntersuchtes Material\tUntersuchtes Material\tMATERIAL EXAMINED\t1\t0\t0\t100\tde
1\t研究标本\t研究标本\t研究标本\tMATERIAL EXAMINED\t1\t0\t0\t100\tzh
1\tMaterial\tMaterial\tMaterial\tMATERIAL EXAMINED\t0\t0\t0\t70\ten/mixed
1\tMaterial\tMaterial\tMaterial\tMATERIAL EXAMINED\t0\t0\t0\t70\tcs
1\tMатериал\tMатериал\tMатериал\tMATERIAL EXAMINED\t0\t0\t0\t70\tru
1\tType material\tType material\tType material\tTYPE MATERIAL\t1\t0\t0\t100\ten/mixed
1\tTypový materiál\tTypový materiál\tTypový materiál\tTYPE MATERIAL\t1\t0\t0\t100\tcs
1\tТиповой материал\tТиповой материал\tТиповой материал\tTYPE MATERIAL\t1\t0\t0\t100\tru
1\tTypusmaterial\tTypusmaterial\tTypusmaterial\tTYPE MATERIAL\t1\t0\t0\t100\tde
1\tHolotype\tHolotype\tHolotype\tTYPE SPECIMENS\t1\t0\t0\t100\ten/mixed
1\tParatype\tParatype\tParatype\tTYPE SPECIMENS\t1\t0\t0\t95\ten/mixed
1\tHolotyp\tHolotyp\tHolotyp\tTYPE SPECIMENS\t1\t0\t0\t100\tcs
1\tГолотип\tГолотип\tГолотип\tTYPE SPECIMENS\t1\t0\t0\t100\tru
1\tLocality\tLocality\tLocality\tLOCALITY\t1\t0\t0\t100\ten/mixed
1\tLocalita\tLocalita\tLocalita\tLOCALITY\t1\t0\t0\t100\tcs
1\tМестонахождение\tМестонахождение\tМестонахождение\tLOCALITY\t1\t0\t0\t100\tru
1\tFundort\tFundort\tFundort\tLOCALITY\t1\t0\t0\t100\tde
1\t产地\t产地\t产地\tLOCALITY\t1\t0\t0\t100\tzh
1\tType locality\tType locality\tType locality\tLOCALITY\t1\t0\t0\t95\ten/mixed
1\tStratigraphy\tStratigraphy\tStratigraphy\tSTRATIGRAPHY\t1\t0\t0\t100\ten/mixed
1\tStratigraphy\tStratigraphy\tStratigraphy\tSTRATIGRAPHY\t1\t0\t0\t100\tcs
1\tСтратиграфия\tСтратиграфия\tСтратиграфия\tSTRATIGRAPHY\t1\t0\t0\t100\tru
1\tStratigraphie\tStratigraphie\tStratigraphie\tSTRATIGRAPHY\t1\t0\t0\t100\tde
1\t地层\t地层\t地层\tSTRATIGRAPHY\t1\t0\t0\t100\tzh
1\tHorizont\tHorizont\tHorizont\tSTRATIGRAPHY\t0\t0\t0\t80\tcs
1\tRemarks\tRemarks\tRemarks\tREMARKS\t1\t0\t0\t100\ten/mixed
1\tNotes\tNotes\tNotes\tREMARKS\t1\t0\t0\t100\tcs
1\tЗамечания\tЗамечания\tЗамечания\tREMARKS\t1\t0\t0\t100\tru
1\tBemerkungen\tBemerkungen\tBemerkungen\tREMARKS\t1\t0\t0\t100\tde
1\t备注\t备注\t备注\tREMARKS\t1\t0\t0\t100\tzh
1\tDiscussion\tDiscussion\tDiscussion\tREMARKS\t1\t0\t0\t95\ten/mixed
1\tOccurrence\tOccurrence\tOccurrence\tOCCURRENCE\t1\t0\t0\t100\ten/mixed
1\tVýskyt\tVýskyt\tVýskyt\tOCCURRENCE\t1\t0\t0\t100\tcs
1\tРаспространение\tРаспространение\tРаспространение\tOCCURRENCE\t1\t0\t0\t100\tru
1\t分布\t分布\t分布\tOCCURRENCE\t1\t0\t0\t100\tzh
1\tDistribution\tDistribution\tDistribution\tOCCURRENCE\t0\t0\t0\t80\ten/mixed
1\tEtymology\tEtymology\tEtymology\tETYMOLOGY\t1\t0\t0\t100\ten/mixed
1\tEtymologie\tEtymologie\tEtymologie\tETYMOLOGY\t1\t0\t0\t100\tcs
1\tЭтимология\tЭтимология\tЭтимология\tETYMOLOGY\t1\t0\t0\t100\tru
1\tSizes\tSizes\tSizes\tSIZE\t1\t0\t0\t100\ten/mixed
1\tDimensions\tDimensions\tDimensions\tSIZE\t1\t0\t0\t100\ten/mixed
1\tMeasurements\tMeasurements\tMeasurements\tSIZE\t1\t0\t0\t100\ten/mixed
1\tDimensions\tDimensions\tDimensions\tSIZE\t1\t0\t0\t100\tcs
1\tРазмеры\tРазмеры\tРазмеры\tSIZE\t1\t0\t0\t100\tru
1\tMaße\tMaße\tMaße\tSIZE\t1\t0\t0\t100\tde
1\tSize\tSize\tSize\tSIZE\t0\t0\t0\t80\ten/mixed
1\tType species\tType species\tType species\tTYPE TAXON\t1\t0\t1\t100\ten/mixed
1\tTypový druh\tTypový druh\tTypový druh\tTYPE TAXON\t1\t0\t1\t100\tcs
1\tТиповой вид\tТиповой вид\tТиповой вид\tTYPE TAXON\t1\t0\t1\t100\tru
1\tType genus\tType genus\tType genus\tTYPE TAXON\t1\t0\t1\t100\ten/mixed
1\tIncluded species\tIncluded species\tIncluded species\tINCLUDED TAXONS\t1\t0\t0\t100\ten/mixed
1\tIncluded taxa\tIncluded taxa\tIncluded taxa\tINCLUDED TAXONS\t1\t0\t0\t100\ten/mixed
1\tZahrnuté druhy\tZahrnuté druhy\tZahrnuté druhy\tINCLUDED TAXONS\t1\t0\t0\t100\tcs
1\t特征\t特征\t特征\tDIAGNOSIS\t1\t0\t0\t100\tzh
1\t描述\t描述\t描述\tDESCRIPTION\t1\t1\t0\t100\tzh
1\t比较\t比较\t比较\tREMARKS\t1\t0\t0\t95\tzh
1\t讨论\t讨论\t讨论\tREMARKS\t1\t0\t0\t90\tzh
1\t模式种\t模式种\t模式种\tTYPE TAXON\t1\t0\t1\t100\tzh
1\t产地与层位\t产地与层位\t产地与层位\tLOCALITY\t1\t0\t0\t100\tzh
1\t产地与地层\t产地与地层\t产地与地层\tLOCALITY\t1\t0\t0\t95\tzh
1\t地层与时代\t地层与时代\t地层与时代\tSTRATIGRAPHY\t1\t0\t0\t100\tzh
1\t分布与时代\t分布与时代\t分布与时代\tOCCURRENCE\t1\t0\t0\t100\tzh
1\t分布时代\t分布时代\t分布时代\tOCCURRENCE\t1\t0\t0\t95\tzh
1\t时代和分布\t时代和分布\t时代和分布\tOCCURRENCE\t1\t0\t0\t100\tzh
1\t时代与分布\t时代与分布\t时代与分布\tOCCURRENCE\t1\t0\t0\t95\tzh
1\t讨论与比较\t讨论与比较\t讨论与比较\tREMARKS\t1\t0\t0\t92\tzh
1\t比较与讨论\t比较与讨论\t比较与讨论\tREMARKS\t1\t0\t0\t92\tzh
1\t模式属\t模式属\t模式属\tTYPE TAXON\t1\t0\t1\t100\tzh
1\t种名来源\t种名来源\t种名来源\tETYMOLOGY\t1\t0\t0\t100\tzh
1\t属名来源\t属名来源\t属名来源\tETYMOLOGY\t1\t0\t0\t100\tzh
1\t壳体度量\t壳体度量\t壳体度量\tSIZE\t1\t0\t0\t100\tzh
1\t保存壳长\t保存壳长\t保存壳长\tSIZE\t0\t0\t0\t80\tzh
1\t同物异名\t同物异名\t同物异名\tSYNONYMY\t1\t0\t0\t100\tzh
1\tДиагноз\tДиагноз\tДиагноз\tDIAGNOSIS\t1\t0\t0\t100\tru
1\tОписание\tОписание\tОписание\tDESCRIPTION\t1\t1\t0\t100\tru
1\tСравнение\tСравнение\tСравнение\tREMARKS\t1\t0\t0\t100\tru
1\tЗамечания\tЗамечания\tЗамечания\tREMARKS\t1\t0\t0\t95\tru
1\tОбсуждение\tОбсуждение\tОбсуждение\tREMARKS\t1\t0\t0\t90\tru
1\tСостав\tСостав\tСостав\tINCLUDED TAXONS\t1\t0\t0\t100\tru
1\tГеологическое и географическое распространение\tГеологическое и географическое распространение\tГеологическое и географическое распространение\tOCCURRENCE\t1\t0\t0\t100\tru
1\tГеологическое распространение\tГеологическое распространение\tГеологическое распространение\tOCCURRENCE\t1\t0\t0\t95\tru
1\tМатериал и местонахождение\tМатериал и местонахождение\tМатериал и местонахождение\tMATERIAL EXAMINED\t1\t0\t0\t100\tru
1\tТиповой вид\tТиповой вид\tТиповой вид\tTYPE TAXON\t1\t0\t1\t100\tru
1\tТиповой род\tТиповой род\tТиповой род\tTYPE TAXON\t1\t0\t1\t100\tru
"""


# Forced English section aliases requested for post-translation / English remapping.
# NOTE: UI/export still uses canonical PaleoN field names; "Discussion" maps to REMARKS.
_SCHEMA_FORCED_ALIAS_ROWS = [
    {"enabled":"1", "label":"Characteristics", "label_regex":"Characteristics", "canonical_label":"Characteristics", "target_field":"DESCRIPTION", "is_strong":"1", "can_start_treatment":"1", "type_species_standalone":"0", "priority":"110", "language":"en/mixed"},
    {"enabled":"1", "label":"Discussion and Comparison", "label_regex":"Discussion and Comparison", "canonical_label":"Discussion and Comparison", "target_field":"REMARKS", "is_strong":"1", "can_start_treatment":"1", "type_species_standalone":"0", "priority":"110", "language":"en/mixed"},
    {"enabled":"1", "label":"Age and Distribution", "label_regex":"Age and Distribution", "canonical_label":"Age and Distribution", "target_field":"OCCURRENCE", "is_strong":"1", "can_start_treatment":"1", "type_species_standalone":"0", "priority":"110", "language":"en/mixed"},
    {"enabled":"1", "label":"Discussion", "label_regex":"Discussion", "canonical_label":"Discussion", "target_field":"REMARKS", "is_strong":"1", "can_start_treatment":"1", "type_species_standalone":"0", "priority":"105", "language":"en/mixed"},
    {"enabled":"1", "label":"Comparison", "label_regex":"Comparison", "canonical_label":"Comparison", "target_field":"REMARKS", "is_strong":"1", "can_start_treatment":"1", "type_species_standalone":"0", "priority":"105", "language":"en/mixed"},
]


def _apply_forced_schema_aliases(df: pd.DataFrame) -> pd.DataFrame:
    """Append critical English aliases even if user's section_schema.tsv is older."""
    if df is None or df.empty:
        return df
    required_cols = list(df.columns)
    existing = set(str(x).strip().lower() for x in df.get("label", pd.Series(dtype=str)).astype(str))
    rows = []
    for row in _SCHEMA_FORCED_ALIAS_ROWS:
        if row["label"].lower() not in existing:
            rows.append({col: row.get(col, "") for col in required_cols})
    if rows:
        df = pd.concat([df, pd.DataFrame(rows, columns=required_cols)], ignore_index=True)
    return df.reset_index(drop=True)


def load_schema(path: pathlib.Path = None) -> pd.DataFrame:
    """
    Loads the active user's TSV schema. Falls back if the file does not exist.
    Returns a DataFrame with indexed rows.
    """
    global _SCHEMA_DF
    if path is None:
        path = _upath("section_schema.tsv")
        if not path.exists():
            path = SCHEMA_FILE  # legacy fallback
    with _SCHEMA_LOCK:
        if _SCHEMA_DF is not None:
            return _SCHEMA_DF
        if path.exists():
            try:
                df = pd.read_csv(path, sep="\t", encoding="utf-8", dtype=str).fillna("")
                df["enabled"] = df["enabled"].astype(str).str.strip()
                df = df[df["enabled"] == "1"].reset_index(drop=True)
                _SCHEMA_DF = _apply_forced_schema_aliases(df)
                return _SCHEMA_DF
            except Exception as exc:
                logging.warning(f"Schema load failed ({exc}); using fallback.")
        df = pd.read_csv(io.StringIO(_MINIMAL_SCHEMA_TSV), sep="\t", dtype=str).fillna("")
        df = df[df["enabled"] == "1"].reset_index(drop=True)
        _SCHEMA_DF = _apply_forced_schema_aliases(df)
        return _SCHEMA_DF


def reload_schema():
    """Forces schema reload (after uploading a new TSV)."""
    global _SCHEMA_DF
    with _SCHEMA_LOCK:
        _SCHEMA_DF = None
    _invalidate_schema_derived_regex_cache()


# ── Gazetteer (seznam známých taxa) ───────────────────────────────────────
# Volitelný měkký bonus do scoring — NE tvrdý filtr. Pokud kandidát
# odpovídá jménu (or rodu) ze seznamu, dostane bonus k duvěře. Nopřítomnost
# v seznamu NIC nepenalizuje — v textech se mohou objevit i jiné, neznámé
# taxony (nově popisované druhy z aktuálního článku tam logicky missing).
_GAZETTEER_GENERA: Optional[set] = None
_GAZETTEER_BINOMIALS: Optional[set] = None
_GAZETTEER_LOCK = threading.Lock()


def load_gazetteer(path: pathlib.Path = None) -> Tuple[set, set]:
    if path is None:
        path = _upath("taxons.txt")
        if not path.exists(): path = GAZETTEER_FILE
    """Loads the list of known taxa. Returns (set of genus names, set of binomials)."""
    global _GAZETTEER_GENERA, _GAZETTEER_BINOMIALS
    with _GAZETTEER_LOCK:
        if _GAZETTEER_GENERA is not None:
            return _GAZETTEER_GENERA, _GAZETTEER_BINOMIALS
        genera: set = set()
        binomials: set = set()
        if path.exists():
            try:
                for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
                    name = line.strip()
                    if not name:
                        continue
                    parts = name.split()
                    if len(parts) >= 1:
                        genera.add(parts[0].lower())
                    if len(parts) >= 2:
                        binomials.add(f"{parts[0]} {parts[1]}".lower())
            except Exception:
                pass
        _GAZETTEER_GENERA = genera
        _GAZETTEER_BINOMIALS = binomials
        return genera, binomials


def reload_gazetteer():
    """Forces gazetteer reload (after uploading a new taxons.txt)."""
    global _GAZETTEER_GENERA, _GAZETTEER_BINOMIALS
    with _GAZETTEER_LOCK:
        _GAZETTEER_GENERA = None
        _GAZETTEER_BINOMIALS = None


def rebuild_taxon_hierarchy() -> Dict[str, int]:
    """
    Recalculates parent-child relationships for all taxa in the DB.

    Algorithm:
    1. Loads candidates sorted by document + page (= order of occurrence
       in the literature, which corresponds to systematic order).
    2. Maintains a rank-stack (stack of parent taxa). Each candidate
       receives the parent = the nearest higher rank in the stack.
    3. Also: if a taxon has the INCLUDED TAXONS field filled and has no parent,
       matching attempts to find one from the DB (genus → family via included taxa).
    4. Returns statistics: {'updated': N, 'already_ok': M, 'skipped': K}.

    Performance: single O(N log N) pass, N+K DB updates in a single transaction.
    """
    _RANK_ORDER = [
        "phylum", "subphylum", "class", "subclass",
        "superorder", "order", "suborder", "infraorder",
        "superfamily", "family", "subfamily", "tribe", "subtribe",
        "genus", "subgenus", "species", "subspecies",
    ]

    con = db()
    # Načíst allchny kandidáty jedním SQL dotazem (sorted = pořadí v lit.)
    rows = con.execute(
        "SELECT id, taxon_name, rank_guess, page_start, document_id, "
        "parent_taxon_name, parent_rank FROM taxon_candidates "
        "WHERE status IN ('approved','pending','needs_review','low_confidence') "
        "ORDER BY document_id, page_start"
    ).fetchall()

    # Načíst INCLUDED TAXONS pole pro párovací krok (genus → family)
    included_map: Dict[int, str] = {}
    for r in con.execute(
        "SELECT candidate_id, field_value FROM occurrence_fields "
        "WHERE field_name='INCLUDED TAXONS' AND field_value != ''"
    ).fetchall():
        included_map[r["candidate_id"]] = r["field_value"]

    # Sestavit index: name_lower → (id, rank) pro rychlé hledání
    name_index: Dict[str, List[Tuple[int, str]]] = {}
    for r in rows:
        key = (r["taxon_name"] or "").strip().lower()
        if key:
            name_index.setdefault(key, []).append((r["id"], r["rank_guess"] or ""))

    stats = {"updated": 0, "already_ok": 0, "skipped": 0}
    updates: List[Tuple[str, str, int]] = []  # (parent_name, parent_rank, cand_id)

    # Per-Document rank-stack
    _stack: Dict[str, Tuple[str, str]] = {}  # rank → (parent_name, parent_rank)
    _cur_doc = None

    for r in rows:
        rank_raw = (r["rank_guess"] or "").strip().lower().split()[0]
        if rank_raw not in _RANK_ORDER:
            stats["skipped"] += 1
            continue

        # Reset stack při přechodu na nový Document
        if r["document_id"] != _cur_doc:
            _stack = {}
            _cur_doc = r["document_id"]

        # Najdi nejbližšího rodiče
        my_pos = _RANK_ORDER.index(rank_raw)
        parent_name = ""
        parent_rank = ""
        for pr in reversed(_RANK_ORDER[:my_pos]):
            if pr in _stack:
                parent_name, parent_rank = _stack[pr]
                break

        # Aktualizuj stack pro tento rank
        _stack[rank_raw] = (r["taxon_name"] or "").strip(), rank_raw

        # Vymaž nižší ranky (jsou teď mimo Context)
        for pr in _RANK_ORDER[my_pos + 1:]:
            _stack.pop(pr, None)

        # Pokud stále nemáme rodiče, zkus INCLUDED TAXONS z vyšší jednotky
        if not parent_name and rank_raw == "species":
            # Pro druh: hledej rod podle prvního slova jména
            genus_prefix = (r["taxon_name"] or "").split()[0].lower()
            if genus_prefix in name_index:
                for pid, prank in name_index[genus_prefix]:
                    if (prank or "").lower() == "genus":
                        parent_name = genus_prefix.capitalize()
                        parent_rank = "genus"
                        break

        # Porovnej s existujícím - aktualizuj jen pokud se změnilo
        old_pn = (r["parent_taxon_name"] or "").strip()
        old_pr = (r["parent_rank"] or "").strip()
        if old_pn == parent_name and old_pr == parent_rank:
            stats["already_ok"] += 1
        else:
            updates.append((parent_name, parent_rank, r["id"]))
            stats["updated"] += 1

    # Dávková aktualizace v jedné transakci
    if updates:
        con.executemany(
            "UPDATE taxon_candidates SET parent_taxon_name=?, parent_rank=? WHERE id=?",
            updates)
        con.commit()
    con.close()
    return stats



# Stejný princip jako u taxonomického gazetteeru: TSV file s termíny,
# kategoriemi a jazyky, předpočítané shody se ukládají do tabulky
# term_matches při schválení kandidáta (viz compute_and_save_term_matches_*).

_MORPH_DF: Optional[pd.DataFrame] = None
_STRAT_DF: Optional[pd.DataFrame] = None
_TERM_LOCK = threading.Lock()

# ── Cached compiled regexes ─────────────────────────────────────────────────
# Regex se buildí jednou při prvním použití a cachuje se v paměti.
# Při reloadu TSV je cache invalidována přes reload_morphology/stratigraphy_terms().
# (dict[term_lower→(canonical,category)], max_ngrams)
_MORPH_REGEX_CACHE: Optional[Tuple[Dict[str, Tuple[str, str]], int]] = None
_STRAT_REGEX_CACHE: Optional[Tuple[Dict[str, Tuple[str, str]], int]] = None

# ── Named-unit regex — kompilujeme JEDNOU na úrovni modulu ─────────────────
# (dříve bylo uvnitř compute_term_matches → zbytečná re.compile() při každém volání)
_NAMED_STRAT_UNIT_TYPES = (
    # EN
    "Formation|Member|Stage|Zone|Biozone|Group|Suite|Series|System|"
    "Beds|Shale|Limestone|Sandstone|Mudstone|Marl|Horizon|Interval|"
    "Supergroup|Subgroup|Subformation|Submember|"
    # CS/SK
    r"souvrství|sovrství|vrstvy|vrstva|úsek|stupeň|zóna|obzor|pásmo|formace|člen|skupina|horizont|biozona|"
    # DE
    r"Schichten|Schicht|Stufe|Gruppe|Horizont|Glied|Unterformation|"
    # FR
    r"Membre|Étage|Sous-formation|Calcaires|Argiles|Série|"
    # RU
    r"свита|толща|горизонт|ярус|зона|биозона|серия|группа|подсвита|слои|пачка|ступень|"
    # ES/IT/PL/SV
    r"Formación|Miembro|Formazione|Membro|Piano|formacja|ogniwo|piętro|Horisont|"
    # ZH
    r"组|段|层|阶|带|群|统|系|期"
)
_NAMED_STRAT_UNIT_RE = re.compile(
    r"([A-ZА-ЯЁČŠŽŘÉÍÚŮÝ\u4e00-\u9fff]"
    r"[A-Za-zА-ЯЁа-яёčšžřéíúuý\u4e00-\u9fff]+"
    r"(?:[-\s]+[A-ZА-ЯЁ\u4e00-\u9fff][A-Za-zА-ЯЁа-яё\u4e00-\u9fff]+)*)"
    r"[\s\-]+(" + _NAMED_STRAT_UNIT_TYPES + r")",
    re.UNICODE | re.IGNORECASE)

# Tokenizer pro compute_term_matches — modul-level konstanta (ne uvnitř funkce)
_TERM_WORD_RE = re.compile(
    r"[\w\u00c0-\u024f\u0400-\u04ff\u4e00-\u9fff]{2,}", re.UNICODE)


def load_morphology_terms(path: pathlib.Path = None) -> pd.DataFrame:
    if path is None:
        path = _upath("paleon_morphology_terms.tsv")
        if not path.exists(): path = MORPHOLOGY_FILE
    """Loads the list of morphological terms (term, canonical, category, language, enabled)."""
    global _MORPH_DF
    with _TERM_LOCK:
        if _MORPH_DF is not None:
            return _MORPH_DF
        if path.exists():
            try:
                df = pd.read_csv(path, sep="\t", dtype=str, encoding="utf-8").fillna("")
                if "enabled" in df.columns:
                    df = df[df["enabled"].astype(str).str.strip() == "1"]
                _MORPH_DF = df.reset_index(drop=True)
                return _MORPH_DF
            except Exception as exc:
                logging.warning(f"Morphology terms load failed ({exc}); using empty set.")
        _MORPH_DF = pd.DataFrame(columns=["term","canonical","category","language","enabled"])
        return _MORPH_DF


def load_stratigraphy_terms(path: pathlib.Path = None) -> pd.DataFrame:
    if path is None:
        path = _upath("stratigraphy_terms.tsv")
        if not path.exists(): path = STRATIGRAPHY_FILE
    """Loads the list of stratigraphic terms (enabled, term, canonical, category, language, match_mode, source_term_en)."""
    global _STRAT_DF
    with _TERM_LOCK:
        if _STRAT_DF is not None:
            return _STRAT_DF
        if path.exists():
            try:
                df = pd.read_csv(path, sep="\t", dtype=str, encoding="utf-8").fillna("")
                if "enabled" in df.columns:
                    df = df[df["enabled"].astype(str).str.strip() == "1"]
                _STRAT_DF = df.reset_index(drop=True)
                return _STRAT_DF
            except Exception as exc:
                logging.warning(f"Stratigraphy terms load failed ({exc}); using empty set.")
        _STRAT_DF = pd.DataFrame(columns=["enabled","term","canonical","category","language","match_mode"])
        return _STRAT_DF


def reload_morphology_terms():
    global _MORPH_DF
    with _TERM_LOCK:
        _MORPH_DF = None


def reload_stratigraphy_terms():
    global _STRAT_DF
    with _TERM_LOCK:
        _STRAT_DF = None


def _load_raw_tsv(path: pathlib.Path, default_cols: list) -> pd.DataFrame:
    """Load a TSV file WITHOUT filtering enabled/disabled rows — for use in the term editors.
    Unlike load_morphology_terms / load_stratigraphy_terms (which strip enabled=0),
    this returns every row so the editor can show and toggle disabled terms."""
    if path.exists():
        try:
            df = pd.read_csv(path, sep="\t", dtype=str, encoding="utf-8").fillna("")
            return df.reset_index(drop=True)
        except Exception as exc:
            logging.warning(f"_load_raw_tsv({path}): {exc}")
    return pd.DataFrame(columns=default_cols)


def get_term_regex(term_type: str) -> Tuple[re.Pattern, Dict[str, Tuple[str, str]]]:
    """
    Compatibility façade for the legacy Morpho/Strat matcher.
    Returns exactly what the legacy build_term_regex() returned:
    (compiled_regex, term_map). Does not perform token/ngram caching.
    """
    df = load_morphology_terms() if term_type == "morphology" else load_stratigraphy_terms()
    return build_term_regex(df)


# ── Systematic sections ─────────────────────────────────────────────────────
_SYST_DF:    Optional[pd.DataFrame] = None
_SYST_RE_CACHE: Optional[re.Pattern]  = None


def load_systematic_sections(path: pathlib.Path = None) -> pd.DataFrame:
    """
    Loads the systematic section headings table (systematic_sections.tsv).
    Columns: enabled, language, heading, priority, weight, pattern, note
    Returns only rows with enabled=1.
    """

    global _SYST_DF
    if path is None:
        path = _upath("systematic_sections.tsv")
        if not path.exists():
            path = SYSTEMATIC_SECTIONS_FILE
    with _TERM_LOCK:
        if _SYST_DF is not None:
            return _SYST_DF
        if path.exists():
            try:
                df = pd.read_csv(path, sep="\t", dtype=str, encoding="utf-8").fillna("")
                if "enabled" in df.columns:
                    df = df[df["enabled"].astype(str).str.strip() == "1"]
                _SYST_DF = df.reset_index(drop=True)
                return _SYST_DF
            except Exception as exc:
                logging.warning(f"Systematic sections load failed ({exc}); using built-in.")
        _SYST_DF = pd.DataFrame(columns=["enabled","language","heading","priority","weight","pattern","note"])
        return _SYST_DF


def reload_systematic_sections():
    global _SYST_DF, _SYST_RE_CACHE
    with _TERM_LOCK:
        _SYST_DF = None
        _SYST_RE_CACHE = None


def build_systematic_re() -> re.Pattern:
    """
    Builds a re.Pattern for detecting systematic section headings.
    Combines:
      1. Base patterns from systematic_sections.tsv (column 'pattern', only enabled=1 and priority core/support)
      2. Original hard-coded patterns (fallback)
    Result is cached in _SYST_RE_CACHE.
    """
    global _SYST_RE_CACHE
    with _TERM_LOCK:
        if _SYST_RE_CACHE is not None:
            return _SYST_RE_CACHE
    df = load_systematic_sections()
    patterns: List[str] = []
    if not df.empty and "pattern" in df.columns:
        # Použít jen core a support záznamy
        active = df[df.get("priority", pd.Series(["core"]*len(df))).isin(["core","support"])]
        for pat in active["pattern"].dropna():
            pat = str(pat).strip()
            if pat:
                # Extrahovat jádro vzoru (odpagesit (?i) a anchory pro použití v SYSTEMATIC_RE)
                # Použijeme celý pattern jako alternativu
                try:
                    re.compile(pat)   # ověřit validitu
                    patterns.append(pat)
                except re.error:
                    pass
    # Hard-coded fallback (puvodní SYSTEMATIC_RE alternativy)
    _FALLBACK = (
        r"\bSystematic\s+pal(?:a?e)ontology\b|"
        r"\bSystematic\s+characteri[sz]ation\b|"
        r"\bSystematics\b|"
        r"\bTaxonomy\b|"
        r"\bSystematic\s+descriptions?\b|"
        r"\bSystematische\s+Pal[aä]ontologie\b|"
        r"\bSyst[eé]matique\b|"
        r"\bSistematica\b|"
        r"\b\u0421\u0438\u0441\u0442\u0435\u043c\u0430\u0442\u0438\u0447\u0435\u0441\u043a\u0430\u044f\s+\u043f\u0430\u043b\u0435\u043e\u043d\u0442\u043e\u043b\u043e\u0433\u0438\u044f\b|"
        r"\b\u0421\u0438\u0441\u0442\u0435\u043c\u0430\u0442\u0438\u0447\u0435\u0441\u043a\u0430\u044f\s+\u0447\u0430\u0441\u0442\u044c\b|"
        r"\b\u7cfb\u7edf\u53e4\u751f\u7269\u5b66\b|\b\u5206\u7c7b\b"
    )
    if not patterns:
        result = re.compile(_FALLBACK, re.IGNORECASE | re.UNICODE)
    else:
        # Kombinujeme patterns z TSV jako alternativu celých řádku
        # TSV patterns již mají (?i) na začátku každého — extrahujeme je do flagu
        clean_parts = []
        for p in patterns:
            # Odstranit všechny inline globální flagy ze začátku patternu.
            # TSV patterns mohou začínat (?i), (?u), (?iu), (?ui) apod. —
            # kompilujeme s re.IGNORECASE | re.UNICODE, takže inline flagy musí
            # být odstraněny (Python 3.11+ vyhodí "global flags not at start").
            clean = re.sub(r"^\(\?[a-zA-Z]+\)", "", p.strip()).strip()
            try:
                re.compile(clean, re.IGNORECASE)
                clean_parts.append(f"(?:{clean})")
            except re.error:
                # Pokud pattern nejde kompilovat, použij jen text nadpisu
                pass
        combined_tsv = "|".join(clean_parts) if clean_parts else _FALLBACK
        result = re.compile(combined_tsv, re.IGNORECASE | re.UNICODE)
    with _TERM_LOCK:
        _SYST_RE_CACHE = result
    return result


def get_systematic_re() -> re.Pattern:
    """Returns the current systematic regex (from TSV or fallback)."""
    return build_systematic_re()


def build_term_regex(df: pd.DataFrame) -> Tuple[re.Pattern, Dict[str, Tuple[str, str]]]:
    """
    Sestaví jeden regex pro allchny termíny (s hranicí slova \\b…\\b, aby
    "test" nechytalo "testimony" apod.) a mapovací slovník
    term.lower() → (canonical, category).
    """
    term_map: Dict[str, Tuple[str, str]] = {}
    parts: List[str] = []
    for _, row in df.iterrows():
        term = str(row.get("term", "")).strip()
        if len(term) < 3:
            continue
        canonical = str(row.get("canonical", "")).strip() or term
        category  = str(row.get("category", "")).strip()
        term_map[term.lower()] = (canonical, category)
        parts.append(re.escape(term))
    parts.sort(key=len, reverse=True)
    if not parts:
        return re.compile(r"(?!x)x"), term_map
    pattern = re.compile(r"\b(" + "|".join(parts) + r")\b", re.IGNORECASE | re.UNICODE)
    return pattern, term_map


# Která OUTPUT_FIELDS pole se prohledávají primárně pro daný typ termínu.
# ROLLBACK: starší Morpho/Strat regex matcher z claude-morfph.py
TERM_PRIMARY_FIELDS: Dict[str, List[str]] = {
    "stratigraphy": ["STRATIGRAPHY"],
    "morphology":   ["DESCRIPTION", "DIAGNOSIS"],
}


def compute_term_matches(
    field_text_map: Dict[str, str],
    raw_block: str,
    term_type: str,
) -> List[Dict[str, str]]:
    """
    Finds term matches (morphology/stratigraphy) in a taxon record.
    Primarily searches the relevant mapped fields (STRATIGRAPHY for
    stratigraphy; DESCRIPTION+DIAGNOSIS for morphology). If those fields
    are not yet filled (mapping has not yet run), falls back to RAW_TAXONOMIC_BLOCK
    as the source. Returns a deduplicated list (one entry per
    canonical term) with the source field from which the match originated.
    """
    df = load_stratigraphy_terms() if term_type == "stratigraphy" else load_morphology_terms()
    primary_fields = TERM_PRIMARY_FIELDS.get(term_type, [])
    regex, term_map = build_term_regex(df)

    found: Dict[str, Dict[str, str]] = {}

    has_primary_content = False
    for fname in primary_fields:
        text = field_text_map.get(fname, "")
        if text and text != NOT_PROVIDED:
            has_primary_content = True
            for m in regex.finditer(text):
                matched = m.group(1)
                canonical, category = term_map.get(matched.lower(), (matched, ""))
                if canonical not in found:
                    found[canonical] = {
                        "term": matched, "canonical": canonical,
                        "category": category, "source_field": fname,
                    }

    if not has_primary_content and raw_block:
        for m in regex.finditer(raw_block):
            matched = m.group(1)
            canonical, category = term_map.get(matched.lower(), (matched, ""))
            if canonical not in found:
                found[canonical] = {
                    "term": matched, "canonical": canonical,
                    "category": category, "source_field": "RAW_TAXONOMIC_BLOCK",
                }

    # ── Druhý pruchod: detekce pojmenovaných stratigraphic units ────────
    # Zachytí vzory jako "Buchava Formation", "Klabava Member", "Stage 3"
    # které nejsou ve slovníku jako vlastní termín, ale jsou duležitou informací.
    if term_type == "stratigraphy":
        _UNIT_TYPES = (
            # EN
            "Formation|Member|Stage|Zone|Biozone|Group|Suite|Series|System|"
            "Beds|Shale|Limestone|Sandstone|Mudstone|Marl|Horizon|Interval|"
            "Supergroup|Subgroup|Subformation|Submember|"
            # CS/SK
            "souvrství|sovrství|vrstvy|vrstva|úsek|stupeň|zóna|obzor|pásmo|formace|člen|skupina|horizont|biozona|"
            # DE
            "Schichten|Schicht|Stufe|Gruppe|Horizont|Glied|Unterformation|"
            # FR
            "Membre|Étage|Sous-formation|Calcaires|Argiles|Série|"
            # RU
            "свита|толща|горизонт|ярус|зона|биозона|серия|группа|подсвита|слои|пачка|ступень|"
            # ES/IT/PL/SV
            "Formación|Miembro|Formazione|Membro|Piano|formacja|ogniwo|piętro|Horisont|"
            # ZH
            "组|段|层|阶|带|群|统|系|期"
        )
        _NAMED_UNIT_RE = re.compile(
            r"([A-ZА-ЯЁČŠŽŘÉÍÚŮÝ\u4e00-\u9fff]"
            r"[A-Za-zА-ЯЁа-яёčšžřéíúuý\u4e00-\u9fff]+"
            r"(?:[-\s]+[A-ZА-ЯЁ\u4e00-\u9fff][A-Za-zА-ЯЁа-яё\u4e00-\u9fff]+)*)"
            r"[\s\-]+(" + _UNIT_TYPES + r")",
            re.UNICODE | re.IGNORECASE)
        # Hledat ve stratigrafickém textu or raw bloku
        _strat_texts = []
        for fname in ["STRATIGRAPHY", "OCCURRENCE", "LOCALITY"]:
            v = field_text_map.get(fname, "")
            if v and v != "Not provided":
                _strat_texts.append(v)
        if not _strat_texts and raw_block:
            _strat_texts = [raw_block[:3000]]
        for _txt in _strat_texts:
            for _m in _NAMED_UNIT_RE.finditer(_txt):
                _full_name = f"{_m.group(1)} {_m.group(2)}"
                _canonical = _full_name  # celý name je kanonický
                if _canonical not in found and len(_m.group(1)) > 2:
                    found[_canonical] = {
                        "term": _full_name,
                        "canonical": _canonical,
                        "category": _m.group(2),  # Formation / Member / Stage...
                        "source_field": "named_unit_detection",
                    }

    return list(found.values())


def save_term_matches(candidate_id: int, term_type: str, matches: List[Dict[str, str]]) -> None:
    """Overwrites the stored term matches of a given type for this candidate."""
    con = db()
    con.execute("DELETE FROM term_matches WHERE candidate_id=? AND term_type=?",
                (candidate_id, term_type))
    for m in matches:
        con.execute(
            "INSERT INTO term_matches (candidate_id,term_type,term,canonical,category,source_field) "
            "VALUES (?,?,?,?,?,?)",
            (candidate_id, term_type, m["term"], m["canonical"], m["category"], m["source_field"]))
    con.commit()
    con.close()


def compute_and_save_term_matches_for_candidate(candidate_id: int) -> None:
    """
    Calculates and saves morphological AND stratigraphic term matches for
    a single candidate. Called automatically on approval (same as
    auto-mapping of fields) — see _batch_update_status, _batch_llm_validate,
    _save_table_edits.
    """
    cand = get_candidate(candidate_id)
    if not cand:
        return
    fields = get_candidate_fields(candidate_id)
    block = cand["block_text"] or ""
    for term_type in ("morphology", "stratigraphy"):
        matches = compute_term_matches(fields, block, term_type)
        save_term_matches(candidate_id, term_type, matches)


def recompute_all_term_matches(progress_cb=None, include_needs_review: bool = True) -> int:
    """Recalculates terms for ALL approved (and optionally needs_review) candidates."""
    statuses = "('approved','needs_review')" if include_needs_review else "('approved')"
    con = db()
    ids = [r["id"] for r in con.execute(
        f"SELECT id FROM taxon_candidates WHERE status IN {statuses}").fetchall()]
    con.close()
    total = len(ids)
    for i, cid in enumerate(ids):
        if progress_cb:
            progress_cb(i, total)
        compute_and_save_term_matches_for_candidate(cid)
    if progress_cb and total:
        progress_cb(total, total)   # finální 100 %
    return total



# ══════════════════════════════════════════════════════════════════════════════
# Deterministické řešení nejednoznačných labels (label → víc target_field)
# ══════════════════════════════════════════════════════════════════════════════
# Některé labely v section_schema.tsv mapují na >1 cílové pole se STEJNÝM
# Score (priority + is_strong). Bez explicitního pravidla by výběr závisel
# na pořadí řádku v TSV → nedeterministické mapování mezi buildy.
#
# Tento slovník řeší POUZE skutečné TIE (shodné Score). Pokud má některé
# pole v TSV vyšší priority, respektujeme volbu autora schématu — sem patří
# jen labely, kde jsou Score shodná a je třeba deterministicky rozhodnout.
#
# Volba je založena na taxonomické konvenci:
#   • "comparison/comparaison"  → REMARKS  (srovnávací poznámky, ne diagnóza)
#   • "depth"                   → LOCALITY (hloubka nálezu = lokalita, ne size)
#   • "series"                  → STRATIGRAPHY (geol. série; "type series" má
#                                  vlastní specifičtější label → TYPE SPECIMENS)
#   • "dépôt"                   → TYPE MATERIAL (fr. „saveení typu")
#   • "geographic distribution" → OCCURRENCE (rozšíření, ne bodová lokalita)
#   • "emended diagnosis/description" → DIAGNOSIS/DESCRIPTION (ne NOM. ACTS)
_AMBIGUOUS_LABEL_PREFERENCE: Dict[str, str] = {
    "comparison":              "REMARKS",
    "comparisons":             "REMARKS",
    "comparaison":             "REMARKS",
    "comparaisons":            "REMARKS",
    "depth":                   "LOCALITY",
    "series":                  "STRATIGRAPHY",
    "dépôt":                   "TYPE MATERIAL",
    "geographic distribution": "OCCURRENCE",
    "emended diagnosis":       "DIAGNOSIS",
    "emended description":     "DESCRIPTION",
    # Pozn.: labely jako "type designation" (→ TYPE TAXON, prio 100 vs
    # NOM.ACTS 95), "bibliographic citation" (→ AUTHOR prio 70 vs REFERENCE
    # 50) a "replacement name" (→ SYNONYMY 100 vs NOM.ACTS 95) mají v TSV
    # jednoznačně vyšší prioritu — jejich volbu autora schématu respektujeme.
}


def _build_section_regex_uncached(df: pd.DataFrame) -> Tuple[re.Pattern, Dict[str, str]]:
    """
    Builds a regex from all labels and a label → target_field mapping dict.
    Internal — do not call directly, use build_section_regex().

    Improvements (2026-07):
    - Uses label_regex (precise escaping for dots, ?, dashes)
    - Strong labels (is_strong=1) and higher priorities match preferentially
    - Sorting: priority DESC + length DESC (longer alternative wins)
    - Deterministic resolution of TIED ambiguities via _AMBIGUOUS_LABEL_PREFERENCE
      (previously depended on row order in TSV → non-deterministic mapping).
    """
    label_to_field: Dict[str, str] = {}
    label_priority: Dict[str, int] = {}
    entries: List[Tuple[int, int, str, str]] = []   # (score, len, regex_alt, label_key)

    has_label_regex = "label_regex" in df.columns
    has_priority    = "priority"    in df.columns
    has_strong      = "is_strong"   in df.columns

    has_lang = "language" in df.columns
    for row in df.itertuples(index=False):
        label  = str(getattr(row, "label",        "") or "").strip()
        field  = str(getattr(row, "target_field", "") or "").strip()
        pat    = str(getattr(row, "label_regex",  "") or "").strip() if has_label_regex else ""
        prio   = int(str(getattr(row, "priority", "0") or "0").strip() or "0") if has_priority else 0
        strong = str(getattr(row, "is_strong", "0") or "0").strip() == "1" if has_strong else False
        lang   = str(getattr(row, "language", "") or "").strip().lower() if has_lang else ""

        _has_cjk = bool(re.search(r"[\u4e00-\u9fff\u3040-\u30ff]", label))
        _min_len = 2 if _has_cjk else 3
        if len(label) < _min_len or not field:
            continue
        # dwc/db exclusion — ale POUZE camelCase DB identifikátory
        # (namePublishedInYear, decimalLongitude), NE prosté nadpisy
        # jako DIAGNOSIS, OCCURRENCE, DESCRIPTION které jsou i běžné nadpisy.
        _is_camel = bool(re.search(r"[a-z][A-Z]", label))       # camelCase
        _is_dwc = (lang in ("dwc/db", "dwc", "db")) and _is_camel

        if pat:
            try:
                re.compile(pat, re.IGNORECASE)
                alt = pat
            except re.error:
                alt = re.escape(label)
        else:
            alt = re.escape(label)

        lk = label.lower()
        score = prio + (10 if strong else 0)
        _existing_score = label_priority.get(lk, -1)
        _take = False
        if lk not in label_to_field:
            _take = True
        elif score > _existing_score:
            _take = True
        elif score == _existing_score and lk in _AMBIGUOUS_LABEL_PREFERENCE:
            # TIE + máme explicitní preferenci → vezmi preferované pole,
            # jinak ponech dosavadní (deterministické, nezávislé na pořadí řádku).
            _pref = _AMBIGUOUS_LABEL_PREFERENCE[lk]
            if field == _pref and label_to_field.get(lk) != _pref:
                _take = True
        if _take:
            label_to_field[lk] = field
            label_priority[lk] = score
        # entries (regex alternativy) přidáváme VŽDY když label není dwc/db —
        # regex hledá výskyt labelu v textu, cílové pole se pak dohledá
        # z label_to_field (což je už deterministické díky logice výše).
        if not _is_dwc:
            entries.append((score, len(label), alt, lk))

    # Deduplikace regex alternativ podle label_key (ruzné řádky téhož labelu
    # s odlišnými poli by jinak přidaly tutéž alternativu vícekrát → nafouklý
    # a pomalejší regex). Zachováme nejvyšší Score pro řazení.
    _seen_alts: Dict[str, Tuple[int, int, str, str]] = {}
    for e in entries:
        _sc, _ln, _alt, _lk = e
        _prev = _seen_alts.get(_alt)
        if _prev is None or _sc > _prev[0]:
            _seen_alts[_alt] = e
    entries = list(_seen_alts.values())

    entries.sort(key=lambda x: (x[0], x[1]), reverse=True)
    alts = [x[2] for x in entries]

    # ── Robustní separátor za labelem ────────────────────────────────────────
    # Dvouúrovňový:
    #  (a) BARE label + separátor [: . ) — –] / newline / 2+ mezer / tab
    #      (tečka povolena pro nadpisy "Diagnosis." apod.)
    #  (b) label + POKRAČOVACÍ slova (jen malá písmena/spojky) + JEN dvojtečka/
    #      pomlčka/newline — zachytí složené nadpisy jako
    #      "Stratigraphic range and distribution:", "Type horizon and locality:".
    #      Malá písmena zaručí, že "Diagnosis Malinky et al." NEmatchne (M velké).
    # Volitelné koncové s/es → plurál (Paratype→Paratypes).
    # CJK fullwidth interpunkce: ：(FF1A) ，(FF0C) 、(3001) ；(FF1B) 。(3002) ）(FF09)
    _CJK_SEP = r"\uff1a\uff0c\u3001\uff1b\u3002\uff09"
    _SEP = (
        r"s?e?s?"
        r"(?:"
        r"(?:[ \t]+[a-z\u00e0-\u00ff\u0430-\u044f\u010d\u0161\u017e\u0159\u00e9&/][\w'/-]*)"
        r"{1,4}[ \t]*(?:[:\u2014\u2013" + _CJK_SEP + r"]|\n|  {2,})"   # multi-word
        r"|"
        r"[ \t]*(?:[:.)\u2014\u2013\-" + _CJK_SEP + r"]|\n|  +|\t)"    # bare label
        r")"
    )
    pattern = re.compile(
        r"(?:^|\n)[ \t]*(" + "|".join(alts) + r")" + _SEP,
        re.IGNORECASE | re.UNICODE,
    ) if alts else re.compile(r"(?!x)x")
    return pattern, label_to_field


def build_section_regex(df: pd.DataFrame) -> Tuple[re.Pattern, Dict[str, str]]:
    """
    Cached façade over _build_section_regex_uncached. Because the schema
    is process-global (changed only via reload_schema, which invalidates the cache),
    so it is safe to cache the result between calls. Without the cache, the massive
    regex (1300+ labels) was recompiled on every block mapping call.
    """
    global _SECTION_REGEX_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        if _SECTION_REGEX_CACHE is not None:
            return _SECTION_REGEX_CACHE
    # Compilation outside the lock (may take time) — race is harmless (idempotent).
    result = _build_section_regex_uncached(df)
    with _SCHEMA_DERIVED_RE_LOCK:
        _SECTION_REGEX_CACHE = result
    return result


def get_strong_labels(df: pd.DataFrame) -> set:
    """Returns the set of labels with is_strong=1 (cached)."""
    global _STRONG_LABELS_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        if _STRONG_LABELS_CACHE is not None:
            return _STRONG_LABELS_CACHE
    result = set(
        str(getattr(row, "label", "")).lower()
        for row in df.itertuples(index=False)
        if str(getattr(row, "is_strong", "0")).strip() == "1"
    )
    with _SCHEMA_DERIVED_RE_LOCK:
        _STRONG_LABELS_CACHE = result
    return result


def get_strong_fields(df: pd.DataFrame) -> set:
    """Returns the set of target_field values where is_strong=1 (cached)."""
    global _STRONG_FIELDS_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        if _STRONG_FIELDS_CACHE is not None:
            return _STRONG_FIELDS_CACHE
    result = set(
        str(getattr(row, "target_field", ""))
        for row in df.itertuples(index=False)
        if str(getattr(row, "is_strong", "0")).strip() == "1"
    )
    with _SCHEMA_DERIVED_RE_LOCK:
        _STRONG_FIELDS_CACHE = result
    return result


# ══════════════════════════════════════════════════════════════════════════════
# REGEX VZORY PRO DETEKCI
# ══════════════════════════════════════════════════════════════════════════════

# Rodové/druhové vzory
FAMILY_RE = re.compile(
    r"^(?P<name>\??[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]*"
    r"(?:idae|inae|ini|oidea|acea|iformes))\b", re.UNICODE)

# Řád: standardní zoologická/paleontologická koncovka -ecida, -ida (ne -idae!)
# Příklady: Circothecida, Orthothecida, Exilithecida, Hyolithida
# Pozn.: negativní lookbehind (?<![i]) zajistí, že -idae NENÍ zachyceno jako -ida + e
ORDER_RE = re.compile(
    r"^(?P<name>\??[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]*"
    r"(?:ecida|arida|(?<!i)ida))\b",
    re.UNICODE)

# Třída: -morpha (Orthothecimorpha, Hyolithinomorpha), -phyta, -opsida aj.
CLASS_RE = re.compile(
    r"^(?P<name>\??[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]*"
    r"(?:morpha|phyta|opsida|ophyceae|ozoa|mycetes))\b",
    re.UNICODE)

SPECIES_RE = re.compile(
    r"^(?P<name>\??[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+"
    r"\??\s+(?:(?:cf\.|aff\.)\s+)?\??"
    r"[a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü][a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+"
    r"(?:\s+(?:sp\.?\s*nov\.?|n\.?\s*sp\.?|sp\.|spp\.|sp\.\s*indet\.))?)"
    r"(?:\s+[A-Z][A-Za-z\-]+,?\s*\d{4})?",
    re.UNICODE)


# Trinomiální a explicitní subspecifické nadpisy.
# Konzervativní: chytá buď "Genus species subspecies", or "Genus species subsp./ssp. epithet".
# Používá se PŘED SPECIES_RE, aby se poddruh nezkrátil na binomium.
SUBSPECIES_RE = re.compile(
    r"^(?P<n>\??[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+"
    r"\??\s+[a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü][a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+"
    r"\s+(?:(?:subsp\.|ssp\.|var\.)\s+)?"
    r"(?!n\.?$|sp\.?$|nov\.?$|indet\.?$)"
    r"[a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü][a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]{1,})"
    r"(?:\s+[A-Z][A-Za-z\-]+,?\s*\d{4})?",
    re.UNICODE)

# Cyrillic binomial (Russian, Ukrainian, etc.)
CYRILLIC_SPECIES_RE = re.compile(
    r"^(?P<name>\??[А-ЯЁ][а-яёА-ЯЁ\-]+\??\s+[а-яё][а-яё\-]+)"
    r"(?:\s+[А-ЯЁ][А-ЯЁа-яёА-Я\-]+,?\s*\d{4})?",
    re.UNICODE)

# Čínský nadpis taxonu: čínské jméno + latinský name + (autor, rok).
# Vzory z reálných publikací (Qian 1977, 2000):
#   "圆管螺属 Circotheca Sysoiev, 1958"   (rod: čín. jméno + rank znak 属)
#   "圆管螺科 Circothecidae Missarzhevsky, 1969"  (čeleď: 科)
#   "密脊脊管螺 Lophotheca multicostata Qian (MS)"  (druh: bez rank znaku)
# Extrahujeme LATINSKÝ name (pro databázi), čínské jméno se zahodí.
# Rank znak (属=genus, 科=family, 目=order, 纲=class) určuje rank.
_CHINESE_RANK_CHARS = {"属": "Genus", "超科": "Superfamily", "科": "Family", "亚科": "Subfamily", "族": "Tribe", "亚族": "Subtribe", "总目": "Superorder", "目": "Order", "亚目": "Suborder",
                       "纲": "Class", "门": "Phylum", "种": "Species"}
CHINESE_TAXON_RE = re.compile(
    r"^[\u4e00-\u9fff]+?"                      # čínské jméno (non-greedy)
    r"(?P<rankchar>[属科目纲门种])?"             # volitelný rank znak
    r"(?=\s)"                                    # musí následovat mezera
    r"\s+"                                       # mezera
    r"(?P<name>[A-Z][A-Za-z\-]+"               # latinský rod
    r"(?:\s+[a-z][a-z\-]+)?)"                  # volitelný druhový epiteton
    r"(?!\s*——)"                                # NEsmělo by následovat em-dash (etymologie "Eo——始")
    r"(?:\s+[A-Z][A-Za-z\-]+\.?)?"             # volitelný autor
    r"(?:\s*,?\s*\d{4})?"                      # volitelný rok
    r"(?:\s*\((?:MS|新种|新属|gen\.\s*et\s*sp\.\s*nov\.)\))?",  # nom. akt
    re.UNICODE)

# Vzor pro detekci etymologických řádku — ty se NIKDY nepovažují za taxon
# nadpis (e.g. „属名来源　Eo——始、古之意" or „种名来源　qiongzhussi——")
_CHINESE_ETYMOLOGY_RE = re.compile(r"[属种]名来源|名称来源|词源", re.UNICODE)


# ── Chinese / multilingual high-rank treatment detector ─────────────────────
# Used before the conservative generic parser for dense systematic PDFs.
_ZH_TAXON_HEADING_RE = re.compile(
    # \s* (místo \s+): čínský rank znak muže být bez mezery přímo před latinským jménem
    # (e.g. "软舌螺动物门Hyolitha" — typický OCR artefakt skenovaných čínských monografií).
    # Zároveň \s* zachytí i split-line případ "圆管螺属\nC ircotheca" protože \s zahrnuje \n.
    r"(?P<zh>[\u4e00-\u9fff]{1,30})(?P<rank_char>总目|亚目|超科|科|亚科|族|亚族|门|纲|目|属|种)\s*"
    r"(?P<latin>[A-Z](?:\s*[A-Za-z-]){3,}(?:\s+[a-z][a-z-]+)?)"
    r"(?:\s+(?P<author>[A-Z](?:\s*[A-Za-z.-]){1,30}|[A-ZА-ЯЁ][A-Za-zА-Яа-яЁё.-]{1,30}))?"
    r"(?:\s*,?\s*(?P<year>1[5-9]\d{2}|20[0-3]\d))?",
    re.UNICODE,
)
_ZH_FIELD_LABELS_FOR_SPLIT = ("模式种", "模式属", "模式标本", "模式材料", "特征", "分类特征", "描述", "讨论与比较", "讨论", "比较", "时代和分布", "分布与时代", "地层与时代", "插图", "图版")


def _clean_latin_taxon_ocr(name: str) -> str:
    """Fixes OCR artefacts in a Latin name (split-word: "C ircotheca" → "Circotheca").

    Opravuje: velké písmeno + 1 mezera + malá písmena (jednopísmenný OCR split).
    ODSTRANĚN druhý regex `\\b([A-Z][a-z]{2,})\\s+([a-z]{3,})\\b` — byl příliš agresivní
    a nesprávně spojoval binomia ("Hyolithes stylus" → "Hyolithestylus").
    """
    if not name:
        return ""
    s = unicodedata.normalize("NFKC", name)
    # Fix single-letter OCR split: "C ircotheca" → "Circotheca"
    s = re.sub(r"\b([A-Z])\s+([a-z]{2,})\b", r"\1\2", s)
    return re.sub(r"\s+", " ", s).strip(" .,;:()[]")


def _infer_rank_from_suffix_and_context(name: str, rank_char: str = "", context: str = "") -> str:
    """Infer exact rank from Chinese rank marker and zoological family-group suffixes."""
    n = _clean_latin_taxon_ocr(name)
    low = n.lower()
    rc_map = {
        "门": "Phylum", "纲": "Class", "总目": "Superorder", "目": "Order", "亚目": "Suborder",
        "超科": "Superfamily", "科": "Family", "亚科": "Subfamily", "族": "Tribe", "亚族": "Subtribe",
        "属": "Genus", "种": "Species",
    }
    if rank_char in rc_map:
        return rc_map[rank_char]
    ctx = context or ""
    # ICZN Article 29.2 family-group suffixes.
    if low.endswith("oidea"):
        return "Superfamily"
    if low.endswith("idae"):
        return "Family"
    if low.endswith("inae"):
        return "Subfamily"
    if low.endswith("ini"):
        return "Tribe"
    if low.endswith("ina"):
        return "Subtribe"
    if re.search(r"(?:总目|超目)", ctx):
        return "Superorder"
    if re.search(r"亚目", ctx):
        return "Suborder"
    if re.search(r"(?:目|超目|亚目)", ctx) or low.endswith(("cida", "ida")):
        return "Order"
    if re.search(r"(?:纲|亚纲|超纲)", ctx) or low.endswith(("imorpha", "morpha")):
        return "Class"
    if re.search(r"门", ctx) or low.endswith(("litha", "zoa", "ozoes")):
        return "Phylum"
    return _rank_from_name_shape(n, "Genus")


def _normalize_chinese_systematic_text(text: str) -> str:
    """Normalises OCR text from a Chinese palaeontological monograph for detection of taxonomic headings.

    Steps:
    1. NFKC unicode normalisation.
    2. Dictionary corrections for specific OCR artefacts (split Latin names, authors).
    3. Generic OCR split-word fix: "A xxxxxx" → "Axxxxxx" for uppercase + space +
       3+ lowercase letters — typical artefact of TH-OCR and similar Chinese scanners.
    4. Insert line breaks before headings (rank-char + Latin name) inside running text.
    5. Split field labels onto their own lines.
    """
    if not text:
        return text
    out = unicodedata.normalize("NFKC", text)

    # 2. Slovníkové OCR opravy — specifické artefakty čínských skenéru (TH-OCR, podobné)
    for bad, good in {
        # Hyolitha / class level
        "H yolitha": "Hyolitha", "H yolithozoes": "Hyolithozoes",
        "H yolithim orpha": "Hyolithimorpha", "H yo-lithimorpha": "Hyolithimorpha",
        "O rthothecim orpha": "Orthothecimorpha",
        "O rthothecioidea": "Orthothecioidea",
        # Order / Family level
        "C ircothecida": "Circothecida", "O rthothecida": "Orthothecida",
        "E xilithecida": "Exilithecida",
        "C ircothecidae": "Circothecidae",
        "H yolithelm inthes": "Hyolithelminthes",
        # Genus level
        "H exitheca": "Hexitheca",
        "C ircotheca": "Circotheca",
        "P aracircotheca": "Paracircotheca",
        "H yolithes": "Hyolithes",
        # Authors
        "Q ian": "Qian", "M arek": "Marek", "M issarzhevsky": "Missarzhevsky",
        "H olm": "Holm", "S ysoiev": "Sysoiev", "L in n arso n": "Linnarson",
        "L innarson": "Linnarson",
        "R u n neg ar": "Runnegar", "R unnegar": "Runnegar",
        # Misc
        "em end .": "emend.", "em end.": "emend.", "nom .": "nom.",
    }.items():
        out = out.replace(bad, good)

    # 3. Generická OCR oprava: velké písmeno + 1 mezera + 3+ malých písmen → slepení
    # Vzory: "C ircotheca" → "Circotheca", "O rthothecida" → "Orthothecida"
    # Bezpečné v čínském Contextu kde "X word" kombinace jsou taxonomická jména.
    out = re.sub(r"\b([A-Z])\s+([a-z]{3,})\b", r"\1\2", out)

    # 4. Vložení \n před nadpisy embeddované v running textu
    # Dvojitý lookbehind:
    #   (?<!\n)           – neopakovat \n pokud nadpis je již na vlastním řádku
    #   (?<![\u4e00-\u9fff]) – nezačínat UVNITŘ čínského složeného slova
    #     (e.g. v "口管螺科Circ..." regex by bez tohoto lookbehind nalezl "管螺科Circ"
    #      a vložil \n před "管螺", přičemž "口" zustane osamoceně → false positive)
    # \s* (místo \s+): rank-znak muže být bezprostředně (bez mezery) před latinským jménem
    out = re.sub(
        r"(?<!\n)(?<![\u4e00-\u9fff])(?P<h>[\u4e00-\u9fff]{1,30}(?:总目|亚目|超科|亚科|亚族|门|纲|目|科|族|属|种)\s*[A-Z][A-Za-z\s-]{3,80}?\s*(?:1[5-9]\d{2}|20[0-3]\d)?)",
        r"\n\g<h>", out)

    # 5. Field-labely na vlastní řádky
    for lab in sorted(_ZH_FIELD_LABELS_FOR_SPLIT, key=len, reverse=True):
        out = re.sub(rf"(?<!\n)\s*({re.escape(lab)})\s+", rf"\n\1 ", out)

    return re.sub(r"\n{3,}", "\n\n", out).strip()


def _match_chinese_highrank_heading(text: str) -> Tuple[Optional[str], str, str]:
    t = _normalize_chinese_systematic_text((text or "").strip())
    line = t.splitlines()[0].strip() if t else ""
    if not line or not re.search(r"[\u4e00-\u9fff]", line) or line.startswith(("插图", "图版")):
        return None, "", ""
    if _CHINESE_ETYMOLOGY_RE.search(line):
        return None, "", ""
    m = _ZH_TAXON_HEADING_RE.match(line)
    if not m:
        return None, "", ""
    name = _clean_latin_taxon_ocr(m.group("latin"))
    if not name or name.lower() in TAXON_STOPWORDS:
        return None, "", ""
    return name, _infer_rank_from_suffix_and_context(name, m.group("rank_char") or "", line), "chinese_highrank_suffix"


def _extract_goldset_style_chinese_treatments(document_id: int, pages: List[PageText], settings: Dict) -> Optional[Dict[str, Any]]:
    parts, page_offsets, off = [], [], 0
    for pg in pages:
        txt = _normalize_chinese_systematic_text(pg.text or "")
        if not txt.strip():
            continue
        if parts:
            parts.append("\n"); off += 1
        start = off; parts.append(txt); off += len(txt); page_offsets.append((start, off, pg.page_number))
    full = "".join(parts)
    if not re.search(r"[\u4e00-\u9fff]", full):
        return None
    matches = []
    for m in _ZH_TAXON_HEADING_RE.finditer(full):
        zh_text = m.group("zh") or ""
        # ── Filtr false-pozitiv zpusobených \s* ──────────────────────────────
        # "模式属 Circotheca" / "模式种 Hyolithes" = typový rod/druh, NENÍ nadpis.
        # Testujeme přítomnost "模式" v čínské části (zh group).
        if "模式" in zh_text:
            continue
        tail = full[m.end():m.end()+260]
        if not any(lbl in tail for lbl in ("特征", "模式种", "模式属", "讨论", "比较", "时代和分布")):
            continue
        name = _clean_latin_taxon_ocr(m.group("latin"))
        # Ochrana proti velmi krátkým jménum (OCR artefakty jako "引目Holm" → "Holm")
        if len(name.split()[0] if name else "") < 5:
            continue
        rank = _infer_rank_from_suffix_and_context(name, m.group("rank_char") or "", m.group(0))
        if name and rank in {"Phylum", "Class", "Superorder", "Order", "Suborder", "Superfamily", "Family", "Subfamily", "Tribe", "Subtribe", "Genus", "Species", "Subspecies"}:
            mm = {"start": m.start(), "end": m.end(), "name": name, "rank": rank, "heading": m.group(0).strip()}
            if not matches or abs(mm["start"] - matches[-1]["start"]) > 10:
                matches.append(mm)
    if len(matches) < 2:
        return None
    def page_for_offset(pos: int) -> int:
        for s, e, pn in page_offsets:
            if s <= pos <= e:
                return pn
        return pages[0].page_number if pages else 1
    con = db(); _delete_candidates_for_document(document_id, con)
    now = datetime.now().isoformat(); accepted = 0
    for i, mm in enumerate(matches):
        end = matches[i+1]["start"] if i+1 < len(matches) else len(full)
        block = re.sub(r"!\[\]\[[^\]]+\]", "", full[mm["start"]:end]).strip()
        # Vyčisti zbývající OCR artefakty v bloku před mapováním fields
        block = re.sub(r"\b([A-Z])\s+([a-z]{3,})\b", r"\1\2", block)
        ps, pe = page_for_offset(mm["start"]), page_for_offset(max(mm["start"], end-1))
        cur = con.execute(
            "INSERT INTO taxon_candidates "
            "(document_id,taxon_name,rank_guess,confidence,status,heading_text,context_before,context_after,"
            "page_start,block_text,created_at,block_end_page,unit_index,debug_json,"
            "boundary_method,boundary_reason,boundary_confidence) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (document_id, mm["name"], mm["rank"], 0.96, "pending", mm["heading"],
             full[max(0, mm["start"]-300):mm["start"]], full[end:min(len(full), end+300)],
             ps, block, now, pe, i,
             json.dumps({"detector": "chinese_goldset_style", "rank_basis": "chinese_marker_or_iczn_suffix"}, ensure_ascii=False),
             "chinese_goldset_style", "heading_to_next_heading", 0.96))
        cid = cur.lastrowid
        try:
            fields = map_sections_from_block(block)
            for fname, fval in fields.items():
                if fval and fval != NOT_PROVIDED:
                    con.execute(
                        "INSERT INTO occurrence_fields (candidate_id,field_name,field_value,source_pages,method) VALUES (?,?,?,?,?)",
                        (cid, fname, fval, f"{ps}-{pe}" if pe != ps else str(ps), "chinese_goldset_style"))
        except Exception as exc:
            logging.warning(f"Chinese field mapping failed: {exc}")
        accepted += 1
    con.commit(); con.close()
    return {"Accepted": accepted, "Low-confidence": 0, "Rejected": 0, "Method": "chinese_goldset_style"}


GENUS_SP_RE = re.compile(
    r"^(?P<name>[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+"
    r"\s+(?:sp\.|spp\.|gen\.\s*indet\.|sp\.\s*indet\.))(?:\s|$)", re.UNICODE)

GENUS_RE = re.compile(
    r"^(?P<name>[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+)"
    r"(?:\s+(?:gen\.\s*nov\.|n\.\s*gen\.)|\s+[A-Z][A-Za-z\-]+,?\s*\d{4}|,?\s*\d{4})?\s*$",
    re.UNICODE)

# Samostatný nadpis "Rod Autor, Rok" BEZ druhového epitetu — typický pro
# rodovou (genus-level) sekci systematického popisu, e.g. "Gracilitheca Sysoev, 1968".
# Striktně ukotveno na celý text (^...$), aby nechytalo fragmenty běžných vět —
# obyčejná věta po "Rok" nekončí, zatímco tento nadpis ano.
GENUS_AUTHOR_RE = re.compile(
    r"""^(?P<name>
            [A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ]
            [a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü]{3,}
        )
        \s+
        (?:[A-Z][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+
           (?:\s*(?:&|and|et|und)\s*[A-Z][A-Za-z\-]+)?
           (?:\s+et\s+al\.)?
        )
        ,?\s*
        (?P<year>1[5-9]\d{2}|20[012]\d)
        \.?\s*$""",
    re.VERBOSE | re.UNICODE)

NEW_TAXON_RE = re.compile(
    r"\b(sp\.\s*nov\.?|n\.\s*sp\.?|gen\.\s*nov\.?|n\.\s*gen\.?|subgen\.\s*nov\.?|n\.\s*gen\.,\s*n\.\s*sp\.?|gen\.\s*et\s*sp\.\s*nov\.?|"
    r"fam\.\s*nov\.?|comb\.\s*nov\.?|nom\.\s*nov\.?|stat\.\s*nov\.?|"
    r"species\s+nova|genus\s+novum|nova\s+species|nový\s+druh|новый\s+вид|новый\s+вид|新种|新属|新科|（新种）|\(新种\)|\(新属\))\b",
    re.IGNORECASE | re.UNICODE)

OPEN_NOMEN_RE = re.compile(
    r"\b(cf\.|aff\.|sp\.\s*indet\.|gen\.\s*et\s*sp\.\s*indet\.|"
    r"incertae\s+sedis|nomen\s+dubium|nomen\s+nudum|indet\.)\b",
    re.IGNORECASE | re.UNICODE)

RANK_LABEL_RE = re.compile(
    r"^(Phylum|Class|Subclass|Superorder|Order|Suborder|Infraorder|Superfamily|"
    r"Family|Subfamily|Tribe|Subtribe|Genus|Subgenus|Species|Subspecies|"
    r"Kmen|Třída|Řád|Čeleď|Rod|Druh|"
    r"Тип|Класс|Подкласс|Отряд|Подотряд|Семейство|Надсемейство|Подсемейство|Род|Подрод|Вид|Подвид|"
    r"Ordnung|Klasse|Gattung|Art|"
    r"Ordning|Klass|Familj|Släkte|"
    r"Ordo|Classis|Familia|"
    r"Famille|Ordre|Genre|Espèce|"
    r"属|种|科|目|纲|门)[\s:\.]*(.*)$",
    re.IGNORECASE | re.UNICODE)

# Kanonizace rank-labelu pro parser. Duležité hlavně proto, že .title()
# u cizojazyčných labels ("Druh", "Вид", "种") nevyrábí hodnoty kompatibilní
# s RANK_OPTIONS a následně se míchá Species/Subspecies/Higher taxon.
_RANK_LABEL_TO_CANONICAL: Dict[str, str] = {
    "phylum": "Phylum", "subphylum": "Phylum", "kmen": "Phylum", "тип": "Phylum", "门": "Phylum",
    "class": "Class", "subclass": "Class", "třída": "Class", "класс": "Class", "подкласс": "Class",
    "klasse": "Class", "klass": "Class", "classis": "Class", "纲": "Class",
    "superorder": "Superorder",
    "order": "Order", "suborder": "Suborder", "infraorder": "Infraorder",
    "řád": "Order", "отряд": "Order", "подотряд": "Order", "ordnung": "Order",
    "ordning": "Order", "ordo": "Order", "ordre": "Order",
    "总目": "Superorder", "目": "Order", "亚目": "Suborder",
    "superfamily": "Superfamily", "family": "Family", "subfamily": "Subfamily",
    "tribe": "Tribe", "subtribe": "Subtribe",
    "čeleď": "Family", "семейство": "Family", "надсемейство": "Superfamily", "подсемейство": "Subfamily",
    "familia": "Family", "famille": "Family", "familj": "Family",
    "超科": "Superfamily", "科": "Family", "亚科": "Subfamily", "族": "Tribe", "亚族": "Subtribe",
    "genus": "Genus", "subgenus": "Subgenus", "rod": "Genus", "род": "Genus", "подрод": "Subgenus",
    "gattung": "Genus", "släkte": "Genus", "genre": "Genus", "属": "Genus",
    "species": "Species", "druh": "Species", "вид": "Species", "art": "Species", "espèce": "Species", "种": "Species",
    "subspecies": "Subspecies", "poddruh": "Subspecies", "подвид": "Subspecies",
    "subsp": "Subspecies", "ssp": "Subspecies", "亚种": "Subspecies",
}
_SPECIES_LEVEL_RANKS = {"species", "subspecies"}
_HIGHER_TYPE_TAXON_RANKS = {"genus", "subgenus", "family", "order", "class", "phylum"}


def _canonical_rank_label(label: str) -> str:
    """Converts a rank label to its canonical English name (Species, Genus, Family…)."""
    key = re.sub(r"[.:]$", "", (label or "").strip().lower())
    return _RANK_LABEL_TO_CANONICAL.get(key, (label or "").strip().title())


def _rank_from_name_shape(name: str, fallback: str = "") -> str:
    """Estimates rank from name morphology (trinomial=Subspecies, binomial=Species…)."""
    raw = (name or "").replace("?", " ").strip()
    _skip = {"cf", "cf.", "aff", "aff.", "subsp", "subsp.", "ssp", "ssp.", "var", "var.",
             "n", "n.", "sp", "sp.", "spp", "spp.", "nov", "nov.", "indet", "indet."}
    parts = []
    for p in re.split(r"\s+", raw):
        q = p.strip(",;:()[]")
        if q and q.lower() not in _skip:
            parts.append(q)
    if not parts:
        return fallback or ""
    if len(parts) >= 3 and parts[0][:1].isupper() and parts[1][:1].islower() and parts[2][:1].islower():
        return "Subspecies"
    if len(parts) >= 2 and parts[0][:1].isupper() and parts[1][:1].islower():
        return "Species"
    if re.search(r"(idae|inae|ini|oidea|acea|iformes)$", parts[0], re.I):
        return "Family"
    return fallback or "Genus"


def _prefer_name_shape_rank(prev_rank: str, name: str, default: str) -> str:
    prev = _canonical_rank_label(prev_rank) if prev_rank else ""
    shaped = _rank_from_name_shape(name, default)
    if shaped in {"Species", "Subspecies"}:
        return prev if prev == "Subspecies" else shaped
    return prev or shaped or default


SYSTEMATIC_RE = re.compile(
    # "pal(?:a?e)ontology" pokrývá BRITSKÝ pravopis "palaeontology" (pal+ae+ontology)
    # I AMERICKÝ "paleontology" (pal+e+ontology) — puvodní "pale?ontology" chytal
    # jen americký tvar a v britsky psaných časopisech (Alcheringa aj.) selhával.
    r"\b(Systematic\s+pal(?:a?e)ontology|Systematics|Taxonomy|Systematic\s+descriptions?|"
    r"Systematic\s+characteri[sz]ation|Systematic\s+section|Now\s+taxa|"
    r"Palaeontological\s+systematics|Description\s+of\s+taxa|"
    r"Taxonomic\s+descriptions?|Formal\s+descriptions?|"
    r"Systematische\s+Paläontologie|Systematik|"
    r"Systématique|Paléontologie\s+systématique|"
    r"Systematická\s+paleontologie|Systematika|"
    r"Систематическая\s+палеонтология|Систематика|"
    r"系统古生物学|分类)\b",
    re.IGNORECASE | re.UNICODE)

# POZN.: bez ukotvení na konci ($) — "References" často sdílí PyMuPDF blok
# s první položkou bibliografie ("References\nBARRANDE, J., 1867…"), což po
# sloučení řádku uvnitř odstavce (viz reading_order) vytvoří jedinou souvislou
# větu. Stačí, že odstavec ZAČÍNÁ tímto slovem.
END_REGION_RE = re.compile(
    r"^\s*(References|Acknowledgements?|Bibliography|Literature\s+cited)\b",
    re.IGNORECASE | re.MULTILINE)

# Detekce stránek abecedního INDEXU (rejstřík na konci monografie).
# Takové stránky typically obsahují řádky ve formátu "Taxon .... 25" or
# "Taxon, 25, 47" — tedy jméno + tečky/čárky + čísla pages. Pokud 70%+ řádku
# na stránce odpovídá tomuto vzoru, stránka je označena jako "index" a
# kandidáti z ní nejsou extrahováni (vyhne se duplicitám bez fields).
_BOOK_INDEX_LINE_RE = re.compile(
    r"^[A-ZА-ЯЁ一-鿿\[\(]"      # begins with capital / CJK
    r"[A-Za-zА-Яа-яёÄÖÜäöüčšžřéíúuý\-\s\.]+?"  # name (with possible spaces/dashes)
    r"[\s\.]{2,}"                 # dots or spaces separating name from pages
    r"[\d,\s–\-]+$",              # page numbers
    re.UNICODE)


# Detekce front-matter obsahu / "indexu Documentu" (Table of Contents).
# Řádky typu "Genus RHIPIDOMELLA ... 10" v obsahu NESMÍ spustit systematickou
# sekci ani vytvořit taxonomické bloky. Bloky se tvoří až od reálné textové části.
_TOC_HEADING_RE = re.compile(
    r"^\s*(_+\s*)?(contents|table\s+of\s+contents|index\s+of\s+contents|"
    r"inhalt|inhaltsverzeichnis|sommaire|table\s+des\s+mati[eè]res|"
    r"содержание|оглавление|目录|目次)\b",
    re.IGNORECASE | re.UNICODE)

_TOC_LEADER_LINE_RE = re.compile(
    r"^\s*.{2,180}?(?:\.{2,}|…{2,}|[-–—_]{3,})\s*"
    r"(?:[ivxlcdm]+|\d+|following\s+\d+)\s*$",
    re.IGNORECASE | re.UNICODE)

_TOC_TRAILING_PAGE_RE = re.compile(
    r"^\s*(?:[A-Z][A-Z\s/&(),.'-]{3,}|[A-Z][A-Za-z][A-Za-z\s/&(),.'-]{3,})\s+"
    r"(?:[ivxlcdm]+|\d+|following\s+\d+)\s*$",
    re.IGNORECASE | re.UNICODE)


def _looks_like_document_index_page(lines: List[str]) -> bool:
    """True for a table-of-contents / front-matter index page."""
    nonempty = [l.strip() for l in lines if l and l.strip()]
    if not nonempty:
        return False
    first_lines = nonempty[:12]
    has_toc_heading = any(_TOC_HEADING_RE.search(l.strip(" _")) for l in first_lines)
    leader_hits = sum(1 for l in nonempty if _TOC_LEADER_LINE_RE.match(l))
    trailing_hits = sum(1 for l in nonempty if _TOC_TRAILING_PAGE_RE.match(l))
    if has_toc_heading and (leader_hits >= 2 or trailing_hits >= 4):
        return True
    if leader_hits >= 6 and (leader_hits / max(1, len(nonempty))) >= 0.25:
        return True
    return False

CAPTION_RE   = re.compile(r"^(Fig\.?|Figure|Plate|Pl\.?|Table|Tab\.?)\s*\d+", re.I)
# POZOR: musí vyžadovat skutečnou strukturu synonymního záznamu — binomium
# následované "Autor, Rok:" a teprve PAK stránkovým/obrázkovým odkazem.
# Bez vyžadování dvojtečky po roce by tento vzor omylem chytal i genuinní
# nadpisy nových druhu typu "Gracilitheca astronauta n. sp. (Figs. 3A–I, 4B)
# Holotype: …", protože ty také obsahují "Fig." někde za sebou.
SYNONYMY_LINE_RE = re.compile(
    r"^(?:\d{4}\s+)?[A-Z][a-zA-Z\-]+\s+[a-z][a-zA-Z\-]+.*?,\s*\d{4}\s*:\s*"
    r".*\b(pl\.|fig\.|p\.|pp\.|стр\.|рис\.)",
    re.I | re.UNICODE)
# Ruský styl synonymiky: "Circotheca billingsi: Мешкова, 1969б, с. 176, табл. LVII, фиг. 1, 2."
# Linka začíná taxonem (Velké+malé) or rodovým jménem, pak ":", autor, rok, citace.
SYNONYMY_LINE_RE_RU = re.compile(
    r"^[A-ZА-ЯЁ][a-zA-Zа-яё\\-]+(?:\\s+[a-zA-Zа-яё\\-]+)?"
    r"\\s*:\\s*"
    r"[A-ZА-ЯЁ][a-zA-Zа-яёäöü\\-]+.*?"
    r",\\s*\\d{4}[a-zа-я]?[\\s;,]+"
    r".*?\\b(с\\\.|стр\\\.|фиг\\\.|табл\\\.|рис\\\.|p\\\.|pl\\\.|figs?\\\.|fig\\\.)",
    re.I | re.UNICODE | re.MULTILINE)
REF_LIKE_RE  = re.compile(r"^[A-Z][A-Za-z\-]+,\s*[A-Z].*\d{4}", re.UNICODE)
AUTHOR_YEAR_RE = re.compile(
    r"(?:\()?"
    r"([A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][A-Za-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]+"
    r"(?:\s+(?:&|and|et)\s+[A-Z][A-Za-z\-]+)?)"
    r"[,\s]+(1[5-9]\d{2}|20[012]\d)"
    r"(?:\))?",
    re.UNICODE)
HYPHEN_BREAK_RE = re.compile(r"(\w)-\s*\n\s*(\w)", re.UNICODE)

TAXON_STOPWORDS = {
    "systematic", "palaeontology", "paleontology", "systematics", "taxonomy",
    "description", "diagnosis", "material", "occurrence", "remarks", "discussion",
    "locality", "stratigraphy", "figure", "table", "small", "large", "shell",
    "conch", "known", "references", "acknowledgements", "type", "species",
    "systematika", "taxonomie",
    # Běžná anglická slova na začátku vět/odstavcu — časté falešné pozitivy
    "this", "the", "a", "an", "to", "in", "on", "for", "with", "we", "it",
    "key", "five", "four", "three", "two", "one", "almost", "subsequently",
    "however", "recently", "published", "downloaded", "please", "taylor",
    "informa", "all", "each", "both", "such", "these", "those", "our",
    "their", "his", "her", "its", "as", "at", "by", "from", "into", "during",
    "including", "according", "based", "following", "here", "thus", "since",
    "because", "although", "additional", "further", "other", "another",
    "specimens", "specimen", "holotype", "paratype", "paratypes",
    # Slovně zapsaná čísla
    "six", "seven", "eight", "nine", "ten", "eleven", "twelve", "thirteen",
    "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen",
    "twenty", "thirty", "forty", "fifty", "hundred", "several", "many",
    "few", "some", "most",
    # Nadpisy úvodních sections
    "mots", "keywords", "abstract", "résumé", "resume", "summary",
    "zusammenfassung", "schlüsselwörter", "ключевые",
    "introduction", "disclosure", "interest", "acknowledgements",
    # Fráze před schema-labely
    "original", "emended", "revised", "general", "differential",
    # ── Doplněno po analýze reálných exportu (2026-07) ─────────────────────
    # Sekční nadpisy chybně detekované jako binomium Genus species:
    "terminology", "faunal", "unfigured", "until", "except", "although",
    "comparison", "introduction", "methods", "method", "results", "conclusions",
    "conclusion", "discussion", "review", "notes", "note", "overview",
    "history", "preservation", "preparation", "collection",
    # Morfologické adjektivy a termíny na začátku diagnóz/popisu
    # ("Monoclaviculate operculum", "Hyolithids characterized" atd.)
    "monoclaviculate", "biclaviculate", "hyolithid", "hyolithids", "orthothecid",
    "orthothecids", "hyolith", "hyoliths",
    "narrow", "wide", "broad", "elongate", "elongated", "compressed",
    "curved", "straight", "conical", "cylindrical", "subcylindrical",
    "circular", "oval", "triangular", "subtriangular", "quadrangular",
    "dorsal", "ventral", "lateral", "apertural", "adapertural",
    "anterior", "posterior", "internal", "external", "apical", "basal",
    "operculum", "aperture", "ligula", "clavicle", "clavicles",
    "ornamented", "smooth", "ribbed", "striated", "sculptured",
    "characterized", "distinguished", "separated", "identified",
    "observed", "figured", "described", "assigned", "attributed",
    "reported", "referred", "considered", "mentioned", "noted",
    "represented", "consisting", "bearing",
    # Noxt sekce a pojmy z recenzních článku
    "repository", "repositories", "collected", "locality", "localities",
    "institutional", "abbreviations", "acknowledgments",
    # Spojky a příslovce v angličtině/němčině/francouzštině
    "whereas", "whilst", "therefore", "nevertheless", "furthermore",
    "moreover", "however", "indeed", "apparently", "presumably",
    "possibly", "probably", "certainly", "apparently",
    # ── OCR-split sekční nadpisy: "Diag nosis" → "Diag" zustane jako FP ──────
    # Přidáme krátké OCR fragmenty sections jako stopwords
    "diag", "diagn", "descr", "morphol", "stratigr", "occurr",
    "emend", "emended", "sensu", "fide", "vide", "cfr", "incl",
    "excl", "stat", "comb", "nov", "nob",
    # ── Měrné a popisné výrazy chybně detekované jako binomia ──────────────
    # "Estimated length", "Maximum width", "Total height" atd.
    "estimated", "estimate", "maximum", "minimum", "average", "approximate",
    "approximately", "total", "overall", "measured", "measurement",
    "length", "width", "height", "thickness", "depth", "diameter",
    "breadth", "ratio", "angle", "curvature", "size",
    # Geometrické a anatomické termíny stojící na začátku věty
    "cross", "section", "outline", "profile", "shape", "form", "type",
    "dorsoventral", "anteroposterior", "transverse", "longitudinal",
    "adapical", "adoral",
    # Geol. časová označení stojící samostatně na začátku nadpisu
    "lower", "upper", "middle", "early", "late", "mid",
    "cambrian", "ordovician", "silurian", "devonian", "carboniferous",
    "permian", "triassic", "jurassic", "cretaceous", "paleocene",
    "eocene", "oligocene", "miocene", "pliocene", "pleistocene",
    # ── P 1.5: Ruské nadpisové stopwords ──────────────────────────────────────
    "аннотация", "введение", "методика", "заключение", "благодарности",
    "результаты", "обсуждение", "выводы", "литература", "резюме",
    "список", "приложение", "дополнение", "таблица", "рисунок", "пластина",
    # ── P 1.5: Čínské stopwords ───────────────────────────────────────────────
    "摘要", "关键词", "引言", "方法", "结论", "致谢", "结果",
    "讨论", "参考文献", "附录", "图版", "表格",
    # ── P 1.5: Latinské nadpisy sections ────────────────────────────────────────
    "conspectus", "enumeratio", "catalogus", "addenda", "corrigenda",
    # ── Časté FP z anglicky psaných paleontologických článku ─────────────────
    # Logistical / metodické nadpisy (Malinky 2002: "Logistical difficulties")
    "logistical", "logistics", "collecting", "preservation", "collection",
    "registration", "accessibility", "difficulties",
    # Morfologické adj/noun páry falešně detekované jako binomia ("outermost layer")
    "outermost", "innermost", "lowermost", "uppermost", "foremost",
    "layer", "layers", "surface", "surfaces", "portion", "portions",
    "region", "regions", "margin", "margins", "edge", "edges",
    "section", "sections", "shell", "specimen", "specimens",
    "feature", "features", "character", "characters", "element",
    "component", "components", "part", "parts", "area", "areas",
    # Posuzovací výrazy
    "general", "typical", "common", "unusual", "unique", "similar",
    "different", "larger", "smaller", "shorter", "longer", "wider",
    # ── Časté FP z geologicko-palaeontologických článku (cave bear paper, 2026-07) ─
    # Záhlaví stránek a institucí
    "département", "departement", "centre", "cahiers", "fichier",
    "citer", "actes", "symposium", "musée", "muséum", "musee", "museum",
    # Statistické a datové termíny
    "statistic", "statistics", "statistical", "measurements", "measurement",
    "radiometric", "comparisons", "comparison", "comparaison",
    "indices", "index", "parameters", "data", "mean", "deviation",
    # Anatomické termíny stojící samostatně (záhlaví tabulek)
    "ramesch", "gamssulzen", "conturines", "zoolithenhohle",
    # Obecné biologicko-geologické termíny detekované jako binomia
    "only", "elements", "cave", "bears", "phylogenetic", "conclusions",
    "phylogenetic", "systematic", "characterisation", "characterization",
    # Francouzské termíny na začátku textu
    "les", "du", "de", "des", "et", "en", "un", "une", "dans", "le", "la",
    "pour", "avec", "que", "qui", "mais", "sur", "par", "au", "aux",
    # Německé termíny
    "die", "der", "den", "das", "des", "ein", "eine", "und", "oder",
    "mit", "von", "aus", "bei", "zur", "zum", "über",
}

# Boilerplate / running-header-footer junk typické pro akademické PDF.
# Tyto řádky NIKDY nesmí být detekovány jako kandidát nadpisu taxonu.
BOILERPLATE_RE = re.compile(
    r"^(?:"
    r"Downloaded by|This article may be used|Published online|"
    r"PLEASE SCROLL DOWN|Taylor\s*(?:and|&)\s*Francis|Informa Ltd|"
    r"To cite this article|To link to this article|Publisher:|"
    r"Terms\s*(?:and|&)\s*Conditions|All rights reserved|"
    r"Registered in England|Registered office|Registered Number|"
    r"Author's personal copy|Author personal copy|"
    r"Disclosure of interest|Conflicts? of interest|"
    r"https?://doi\.org/|https?://www\.|http://dx\.doi|"   # DOI/URL řádky z PDF
    r"Cambridge Core|cambridge\.org|IP address:\s*\d|"      # Cambridge Core download banners
    r"Citer ce document|Cite this document|"                # Cover-page citation blocks
    r"Fichier pdf g[eé]n[eé]r[eé]|"                       # Persee.fr cover page
    r"Département du Rh[oô]ne|Centre de Conservation|"     # French journal running headers
    r"Cahiers scientifiques.*H[oô]rs.s[eé]rie|"           # Journal title in header
    r"D[eé]partement du Rh[oô]ne\s*[-–]\s*Mus[eé]um|"    # Département header variant
    r"Actes du \d+e? [Ss]ymposium"                         # Conference proceedings header
    r")|"
    # Generický vzor běžícího záhlaví: cokoliv krátkého / Časopis (Rok) pagesy
    r"^.{2,45}/\s*.{3,60}\(\d{4}\)\s*\d",
    re.IGNORECASE)

# ══════════════════════════════════════════════════════════════════════════════
# DATOVÉ TŘÍDY
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class PageText:
    page_number: int
    text: str
    method: str
    ocr_note: str = NOT_PROVIDED
    layout_note: str = NOT_PROVIDED

@dataclass
class TextUnit:
    page_number: int
    text: str
    is_bold: bool = False
    is_italic: bool = False
    zone_flags: str = ""   # "systematic", "references", "synonymy", "caption", …
    bbox: Optional[Tuple[float,float,float,float]] = None

# ══════════════════════════════════════════════════════════════════════════════
# DATABÁZE
# ══════════════════════════════════════════════════════════════════════════════

def init_db() -> None:
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    EXPORTS_DIR.mkdir(exist_ok=True)
    UPLOADS_DIR.mkdir(exist_ok=True)
    con = db()  # vždy správná cesta — legacy i user-specific (dle session)
    con.executescript("""
    PRAGMA journal_mode=WAL;
    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY,
        filename TEXT NOT NULL,
        path TEXT,
        lang TEXT DEFAULT 'en',
        page_count INTEGER DEFAULT 0,
        char_count INTEGER DEFAULT 0,
        notes TEXT DEFAULT '',
        created_at TEXT,
        doi TEXT DEFAULT '',
        pub_authors TEXT DEFAULT '',
        pub_journal TEXT DEFAULT '',
        pub_year INTEGER DEFAULT NULL,
        pub_volume TEXT DEFAULT '',
        pub_pages TEXT DEFAULT '',
        pub_title TEXT DEFAULT ''
    );
    CREATE TABLE IF NOT EXISTS pages (
        id INTEGER PRIMARY KEY,
        document_id INTEGER NOT NULL,
        page_number INTEGER NOT NULL,
        text TEXT,
        method TEXT,
        ocr_note TEXT,
        layout_note TEXT
    );
    CREATE TABLE IF NOT EXISTS taxon_candidates (
        id INTEGER PRIMARY KEY,
        document_id INTEGER NOT NULL,
        taxon_name TEXT NOT NULL,
        rank_guess TEXT DEFAULT '',
        confidence REAL DEFAULT 0.0,
        status TEXT DEFAULT 'pending',
        heading_text TEXT,
        context_before TEXT DEFAULT '',
        context_after TEXT DEFAULT '',
        page_start INTEGER DEFAULT 0,
        block_text TEXT DEFAULT '',
        block_end_page INTEGER DEFAULT 0,
        unit_index INTEGER DEFAULT -1,
        debug_json TEXT DEFAULT '{}',
        llm_json TEXT DEFAULT '{}',
        created_at TEXT,
        parent_taxon_name TEXT DEFAULT '',
        parent_rank TEXT DEFAULT '',
        parent_id INTEGER DEFAULT NULL,
        manual_block_text TEXT DEFAULT NULL,
        active_block_source TEXT DEFAULT 'parser',
        boundary_method TEXT DEFAULT 'parser',
        boundary_reason TEXT DEFAULT '',
        boundary_confidence REAL DEFAULT NULL,
        block_version INTEGER DEFAULT 1,
        manual_edited_at TEXT DEFAULT NULL,
        manual_edited_by TEXT DEFAULT NULL,
        start_unit_id INTEGER DEFAULT NULL,
        end_unit_id INTEGER DEFAULT NULL
    );
    CREATE TABLE IF NOT EXISTS occurrence_fields (
        id INTEGER PRIMARY KEY,
        candidate_id INTEGER NOT NULL,
        field_name TEXT NOT NULL,
        field_value TEXT DEFAULT '',
        source_pages TEXT DEFAULT '',
        method TEXT DEFAULT 'manual'
    );
    CREATE TABLE IF NOT EXISTS term_matches (
        id INTEGER PRIMARY KEY,
        candidate_id INTEGER NOT NULL,
        term_type TEXT NOT NULL,
        term TEXT NOT NULL,
        canonical TEXT NOT NULL,
        category TEXT DEFAULT '',
        source_field TEXT DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_cand_doc ON taxon_candidates(document_id);
    CREATE INDEX IF NOT EXISTS idx_occ_cand ON occurrence_fields(candidate_id);
    CREATE INDEX IF NOT EXISTS idx_term_cand ON term_matches(candidate_id);
    CREATE INDEX IF NOT EXISTS idx_term_lookup ON term_matches(term_type, canonical);
    -- FTS5: fulltextové vyhledávání přes taxon_name + block_text + pole
    CREATE VIRTUAL TABLE IF NOT EXISTS fts_candidates USING fts5(
        taxon_name, block_text, all_fields,
        tokenize="unicode61 remove_diacritics 1"
    );
    CREATE TABLE IF NOT EXISTS goldset_documents (
        id INTEGER PRIMARY KEY,
        filename TEXT NOT NULL,
        path TEXT,
        linked_document_id INTEGER,
        created_at TEXT,
        FOREIGN KEY(linked_document_id) REFERENCES documents(id)
    );
    CREATE TABLE IF NOT EXISTS goldset_records (
        id INTEGER PRIMARY KEY,
        goldset_doc_id INTEGER NOT NULL,
        record_n INTEGER NOT NULL,
        taxon_name TEXT NOT NULL,
        rank TEXT DEFAULT '',
        block_text TEXT DEFAULT '',
        expected_fields_json TEXT DEFAULT '{}',
        created_at TEXT,
        FOREIGN KEY(goldset_doc_id) REFERENCES goldset_documents(id)
    );
    """)
    # Migrations – bezpečné ALTER TABLE (SQLite ignoruje duplicity přes try/except)
    try:
        con.execute("ALTER TABLE documents ADD COLUMN notes TEXT DEFAULT ''")
        con.commit()
    except Exception:
        pass
    try:
        con.execute("ALTER TABLE taxon_candidates ADD COLUMN unit_index INTEGER DEFAULT -1")
        con.commit()
    except Exception:
        pass
    # Bod 3: FTS5 virtual table (pokud neexists)
    try:
        con.execute(
            'CREATE VIRTUAL TABLE IF NOT EXISTS fts_candidates USING fts5('
            'taxon_name, block_text, all_fields,'
            'tokenize="unicode61 remove_diacritics 1")')
        con.commit()
    except Exception:
        pass
    # Bod 4: metadata publikace
    for _col, _type in [
        ("doi",         "TEXT DEFAULT ''"),
        ("pub_authors", "TEXT DEFAULT ''"),
        ("pub_journal", "TEXT DEFAULT ''"),
        ("pub_year",    "INTEGER DEFAULT NULL"),
        ("pub_volume",  "TEXT DEFAULT ''"),
        ("pub_pages",   "TEXT DEFAULT ''"),
        ("pub_title",   "TEXT DEFAULT ''"),
    ]:
        try:
            con.execute(f"ALTER TABLE documents ADD COLUMN {_col} {_type}")
            con.commit()
        except Exception:
            pass
    # Feature 3: parent-child hierarchy columns
    for _col, _type in [
        ("parent_taxon_name", "TEXT DEFAULT ''"),
        ("parent_rank",       "TEXT DEFAULT ''"),
        ("parent_id",         "INTEGER DEFAULT NULL"),
    ]:
        try:
            con.execute(f"ALTER TABLE taxon_candidates ADD COLUMN {_col} {_type}")
            con.commit()
        except Exception:
            pass
    # Block Editor columns
    for _col, _type in [
        ("manual_block_text",   "TEXT DEFAULT NULL"),
        ("active_block_source", "TEXT DEFAULT 'parser'"),
        ("boundary_method",     "TEXT DEFAULT 'parser'"),
        ("boundary_reason",     "TEXT DEFAULT ''"),
        ("boundary_confidence", "REAL DEFAULT NULL"),
        ("block_version",       "INTEGER DEFAULT 1"),
        ("manual_edited_at",    "TEXT DEFAULT NULL"),
        ("manual_edited_by",    "TEXT DEFAULT NULL"),
        ("start_unit_id",       "INTEGER DEFAULT NULL"),
        ("end_unit_id",         "INTEGER DEFAULT NULL"),
    ]:
        try:
            con.execute(f"ALTER TABLE taxon_candidates ADD COLUMN {_col} {_type}")
            con.commit()
        except Exception:
            pass
    try:
        # pages_version se zvyšuje při KAŽDÉ změně obsahu stránek Documentu
        # (re-index) — slouží jako invalidační klíč pro cache text units
        # (viz get_cached_text_units / _bump_pages_version níže). Reset
        # detekce ani schvalování kandidátu verzi NEmění, protože text
        # stránek se tím nemění.
        con.execute("ALTER TABLE documents ADD COLUMN pages_version INTEGER DEFAULT 0")
        con.commit()
    except Exception:
        pass
    # ── Výkonnostní indexy (bezpečné – IF NOT EXISTS) ─────────────────────────
    # Bez indexu dotazy nad tisíci záznamy provádějí full-table scan.
    for _idx_sql in [
        "CREATE INDEX IF NOT EXISTS idx_tc_status          ON taxon_candidates(status)",
        "CREATE INDEX IF NOT EXISTS idx_tc_docid           ON taxon_candidates(document_id)",
        "CREATE INDEX IF NOT EXISTS idx_tc_rank            ON taxon_candidates(rank_guess)",
        "CREATE INDEX IF NOT EXISTS idx_tc_taxon_name      ON taxon_candidates(taxon_name)",
        "CREATE INDEX IF NOT EXISTS idx_tc_parent          ON taxon_candidates(parent_taxon_name)",
        "CREATE INDEX IF NOT EXISTS idx_tc_name_rank       ON taxon_candidates(taxon_name, rank_guess)",
        "CREATE INDEX IF NOT EXISTS idx_of_cand_field      ON occurrence_fields(candidate_id, field_name)",
        "CREATE INDEX IF NOT EXISTS idx_of_field_value     ON occurrence_fields(field_name)",
    ]:
        try:
            con.execute(_idx_sql); con.commit()
        except Exception:
            pass
    con.close()


def db() -> sqlite3.Connection:
    """Returns a DB connection for the active user (or legacy during migration)."""
    u = st.session_state.get("pn_user") if st is not None else None
    if u:
        path = USERS_DIR / _sanitize_username(u) / "paleon.db"
        path.parent.mkdir(parents=True, exist_ok=True)
    else:
        path = DB_FILE
    con = sqlite3.connect(str(path), check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    return con


# ══════════════════════════════════════════════════════════════════════════════
# CACHE TEXT UNITS  (výkon extract_block_for_candidate)
# ══════════════════════════════════════════════════════════════════════════════
#
# extract_block_for_candidate() dřív volalo build_text_units(pages_obj) —
# tj. kompletní re-segmentaci celého Documentu — při KAŽDÉM volání. Při
# hromadném schválení N kandidátu se tak Document re-segmentoval N×.
# Cache níže je proces-wide (přežívá napříč Streamlit reruny i mezi
# jednotlivými akcemi uživatele, ne jen po dobu jednoho requestu),
# klíčovaná (document_id, pages_version) — díky explicitnímu verznímu
# čísluv `documents.pages_version` je invalidace spolehlivá i robustní:
# nezávisí na hashování obsahu ani na "náhodě", že se obsah nezmění.
_TEXT_UNITS_CACHE: Dict[Tuple[int, int], List[Any]] = {}
_TEXT_UNITS_CACHE_LOCK = threading.Lock()


def _bump_pages_version(document_id: int, con: Optional[sqlite3.Connection] = None) -> None:
    """
    Increments the pages_version of a document — MUST be called after every change to
    page content (thpicky re-index). Bez tohoto by cache text units vracela
    zastaralá data po re-indexaci téhož Documentu.
    """
    own_con = con is None
    if own_con:
        con = db()
    con.execute(
        "UPDATE documents SET pages_version = COALESCE(pages_version,0) + 1 WHERE id=?",
        (document_id,))
    if own_con:
        con.commit(); con.close()


def invalidate_text_units_cache(document_id: int) -> None:
    """Removes all cached text-unit variants for the given document."""
    with _TEXT_UNITS_CACHE_LOCK:
        for key in [k for k in _TEXT_UNITS_CACHE if k[0] == document_id]:
            del _TEXT_UNITS_CACHE[key]


def get_cached_text_units(document_id: int) -> List[Any]:
    """
    Returns the document's text units — from cache if a valid variant exists
    for the current pages_version, jinak je sestaví (build_text_units) a
    zacachuje. Voláno hlavně z extract_block_for_candidate(), kde se
    dřív stejná re-segmentace opakovala zbytečně u každého kandidáta.
    """
    con = db()
    row = con.execute(
        "SELECT pages_version FROM documents WHERE id=?", (document_id,)
    ).fetchone()
    pages_version = (row["pages_version"] if row and row["pages_version"] is not None else 0)
    cache_key = (document_id, pages_version)

    with _TEXT_UNITS_CACHE_LOCK:
        cached = _TEXT_UNITS_CACHE.get(cache_key)
    if cached is not None:
        con.close()
        return cached

    pages_rows = con.execute(
        "SELECT page_number, text, method, ocr_note, layout_note "
        "FROM pages WHERE document_id=? ORDER BY page_number",
        (document_id,)
    ).fetchall()
    con.close()

    pages_obj = [
        PageText(r["page_number"], r["text"] or "", r["method"],
                 r["ocr_note"] or NOT_PROVIDED, r["layout_note"] or NOT_PROVIDED)
        for r in pages_rows
    ]
    units = build_text_units(pages_obj)
    # Duležité pro extract_block_for_candidate(): blokový filtr pracuje se
    # zone_flags (caption, references, document_index...). Při čtení z cache
    # proto musí mít jednotky stejné flagy jako při detekci kandidátu.
    assign_zone_flags(units)

    with _TEXT_UNITS_CACHE_LOCK:
        # Uklidit staré verze téhož Documentu, ať cache neroste bez mezí.
        for key in [k for k in _TEXT_UNITS_CACHE if k[0] == document_id and k[1] != pages_version]:
            del _TEXT_UNITS_CACHE[key]
        _TEXT_UNITS_CACHE[cache_key] = units
    return units


def _delete_candidates_for_document(document_id: int, con: Optional[sqlite3.Connection] = None) -> None:
    """
    Deletes ALL candidates of the given document and their dependent data
    (occurrence_fields, term_matches) — ve SPRÁVNÉM pořadí.

    KRITICKÉ: occurrence_fields a term_matches se MUSÍ smazat PŘED
    taxon_candidates. Opačné pořadí je bug — subquery
    "WHERE candidate_id IN (SELECT id FROM taxon_candidates WHERE
    document_id=?)" po smazání taxon_candidates už nic nenajde a závislé
    řádky zustanou v DB osiřelé napořád (dřívější stav tohoto kódu).

    Lze volat s vlastním otevřeným spojením (`con=`, nekomituje/nezavírá —
    pro použití uvnitř větší transakce) or bez něj (otevře/commitne/
    zavře vlastní spojení).
    """
    own_con = con is None
    if own_con:
        con = db()
    con.execute("""DELETE FROM occurrence_fields WHERE candidate_id IN
                   (SELECT id FROM taxon_candidates WHERE document_id=?)""", (document_id,))
    con.execute("""DELETE FROM term_matches WHERE candidate_id IN
                   (SELECT id FROM taxon_candidates WHERE document_id=?)""", (document_id,))
    con.execute("DELETE FROM taxon_candidates WHERE document_id=?", (document_id,))
    if own_con:
        con.commit(); con.close()


# ══════════════════════════════════════════════════════════════════════════════
# EXTRAKCE TEXTU  (PDF dvousloupcová, DOCX, TXT)
# ══════════════════════════════════════════════════════════════════════════════

def detect_columns(blocks: list, width: float) -> bool:
    """
    Two-column heuristic based on x-positions of block centres.
    Ported from the original paleon.py (Martin's approach).
    """
    xs = [((b[0]+b[2])/2.0) for b in blocks if len(b)>=5 and str(b[4]).strip()]
    if len(xs) < 8:
        return False
    left   = sum(1 for x in xs if x < width * 0.45)
    right  = sum(1 for x in xs if x > width * 0.55)
    middle = sum(1 for x in xs if width*0.45 <= x <= width*0.55)
    return left >= 3 and right >= 3 and middle <= max(2, len(xs)*0.25)


# Generický running-header/footer text typický pro akademické PDF
# (na rozdíl od BOILERPLATE_RE, který chytá specifické fráze, toto chytá
# tvar "Krátký text + velká písmena + číslo stránky" typický pro záhlaví).
_HEADER_LIKE_RE = re.compile(
    r"^[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ][\w\s\.,&–\-]{2,60}\d{1,4}\s*$"
)


def _is_margin_block(
    x0: float, y0: float, x1: float, y1: float, txt: str,
    width: float, height: float,
) -> bool:
    """
    Decides whether a block is a running header/footer, side watermark
    or other marginalý artefakt, který NESMÍ skončit v hlavním textovém
    proudu stránky (ani v detekci kandidátu, ani v RAW_TAXONOMIC_BLOCK).

    Signály (kterýkoliv stačí):
      1. Extrémní poměr pages bbox (úzký a vysoký/široký) → rotovaný text
         podél okraje stránky (typický vodoznak "Downloaded by …").
      2. Blok leží v horních/dolních ~6 % výšky stránky a je krátký →
         typické umístění běžícího záhlaví/zápatí s číslem pagesy.
      3. Text odpovídá BOILERPLATE_RE (explicitní fráze nakladatelu).
    """
    bw = max(x1 - x0, 0.01)
    bh = max(y1 - y0, 0.01)
    aspect = bh / bw

    # Signál 1: rotovaný/popagesní text (vysoký úzký pruh u okraje stránky)
    if aspect > 5 and bw < width * 0.06:
        return True

    # Signál 2: krátký text v horním/dolním pruhu stránky
    in_top_strip    = y0 < height * 0.062
    in_bottom_strip = y1 > height * 0.94
    if (in_top_strip or in_bottom_strip) and len(txt) < 100:
        compact = re.sub(r"\s+", " ", txt).strip()
        if _HEADER_LIKE_RE.match(compact) or compact.isupper():
            return True

    # Signál 3: explicitní nakladatelský balast
    if BOILERPLATE_RE.search(txt):
        return True

    return False


def reading_order(fitz_page: Any, settings: Dict) -> Tuple[str, str]:
    """
    Extracts text with the correct reading order.
    Two-column layout: sort left column firstevý sloupec, pak pravý.
    Běžící záhlaví/zápatí a popagesní vodoznaky jsou odpagesěny PŘED
    sestavením textu stránky — nesmí kontaminovat RAW_TAXONOMIC_BLOCK.
    """
    raw = []
    width  = float(fitz_page.rect.width)
    height = float(fitz_page.rect.height)
    for b in fitz_page.get_text("blocks"):
        if len(b) < 5:
            continue
        x0, y0, x1, y1, txt = b[:5]
        txt = (txt or "").strip()
        if not txt:
            continue
        if _is_margin_block(float(x0), float(y0), float(x1), float(y1), txt, width, height):
            continue
        raw.append((float(x0), float(y0), float(x1), float(y1), txt))

    two_col = settings.get("use_column_detection", True) and detect_columns(raw, width)
    if two_col:
        ordered = sorted(raw, key=lambda b: (0 if (b[0]+b[2])/2.0 < width/2 else 1, b[1], b[0]))
        note = "two-column layout detected – reading order applied"
    else:
        ordered = sorted(raw, key=lambda b: (b[1], b[0]))
        note = NOT_PROVIDED

    # Spojovat bloky DVOJITÝM newline (paragraph-level granularita pro
    # build_text_units), ale uvnitř KAŽDÉHO bloku sloučit jednotlivé
    # vizuální řádky zpět do plynulého textu — PDF justifikovaný text
    # obsahuje tvrdé zalomení na konci každého vizuálního řádku, což by
    # jinak rozbilo větu na samostatné řádky ("provided\nthe\nfollowing…").
    joined_blocks = []
    for b in ordered:
        block_text = b[4]
        # POŘADÍ JE KRITICKÉ:
        # 1) Nojdřív dehyphenace na PŮVODNÍCH zalomeních řádku uvnitř bloku
        #    (zde ještě exists skutečný \n mezi "low-" a "er", takže
        #    HYPHEN_BREAK_RE muže spojit "low-\ner" → "lower").
        block_text = HYPHEN_BREAK_RE.sub(r"\1\2", block_text)
        # 2) Teprve POTOM sloučit zbylé konce vizuálních řádku do mezery,
        #    ale zachovat skutečné odstavcové zlomy (dvojitý newline) beze změny.
        block_text = re.sub(r"(?<!\n)\n(?!\n)", " ", block_text)
        block_text = re.sub(r"[ \t]{2,}", " ", block_text).strip()
        joined_blocks.append(block_text)

    text = "\n\n".join(joined_blocks)
    return text, note


def _fix_ocr_diacritics(text: str) -> str:
    """
    Fixes typical OCR errors when scanning Czech/Slovak/German/Latin-script
    textsdeckých textu.  Dvě hlavní kategorie problému:

    A) Mark přilepený k písmenu (bez mezery):
       „Dlouha´" → „Dlouhá", „Ko¨r" → „Kör", „rˇ" → „ř"
       Řeší se explicitní substituční tabulkou (NFC normalizace nepostačuje,
       protože ´ (U+00B4) a ˇ (U+02C7) jsou spacing modifiers, ne combining).

    B) Mezera vložená OCR mezi písmeno a diakritiku a/or mezi slabiky:
       „Ty´ rˇ ovice" → „Týřovice"
       „Sˇ a´ rka" → „Šárka", „Nova´ k" → „Novák"
       Po nahrazení mark→composed letter zustanou zbytečné mezery uvnitř
       slov; odpagesí je heuristika pro osamocené krátké fragmenty.
    """
    import unicodedata

    # NFC nejprve – zvládne true combining marks (U+030x série)
    text = unicodedata.normalize("NFC", text)

    # ── Private Use Area znaky → oddělovač odstavce ─────────────────────────
    # U+F8E7 a sousední PUA znaky se v PDF z Cambridge Core, JSTOR apod.
    # používají jako náhrada za bullet/section-break, ale při extrakci PyMuPDF
    # je vloží doslova; bez nahrazení pak sekce nejsou odděleny newlinem a
    # map_sections_from_block je nerozezná jako hranice pole.
    text = re.sub(r"[\uf8e0-\uf8ff\uf000-\uf0ff]", "\n\n", text)

    # ── Control characters ────────────────────────────────────────────────────
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", " ", text)

    # ── Ligatures typické pro OCR starých tisku ──────────────────────────────
    text = (text
            .replace("\ufb00", "ff").replace("\ufb01", "fi")
            .replace("\ufb02", "fl").replace("\ufb03", "ffi").replace("\ufb04", "ffl"))

    # ── Substituční tabulka: spacing diacritic + písmeno → composed ──────────
    # Pořadí: nejdřív delší vzory (dvoupísmenné jako „ch"), pak kratší.
    # Použity znaky: ´ (U+00B4 ACUTE ACCENT), ˇ (U+02C7 CARON),
    #                ˚ (U+02DA RING ABOVE), ¨ (U+00A8 DIAERESIS/UMLAUT),
    #                ` (U+0060 GRAVE), ʹ (U+02B9), ′ (U+2032 PRIME)
    _ACUTE  = r"[´\u02B9\u2032\u0060\u2019]"  # acute-like spacing marks
    _CARON  = r"[ˇ\u02C7]"                     # caron / háček
    _RING   = r"[˚\u02DA]"                     # ring above
    _UMLAUT = r"[¨\u00A8]"                     # diaeresis / umlaut

    _TABLE = [
        # Acute — malá/velká
        (rf"A{_ACUTE}", "Á"), (rf"a{_ACUTE}", "á"),
        (rf"E{_ACUTE}", "É"), (rf"e{_ACUTE}", "é"),
        (rf"I{_ACUTE}", "Í"), (rf"i{_ACUTE}", "í"),
        (rf"O{_ACUTE}", "Ó"), (rf"o{_ACUTE}", "ó"),
        (rf"U{_ACUTE}", "Ú"), (rf"u{_ACUTE}", "ú"),
        (rf"Y{_ACUTE}", "Ý"), (rf"y{_ACUTE}", "ý"),
        # Acute reversed (mark first)
        (rf"{_ACUTE}A", "Á"), (rf"{_ACUTE}a", "á"),
        (rf"{_ACUTE}E", "É"), (rf"{_ACUTE}e", "é"),
        (rf"{_ACUTE}I", "Í"), (rf"{_ACUTE}i", "í"),
        (rf"{_ACUTE}O", "Ó"), (rf"{_ACUTE}o", "ó"),
        (rf"{_ACUTE}U", "Ú"), (rf"{_ACUTE}u", "ú"),
        (rf"{_ACUTE}Y", "Ý"), (rf"{_ACUTE}y", "ý"),
        # Caron / háček — malá/velká
        (rf"C{_CARON}", "Č"), (rf"c{_CARON}", "č"),
        (rf"D{_CARON}", "Ď"), (rf"d{_CARON}", "ď"),
        (rf"E{_CARON}", "Ě"), (rf"e{_CARON}", "ě"),
        (rf"N{_CARON}", "Ň"), (rf"n{_CARON}", "ň"),
        (rf"R{_CARON}", "Ř"), (rf"r{_CARON}", "ř"),
        (rf"S{_CARON}", "Š"), (rf"s{_CARON}", "š"),
        (rf"T{_CARON}", "Ť"), (rf"t{_CARON}", "ť"),
        (rf"Z{_CARON}", "Ž"), (rf"z{_CARON}", "ž"),
        # Caron reversed
        (rf"{_CARON}C", "Č"), (rf"{_CARON}c", "č"),
        (rf"{_CARON}D", "Ď"), (rf"{_CARON}d", "ď"),
        (rf"{_CARON}E", "Ě"), (rf"{_CARON}e", "ě"),
        (rf"{_CARON}N", "Ň"), (rf"{_CARON}n", "ň"),
        (rf"{_CARON}R", "Ř"), (rf"{_CARON}r", "ř"),
        (rf"{_CARON}S", "Š"), (rf"{_CARON}s", "š"),
        (rf"{_CARON}T", "Ť"), (rf"{_CARON}t", "ť"),
        (rf"{_CARON}Z", "Ž"), (rf"{_CARON}z", "ž"),
        # Ring above — u / Ů (česky), å / Å (skandinávsky)
        (rf"U{_RING}", "Ů"), (rf"u{_RING}", "u"),
        (rf"A{_RING}", "Å"), (rf"a{_RING}", "å"),
        (rf"{_RING}U", "Ů"), (rf"{_RING}u", "u"),
        (rf"{_RING}A", "Å"), (rf"{_RING}a", "å"),
        # Umlaut / diaeresis — německy, estonsky, švédsky
        (rf"A{_UMLAUT}", "Ä"), (rf"a{_UMLAUT}", "ä"),
        (rf"O{_UMLAUT}", "Ö"), (rf"o{_UMLAUT}", "ö"),
        (rf"U{_UMLAUT}", "Ü"), (rf"u{_UMLAUT}", "ü"),
        (rf"{_UMLAUT}A", "Ä"), (rf"{_UMLAUT}a", "ä"),
        (rf"{_UMLAUT}O", "Ö"), (rf"{_UMLAUT}o", "ö"),
        (rf"{_UMLAUT}U", "Ü"), (rf"{_UMLAUT}u", "ü"),
    ]
    for pat, repl in _TABLE:
        text = re.sub(pat, repl, text)

    # ── Mezerami oddělené diakritiky ─────────────────────────────────────────
    # Po výše provedené substituci mohou zustat mezery uvnitř slov, kde OCR
    # každou slabiku/souhlásku extrahoval zvlášť. Heuristika:
    #
    # 1) Osamocený znak s háčkem/čárkou (1-2 znaky) obklopený delšími fragmenty
    #    → pravděpodobně uprostřed slova: „Tý ř ovice" → „Týřovice"
    #    Podmínka: předchozí fragment ≥ 2 znaky, následující ≥ 2 znaky,
    #    osamocený znak je písmeno s diakritikou.
    DIACRITIC_SOLO = re.compile(
        r"(\b\w{1,6})\s+([áčďéěíňóřšťúuýžÁČĎÉĚÍŇÓŘŠŤÚŮÝŽäöüÄÖÜåÅ])\s+(\w{2,})",
        re.UNICODE
    )
    def _join_solo(m: re.Match) -> str:
        pre, solo, post = m.group(1), m.group(2), m.group(3)
        if pre[-1] in ".,;:!?0123456789)]}":
            return m.group(0)
        return pre + solo + post

    # Aplikovat opakovaně (kaskáda: „Sˇ á rka" → „Šárka" potřebuje 2 pruchody)
    for _ in range(3):
        prev = text
        text = DIACRITIC_SOLO.sub(_join_solo, text)
        if text == prev:
            break

    # Případ 2: fragment zakončen hákovanou souhláskou + mezera + pokračování
    # „Zahoř any" → „Zahořany", „Nahoř any" → „Nahořany"
    # Hákové souhlásky (ř č š ž ď ť ň) téměř nikdy nestojí na konci samostatného
    # slova — pokud za nimi následuje mezera a malé písmeno, jde o OCR split.
    CARON_END_SPLIT = re.compile(
        r"([řčšžďťňŘČŠŽĎŤŇ])\s+([a-záčďéěíňóřšťúuýž]\w*)",
        re.UNICODE
    )
    for _ in range(2):
        prev = text
        text = CARON_END_SPLIT.sub(lambda m: m.group(1) + m.group(2), text)
        if text == prev:
            break

    # Případ 3: fragment zakončen u + mezera + 1-2 znakové pokračování
    # „Dvu r" → „Dvur", „lu v" → „luv"
    # u stojící na konci fragmentu je téměř vždy nedokončené slovo.
    RING_U_SPLIT = re.compile(
        r"(u)\s+([a-záčďéěíňóřšťúuýž]{1,2})\b",
        re.UNICODE
    )
    for _ in range(2):
        prev = text
        text = RING_U_SPLIT.sub(lambda m: m.group(1) + m.group(2), text)
        if text == prev:
            break

    # Případ 4: KRÁTKÝ fragment (2-4 znaky) zakončen ostrou diakritikou +
    # mezera + krátké pokračování → příjmení or pádová koncovka
    # „Krá luv" → „Králuv", „Nová k" → „Novák"
    # PODMÍNKA délky ≤ 4: „Dlouhá hora" (6 znaku) se NESPOJÍ — Dlouhá
    # je complete adjektivum, hora je samostatné slovo.
    ACUTE_VOWEL_FINAL_SPLIT = re.compile(
        r"\b(\w{2,4}[áéíóúý])\s+(\w{1,4})\b(?![A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ])"
        r"(?=[\s,;.(]|$)",
        re.UNICODE
    )
    for _ in range(2):
        prev = text
        text = ACUTE_VOWEL_FINAL_SPLIT.sub(
            lambda m: m.group(1) + m.group(2), text)
        if text == prev:
            break

    # ── Rozházená písmena OCR ("D i l y t e s" → "Dilytes") ─────────────────
    # Starší skeny (Barrande, Holm, Novák) občas obsahují slova, kde OCR
    # extrahuje každé písmeno zvlášť oddělené mezerou. Podmínka: velké
    # písmeno následované 3 or více (mezera + malé písmeno) → sloučit.
    # Lookbehind / lookahead (?<!\w) / (?!\w) zabrání shodě uvnitř normálních slov.
    # Minimální délka 4 znaky (A b c = 3 malá) → vyhýbá se zkratkám.
    _SPACED_WORD_RE = re.compile(
        r"(?<!\w)"
        r"([A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜА-ЯЁ])"
        r"((?:[ 	][a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûüа-яё]){2,})"
        r"(?!\w)",
        re.UNICODE
    )
    text = _SPACED_WORD_RE.sub(
        lambda m: m.group(1) + m.group(2).replace(" ", "").replace("	", ""),
        text
    )
    # ── OCR word-split pro čínské práce: "A mbrolinevitus" → "Ambrolinevitus" ─
    # Čínské papíry (skenované) občas rozlomí latinské jméno na hranici řádku tak,
    # že první 1–2 písmena jsou oddělena mezerou od zbytku slova.
    # Vzory:
    #   "A mbrolinevitus" → "Ambrolinevitus" (1 velké + mezera + zbytek)
    #   "p latyp" → "platyp" (1 malé + mezera + zbytek, uvnitř slova)
    # Bezpečnostní podmínka: první fragment musí mít 1–2 znaky, pokračování ≥3 znaky
    # a nesmí jít o volné jednopísmenné zkratky před novým slovem (s mezerou na obou pagesách).
    _WORD_SPLIT_ZH_RE = re.compile(
        r"(?<![A-Za-z])([A-Z]{1,2}) ([a-z]{3,})(?=[^a-zA-Z]|$)",
        re.UNICODE)
    text = _WORD_SPLIT_ZH_RE.sub(lambda m: m.group(1) + m.group(2), text)
    # ── Krátká cyriličká strukturní slova (versálky) ────────────
    # Sovětské monografie používají "versálky" (rozepsané velké) pro rankové labely,
    # kde finální písmena jsou VELKÁ i u malých slov: "RoД" = Род (no, it's Р о Д).
    _CYRILLIC_SHORT = {
        # Род
        "Р о Д": "Род",   # R o D (versal D)
        "Р О Д": "Род",   # R O D (all caps)
        "Р о д": "Род",   # R o d (correct lc)
        # Вид
        "В и Д": "Вид",   # V i D
        "В И Д": "Вид",   # V I D
        # Тип
        "Т и п": "Тип",   # T i p
        "Т И П": "Тип",   # T I P
        # Семейство
        "С е м е й с т в о": "Семейство",  # С е м е й с т в о
        "С Е М Е Й С Т В О": "Семейство",  # С Е М Е Й С Т В О
        # Отряд
        "О т р я д": "Отряд",  # О т р я д
        "О Т Р Я Д": "Отряд",  # О Т Р Я Д
        # Класс
        "К л а с с": "Класс",  # К л а с с
        "К Л А С С": "Класс",  # К Л А С С
        # Сравнение
        "С р а в н е н и е": "Сравнение",
        # Состав
        "С о с т а в": "Состав",  # С о с т а в
    }
    for _old_s, _new_s in _CYRILLIC_SHORT.items():
        text = text.replace(_old_s, _new_s)


    return text





def _paragraphize(text: str) -> str:
    """
    Normalises whitespace and fixes typical OCR problems.

    Step order:
    1. Fixava diakritiky (OCR chyby: Ty´ rˇ ovice → Týřovice)
    2. CR/LF → LF, redukce vícenásobných mezer
    3. Sloučení soft-wrap řádku: jednoduchý \\n kde věta pokračuje malým
       písmenem se sloučí na mezeru (nespojuje pokud příští řádek začíná
       velkým písmenem or je prázdný).
    4. Normalizace prázdných řádku (3+ → 2).
    """
    text = _fix_ocr_diacritics(text)
    text = re.sub(r"\r\n|\r", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    # Soft wrap: hyphen-break
    text = HYPHEN_BREAK_RE.sub(r"\1\2", text)
    # Soft wrap: řádek pokračuje malým písmenem → sloučit na mezeru
    text = re.sub(
        r"([a-záčďéěíňóřšťúuýžäöü,;])\n([a-záčďéěíňóřšťúuýžäöü])",
        r"\1 \2", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    # ── OCR split-word oprava pro sekční nadpisy a latinské termíny ──────────
    # "Diag nosis" → "Diagnosis", "Descr iption" → "Description" atd.
    # Vzor: 3-8 písmen začínající VELKÝM + mezera + 3-8 malých písmen
    # (pouze pokud obě části dohromady tvoří slovo bez mezerový výsledek)
    _KNOWN_SECTION_SPLITS = [
        (r"\bDiag\s+nosis\b", "Diagnosis"),
        (r"\bDiagn\s+osis\b", "Diagnosis"),
        (r"\bDescr\s+iption\b", "Description"),
        (r"\bOccurr\s+ence\b", "Occurrence"),
        (r"\bStratigr\s+aphy\b", "Stratigraphy"),
        (r"\bMorphol\s+ogy\b", "Morphology"),
        (r"\bSynon\s+ymy\b", "Synonymy"),
        (r"\bEtymol\s+ogy\b", "Etymology"),
        (r"\bPreserv\s+ation\b", "Preservation"),
        (r"\bLocal\s+ity\b", "Locality"),
        (r"\bMater\s+ial\b", "Material"),
        (r"\bDiscuss\s+ion\b", "Discussion"),
        (r"\bCompar\s+ison\b", "Comparison"),
        (r"\bRem\s+arks\b", "Remarks"),
    ]
    for _pat, _repl in _KNOWN_SECTION_SPLITS:
        text = re.sub(_pat, _repl, text)
    return text.strip()


def extract_pages_from_file(path: pathlib.Path, settings: Dict) -> List[PageText]:
    ext = path.suffix.lower()
    if ext == ".txt":
        return _extract_txt(path)
    if ext == ".docx":
        return _extract_docx(path)
    if ext == ".pdf":
        return _extract_pdf(path, settings)
    return [PageText(1, "", "unsupported", f"Extension {ext} not supported")]


def _extract_txt(path: pathlib.Path) -> List[PageText]:
    text = path.read_text(encoding="utf-8", errors="replace")
    chunks = text.split("\x0c") if "\x0c" in text else [
        text[i:i+3000] for i in range(0, len(text), 3000)
    ]
    return [PageText(i+1, _paragraphize(c), "txt") for i, c in enumerate(chunks) if c.strip()]


def _extract_docx(path: pathlib.Path) -> List[PageText]:
    if not HAS_DOCX:
        return [PageText(1, "", "docx_failed", "python-docx not installed")]
    try:
        doc = _DocxDoc(str(path))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for tbl in doc.tables:
            for row in tbl.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        full = _paragraphize("\n".join(parts))
        full = _normalize_chinese_systematic_text(full)

        # ── Inteligentní chunking na hranicích paragrafu ──────────────────
        # Místo tvrdého chunking každých 3000 znaku rozdělíme na logické
        # bloky (odstavce oddělené prázdným řádkem). Každý blok se stane
        # samostatnou "stránkou". Bloky delší než MAX_CHUNK se dále rozdělí
        # jen tehdy, pokud neexists kratší přirozená hranice.
        # Výhoda: pro čínské/ruské Documenty každý taxonomický záznam
        # (oddělený prázdnou řádkou) tvoří samostatnou stránku — PaleoN pak
        # správně detekuje kandidáty a extrahuje bloky bez cross-page šumu.
        MAX_CHUNK = 4000
        raw_paras = [p.strip() for p in full.split("\n\n") if p.strip()]
        pages: List[str] = []
        current = ""
        for para in raw_paras:
            if not current:
                current = para
            elif len(current) + len(para) + 2 <= MAX_CHUNK:
                current = current + "\n\n" + para
            else:
                pages.append(current)
                current = para
        if current:
            pages.append(current)

        # Fallback: pokud bychom skončili s nula stránkami, použij puvodní split
        if not pages:
            pages = [full[i:i+3000] for i in range(0, len(full), 3000)]

        return [PageText(i + 1, c, "docx") for i, c in enumerate(pages) if c.strip()]
    except Exception as e:
        return [PageText(1, "", "docx_failed", str(e))]


def _extract_pdf(path: pathlib.Path, settings: Dict) -> List[PageText]:
    if HAS_FITZ:
        try:
            return _extract_pdf_fitz(path, settings)
        except Exception as e:
            logging.warning(f"PyMuPDF failed: {e}")
    if HAS_PDFPLUMBER:
        try:
            return _extract_pdf_plumber(path, settings)
        except Exception as e:
            return [PageText(1, "", "pdf_failed", str(e))]
    return [PageText(1, "", "no_pdf_lib", "Install pymupdf or pdfplumber")]


def _extract_pdf_fitz(path: pathlib.Path, settings: Dict) -> List[PageText]:
    pdf = fitz.open(str(path))
    out: List[PageText] = []
    min_chars = int(settings.get("pdf_min_chars", 80))
    for i, page in enumerate(pdf):
        txt, note = reading_order(page, settings)
        method = "pymupdf_blocks"
        if len(txt.strip()) < min_chars and settings.get("ocr_enabled") and HAS_TESSERACT:
            otxt, onote = _ocr_page(path, i, settings)
            if otxt.strip():
                txt, note, method = otxt, onote, "tesseract_ocr"
        out.append(PageText(i+1, _paragraphize(txt), method, NOT_PROVIDED, note))
    pdf.close()
    return out


def _extract_pdf_plumber(path: pathlib.Path, settings: Dict) -> List[PageText]:
    out: List[PageText] = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages):
            txt = page.extract_text() or ""
            out.append(PageText(i+1, _paragraphize(HYPHEN_BREAK_RE.sub(r"\1\2", txt)),
                                "pdfplumber"))
    return out


def _render_page_image(path: pathlib.Path, page_idx: int, dpi: int = 300) -> bytes:
    """Renders a PDF page as PNG bytes (high DPI for OCR)."""
    pdf  = fitz.open(str(path))
    page = pdf.load_page(page_idx)
    zoom = dpi / 72.0
    pix  = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    png  = pix.tobytes("png")
    pdf.close()
    return png


def _ocr_page(path: pathlib.Path, page_idx: int, settings: Dict) -> Tuple[str, str]:
    """
    Multi-engine OCR of a PDF page.
    Order: easyocr (if available and GPU) → PyMuPDF built-in OCR
    → pytesseract → chyba.
    Vrací (text, metoda_label).
    """
    errors = []

    # ── 1. easyocr (žádná externí binárka, GPU-akcelerované) ──────────────
    if HAS_EASYOCR:
        try:
            lang_raw = settings.get("ocr_languages", "en")
            # Převod tesseract kódu → easyocr (eng→en, ces→cs, deu→de …)
            _tess_to_easy = {"eng":"en","ces":"cs","deu":"de","fra":"fr",
                             "rus":"ru","chi_sim":"ch_sim","jpn":"ja"}
            easy_langs = []
            for part in lang_raw.replace("+", ",").split(","):
                easy_langs.append(_tess_to_easy.get(part.strip(), part.strip()))
            easy_langs = list(dict.fromkeys(easy_langs))[:3]  # max 3, deduplikace
            if not easy_langs:
                easy_langs = ["en"]
            # Lazy init readeru (sdílený přes session_state kvuli výkonu)
            _cache_key = f"_easyocr_reader_{'_'.join(easy_langs)}"
            if _cache_key not in st.session_state:
                st.session_state[_cache_key] = _easyocr.Reader(
                    easy_langs, gpu=True, verbose=False)
            reader = st.session_state[_cache_key]
            png = _render_page_image(path, page_idx, dpi=300)
            result = reader.readtext(png, detail=0, paragraph=True)
            txt = "\n".join(result)
            if txt.strip():
                return _paragraphize(txt), "easyocr"
        except Exception as exc:
            errors.append(f"easyocr: {exc}")

    # ── 2. PyMuPDF built-in OCR (vyžaduje tesseract, ale přes fitz API) ──
    if HAS_FITZ:
        try:
            lang_fitz = settings.get("ocr_languages", "eng").split("+")[0]
            pdf  = fitz.open(str(path))
            page = pdf.load_page(page_idx)
            tp   = page.get_textpage_ocr(language=lang_fitz, dpi=300, full=True)
            txt  = page.get_text(textpage=tp)
            pdf.close()
            if txt.strip():
                return _paragraphize(txt), "fitz_ocr"
        except Exception as exc:
            errors.append(f"fitz_ocr: {exc}")

    # ── 3. pytesseract (přímé volání) ─────────────────────────────────────
    if HAS_TESSERACT:
        try:
            png  = _render_page_image(path, page_idx, dpi=300)
            img  = _PILImage.open(io.BytesIO(png))
            lang = settings.get("ocr_languages", "eng")
            cfg  = "--oem 3 --psm 6"
            txt  = pytesseract.image_to_string(img, lang=lang, config=cfg)
            if txt.strip():
                return _paragraphize(txt), "tesseract_ocr"
        except Exception as exc:
            errors.append(f"tesseract: {exc}")

    return "", "ocr_failed: " + " | ".join(errors)


# ══════════════════════════════════════════════════════════════════════════════
# DETEKČNÍ ENGINE  (vylepšená verze)
# ══════════════════════════════════════════════════════════════════════════════

# ── Druhotná segmentace (re-segmentace) ────────────────────────────────────────
# U nativních PDF dělí PyMuPDF text na čisté odstavcové bloky automaticky.
# U OCR'd/skenovaných PDF je granularita bloku hrubší — "Class:", "Order:",
# "Family:", "Genus Name Autor, Rok" se mohou ocitnout SLOUČENÉ v jediném
# odstavci/bloku. Tato funkce takový blok dodatečně rozseká podle
# strukturálních kotev (rank-labely, schema-labely, popisky obrázku,
# "n. sp."/"sp. nov." nadpisy), takže detekce kandidátu dostane granularitu
# nezávislou na kvalitě PDF extrakce.

# Rank-labely jako kotvy uprostřed textu — "Genus Nophrotheca Marek, 1966"
# uprostřed odstavce signalizuje začátek nového nadpisu.
# Samostatný "Class:"/"Order:"/"Family:" label — pokud takový stojí TĚSNĚ
# před nadpisem rodu, patří logicky k NĚMU (viz vzorový Document:
# "Family: GRACILITHECIDAE Sysoev, 1972" je součástí záznamu rodu Gracilitheca).
_RANK_PREFIX_LABEL_RE = re.compile(
    r"^(Class|Order|Family|Subfamily|Superfamily|Tribe|Subtribe|Семейство|Отряд|Класс|Надсемейство|科|目|纲|门)"
    r"\\s*:?\\s+[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽА-ЯЁ]",
    re.UNICODE)

_RESPLIT_RANK_RE = re.compile(
    r"(?<![\w])"
    r"(Class|Order|Family|Subfamily|Superfamily|Tribe|Subtribe|"
    r"Genus|Subgenus|Species|Subspecies|Suborder|Infraorder|Phylum|Subphylum|"
    r"Kmen|Třída|Podtřída|Řád|Podřád|Čeleď|Podčeleď|Rod|Podrod|Druh|Poddruh|"
    r"Кмен|Класс|Отряд|Семейство|Род|Вид)"
    r"\s*:?\s+(?=[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜА-ЯЁ])",
    re.UNICODE)

# Resplit na koncovkách (pro Documenty bez explicitních rank-labelu):
# -ida (řád), -morpha (třída), -idae (čeleď) jako kotvy pro nový blok.
# Záměrně oddělené od _RESPLIT_RANK_RE — aplikuje se jako sekundární pruchod.
_RESPLIT_SUFFIX_RE = re.compile(
    r"(?<!\w)"
    r"(?=[A-Z][A-Za-z\-]+"
    r"(?:ecida|(?<!i)ida|morpha|phyta|idae|inae|oidea)"
    r"(?:\s+[A-Z][A-Za-z]+,?\s+\d{4}))"
    r"(?=[A-Z])",
    re.UNICODE)

# Popisky obrázku/tabulek jako kotvy — "Fig. 2. Nophrotheca sophia..." s tečkou
# za číslem a velkým písmenem hned po ní (na rozdíl od inline odkazu "(Fig. 3D, E)").
_RESPLIT_CAPTION_RE = re.compile(
    r"(?<![\w(])"
    r"(?:Fig(?:s)?\.?|Figure(?:s)?|Plate(?:s)?|Pl\.?|Table(?:s)?|Tab\.?)"
    r"\s*\d+[A-Za-z]?\.\s+(?=[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ])",
    re.UNICODE)

# Nový-taxon marker jako kotva — POZOR: "n. sp." / "sp. nov." se v textu
# objevuje JAK ve vlastním nadpisu, TAK opakovaně ve zpětných odkazech uvnitř
# Discussion/Occurrence sections stejného druhu ("Nophrotheca sophia n. sp.
# clearly fits with concept of…", "…together with the Nophrotheca betula…").
# Klíčové rozlišení: skutečný NADPIS je vždy buď na konci kusu textu, or
# bezprostředně následován "(Figs …)" — zatímco zpětný odkaz pokračuje další
# prózou. Lookahead toto vynucuje a eliminuje naprostou většinu falešných
# rozdělení bez nutnosti seznamu zakázaných slov.
_RESPLIT_NEWTAXON_RE = re.compile(
    r"(?<![\w(])"
    r"([A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ][a-záčďéěíňóřšťúuýž]+\s+"
    r"[a-záčďéěíňóřšťúuýž]+\s+(?:n\.\s*sp\.|sp\.\s*nov\.))"
    r"(?=\s*\(|\s*$|\s*\n)",
    re.UNICODE)


def _resplit_unit_text(text: str, schema_label_pattern: Optional[re.Pattern]) -> List[str]:
    """
    Splits a single paragraph into sub-paragraphs based on structural anchors.
    Returns a listnam textových úseku (min. 1 prvek = puvodní text beze změny,
    pokud žádná kotva nebyla nalezena).

    DVOUFÁZOVÉ ZPRACOVÁNÍ (duležité pořadí):
      Fáze 1 — nejdřív rozdělit POUZE podle popisku obrázku (Fig./Plate/Table).
               Popisek typically zní "Fig. 2. Nophrotheca sophia n. sp.: A, B. …" —
               pokud bychom dál dělili I uvnitř tohoto kusu podle rank-labelu
               or n. sp. markeru, vznikl by z "Nophrotheca sophia n. sp.: A, B…"
               falešný kandidát nadpisu (vypadá přesně jako skutečný nadpis druhu).
      Fáze 2 — teprve kusy, které NEJSOU popiskem obrázku, dál rozdělit podle
               rank-labelu, schema-labelu a n. sp./sp. nov. markeru.
    """
    # ── Fáze 1: rozdělení podle popisku obrázku ──────────────────────────────
    caption_cuts = sorted(
        m.start() for m in _RESPLIT_CAPTION_RE.finditer(text)
    )
    caption_cuts = [c for c in caption_cuts if c != 0]

    if caption_cuts:
        stage1: List[str] = []
        prev = 0
        for cp in caption_cuts:
            piece = text[prev:cp].strip()
            if piece:
                stage1.append(piece)
            prev = cp
        tail = text[prev:].strip()
        if tail:
            stage1.append(tail)
    else:
        stage1 = [text]

    # ── Fáze 2: další dělení jen u ne-popiskových kusu ───────────────────────
    final_pieces: List[str] = []
    for piece in stage1:
        if CAPTION_RE.match(piece):
            # Toto JE popisek obrázku — necháme ho vcelku, dál nedělíme.
            final_pieces.append(piece)
            continue

        cut_points: set = set()
        for m in _RESPLIT_RANK_RE.finditer(piece):
            cut_points.add(m.start(1))
        for m in _RESPLIT_NEWTAXON_RE.finditer(piece):
            cut_points.add(m.start(1))
        if schema_label_pattern is not None:
            for m in schema_label_pattern.finditer(piece):
                cut_points.add(m.start(1))
        cut_points.discard(0)

        if not cut_points:
            final_pieces.append(piece)
            continue

        sorted_cuts = sorted(cut_points)
        prev2 = 0
        for cp in sorted_cuts:
            sub = piece[prev2:cp].strip()
            if sub:
                final_pieces.append(sub)
            prev2 = cp
        tail2 = piece[prev2:].strip()
        if tail2:
            final_pieces.append(tail2)

    return final_pieces if final_pieces else [text]


def build_text_units(pages: List[PageText]) -> List[TextUnit]:
    """
    Splits pages into TextUnit objects (paragraphs).
    Applies secondary re-segmentation na strukturální kotvy — robustní i pro
    OCR'd PDF, kde PyMuPDF/OCR vrstva nedělí text na čisté odstavce.
    """
    # Schema-label regex pro re-segmentaci (stejné labely jako section mapping)
    try:
        schema = load_schema()
        schema_pat, _ = build_section_regex(schema)
    except Exception:
        schema_pat = None

    units: List[TextUnit] = []
    for pg in pages:
        paras = re.split(r"\n{2,}", pg.text)
        for para in paras:
            para = para.strip()
            if not para:
                continue
            sub_pieces = _resplit_unit_text(para, schema_pat)
            for piece in sub_pieces:
                piece = piece.strip()
                if not piece:
                    continue
                units.append(TextUnit(
                    page_number=pg.page_number,
                    text=piece,
                ))
    return units


def assign_zone_flags(units: List[TextUnit]) -> None:
    """
    Assigns zone_flags to each TextUnit based on context.
    Modifies units in-place.

    Nový flag "pre_systematic": pokud Document OBSAHUJE explicitní systematickou
    sekci, bloky PŘED ní dostanou tento flag → score_candidate je penalizuje
    výrazněji (pravděpodobnost FP z úvodu/metodiky je vyšší).
    """
    # Předsken stránek: odlišit reálný text od obsahu / indexu Documentu.
    # Duležité: "Systematic paleontology" v obsahu NESMÍ aktivovat systematic zónu.
    _SYS_RE = get_systematic_re()  # dynamicky z TSV or fallback
    _index_pages: set = set()          # abecední rejstřík na konci knihy
    _doc_index_pages: set = set()      # obsah / front-matter index Documentu
    _page_lines: Dict[int, List[str]] = {}
    for u in units:
        _page_lines.setdefault(u.page_number, []).extend(u.text.splitlines())

    for _pn, _lines in _page_lines.items():
        _nonempty = [l for l in _lines if l.strip()]
        if not _nonempty:
            continue
        _book_matches = sum(1 for l in _nonempty if _BOOK_INDEX_LINE_RE.match(l.strip()))
        if _book_matches / len(_nonempty) >= 0.70:
            _index_pages.add(_pn)
        if _looks_like_document_index_page(_lines):
            _doc_index_pages.add(_pn)

    has_systematic_section = any(
        _SYS_RE.search(u.text.strip()) and u.page_number not in _doc_index_pages
        for u in units
    )

    in_systematic = False
    systematic_encountered = False
    in_references  = False
    in_synonymy    = False

    for i, u in enumerate(units):
        flags: List[str] = []
        t = u.text.strip()

        # Abecední index / rejstřík → přeskočit při detekci kandidátu
        if u.page_number in _index_pages:
            flags.append("book_index")
        # Obsah / index Documentu ve front matter → nikdy nevytvářet kandidáty
        # a nikdy z něj nespouštět systematickou ani references zónu.
        in_document_index_page = u.page_number in _doc_index_pages
        if in_document_index_page:
            flags.append("document_index")
            u.zone_flags = ",".join(flags)
            continue

        # Systematická sekce — jen v reálné textové části, ne v obsahu.
        if _SYS_RE.search(t):
            in_systematic = True
            systematic_encountered = True
        if END_REGION_RE.search(t):
            in_references = True
            in_systematic = False
        # Reset synonymy po prázdném řádku or novém nadpisu
        if in_synonymy and (not SYNONYMY_LINE_RE.match(t)):
            in_synonymy = False

        if in_systematic:
            flags.append("systematic")
        if in_references:
            flags.append("references")
        # Blok PŘED systematickou sections (pokud vubec exists) → extra penalizace
        if has_systematic_section and not systematic_encountered and not in_systematic:
            flags.append("pre_systematic")

        # Caption
        if CAPTION_RE.match(t):
            flags.append("caption")
        # Synonymy řádek — "synonymy" se přidá jen JEDNOU (dřív se tu
        # přidávalo dvakrát: jednou při přímé shodě SYNONYMY_LINE_RE a
        # znovu vzápětí přes "if in_synonymy" — neškodilo to, protože
        # score_candidate flags převádí na set(), ale bylo to matoucí
        # a zbytečné).
        if SYNONYMY_LINE_RE.match(t):
            in_synonymy = True
        if in_synonymy:
            flags.append("synonymy")
        # Reference-like
        if REF_LIKE_RE.match(t) and len(t.split()) > 4:
            flags.append("reference_like")

        u.zone_flags = ",".join(flags)


def _clean_candidate(text: str) -> str:
    """Removes trailing parentheses with fig/pl/tab references."""
    text = re.sub(r"\s*\([^)]{1,80}(?:fig|pl|plate|table)[^)]*\)\s*$", "", text, flags=re.I)
    return text.strip()


# Druhé slovo (domnělý druhový epiteton) nesmí být spojka/předložka —
# zabraňuje falešným shodám typu "Gracilitheca and Nophrotheca..." (name článku).
EPITHET_STOPWORDS = {
    # Gramatické funkční slova
    "and", "or", "in", "of", "the", "with", "from", "to", "et", "und",
    "is", "are", "was", "were", "has", "have", "by", "at", "on", "for",
    # Morfologické podstatné jméno jako druhý člen falešného binomia
    "layer", "layers", "region", "regions", "surface", "surfaces",
    "section", "sections", "portion", "portions", "margin", "margins",
    "edge", "edges", "line", "lines", "groove", "grooves", "rib", "ribs",
    "zone", "zones", "band", "bands", "pore", "pores", "node", "nodes",
    "fold", "folds", "lobe", "lobes", "plate", "plates", "scar", "scars",
    "angle", "angles", "axis", "axes", "wall", "walls", "base", "apex",
    "specimens", "specimen", "material", "collection", "collections",
    "formation", "formations", "limestone", "sandstone", "shale",
    # Angličtina: druhá slova v "adjective noun" nadpisech sections
    "characteristics", "characteristic", "description", "discussion",
    "remarks", "comment", "comments", "comparison", "comparisons",
    "assignment", "affinity", "affinities", "identification",
    "occurrence", "occurrences", "distribution", "range", "ranges",
    "preservation", "collecting", "registration", "methods",
}

def _match_name(text: str, prev_rank: str = "") -> Tuple[Optional[str], str, str]:
    """
    Tries binomial, generic and family-group patterns.
    Returns (name, rank, pattern) or (None, '', '').
    Podporuje latinu, Cyrilici a vzor 'Rank Name Author, Year'.
    """
    t = _clean_candidate(text)
    words = t.split()
    first_word = words[0].strip("?.,;:\"'").lower() if words else ""
    second_word = words[1].strip("?.,;:\"'").lower() if len(words) > 1 else ""
    if second_word in EPITHET_STOPWORDS:
        return None, "", ""

    # Vzor "Family Hyolithidae Smith, 1900" or "Genus Examplius Smith, 1900"
    # RANK_LABEL_RE → použij rank jako Context a name jako target
    rm = RANK_LABEL_RE.match(t)
    if rm:
        if rm.group(2).strip():
            detected_rank = _canonical_rank_label(rm.group(1))
            sub = rm.group(2).strip()
            sub_name, _, sub_pat = _match_name(sub, detected_rank)
            if sub_name:
                return sub_name, detected_rank, sub_pat
        # Slovo, které by JINAK bylo rank-labelem (Class/Order/Family/Genus/…),
        # se tu použilo jako běžné slovo na začátku věty ("Class discussion
        # of hyolith affinities remains controversial.") a pokus interpretovat
        # zbytek jako jméno taxonu selhal. Taková věta NIKDY není skutečný
        # nadpis — pokud bychom tu nekončili, mohla by dál projít přes
        # SPECIES_RE níže jako falešné binomium (Velké-slovo + malé-slovo
        # strukturálně vypadá jako "Genus species", i když jde o obyčejnou
        # anglickou/českou/atd. větu).
        return None, "", ""

    # Stopwords check (až po pokusu s rank-label prefixem)
    if first_word in TAXON_STOPWORDS:
        return None, "", ""

    # Cyrillic binomial (ruské, ukrajinské názvy)
    cm = CYRILLIC_SPECIES_RE.match(t)
    if cm:
        name = cm.group("name").strip()
        if name:
            return name, prev_rank or "Species", "cyrillic_species"

    # Čínský nadpis: čínské jméno + latinský name (+ rank znak)
    if t and "\u4e00" <= t[0] <= "\u9fff":   # začíná čínským znakem
        # Přeskočit etymologické řádky: "属名来源 Eo——始、古之意"
        if _CHINESE_ETYMOLOGY_RE.search(t):
            return None, "", ""
        zm = CHINESE_TAXON_RE.match(t)
        if zm:
            name = zm.group("name").strip()
            # ── Oprava OCR artifact: "A mbrolinevitus" → "Ambrolinevitus" ──────
            # V čínských pracích OCR někdy rozlomí slovo na hranici řádku tak,
            # že první písmeno/slabika je oddělena mezerou od zbytku slova.
            # Pravidlo: začáteční 1–2 písmena + mezera + pokračování (3+ znaku) = jeden celek.
            name = re.sub(r"^([A-Z]{1,2}) ([a-z]{3,})", r"\1\2", name)
            # Opravit rozlomení druhového epitetu: "p latyp" → "platyp" (uvnitř jména)
            name = re.sub(r"([A-Za-z]{2,}) ([a-z]{1,3}) ([a-z]{3,})\b",
                          lambda m: m.group(1) + " " + m.group(2) + m.group(3),
                          name)
            if name and name.lower() not in TAXON_STOPWORDS:
                rank_char = zm.group("rankchar")
                zh_rank = _CHINESE_RANK_CHARS.get(rank_char or "", "")
                # Rank: z rank znaku, jinak podle přítomnosti druhového epitetu
                if not zh_rank:
                    zh_rank = "Species" if " " in name else "Genus"
                return name, zh_rank or prev_rank, "chinese_taxon"

    for rx, rank_guess, pname in [
        (FAMILY_RE,   "Family",   "family"),
        (ORDER_RE,    "Order",    "order"),
        (CLASS_RE,    "Class",    "class_taxon"),
        (SPECIES_RE,  prev_rank or "Species",  "species"),
        (GENUS_SP_RE, prev_rank or "Species",  "genus_sp"),
    ]:
        m = rx.match(t)
        if m:
            name = m.group("name").strip()
            if name.lower() not in TAXON_STOPWORDS:
                return name, rank_guess, pname

    # Samostatný "Rod Autor, Rok" bez druhového epitetu (genus-level nadpis).
    # Striktně ukotveno na celý text (whole-line match), takže riziko
    # falešné shody s běžnou větou je minimální.
    gam = GENUS_AUTHOR_RE.match(t)
    if gam:
        bare_name = gam.group("name").strip()
        if bare_name.lower() not in TAXON_STOPWORDS:
            # Vrátit CELÝ zápis "Rod Autor, Rok" (ne jen holý rod) —
            # konzistentní s formátem vzorových records ("Gracilitheca Sysoev, 1968").
            full_name = gam.group(0).strip().rstrip(".")
            return full_name, prev_rank or "Genus", "genus_author"

    # Jednoslovný rod jen se silnou podmínkou
    if prev_rank.lower() in ("genus", "subgenus"):
        m = GENUS_RE.match(t)
        if m:
            name = m.group("name").strip()
            if name.lower() not in TAXON_STOPWORDS:
                return name, "Genus", "genus"

    # Name s explicitním rankovým suffixem bez autora:
    # "CRISPATELLA gen.", "HYOLITHIDAE fam. nov.", "Examplius gen. nov."
    _RANK_SUFFIX_RE = re.compile(
        r"^([A-ZА-ЯЁ一-鿿][A-Za-zА-ЯЁа-яёčšžřéíúuý一-鿿]+)\s+"
        r"(gen|fam|ord|cl|sp)\.?\s*(?:nov\.?)?$",
        re.IGNORECASE | re.UNICODE)
    _sfx_m = _RANK_SUFFIX_RE.match(t)
    if _sfx_m:
        _sfx_name = _sfx_m.group(1).strip()
        _sfx_type = _sfx_m.group(2).lower()
        _sfx_rank = {"gen": "Genus", "fam": "Family", "ord": "Order",
                     "cl": "Class", "sp": "Species"}.get(_sfx_type, "")
        if _sfx_rank and _sfx_name.lower() not in TAXON_STOPWORDS:
            return _sfx_name, prev_rank or _sfx_rank, "rank_suffix"

    return None, "", ""


def _is_ordinary_sentence(text: str) -> bool:
    """
    Returns True if the text is probably NOT a taxon heading but a regular sentence
    or descriptionný výraz. Slouží jako penalizace ve score_candidate.

    Rozšířeno po analýze reálných exportu (2026-07):
    - Zachytí krátká 2-slovná slovní spojení, která nejsou binomiem
      ("Faunal comparison", "Monoclaviculate operculum") — SPECIES_RE
      je chytí jako Genus+species, ale jsou to morfologické výrazy/nadpisy.
    - Zachytí věty zjevně začínající participiem or přídavným jménem
      ("Hyolithids characterized by…", "Unfigured incomplete compressed…").
    """
    t = text.strip()

    # Příliš dlouhé na nadpis
    if len(t) > 200:
        return True

    words = t.split()
    n_words = len(words)

    # Věta s tečkou na konci a více než 6 slovy bez nomenklaturního aktu
    if t.endswith(".") and n_words > 6 and not NEW_TAXON_RE.search(t):
        return True

    # Slova indikující větu (kopula, pomocná slovesa)
    if re.search(r"\b(is|are|was|were|has|have|known|found|consists|shows|"
                 r"comes|occurs|represents|suggests|indicates|reveals)\b", t, re.I):
        if n_words > 5:
            return True

    # Dvě slova — buď legitimní rod/druh (chytí SPECIES_RE) or morfologická fráze.
    # Rozlišení: pravý druhový epiteton je latinský (jen písmena, pomlčky, tečky).
    # Fráze jako "Faunal comparison", "Monoclaviculate operculum" nesplňují
    # kritérium latinského epitetu — slova jsou příliš obecná (Angličtina/adj.).
    if n_words == 2:
        second = words[1].lower().rstrip(".,;:")
        # Druhé slovo je běžné anglické slovo, ne pravý druhový epiteton
        _ENG_NONADJ = {
            "comparison", "operculum", "used", "incomplete", "other",
            "conch", "aperture", "specimen", "material", "fauna",
            "section", "plate", "type", "form", "stage",
            "characterized", "described", "discussed", "assigned",
            "referred", "considered", "observed", "reported", "figured",
        }
        if second in _ENG_NONADJ:
            return True
        # Druhé slovo příliš krátké pro druhový epiteton (min. 3 znaky)
        if len(second) < 3 and not second.endswith("."):
            return True

    # Particip or adjektivum na prvním místě (typický začátek diagnózy, ne nadpisu)
    # "Characterized by", "Distinguished from", "Represented by" atd.
    _PARTICIP_STARTERS = {
        "characterized", "distinguished", "separated", "represented",
        "described", "assigned", "attributed", "reported", "referred",
        "consisting", "bearing", "showing", "differing", "resembling",
        "unfigured", "compressed", "deformed", "poorly", "well",
    }
    if words and words[0].lower() in _PARTICIP_STARTERS:
        return True

    return False


_STRONG_SECTION_RE_CACHE: Optional[re.Pattern] = None
_PURE_LABEL_RE_CACHE: Optional[re.Pattern] = None
# Cache pro nejdražší schema-odvozené struktury (invalidace v reload_schema).
# build_section_regex iteruje 1300+ řádku a kompiluje obří regex — bez cache
# se to dělo při KAŽDÉM map_sections_from_block/annotate_raw_block.
_SECTION_REGEX_CACHE: Optional[Tuple[re.Pattern, Dict[str, str]]] = None
_STRONG_FIELDS_CACHE: Optional[set] = None
_STRONG_LABELS_CACHE: Optional[set] = None
_LABEL_ONLY_PAT_CACHE: Optional[re.Pattern] = None
_SCHEMA_DERIVED_RE_LOCK = threading.Lock()


def _invalidate_schema_derived_regex_cache() -> None:
    """
    Invalidates cached regexes derived from the schema (strong-section and
    pure-section-label). MUSÍ se zavolat vždy, když se mění samotné schéma
    (viz reload_schema) — jinak by se dál používal starý, zastaralý regex.
    """
    global _STRONG_SECTION_RE_CACHE, _PURE_LABEL_RE_CACHE
    global _SECTION_REGEX_CACHE, _STRONG_FIELDS_CACHE, _STRONG_LABELS_CACHE
    global _LABEL_ONLY_PAT_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        _STRONG_SECTION_RE_CACHE = None
        _PURE_LABEL_RE_CACHE = None
        _SECTION_REGEX_CACHE = None
        _STRONG_FIELDS_CACHE = None
        _STRONG_LABELS_CACHE = None
        _LABEL_ONLY_PAT_CACHE = None


def _get_strong_section_regex(schema: pd.DataFrame) -> re.Pattern:
    global _STRONG_SECTION_RE_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        if _STRONG_SECTION_RE_CACHE is not None:
            return _STRONG_SECTION_RE_CACHE
        labels = sorted(
            {str(row.get("label", "")).strip() for _, row in schema.iterrows()
             if str(row.get("is_strong", "0")).strip() == "1"
             and len(str(row.get("label", "")).strip()) >= 3},
            key=len, reverse=True)
        pat = (re.compile("|".join(re.escape(l) for l in labels), re.IGNORECASE | re.UNICODE)
               if labels else re.compile(r"(?!x)x"))  # nikdy nic nenajde
        _STRONG_SECTION_RE_CACHE = pat
        return pat


def _after_has_strong_section(after_text: str, strong_fields: set) -> bool:
    """
    Returns True if the texts after the candidate contain a strong section label.
    This is the most reliablelnější Contextový signál.

    VÝKON: dřív se tu při KAŽDÉM volání (tj. pro každého detekovaného
    kandidáta) iterovalo celé schéma přes pandas iterrows() a pro každý
    "strong" řádek se dělal samostatný re.search() — u schématu s ~200
    řádky (multilinguální TSV) to je 200 Python-úrovňových operací na
    kandidáta. Teď se jeden kombinovaný regex sestaví jen JEDNOU a
    cachuje (invalidace při změně schématu — viz reload_schema).
    """
    schema = load_schema()
    pat = _get_strong_section_regex(schema)
    return bool(pat.search(after_text))


# Znaky téměř nikdy se nevyskytující v normální taxonomické próze, ale
# typické pro OCR šum z map/legend/diagramu ("Cam bri: —sy ly [| srauseiee").
_GARBAGE_CHARS_RE = re.compile(r"[\[\]|<>]")


def _garbage_penalty(text: str) -> float:
    """
    Returns a penalty (0 to -0.50) based on the density of OCR noise in the text.
    Combines the presenceomnost neobvyklých znaku a podíl extrémně krátkých
    "slov" (typický projev OCR fragmentace map/legend/tabulek).
    """
    if _GARBAGE_CHARS_RE.search(text):
        return -0.50
    words = [w.strip(".,;:()") for w in text.split()]
    words = [w for w in words if w]
    if len(words) >= 3:
        very_short = sum(1 for w in words if len(w) <= 2 and not w.isdigit())
        if very_short / len(words) > 0.35:
            return -0.35
    return 0.0


def score_candidate(
    unit: TextUnit,
    name: str,
    rank: str,
    pattern: str,
    after_text: str,
    settings: Dict,
    strong_fields: set,
) -> Tuple[float, Dict]:
    """
    Weighted scoring of a candidate. Returns (score 0–1, debug dict).
    """
    flags = set(f for f in unit.zone_flags.split(",") if f)
    debug: Dict = {"bonuses": {}, "penalties": {}}
    score = 0.40   # základní Score za detekci jména

    # ── Bonusy ──────────────────────────────────────────────────────────────
    if "family" in pattern:
        score += 0.20;  debug["bonuses"]["family_pattern"] = 0.20
    elif "order" in pattern:
        score += 0.20;  debug["bonuses"]["order_pattern"] = 0.20
    elif "class_taxon" in pattern:
        score += 0.20;  debug["bonuses"]["class_pattern"] = 0.20
    elif "species" in pattern or "genus_sp" in pattern:
        score += 0.15;  debug["bonuses"]["species_pattern"] = 0.15
    elif "genus" in pattern:
        # Rodová úroveň ("Genus Examplius Sysoev, 1968" / bare "Examplius"
        # s rank Contextem) dřív NEDOSTÁVALA žádný strukturní bonus — jen
        # family a species patterny ho měly. Rodové nadpisy tak systematicky
        # skórovaly níž, přestože jsou stejně platnou hranicí bloku jako
        # family/species (viz extract_block_for_candidate). Bonus o 0.05
        # nižší než species, protože holé rodové jméno (bez binomia) je
        # o něco méně specifický signál.
        # P 1.3: holý rod + autor + rok → silnější signal než species pattern
        genus_bonus = 0.15 if pattern == "genus_author" else 0.10
        score += genus_bonus;  debug["bonuses"]["genus_pattern"] = genus_bonus

    if NEW_TAXON_RE.search(unit.text):
        # P 1.1: v krátkém nadpisu (<100 znaku) je n. sp. silnější signál
        nsp_bonus = 0.35 if len(unit.text.strip()) < 100 else 0.25
        score += nsp_bonus;  debug["bonuses"]["new_taxon_marker"] = nsp_bonus

    if OPEN_NOMEN_RE.search(unit.text):
        score += 0.05;  debug["bonuses"]["open_nomen"] = 0.05

    if rank:
        score += 0.10;  debug["bonuses"]["rank_context"] = 0.10

    if len(unit.text) <= 120:
        score += 0.10;  debug["bonuses"]["short_heading"] = 0.10

    if unit.is_bold or unit.is_italic:
        score += 0.10;  debug["bonuses"]["font_style"] = 0.10

    if "systematic" in flags:
        score += 0.15;  debug["bonuses"]["systematic_context"] = 0.15

    if _after_has_strong_section(after_text, strong_fields):
        score += 0.20;  debug["bonuses"]["strong_section_follows"] = 0.20

    if AUTHOR_YEAR_RE.search(unit.text):
        score += 0.10;  debug["bonuses"]["author_year"] = 0.10

    # Gazetteer (taxons.txt) — měkký bonus, NE tvrdý filtr. Nopřítomnost
    # jména v seznamu nic nepenalizuje (nově popisované druhy v aktuálním
    # článku tam logicky missing); přítomnost zvyšuje duvěru.
    genera, binomials = load_gazetteer()
    name_words = name.lower().split()
    gaz_bonus = float(settings.get("gazetteer_bonus", 0.10))
    if name_words:
        if len(name_words) >= 2 and f"{name_words[0]} {name_words[1]}" in binomials:
            score += gaz_bonus * 1.5
            debug["bonuses"]["gazetteer_binomial"] = round(gaz_bonus*1.5, 3)
        elif name_words[0] in genera:
            score += gaz_bonus
            debug["bonuses"]["gazetteer_genus"] = gaz_bonus

    # ── Penalizace suspektních OCR fragmentu ─────────────────────────────────
    # Krátká první slova (≤5 znaku) která nejsou v gazetteeru jsou typically
    # OCR-split artefakty ("Speci|men", "Diag|nosis", "Descr|iption") or
    # obecná anglická slova. Pokud je gazetteer neprázdný a jméno v něm missing,
    # aplikujeme výraznou penalizaci.
    _first_w = name_words[0] if name_words else ""
    _gaz_non_empty = bool(genera)
    if _first_w and len(_first_w) <= 5 and _gaz_non_empty and _first_w not in genera:
        score -= 0.45
        debug["penalties"]["short_unknown_genus"] = -0.45
    elif _first_w and len(_first_w) <= 4 and _first_w not in genera:
        # I bez gazetteeru: 4 a méně znaku je velmi suspicious (SPECIES_RE
        # vyžaduje ≥1 char rod – ale "Sp" "Gen" "Fam" apod. nejsou validní rody)
        score -= 0.35
        debug["penalties"]["very_short_genus"] = -0.35

    # ── Pokuty ───────────────────────────────────────────────────────────────
    garbage_pen = _garbage_penalty(unit.text)
    if garbage_pen < 0:
        score += garbage_pen
        debug["penalties"]["garbage_chars"] = garbage_pen

    if "references" in flags or "reference_like" in flags:
        score -= 0.50;  debug["penalties"]["references"] = -0.50

    if "book_index" in flags:
        score -= 0.85;  debug["penalties"]["book_index"] = -0.85

    if "caption" in flags:
        score -= 0.45;  debug["penalties"]["caption"] = -0.45

    if "synonymy" in flags:
        score -= 0.45;  debug["penalties"]["synonymy"] = -0.45

    if _is_ordinary_sentence(unit.text):
        score -= 0.30;  debug["penalties"]["ordinary_sentence"] = -0.30

    # Pozn.: dřív tu byla i podmínka "systematic_heading" not in flags —
    # ten flag ale NIKDE v assign_zone_flags() nevznikal (mrtvý kód, který
    # jen matoucně naznačoval funkčnost, jež neexistovala). Odpagesěno.
    if "systematic" not in flags:
        pen = float(settings.get("outside_systematic_penalty", 0.20))
        if "pre_systematic" in flags:
            # P 1.2: blok PŘED systematickou sections — agresivní penalizace.
            pen = min(1.0, pen + 0.45)   # místo 0.30
            debug["penalties"]["pre_systematic_extra"] = -0.45
            # Anulovat n. sp. bonus (v úvodu/abstraktu je to jen citace)
            if debug.get("bonuses", {}).get("new_taxon_marker"):
                score -= debug["bonuses"]["new_taxon_marker"]
                debug["penalties"]["pre_systematic_nsp_cancel"] = -debug["bonuses"]["new_taxon_marker"]
            # Anulovat gazetteer bonusy (taxa z gazetteeru se v úvodu citují)
            for gk in ("gazetteer_binomial", "gazetteer_genus"):
                if debug.get("bonuses", {}).get(gk):
                    score -= debug["bonuses"][gk]
                    debug["penalties"][f"pre_systematic_{gk}_cancel"] = -debug["bonuses"][gk]
        score -= pen;   debug["penalties"]["outside_systematic"] = -pen

    if len(name.split()) == 1 and rank.lower() not in ("genus","subgenus") and \
            not re.search(r"(idae|inae|ini|oidea)$", name, re.I):
        score -= 0.30;  debug["penalties"]["single_word_no_rank"] = -0.30

    score = round(max(0.0, min(1.0, score)), 4)
    debug["final_score"] = score
    return score, debug


def detect_candidates(
    document_id: int,
    pages: List[PageText],
    settings: Dict,
) -> Dict[str, Any]:
    """
    Main detector. Returns diagnostics and writes candidates to the DB.
    """
    if settings.get("use_chinese_goldset_indexing", True):
        try:
            zh_diag = _extract_goldset_style_chinese_treatments(document_id, pages, settings)
            if zh_diag:
                return zh_diag
        except Exception as exc:
            logging.warning(f"Chinese goldset-style indexing failed; falling back: {exc}")
    schema = load_schema()
    strong_fields = get_strong_fields(schema)
    units = build_text_units(pages)
    assign_zone_flags(units)

    min_conf = float(settings.get("taxon_min_confidence", 0.60))
    min_low  = float(settings.get("taxon_low_confidence", 0.45))

    diag = {
        "Pages": len(pages),
        "Text units": len(units),
        "Characters": sum(len(u.text) for u in units),
        "Systematic region": any("systematic" in u.zone_flags for u in units),
        "Candidates found": 0,
        "Accepted": 0,
        "Low-confidence": 0,
        "Rejected": 0,
        "Rejected detail": [],
    }

    con = db()
    # Pozn.: dřív se tu mazalo taxon_candidates JAKO PRVNÍ a teprve pak
    # occurrence_fields přes subquery na taxon_candidates — subquery po
    # smazání taxon_candidates už nic nenajde, takže occurrence_fields
    # (a term_matches) zustávaly osiřelé. _delete_candidates_for_document
    # maže ve správném pořadí (závislosti první).
    _delete_candidates_for_document(document_id, con=con)

    # Post-detekce deduplication: zakázat pre_systematic kandidáty jejichž
    # jméno se vyskytuje i v systematické zóně (jsou to citační zmínky).
    _pre_sys_seen: set = set()    # jména z pre_systematic zóny
    _sys_seen:     set = set()    # jména ze systematic zóny

    prev_rank = ""
    # Feature 3: rank stack pro parent-child hierarchii
    # P 5.1: rozšířit o subspecies
    _RANK_ORDER = ["phylum","subphylum","class","subclass","order","suborder",
                   "superfamily","family","subfamily","tribe","genus","subgenus",
                   "species","subspecies"]
    _rank_stack: Dict[str, Tuple[int, str]] = {}

    def _get_parent_for_rank(rank: str) -> Tuple[Optional[int], str, str]:
        r = rank.lower().split()[0] if rank else ""
        try:
            my_pos = _RANK_ORDER.index(r)
        except ValueError:
            my_pos = len(_RANK_ORDER)
        for pr in reversed(_RANK_ORDER[:my_pos]):
            if pr in _rank_stack:
                pid, pname = _rank_stack[pr]
                return pid, pname, pr
        return None, "", ""

    for i, u in enumerate(units):
        flags = set(f for f in u.zone_flags.split(",") if f)

        if CAPTION_RE.match(u.text.strip()):
            continue
        if "references" in flags:
            continue
        # Přeskočit stránky abecedního indexu i front-matter obsahu Documentu.
        # Candidates z nich nemají taxonomický blok — jsou to jen odkazy na stránky.
        if "book_index" in flags or "document_index" in flags:
            continue
        # Boilerplate guard – running header/footer junk z akademických PDF
        # (e.g. "Downloaded by [...] at ... 2013") nesmí nikdy vygenerovat kandidáta.
        if BOILERPLATE_RE.search(u.text.strip()):
            continue
        # Section label guard – nezaměnit "Diagnosis:" za nadpis taxonu
        if _is_pure_section_label(u.text.strip(), schema):
            continue
        # Synonymy entry guard
        if SYNONYMY_LINE_RE.match(u.text.strip()):
            continue

        # Rank-only řádek → Context pro příští
        rm = RANK_LABEL_RE.match(u.text.strip())
        if rm and not rm.group(2).strip():
            prev_rank = rm.group(1).title()
            continue

        name, rank, pattern = _match_name(u.text.strip(), prev_rank)
        if not name:
            continue
        # P 5.1: odvod rank z tvaru jména pokud missing
        # Rozšířeno: rozpoznává poddruh (trinomen), nadčeleď, řád i podrod.
        if not rank:
            _nm = name.strip()
            _words = _nm.split()
            # Odečíst případný "(Autor, rok)" pro počítání slov epitetu
            _epithet_words = [w for w in _words
                              if not w.startswith("(") and not w[0].isdigit()]
            if re.search(r"(oidea)$", _nm, re.I):
                rank = "Superfamily"
            elif re.search(r"(idae)$", _nm, re.I):
                rank = "Family"
            elif re.search(r"(inae)$", _nm, re.I):
                rank = "Subfamily"
            elif re.search(r"(ini)$", _nm, re.I) and _nm[0].isupper():
                rank = "Tribe"
            elif re.search(r"(iformes|ida|ina)$", _nm, re.I) and len(_words) == 1 \
                    and _nm[0].isupper():
                # řádové/podřádové zakončení (jednoslovné, velké počáteční)
                rank = "Order"
            elif "(" in _nm and ")" in _nm and len(_epithet_words) >= 2:
                # binomen s podrodem: "Genus (Subgenus) species"
                # → stále species-level, ale zachytíme jako Species
                rank = "Species"
            elif len([w for w in _epithet_words
                      if w and w[0].islower()]) >= 2 and len(_epithet_words) >= 3:
                # trinomen: "Genus epithet subepithet" (2 malá slova za rodem)
                # → poddruh
                rank = "Subspecies"
            elif " " in _nm:
                rank = "Species"
            elif len(_nm) > 3 and _nm[0].isupper() and _nm[1:].islower():
                rank = "Genus"

        diag["Candidates found"] += 1
        after = "\n".join(u2.text for u2 in units[i+1:i+7])
        before = "\n".join(u2.text for u2 in units[max(0, i-3):i])

        score, debug_info = score_candidate(u, name, rank, pattern, after, settings, strong_fields)
        # Sledovat jména dle zóny pro post-detekci deduplication
        _norm_name = name.strip().lower()
        _flags_set = set(f for f in u.zone_flags.split(",") if f)
        if "pre_systematic" in _flags_set:
            _pre_sys_seen.add(_norm_name)
        elif "systematic" in _flags_set:
            _sys_seen.add(_norm_name)

        # Strážce titulku článku: úplně PRVNÍ jednotka Documentu je téměř
        # vždy nadpis článku ("Gracilitheca astronauta n. sp. and Nophrotheca
        # sophia n. sp. (Hyolitha, Orthothecida) from the Cambrian…"), což
        # vypadá jako legitimní nadpis druhu, ale NENÍ jím. Vyžadovat extra
        # silný dukaz (nomenklaturní akt I gazetteer shoda současně).
        if i == 0:
            has_strong_evidence = (
                debug_info.get("bonuses", {}).get("new_taxon_marker") and
                ("gazetteer_genus" in debug_info.get("bonuses", {}) or
                 "gazetteer_binomial" in debug_info.get("bonuses", {}))
            )
            if not has_strong_evidence:
                score = round(max(0.0, score - 0.35), 4)
                debug_info.setdefault("penalties", {})["likely_article_title"] = -0.35

        # P 1.2: pre_systematic s nízkým Score → rovnou rejected
        if "pre_systematic" in _flags_set and score < 0.50:
            score = round(max(0.0, score - 0.10), 4)  # extra trest pro jistotu
            debug_info.setdefault("penalties", {})["pre_systematic_autorej"] = -0.10
        if score >= min_conf:
            status = "pending"
            diag["Accepted"] += 1
        elif score >= min_low:
            status = "low_confidence"
            diag["Low-confidence"] += 1
        else:
            diag["Rejected"] += 1
            diag["Rejected detail"].append({
                "text": u.text[:80], "page": u.page_number,
                "score": score, "debug": debug_info
            })
            # POZOR: prev_rank se ZÁMĚRNĚ NEMAŽE tady. Dřív se mazal hned
            # při jakémkoli úspěšném _match_name() zásahu, BEZ OHLEDU na
            # finální Score — takže když mezi "Family:" a skutečným rodovým
            # nadpisem ležela jednotka, kterou _match_name() omylem (byť jen
            # slabě) rozpoznal jako jméno, ale která byla nakonec zamítnuta
            # (rejected), rank Context se ztratil a skutečný nadpis o pár
            # jednotek dál už žádný rank Context neměl. Rank Context teď
            # přežívá až do PRVNÍHO kandidáta, který byl skutečně přijat
            # (pending/low_confidence) — viz větev níže.
            continue

        # Kandidát byl přijat (pending/low_confidence) → rank Context je
        # "spotřebován" a čeká na další rank-label or taxon nadpis.
        prev_rank = ""

        _pid, _pname, _prank = _get_parent_for_rank(rank or "")
        cur = con.execute("""
            INSERT INTO taxon_candidates
            (document_id, taxon_name, rank_guess, confidence, status,
             heading_text, context_before, context_after, page_start,
             unit_index, debug_json, created_at,
             parent_taxon_name, parent_rank, parent_id)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (document_id, name, rank or NOT_PROVIDED, score, status,
              u.text[:500], before[:500], after[:500],
              u.page_number, i, json.dumps(debug_info, ensure_ascii=False),
              datetime.now().isoformat(), _pname, _prank, _pid))
        _cid = cur.lastrowid
        _nrank = (rank or "").lower().split()[0]
        if _nrank in _RANK_ORDER:
            _rank_stack[_nrank] = (_cid, name)
            for _lr in _RANK_ORDER[_RANK_ORDER.index(_nrank)+1:]:
                _rank_stack.pop(_lr, None)

    con.commit()
    con.close()

    # ── Druhý pruchod: extrakce bloku a automatické mapování sections ───────────
    # Pro VŠECHNY nově detekované kandidáty (včetně rejected):
    #   1. Extrahujeme verbatim blok textu (z již cachovaných text units).
    #   2. Uložíme block_text přímo do taxon_candidates.
    #   3. Spustíme auto_map_candidate_fields — vyplníme prázdná pole
    #      (metoda='auto'); manuálně editovaná pole se nepřepisují.
    # Text units jsou stále v cache z tohoto volání build_text_units()
    # → druhý pruchod je výcountně levný.
    try:
        con2 = db()
        new_cands = con2.execute(
            "SELECT id, unit_index, page_start FROM taxon_candidates "
            "WHERE document_id=? ORDER BY unit_index",
            (document_id,)
        ).fetchall()
        # Zjistit jazyk Documentu pro překladový pruchod
        doc_lang_row = con2.execute(
            "SELECT lang FROM documents WHERE id=?", (document_id,)
        ).fetchone()
        doc_lang = (doc_lang_row["lang"] or "en") if doc_lang_row else "en"
        con2.close()

        n_blocks = 0
        n_fields = 0
        n_translated = 0
        for cand_row in new_cands:
            block = extract_block_for_candidate(
                cand_row["id"], document_id, cand_row["page_start"])
            if block:
                con3 = db()
                con3.execute(
                    "UPDATE taxon_candidates SET block_text=? WHERE id=? "
                    "AND (block_text IS NULL OR block_text='')",
                    (block, cand_row["id"]))
                con3.commit(); con3.close()
                n_blocks += 1
                n_fields += auto_map_candidate_fields(cand_row["id"], block)

                # Přeložit pole LLM, pokud Document není anglicky a je překlad povolen.
                # Překlad následuje AŽ PO auto_map, aby měl co překládat.
                n_translated += auto_translate_candidate_fields(
                    cand_row["id"], doc_lang, settings)

        diag["Blocks extracted"] = n_blocks
        diag["Fields auto-mapped"] = n_fields
        if n_translated:
            diag["Fields translated"] = n_translated
    except Exception as exc:
        logging.warning(f"Auto-map second pass failed: {exc}")

    # Feature 5: Cross-reference synonym → automatické propojení records
    try:
        n_links = _link_synonyms_for_document(document_id)
        if n_links:
            diag["Synonym links"] = n_links
    except Exception as _e5:
        logging.warning(f"Synonym linking failed: {_e5}")

    # Post-detekce: auto-reject pre_systematic kandidátu jejichž jméno
    # se vyskytuje i v systematické zóně → jsou to citační zmínky, ne popisy.
    _duplicates = _pre_sys_seen & _sys_seen
    if _duplicates:
        try:
            _dup_con = db()
            _dup_rows = _dup_con.execute(
                "SELECT id, taxon_name, status FROM taxon_candidates "
                "WHERE document_id=? AND status NOT IN ('approved','rejected')",
                (document_id,)).fetchall()
            _dup_ids = []
            for _dr in _dup_rows:
                if (_dr["taxon_name"] or "").strip().lower() in _duplicates:
                    # Candidates s pre_systematic penalizací mají confidence výrazně nižší
                    # než systematičtí — odmítnout ty s confidence < threshold * 0.75
                    _dup_ids.append(_dr["id"])
            if _dup_ids:
                _dup_con.execute(
                    f"UPDATE taxon_candidates SET status='rejected' "
                    f"WHERE id IN ({','.join('?' * len(_dup_ids))})",
                    _dup_ids)
                _dup_con.commit()
                diag["Pre-systematic duplicates rejected"] = len(_dup_ids)
            _dup_con.close()
        except Exception as _e_dup:
            logging.warning(f"Pre-systematic dedup failed: {_e_dup}")

    return diag


def _get_pure_label_regex(schema: pd.DataFrame) -> re.Pattern:
    """
    Builds ONE regex for _is_pure_section_label() — replaces the equivalent
    of three condínkám dřívějšího per-řádkového porovnávání:
      (a) core == label            (řádek je JEN label, po odpagesění
                                     koncových ':', '.', mezer)
      (b) label bezprostředně následovaný ':'
      (c) label bezprostředně následovaný '.'
    Regex: label na začátku, pak buď jeden z ':'/'.' (b,c), or libovolná
    kombinace mezer/':'/​'.' až do konce řetězce (a).
    """
    global _PURE_LABEL_RE_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        if _PURE_LABEL_RE_CACHE is not None:
            return _PURE_LABEL_RE_CACHE
        labels = sorted(
            {str(row.get("label", "")).strip() for _, row in schema.iterrows()
             if len(str(row.get("label", "")).strip()) >= 4},
            key=len, reverse=True)
        if labels:
            alt = "|".join(re.escape(l) for l in labels)
            pat = re.compile(
                r"^(?:" + alt + r")(?:[:.]|[\s:.]*$)",
                re.IGNORECASE | re.UNICODE)
        else:
            pat = re.compile(r"(?!x)x")  # nikdy nic nenajde
        _PURE_LABEL_RE_CACHE = pat
        return pat


def _is_pure_section_label(line: str, schema: pd.DataFrame) -> bool:
    """
    True if this line is a section label (belongs INSIDE a block,
    not as a new headingpis taxonu). Používá načtené schema.

    VÝKON: tohle se volá pro KAŽDOU textovou jednotku Documentu (hlavní
    detekční smyčka) — dřív se tu pro každou jednotku iterovalo CELÉ
    schéma (stovky řádku) přes pandas iterrows(). U víceset-stránkového
    Documentu s tisíci jednotkami to dělalo statisíce Python-úrovňových
    porovnání. Teď je to jeden předkompilovaný regex, sestavený jen
    jednou a cachovaný (invalidace při změně schématu — viz reload_schema).
    """
    pat = _get_pure_label_regex(schema)
    return bool(pat.match(line.strip()))


# ══════════════════════════════════════════════════════════════════════════════
# SEGMENTACE BLOKU
# ══════════════════════════════════════════════════════════════════════════════

def extract_block_for_candidate(
    candidate_id: int,
    document_id: int,
    page_start: int,
) -> str:
    """
    Extracts the verbatim text block for an approved candidate.

    KEY CHANGE: the block is NESTAVÍ hrubým hledáním podřetězce (substring
    search) v konkatenovaném textu stránek — to je křehké u dvousloupcového
    layoutu, kde se konec předchozího záznamu muže v lineárním textu ocitnout
    za nadpisem následujícího (text z pravého sloupce pagesy N muže v
    pořadí čtení následovat až PO nadpisu na pagesě N+1 v levém sloupci).

    Místo toho se blok sestaví ze STEJNÉHO seznamu odstavcových jednotek
    (units), který použila detekce kandidátu (detect_candidates) — každý
    kandidát má saveený přesný `unit_index` (pořadí v tomto seznamu).
    Blok = units[muj_index : index_dalšího_kandidáta], spojené \n\n.
    Tím je hranice bloku 100% konzistentní s tím, co bylo skórováno jako
    nadpis, a nemuže dojít k duplicitnímu/chybnému nalezení přes find().

    HRANICE BLOKU: jako hranici bereme VŠECHNY kandidáty kromě 'rejected'
    — tedy i 'needs_review' a 'low_confidence'. Ty totiž typically JSOU
    skutečné nadpisy (jen s nejistým Score) a pokud by se ignorovaly,
    jejich text (a případně i další skutečný nadpis za nimi) by se
    omylem vstřebal do PŘEDCHOZÍHO bloku. Naopak 'rejected' hranicí
    NENÍ — to jsou detekce vyhodnocené jako falešně pozitivní (nejde
    o skutečný nadpis), takže jejich text správně patří do bloku,
    který je obklopuje.
    """
    con = db()
    cand = con.execute(
        "SELECT * FROM taxon_candidates WHERE id=?", (candidate_id,)
    ).fetchone()
    if not cand:
        con.close()
        return ""

    # Všichni kandidáti Documentu seřazení podle unit_index (= pořadí v textu).
    # 'rejected' se VYNECHÁVÁ (nejsou hranicí), all ostatní SE POČÍTÁ.
    all_cands = con.execute(
        """SELECT id, unit_index FROM taxon_candidates
           WHERE document_id=? AND status != 'rejected'
           ORDER BY unit_index""",
        (document_id,)
    ).fetchall()

    has_pages = con.execute(
        "SELECT COUNT(*) FROM pages WHERE document_id=?", (document_id,)
    ).fetchone()[0]
    con.close()

    if not has_pages:
        return ""

    # Fallback pro staré záznamy bez unit_index (zpětná kompatibilita
    # s kandidáty vytvořenými před zavedením tohoto pole, unit_index=-1)
    if cand["unit_index"] is None or cand["unit_index"] < 0:
        con = db()
        pages_rows = con.execute(
            "SELECT page_number, text, method, ocr_note, layout_note "
            "FROM pages WHERE document_id=? ORDER BY page_number",
            (document_id,)
        ).fetchall()
        con.close()
        return _extract_block_legacy_substring(cand, all_cands, pages_rows)

    # Text units z per-Documentové cache — dřív se tu volalo
    # build_text_units(pages_obj) ZNOVU při KAŽDÉM volání této funkce,
    # takže hromadné schválení 50 kandidátu re-segmentovalo Document 50×.
    units = get_cached_text_units(document_id)

    my_idx = cand["unit_index"]
    if my_idx >= len(units):
        return ""

    # Najít unit_index předchozího a dalšího kandidáta (ohraničí blok)
    cand_unit_indices = sorted(r["unit_index"] for r in all_cands if r["unit_index"] is not None and r["unit_index"] >= 0)
    next_idx = None
    prev_idx = None
    for ui in cand_unit_indices:
        if ui > my_idx and next_idx is None:
            next_idx = ui
        if ui < my_idx:
            prev_idx = ui

    # Pohltit bezprostředně PŘEDCHÁZEJÍCÍ "Class:"/"Order:"/"Family:" labely —
    # patří logicky k TOMUTO (genus-level) nadpisu, ne k předchozímu záznamu
    # (viz vzor: "Family: GRACILITHECIDAE Sysoev, 1972" patří k rodu Gracilitheca).
    start_idx = my_idx
    lower_bound = (prev_idx + 1) if prev_idx is not None else 0
    while (start_idx - 1 >= lower_bound and
           _RANK_PREFIX_LABEL_RE.match(units[start_idx - 1].text.strip())):
        start_idx -= 1

    end_slice = next_idx if next_idx is not None else len(units)
    block_units = units[start_idx:end_slice]

    # Vyfiltrovat popisky obrázku/tabulek a OCR šum z map/legend/diagramu
    # PŘÍMO z obsahu bloku — ne jen z kandidátní detekce. Vlastní nadpis
    # taxonu (na pozici heading_rel_idx, NE nutně na indexu 0 — viz pohlcení
    # předcházejících Class:/Order:/Family: labels výše) se NIKDY nefiltruje.
    heading_rel_idx = my_idx - start_idx
    kept_units = []
    for i, u in enumerate(block_units):
        if i == heading_rel_idx:
            kept_units.append(u)   # nadpis taxonu vždy zachovat
            continue
        stripped = u.text.strip()
        _unit_flags = set(f for f in u.zone_flags.split(",") if f)
        if "book_index" in _unit_flags or "document_index" in _unit_flags:
            continue   # rejstřík/obsah Documentu nepatří do taxonomického bloku
        if CAPTION_RE.match(stripped):
            continue   # popisek obrázku/tabulky/desky
        if _garbage_penalty(u.text) <= -0.35:
            continue   # OCR šum z map/legend/diagramu
        if BOILERPLATE_RE.search(stripped):
            continue   # nakladatelský balast / běžící záhlaví
        if len(stripped) < 5:
            continue   # zbytkové OCR smetí ("bt", "7)" apod.)
        kept_units.append(u)

    # Posledních 1-3 jednotky END_SLICE jsou "Class:"/"Order:"/"Family:" labely,
    # které předcházejí DALŠÍMU (genus-level) kandidátovi — patří logicky k NĚMU,
    # ne k aktuálnímu bloku (viz vzor: "Family: GRACILITHECIDAE Sysoev, 1972"
    # patří k záznamu rodu Gracilitheca, ne k předchozímu druhu).
    while (len(kept_units) > 1 and
           _RANK_PREFIX_LABEL_RE.match(kept_units[-1].text.strip())):
        kept_units.pop()

    result = "\n\n".join(u.text for u in kept_units).strip()

    # Save block_end_page = poslední stránka bloku, aby PDF viewer zobrazil
    # allchny stránky záznamu (vícestránkové záznamy).
    if kept_units:
        _blk_end_pg = max(u.page_number for u in kept_units)
        _blk_start_pg = int(cand["page_start"] or 1)
        if _blk_end_pg >= _blk_start_pg:
            try:
                _ep_con = db()
                _ep_con.execute(
                    "UPDATE taxon_candidates SET block_end_page=? WHERE id=?",
                    (_blk_end_pg, candidate_id))
                _ep_con.commit(); _ep_con.close()
            except Exception:
                pass

    # Tvrdá hranice: blok NIKDY nesmí přetéct do Acknowledgements/References.
    m = END_REGION_RE.search(result)
    if m and m.start() > 50:
        result = result[:m.start()].rstrip()

    # Odstraň Cambridge Core / DOI bannerové řádky z bloku — jsou to
    # artefakty stahování PDF z webu ("https://doi.org/… Downloaded from
    # https://www.cambridge.org/core IP address:…") které se přimíchají
    # do textu, když PyMuPDF čte PDF s URL na každé stránce.
    _URL_LINE_RE = re.compile(
        r"^https?://\S+.*$|"
        r"^.*cambridge\.org/core.*$|"
        r"^.*IP address:.*$|"
        r"^.*terms of use.*$",
        re.MULTILINE | re.IGNORECASE)
    result = _URL_LINE_RE.sub("", result)
    result = re.sub(r"\n{3,}", "\n\n", result).strip()

    # Normalizuj typografické varianty labels i v saveeném bloku
    # (aby byl DB text čistý pro manuální prohlížení i re-mapping)
    result = _normalize_label_variants(result)

    return result


def _extract_block_legacy_substring(cand, all_cands, pages_rows) -> str:
    """
    Fallback method using substring search — only for candidates without a stored
    unit_index (vytvořené starší verzí aplikace). Noví kandidáti vždy
    používají přesnou unit_index metodu výše.
    """
    cand_ids = [r["id"] for r in all_cands]
    try:
        idx = cand_ids.index(cand["id"])
    except ValueError:
        return ""

    start_page = cand["page_start"]
    end_page = None
    # Bez unit_index nelze spolehlivě najít "dalšího" kandidáta podle pořadí;
    # použijeme nejbližší vyšší page_start jako hrubý odhad.
    con = db()
    next_row = con.execute(
        """SELECT page_start FROM taxon_candidates
           WHERE document_id=? AND id!=? AND page_start>=?
           ORDER BY page_start LIMIT 1""",
        (cand["document_id"], cand["id"], start_page)
    ).fetchone()
    con.close()
    if next_row:
        end_page = next_row["page_start"]

    pages_filtered = [r for r in pages_rows if r["page_number"] >= start_page
                      and (end_page is None or r["page_number"] <= end_page)]
    if not pages_filtered:
        return ""

    heading = cand["heading_text"] or ""
    full_text = ""
    for pi, pg in enumerate(pages_filtered):
        pg_text = pg["text"] or ""
        if pi == 0:
            pos = pg_text.find(heading[:60]) if heading else -1
            if pos != -1:
                pg_text = pg_text[pos:]
        full_text += pg_text + "\n"

    result = full_text.strip()
    m = END_REGION_RE.search(result)
    if m and m.start() > 50:
        result = result[:m.start()].rstrip()
    return result


# ══════════════════════════════════════════════════════════════════════════════
# SECTION MAPPING  (TSV-driven)
# ══════════════════════════════════════════════════════════════════════════════

# Synonymický year-line: "1891 Hyolithes signatulus NOV.; ..."
# Používá se v map_sections_from_block() pro auto-detekci synonymiky v blocích,
# kde missing explicitní label "Synonymy:" (typically starší literatury).
# Podmínka: 4-místný rok + mezera + Velké písmeno + mezera + malé písmeno.
_SYNONYMY_YEAR_LINE_RE = re.compile(
    r"(?m)^(\d{4})\s+"
    r"[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽÄÖÜÀÂÆÇÈÊËÎÏÔÙÛÜ]"
    r"[a-zA-Záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü\-]*"
    r"\s+[a-záčďéěíňóřšťúuýžäöüàâæçèêëîïôùûü]",
    re.UNICODE
)


def _get_label_only_pattern(schema: pd.DataFrame) -> Optional[re.Pattern]:
    """
    Cached regex for detecting section labels WITHOUT requiring a newline before
    them (usedoužívá se k vkládání \\n před labely uprostřed odstavce). Invalidace
    přes reload_schema → _invalidate_schema_derived_regex_cache.
    """
    global _LABEL_ONLY_PAT_CACHE
    with _SCHEMA_DERIVED_RE_LOCK:
        if _LABEL_ONLY_PAT_CACHE is not None:
            # Sentinel: prázdný never-match znamená "spočítáno, ale žádné labely"
            return _LABEL_ONLY_PAT_CACHE if _LABEL_ONLY_PAT_CACHE.pattern != "(?!x)x" else None
    pat: Optional[re.Pattern] = None
    try:
        has_lr = "label_regex" in schema.columns
        has_pr = "priority" in schema.columns
        _tmp = []
        for r in schema.itertuples(index=False):
            label = str(getattr(r, "label", "") or "").strip()
            field = str(getattr(r, "target_field", "") or "").strip()
            _has_cjk = bool(re.search(r"[\u4e00-\u9fff\u3040-\u30ff]", label))
            _min_len = 2 if _has_cjk else 3
            if len(label) < _min_len or not field:
                continue
            pat_col = str(getattr(r, "label_regex", "") or "").strip() if has_lr else ""
            prio    = int(str(getattr(r, "priority", "0") or "0").strip() or "0") if has_pr else 0
            alt = pat_col if pat_col else re.escape(label)
            try:
                re.compile(alt, re.IGNORECASE)
                _tmp.append((prio, len(label), alt))
            except re.error:
                _tmp.append((prio, len(label), re.escape(label)))
        _tmp.sort(key=lambda x: (x[0], x[1]), reverse=True)
        label_parts = [x[2] for x in _tmp]
        if label_parts:
            # Robustní detekce labelu UPROSTŘED odstavce (typický OCR artefakt
            # kde missing zalomení: "…concave.Description: conch straight…").
            # Podmínky (aby se předešlo falešným pozitivum):
            #  - před labelem je věta-ukončující interpunkce [.!?;)] (0-3 mezer)
            #  - PO labelu MUSÍ následovat dvojtečka / em-dash / en-dash
            #    (silný signál nadpisu; brání matchování prózových slov)
            pat = re.compile(
                r"(?<!\n)(?<!\A)"
                r"(?P<delim>[.!?;)\u3002\uff01\uff1f\uff1b\uff0c]\s{0,3})"
                r"(?P<label>" + "|".join(label_parts) + r")"
                r"(?P<sep>s?e?s?\s*[:\u2014\u2013\uff1a])",
                re.IGNORECASE | re.UNICODE
            )
    except Exception:
        pat = None
    with _SCHEMA_DERIVED_RE_LOCK:
        _LABEL_ONLY_PAT_CACHE = pat if pat is not None else re.compile(r"(?!x)x")
    return pat


def _normalize_label_variants(text: str) -> str:
    """
    Normalises typographic variants of section labels in block text
    BEFORE field detection pomocí regex. Zachovává obsah sections, mění jen formát labels.

    Opravuje tyto varianty (platí pro allchna pole — Diagnosis, Description…):
      • „D i a g n o s i s:" — mezery mezi písmeny (justified OCR, starší tisk)
      • „D.i.a.g.n.o.s.i.s." — tečky mezi písmeny (small-caps simulace)
      • „D  i  a  g" — dvojité mezery (OCR artifact z tučných písmen)
      • „Descrip-\ntion:" — dělení slova na konci řádku
      • „ﬁgure" → „figure" — typografické ligatury
      • „De\u00adscription" — soft-hyphen (neviditelný v textu, rozbije matching)
    """
    # 1. OCR/PDF ligatury + ® artefakt (starý OCR nahrazoval ﬁ znakem ®)
    _LIG = [("ﬁ","fi"),("ﬂ","fl"),("ﬀ","ff"),("ﬃ","ffi"),("ﬄ","ffl"),
            ("ﬅ","st"),("ﬆ","st"),("Ꜳ","AA"),("ꜳ","aa"),("Ǳ","DZ"),("ǲ","Dz"),
            ("ǳ","dz"),("ĳ","ij"),("Ĳ","IJ"),("ŉ","nʼ")]
    # ® (U+00AE) jako OCR artefakt za ﬁ ligaturou: "dif®culties" → "difficulties"
    # Bezpečné: ® uvnitř slova je vždy OCR chyba; skutečný znak ® je vždy osamocený.
    text = re.sub(r"(?<=[A-Za-z])®(?=[A-Za-z])", "fi", text)
    # Opravit i ¬ (U+00AC, NOT sign) — v PDF se občas zaměnil za - or ﬁ
    text = re.sub(r"(?<=[A-Za-z])¬(?=[A-Za-z])", "-", text)
    for src, dst in _LIG:
        text = text.replace(src, dst)

    # 2. Soft-hyphen (U+00AD) — neviditelný rozbíječ slov
    text = text.replace("\u00ad", "")

    # 3. Hyphenation na konci řádku uvnitř slova:
    #    „Descrip-\ntion" → „Description"
    #    Jen lowercase continuation (ne věta za pomlčkou)
    text = re.sub(r"([A-Za-záčďéěíňóřšťúuýžäöüàâèêëîïôùûü])-\n([a-záčďéěíňóřšťúuýžäöüàâèêëîïôùûü])",
                  r"\1\2", text)

    # 4. Spaced single letters: „D i a g n o s i s" → „Diagnosis"
    #    Podmínka: ≥3 jednotková slova + 1 závěrečné (celkem ≥4 písmena)
    #    → kratší sekvence (n sp., U S A) jsou bezpečně přeskočeny
    text = re.sub(
        r"(?<![A-Za-z\u00c0-\u024f])([A-Za-z] ){3,}[A-Za-z](?![A-Za-z\u00c0-\u024f])",
        lambda m: m.group(0).replace(" ", ""), text)

    # 4b. Double-spaced: „D  i  a  g" → „Diag"
    text = re.sub(
        r"(?<![A-Za-z])([A-Za-z]  ){2,}[A-Za-z](?![A-Za-z])",
        lambda m: re.sub(r" {2}", "", m.group(0)), text)

    # 4c. ALL-CAPS word-chunks: „STRA TI GRA PHY" → „STRATIGRAPHY"
    #     Artefakt z PDF 2-column layoutu. Podmínka: ≥3 ALL-CAPS tokeny.
    text = re.sub(
        r"(?<![A-Za-z])([A-Z]{2,} ){2,}[A-Z]{2,}(?![a-z])",
        lambda m: m.group(0).replace(" ", ""), text)

    # 5. Dotted single letters: „D.i.a.g.n.o.s.i.s." → „Diagnosis"
    #    Podmínka: ≥4 X. skupiny (U.S.A. = 3 → bezpečně přeskočeno)
    text = re.sub(
        r"(?<![A-Za-z])([A-Za-z]\.){4,}[A-Za-z]\.?(?![A-Za-z])",
        lambda m: re.sub(r"\.", "", m.group(0)), text)

    return text


def map_sections_from_block(block_text: str, rank: str = "") -> Dict[str, str]:
    """
    Maps sections inside a verbatim block to canonical fields.
    Returns {field_name: verbatim_text}.
    Nozapisuje 'Not provided' – chybějící pole jsou prostě vynechána.

    OPRAVY (2026-07):
    A) Před mapováním se do bloku vloží newliny před sekční labely, které
       se nacházejí uprostřed odstavce (typický artefakt PDFu kde missing
       zalomení řádku po tečce — „Description.Text Discussion.Text").
       Bez tohoto kroku regex vyžadující (^|\\n) před labelem nenajde nic.

    B) Hodnoty fields se normalizují: soft-wrap single-\\n se sloučí na
       mezeru a text každého pole tvoří jeden odstavec (pokud nezačíná
       dalším labelem). Tím se vyřeší „Description" obsahující uprostřed
       „Discussion" — ten se teď správně detekuje jako hranice pole.
    """
    schema = load_schema()
    sec_re, label_to_field = build_section_regex(schema)

    # ── Pre-processing: vložit \\n před labely uprostřed odstavce ────────────
    # Bez tohoto kroku by label uvnitř věty (bez preceding newline) nebyl
    # rozpoznán regexem sec_re (který vyžaduje ^|\\n před labelem).
    # Regex se cachuje (viz _LABEL_ONLY_PAT_CACHE) — dřív se přestavoval
    # při KAŽDÉM volání (1300+ labels → drahá kompilace na každý blok).
    _label_only_pat = _get_label_only_pattern(schema)

    # ── KROK 0: Normalizace typografických variant labels ────────────────────
    # Musí proběhnout PŘED vším ostatním (boilerplate strip, label_only_pat).
    block_text = _normalize_label_variants(block_text)

    # ── Strip embedded publication interruptions ─────────────────────────────
    # Záhlaví stránek, čísla pages a jiný boilerplate vložený uprostřed bloku
    # rozbijí detekci labels — "Description:" za "210\nM. Valent / Annales..."
    # není nalezeno, protože regex vidí pouze předchozí newline.
    _block_boilerplate_res = [
        re.compile(r"(?m)^\s*\d{1,4}\s*$"),                         # holé číslo pagesy
        re.compile(r"(?m)^.{2,40}/\s*.{5,60}\(\d{4}\)\s*\d+.*$"),  # "Autor / Časopis (rok) str."
        re.compile(r"(?im)^(centre de conservation|département du rh|cahiers scientifiques"
                   r"|citer ce document|fichier pdf|author.s personal copy"
                   r"|downloaded from|copyright ©|all rights reserved"
                   r"|this article|accepted manuscript)\b.*$"),
    ]
    processed = block_text
    for _bpre in _block_boilerplate_res:
        try:
            processed = _bpre.sub("", processed)
        except Exception:
            pass
    processed = re.sub(r"\n{3,}", "\n\n", processed)

    if _label_only_pat:
        # Vložit \n před label — ale zachovat délky (neměnit obsah)
        processed = _label_only_pat.sub(
            lambda m: m.group("delim") + "\n" + m.group("label") + m.group("sep"),
            processed
        )

    # ── Najdi allchny labely s jejich pozicemi ────────────────────────────────
    # 4-tuple: (start_pos, field, text_start_after_label, label_text)
    positions: List[Tuple[int, str, int, str]] = []
    for m in sec_re.finditer(processed):
        label = m.group(1)
        field = label_to_field.get(label.lower(), "")
        if field:
            positions.append((m.start(1), field, m.end(), label))

    # ── Synonymy heuristika: detekce bloku začínajících rokem ───────────────
    # Pokud schema nenašlo žádný "Synonymy:" label, ale v bloku jsou řádky
    # formátu "1891 Hyolithes signatulus NOV.; ..." (year-lines), injektujeme
    # pozici jako SYNONYMY bez labelu — tak text se správně přiřadí.
    #
    # Podmínky pro injekci (ochrana před falešnými pozitivy):
    #   A) ≥2 year-lines → jednoznačný synonymický blok (i krátká synonymika)
    #   B) 1 year-line AND odpovídá přísnějšímu SYNONYMY_LINE_RE (má pl./fig./p.)
    has_synonymy_label = any(f == "SYNONYMY" for _, f, _, _ in positions)
    if not has_synonymy_label:
        yr_matches = list(_SYNONYMY_YEAR_LINE_RE.finditer(processed))
        strict_matches = [m for m in SYNONYMY_LINE_RE.finditer(processed)]
        ru_matches = list(SYNONYMY_LINE_RE_RU.finditer(processed))
        # P 3.3: přísnější heuristika — ≥5 year-lines bez labelu, or ≥2 na začátku bloku
        _first_pos = yr_matches[0].start(1) if yr_matches else 9999
        _at_start   = _first_pos < 80                      # první year-line blízko začátku bloku
        inject_synonymy = (
            len(yr_matches) >= 5                                    # EN/CS: ≥5 year-lines (méně FP)
            or (_at_start and len(yr_matches) >= 2)                 # na začátku bloku: ≥2 stačí
            or (len(yr_matches) >= 2 and not positions)             # celý blok = jen synonymy
            or len(ru_matches) >= 3                                  # RU: ≥3 citace
        )
        if inject_synonymy:
            first_yr_pos = yr_matches[0].start(1)
            # Injektujeme jen pokud je year-line před existujícími pozicemi
            # NEBO pokud positions je prázdné (celý blok je synonymika).
            if not positions or first_yr_pos < positions[0][0]:
                # text_start = first_yr_pos: rok je součástí obsahu synonymie
                positions.append((first_yr_pos, "SYNONYMY", first_yr_pos, "Synonymy"))

    if not positions:
        return {}

    positions.sort(key=lambda x: x[0])

    # ── Deduplikace + řešení překryvu ────────────────────────────────────────
    # 1) Stejná pozice + pole → zachovat jednu.
    # 2) Překrývající se labely (start2 spadá do labelu1) → zachovat DELŠÍ label
    #    (specifičtější: "Type species" > "Type", "Stratigraphic range and
    #    distribution" > "Stratigraphic range").
    seen_pos = set()
    dedup_positions = []
    for p in positions:
        key = (p[0], p[1])
        if key not in seen_pos:
            seen_pos.add(key)
            dedup_positions.append(p)
    # Overlap resolution: procházet seřazené, zahodit kratší label který
    # začíná uvnitř textového rozsahu předchozího labelu (do jeho text_startu).
    _resolved: List[Tuple[int, str, int, str]] = []
    for p in dedup_positions:
        _pos, _field, _ts, _label = p
        if _resolved:
            _lp, _lf, _lts, _ll = _resolved[-1]
            # Pokud tento label začíná ještě uvnitř labelu předchozího
            # (tzn. jsou to konkurenční varianty téhož místa)
            if _pos < _lts:
                # Ponechat ten s DELŠÍM labelem (specifičtější)
                if len(_label) > len(_ll):
                    _resolved[-1] = p
                # jinak předchozí (delší) zustává, tento zahodit
                continue
        _resolved.append(p)
    positions = _resolved

    result: Dict[str, str] = {}
    for i, (pos, field, text_start, label) in enumerate(positions):
        next_pos = positions[i+1][0] if i + 1 < len(positions) else len(processed)
        chunk = processed[text_start:next_pos].strip()
        if not chunk:
            continue

        # ── Normalizace odstavce: sloučit soft-wrap single-\\n → mezera ──────
        # Zachovat dvojitý \\n (skutečný odstavec), sloučit jednoduchý \\n
        # kde pokračuje text malým písmenem (soft wrap z OCR).
        chunk = re.sub(
            r"([a-záčďéěíňóřšťúuýžäöü,;])\n([a-záčďéěíňóřšťúuýžäöü])",
            r"\1 \2", chunk)
        chunk = re.sub(r"\n{3,}", "\n\n", chunk).strip()

        # ── Čištění artefaktu na začátku chunku ────────────────────────────
        # Odpagesit zbytkové oddělovače/interpunkci co zustaly po labelu
        # (e.g. ": " or "— " které separátor nespotřeboval).
        chunk = re.sub(r"^[\s:.\u2014\u2013)\-–—;,]+", "", chunk).strip()
        # Přeskočit prázdné or příliš krátké zbytky (< 2 znaky = artefakt)
        if len(chunk) < 2:
            continue

        # ── TYPE TAXON ořez: "Type species" hodnota je typically JEN binomen
        # + max. jedna věta (autor, rok, odkaz). Pokud chunk obsahuje delší
        # text (protože další label chyběl), ořízni na první větu končící
        # tečkou následovanou velkým písmenem or koncem. Zabraňuje tomu,
        # aby se celý Diagnosis vsákl do TYPE TAXON když missing "Diagnosis:".
        if field == "TYPE TAXON" and len(chunk) > 200:
            m_cut = re.search(r"\.\s+(?=[A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ])", chunk[:250])
            if m_cut:
                chunk = chunk[:m_cut.start() + 1].strip()

        # P 3.2: pro DIAGNOSIS rozliš "Emended diagnosis" vs "Original diagnosis"
        _label_low = (label or "").lower()
        _is_emended   = any(kw in _label_low for kw in ("emend", "revised", "amended"))
        _is_original  = any(kw in _label_low for kw in ("original", "original diagnosis"))
        _prefix = ""
        if field == "DIAGNOSIS" and (_is_emended or _is_original):
            _prefix = "[Emended diagnosis] " if _is_emended else "[Original diagnosis] "

        if field not in result:
            result[field] = _prefix + chunk if _prefix else chunk
        else:
            # Stejné pole z více labels (Holotype + Paratype → TYPE SPECIMENS):
            # zřetězit v pořadí výskytu, oddělit odstavcem.
            _suffix = _prefix + chunk if _prefix else chunk
            result[field] = result[field] + "\n\n" + _suffix

    # ── P 4.2: DIAGNOSIS ↔ DESCRIPTION disambiguace (konzervativní) ──────────
    # Reklasifikace se provádí POUZE když je jedno pole prázdné a druhé nese
    # jasné signály toho typu. Na rozdíl od dřívější verze NEMAŽE puvodní pole
    # destruktivně, když je jeho label explicitní — jen v jednoznačných
    # případech přesune obsah. Pokud si nejsme jistí, ponecháme jak je.
    _diag = result.get("DIAGNOSIS", "").strip()
    _desc = result.get("DESCRIPTION", "").strip()
    if _desc and not _diag:
        # DESCRIPTION → DIAGNOSIS jen když je KRÁTKÝ a diagnostický ve stylu
        # (čárkami oddělený výčet znaku, žádné dlouhé věty). Diagnózy hyolitu
        # bývají telegrafické: "Shell small, conical, apical angle 8°, …".
        _short = len(_desc) < 200
        _listy = _desc.count(",") >= 2
        _no_long_sentences = not re.search(r"[.!?]\s+[A-Z][a-z]+\s+\w+\s+\w+\s+\w+", _desc)
        if _short and _listy and _no_long_sentences:
            result["DIAGNOSIS"] = _desc
            del result["DESCRIPTION"]
    elif _diag and not _desc:
        # DIAGNOSIS → DESCRIPTION jen když je DLOUHÝ a narativní (4+ vět).
        # Diagnóza je stručná; delší narativní text je popis.
        _sentences = [s for s in re.split(r"[.!?]+", _diag) if len(s.strip()) > 10]
        if len(_diag) > 500 and len(_sentences) >= 5:
            result["DESCRIPTION"] = _diag
            del result["DIAGNOSIS"]

    # ── REST: text před prvním labelem, který nebyl přiřazen ─────────────────
    # Pokud blok začíná textem PŘED prvním nalezeným labelem, saveit ho jako REST.
    if positions:
        _first_text_pos = positions[0][0]
        _rest_text = processed[:_first_text_pos].strip()
        if _rest_text and len(_rest_text) > 20 and "REST" not in result:
            result["REST"] = _rest_text

    # ── Rank-aware TYPE TAXON ↔ TYPE SPECIMENS korekce ───────────────────────
    # PRAVIDLO taxonomické nomenklaktury:
    #   species / subspecies → mají HOLOTYP/PARATYP → TYPE SPECIMENS
    #   genus / family / … → mají typový druh/rod  → TYPE TAXON
    #
    # Pokud schema omylem přiřadilo špatné pole (label "type species" v bloku
    # druhu, or "holotype" v bloku rodu), opravíme to zde.
    if rank:
        _rank_norm = (rank or "").lower().split()[0]
        # Přeložíme přes alias (subspecies→species, subgenus→genus…)
        if _rank_norm in _REQUIRED_FIELDS_BY_RANK:
            _rk = _rank_norm
        else:
            _rk = _RANK_ALIASES.get(_rank_norm, "")

        _is_species_grade = _rk in ("species", "subspecies")
        _is_higher        = _rk in ("genus", "family", "order", "class", "phylum")

        if _is_species_grade:
            # Druh/poddruh nemá TYPE TAXON — přejmenovat na TYPE SPECIMENS,
            # pokud TYPE SPECIMENS ještě není filled.
            if "TYPE TAXON" in result and "TYPE SPECIMENS" not in result:
                _tt_val = result.pop("TYPE TAXON")
                # Bezpečnostní kontrola: TYPE TAXON u druhu typically vypadá jako
                # „Genus epithet Autor, 1900" or „typový druh: …" — krátký text.
                # Pokud je příliš krátký pro typ. exemplář, přeřadit jen jako NOTE.
                if len(_tt_val) < 15:
                    result.setdefault("REMARKS", _tt_val)
                else:
                    result["TYPE SPECIMENS"] = _tt_val
            # Holotype/paratype zmínky v REST → přesunout do TYPE SPECIMENS
            _rest_val = result.get("REST", "")
            if _rest_val and "TYPE SPECIMENS" not in result:
                if any(kw in _rest_val.lower() for kw in TYPE_SPECIMEN_KEYWORDS):
                    result["TYPE SPECIMENS"] = _rest_val
                    result.pop("REST", None)

        elif _is_higher:
            # Rod/čeleď/… nemá TYPE SPECIMENS jako heading-label — mohlo by
            # vzniknout jen chybou. Přejmenovat na TYPE TAXON pokud TYPE TAXON missing.
            # POZOR: holotype/paratype zmínky v textu jsou legitimní (cit. orig. popisu)
            # → přejmenováváme POUZE pokud se jedná o schema-label, ne inline zmínku.
            if "TYPE SPECIMENS" in result and "TYPE TAXON" not in result:
                _ts_val = result["TYPE SPECIMENS"]
                # Heuristika: pokud hodnota neobsahuje holotype/paratype keywords,
                # pravděpodobně jde o typ. druh/rod → přeřadit do TYPE TAXON
                if not any(kw in (_ts_val or "").lower() for kw in TYPE_SPECIMEN_KEYWORDS):
                    result["TYPE TAXON"] = result.pop("TYPE SPECIMENS")

    return result


def annotate_raw_block(block_text: str) -> str:
    """
    Returns RAW_TAXONOMIC_BLOCK with a canonical category tag ("[DIAGNOSIS]",
    "[DESCRIPTION]" apod.) přidaným NA ZAČÁTEK každého odstavce, kde byl
    rozpoznán sekční label ze schématu (libovolný jazyk) — text samotný
    zustává VERBATIM, štítek se pouze PŘIDÁ před něj, nic se nemaže ani
    nepřepisuje. Odstavce bez rozpoznaného labelu zustávají beze změny.

    Příklad: "Diagnóza: Skořápka je..." → "[DIAGNOSIS] Diagnóza: Skořápka je…"

    Použití: čitelnější náhled RAW bloku v Editoru a v exportech, zejména
    u vícejazyčných textu, kde label sám o sobě nenapovídá kategorii anglicky.
    """
    schema = load_schema()
    sec_re, label_to_field = build_section_regex(schema)

    # Najít pozice rozpoznaných labels, ale POUZE pokud label stojí na
    # začátku odstavce (po \n\n or na úplném začátku textu) — jinak by
    # se mohl objevit štítek uprostřed věty, kde label jen náhodně padne
    # na slovo shodné se sekčním labelem.
    paragraphs = re.split(r"(\n{2,})", block_text)   # zachovat oddělovače
    out_parts: List[str] = []
    for part in paragraphs:
        if part.strip() == "" or part.startswith("\n"):
            out_parts.append(part)
            continue
        # sec_re očekává "(?:^|\n)LABEL[:.\n]" — ^ je kotva na začátku
        # ŘETĚZCE, takže part (samostatný odstavec) lze testovat přímo.
        m = sec_re.match(part)
        if m:
            label = m.group(1)
            field = label_to_field.get(label.lower(), "")
            if field:
                out_parts.append(f"[{field}] {part}")
                continue
        out_parts.append(part)

    return "".join(out_parts)


# ══════════════════════════════════════════════════════════════════════════════
# LM STUDIO
# ══════════════════════════════════════════════════════════════════════════════

def lm_models(settings: Dict) -> List[str]:
    """
    Returns the list of models available in LM Studio.
    Raises RuntimeError with a problem descriptionu pokud server není dostupný.
    """
    base = settings.get("lmstudio_base_url", "http://localhost:1234/v1").rstrip("/")
    try:
        req = urllib.request.Request(base + "/models", method="GET")
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read().decode())
        return [x.get("id","") for x in data.get("data",[]) if x.get("id")]
    except ConnectionRefusedError:
        raise RuntimeError(
            f"LM Studio is not running at {base}. Start LM Studio and load a model.")
    except urllib.error.URLError as e:
        reason = str(getattr(e, "reason", e))
        raise RuntimeError(f"Cannot connect to LM Studio ({base}): {reason}")
    except Exception as exc:
        raise RuntimeError(f"Error loading models: {exc}")


def lm_chat(settings: Dict, system_prompt: str, user_prompt: str) -> str:
    """
    Calls LM Studio /chat/completions.
    Nover creates records – only returns suggestions.
    Vyvolá RuntimeError s akčním popisem při každém selhání.
    """
    if not settings.get("llm_enabled"):
        raise RuntimeError("LLM is not enabled — enable it in Settings → LM Studio.")
    model = settings.get("lmstudio_model", "").strip()
    if not model:
        raise RuntimeError("No model selected — choose a model in Settings → LM Studio.")
    base = settings.get("lmstudio_base_url","http://localhost:1234/v1").rstrip("/")
    payload = {
        "model": model,
        "messages": [
            {"role":"system","content":system_prompt},
            {"role":"user","content":user_prompt},
        ],
        "temperature": float(settings.get("llm_temperature", 0.0)),
    }
    req = urllib.request.Request(
        base + "/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type":"application/json"},
        method="POST",
    )
    timeout = int(settings.get("llm_timeout", 180))
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode())
    except ConnectionRefusedError:
        raise RuntimeError(
            f"LM Studio is not running or unreachable at {base}. "
            "Spusťte LM Studio a načtěte model.")
    except urllib.error.URLError as e:
        reason = str(getattr(e, "reason", e))
        if "refused" in reason.lower():
            raise RuntimeError(
                f"LM Studio unreachable at {base} — zkontrolujte, zda běží a "
                "v Settings → Local Inference Server je povolen.")
        raise RuntimeError(f"Notwork error communicating with LM Studio: {reason}")
    except urllib.error.HTTPError as e:
        body = ""
        try:
            body = e.read().decode("utf-8", errors="replace")[:200]
        except Exception:
            pass
        if e.code == 404:
            raise RuntimeError(
                f"LM Studio returned 404 — wrong base URL? ('{base}')")
        if e.code in (500, 503):
            raise RuntimeError(
                f"LM Studio returned {e.code} — model pravděpodobně není načten. "
                f"Detail: {body}")
        raise RuntimeError(f"HTTP {e.code} od LM Studio: {body}")
    except TimeoutError:
        raise RuntimeError(
            f"Timeout po {timeout}s — model příliš pomalý or nedostupný. "
            "Zkuste zvýšit Timeout v Settings → LM Studio.")
    try:
        return data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as e:
        raise RuntimeError(
            f"Unexpected response format from LM Studio: {str(data)[:200]}")


def lm_parse_json(raw: str) -> Optional[Dict]:
    """Parses JSON from LLM output, removing markdown backticks."""
    raw = re.sub(r"^```(?:json)?\s*", "", raw.strip(), flags=re.I)
    raw = re.sub(r"\s*```$", "", raw.strip())
    try:
        return json.loads(raw)
    except Exception:
        # Zkus najít JSON uvnitř textu
        m = re.search(r"\{.*\}", raw, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
    return None


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════

def get_candidate_fields(candidate_id: int) -> Dict[str, str]:
    con = db()
    rows = con.execute(
        "SELECT field_name, field_value FROM occurrence_fields WHERE candidate_id=?",
        (candidate_id,)
    ).fetchall()
    con.close()
    return {r["field_name"]: r["field_value"] for r in rows}


def get_candidate(cid: int) -> Optional[Dict]:
    con = db()
    r = con.execute("SELECT * FROM taxon_candidates WHERE id=?", (cid,)).fetchone()
    con.close()
    return dict(r) if r else None


# ══════════════════════════════════════════════════════════════════════════════
# ACTIVE BLOCK + FIELD EVENTS + BOUNDARY SUGGESTIONS
# ══════════════════════════════════════════════════════════════════════════════

_TYPE_KIND_ORDER = ["holotype", "lectotype", "neotype", "syntype", "paratype", "paralectotype"]
_TYPE_KIND_RE = re.compile(
    r"\b(holotype|lectotype|neotype|syntype|paratype|paralectotype|"
    r"голотип|лектотип|неотип|синтип|паратип|паралектотип|"
    r"正模|选模|新模|合模|副模)\b",
    re.IGNORECASE | re.UNICODE)
_TYPE_MATERIAL_SIGNAL_RE = re.compile(
    r"\b(holotype|lectotype|neotype|syntype|paratype|paralectotype|type\s+material|type\s+specimen|"
    r"голотип|лектотип|неотип|синтип|паратип|типовой\s+материал|正模|选模|新模|副模|模式材料|模式标本)\b",
    re.IGNORECASE | re.UNICODE)
_TYPE_TAXON_LINE_RE = re.compile(
    r"(?im)^\s*((?:type\s+species|type\s+genus|typový\s+druh|typový\s+rod|"
    r"типовой\s+вид|типовой\s+род|espèce\s+type|genre\s+type|Typusart|Typusgattung|模式种|模式属)"
    r"\s*[:.–—-]?\s*.+)$")
_INCLUDED_LINE_RE = re.compile(
    r"(?im)^\s*((?:included|included\s+species|included\s+genera|included\s+taxa|species\s+included|genera\s+included|"
    r"included\s+taxons|zahrnuté|zahrnuté\s+taxony|включенные|состав|包含|归入)"
    r"\s*[:.–—-]?\s*.+)$")
_FIELD_CONTINUATION_RE = re.compile(
    r"(?i)^(diagnosis|description|redescription|material|material\s+examined|type\s+material|holotype|lectotype|paratype|"
    r"type\s+locality|type\s+horizon|occurrence|distribution|remarks|discussion|etymology|synonymy|figures?)\b")


def get_active_block_text(cand: Dict[str, Any]) -> str:
    """Returns the active block: manual if active and non-empty; otherwise the parser block_text."""
    if not cand:
        return ""
    manual = (cand.get("manual_block_text") or "").strip()
    source = (cand.get("active_block_source") or "parser").strip().lower()
    if source == "manual" and manual:
        return manual
    return cand.get("block_text") or ""


def set_active_block_source(candidate_id: int, source: str) -> None:
    source = "manual" if source == "manual" else "parser"
    con = db()
    con.execute("UPDATE taxon_candidates SET active_block_source=?, block_version=COALESCE(block_version,1)+1 WHERE id=?",
                (source, candidate_id))
    con.commit(); con.close()
    _fts_update_candidate(candidate_id)


def save_manual_block(candidate_id: int, text_value: str, activate: bool = True) -> None:
    con = db()
    user = st.session_state.get("pn_user", "") if st is not None else ""
    source = "manual" if activate else "parser"
    con.execute(
        """UPDATE taxon_candidates
           SET manual_block_text=?, active_block_source=?, boundary_method='manual',
               boundary_reason='manual_block_override', manual_edited_at=?, manual_edited_by=?,
               block_version=COALESCE(block_version,1)+1
           WHERE id=?""",
        (text_value or "", source, datetime.now().isoformat(), user, candidate_id))
    con.commit(); con.close()
    _fts_update_candidate(candidate_id)


def ensure_candidate_block(candidate_id: int) -> str:
    cand = get_candidate(candidate_id)
    if not cand:
        return ""
    active = get_active_block_text(cand)
    if active:
        return active
    block = extract_block_for_candidate(candidate_id, cand["document_id"], cand["page_start"])
    if block:
        con = db()
        con.execute("""UPDATE taxon_candidates
                       SET block_text=?, boundary_method='parser', boundary_reason=COALESCE(NULLIF(boundary_reason,''),'extracted_on_demand'),
                           boundary_confidence=COALESCE(boundary_confidence,0.70)
                       WHERE id=?""", (block, candidate_id))
        con.commit(); con.close()
    return block or ""


def _field_confidence(method: str, field: str, label: str = "") -> float:
    base = {"schema_label": 0.92, "inline_type": 0.90, "rank_aware_derivation": 0.88, "content_fallback": 0.72}.get(method, 0.70)
    if field in {"TYPE MATERIAL", "TYPE TAXON", "TYPE SPECIMENS"}:
        base += 0.03
    if label:
        base += 0.02
    return min(base, 0.99)


def detect_field_events(block_text: str) -> List[Dict[str, Any]]:
    """Returns sorted field/taxon events in the active block for segmentation, debugging and GUI."""
    events: List[Dict[str, Any]] = []
    if not block_text:
        return events
    schema = load_schema()
    sec_re, label_to_field = build_section_regex(schema)
    for m in sec_re.finditer(block_text):
        label = m.group(1)
        field = label_to_field.get(label.lower(), "")
        if field:
            events.append({
                "pos": int(m.start(1)), "end": int(m.end()), "kind": "field_label",
                "field": field, "label": label, "method": "schema_label",
                "confidence": _field_confidence("schema_label", field, label),
            })
    for m in _TYPE_MATERIAL_SIGNAL_RE.finditer(block_text):
        line_start = block_text.rfind("\n", 0, m.start()) + 1
        existing_near = any(abs(ev["pos"] - line_start) < 5 and ev.get("field") == "TYPE MATERIAL" for ev in events)
        if not existing_near:
            events.append({"pos": int(line_start), "end": int(m.end()), "kind": "field_label",
                           "field": "TYPE MATERIAL", "label": m.group(1), "method": "inline_type", "confidence": 0.91})
    for m in _TYPE_TAXON_LINE_RE.finditer(block_text):
        events.append({"pos": int(m.start(1)), "end": int(m.end(1)), "kind": "field_label",
                       "field": "TYPE TAXON", "label": "type_taxon", "method": "inline_type", "confidence": 0.93})
    first_line = block_text.strip().splitlines()[0] if block_text.strip() else ""
    try:
        name, rank, pat = _match_name(first_line)
        if name:
            events.append({"pos": 0, "end": len(first_line), "kind": "taxon_heading",
                           "field": "", "label": first_line, "rank": rank, "method": pat, "confidence": 0.90})
    except Exception:
        pass
    events.sort(key=lambda e: (e["pos"], 0 if e["kind"] == "taxon_heading" else 1))
    dedup: List[Dict[str, Any]] = []
    for ev in events:
        dup_i = next((i for i, d in enumerate(dedup)
                      if abs(d["pos"]-ev["pos"]) <= 2 and d.get("field") == ev.get("field") and d["kind"] == ev["kind"]), None)
        if dup_i is None:
            dedup.append(ev)
        elif ev.get("confidence", 0) > dedup[dup_i].get("confidence", 0):
            dedup[dup_i] = ev
    return dedup


def segment_fields_by_events(block_text: str, events: Optional[List[Dict[str, Any]]] = None) -> Tuple[Dict[str, str], Dict[str, Any]]:
    """Segments fields by events: text from a label up to the next field/taxon event."""
    events = events if events is not None else detect_field_events(block_text)
    field_events = [e for e in events if e.get("kind") == "field_label" and e.get("field")]
    field_events.sort(key=lambda e: e["pos"])
    fields: Dict[str, str] = {}
    debug: Dict[str, Any] = {"events": events, "field_confidence": {}}
    for i, ev in enumerate(field_events):
        next_positions = [e["pos"] for e in field_events[i+1:] if e["pos"] > ev["pos"]]
        end_pos = min(next_positions) if next_positions else len(block_text)
        start_val = ev.get("end", ev["pos"])
        value = block_text[start_val:end_pos].strip(" \t\n:;.-–—")
        field = ev["field"]
        if not value:
            continue
        if field in fields:
            if value not in fields[field]:
                fields[field] = fields[field].rstrip() + "\n" + value
        else:
            fields[field] = value
        debug["field_confidence"][field] = {
            "confidence": ev.get("confidence", 0.75),
            "method": ev.get("method", "field_event"),
            "label": ev.get("label", ""),
            "pos": ev.get("pos"),
        }
    return fields, debug


def _extract_type_taxon_name_from_line(line: str) -> str:
    if not line:
        return ""
    val = re.sub(
        r"(?i)^\s*(type\s+species|type\s+genus|typový\s+druh|typový\s+rod|типовой\s+вид|типовой\s+род|"
        r"espèce\s+type|genre\s+type|Typusart|Typusgattung|模式种|模式属)\s*[:.–—-]?\s*", "", line).strip()
    return val or line.strip()


def derive_type_and_included_fields(candidate_id: int, block_text: str = "") -> Dict[str, str]:
    """Rank-aware derivace TYPE MATERIAL/TYPE SPECIMENS/TYPE TAXON/INCLUDED TAXONS."""
    cand = get_candidate(candidate_id)
    if not cand:
        return {}
    rank = (cand.get("rank_guess") or "").strip().lower()
    fields = get_candidate_fields(candidate_id)
    block = block_text or get_active_block_text(cand)
    out: Dict[str, str] = {}
    if rank in {"species", "subspecies"}:
        src = fields.get("TYPE MATERIAL") or fields.get("TYPE SPECIMENS") or ""
        if not src and block and _TYPE_MATERIAL_SIGNAL_RE.search(block):
            m = _TYPE_MATERIAL_SIGNAL_RE.search(block)
            start = max(0, block.rfind("\n", 0, m.start()))
            end_candidates = [x for x in [
                block.find("\nDiagnosis", m.end()), block.find("\nDescription", m.end()),
                block.find("\nRemarks", m.end()), block.find("\nOccurrence", m.end())] if x != -1]
            end = min(end_candidates) if end_candidates else min(len(block), m.end()+1200)
            src = block[start:end].strip()
        if src:
            out["TYPE MATERIAL"] = src
            lines = [ln.strip() for ln in re.split(r"[\n;]", src) if _TYPE_KIND_RE.search(ln)]
            prioritized = []
            for kind in ["holotype", "lectotype", "neotype", "syntype", "paratype", "paralectotype"]:
                for ln in lines:
                    if kind in ln.lower() and ln not in prioritized:
                        prioritized.append(ln)
            out["TYPE SPECIMENS"] = "; ".join(prioritized[:8]) if prioritized else src
            try:
                parsed = _parse_type_specimens(out["TYPE SPECIMENS"])
                if parsed.get("type_kinds"):
                    out["TYPE_SPECIMEN_KIND"] = parsed["type_kinds"]
                if parsed.get("institution_codes"):
                    out["INSTITUTION_CODE"] = parsed["institution_codes"]
                if parsed.get("catalog_numbers"):
                    out["CATALOG_NUMBER"] = parsed["catalog_numbers"]
            except Exception:
                pass
    elif rank in {"genus", "subgenus", "family", "order", "class", "phylum"}:
        line = fields.get("TYPE TAXON") or ""
        if not line and block:
            tm = _TYPE_TAXON_LINE_RE.search(block)
            if tm:
                line = tm.group(1).strip()
        if line:
            out["TYPE TAXON"] = line.strip()
            inc_name = _extract_type_taxon_name_from_line(line)
            existing_inc = fields.get("INCLUDED TAXONS", "")
            if inc_name:
                if existing_inc and inc_name.lower() not in existing_inc.lower():
                    out["INCLUDED TAXONS"] = existing_inc.rstrip() + "; " + inc_name
                elif not existing_inc:
                    out["INCLUDED TAXONS"] = inc_name
        elif block:
            im = _INCLUDED_LINE_RE.search(block)
            if im and not fields.get("INCLUDED TAXONS"):
                out["INCLUDED TAXONS"] = im.group(1).strip()
    return {k: v for k, v in out.items() if v and str(v).strip()}


def update_candidate_debug(candidate_id: int, **items) -> None:
    cand = get_candidate(candidate_id)
    if not cand:
        return
    try:
        dbg = json.loads(cand.get("debug_json") or "{}")
        if not isinstance(dbg, dict):
            dbg = {}
    except Exception:
        dbg = {}
    dbg.update(items)
    con = db()
    con.execute("UPDATE taxon_candidates SET debug_json=? WHERE id=?",
                (json.dumps(dbg, ensure_ascii=False), candidate_id))
    con.commit(); con.close()


def suggest_boundary_for_candidate(candidate_id: int, lookahead: int = 80) -> Dict[str, Any]:
    """Suggests end_unit_id and a reason based on text units after the candidate."""
    cand = get_candidate(candidate_id)
    if not cand:
        return {}
    try:
        units = get_cached_text_units(cand["document_id"])
    except Exception as exc:
        return {"candidate_id": candidate_id, "error": str(exc), "reason": "text_units_unavailable", "confidence": 0.0}
    if not units:
        return {"candidate_id": candidate_id, "reason": "no_text_units", "confidence": 0.0}
    start = int(cand.get("start_unit_id") or cand.get("unit_index") or 0)
    start = max(0, min(start, len(units)-1))
    best_end = int(cand.get("end_unit_id") or start)
    reason = "lookahead_limit"
    confidence = 0.55
    alternatives = []
    for i in range(start+1, min(len(units), start+lookahead)):
        u = units[i]
        t = (u.text or "").strip()
        if not t:
            continue
        zf = u.zone_flags or ""
        if "references" in zf:
            best_end, reason, confidence = i-1, "references_zone", 0.94
            break
        if "caption" in zf and i > start+2:
            alternatives.append({"end_unit_id": i-1, "reason": "before_caption", "confidence": 0.70})
        try:
            nm, rk, pat = _match_name(t)
        except Exception:
            nm, rk, pat = None, "", ""
        if nm:
            rk_l = (rk or "").lower()
            if i > start+1:
                best_end, reason, confidence = (
                    i-1, f"next_taxon_heading:{rk}",
                    0.90 if rk_l in {"species", "subspecies", "genus", "family", "order", "class", "phylum"} else 0.78)
                alternatives.append({"end_unit_id": best_end, "reason": reason, "confidence": confidence})
                break
        if _FIELD_CONTINUATION_RE.match(t) or _TYPE_MATERIAL_SIGNAL_RE.search(t) or _TYPE_TAXON_LINE_RE.search(t):
            best_end = i; reason = "continued_by_field_or_type_signal"; confidence = max(confidence, 0.82); continue
        if u.page_number != units[start].page_number and i <= start+4:
            best_end = i; reason = "continued_across_page"; confidence = max(confidence, 0.76)
    return {"candidate_id": candidate_id, "suggested_start_unit_id": start,
            "suggested_end_unit_id": max(start, best_end), "reason": reason,
            "confidence": round(float(confidence), 3), "alternatives": alternatives[:5]}


def update_field_confidence_debug(candidate_id: int, field_debug: Dict[str, Any], boundary_suggestion: Optional[Dict[str, Any]] = None) -> None:
    payload = {"field_events": field_debug.get("events", []), "field_confidence": field_debug.get("field_confidence", {})}
    if boundary_suggestion:
        payload["boundary_suggestion"] = boundary_suggestion
    update_candidate_debug(candidate_id, **payload)


def run_post_approval_pipeline(candidate_id: int) -> None:
    block = ensure_candidate_block(candidate_id)
    if not block:
        update_candidate_debug(candidate_id, post_approval="no_block"); return
    auto_map_candidate_fields(candidate_id, block)
    events = detect_field_events(block)
    fields_by_events, field_debug = segment_fields_by_events(block, events)
    if fields_by_events:
        save_fields(candidate_id, fields_by_events, method="field_event_segmenter")
    derived = derive_type_and_included_fields(candidate_id, block)
    if derived:
        save_fields(candidate_id, derived, method="rank_aware_derivation")
    fields = get_candidate_fields(candidate_id)
    cand = get_candidate(candidate_id) or {}
    comp, missing = _completeness_check(cand.get("rank_guess", ""), fields)
    boundary = None
    try:
        boundary = suggest_boundary_for_candidate(candidate_id)
    except Exception as exc:
        boundary = {"error": str(exc)}
    try:
        compute_and_save_term_matches_for_candidate(candidate_id)
    except Exception as exc:
        update_candidate_debug(candidate_id, term_matching_error=str(exc))
    update_field_confidence_debug(candidate_id, field_debug, boundary)
    update_candidate_debug(
        candidate_id, post_approval="ok",
        active_block_source=cand.get("active_block_source", "parser"),
        completeness_score=round(float(comp), 3),
        missing_required_fields=missing, boundary_checked=True)



def get_candidates_fields_bulk(candidate_ids: List[int]) -> Dict[int, Dict[str, str]]:
    """
    Loads fields for multiple candidates at once (one SQL query instead of N).
    Returns {candidate_id: {field_name: field_value}}.
    Použití v smyčkách kde se volá get_candidate_fields() pro každý prvek.
    """
    if not candidate_ids:
        return {}
    ph = ",".join("?" * len(candidate_ids))
    con = db()
    rows = con.execute(
        f"SELECT candidate_id, field_name, field_value FROM occurrence_fields "
        f"WHERE candidate_id IN ({ph})",
        candidate_ids,
    ).fetchall()
    con.close()
    result: Dict[int, Dict[str, str]] = {cid: {} for cid in candidate_ids}
    for r in rows:
        result[r["candidate_id"]][r["field_name"]] = r["field_value"]
    return result


def save_fields(candidate_id: int, fields: Dict[str, str], method: str = "manual") -> None:
    con = db()
    for fname, fval in fields.items():
        if not fval or fval == NOT_PROVIDED:
            continue
        existing = con.execute(
            "SELECT id FROM occurrence_fields WHERE candidate_id=? AND field_name=?",
            (candidate_id, fname)
        ).fetchone()
        if existing:
            con.execute(
                "UPDATE occurrence_fields SET field_value=?, method=? WHERE id=?",
                (fval, method, existing["id"])
            )
        else:
            con.execute(
                "INSERT INTO occurrence_fields (candidate_id,field_name,field_value,method) VALUES (?,?,?,?)",
                (candidate_id, fname, fval, method)
            )
    con.commit()
    con.close()
    # Aktualizovat FTS5 index
    _fts_update_candidate(candidate_id)
    # term_matches_updated_after_save_fields: starší Morpho/Strat matcher po ruční editaci fields
    try:
        compute_and_save_term_matches_for_candidate(candidate_id)
    except Exception as exc:
        logging.debug(f"Term matches update failed for {candidate_id}: {exc}")


def _get_or_create_manual_document(filename: str = "Manual records") -> int:
    """Returns the ID of the virtual document for manual records, creating it if necessary."""
    con = db()
    row = con.execute("SELECT id FROM documents WHERE filename=? ORDER BY id LIMIT 1", (filename,)).fetchone()
    if row:
        doc_id = int(row["id"])
        con.close()
        return doc_id
    now = datetime.now().isoformat()
    cur = con.execute(
        "INSERT INTO documents (filename,path,lang,page_count,char_count,notes,created_at,pub_title) "
        "VALUES (?,?,?,?,?,?,?,?)",
        (filename, "manual://records", "manual", 1, 0,
         "Virtual document for manually created taxonomic records", now, filename),
    )
    doc_id = int(cur.lastrowid)
    con.execute(
        "INSERT INTO pages (document_id,page_number,text,method,ocr_note,layout_note) VALUES (?,?,?,?,?,?)",
        (doc_id, 1, "Manual records", "manual", NOT_PROVIDED, NOT_PROVIDED),
    )
    con.commit(); con.close()
    return doc_id


def create_manual_taxon_record(
    taxon_name: str,
    rank: str,
    block_text: str = "",
    document_id: Optional[int] = None,
    status: str = "approved",
    page_start: int = 1,
    fields: Optional[Dict[str, str]] = None,
    source_note: str = "manual_editor",
) -> int:
    """Fully creates a new taxonomic record manually from the Editor."""
    taxon_name = (taxon_name or "").strip()
    if not taxon_name:
        raise ValueError("Taxon name is required.")
    rank = (rank or "").strip() or _rank_from_name_shape(taxon_name, "Genus")
    status = status if status in STATUS_OPTIONS else "approved"
    block_text = (block_text or "").strip()
    document_id = int(document_id or _get_or_create_manual_document())
    now = datetime.now().isoformat()
    clean_fields = {
        str(k).strip(): str(v).strip()
        for k, v in (fields or {}).items()
        if str(k).strip() and v and str(v).strip() and str(v).strip() != NOT_PROVIDED
    }
    if not block_text:
        block_parts = [taxon_name]
        for k, v in clean_fields.items():
            block_parts.append(f"{k}: {v}")
        block_text = "\n".join(block_parts).strip() or taxon_name

    con = db()
    try:
        cur = con.execute(
            "INSERT INTO taxon_candidates "
            "(document_id,taxon_name,rank_guess,confidence,status,heading_text,context_before,context_after," 
            "page_start,block_text,created_at,block_end_page,unit_index,debug_json," 
            "manual_block_text,active_block_source,boundary_method,boundary_reason,boundary_confidence," 
            "block_version,manual_edited_at,manual_edited_by) "
            "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (document_id, taxon_name, rank, 1.0, status, taxon_name, "", "",
             int(page_start or 1), block_text, now, int(page_start or 1), -1,
             json.dumps({"created_by": source_note, "manual_record": True}, ensure_ascii=False),
             block_text, "manual", "manual", "created_in_editor", 1.0,
             1, now, st.session_state.get("pn_user", "") if st is not None else ""),
        )
        cid = int(cur.lastrowid)
        con.commit()
    except Exception:
        con.rollback(); con.close(); raise
    con.close()

    if clean_fields:
        save_fields(cid, clean_fields, method=source_note)
    try:
        auto_map_candidate_fields(cid, block_text)
    except Exception as exc:
        update_candidate_debug(cid, manual_automap_error=str(exc))
    try:
        derived = derive_type_and_included_fields(cid, block_text)
        if derived:
            save_fields(cid, derived, method="manual_rank_aware_derivation")
    except Exception as exc:
        update_candidate_debug(cid, manual_derived_error=str(exc))
    try:
        compute_and_save_term_matches_for_candidate(cid)
    except Exception as exc:
        update_candidate_debug(cid, manual_term_matching_error=str(exc))
    try:
        _fts_update_candidate(cid)
    except Exception:
        pass
    try:
        all_fields = get_candidate_fields(cid)
        comp, missing = _completeness_check(rank, all_fields)
        update_candidate_debug(cid, manual_record="created", completeness_score=round(float(comp), 3), missing_required_fields=missing)
    except Exception:
        pass
    return cid



# Regex pro inline odkazu na obrázky/tabule v taxonomickém bloku.
# Zachytí: Text-fig. 3, Pl. 1 figs. 1-8, (Figs. 3A-I, 4B), 图版I, Рис. 5.
# Vyloučí popisky obrázku (Caption pattern: na začátku řádku + ". Velké").
# ── Agregace odkazu na obrázky / tabule ───────────────────────────────
# Zachycuje: Text-fig. 3, Pl. 1, figs. 1–8, (Fig. 7E, J), 图版I
# Popisky (Caption): „Fig. 3. Description.“ — tvrzé '. VELKÉ' za číslem → vyloučit.
_FIG_REF_RE = re.compile(
    r"(?:Text-figs?\.?|Figs?\.?|Pl\.?s?|Plate\s*s?|图版|Рис\.?|табл\.?)"
    r"\s*"
    r"(?:[IVX]{1,6}|[0-9][0-9A-Za-z]?)"
    r"(?:[\s,.\-–]*(?:figs?\.?\s*)?(?:[IVX]{1,6}|[0-9][0-9A-Za-z]?))*",
    re.IGNORECASE | re.UNICODE
)
# Caption: prefix + číslo + TECKA + MEZERA + VELKÉ PÍSMENO
_FIG_CAPTION_LINE_RE = re.compile(
    r"(?:^|\n)\s*(?:Text-fig|Fig|Plate|Pl)[s]?\.?\s*"
    r"[0-9IVX][0-9A-Za-z,\-]*"
    r"\.\s+[A-Z]",
    re.MULTILINE | re.UNICODE
)


def _collect_figure_refs(block_text: str) -> str:
    """
    Aggregates inline image/plate references from a taxonomic block
    into a single string (odděleno "; "). Popisky obrázku (Caption) se vynechají.

    Zahrnuje:   Pl. 1, figs. 1-8  |  (Fig. 7E, J)  |  Text-fig. 3  |  图版I
    Vyloučí:   'Fig. 3. Detailed view...' (Caption: tečka + VELKÉ za číslem)
    """
    if not block_text:
        return ""

    # Zjistit rozsahy řádku, které jsou popisky (Caption).
    # POZOR: m.start() muže ukazovat na \n (v re.MULTILINE módu ^ vzáženě
    # matchuje \n jako součást \s*), proto hledáme skutečný začátek
    # klíčovho slova přeskočením úvodních whitespace znaku.
    _NL = chr(10)  # newline bez přímého použití '\n' (ochrana heredocu)
    caption_ranges: List[tuple] = []
    for m in _FIG_CAPTION_LINE_RE.finditer(block_text):
        # Najdi skutečnou pozici klíčového slova (přeskoč \n / mezery)
        kw_pos = m.start()
        while kw_pos < m.end() and block_text[kw_pos] in (_NL, chr(13), ' ', chr(9)):
            kw_pos += 1
        line_start = block_text.rfind(_NL, 0, kw_pos) + 1
        line_end   = block_text.find(_NL, kw_pos)
        if line_end == -1:
            line_end = len(block_text)
        caption_ranges.append((line_start, line_end))

    def _is_caption(pos: int) -> bool:
        return any(s <= pos <= e for s, e in caption_ranges)

    refs: List[str] = []
    seen: set = set()
    for m in _FIG_REF_RE.finditer(block_text):
        if _is_caption(m.start()):
            continue
        ref = m.group(0).strip().rstrip(".,;")
        key = re.sub(r"\s+", "", ref.lower())
        if key and key not in seen and len(key) >= 3:
            seen.add(key)
            refs.append(ref)

    return "; ".join(refs)

# ══════════════════════════════════════════════════════════════════════════════
# Content-based fallback extractory — vyplní pole i BEZ explicitního labelu
# ══════════════════════════════════════════════════════════════════════════════

# Type-specimen řádky bez explicitního "Type specimens:" labelu
_HOLOTYPE_INLINE_RE = re.compile(
    r"(?:^|\n)\s*(?:Holo|Lecto|Noo|Para|Syn)typ(?:e|us)?[sy]?\b[^\n]{0,300}",
    re.IGNORECASE | re.UNICODE)

# "Type horizon and locality" a podobné kombinované fráze
_TYPE_HORIZON_RE = re.compile(
    r"(?:^|\n)\s*Type\s+(?:horizon|locality|stratum|level)[^\n]{0,300}",
    re.IGNORECASE | re.UNICODE)

# Etymologie bez labelu: "The specific name refers to…", "named after…"
_ETYMOLOGY_INLINE_RE = re.compile(
    r"(?:^|\n)[^\n]*?(?:specific\s+(?:name|epithet)|named\s+(?:after|for|in\s+honou?r)"
    r"|derivation\s+of\s+(?:the\s+)?name|refers\s+to\s+the|means\s+\w+\s+in\s+(?:Latin|Greek))"
    r"[^\n]{0,200}",
    re.IGNORECASE | re.UNICODE)

# Occurrence bez labelu: "known from…", "found in…", "occurs in…"
_OCCURRENCE_INLINE_RE = re.compile(
    r"(?:^|\n)[^\n]*?(?:has\s+been\s+found\s+(?:only\s+)?in|is\s+known\s+from"
    r"|occurs?\s+in\s+the|recorded\s+from|distributed\s+in|representatives?\s+of)"
    r"[^\n]{0,300}",
    re.IGNORECASE | re.UNICODE)

# Stratigraphy bez labelu: kombinace jednotky + zóny/formace/stupně
_STRAT_INLINE_RE = re.compile(
    r"(?:^|\n)[^\n]*?"
    r"(?:Formation|Member|Biozone|Zone|Stage|Series|Cambrian|Ordovician|Silurian"
    r"|Devonian|Carboniferous|Permian|Triassic|Jurassic|Cretaceous|souvrstv[íi]"
    r"|\u0441\u0432\u0438\u0442\u0430|\u044f\u0440\u0443\u0441)"
    r"[^\n]{0,250}",
    re.IGNORECASE | re.UNICODE)


def _content_fallback_extract(block_text: str, already: Dict[str, str]) -> Dict[str, str]:
    """
    Extrahuje pole z obsahu bloku i BEZ explicitního labelu, pomocí
    obsahových vzoru. Vyplňuje jen pole, která ještě nejsou v `already`.
    Konzervativní — přidá pole jen když má vzor jasnou shodu.
    """
    out: Dict[str, str] = {}

    def _has(f: str) -> bool:
        v = already.get(f, "") or out.get(f, "")
        return bool(v and v.strip() and v != NOT_PROVIDED)

    # TYPE SPECIMENS z inline Holotype/Paratype řádku
    if not _has("TYPE SPECIMENS"):
        _tsm = _HOLOTYPE_INLINE_RE.findall(block_text)
        if _tsm:
            _joined = " ".join(m.strip() for m in _tsm[:6])
            if len(_joined) > 10:
                out["TYPE SPECIMENS"] = _joined.strip()

    # LOCALITY / STRATIGRAPHY z "Type horizon and locality"
    if not _has("LOCALITY") or not _has("STRATIGRAPHY"):
        _thm = _TYPE_HORIZON_RE.search(block_text)
        if _thm:
            _val = _thm.group(0).strip()
            if len(_val) > 15:
                if not _has("STRATIGRAPHY"):
                    out["STRATIGRAPHY"] = _val
                if not _has("LOCALITY"):
                    out["LOCALITY"] = _val

    # ETYMOLOGY z inline formulací
    if not _has("ETYMOLOGY"):
        _em = _ETYMOLOGY_INLINE_RE.search(block_text)
        if _em:
            _val = _em.group(0).strip()
            if 15 < len(_val) < 300:
                out["ETYMOLOGY"] = _val

    # OCCURRENCE z inline "known from / found in"
    if not _has("OCCURRENCE"):
        _om = _OCCURRENCE_INLINE_RE.search(block_text)
        if _om:
            _val = _om.group(0).strip()
            if len(_val) > 20:
                out["OCCURRENCE"] = _val

    # STRATIGRAPHY z inline geologických jednotek (jen pokud stále missing)
    if not _has("STRATIGRAPHY") and "STRATIGRAPHY" not in out:
        _sm = _STRAT_INLINE_RE.search(block_text)
        if _sm:
            _val = _sm.group(0).strip()
            if len(_val) > 15:
                out["STRATIGRAPHY"] = _val

    return out


# ── CJK / Cyrillic detection + LLM-assisted extraction ───────────────────────

def _is_cjk_dominant(text: str, threshold: float = 0.12) -> bool:
    """Returns True if the text contains ≥ threshold proportion of CJK or Cyrillic characters.

    Práh 12 % je záměrně nízký — smíšené čínsko-latinské vědecké texty mají
    přibližně 20–60 % čínských znaku (latinská jména taxa ten podíl snižují).
    """
    if not text:
        return False
    cjk = sum(
        1 for c in text
        if (0x4E00 <= ord(c) <= 0x9FFF)   # CJK Unified Ideographs
        or (0x3400 <= ord(c) <= 0x4DBF)   # CJK Extension A
        or (0xF900 <= ord(c) <= 0xFAFF)   # CJK Compatibility Ideographs
        or (0x0400 <= ord(c) <= 0x04FF)   # Cyrillic
    )
    return (cjk / max(len(text), 1)) >= threshold


def _detect_script(text: str) -> str:
    """Identifies the dominant script of the block: 'zh', 'ru', 'latin', or 'mixed'."""
    if not text:
        return "latin"
    cjk_count = sum(
        1 for c in text
        if (0x4E00 <= ord(c) <= 0x9FFF)
        or (0x3400 <= ord(c) <= 0x4DBF)
    )
    cyr_count = sum(1 for c in text if 0x0400 <= ord(c) <= 0x04FF)
    total = max(len(text), 1)
    if cjk_count / total >= 0.10:
        return "zh"
    if cyr_count / total >= 0.20:
        return "ru"
    if (cjk_count + cyr_count) / total >= 0.08:
        return "mixed"
    return "latin"


def llm_extract_fields_cjk_ru(block_text: str, settings: Dict) -> Dict[str, str]:
    """LLM-assisted field extraction for blocks dominated by CJK or Cyrillic script.

    Strategy (NOT treložit-nejprve):
      1. Pošle raw blok (čínsky/rusky) přímo LLM s instrukcí
         LLM_CJK_RU_EXTRACTION_PROMPT.
      2. LLM rozumí čínštině/ruštině nativně, identifikuje sekce semanticky
         (bez závislosti na mezerách or regex vzorech), a vrací JSON
         s anglickými překlady hodnot fields.
      3. Latinská jména taxa, autoři a roky jsou zachována verbatim —
         instrukce promptu to explicitně zakazuje.

    Výhody oproti přístupu "přeložit-nejprve":
      • Latinská jména taxa nejsou poškozena překladem.
      • Jeden LLM call místo dvou (překlad + extrakce).
      • LLM chápe Context v puvodním jazyce lépe než přeložený text.
      • Vhodné pro smíšené texty (čínština + latinská nomenklatura).

    Vrací dict {FIELD_NAME: value} or {} při chybě / LLM vypnuto.
    """
    if not settings.get("llm_enabled"):
        return {}
    if not block_text or not block_text.strip():
        return {}
    try:
        raw = lm_chat(
            settings,
            LLM_CJK_RU_EXTRACTION_PROMPT,
            block_text[:6000],  # limit na 6000 znaku — blok taxonu bývá kratší
        )
        parsed = lm_parse_json(raw)
        if not parsed:
            return {}
        fields_raw = parsed.get("fields", {})
        if not isinstance(fields_raw, dict):
            return {}
        # Normalizace: ořez whitespace, odmítnutí prázdných hodnot
        result: Dict[str, str] = {}
        for k, v in fields_raw.items():
            k2 = str(k).strip().upper()
            v2 = str(v).strip() if v else ""
            if k2 and v2 and v2.lower() not in ("not provided", "n/a", "none", ""):
                result[k2] = v2
        return result
    except Exception as exc:
        logging.warning("llm_extract_fields_cjk_ru: chyba — %s", exc)
        return {}


# Minimální count vyplněných fields (mimo AUTHOR/TAXON), pod nímž se spustí
# CJK/RU LLM pass jako gap-filler.
_CJK_RU_LLM_FIELD_THRESHOLD = 3


def auto_map_candidate_fields(candidate_id: int, block_text: str) -> int:
    """
    Automatically maps block text sections to standardised fields (AUTHOR,
    DIAGNOSIS, DESCRIPTION atd.) pomocí schema regex.

    Strategie: přepisovat pouze PRÁZDNÁ pole — manuálně editovaná or
    dřívějším pruchodem vyplněná pole zustávají beze změny. Vrací count
    nově vyplněných fields.

    Používá `method='auto'` tak, aby bylo v DB vidět, která pole přišla
    z automatického mapování (na rozdíl od 'manual' or 'llm'), a uživatel
    mohl v Editoru snadno identifikovat co zkontrolovat.
    """
    if not block_text or not block_text.strip():
        return 0
    # Načteme rank kandidáta pro rank-aware korekci TYPE TAXON↔TYPE SPECIMENS
    _rank_for_map = ""
    try:
        _con_r = db()
        _r_row = _con_r.execute(
            "SELECT rank_guess FROM taxon_candidates WHERE id=?", (candidate_id,)
        ).fetchone()
        _con_r.close()
        if _r_row:
            _rank_for_map = (_r_row["rank_guess"] or "").strip()
    except Exception:
        pass
    mapped = map_sections_from_block(block_text, rank=_rank_for_map)

    existing = get_candidate_fields(candidate_id)

    # ── Content-based fallback: vyplní pole bez explicitního labelu ──────────
    # Sloučit existující + již namapovaná pole pro kontrolu "co ještě missing"
    _combined = dict(existing)
    _combined.update(mapped)
    _fallback = _content_fallback_extract(block_text, _combined)
    for _f, _v in _fallback.items():
        if _f not in mapped:
            mapped[_f] = _v

    # FIGURES z inline odkazu (pokud schema nezachytil label "Figures:")
    if "FIGURES" not in mapped:
        figs = _collect_figure_refs(block_text)
        if figs:
            mapped["FIGURES"] = figs

    # SIZE_PARSED z SIZE pole (strukturovaný rozměr)
    if "SIZE" in mapped and "SIZE_PARSED" not in mapped:
        try:
            _sz = _parse_size_field(mapped["SIZE"])
            _sz_str = _size_dict_to_str(_sz)
            if _sz_str:
                mapped["SIZE_PARSED"] = _sz_str
        except Exception:
            pass

    # Strukturovaný typový materiál z TYPE SPECIMENS
    if "TYPE SPECIMENS" in mapped:
        try:
            _ts = _parse_type_specimens(mapped["TYPE SPECIMENS"])
            if _ts.get("type_kinds") and "TYPE_SPECIMEN_KIND" not in mapped:
                mapped["TYPE_SPECIMEN_KIND"] = _ts["type_kinds"]
            if _ts.get("institution_codes") and "INSTITUTION_CODE" not in mapped:
                mapped["INSTITUTION_CODE"] = _ts["institution_codes"]
            if _ts.get("catalog_numbers") and "CATALOG_NUMBER" not in mapped:
                mapped["CATALOG_NUMBER"] = _ts["catalog_numbers"]
        except Exception:
            pass

    # ── CJK / Cyrillic LLM gap-fill ──────────────────────────────────────────
    # Pokud je blok dominován čínštinou or ruštinou A regex extrakce přinesla
    # málo fields (< _CJK_RU_LLM_FIELD_THRESHOLD), spustíme specializovaný
    # LLM pass. Tento přístup je výrazně lepší než "přeložit-nejprve", protože:
    #   • Latinská jména taxa jsou zachována verbatim (LLM dostane instrukci).
    #   • Jeden LLM call místo dvou (překlad + extrakce).
    #   • LLM chápe sémantiku čínského/ruského textu nativně.
    # LLM pass se spustí pouze pokud je LLM povoleno v nastavení.
    _content_fields = {
        k for k in mapped
        if k not in {"AUTHOR", "TAXON", "TAXON_NAME_VERBATIM",
                     "TAXON_RANK_AS_WRITTEN", "FIGURES", "RECORD_ID"}
    }
    _llm_sourced_keys: set = set()
    if (
        _is_cjk_dominant(block_text)
        and len(_content_fields) < _CJK_RU_LLM_FIELD_THRESHOLD
    ):
        try:
            _settings = get_settings()
        except Exception:
            _settings = {}
        if _settings.get("llm_enabled"):
            _llm_fields = llm_extract_fields_cjk_ru(block_text, _settings)
            for _lf, _lv in _llm_fields.items():
                if _lf not in mapped and _lv:
                    mapped[_lf] = _lv
                    _llm_sourced_keys.add(_lf)
            if _llm_fields:
                logging.info(
                    "auto_map_candidate_fields[%d]: CJK/RU LLM pass added %d fields: %s",
                    candidate_id, len(_llm_sourced_keys), list(_llm_sourced_keys),
                )

    if not mapped:
        return 0
    to_save_auto: Dict[str, str] = {}
    to_save_llm:  Dict[str, str] = {}
    for field, value in mapped.items():
        if not value or not value.strip():
            continue
        if existing.get(field):   # jen prázdná pole
            continue
        if field in _llm_sourced_keys:
            to_save_llm[field] = value
        else:
            to_save_auto[field] = value
    if to_save_auto:
        save_fields(candidate_id, to_save_auto, method="auto")
    if to_save_llm:
        save_fields(candidate_id, to_save_llm, method="llm")
    return len(to_save_auto) + len(to_save_llm)


# Pole, která se nepřekládají — jde o vlastní jména, citace, čísla
_TRANSLATION_SKIP_FIELDS = {
    "AUTHOR", "RECORD_ID", "SOURCE_DOCUMENT", "FIGURES", "REFERENCE",
    "TAXON", "TAXON_NAME_VERBATIM", "TAXON_RANK_AS_WRITTEN",
    "NOMENCLATURAL ACTS",  # obsahuje citace zákonu a kódu – zachovat verbatim
}

# Languages které se nepřekládají (Documenty v angličtině or bez jazyka)
_ENGLISH_LANG_CODES = {"en", "eng", "english", "en-gb", "en-us", ""}

# Mapování anglických nadpisu sections → cílové DB pole.
# Aplikuje se automaticky PO překladu (non-EN Documenty).
# Klíče jsou lowercase verze toho, pod čím LLM sekci saveí
# (= puvodní nadpis zdrojového textu přeložený do AJ).
_POST_TRANSLATION_FIELD_MAP: Dict[str, str] = {
    "characteristics":               "DESCRIPTION",
    "discussion and comparison":     "REMARKS",    # 讨论与比较 / 比较与讨论
    "comparison and discussion":     "REMARKS",
    "age and distribution":          "OCCURRENCE", # 时代和分布
    "age and occurrence":            "OCCURRENCE",
    "distribution and age":          "OCCURRENCE",
    "geological and geographical distribution": "OCCURRENCE",
    "discussion":                    "REMARKS",
    "comparison":                    "REMARKS",
    "type species":                  "TYPE TAXON",
    "type genus":                    "TYPE TAXON",
}


def remap_fields_post_translation(candidate_id: int) -> int:
    """
    After translation, re-checks all record fields and remaps those
    whose name matchesvídá _POST_TRANSLATION_FIELD_MAP.

    Příklad: sekce "Characteristics" → saveena pod klíčem "CHARACTERISTICS"
    → tato funkce ji přesune do "DESCRIPTION".

    Pokud cílové pole již exists, hodnoty se sloučí (oddělovač "\\n\\n").
    Vrací count přemapovaných fields.
    """
    rows = get_candidate_fields(candidate_id)   # {FIELD_NAME: value}
    to_save:   Dict[str, str] = {}
    to_delete: List[str] = []

    for field, value in rows.items():
        target = _POST_TRANSLATION_FIELD_MAP.get(field.strip().lower())
        if not target or target == field:
            continue
        # Sloučit s případnou existující hodnotou v cílovém poli
        existing = rows.get(target, "")
        if existing:
            merged = existing.rstrip() + "\n\n" + value.strip()
        else:
            merged = value
        to_save[target] = merged
        to_delete.append(field)

    if to_save:
        save_fields(candidate_id, to_save, method="remap_post_translation")

    if to_delete:
        con = db()
        for f in to_delete:
            con.execute(
                "DELETE FROM occurrence_fields WHERE candidate_id=? AND field_name=?",
                (candidate_id, f),
            )
        con.commit()
        con.close()

    return len(to_delete)


def auto_translate_candidate_fields(
    candidate_id: int,
    doc_lang: str,
    settings: Dict[str, Any],
    **kwargs,
) -> int:
    """
    Translates the filled fields of a record from `doc_lang` to English using an LLM
    and updates je formátem:

        English translation (puvodní originální text)

    Parametry kwargs:
      force=True      — přeloží i bez llm_auto_translate; přeskočí jazykovou kontrolu
                        pro unknown jazyk (lang=""); šíří výjimky LM Studio do UI
      force_lang=str  — přepíše doc_lang pro prompt (e.g. "cs" pokud DB nemá lang)

    Pole uvedená v `_TRANSLATION_SKIP_FIELDS` se přeskakují.
    Vrací count skutečně přeložených fields; vyvolá RuntimeError při chybě LM Studio
    pokud force=True (jinak vrátí 0 a zaloguje).
    """
    force: bool = kwargs.get("force", False)
    force_lang: str = kwargs.get("force_lang", "") or ""

    if not settings.get("llm_enabled"):
        if force:
            raise RuntimeError("LLM is not enabled — enable it in Settings → LM Studio.")
        return 0
    if not force and not settings.get("llm_auto_translate"):
        return 0  # auto-translate vypnut; force=True toto přeskočí

    # Languageová kontrola: pro force překlad povolíme i unknown jazyk (""),
    # jen explicitní English (en/eng/…) je vždy přeskočena.
    _lang_norm = (force_lang or doc_lang or "").strip().lower()
    _explicit_english = {"en", "eng", "english", "en-gb", "en-us"}
    if _lang_norm in _explicit_english:
        return 0
    # Při auto-překladu bez force: unknown jazyk přeskočit (mohlo by jít o angličtinu)
    if not force and _lang_norm == "":
        return 0

    raw_block: str = kwargs.get("block", "") or ""

    existing = get_candidate_fields(candidate_id)

    # Sestavit seznam fields k překladu
    to_translate = {
        fname: fval
        for fname, fval in existing.items()
        if fname not in _TRANSLATION_SKIP_FIELDS
        and fval and fval not in (NOT_PROVIDED, "")
        and len(fval.strip()) > 10
    }

    # ── Fallback: nejsou namapovaná pole → přelož raw blok a auto-mapuj ──
    if not to_translate:
        if not raw_block or len(raw_block.strip()) < 20:
            return 0   # nemáme ani raw blok, vzdáme to
        # Přeložíme celý raw blok jako jeden text
        _lang_for_prompt = (force_lang or doc_lang or "").strip().lower()
        lang_name_fb = {
            "cs": "Czech", "cz": "Czech",
            "ru": "Russian", "rus": "Russian",
            "de": "German", "deu": "German",
            "fr": "French", "fra": "French",
            "zh": "Chinese", "zho": "Chinese", "chi": "Chinese",
            "la": "Latin", "lat": "Latin",
            "pl": "Polish", "sv": "Swedish", "sk": "Slovak",
            "it": "Italian", "es": "Spanish", "hu": "Hungarian",
        }.get(_lang_for_prompt, _lang_for_prompt.upper() or "unknown")
        fb_prompt = (
            f"Translate the following palaeontological taxonomic text from "
            f"{lang_name_fb} to English. Preserve all Latin taxon names, "
            f"author names, catalogue numbers and stratigraphic unit names verbatim. "
            f"Return ONLY the translated text, no commentary.\n\n{raw_block}"
        )
        try:
            fb_resp = lm_chat(settings, "", fb_prompt)
        except Exception as _e:
            if force:
                raise
            logging.warning("auto_translate raw-block fallback: %s", _e)
            return 0
        translated_block = fb_resp.strip()
        if not translated_block:
            return 0
        # Save přeložený blok do candidate_fields jako RAW_TRANSLATION
        save_fields(candidate_id,
                    {"RAW_TRANSLATION": translated_block},
                    method="auto_translated_raw")
        # Pokusit se z přeloženého textu extrahovat pole
        n_mapped = auto_map_candidate_fields(candidate_id, translated_block)
        # Přemapovat sekce (Characteristics → DESCRIPTION atd.)
        remap_fields_post_translation(candidate_id)
        # Vrátit 1 (přeložili jsme blok) + count namapovaných fields
        return 1 + n_mapped

    # Sestavit user prompt s hodnotami k překladu
    fields_block = "\n".join(
        f"{fname}:\n{val}" for fname, val in to_translate.items()
    )
    _lang_for_prompt = (force_lang or doc_lang or "").strip().lower()
    lang_name = {
        "cs": "Czech", "cz": "Czech",
        "ru": "Russian", "rus": "Russian",
        "de": "German", "deu": "German",
        "fr": "French", "fra": "French",
        "zh": "Chinese", "zho": "Chinese", "chi": "Chinese",
        "la": "Latin", "lat": "Latin",
        "pl": "Polish", "pol": "Polish",
        "sk": "Slovak", "slk": "Slovak",
        "sv": "Swedish", "swe": "Swedish",
        "it": "Italian", "ita": "Italian",
        "es": "Spanish", "spa": "Spanish",
        "hu": "Hungarian", "hun": "Hungarian",
    }.get(_lang_for_prompt, _lang_for_prompt.upper() or "unknown")

    user_prompt = (
        f"Source language: {lang_name}\n\n"
        f"Fields to translate:\n\n{fields_block}"
    )
    system_prompt = settings.get("llm_translation_prompt", LLM_TRANSLATION_PROMPT)

    try:
        raw = lm_chat(settings, system_prompt, user_prompt)
        result = lm_parse_json(raw)
    except Exception as exc:
        logging.warning(f"Translation LLM call failed for candidate {candidate_id}: {exc}")
        if force:
            raise  # UI dostane přesný popis chyby (RuntimeError z lm_chat)
        return 0

    if not result or "fields" not in result:
        if force:
            raise RuntimeError(
                "LLM nevrátilo platný JSON s překladem. "
                "Zkuste jiný model or snižte teplotu na 0.")
        return 0

    translated_fields = result["fields"]
    n_translated = 0
    to_save: Dict[str, str] = {}

    for fname, en_text in translated_fields.items():
        if not en_text or not en_text.strip():
            continue
        original = to_translate.get(fname, "")
        en_clean = en_text.strip()

        # Formát: přeložený anglický text (puvodní originální text)
        # Pokud je překlad identický s originálem (LLM vrátil beze změny),
        # nevkládat závorky — netřeba duplikovat text.
        if _norm_search(en_clean) == _norm_search(original):
            combined = original   # beze změny
        else:
            # Originál do závorek — zkrátit pokud je velmi dlouhý (>400 znaku)
            orig_display = (
                original[:400] + "…" if len(original) > 400 else original
            )
            combined = f"{en_clean} ({orig_display})"

        to_save[fname] = combined
        n_translated += 1

    if to_save:
        save_fields(candidate_id, to_save, method="auto_translated")

    # Po překladu přemapuj sekce na standardní DB pole
    # (e.g. "CHARACTERISTICS" → "DESCRIPTION", "DISCUSSION AND COMPARISON" → "DISCUSSION")
    remap_fields_post_translation(candidate_id)

    return n_translated



# ══════════════════════════════════════════════════════════════════════════════
# FEATURE 1: Parser katalogových čísel typového materiálu
# ══════════════════════════════════════════════════════════════════════════════

# Slovník institucí/depozitářu. Klíč = kód, hodnota = plné jméno.
_INSTITUTION_CODES: Dict[str, str] = {
    # Ruské / sovětské
    "ЯФАН": "Yakutian Branch AN SSSR", "ЯФАН": "Yakutian Branch AN SSSR",
    "ГИН": "Geological Institute RAS", "ПИН": "Palaeontological Institute RAS",
    "ЦНИГР": "Central Research Geological Museum",
    # Mezinárodní
    "USNM": "Smithsonian Institution NMNH", "NHMUK": "Natural History Museum London",
    "BMNH": "British Museum Natural History", "MCZ": "Harvard MCZ",
    "AMNH": "American Museum of Natural History", "YPM": "Yale Peabody Museum",
    "NIGP": "Nanjing Institute of Geology and Palaeontology",
    "IVPP": "Institute of Vertebrate Paleontology and Paleoanthropology",
    # České
    "NM": "National Museum Prague", "NMP": "National Museum Prague",
    "CGS": "Czech Geological Survey", "CGS": "Czech Geological Survey",
    # Polské / švédské / německé
    "ZPAL": "Institute of Paleobiology Warsaw",
    "LO": "Lund University Collections", "SGU": "Swedish Geological Survey",
    "SMF": "Senckenberg Research Institute",
    # Obecné
    "LM": "Local Museum (unspecified)",
}

# Regex pro catalog numbers (flexibilní – podporuje RU, EN, ZH styly)
_CATNO_RE = re.compile(
    r"""
    (?:
        # Full institutional code (2–6 uppercase) + optional collection series
        # (1–2 letters, e.g. SNM *Z* 15784, PIN *N* 3302/25) + catalogue number.
        (?P<inst1>[A-ZА-ЯЁ]{2,6})
        (?:\s+(?P<coll>[A-ZА-ЯЁ]{1,2}))?
        \s*[№#]?\s*(?:No\.?\s*)?
        (?P<no1>\d+(?:[\-\/][\dA-Z]+){0,3})
    |
        # Jednopísmenný prefix + víceciferné číslo (L 41315)
        (?P<inst3>[A-ZА-ЯЁ])\s+(?P<no3>\d{3,})
    |
        # Jen číslo s explicitním № or #
        [№#]\s*(?P<no2>\d+(?:[\-\/]\d+)?)
    )
    """,
    re.VERBOSE | re.UNICODE
)

_TYPE_KIND_RE = re.compile(
    r"(?:^|(?<=[^A-Za-z\u0430-\u044f\u0451]))"
    r"("
    # EN/international base (…type), + German …typus, + optional plural s/y/en
    r"Holotypus|Lectotypus|Nootypus|Syntypus|Paratypus|Paralectotypus"
    r"|Holotype|Lectotype|Nootype|Syntype|Paratype|Paralectotype|Topotype|Allotype"
    # Czech plurals: Syntypy, Paratypy, Holotypy (…typ + y)
    r"|Syntypy|Paratypy|Holotypy|Lektotyp|Paratyp|Holotyp|Syntyp|Nootyp"
    # Russian
    r"|Голотип|Лектотип|Паратип|Неотип|Синтип|Паралектотип"
    # Chinese
    r"|正模标本|正模|副模标本|副模|选模|新模|地模|全模|模式标本"
    r")",
    re.IGNORECASE | re.UNICODE)



def _parse_type_specimens(text: str) -> Dict[str, str]:
    """
    Extracts structured fields from the verbatim TYPE SPECIMENS text:
      - type_kinds: seřazený seznam typu (Holotype, Paratype…)
      - institution_codes: kódy depozitářu (ЯФАН, GIN, USNM…)
      - catalog_numbers: katalogová čísla (32/12, 390614…)

    Příklady vstupu:
      "Holotype: ЯФАН № 32/12, р. Алдан"
      "Holotype USNM 390614; Paratype USNM 390615-390618"
      "正模: NIGP 134291"
    """
    if not text:
        return {"type_kinds": "", "institution_codes": "", "catalog_numbers": ""}

    kinds = []
    seen_kinds: set = set()
    for m in _TYPE_KIND_RE.finditer(text):
        k = m.group(1).strip()
        _kl = k.lower()
        # Normalizovat allchny jazykové varianty na kanonickou angličtinu.
        # Klíčem je lowercase forma (regex je IGNORECASE).
        _norm_map = {
            # German -typus → -type
            "holotypus": "Holotype", "lectotypus": "Lectotype",
            "neotypus": "Nootype", "syntypus": "Syntype",
            "paratypus": "Paratype", "paralectotypus": "Paralectotype",
            # Czech (-typ / -typy plural)
            "holotyp": "Holotype", "holotypy": "Holotype",
            "lektotyp": "Lectotype", "neotyp": "Nootype",
            "syntyp": "Syntype", "syntypy": "Syntype",
            "paratyp": "Paratype", "paratypy": "Paratype",
            # Russian
            "голотип": "Holotype", "лектотип": "Lectotype",
            "паратип": "Paratype", "неотип": "Nootype",
            "синтип": "Syntype", "паралектотип": "Paralectotype",
            # Chinese
            "正模": "Holotype", "正模标本": "Holotype",
            "副模": "Paratype", "副模标本": "Paratype",
            "选模": "Lectotype", "新模": "Nootype",
            "地模": "Topotype", "全模": "Syntype",
            "模式标本": "Type specimen",
        }
        k = _norm_map.get(_kl, k.title() if k.isascii() else k)
        if k.lower() not in seen_kinds:
            seen_kinds.add(k.lower())
            kinds.append(k)

    institutions: List[str] = []
    catalog_nos: List[str] = []
    seen_inst: set = set()
    seen_no: set = set()
    for m in _CATNO_RE.finditer(text):
        inst = (m.group("inst1") or m.group("inst3") or "").strip().upper()
        coll = (m.group("coll") or "").strip().upper()
        no_core = (m.group("no1") or m.group("no2") or m.group("no3") or "").strip()
        # Sběrová série se přidá k číslu (SNM Z 15784 → katalog "Z 15784")
        no = (f"{coll} {no_core}".strip() if coll else no_core)
        if inst and inst not in seen_inst:
            # Ověřit délku (příliš krátká písmena jsou false positives).
            # Jednopísmenný institucionální prefix (inst3) je povolen jen
            # když má číslo aspoň 3 cifry (viz regex) — tady stačí ≥1 znak.
            if len(inst) >= 1:
                seen_inst.add(inst)
                institutions.append(inst)
        if no and no not in seen_no and len(no_core) >= 2:
            seen_no.add(no)
            catalog_nos.append(no)

    return {
        "type_kinds":        "; ".join(kinds),
        "institution_codes": "; ".join(institutions),
        "catalog_numbers":   "; ".join(catalog_nos),
    }


# ══════════════════════════════════════════════════════════════════════════════
# FEATURE 2: Parser rozměru (SIZE field)
# ══════════════════════════════════════════════════════════════════════════════

# Mapování variant měřených veličin na kanonické klíče
_SIZE_PARAM_ALIASES: Dict[str, str] = {
    # Délka
    "length": "length", "lenght": "length", "l": "length",
    "длина": "length", "壳长": "length", "lon": "length",
    # Šířka
    "width": "width", "w": "width", "ширина": "width", "壳宽": "width",
    "šířka": "width",
    # Výška
    "height": "height", "h": "height", "высота": "height", "壳高": "height",
    "výška": "height", "haut": "height",
    # Pruměr
    "diameter": "diameter", "diam": "diameter", "диаметр": "diameter",
    # Apikální úhel
    "apical angle": "apical_angle", "angle": "apical_angle",
    "угол расхождения": "apical_angle", "угол": "apical_angle",
    "生长角": "apical_angle", "growth angle": "apical_angle",
    # W/H ratio
    "w/h": "wh_ratio", "w/h ratio": "wh_ratio", "ш/в": "wh_ratio",
    "отношение ш/в": "wh_ratio", "切面比率": "wh_ratio",
    # Tloušťka stěny
    "wall thickness": "wall_thickness", "thickness": "wall_thickness",
    "толщина стенки": "wall_thickness",
}

# Číselný pattern — podporuje desetinnou čárku (RU) i tečku (EN)
_NUMBER_RE = re.compile(r"\d+[.,]\d+|\d+")

# Inline measurement: "length 5.2 mm", "Длина 6,75 мм", "壳长 1.9 毫米"
_INLINE_MEAS_RE = re.compile(
    r"(?P<param>"
    + "|".join(re.escape(p) for p in sorted(_SIZE_PARAM_ALIASES, key=len, reverse=True))
    + r")"
    r"\s*[=:–—]?\s*"
    r"(?P<val>\d+[.,]?\d*(?:\s*[–—-]\s*\d+[.,]?\d*)?)"
    r"\s*(?P<unit>mm|мм|毫米|cm|°|deg)?",
    re.IGNORECASE | re.UNICODE
)


def _parse_size_field(text: str) -> Dict[str, str]:
    """
    Extracts structured dimensions from the verbatim SIZE text.
    Returns a dict {canonický_parametr: "hodnota jednotka"}.

    Podporuje:
      - Inline: "length 21.0 mm, apical angle 8°, W/H = 1.4"
      - Ruský styl: "Длина 6,75 мм, Ширина 0,90 мм, Угол 3,5°"
      - Čínský styl: "壳长 1.9 毫米，口端宽 0.48 毫米"
      - ASCII tabulky (hlavička + datové řádky)
    """
    if not text:
        return {}

    result: Dict[str, str] = {}

    # Normalizace: desetinná čárka RU/CS → tečka (6,75 → 6.75).
    # re.sub s backreferencí — pozor na escaping.
    normalized = re.sub(r"(\d),\s*(\d)", lambda m: m.group(1)+"."+m.group(2), text)

    for m in _INLINE_MEAS_RE.finditer(normalized):
        param_raw = m.group("param").lower().strip()
        canonical = _SIZE_PARAM_ALIASES.get(param_raw, param_raw)
        val = m.group("val").replace(",", ".").strip()
        unit = (m.group("unit") or "").strip()
        if unit in ("мм", "毫米"):
            unit = "mm"
        stored = f"{val} {unit}".strip() if unit else val
        if canonical not in result:
            result[canonical] = stored

    # Pokus o ASCII tabulku: hledáme řádky se dvěma a více čísly
    lines = text.split("\n")
    # Heuristika: řádky kde jsou min. 2 čísla oddělená mezerami → tabulka
    number_rows = [l for l in lines if len(_NUMBER_RE.findall(l)) >= 2]
    if len(number_rows) >= 2:
        # Přidat jako surový řetězec pro ruční kontrolu
        if "table_raw" not in result:
            result["table_raw"] = " | ".join(r.strip() for r in number_rows[:5])

    return result


def _size_dict_to_str(d: Dict[str, str]) -> str:
    """Formats the result of _parse_size_field as a human-readable string for export."""
    if not d:
        return ""
    parts = [f"{k}: {v}" for k, v in d.items() if k != "table_raw"]
    if "table_raw" in d:
        parts.append(f"[table: {d['table_raw'][:80]}]")
    return "; ".join(parts)


# ══════════════════════════════════════════════════════════════════════════════
# FEATURE 4: Validace completenessi záznamu podle ranku
# ══════════════════════════════════════════════════════════════════════════════

# Pole požadovaná pro každý rank. Hodnoty jsou (field, weight) kde weight=1 je
# povinné a weight=0.5 je doporučené.
_REQUIRED_FIELDS_BY_RANK: Dict[str, List[Tuple[str, float]]] = {
    # PRAVIDLO: species/subspecies mají TYPE SPECIMENS (konkrétní exempláře)
    #           genus a vyšší mají TYPE TAXON (typový druh / rod)
    # Pozn.: LOCALITY pro druhy je optional (0.5) — data bývají v OCCURRENCE.
    # _completeness_check() akceptuje OCCURRENCE jako alternativu k LOCALITY.
    "species": [
        ("DIAGNOSIS",       1.0),
        ("TYPE SPECIMENS",  1.0),  # HOLOTYP, PARATYPUS atd. — jen pro species/subspecies
        ("LOCALITY",        0.5),  # alternativa: OCCURRENCE
        ("STRATIGRAPHY",    0.8),
        ("DESCRIPTION",     0.8),
        ("FIGURES",         0.5),
        ("OCCURRENCE",      0.5),
    ],
    # Subspecies: identické požadavky jako species — TYPE SPECIMENS (ne TYPE TAXON!)
    "subspecies": [
        ("DIAGNOSIS",       1.0),
        ("TYPE SPECIMENS",  1.0),
        ("LOCALITY",        0.5),
        ("STRATIGRAPHY",    0.8),
        ("DESCRIPTION",     0.8),
        ("FIGURES",         0.5),
        ("OCCURRENCE",      0.5),
    ],
    "genus": [
        ("TYPE TAXON",      1.0),  # typový druh — jen pro genus a vyšší
        ("DIAGNOSIS",       1.0),
        ("DESCRIPTION",     0.5),
        ("OCCURRENCE",      0.5),
        ("INCLUDED TAXONS", 0.5),
    ],
    "family": [
        ("TYPE TAXON",      1.0),  # typový rod
        ("DIAGNOSIS",       1.0),
        ("INCLUDED TAXONS", 0.5),
        ("OCCURRENCE",      0.5),
    ],
    "order": [
        ("TYPE TAXON",      0.8),  # typová čeleď (méně povinné)
        ("DIAGNOSIS",       1.0),
        ("INCLUDED TAXONS", 0.5),
        ("OCCURRENCE",      0.5),
    ],
    "class": [
        ("TYPE TAXON",      0.8),
        ("DIAGNOSIS",       1.0),
        ("INCLUDED TAXONS", 0.5),
    ],
    "phylum": [
        ("DIAGNOSIS",       1.0),
        ("INCLUDED TAXONS", 0.5),
    ],
}

# Zástupná jména ranku pro normalizaci.
# KLÍČOVÉ PRAVIDLO: species + subspecies → TYPE SPECIMENS (ne TYPE TAXON)
#                   genus + vyšší → TYPE TAXON (ne TYPE SPECIMENS)
_RANK_ALIASES: Dict[str, str] = {
    # species-grade (potřebují TYPE SPECIMENS)
    "species": "species", "druh": "species", "вид": "species", "art": "species",
    "espèce": "species", "specie": "species", "especie": "species",
    "subspecies": "species",   # poddruh → stejné požadavky jako druh (TYPE SPECIMENS!)
    "subsp": "species", "ssp": "species", "var": "species", "variety": "species",
    "forma": "species", "form": "species",
    # genus-grade (potřebují TYPE TAXON)
    "genus": "genus", "rod": "genus", "род": "genus", "gattung": "genus",
    "subgenus": "genus", "subgen": "genus", "podrod": "genus",
    # family-grade (potřebují TYPE TAXON)
    "family": "family", "čeleď": "family", "семейство": "family", "famille": "family",
    "familia": "family", "familie": "family",
    "subfamily": "family", "subfamilia": "family",
    "tribe": "family", "tribus": "family",
    "superfamily": "family", "superfamilia": "family",
    # order-grade
    "order": "order", "řád": "order", "отряд": "order", "ordnung": "order",
    "ordo": "order",
    "suborder": "order", "infraorder": "order", "superorder": "order",
    # class-grade
    "class": "class", "třída": "class", "класс": "class", "klasse": "class",
    "classis": "class",
    "subclass": "class", "infraclass": "class",
    # phylum-grade
    "phylum": "phylum", "kmen": "phylum", "тип": "phylum",
    "subphylum": "phylum",
    # P 6.1: Russian rank aliases
    "подрод": "genus", "подсемейство": "family", "подвид": "species",
    "отдел": "order", "подотряд": "order",
    # P 7.1: Chinese rank aliases
    "亚属": "genus", "亚科": "family", "亚种": "species", "亚目": "order",
    "亚纲": "class", "亚门": "phylum",
}


def _completeness_check(
    rank_raw: str,
    fields: Dict[str, str],
) -> Tuple[float, List[str]]:
    """
    Checks the completeness of a record for the given rank.
    Returns (score 0–1, list of missing fivinných fields).

    Score: součet vah přítomných fields / součet vah allch požadovaných fields.
    Pole se považuje za přítomné pokud exists a není NOT_PROVIDED.
    """
    # Normalizace ranku: subspecies má explicitní klíč v _REQUIRED_FIELDS_BY_RANK,
    # proto zkusíme nejprve přímou shodu (zachová subspecies jako subspecies,
    # ne jen alias na species).
    rank_raw_norm = (rank_raw or "").lower().split()[0]
    if rank_raw_norm in _REQUIRED_FIELDS_BY_RANK:
        rank = rank_raw_norm          # přesná shoda (subspecies, species, genus…)
    else:
        rank = _RANK_ALIASES.get(rank_raw_norm, "")
    if not rank or rank not in _REQUIRED_FIELDS_BY_RANK:
        return 1.0, []  # unknown rank → žádné požadavky

    required = _REQUIRED_FIELDS_BY_RANK[rank]
    total_weight = sum(w for _, w in required)
    if total_weight == 0:
        return 1.0, []

    # Pro species/subspecies: LOCALITY je alternativou OCCURRENCE
    _species_grade = rank in ("species", "subspecies")
    occ_filled = bool((fields.get("OCCURRENCE","") or "").strip() and
                      fields.get("OCCURRENCE","") != NOT_PROVIDED)

    # TYPE SPECIMENS se považují za přítomné pokud je v MATERIAL EXAMINED
    # zmíněn holotyp/paratyp/lektotyp — jen pro species-grade taxony
    _MAT = (fields.get("MATERIAL EXAMINED","") or "").lower()
    _type_in_material = any(kw in _MAT for kw in TYPE_SPECIMEN_KEYWORDS)

    filled_weight = 0.0
    missing: List[str] = []
    for field, weight in required:
        val = fields.get(field, "")
        filled = bool(val and val.strip() and val != NOT_PROVIDED)
        # LOCALITY pro species/subspecies – akceptuj OCCURRENCE jako alternativu
        if field == "LOCALITY" and not filled and _species_grade and occ_filled:
            filled_weight += weight
            continue
        # TYPE SPECIMENS pro species/subspecies – akceptuj holotype/paratype v MATERIAL EXAMINED
        if field == "TYPE SPECIMENS" and not filled and _species_grade and _type_in_material:
            filled_weight += weight
            continue   # nepřidávat do missing (data jsou v MATERIAL EXAMINED)
        if filled:
            filled_weight += weight
        elif weight >= 1.0:
            missing.append(field)

    # Score: povinná pole jsou základ (weight ≥ 1.0),
    # volitelná tvoří bonus. Povinná pole určují "completeness",
    # volitelná ji vylepšují max o 20 %.
    # POZOR: alternativy (OCCURRENCE ≡ LOCALITY, MATERIAL EXAMINED ≡ TYPE SPECIMENS)
    # jsou zahrnuty jak v loop výše (missing list) tak i v score výpočtu níže.
    mandatory_w = sum(w for _, w in required if w >= 1.0)
    optional_w  = sum(w for _, w in required if w < 1.0)

    def _fld_filled(f: str, w: float) -> bool:
        """Checks whether field f is filled, including alternatives for species-grade records."""
        val = fields.get(f, "")
        filled = bool(val and val.strip() and val != NOT_PROVIDED)
        if not filled and _species_grade:
            # LOCALITY → akceptuj OCCURRENCE jako alternativu
            if f == "LOCALITY" and occ_filled:
                return True
            # TYPE SPECIMENS → akceptuj holotyp/paratyp v MATERIAL EXAMINED
            if f == "TYPE SPECIMENS" and _type_in_material:
                return True
        return filled

    mand_filled = sum(w for f, w in required if w >= 1.0 and _fld_filled(f, w))
    opt_filled  = sum(w for f, w in required if w < 1.0  and _fld_filled(f, w))
    mand_score = mand_filled / mandatory_w if mandatory_w else 1.0
    opt_bonus  = (opt_filled / optional_w * 0.2) if optional_w else 0.0
    score = round(min(1.0, mand_score * 0.8 + opt_bonus), 3)
    return score, missing


# ══════════════════════════════════════════════════════════════════════════════
# FEATURE 5: Cross-reference synonym → propojení records v rámci Documentu
# ══════════════════════════════════════════════════════════════════════════════

# Extrakce taxonových jmen ze synonymiky
_SYN_TAXON_RE = re.compile(
    r"(?:^|\n)"
    r"(?:\d{4}\s+)?"                              # volitelný rok na začátku
    r"(?P<name>"
    r"[A-Z][a-z]+"                                  # rod
    r"(?:\s+[a-z][a-z\-]+)?"                      # volitelný druh
    r")",
    re.MULTILINE | re.UNICODE
)


def _extract_names_from_synonymy(synonymy_text: str) -> set:
    """
    Extracts a set of taxon names from the synonymy section.
    Goal: link records withini stejného Documentu.
    """
    names = set()
    for m in _SYN_TAXON_RE.finditer(synonymy_text or ""):
        name = m.group("name").strip()
        if len(name) > 3 and name.lower() not in TAXON_STOPWORDS:
            names.add(name.lower())
    return names


def _link_synonyms_for_document(document_id: int) -> int:
    """
    For all approved candidates of the document:
    1. Reads their SYNONYMY field.
    2. Extrahuje jména.
    3. Pokud jiný kandidát v témž Documentu nese toto jméno → nastaví
       RELATED_RECORD_ID na propojený záznam.
    Vrací count nových propojení.
    """
    con = db()
    # Všichni kandidáti Documentu s jejich poli
    cands = con.execute(
        "SELECT id, taxon_name, rank_guess FROM taxon_candidates "
        "WHERE document_id=? AND status IN ('approved','pending','low_confidence')",
        (document_id,)
    ).fetchall()

    if not cands:
        con.close()
        return 0

    # Načíst SYNONYMY pole pro každého kandidáta
    cand_ids = [c["id"] for c in cands]
    syn_rows = con.execute(
        f"SELECT candidate_id, field_value FROM occurrence_fields "
        f"WHERE field_name='SYNONYMY' AND candidate_id IN ({','.join('?'*len(cand_ids))})",
        cand_ids
    ).fetchall()
    syn_map: Dict[int, str] = {r["candidate_id"]: r["field_value"] for r in syn_rows}

    # Index: taxon_name.lower() → candidate_id
    name_to_id: Dict[str, int] = {
        c["taxon_name"].lower().strip(): c["id"] for c in cands
        if c["taxon_name"]
    }
    # Přidat i jen rodové jméno (první slovo) pro větší pokrytí
    for c in cands:
        first = (c["taxon_name"] or "").split()[0].lower().strip()
        if first and first not in name_to_id:
            name_to_id[first] = c["id"]

    links_added = 0
    for cand in cands:
        synonymy = syn_map.get(cand["id"], "")
        if not synonymy:
            continue
        mentioned_names = _extract_names_from_synonymy(synonymy)
        for mname in mentioned_names:
            related_id = name_to_id.get(mname)
            if related_id and related_id != cand["id"]:
                # Save jako RELATED_RECORD_ID (pokud ještě není)
                existing = con.execute(
                    "SELECT id FROM occurrence_fields "
                    "WHERE candidate_id=? AND field_name='RELATED_RECORD_ID'",
                    (cand["id"],)
                ).fetchone()
                if not existing:
                    con.execute(
                        "INSERT INTO occurrence_fields "
                        "(candidate_id, field_name, field_value, method) VALUES (?,?,?,?)",
                        (cand["id"], "RELATED_RECORD_ID",
                         str(related_id), "auto_synonym_link")
                    )
                    links_added += 1
    con.commit()
    con.close()
    return links_added

def build_output_row(cand: sqlite3.Row, fields: Dict[str, str], doc_row: sqlite3.Row) -> Dict[str, str]:
    """Builds a complete output row with all OUTPUT_FIELDS."""
    # sqlite3.Row nepodporuje .get() s defaultem → převedeme na dict hned na začátku
    cand    = dict(cand)
    doc_row = dict(doc_row)
    row: Dict[str, str] = {f: NOT_PROVIDED for f in OUTPUT_FIELDS}
    row["RECORD_ID"]          = f"{str(doc_row['id'])}-{str(cand['id'])}"
    row["SOURCE_DOCUMENT"]    = doc_row["filename"]
    row["SOURCE_DOCUMENT_SHORT_NAME"] = pathlib.Path(doc_row["filename"]).stem
    row["SOURCE_LANGUAGE"]    = doc_row["lang"] or NOT_PROVIDED
    row["TAXON_NAME_VERBATIM"] = cand["taxon_name"]
    row["TAXON_RANK_AS_WRITTEN"] = cand["rank_guess"] or NOT_PROVIDED
    row["RAW_TAXONOMIC_BLOCK"] = cand["block_text"] or NOT_PROVIDED
    row["BLOCK_START_LOCATOR"] = f"Page {cand['page_start']}"
    row["EXTRACTION_CONFIDENCE"] = str(round(cand["confidence"],3))
    row["SOURCE PAGES"]        = str(cand["page_start"])

    # Přiřazená pole
    for fname, fval in fields.items():
        if fname in row:
            row[fname] = fval

    # Fallback AUTHOR z nadpisu — extrahuj i rok zvlášť
    if row["AUTHOR"] == NOT_PROVIDED:
        m = AUTHOR_YEAR_RE.search(cand["heading_text"] or "")
        if m:
            row["AUTHOR"] = m.group(0)
    # Rok publikace: pokud AUTHOR_YEAR_RE má rok, saveíme ho zvlášť
    if row.get("YEAR_OF_PUBLICATION", NOT_PROVIDED) == NOT_PROVIDED:
        _hy = cand.get("heading_text") or cand.get("taxon_name") or ""
        _ym = re.search(r"\b(1[5-9]\d{2}|20[012]\d)\b", _hy)
        if _ym:
            row["YEAR_OF_PUBLICATION"] = _ym.group(1)

    # Open nomenclature
    if row["OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS"] == NOT_PROVIDED:
        m = OPEN_NOMEN_RE.search(cand["taxon_name"] or "")
        if m:
            row["OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS"] = m.group(0)

    # Feature 1: TYPE SPECIMENS → strukturované katalogové číslo
    type_spec_text = row.get("TYPE SPECIMENS", "") or ""
    if type_spec_text and type_spec_text != NOT_PROVIDED:
        ts = _parse_type_specimens(type_spec_text)
        if ts["type_kinds"]:        row["TYPE_SPECIMEN_KIND"] = ts["type_kinds"]
        if ts["institution_codes"]: row["INSTITUTION_CODE"]   = ts["institution_codes"]
        if ts["catalog_numbers"]:   row["CATALOG_NUMBER"]     = ts["catalog_numbers"]

    # Feature 2: SIZE → strukturované rozměry
    size_text = row.get("SIZE", "") or ""
    if size_text and size_text != NOT_PROVIDED:
        size_d = _parse_size_field(size_text)
        if size_d:
            row["SIZE_PARSED"] = _size_dict_to_str(size_d)

    # Feature 3: parent-child hierarchie
    parent_name = str(cand["parent_taxon_name"] or "") if "parent_taxon_name" in cand.keys() else ""
    parent_rank = str(cand["parent_rank"] or "")      if "parent_rank"       in cand.keys() else ""
    if parent_name: row["PARENT_TAXON_NAME"] = parent_name
    if parent_rank: row["PARENT_RANK"]       = parent_rank

    # Feature 4: completeness validation
    comp_score, missing_f = _completeness_check(
        cand["rank_guess"] or "", fields)
    row["COMPLETENESS_SCORE"]     = str(comp_score)
    row["MISSING_REQUIRED_FIELDS"] = "; ".join(missing_f) if missing_f else ""

    return row


def export_to_xlsx(
    rows: List[Dict[str, str]],
    review_rows: List[Dict],
    references: Optional[List[str]] = None,
) -> bytes:
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl is not installed.")
    wb = openpyxl.Workbook()
    h_fill = PatternFill("solid", fgColor="1F4E79")
    h_font = Font(bold=True, color="FFFFFF", size=9)
    h_aln  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c_aln  = Alignment(vertical="top", wrap_text=True)

    # Sheet 1 – Taxonomic extraction
    ws = wb.active
    ws.title = "Taxonomic extraction"
    ws.append(OUTPUT_FIELDS)
    for cell in ws[1]:
        cell.fill, cell.font, cell.alignment = h_fill, h_font, h_aln
    for row in rows:
        ws.append([row.get(f, NOT_PROVIDED) for f in OUTPUT_FIELDS])
        for cell in ws[ws.max_row]:
            cell.alignment = c_aln
    for i, col in enumerate(OUTPUT_FIELDS, 1):
        w = 60 if col in ("RAW_TAXONOMIC_BLOCK","DIAGNOSIS","DESCRIPTION","SYNONYMY") else 22
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Sheet 1b – Parsed fields (Feature 1, 2, 3, 4)
    ws1b = wb.create_sheet("Parsed fields")
    parsed_cols = ["RECORD_ID", "TAXON_NAME_VERBATIM", "TAXON_RANK_AS_WRITTEN",
                   "PARENT_TAXON_NAME", "PARENT_RANK",
                   "TYPE_SPECIMEN_KIND", "INSTITUTION_CODE", "CATALOG_NUMBER",
                   "SIZE_PARSED",
                   "COMPLETENESS_SCORE", "MISSING_REQUIRED_FIELDS"]
    ws1b.append(parsed_cols)
    for cell in ws1b[1]:
        cell.fill = PatternFill("solid", fgColor="1A5276")
        cell.font = Font(bold=True, color="FFFFFF", size=9)
        cell.alignment = h_aln
    for row in rows:
        ws1b.append([row.get(f, NOT_PROVIDED) for f in parsed_cols])
        for cell in ws1b[ws1b.max_row]:
            cell.alignment = c_aln
    for i, col in enumerate(parsed_cols, 1):
        ws1b.column_dimensions[get_column_letter(i)].width = 30
    ws1b.freeze_panes = "A2"

    # Sheet 2 – Review / low confidence
    ws2 = wb.create_sheet("Review queue")
    review_cols = ["ID","Document","Taxon","Page","Confidence","Status","Heading","Notes"]
    ws2.append(review_cols)
    for cell in ws2[1]:
        cell.fill = PatternFill("solid", fgColor="843C0C")
        cell.font = Font(bold=True, color="FFFFFF", size=9)
    for r in review_rows:
        ws2.append([r.get(c, "") for c in review_cols])

    # Sheet 3 – Reference / literatura (volitelná, jen pokud bylo něco vybráno)
    if references:
        ws3 = wb.create_sheet("References")
        ws3.append(["#", "Reference"])
        for cell in ws3[1]:
            cell.fill = PatternFill("solid", fgColor="3f6212")
            cell.font = Font(bold=True, color="FFFFFF", size=9)
        for i, ref in enumerate(references, 1):
            ws3.append([i, ref])
            ws3.cell(row=ws3.max_row, column=2).alignment = c_aln
        ws3.column_dimensions["A"].width = 6
        ws3.column_dimensions["B"].width = 100

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_to_docx(rows: List[Dict[str, str]]) -> bytes:
    if not HAS_DOCX:
        raise RuntimeError("python-docx is not installed.")
    doc = _DocxDoc()
    doc.add_heading("PaleoN – Taxonomic Extraction Dossier", 0)
    doc.add_paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Records: {len(rows)}")
    ORDERED = [
        "TAXON_NAME_VERBATIM","TAXON_RANK_AS_WRITTEN","AUTHOR","SYNONYMY",
        "TYPE SPECIMENS","TYPE MATERIAL","MATERIAL EXAMINED","ETYMOLOGY",
        "LOCALITY","STRATIGRAPHY","DIAGNOSIS","DESCRIPTION","SIZE",
        "REMARKS","OCCURRENCE","FIGURES","REFERENCE",
        "SOURCE PAGES","EXTRACTION_NOTES",
        "RAW_TAXONOMIC_BLOCK",
    ]
    for row in rows:
        doc.add_heading(row.get("TAXON_NAME_VERBATIM","?"), level=1)
        tbl = doc.add_table(rows=0, cols=2)
        tbl.style = "Table Grid"
        for field in ORDERED:
            val = row.get(field, NOT_PROVIDED)
            if val and val != NOT_PROVIDED:
                cells = tbl.add_row().cells
                cells[0].paragraphs[0].add_run(field).bold = True
                cells[1].text = val
        doc.add_page_break()
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_to_json(rows: List[Dict[str, str]]) -> bytes:
    payload = {
        "app": APP_NAME, "version": APP_VERSION,
        "generated": datetime.now().isoformat(),
        "records": rows,
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


# ══════════════════════════════════════════════════════════════════════════════
# GOLD SET — parser DOCX + evaluátor
# ══════════════════════════════════════════════════════════════════════════════

# Mapování gold-set labels na kanonická PaleoN pole (SECTION_FIELDS).
# Klíče jsou lowercase bez teček — porovnáváme s label.lower().rstrip('.')
_GOLDSET_LABEL_TO_FIELD: Dict[str, str] = {
    # DIAGNOSIS
    "diagnosis": "DIAGNOSIS", "original diagnosis": "DIAGNOSIS",
    "emended diagnosis": "DIAGNOSIS", "revised diagnosis": "DIAGNOSIS",
    "amended diagnosis": "DIAGNOSIS", "differential diagnosis": "DIAGNOSIS",
    # DESCRIPTION
    "description": "DESCRIPTION", "redescription": "DESCRIPTION",
    "re-description": "DESCRIPTION", "general description": "DESCRIPTION",
    "morphology": "DESCRIPTION", "shell morphology": "DESCRIPTION",
    # SYNONYMY
    "synonymy": "SYNONYMY", "synonymy and citations": "SYNONYMY",
    "syn": "SYNONYMY", "history": "SYNONYMY",
    "nomenclatural history": "SYNONYMY", "taxonomic history": "SYNONYMY",
    # MATERIAL EXAMINED
    "material": "MATERIAL EXAMINED", "material examined": "MATERIAL EXAMINED",
    "studied material": "MATERIAL EXAMINED", "examined material": "MATERIAL EXAMINED",
    "specimens examined": "MATERIAL EXAMINED",
    # TYPE SPECIMENS
    "holotype": "TYPE SPECIMENS", "paratypes": "TYPE SPECIMENS",
    "type specimens": "TYPE SPECIMENS", "type material": "TYPE SPECIMENS",
    "type specimens described": "TYPE SPECIMENS",
    # ETYMOLOGY
    "etymology": "ETYMOLOGY",
    # SIZE
    "dimensions": "SIZE", "size": "SIZE", "measurements": "SIZE",
    # OCCURRENCE
    "occurrence": "OCCURRENCE", "distribution": "OCCURRENCE",
    "stratigraphic range": "OCCURRENCE",
    "stratigraphic range and distribution": "OCCURRENCE",
    "stratigraphic and geographic distribution": "OCCURRENCE",
    "geographical distribution": "OCCURRENCE",
    "stratigraphical range and distribution": "OCCURRENCE",
    # STRATIGRAPHY
    "type unit and locality": "STRATIGRAPHY",
    "type horizon and locality": "STRATIGRAPHY",
    "stratigraphy": "STRATIGRAPHY", "biostratigraphy": "STRATIGRAPHY",
    # LOCALITY
    "locality": "LOCALITY", "type locality": "LOCALITY",
    "collection locality": "LOCALITY",
    # TAXONOMIC PLACEMENT
    "class": "TAXONOMIC PLACEMENT", "order": "TAXONOMIC PLACEMENT",
    "family": "TAXONOMIC PLACEMENT", "genus": "TAXONOMIC PLACEMENT",
    "subgenus": "TAXONOMIC PLACEMENT",
    # TYPE TAXON
    "type species": "TYPE TAXON", "type genus": "TYPE TAXON",
    "type taxon": "TYPE TAXON", "type of genus": "TYPE TAXON",
    "type of the genus": "TYPE TAXON", "type of the family": "TYPE TAXON",
    # Cave bear paper / standard zool. nomenclature labels:
    "locus typicus": "LOCALITY", "basis data": "LOCALITY",
    "stratum typicum": "STRATIGRAPHY", "storage": "LOCALITY",
    "derivatio nominis": "ETYMOLOGY", "derivation of name": "ETYMOLOGY",
    "differential diagnosis": "DIAGNOSIS", "diagnosis differentialis": "DIAGNOSIS",
    "other sites":          "OCCURRENCE",
    "other species":        "INCLUDED TAXONS",    # "Other species." in genus blocks
    "other taxa":           "INCLUDED TAXONS",
    "material":             "MATERIAL EXAMINED",   # standalone "Material." in species
    "examined material":    "MATERIAL EXAMINED",
    "studied material":     "MATERIAL EXAMINED",
    "material examined":    "MATERIAL EXAMINED",
    "referred material":    "MATERIAL EXAMINED",
    "additional material":  "MATERIAL EXAMINED",
    # P 6.3: Russian section labels (extended)
    "характеристика": "DESCRIPTION", "геологическое распространение": "OCCURRENCE",
    "палеоэкология": "REMARKS", "местонахождение": "LOCALITY",
    "стратиграфическое и географическое распространение": "OCCURRENCE",
    "диагноз": "DIAGNOSIS", "описание": "DESCRIPTION",
    "сравнение": "REMARKS", "замечания": "REMARKS", "обсуждение": "REMARKS",
    "состав": "INCLUDED TAXONS",
    "геологическое и географическое распространение": "OCCURRENCE",
    "материал и местонахождение": "MATERIAL EXAMINED",
    "типовой вид": "TYPE TAXON", "типовой род": "TYPE TAXON",
    "размеры": "SIZE",
    # P 7.1: Chinese rank labels (新种/新属 → handled by NEW_TAXON_RE)
    # P 7.2: Chinese section labels (extended)
    "特征": "DIAGNOSIS", "分类特征": "DIAGNOSIS",
    "化石描述": "DESCRIPTION", "描述": "DESCRIPTION", "系统分类": "TAXONOMIC PLACEMENT",
    "地层与时代": "STRATIGRAPHY", "产地与层位": "LOCALITY",
    "产地与地层": "LOCALITY",
    "分布与时代": "OCCURRENCE", "分布时代": "OCCURRENCE",
    "时代和分布": "OCCURRENCE", "时代与分布": "OCCURRENCE",
    "比较": "REMARKS", "讨论": "REMARKS",
    "讨论与比较": "REMARKS", "比较与讨论": "REMARKS",
    "模式种": "TYPE TAXON", "模式属": "TYPE TAXON",
    "种名来源": "ETYMOLOGY", "属名来源": "ETYMOLOGY",
    "壳体度量": "SIZE", "保存壳长": "SIZE",
    "同物异名": "SYNONYMY",
    # INCLUDED TAXONS
    "species included":      "INCLUDED TAXONS",
    "included species":      "INCLUDED TAXONS",
    "included genera":       "INCLUDED TAXONS",
    "included genera.":      "INCLUDED TAXONS",
    "genera included":       "INCLUDED TAXONS",
    "included taxa":         "INCLUDED TAXONS",
    "genera":                "INCLUDED TAXONS",   # standalone "Genera." in family blocks
    "species":               "INCLUDED TAXONS",   # "Species." heading in genus treatments
    # REMARKS / DISCUSSION
    "discussion": "REMARKS", "remarks": "REMARKS",
    "notes": "REMARKS", "information": "REMARKS",
    "comparative remarks": "REMARKS", "comparative discussion": "REMARKS",
    # FIGURES
    "figures": "FIGURES",
}

_GOLDSET_KNOWN_LABELS: set = set(_GOLDSET_LABEL_TO_FIELD.keys())

_GOLDSET_MARKER_RE = re.compile(
    # Format A – EN: "Record N – rank"  (rank required; taxon name optional on same line)
    # Format B – CS: "Record N TaxonName"
    r"^(?:"
    r"Record\s+(\d+)\s*[-–—]+\s*"
    r"(genus|species|family|subfamily|class|subclass|order|suborder"
    r"|subgenus|phylum|tribe|superfamily|infraorder)"
    r"(?:\s+(.+?))?"
    r"|Z[aá]znam\s+(\d+)\s*(.+?)"
    r")$",
    re.IGNORECASE,
)
_GOLDSET_INLINE_RE = re.compile(
    r"^([A-ZÁČĎÉĚÍŇÓŘŠŤÚŮÝŽ][^:\n]{2,50}):\s+(.+)$",
)


def _infer_taxon_name_from_body(body: List[str], rank: str) -> str:
    """
    Derives the taxon name from the record body if it was not on the marker line.
    Used when formátu "Record N – rank" (jen rank na markeru, jméno v těle).

    Strategie:
      1. Hledá odstavec začínající rank-labelem shodným s rank
         ("Family Hyolithidae Smith, 1900" → "Hyolithidae Smith, 1900")
      2. Pokud odstavec začíná JINÝM rank-labelem (merged paragraph),
         hledá cílový rank-label uvnitř textu
         ("Class Foo 1926 Order Bar 1916" s rank=Order → "Bar 1916")
      3. Pro druh (species) detekuje binomium i s kvalifikátory (?, cf., aff.)
      4. Fallback: první neprázdný odstavec
      Výsledek čistí od "new genus/species/…" suffixu a obr. odkazu.
    """
    if not rank:
        return ""
    rank_lower = rank.lower()

    _rank_cats: Dict[str, set] = {
        "genus":       {"genus", "genu"},
        "subgenus":    {"subgenus"},
        "species":     set(),
        "family":      {"family", "familia"},
        "subfamily":   {"subfamily"},
        "order":       {"order", "ordo"},
        "suborder":    {"suborder"},
        "class":       {"class", "classis"},
        "subclass":    {"subclass"},
        "phylum":      {"phylum"},
        "tribe":       {"tribe"},
        "superfamily": {"superfamily"},
        "infraorder":  {"infraorder"},
    }
    acceptable = _rank_cats.get(rank_lower, {rank_lower})

    # Regex pro ořez nomenklaturních aktu a obrázkových odkazu na konci
    _NOM_STRIP = re.compile(
        r"\s+(?:new\s+(?:genus|species|subspecies|combination|family)|"
        r"(?:gen|sp|fam|ord|cl|subsp|subgen)\.?\s*(?:nov|n)\.?|"
        r"(?:gen|sp)\.?\s*n\.|n\.?\s*(?:gen|sp|fam)\.?|"
        r"comb\.?\s*(?:nov|n)\.?)"
        r".*$",
        re.IGNORECASE,
    )
    _FIG_STRIP = re.compile(
        r"\s+(?:Figs?\.?|Figures?|Pls?\.?|Plates?)\s*[\d\.,\-\u2013\u2014\s\u2013\u2014]+.*$",
        re.IGNORECASE,
    )

    def _clean(raw: str) -> str:
        s = raw.strip()
        s = _NOM_STRIP.sub("", s).strip()
        s = _FIG_STRIP.sub("", s).strip()
        return s

    # Regex pro inline vyhledávání rank-labelu uvnitř odstavce
    # (pro případ sloučených odstavcu jako "Class Foo 1926 Order Bar 1916")
    if acceptable:
        _labels_esc = "|".join(re.escape(lbl) for lbl in sorted(acceptable, key=len, reverse=True))
        _inline_rank_re = re.compile(
            r"\b(?:" + _labels_esc + r")\s+([A-Z\?!][^\n]{1,120})",
            re.IGNORECASE,
        )
    else:
        _inline_rank_re = None

    for para in body[:10]:
        para = para.strip()
        if not para:
            continue

        # ── 1. Rank-label na začátku odstavce ────────────────────────────
        rm = RANK_LABEL_RE.match(para)
        if rm:
            label_word = rm.group(1).lower()
            rest = rm.group(2).strip()
            if label_word in acceptable and rest:
                return _clean(rest.split("\n")[0])

            # ── 2. Rank-label uvnitř sloučeného odstavce ─────────────────
            # (RANK_LABEL_RE chytil jiný rank na začátku → hledáme dál)
            if _inline_rank_re:
                im = _inline_rank_re.search(para)
                if im:
                    candidate = im.group(1).strip()
                    # Ořez za dalším rank-labelem or koncem věty
                    candidate = re.split(
                        r"\s+\b(?:Class|Subclass|Phylum|Order|Suborder|Family|Subfamily"
                        r"|Genus|Subgenus|Tribe|Superfamily|Infraorder)\b",
                        candidate, flags=re.I,
                    )[0]
                    candidate = candidate.split("\n")[0].strip()
                    if candidate and candidate.lower() not in TAXON_STOPWORDS:
                        return _clean(candidate)

        # ── 3. Detekce druhu / poddruhu (binomium s volitelnými kvalif.) ─
        elif rank_lower in ("species", "subspecies"):
            # Zachytí: "Genus epithet", "Genus ? epithet", "Genus cf. epithet",
            #           "Genus aff. epithet", "Genus ? cf. epithet", …
            sp_re = re.compile(
                r"^([A-Z][A-Za-z\u00c0-\u024f\-]+)"    # genus (s diakritikou)
                r"(?:\s+(?:\?|cf\.|aff\.?|sensu|ex|nov\.?)){0,2}"  # 0–2 kvalif.
                r"\s+([a-z\?][A-Za-z\u00c0-\u024f\.\-]*)"         # epithet
            )
            sm = sp_re.match(para)
            if sm:
                raw = para.split("(")[0].strip()   # ořez "(Autor, rok)"
                return _clean(raw[:150])
            # Taky "sp." samostatně: "Cavernolites sp."
            sp_only = re.compile(r"^([A-Z][A-Za-z\-]+)\s+sp\.\s*$", re.I)
            if sp_only.match(para):
                return para.strip()
            # Poslední záchrana pro druhy: první řádek začínající velkým slovem
            # které není label (Holotype:, Locus:, Description:, …)
            words = para.split()
            if (len(words) >= 1
                    and words[0][0].isupper()
                    and not re.match(r"^[A-Z][a-z]{2,}\s*:", para)
                    and words[0].lower() not in TAXON_STOPWORDS
                    and not para.endswith(".")):
                raw = para.split("(")[0].strip()
                return _clean(raw[:150])

    # ── 4. Fallback ───────────────────────────────────────────────────────
    for para in body[:3]:
        if para.strip():
            return para.strip()[:200]
    return ""


def parse_goldset_docx(path: str) -> List[Dict[str, Any]]:
    """
    Parses a PaleoN gold-set DOCX. Supports two marker formats:
      EN: "Record N -- genus TaxonName"
      CS: "Record N TaxonName"  (or "Record N" s diakritikou)

    Vrací seznam records:
    {
      "n": int,
      "taxon_name": str,
      "rank": str,
      "block_text": str,
      "fields": {paleon_field_name: verbatim_text, ...}
    }
    """
    try:
        import docx as _docx_mod
    except ImportError:
        raise RuntimeError(
            "Library python-docx není nainstalována. "
            "Spusťte: pip install python-docx --break-system-packages")

    doc = _docx_mod.Document(path)
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    records = []
    current: Optional[Dict[str, Any]] = None
    current_paras: List[str] = []

    def _finish(rec: Dict, body: List[str]) -> None:
        """Extracts fields from the record body and populates them into rec."""
        fields: Dict[str, str] = {}
        label: Optional[str] = None
        content: List[str] = []

        def _flush():
            if label and content:
                field_key = _GOLDSET_LABEL_TO_FIELD.get(label.lower().rstrip("."), "")
                if field_key:
                    existing = fields.get(field_key, "")
                    addition = " ".join(content).strip()
                    fields[field_key] = (existing + " " + addition).strip() if existing else addition

        for line in body:
            low = line.lower().rstrip(".").strip()
            if low in _GOLDSET_KNOWN_LABELS and len(line.split()) <= 7:
                _flush()
                label = line.rstrip(".").strip()
                content = []
            else:
                m = _GOLDSET_INLINE_RE.match(line)
                if m and len(m.group(1).split()) <= 6:
                    _flush()
                    label = m.group(1).strip()
                    content = [m.group(2).strip()]
                else:
                    content.append(line)
        _flush()

        rec["block_text"] = "\n".join(body)
        rec["fields"] = fields
        # Pokud taxon_name nebyl na řádku markeru, odvodit z těla záznamu
        if not rec.get("taxon_name"):
            rec["taxon_name"] = _infer_taxon_name_from_body(body, rec.get("rank", ""))

    for para in paras:
        m = _GOLDSET_MARKER_RE.match(para)
        if m:
            if current is not None:
                _finish(current, current_paras)
                records.append(current)
            if m.group(1):  # EN format – group(2)=rank (required), group(3)=name (optional)
                rank = m.group(2).strip().title()
                taxon_name = (m.group(3) or "").strip()
                n = int(m.group(1))
            else:           # CS format
                rank = ""
                taxon_name = m.group(5).strip()
                n = int(m.group(4))
            current = {"n": n, "taxon_name": taxon_name, "rank": rank}
            current_paras = []
        elif current is not None:
            current_paras.append(para)

    if current is not None:
        _finish(current, current_paras)
        records.append(current)

    return records


def import_goldset_to_db(path: str, linked_doc_id: Optional[int] = None) -> int:
    """
    Parses a gold-set DOCX and saves it to the DB tables goldset_documents +
    goldset_records. Vrací goldset_doc_id.
    """
    records = parse_goldset_docx(path)
    filename = pathlib.Path(path).name
    now = datetime.now().isoformat(timespec="seconds")
    con = db()
    cur = con.execute(
        "INSERT INTO goldset_documents (filename, path, linked_document_id, created_at) "
        "VALUES (?,?,?,?)", (filename, str(path), linked_doc_id, now))
    gid = cur.lastrowid
    for r in records:
        con.execute(
            "INSERT INTO goldset_records "
            "(goldset_doc_id, record_n, taxon_name, rank, block_text, expected_fields_json, created_at) "
            "VALUES (?,?,?,?,?,?,?)",
            (gid, r["n"], r["taxon_name"], r.get("rank", ""),
             r.get("block_text", ""),
             json.dumps(r.get("fields", {}), ensure_ascii=False),
             now))
    con.commit(); con.close()
    return gid


@dataclass
class GoldEvalResult:
    """Result of comparing one gold-set record with PaleoN detection."""
    gold_n: int
    gold_name: str
    gold_rank: str
    matched_cand_id: Optional[int]
    matched_name: Optional[str]
    detected: bool            # byl taxon vubec detekován?
    name_score: float         # fuzzy shoda jména (0–1)
    block_contains_gold: bool # blok PaleoN obsahuje klíčový text gold setu?
    block_too_short: bool     # blok kratší než gold (missing konec)
    block_too_long: bool      # blok delší než gold o >40 % (pohltil cizí text)
    field_results: Dict[str, bool]  # pole → správně namapováno?


def evaluate_goldset(goldset_doc_id: int, linked_doc_id: int) -> List[GoldEvalResult]:
    """
    Compares gold-set records with PaleoN detection for the given document.

    Taxon match:
    - Normalizovaný name gold setu (bez autoru/roku) se porovná fuzzy
      metodou _norm_search se allmi kandidáty Documentu.
    - Práh: první token (genus) musí sedět + celková fuzzy shoda ≥ 0.55.

    Hodnocení bloku:
    - První odstavec gold set bloku musí být obsažen v PaleoN bloku
      (block_contains_gold). Pokud ne, blok je too_short or nesprávný.
    - Délka PaleoN bloku vs. gold: pokud je PaleoN blok delší o >40 %,
      označíme jako block_too_long (pravděpodobně pohltil sousední taxon).

    Hodnocení fields:
    - Pro každé pole v gold set expected_fields: hledáme, zda PaleoN
      occurrence_fields obsahuje alespoň 30 % textu gold hodnoty
      (substring or fuzzy shoda).
    """
    con = db()
    gold_records = con.execute(
        "SELECT * FROM goldset_records WHERE goldset_doc_id=? ORDER BY record_n",
        (goldset_doc_id,)).fetchall()
    cands = con.execute(
        "SELECT tc.*, d.filename FROM taxon_candidates tc "
        "JOIN documents d ON tc.document_id=d.id "
        "WHERE tc.document_id=? ORDER BY tc.unit_index",
        (linked_doc_id,)).fetchall()
    con.close()

    def _name_score(gold: str, cand: str) -> float:
        """Fuzzy match: compare normalised first words (genus) + the full string."""
        gn = _norm_search(gold)
        cn = _norm_search(cand)
        # První token (rod) musí sedět přesně
        g_first = gn.split()[0] if gn.split() else ""
        c_first = cn.split()[0] if cn.split() else ""
        if g_first and c_first and g_first not in cn and c_first not in gn:
            return 0.0
        return difflib.SequenceMatcher(None, gn[:60], cn[:60]).ratio()

    def _field_ok(gold_val: str, paleon_val: str) -> bool:
        """True if paleon_val contains at least 30 % of the keywords of gold_val."""
        if not gold_val or not paleon_val:
            return False
        gw = set(_norm_search(gold_val).split())
        pw = set(_norm_search(paleon_val).split())
        # Odfiltrovat stop slova (krátká, číslice)
        gw = {w for w in gw if len(w) > 3}
        if not gw:
            return len(_norm_search(gold_val)) > 0 and _norm_search(gold_val)[:30] in _norm_search(paleon_val)
        return len(gw & pw) / len(gw) >= 0.30

    results = []
    for gr in gold_records:
        gold_name = gr["taxon_name"]
        expected = json.loads(gr["expected_fields_json"] or "{}")
        gold_block = gr["block_text"] or ""
        # Klíčová věta gold setu (první neprázdný odstavec těla)
        gold_first_para = next(
            (l.strip() for l in gold_block.split("\n") if len(l.strip()) > 30), "")

        # Najít nejlépe pasujícího kandidáta
        best_cand = None
        best_score = 0.0
        for c in cands:
            sc = _name_score(gold_name, c["taxon_name"] or "")
            if sc > best_score:
                best_score = sc
                best_cand = c

        detected = best_score >= 0.55

        if not detected:
            results.append(GoldEvalResult(
                gold_n=gr["record_n"], gold_name=gold_name, gold_rank=gr["rank"],
                matched_cand_id=None, matched_name=None,
                detected=False, name_score=best_score,
                block_contains_gold=False, block_too_short=False, block_too_long=False,
                field_results={k: False for k in expected}))
            continue

        # Porovnat bloky
        paleon_block = best_cand["block_text"] or ""
        if not paleon_block:
            paleon_block = extract_block_for_candidate(
                best_cand["id"], linked_doc_id, best_cand["page_start"])

        block_contains = (
            _norm_search(gold_first_para[:80]) in _norm_search(paleon_block)
            if gold_first_para else True)
        # Délkové porovnání jen pokud gold blok není prázdný
        gold_len = len(gold_block)
        pal_len  = len(paleon_block)
        block_short = gold_len > 100 and pal_len < gold_len * 0.6
        block_long  = gold_len > 100 and pal_len > gold_len * 1.4

        # Porovnat pole
        con2 = db()
        cand_fields = {
            r["field_name"]: r["field_value"]
            for r in con2.execute(
                "SELECT field_name, field_value FROM occurrence_fields WHERE candidate_id=?",
                (best_cand["id"],)).fetchall()
        }
        con2.close()

        field_results = {}
        for gold_label, gold_val in expected.items():
            paleon_field = _GOLDSET_LABEL_TO_FIELD.get(gold_label.lower().rstrip("."), "")
            if not paleon_field:
                field_results[gold_label] = False
                continue
            paleon_val = cand_fields.get(paleon_field, "")
            field_results[gold_label] = _field_ok(gold_val, paleon_val)

        results.append(GoldEvalResult(
            gold_n=gr["record_n"], gold_name=gold_name, gold_rank=gr["rank"],
            matched_cand_id=best_cand["id"], matched_name=best_cand["taxon_name"],
            detected=True, name_score=round(best_score, 3),
            block_contains_gold=block_contains,
            block_too_short=block_short, block_too_long=block_long,
            field_results=field_results))

    return results


def export_goldset_report_xlsx(results: List[GoldEvalResult],
                                goldset_name: str, doc_name: str) -> bytes:
    """Export benchmark reportu jako XLSX."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        raise RuntimeError("Library openpyxl is not installed.")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gold Set Benchmark"

    GREEN = PatternFill("solid", fgColor="C6EFCE")
    RED   = PatternFill("solid", fgColor="FFC7CE")
    YEL   = PatternFill("solid", fgColor="FFEB9C")

    # Souhrn
    n_total   = len(results)
    n_det     = sum(1 for r in results if r.detected)
    n_fp      = 0  # přesnost nelze spočítat bez znalosti allch kandidátu bez gold
    recall    = n_det / n_total if n_total else 0
    n_short   = sum(1 for r in results if r.block_too_short)
    n_long    = sum(1 for r in results if r.block_too_long)
    n_no_cont = sum(1 for r in results if r.detected and not r.block_contains_gold)

    all_gold_fields = []
    for r in results:
        for k in r.field_results:
            if k not in all_gold_fields:
                all_gold_fields.append(k)
    field_acc = {}
    for f in all_gold_fields:
        vals = [r.field_results[f] for r in results if f in r.field_results]
        field_acc[f] = sum(vals)/len(vals) if vals else 0

    summary = [
        ("Gold set file", goldset_name),
        ("Indexed document", doc_name),
        ("Gold set records", n_total),
        ("Detected", n_det),
        ("Recall", f"{recall:.1%}"),
        ("Boundary too short (block too short)", n_short),
        ("Boundary too long (block too long)", n_long),
        ("Block does not contain the key gold-set text", n_no_cont),
    ]
    for k, v in summary:
        ws.append([k, v])
    ws.append([])
    ws.append(["Field mapping accuracy:"])
    for f, acc in sorted(field_acc.items(), key=lambda x: -x[1]):
        ws.append([f"  {f}", f"{acc:.1%}"])
    ws.append([])

    # Detail table
    headers = ["#", "Gold taxon", "Rank", "Detected",
               "Name match", "Matched candidate",
               "Blok OK", "Too short", "Too long"] + all_gold_fields
    ws.append(headers)
    hdr_row = ws.max_row
    for col, h in enumerate(headers, 1):
        cell = ws.cell(hdr_row, col)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)

    for r in results:
        row = [
            r.gold_n, r.gold_name, r.gold_rank,
            "✅" if r.detected else "❌",
            f"{r.name_score:.2f}" if r.detected else "—",
            r.matched_name or "—",
            "✅" if r.block_contains_gold else ("—" if not r.detected else "❌"),
            "⚠️" if r.block_too_short else "",
            "⚠️" if r.block_too_long else "",
        ]
        for f in all_gold_fields:
            ok = r.field_results.get(f)
            row.append("✅" if ok is True else ("❌" if ok is False else "—"))
        ws.append(row)
        dr = ws.max_row
        # Barva řádku dle detekce
        fill = GREEN if r.detected and r.block_contains_gold else (RED if not r.detected else YEL)
        for col in range(1, 5):
            ws.cell(dr, col).fill = fill
        # Barva fields
        for ci, f in enumerate(all_gold_fields, 10):
            ok = r.field_results.get(f)
            if ok is True:
                ws.cell(dr, ci).fill = GREEN
            elif ok is False:
                ws.cell(dr, ci).fill = RED

    # Šířky sloupcu
    for col_i in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col_i)].width = 18

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()



def export_to_txt(
    rows: List[Dict[str, str]],
    annotate_raw: bool = False,
    references: Optional[List[str]] = None,
) -> bytes:
    """
    Plain text export – records one after another, fields separated by commas.
    Suitable forontrolu a archivaci.

    annotate_raw: pokud True, RAW_TAXONOMIC_BLOCK se vypíše s kategoriemi
    na začátku rozpoznaných odstavcu (viz annotate_raw_block()) — text
    samotný zustává verbatim, jen se přidá štítek.

    references: volitelný seznam vybraných referencí/literatury, který se
    připojí na konec exportu jako samostatná sekce "LITERATURA".
    """
    lines: List[str] = []
    lines.append(f"PaleoN {APP_VERSION} – text export")
    lines.append(f"Datum: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"Number of records: {len(rows)}")
    lines.append("=" * 80)
    SHOW_FIELDS = [
        "RECORD_ID", "TAXON_NAME_VERBATIM", "TAXON_RANK_AS_WRITTEN",
        "AUTHOR", "SOURCE PAGES",
        "SYNONYMY", "TYPE SPECIMENS", "TYPE MATERIAL", "MATERIAL EXAMINED",
        "DIAGNOSIS", "DESCRIPTION", "SIZE", "LOCALITY", "STRATIGRAPHY",
        "OCCURRENCE", "REMARKS", "ETYMOLOGY", "FIGURES", "REFERENCE",
        "EXTRACTION_NOTES", "RAW_TAXONOMIC_BLOCK",
    ]
    for i, row in enumerate(rows, 1):
        lines.append(f"\n{'=' * 80}")
        lines.append(f"RECORD {i}  –  {row.get('TAXON_NAME_VERBATIM', '?')}")
        lines.append("=" * 80)
        for field in SHOW_FIELDS:
            val = row.get(field, "Not provided")
            if val and val != "Not provided":
                lines.append(f"\n{field}:")
                if field == "RAW_TAXONOMIC_BLOCK" and annotate_raw:
                    lines.append(annotate_raw_block(val))
                else:
                    lines.append(val)
        lines.append("")

    if references:
        lines.append(f"\n{'=' * 80}")
        lines.append(f"LITERATURA ({len(references)} items)")
        lines.append("=" * 80)
        for i, ref in enumerate(references, 1):
            lines.append(f"\n[{i}] {ref}")

    return "\n".join(lines).encode("utf-8")


def export_to_pdf(rows: List[Dict[str, str]]) -> bytes:
    """
    PDF export using fpdf2. If fpdf2 is not available, returns HTML as a fallback.
    Usesžívá latin-1 safe encoding pro core fonty (Helvetica).
    """
    if not HAS_FPDF:
        return _export_to_html_print(rows).encode("utf-8")

    SHOW = [
        ("TAXON_RANK_AS_WRITTEN", "Rank"),
        ("AUTHOR",                "Author"),
        ("SOURCE PAGES",          "Pages"),
        ("SYNONYMY",              "Synonymy"),
        ("TYPE SPECIMENS",        "Types"),
        ("MATERIAL EXAMINED",     "Material"),
        ("DIAGNOSIS",             "Diagnosis"),
        ("DESCRIPTION",           "Description"),
        ("SIZE",                  "Dimensions"),
        ("LOCALITY",              "Locality"),
        ("STRATIGRAPHY",          "Stratigraphy"),
        ("REMARKS",               "Notes"),
        ("ETYMOLOGY",             "Etymology"),
        ("OCCURRENCE",            "Occurrence"),
        ("EXTRACTION_NOTES",      "Extraction notes"),
    ]

    def _safe(text: str, maxlen: int = 800) -> str:
        """Latin-1 safe string pro fpdf2 core fonty."""
        return text[:maxlen].encode("latin-1", errors="replace").decode("latin-1")

    pdf = _FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # Titulek
    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(pdf.epw, 9, f"PaleoN {APP_VERSION}  -  Taxonomic Extraction", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 8)
    pdf.cell(pdf.epw, 6,
             f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}   Records: {len(rows)}",
             new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    for i, row in enumerate(rows, 1):
        taxon = _safe(row.get("TAXON_NAME_VERBATIM", "?"), 120)

        # Barevný nadpis taxonu
        pdf.set_fill_color(31, 78, 121)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(pdf.epw, 7, f"{i}.  {taxon}", fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
        pdf.ln(1)

        # Datová pole
        pdf.set_font("Helvetica", "", 9)
        for field, label in SHOW:
            val = row.get(field, "")
            if not val or val == NOT_PROVIDED:
                continue
            val_safe = _safe(val, 600)
            label_safe = _safe(label)
            pdf.set_font("Helvetica", "B", 8)
            pdf.cell(35, 4, label_safe + ":", new_x="RIGHT", new_y="TOP")
            pdf.set_font("Helvetica", "", 8)
            # multi_cell musí začínat na správné x pozici
            x_after_label = pdf.get_x()
            pdf.multi_cell(pdf.epw - 35, 4, val_safe)
            pdf.ln(0.5)

        pdf.ln(3)
        if i < len(rows):
            pdf.set_draw_color(180, 180, 180)
            pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
            pdf.ln(3)

    return bytes(pdf.output())


def _export_to_html_print(rows: List[Dict[str, str]]) -> str:
    """HTML fallback for browser printing to PDF (fpdf2 not available)."""
    SHOW = [
        "TAXON_NAME_VERBATIM","AUTHOR","SYNONYMY","TYPE SPECIMENS",
        "TYPE MATERIAL","MATERIAL EXAMINED","DIAGNOSIS","DESCRIPTION",
        "SIZE","LOCALITY","STRATIGRAPHY","REMARKS","OCCURRENCE",
        "ETYMOLOGY","FIGURES","REFERENCE","EXTRACTION_NOTES",
    ]
    html = ["""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
  body{font-family:Georgia,serif;font-size:10pt;margin:2cm}
  h1{font-size:13pt;border-bottom:2px solid #1F4E79;color:#1F4E79;margin-top:1.5em}
  .label{font-weight:bold;color:#555;font-size:9pt;margin-top:0.5em}
  .val{margin:0 0 0.3em 0;white-space:pre-wrap}
  @media print{h1{page-break-before:always}h1:first-child{page-break-before:avoid}}
</style></head><body>"""]
    html.append(f"<p><strong>PaleoN {APP_VERSION}</strong> – {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                f" – {len(rows)} records</p>")
    for i, row in enumerate(rows, 1):
        taxon = row.get("TAXON_NAME_VERBATIM", "?")
        html.append(f"<h1>{i}. {taxon}</h1>")
        for field in SHOW:
            val = row.get(field, "")
            if val and val != "Not provided":
                html.append(f'<div class="label">{field}</div>')
                import html as _html
                html.append(f'<p class="val">{_html.escape(val[:1200])}</p>')
    html.append("</body></html>")
    return "\n".join(html)


# ══════════════════════════════════════════════════════════════════════════════
# NASTAVENÍ – persistentní saveení + session state
# ══════════════════════════════════════════════════════════════════════════════

def save_settings_to_disc(s: Dict[str, Any]) -> None:
    """Saves settings (excluding prompts) to a JSON file."""
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    # Prompty ukládáme separátně; tady ukládáme jen numeriku/bool/string
    exclude = {"llm_validation_prompt", "llm_boundary_prompt", "llm_field_prompt"}
    data = {k: v for k, v in s.items() if k not in exclude}
    SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_settings_from_disc() -> Dict[str, Any]:
    """Loads settings from disc; fills in missing keys from DEFAULT_SETTINGS."""
    base = DEFAULT_SETTINGS.copy()
    if SETTINGS_FILE.exists():
        try:
            loaded = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
            base.update(loaded)
        except Exception:
            pass
    # Vždy načti prompty zvlášť
    base.update(load_prompts())
    return base


def save_prompts(s: Dict[str, Any]) -> None:
    """Saves LLM prompts to a separate JSON file."""
    BASE_DIR.mkdir(parents=True, exist_ok=True)
    data = {
        "llm_validation_prompt": s.get("llm_validation_prompt", LLM_VALIDATION_PROMPT),
        "llm_boundary_prompt":   s.get("llm_boundary_prompt",   LLM_BOUNDARY_PROMPT),
        "llm_field_prompt":      s.get("llm_field_prompt",      LLM_FIELD_PROMPT),
    }
    PROMPTS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_prompts() -> Dict[str, str]:
    """Loads saved prompts or returns defaults."""
    if PROMPTS_FILE.exists():
        try:
            d = json.loads(PROMPTS_FILE.read_text(encoding="utf-8"))
            return {
                "llm_validation_prompt": d.get("llm_validation_prompt", LLM_VALIDATION_PROMPT),
                "llm_boundary_prompt":   d.get("llm_boundary_prompt",   LLM_BOUNDARY_PROMPT),
                "llm_field_prompt":      d.get("llm_field_prompt",      LLM_FIELD_PROMPT),
            }
        except Exception:
            pass
    return {
        "llm_validation_prompt": LLM_VALIDATION_PROMPT,
        "llm_boundary_prompt":   LLM_BOUNDARY_PROMPT,
        "llm_field_prompt":      LLM_FIELD_PROMPT,
    }


def get_settings() -> Dict[str, Any]:
    if "paleon_settings" not in st.session_state:
        st.session_state["paleon_settings"] = load_settings_from_disc()
    return st.session_state["paleon_settings"]


# ══════════════════════════════════════════════════════════════════════════════
# UI KOMPONENTY
# ══════════════════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════════════════
# CSS + GUI UTILITY
# ══════════════════════════════════════════════════════════════════════════════

_CSS = """
<style>
/* ── Compact metric ── */
[data-testid="stMetricValue"] { font-size: 1.45rem !important; font-weight: 700; }
[data-testid="stMetricLabel"] { font-size: 0.75rem !important; color: #666; }

/* ── Status badge ── */
.badge {
    display: inline-block; padding: 2px 9px; border-radius: 10px;
    font-size: 0.72rem; font-weight: 600; letter-spacing: 0.03em;
}
.b-approved { background:#166534; color:#fff; }
.b-pending  { background:#92400e; color:#fff; }
.b-rejected { background:#991b1b; color:#fff; }
.b-low      { background:#1e40af; color:#fff; }
.b-review   { background:#5b21b6; color:#fff; }

/* ── Fill bar ── */
.fill-wrap { background:#e5e7eb; border-radius:4px; height:6px; margin:3px 0; }
.fill-bar  { height:6px; border-radius:4px; background:linear-gradient(90deg,#3b82f6,#10b981); }

/* ── Editor header ── */
.taxon-header {
    background: #1e293b;
    color: #f1f5f9; padding: 10px 16px; border-radius: 8px;
    margin-bottom: 12px; font-size: 1.05rem; font-weight: 700;
}

/* ── Compact data_editor ── */
[data-testid="stDataFrameContainer"] { font-size: 0.82rem; }

/* ── Tighter expanders ── */
.streamlit-expanderHeader { padding: 0.35rem 0.7rem !important; font-size: 0.88rem !important; }

/* ── Sidebar divider ── */
[data-testid="stSidebar"] hr { margin: 8px 0; }

/* ── Tab content padding ── */
[data-testid="stTabContent"] { padding-top: 12px; }

/* ── Progress text ── */
.prog-label { font-size: 0.78rem; color: #555; margin-bottom: 2px; }

/* ── Compact button row ── */
.btn-row button { margin: 0 2px !important; }

/* ── Warning box ── */
.warn-box {
    background: #fef3c7; border: 1px solid #f59e0b;
    border-radius: 6px; padding: 8px 12px;
    font-size: 0.83rem; color: #78350f;
}

/* ══════════════════════════════════════════════════════════════════════
   COMPACT LAYOUT — less whitespace across the entire application
   ══════════════════════════════════════════════════════════════════════ */

/* Main content block — sufficient top padding so content does not slide under the fixed Streamlit toolbar */
.block-container { padding-top: 3.5rem !important; padding-bottom: 1.6rem !important; }

/* Vertical spacing between widgets in general */
[data-testid="stVerticalBlock"] > div { gap: 0.35rem !important; }
div[data-testid="stVerticalBlockBorderWrapper"] { gap: 0.3rem !important; }

/* st.divider() — thinner spacing */
hr { margin: 0.5rem 0 !important; }

/* Headings h1-h3 — less top/bottom padding */
h1, h2, h3 { margin-top: 0.3rem !important; margin-bottom: 0.3rem !important; padding-top: 0 !important; }

/* Paragraphs / markdown blocks */
[data-testid="stMarkdownContainer"] p { margin-bottom: 0.25rem !important; }

/* Expander inner padding */
[data-testid="stExpander"] details summary { padding: 0.3rem 0.6rem !important; }
[data-testid="stExpander"] .streamlit-expanderContent { padding: 0.4rem 0.7rem !important; }

/* Columns — smaller gap between them */
[data-testid="stHorizontalBlock"] { gap: 0.5rem !important; }

/* Widgets — less bottom spacing */
[data-testid="stWidgetLabel"] { margin-bottom: 0.1rem !important; padding-bottom: 0 !important; }
.stTextInput, .stSelectbox, .stTextArea, .stNumberInput, .stSlider,
.stCheckbox, .stToggle, .stMultiSelect, .stDateInput { margin-bottom: 0.15rem !important; }

/* Buttons — smaller vertical padding */
.stButton button, .stDownloadButton button { padding: 0.3rem 0.8rem !important; }

/* Metrics — less surrounding space */
[data-testid="stMetric"] { padding: 0.2rem 0 !important; }

/* Caption text — less space above/below */
[data-testid="stCaptionContainer"] { margin-top: 0 !important; margin-bottom: 0.2rem !important; }

/* ── Tabs — horizontal scroll in case of overflow ────────────────────────
   We use ONLY .stTabs scoped selectors (never [role="tablist"] /
   [role="tab"] without scope — they are too generic and interfere with the
   internal Streamlit toolbar, causing visual clipping of the tab bar at the top). */
.stTabs [data-baseweb="tab-list"] {
    overflow-x: auto !important;
    flex-wrap: nowrap !important;
    gap: 6px !important;          /* gap between tabs */
    overflow-y: visible !important;
}
.stTabs [data-baseweb="tab"] {
    white-space: nowrap !important;
    min-width: 0 !important;
    flex-shrink: 1 !important;
}

/* Tabs — menší padding nad obsahem */
.stTabs [data-baseweb="tab-panel"] { padding-top: 0.6rem !important; }

/* Data editor / dataframe — less surrounding space */
[data-testid="stDataFrame"], [data-testid="stDataEditor"] { margin-bottom: 0.3rem !important; }

/* Main tabs — pipe as separator + bold text */
[data-baseweb="tab-list"] [data-baseweb="tab"]:not(:first-child)::before {
    content: "|";
    padding-right: 0.55em;
    color: rgba(128, 128, 128, 0.55);
    font-weight: 300;
    font-size: 1.05em;
    pointer-events: none;
}
[data-baseweb="tab"] p,
[data-baseweb="tab"] span {
    font-weight: 700 !important;
}
</style>
"""

# ── Tmavý režim — forcovaný CSS override (Streamlit vlastní theme picker
#    neřešíme, přepisujeme přímo barvy přes data-testid selektory) ──────────
_CSS_DARK = """
<style>
:root {
    --pn-bg: #0e1117;
    --pn-bg-secondary: #1a1d24;
    --pn-bg-tertiary: #262730;
    --pn-text: #e6e6e6;
    --pn-text-muted: #9aa0aa;
    --pn-border: #363945;
    --pn-accent: #3b82f6;
}

.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"],
[data-testid="stBottomBlockContainer"] {
    background-color: var(--pn-bg) !important;
    color: var(--pn-text) !important;
}
[data-testid="stHeader"] { background-color: transparent !important; }

[data-testid="stSidebar"] {
    background-color: var(--pn-bg-secondary) !important;
    border-right: 1px solid var(--pn-border) !important;
}
[data-testid="stSidebar"] * { color: var(--pn-text) !important; }

body, .stMarkdown, p, span, label, li, div { color: var(--pn-text); }
h1, h2, h3, h4, h5, h6 { color: var(--pn-text) !important; }
[data-testid="stCaptionContainer"] { color: var(--pn-text-muted) !important; }

/* Input elements */
.stTextInput input, .stTextArea textarea, .stNumberInput input,
.stSelectbox [data-baseweb="select"] > div,
.stMultiSelect [data-baseweb="select"] > div,
.stDateInput input {
    background-color: var(--pn-bg-tertiary) !important;
    color: var(--pn-text) !important;
    border-color: var(--pn-border) !important;
}
[data-baseweb="popover"] { background-color: var(--pn-bg-tertiary) !important; }
[role="listbox"] { background-color: var(--pn-bg-tertiary) !important; }
[role="option"] { color: var(--pn-text) !important; }

/* Buttons */
.stButton button, .stDownloadButton button {
    background-color: var(--pn-bg-tertiary) !important;
    color: var(--pn-text) !important;
    border: 1px solid var(--pn-border) !important;
}
.stButton button:hover, .stDownloadButton button:hover {
    border-color: var(--pn-accent) !important;
    color: var(--pn-accent) !important;
}
.stButton button[kind="primary"] {
    background-color: var(--pn-accent) !important;
    color: #fff !important;
    border: none !important;
}

/* Expanders */
[data-testid="stExpander"] {
    background-color: var(--pn-bg-secondary) !important;
    border: 1px solid var(--pn-border) !important;
    border-radius: 6px;
}
.streamlit-expanderHeader, [data-testid="stExpander"] summary { color: var(--pn-text) !important; }

/* Tabs */
.stTabs [data-baseweb="tab-list"] {
    background-color: var(--pn-bg-secondary) !important;
    border-bottom: 1px solid var(--pn-border) !important;
}
.stTabs [data-baseweb="tab"] { color: var(--pn-text-muted) !important; }
.stTabs [aria-selected="true"] { color: var(--pn-accent) !important; }

/* Tables / data editor */
[data-testid="stDataFrame"], [data-testid="stDataEditor"] {
    background-color: var(--pn-bg-secondary) !important;
    border: 1px solid var(--pn-border) !important;
}

/* Metrics */
[data-testid="stMetric"] {
    background-color: var(--pn-bg-secondary) !important;
    border-radius: 6px; padding: 6px 10px !important;
}
[data-testid="stMetricLabel"] { color: var(--pn-text-muted) !important; }
[data-testid="stMetricValue"] { color: var(--pn-text) !important; }

/* Dividers */
hr { border-color: var(--pn-border) !important; }

/* Notifications / alerts */
[data-testid="stAlert"] {
    background-color: var(--pn-bg-tertiary) !important;
    color: var(--pn-text) !important;
}

/* Code */
code, pre { background-color: var(--pn-bg-tertiary) !important; color: #e6e6e6 !important; }

/* Custom warn-box — override to readable colours for dark background */
.warn-box { background:#3a2f10 !important; border-color:#f59e0b !important; color:#fde68a !important; }

/* Taxon-header gradient kept — already dark and high-contrast */

/* Toggle/checkbox labels */
[data-testid="stWidgetLabel"] p { color: var(--pn-text) !important; }

/* ── File uploader ─────────────────────────────────────────────────────────*/
[data-testid="stFileUploader"] {
    background-color: var(--pn-bg-secondary) !important;
    border-radius: 8px !important;
    padding: 4px !important;
}
[data-testid="stFileUploaderDropzone"] {
    background-color: var(--pn-bg-tertiary) !important;
    border: 1.5px dashed var(--pn-border) !important;
    border-radius: 6px !important;
}
[data-testid="stFileUploaderDropzone"] small,
[data-testid="stFileUploaderDropzone"] p,
[data-testid="stFileUploaderDropzone"] span:not([data-testid]) {
    color: var(--pn-text-muted) !important;
}
[data-testid="stFileUploaderDropzone"] > section > button {
    background-color: var(--pn-bg-secondary) !important;
    color: var(--pn-text) !important;
    border: 1px solid var(--pn-border) !important;
}
[data-testid="stFileUploader"] [data-testid="stFileUploaderFileName"] {
    color: var(--pn-text) !important;
}
[data-testid="stFileUploader"] [data-testid="stFileUploaderDeleteBtn"] button {
    background: transparent !important;
    color: var(--pn-text-muted) !important;
}

/* ── DataFrame / data editor inner dark styling ─────────────────────────── */
/* Container wrappers */
[data-testid="stDataFrameContainer"] {
    background-color: var(--pn-bg-secondary) !important;
    border-radius: 6px !important;
    overflow: hidden !important;
}
/* Glide Data Grid main wrappers (canvas-based — styles the chrome around canvas) */
.dvn-stack,
[class*="dvn-"],
[class*="glideDataGrid"] {
    background-color: var(--pn-bg-secondary) !important;
    color: var(--pn-text) !important;
}
/* Column headers row (non-canvas layer overlay) */
[data-testid="stDataFrameContainer"] > div > div,
[data-testid="stDataEditor"] > div > div {
    background-color: var(--pn-bg-secondary) !important;
}
/* DataEditor toolbar (add row / search bar) */
[data-testid="stDataEditor"] [data-testid="stDataEditorHeader"] {
    background-color: var(--pn-bg-secondary) !important;
    border-bottom: 1px solid var(--pn-border) !important;
}
[data-testid="stDataEditor"] [data-testid="stDataEditorHeader"] * {
    color: var(--pn-text) !important;
    background-color: var(--pn-bg-secondary) !important;
}
/* DataEditor cell editing input */
[data-testid="stDataEditor"] input,
[data-testid="stDataEditor"] textarea {
    background-color: var(--pn-bg-tertiary) !important;
    color: var(--pn-text) !important;
    border: 1px solid var(--pn-border) !important;
}
/* Scrollbar in tables */
[data-testid="stDataFrameContainer"] ::-webkit-scrollbar,
[data-testid="stDataEditor"] ::-webkit-scrollbar {
    background: var(--pn-bg-tertiary) !important;
    width: 8px; height: 8px;
}
[data-testid="stDataFrameContainer"] ::-webkit-scrollbar-thumb,
[data-testid="stDataEditor"] ::-webkit-scrollbar-thumb {
    background: var(--pn-border) !important;
    border-radius: 4px;
}
/* pn-term-table custom HTML tables */
.pn-term-table { background-color: var(--pn-bg-secondary) !important; color: var(--pn-text) !important; }
.pn-term-table th { color: var(--pn-text-muted) !important; border-bottom-color: var(--pn-border) !important; }
.pn-term-table td { color: var(--pn-text) !important; border-bottom-color: rgba(100,116,139,0.2) !important; }
.pn-result-name { color: var(--pn-text) !important; }
.pn-result-meta { color: var(--pn-text-muted) !important; }
.pn-cat-chip { border-color: var(--pn-border) !important; color: var(--pn-text) !important; }
</style>
"""


def inject_css(theme: str = "light") -> None:
    st.markdown(_CSS, unsafe_allow_html=True)
    if theme == "dark":
        st.markdown(_CSS_DARK, unsafe_allow_html=True)


def _badge_html(status: str) -> str:
    cls = {"approved":"b-approved","pending":"b-pending","rejected":"b-rejected",
           "low_confidence":"b-low","needs_review":"b-review"}.get(status,"b-pending")
    label = {"approved":"✅ approved","pending":"⏳ pending","rejected":"❌ rejected",
              "low_confidence":"🔵 low-conf","needs_review":"🔍 review"}.get(status, status)
    return f'<span class="badge {cls}">{label}</span>'


def _fill_bar_html(filled: int, total: int) -> str:
    pct = int(filled / total * 100) if total else 0
    color = "#ef4444" if pct < 30 else "#f59e0b" if pct < 70 else "#10b981"
    return (f'<div class="fill-wrap"><div class="fill-bar" '
            f'style="width:{pct}%;background:{color};"></div></div>'
            f'<span style="font-size:0.75rem;color:#555">{filled}/{total} fields</span>')


def _conf_badge(score: float) -> str:
    if score >= 0.80: return "🟢"
    if score >= 0.60: return "🟡"
    return "🔴"


def _status_badge(status: str) -> str:
    return {"approved":"✅","rejected":"❌","pending":"⏳",
            "low_confidence":"🔵","needs_review":"🔍"}.get(status,"❓")


def _norm_search(text: str) -> str:
    """
    Normalises a string for diacritic-aware comparison:
    converts to lower case and stripsraní diakritiku (NFD → zachová jen ASCII +
    základní latinku bez kombinačních znaku). Díky tomu 'gracilitheca'
    najde 'Gracilithéca', 'diakritika' najde 'Diakritika', a zároveň
    hledání s diakritikou ('Šatný') najde i variantu bez ('Satny').

    Použito v živém filtering v Editoru i Dossierech pro správnou podporu
    českých, německých, francouzských a ruských (Cyrilici) názvu taxa.
    """
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in nfkd if not unicodedata.combining(ch)).lower()


def _count_filled_fields(candidate_id: int) -> Tuple[int, int]:
    """Returns (filled, total) for output section fields."""
    con = db()
    filled = con.execute(
        "SELECT COUNT(*) FROM occurrence_fields WHERE candidate_id=? AND field_value!=''",
        (candidate_id,)
    ).fetchone()[0]
    con.close()
    return filled, len(SECTION_FIELDS)


# ══════════════════════════════════════════════════════════════════════════════
# KLÁVESOVÉ ZKRATKY  (A = schválit, R = odmítnout)
# ══════════════════════════════════════════════════════════════════════════════

def _inject_keyboard_shortcuts(scope_id: str, approve_marker: str, reject_marker: str) -> None:
    """
    Injects an invisible JS listener that on key press A/R (outside text
    fields and modmo modifikátory Ctrl/Cmd/Alt) klikne na VIDITELNÉ tlačítko,
    jehož popisek obsahuje zadaný marker. Marker je zároveň zobrazen přímo
    v popisku tlačítka (e.g. "✅ Approve (A)"), takže slouží i jako
    nápověda pro uživatele.

    scope_id odlišuje listenery mezi taby (Review vs. Editor) — bez toho
    by se při každém rerunu mohly hromadit duplicitní listenery.
    Viditelnost tlačítka (offsetParent !== null) zajišťuje, že se klikne
    jen na tlačítko v aktuálně aktivním/otevřeném panelu, ne na skryté
    duplicity v neaktivních tabech or zavřených expanderech.
    """
    import streamlit.components.v1 as components
    flag = f"__paleon_kbd_{scope_id}"
    components.html(f"""
    <script>
    (function() {{
        const doc = window.parent.document;
        if (doc.{flag}) return;
        doc.{flag} = true;
        doc.addEventListener('keydown', function(e) {{
            const tag = (e.target.tagName || '').toLowerCase();
            if (tag === 'input' || tag === 'textarea' || e.target.isContentEditable) return;
            if (e.metaKey || e.ctrlKey || e.altKey || e.shiftKey) return;
            const key = e.key.toLowerCase();
            let marker = null;
            if (key === 'a') marker = {json.dumps(approve_marker)};
            else if (key === 'r') marker = {json.dumps(reject_marker)};
            if (!marker) return;
            const buttons = doc.querySelectorAll('button');
            for (const b of buttons) {{
                if (b.offsetParent === null) continue;  // skrytý tab/expander
                if (b.innerText && b.innerText.includes(marker)) {{
                    b.click();
                    e.preventDefault();
                    break;
                }}
            }}
        }});
    }})();
    </script>
    """, height=0, width=0)


# ══════════════════════════════════════════════════════════════════════════════
# NÁHLED SKUTEČNÉ PDF STRÁNKY  (Library → Prohlížeč pages)
# ══════════════════════════════════════════════════════════════════════════════

def _render_pdf_page_image_impl(path_str: str, page_number: int, mtime: float,
                                 zoom: float = 1.6) -> Optional[bytes]:
    """
    Renders one PDF page (1-indexed, matching page_number
    in the DB) as a PNG obrázek pro vizuální ověření OCR chyb. Vrací None,
    pokud PyMuPDF není dostupný, file není PDF or stránka neexists.

    `mtime` je součástí cache klíče jen proto, aby se cache invalidovala,
    pokud by se file na discu někdy přepsal (re-upload stejného jména).
    """
    if not HAS_FITZ:
        return None
    path = pathlib.Path(path_str)
    if path.suffix.lower() != ".pdf" or not path.exists():
        return None
    try:
        pdf = fitz.open(str(path))
        idx = page_number - 1
        if idx < 0 or idx >= len(pdf):
            pdf.close()
            return None
        page = pdf[idx]
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        png_bytes = pix.tobytes("png")
        pdf.close()
        return png_bytes
    except Exception:
        return None


# Cachujeme render jen pokud běžíme pod Streamlitem (st.cache_data vyžaduje
# aktivní runtime) — mimo Streamlit (e.g. při syntax-checku) použijeme
# funkci bez cache.
if st is not None:
    _render_pdf_page_image = st.cache_data(show_spinner=False)(_render_pdf_page_image_impl)
else:
    _render_pdf_page_image = _render_pdf_page_image_impl


def _show_pdf_page_inline(
    document_id: int,
    page_number:  int,
    *,
    label: str  = "",
    zoom:  float = 1.6,
    key:   str  = "",
) -> None:
    """
    Displays one PDF page directly in the application as an image (PyMuPDF).
    Safe alternativa k nefunkčním file:// hyperlinkum — prohlížeče
    blokují navigaci z http://localhost na file://.

    Parametry:
        document_id  – ID z tabulky documents
        page_number  – číslo stránky (1-indexováno)
        label        – nadpis nad náhledem (prázdné = automatický)
        zoom         – zoom factor (1.6 = dobré čitelné, 2.0 = větší detail)
        key          – Streamlit klíč pro cache (musí být unikátní v Contextu)
    """
    if not HAS_FITZ:
        st.caption(t("pdf_preview_requires_pymupdf"))
        return

    # Cache path lookupu v session_state — zamezí opakovaným DB dotazum
    _cache_key = f"pdf_doc_info_{document_id}"
    if _cache_key not in st.session_state:
        con = db()
        _dr = con.execute(
            "SELECT path, filename FROM documents WHERE id=?", (document_id,)
        ).fetchone()
        con.close()
        st.session_state[_cache_key] = (
            (str(_dr["path"]), str(_dr["filename"])) if _dr else None)
    _doc_info = st.session_state[_cache_key]
    if not _doc_info:
        st.caption(t("file_path_not_in_db"))
        return
    doc_path = pathlib.Path(_doc_info[0])
    doc_row = {"path": _doc_info[0], "filename": _doc_info[1]}
    if not doc_row or not doc_row["path"]:
        st.caption(t("file_path_not_in_db"))
        return

    doc_path = pathlib.Path(doc_row["path"])
    if not doc_path.exists():
        # Zkusit relativní cesta z uploads
        alt = _upath("../uploads") / doc_path.name
        if alt.exists():
            doc_path = alt
        else:
            st.caption(tt(f"ℹ️ File not found: `{doc_path.name}`", f"ℹ️ File not found: `{doc_path.name}`"))
            return

    if doc_path.suffix.lower() != ".pdf":
        st.caption(t("preview_pdf_only"))
        return

    mtime = doc_path.stat().st_mtime
    png = _render_pdf_page_image(str(doc_path), page_number, mtime, zoom)
    if png is None:
        st.caption(tt(f"ℹ️ Page {page_number} could not be rendered.", f"ℹ️ Page {page_number} could not be rendered."))
        return

    cap = label or f"📄 **{doc_row['filename']}** — str. {page_number}"
    st.caption(cap)
    st.image(png, use_container_width=True)




# ══════════════════════════════════════════════════════════════════════════════
# FUZZY DETEKCE DUPLICITNÍCH TAXONŮ NAPŘÍČ DocumentY
# ══════════════════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════════════════
# BOD 3: FTS5 — fulltext search (263× rychlejší než LIKE %query%)
# ══════════════════════════════════════════════════════════════════════════════

def _fts_sanitize(query: str) -> str:
    """Converts a user query to a safe FTS5 expression."""
    q = query.strip()
    if not q:
        return ""
    has_wildcard = q.endswith("*")
    base = q.rstrip("*").strip()
    for ch in ['"', '(', ')', ':', '-', '+', '^', '!', '@', '~']:
        base = base.replace(ch, " ")
    base = " ".join(base.split())
    if not base:
        return ""
    if has_wildcard:
        return f'"{base}"*' if " " in base else f"{base}*"
    return f'"{base}"' if " " in base else base


def fts_search(query: str, limit: int = 500) -> List[int]:
    """
    Full-text search via FTS5 index — 100–1000× faster than LIKE.
    Returns a list of canddidate_id seřazených podle relevance (rank).
    Při selhání (FTS5 nedostupné) vrátí prázdný seznam → caller použije LIKE.
    """
    fts_q = _fts_sanitize(query)
    if not fts_q:
        return []
    try:
        con = db()
        rows = con.execute(
            "SELECT rowid FROM fts_candidates WHERE fts_candidates MATCH ? "
            "ORDER BY rank LIMIT ?",
            (fts_q, limit),
        ).fetchall()
        con.close()
        return [r[0] for r in rows]
    except Exception:
        return []


def _fts_update_candidate(candidate_id: int) -> None:
    """
    Updates the FTS5 index for a single candidate.
    Called automatically after save_fields().
    """
    try:
        con = db()
        cand = con.execute(
            "SELECT taxon_name, block_text FROM taxon_candidates WHERE id=?",
            (candidate_id,),
        ).fetchone()
        if not cand:
            con.close()
            return
        fields = con.execute(
            "SELECT field_name, field_value FROM occurrence_fields "
            "WHERE candidate_id=? AND field_value != ''",
            (candidate_id,),
        ).fetchall()
        all_fields_text = " ".join(
            f"{r['field_name']}: {r['field_value']}"
            for r in fields
            if r["field_value"] and r["field_value"] != NOT_PROVIDED
        )
        con.execute("DELETE FROM fts_candidates WHERE rowid=?", (candidate_id,))
        con.execute(
            "INSERT INTO fts_candidates(rowid, taxon_name, block_text, all_fields) "
            "VALUES (?, ?, ?, ?)",
            (candidate_id,
             (cand["taxon_name"] or "").strip(),
             (cand["block_text"] or "")[:50_000],
             all_fields_text[:20_000]),
        )
        con.commit()
        con.close()
    except Exception as _fts_exc:
        logging.debug(f"FTS update failed for {candidate_id}: {_fts_exc}")


def fts_rebuild_all() -> int:
    """
    Full rebuild of the FTS5 index from all data in the DB.
    Called manually from the UI (Settings → Data) or after bulk import.
    Vrací count indexovaných kandidátu.
    """
    try:
        con = db()
        con.execute("DELETE FROM fts_candidates")
        con.commit()
        cands = con.execute(
            "SELECT id, taxon_name, block_text FROM taxon_candidates "
            "WHERE TRIM(taxon_name) != ''"
        ).fetchall()
        n = 0
        for cand in cands:
            fields = con.execute(
                "SELECT field_name, field_value FROM occurrence_fields "
                "WHERE candidate_id=? AND field_value != ''",
                (cand["id"],),
            ).fetchall()
            all_fields_text = " ".join(
                f"{r['field_name']}: {r['field_value']}"
                for r in fields
                if r["field_value"] and r["field_value"] != NOT_PROVIDED
            )
            con.execute(
                "INSERT INTO fts_candidates(rowid, taxon_name, block_text, all_fields) "
                "VALUES (?, ?, ?, ?)",
                (cand["id"],
                 (cand["taxon_name"] or "").strip(),
                 (cand["block_text"] or "")[:50_000],
                 all_fields_text[:20_000]),
            )
            n += 1
        con.commit()
        con.close()
        return n
    except Exception as exc:
        logging.warning(f"FTS rebuild failed: {exc}")
        return 0



def find_fuzzy_taxon_duplicates(threshold: float = 0.86) -> List[Dict[str, Any]]:
    """
    Finds candidate pairs with a similar (but not identical) taxon name —
    typically the samený taxon s drobně odlišným pravopisem napříč Documenty
    (překlep, OCR chyba, jiná transliterace). Přesné shody (case-insensitive)
    se přeskakují — to jsou očekávané opakované výskyty, ne problém.

    Candidates bucketují se podle prvního písmene názvu (case-insensitive),
    aby se omezil count porovnání u větších knihoven.
    """
    con = db()
    rows = con.execute("""
        SELECT tc.id, tc.taxon_name, tc.status, tc.document_id, tc.page_start,
               d.filename
        FROM taxon_candidates tc
        JOIN documents d ON tc.document_id = d.id
        WHERE tc.status != 'rejected' AND TRIM(tc.taxon_name) != ''
    """).fetchall()
    con.close()

    buckets: Dict[str, list] = {}
    for r in rows:
        name = (r["taxon_name"] or "").strip()
        if not name:
            continue
        key = name[0].lower()
        buckets.setdefault(key, []).append(r)

    pairs = []
    for items in buckets.values():
        n = len(items)
        for i in range(n):
            for j in range(i + 1, n):
                a, b = items[i], items[j]
                na, nb = (a["taxon_name"] or "").strip(), (b["taxon_name"] or "").strip()
                if na.lower() == nb.lower():
                    continue
                ratio = difflib.SequenceMatcher(None, na.lower(), nb.lower()).ratio()
                if ratio >= threshold:
                    pairs.append({
                        "id_a": a["id"], "name_a": na, "doc_a": a["filename"],
                        "page_a": a["page_start"], "status_a": a["status"],
                        "id_b": b["id"], "name_b": nb, "doc_b": b["filename"],
                        "page_b": b["page_start"], "status_b": b["status"],
                        "podobnost": round(ratio, 3),
                    })
    pairs.sort(key=lambda p: -p["podobnost"])
    return pairs


def merge_candidate_fields(
    target_id: int,
    source_id: int,
    overwrite: bool = False,
) -> int:
    """
    Copies filled fields from `source_id` to `target_id`.
    If `overwrite=False`, skeskočí pole která target již má.
    Vrací count zkopírovaných fields.

    Použití: target je hlavní záznam, source je duplicita z jiného Documentu.
    """
    src_fields = get_candidate_fields(source_id)
    tgt_fields = get_candidate_fields(target_id)
    to_copy: Dict[str, str] = {}
    for fname, fval in src_fields.items():
        if not fval or fval == NOT_PROVIDED:
            continue
        if overwrite or not tgt_fields.get(fname):
            to_copy[fname] = fval
    if to_copy:
        save_fields(target_id, to_copy, method="merged")
    return len(to_copy)


def link_as_synonym(candidate_id: int, related_id: int) -> None:
    """
    Links two records as synonyms:
    - inserts related_id into RELATED_RECORD_ID of candidate_idted_id
    - do RELATED_RECORD_ID related_id vloží candidate_id
    (obousměrné propojení)
    """
    for cid, rid in [(candidate_id, related_id), (related_id, candidate_id)]:
        con = db()
        existing = con.execute(
            "SELECT id FROM occurrence_fields "
            "WHERE candidate_id=? AND field_name='RELATED_RECORD_ID'",
            (cid,)).fetchone()
        if existing:
            # Přidat k existujícím (odděleno středníkem)
            cur_val = con.execute(
                "SELECT field_value FROM occurrence_fields WHERE id=?",
                (existing["id"],)).fetchone()["field_value"] or ""
            ids_set = set(cur_val.split(";"))
            ids_set.add(str(rid))
            ids_set.discard("")
            con.execute(
                "UPDATE occurrence_fields SET field_value=? WHERE id=?",
                (";".join(sorted(ids_set)), existing["id"]))
        else:
            con.execute(
                "INSERT INTO occurrence_fields "
                "(candidate_id,field_name,field_value,method) VALUES (?,?,?,?)",
                (cid, "RELATED_RECORD_ID", str(rid), "synonym_link"))
        con.commit(); con.close()



# ══════════════════════════════════════════════════════════════════════════════
# OCR-TYPO DETEKCE V AUTORSKÝCH JMÉNECH  (pole AUTHOR)
# ══════════════════════════════════════════════════════════════════════════════

_AUTHOR_TOKEN_RE = re.compile(r"[A-ZÀ-ÖØ-Þ][a-zà-öø-ÿ'’\-]{2,}")
_AUTHOR_STOPWORDS = {
    "et", "al", "in", "emend", "new", "and", "und", "sp", "nov", "the",
    "non", "sensu", "auctorum", "partim",
}


def _extract_author_tokens(value: str) -> List[str]:
    """Extracts probable surname tokens from free-text AUTHOR field content."""
    tokens = _AUTHOR_TOKEN_RE.findall(value or "")
    return [t for t in tokens if t.lower() not in _AUTHOR_STOPWORDS]


def find_author_ocr_typos(threshold: float = 0.82, min_len: int = 4) -> List[Dict[str, Any]]:
    """
    Compares unique author names across all filled AUTHOR fields
    and finds pairs s vysokou textovou podobností, ale nikoliv totožné —
    a typical OCR-error signature (e.g. "Sysoey" vs "Sysoev", "Val'kovy" vs
    "Val'kov"). Vrací návrhy seřazené podle podobnosti; jméno s VÍCE
    výskyty je označeno jako pravděpodobně správné, řidší varianta jako
    možný překlep. These are suggestions for manual verification, nic se needituje
    automaticky.
    """
    con = db()
    rows = con.execute("""
        SELECT of.candidate_id, of.field_value, tc.taxon_name, tc.page_start,
               d.filename
        FROM occurrence_fields of
        JOIN taxon_candidates tc ON of.candidate_id = tc.id
        JOIN documents d ON tc.document_id = d.id
        WHERE of.field_name = 'AUTHOR' AND of.field_value != ''
    """).fetchall()
    con.close()

    occurrences: Dict[str, List[Dict[str, Any]]] = {}
    for r in rows:
        for tok in _extract_author_tokens(r["field_value"]):
            if len(tok) < min_len:
                continue
            occurrences.setdefault(tok, []).append({
                "candidate_id": r["candidate_id"], "taxon_name": r["taxon_name"],
                "filename": r["filename"], "page_start": r["page_start"],
                "field_value": r["field_value"],
            })

    names = sorted(occurrences.keys())
    seen_pairs = set()
    suggestions = []
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = names[i], names[j]
            if a[0].lower() != b[0].lower():
                continue
            if abs(len(a) - len(b)) > 2:
                continue
            ratio = difflib.SequenceMatcher(None, a.lower(), b.lower()).ratio()
            if threshold <= ratio < 1.0:
                pid = tuple(sorted([a, b]))
                if pid in seen_pairs:
                    continue
                seen_pairs.add(pid)
                cnt_a, cnt_b = len(occurrences[a]), len(occurrences[b])
                if cnt_a >= cnt_b:
                    common, common_cnt, rare, rare_cnt = a, cnt_a, b, cnt_b
                else:
                    common, common_cnt, rare, rare_cnt = b, cnt_b, a, cnt_a
                suggestions.append({
                    "pravděpodobně_správně": common, "výskytu_správně": common_cnt,
                    "možný_překlep": rare, "výskytu_překlepu": rare_cnt,
                    "podobnost": round(ratio, 3),
                    "ukázka": occurrences[rare][0],
                })
    suggestions.sort(key=lambda s: -s["podobnost"])
    return suggestions


# ══════════════════════════════════════════════════════════════════════════════
# UKONČENÍ APLIKACE
# ══════════════════════════════════════════════════════════════════════════════


def _terminate_application() -> None:
    """
    Terminates the running PaleoN process (Streamlit server) and attempts to close
    the terminalu (cmd/PowerShell na Windows, shell na macOS/Linuxu), ze kterého
    byl proces spuštěn.

    Mechanismus:
      1. Zjistí PID nadřazeného procesu (typically cmd.exe / shell terminálu).
      2. Pokusí se ukončit i tento nadřazený proces — to je krok, který
         na Windows spolehlivě zavře celé okno cmd/PowerShell.
      3. Nakonec ukončí vlastní (Streamlit) proces přes os._exit(), což
         okamžitě zastaví server bez čekání na úklid Streamlit runtime.

    Na macOS/Linuxu závisí úspěšnost kroku 2 na typu terminálu — některé
    terminály (Terminal.app, gnome-terminal spuštěné přímo s příkazem)
    se zavřou spolu se shellem, jiné (tmux, screen, IDE integrovaný
    terminál) mohou zustat otevřené i po ukončení shellu.
    """
    import os as _os
    import sys as _sys
    import signal as _signal
    import platform as _platform

    my_pid = _os.getpid()
    parent_pid = _os.getppid()

    try:
        if _platform.system() == "Windows":
            # /T = ukončit i potomky, /F = vynutit. Cílíme na rodičovský
            # proces (cmd/PowerShell), což na Windows zavře celé okno.
            _os.system(f"taskkill /F /T /PID {parent_pid} >NUL 2>&1")
        else:
            # macOS / Linux — poslat SIGTERM nadřazenému shellu; pokud
            # terminál ukončuje okno spolu se shellem, zavře se i okno.
            try:
                _os.kill(parent_pid, _signal.SIGTERM)
            except Exception:
                pass
    except Exception:
        pass  # nadřazený proces se nepodařilo ukončit — pokračujeme aspoň vlastním ukončením

    # Vlastní proces ukončit vždy, i kdyby se předchozí krok nezdařil.
    _os._exit(0)


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

def sidebar_ui():
    st.sidebar.title(f"🦕 {APP_NAME}")
    st.sidebar.caption(f"v{APP_VERSION}")
    s = get_settings()

    # ── Aktuální uživatel + přepínač ─────────────────────────────────────
    _cur_user = st.session_state.get("pn_user", "")
    _ucol1, _ucol2 = st.sidebar.columns([3, 1])
    _ucol1.markdown(
        f"<div style='font-size:0.78rem;margin:2px 0'>"  
        f"👤 <b>{_cur_user}</b></div>",
        unsafe_allow_html=True)
    if _ucol2.button("⇄", key="sb_switch_user", help="Switch user"):
        st.session_state.pop("pn_user", None)
        st.rerun()
    st.sidebar.divider()

    # ── Language / Language ─────────────────────────────────────────────────────
    current_lang = st.session_state.get("app_lang", s.get("lang", "en"))
    st.session_state["app_lang"] = "en"

    # ── Dark/Light mode ─────────────────────────────────────────────────────
    is_dark = s.get("theme", "dark") == "dark"
    toggled_dark = st.sidebar.toggle(t("dark_mode"), value=is_dark, key="sb_theme_toggle")
    new_theme = "dark" if toggled_dark else "light"
    if new_theme != s.get("theme", "dark"):
        s["theme"] = new_theme
        st.session_state["paleon_settings"] = s
        save_settings_to_disc(s)
        st.rerun()
    st.sidebar.divider()

    # ── Rychlé statistiky ────────────────────────────────────────────────────
    con = db()
    n_docs   = con.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
    n_pend   = con.execute("SELECT COUNT(*) FROM taxon_candidates WHERE status='pending'").fetchone()[0]
    n_appr   = con.execute("SELECT COUNT(*) FROM taxon_candidates WHERE status='approved'").fetchone()[0]
    n_low    = con.execute("SELECT COUNT(*) FROM taxon_candidates WHERE status='low_confidence'").fetchone()[0]
    con.close()

    sc1, sc2 = st.sidebar.columns(2)
    sc1.metric(t("docs"),     n_docs)
    sc2.metric(t("approved"), n_appr)
    sc3, sc4 = st.sidebar.columns(2)
    sc3.metric(t("pending"),  n_pend)
    sc4.metric(t("low_conf"), n_low)
    st.sidebar.divider()

    # ── LM Studio — kompaktní status + detail v expanderu ───────────────────
    _lm_detected = st.session_state.get("lm_studio_detected", False)
    _lm_enabled  = s.get("llm_enabled", False)
    _lm_model    = s.get("lmstudio_model", "") or ""
    _lm_models   = st.session_state.get("lm_models", [])

    if _lm_detected and _lm_enabled:
        _chip_label = f"🤖 LM Studio · {_lm_model[:35] or '?'}"
        st.sidebar.success(_chip_label)
    elif _lm_detected:
        st.sidebar.info("🔌 LM Studio available — LLM is disabled")
    else:
        st.sidebar.caption("⚫ LM Studio not found")

    with st.sidebar.expander("⚙️ LM Studio — settings", expanded=False):
        s["llm_enabled"] = st.toggle(
            "Enable LLM", value=_lm_enabled,
            help="Enables the LLM assistant (translation, extraction, validation).")
        s["lmstudio_base_url"] = st.text_input(
            "Base URL", value=s.get("lmstudio_base_url","http://localhost:1234/v1"),
            label_visibility="collapsed",
            placeholder="http://localhost:1234/v1")

        if s["llm_enabled"]:
            m_col, r_col = st.columns([3, 1])
            if r_col.button("🔄", key="sb_refresh", help="Reload models from LM Studio"):
                try:
                    _fresh = lm_models(s)
                    if _fresh:
                        st.session_state["lm_models"] = _fresh
                        st.session_state["lm_studio_detected"] = True
                        st.toast(tt(f"✅ {len(_fresh)} models", f"✅ {len(_fresh)} models"), icon="✅")
                    else:
                        st.toast("Server is running, but no models were found", icon="⚠️")
                except RuntimeError as _e:
                    st.toast(str(_e)[:80], icon="❌")
                st.rerun()

            if _lm_models:
                _cur = s.get("lmstudio_model","")
                _idx = _lm_models.index(_cur) if _cur in _lm_models else 0
                s["lmstudio_model"] = m_col.selectbox(
                    "Model", _lm_models, index=_idx, key="sb_model",
                    label_visibility="collapsed")
            else:
                s["lmstudio_model"] = m_col.text_input(
                    "Model ID", value=_lm_model, key="sb_model_txt",
                    label_visibility="collapsed", placeholder="model-id")

            _t1, _t2 = st.columns([1, 1])
            _t1.markdown("<div style='padding-top:8px'>Temperature</div>",
                         unsafe_allow_html=True)
            s["llm_temperature"] = _t2.number_input(
                "Temperature", 0.0, 1.0, float(s.get("llm_temperature", 0.0)), 0.05,
                key="sb_temp", label_visibility="collapsed")
            s["llm_timeout"] = st.number_input(
                "Timeout (s)", 30, 600, int(s.get("llm_timeout", 180)),
                key="sb_timeout")

            st.divider()
            if st.button(t("redetect_lmstudio"), key="sb_redetect"):
                # Vymazat cache detekce → příštím rerunu proběhne znovu
                for _k in ["lm_studio_detected", "lm_models"]:
                    st.session_state.pop(_k, None)
                st.rerun()

            s["llm_auto_translate"] = st.toggle(
                "🌐 Translate during indexing (non-EN)",
                value=s.get("llm_auto_translate", False),
                key="sb_auto_translate",
                help="Starsi volba — v Knihovne je nyni toggle primo u uploadu.")
        else:
            st.caption(t("llm_disabled_caption"))

    st.sidebar.divider()
    if st.sidebar.button(t("save_settings"), key="sb_save",
                         help="Saves to paleon_data/settings.json"):
        save_settings_to_disc(s)
        save_prompts(s)
        st.sidebar.success("✓")

    st.session_state["paleon_settings"] = s

    # ── Ukončit aplikaci ─────────────────────────────────────────────────────
    st.sidebar.divider()
    if st.sidebar.button(t("close_app"), key="sb_close_app"):
        st.session_state["sb_confirm_close"] = True

    if st.session_state.get("sb_confirm_close"):
        st.sidebar.warning(f"⚠️ {t('confirm_close')}")
        cc1, cc2 = st.sidebar.columns(2)
        if cc1.button(t("yes"), key="sb_close_yes", type="primary"):
            st.sidebar.info("…")
            _terminate_application()
        if cc2.button(t("cancel"), key="sb_close_no"):
            st.session_state.pop("sb_confirm_close", None)
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 – KNIHOVNA
# ══════════════════════════════════════════════════════════════════════════════

def _delete_document(doc_id: int) -> None:
    con = db()
    _delete_candidates_for_document(doc_id, con=con)
    con.execute("DELETE FROM pages WHERE document_id=?", (doc_id,))
    con.execute("DELETE FROM documents WHERE id=?", (doc_id,))
    con.commit(); con.close()
    # Document (a případně i jeho id, pokud SQLite rowid znovu přidělí)
    # už neexists — uklidit i cache text units, ať nezustávají viset.
    invalidate_text_units_cache(doc_id)


def _reindex_document(doc_id: int, path: pathlib.Path, s: Dict) -> Dict:
    pages = extract_pages_from_file(path, s)
    char_count = sum(len(p.text) for p in pages)
    # Auto-OCR: pokud je nativní text chudý a Tesseract je k dispozici, opakuj s OCR
    avg_ch = (char_count // len(pages)) if pages else 0
    if avg_ch < int(s.get("pdf_min_chars", 80)) and _HAS_ANY_OCR():
        _ocr_s = dict(s); _ocr_s["ocr_enabled"] = True
        pages_ocr = extract_pages_from_file(path, _ocr_s)
        if sum(len(p.text) for p in pages_ocr) > char_count:
            pages = pages_ocr
            char_count = sum(len(p.text) for p in pages)
    con = db()
    _delete_candidates_for_document(doc_id, con=con)
    con.execute("DELETE FROM pages WHERE document_id=?", (doc_id,))
    for pg in pages:
        con.execute(
            "INSERT INTO pages (document_id,page_number,text,method,ocr_note,layout_note) VALUES (?,?,?,?,?,?)",
            (doc_id, pg.page_number, pg.text, pg.method, pg.ocr_note, pg.layout_note))
    con.execute("UPDATE documents SET page_count=?, char_count=? WHERE id=?",
                (len(pages), char_count, doc_id))
    # Obsah stránek se změnil → zneplatnit cache text units (jinak by
    # extract_block_for_candidate dál vracelo bloky ze STARÉHO textu).
    _bump_pages_version(doc_id, con=con)
    con.commit(); con.close()
    invalidate_text_units_cache(doc_id)
    return detect_candidates(doc_id, pages, s)


def tab_library():
    st.header(t("library_header"))

    # ── Přepínač pro prohlížení jiných uživatelu (read-only) ────────────
    _all_u = _get_all_users()
    _cur   = st.session_state.get("pn_user", "")
    if len(_all_u) > 1:
        _u_names = [u["username"] for u in _all_u]
        _view_u  = st.session_state.get("lib_view_user", _cur)
        def _fmt_user(u):
            is_a = any(x["username"]==u and x["is_admin"] for x in _all_u)
            suffix = "  (ja)" if u==_cur else "  (read-only)"
            prefix = "Admin " if is_a else ""
            return f"{prefix}{u}{suffix}"
        _selected_u = st.selectbox(
            "Browse user library:",
            options=_u_names,
            index=_u_names.index(_view_u) if _view_u in _u_names else 0,
            key="lib_view_user_sel",
            format_func=_fmt_user)
        if _selected_u != _view_u:
            st.session_state["lib_view_user"] = _selected_u
            st.rerun()
        if _selected_u != _cur:
            st.warning(
                f"Prohlizte knihovnu uzivatele **{_selected_u}** (read-only). "
                "Akce jako schvaleni a export funguji jen na vlastni knihovne.")
            st.session_state["lib_readonly_user"] = _selected_u
        else:
            st.session_state.pop("lib_readonly_user", None)
    st.divider()
    s = get_settings()

    # ── Settings indexace (OCR + Detekce) ──────────────────────────────────
    with st.expander(t("indexing_settings"), expanded=False):
        _ic1, _ic2, _ic3 = st.columns(3)
        with _ic1:
            st.caption(t("ocr_caption"))
            s["ocr_enabled"] = st.toggle(
                "Enable OCR (Tesseract)", value=s.get("ocr_enabled", True),
                key="lib_ocr_enabled",
                help="Automatically runs OCR on pages with insufficient native text." +
                     (" ✓ OCR available (" + ("easyocr " if HAS_EASYOCR else "") + ("fitz " if HAS_FITZ else "") + ("tesseract" if HAS_TESSERACT else "") + ")." if _HAS_ANY_OCR() else " ⚠️ No OCR engine found (install easyocr: pip install easyocr)."))
            s["ocr_languages"] = st.text_input(
                "Languages (Tesseract codes)", value=s.get("ocr_languages", "eng+ces+rus+deu+fra"),
                key="lib_ocr_langs", help="E.g. eng+ces+rus+deu+fra+chi_sim")
            s["pdf_min_chars"] = st.number_input(
                "OCR threshold (chars/page)", 10, 500, int(s.get("pdf_min_chars", 80)),
                key="lib_min_chars",
                help="Pages with fewer characters than this threshold will be OCRed (if OCR is enabled).")
            s["use_column_detection"] = st.toggle(
                "Two-column PDF", value=s.get("use_column_detection", True),
                key="lib_two_col")
        with _ic2:
            st.caption(t("candidate_detection_caption"))
            s["taxon_min_confidence"] = st.slider(
                "Min. score (pending)", 0.3, 0.9,
                float(s.get("taxon_min_confidence", 0.60)), 0.05, key="lib_min_conf")
            s["taxon_low_confidence"] = st.slider(
                "Min. score (low-conf)", 0.2, 0.7,
                float(s.get("taxon_low_confidence", 0.45)), 0.05, key="lib_low_conf")
            s["outside_systematic_penalty"] = st.slider(
                "Penalizace mimo syst. sekci", 0.0, 0.5,
                float(s.get("outside_systematic_penalty", 0.20)), 0.05, key="lib_sys_pen")
            s["show_rejected"] = st.toggle(
                "Show rejected in Review", value=s.get("show_rejected", False),
                key="lib_show_rej")
        with _ic3:
            st.caption(t("save_caption"))
            if st.button(t("save_settings"), key="lib_save_settings", type="primary"):
                save_settings_to_disc(s)
                st.success(t("settings_saved_ok"))
            st.session_state["paleon_settings"] = s
            st.caption("OCR: " + ("✅ " + ", ".join(filter(None, ["easyocr" if HAS_EASYOCR else "", "fitz" if HAS_FITZ else "", "tesseract" if HAS_TESSERACT else ""])) if _HAS_ANY_OCR() else t("ocr_no_engine")))

    # ── Upload (podpora více fileu najednou) ────────────────────────────────
    with st.expander(t("upload_new"), expanded=True):
        # ── Dynamic key pro file_uploader: změna klíče = skutečné vymazání fronty
        if "uploader_key_counter" not in st.session_state:
            st.session_state["uploader_key_counter"] = 0
        _up_key = f"uploader_{st.session_state['uploader_key_counter']}"

        u1, u2, u3 = st.columns([3, 1, 1])
        uploaded_files = u1.file_uploader(
            "PDF / DOCX / TXT — drag files here or click. Indexing starts after pressing Upload.",
            type=["pdf","docx","txt"],
            key=_up_key, label_visibility="collapsed",
            accept_multiple_files=True)
        lang_opt = u2.selectbox(
            "Language", ["en","cs","de","fr","ru","zh","mixed"],
            key="lang_sel", label_visibility="collapsed",
            help="The language will be used for all uploaded files.")
        ocr_force = u3.toggle(
            "🔍 OCR", value=s.get("ocr_enabled", True),
            key="up_ocr_force",
            help="Run OCR during upload.",
            disabled=not _HAS_ANY_OCR())
        n_files = len(uploaded_files) if uploaded_files else 0

        # ── Toggle: přeložit po indexaci ─────────────────────────────────────
        # Zobrazit kdykoli je LLM dostupné (bez ohledu na jazyk / stav uploaderu).
        # Streamlit zachová hodnotu přes reruns díky klíči widgetu.
        _do_translate_after = False
        if s.get("llm_enabled"):
            _do_translate_after = st.toggle(
                "🌐 Translate and map fields after indexing",
                value=False,
                key="lib_translate_after",
                help=(
                    "Po kazdem nahranem Documentu:\n1. Prelozi nalezena fields do anglictiny (preskoci, pokud jsou jiz EN or byly extrahovany primym LLM pasem pro CJK/rustinu).\n2. Premapuje sekce jako 'Characteristics' -> DESCRIPTION, 'Age and Distribution' -> OCCURRENCE atd.\n\nVyzaduje: LM Studio s nactenym modelem."
                ),
            )

        # ── Manuální indexace: filey se indexují až po stisku tlačítka ──────
        _up_hash = hash(tuple(sorted(f.name + str(f.size) for f in uploaded_files))) if uploaded_files else 0
        _up_hash_key = "lib_last_upload_hash"
        _already_done = st.session_state.get(_up_hash_key) == _up_hash and _up_hash != 0

        do_up = False
        if uploaded_files and not _already_done:
            _btn_col, _clr_col = st.columns([3, 1])
            _do_upload_btn = _btn_col.button(
                f"⬆️ Upload and index ({n_files} file{'u' if n_files != 1 else ''})",
                key="lib_do_upload_btn", type="primary",
                help="Click to start indexing all queued files.")
            if _clr_col.button("🗑️ Clear queue", key="lib_clear_queue",
                               help="Removes files from the queue without uploading."):
                st.session_state["uploader_key_counter"] += 1
                st.rerun()
            do_up = _do_upload_btn
        elif uploaded_files and _already_done:
            st.caption(tt(f"✓ {n_files} files processed. Add new files for further indexing.", f"✓ {n_files} files processed. Add new files for further indexing."))
            do_up = False

        if uploaded_files and do_up:
            # Lokální kopie settings s OCR přepsaným dle inline přepínače
            _up_s = dict(s)
            _up_s["ocr_enabled"] = ocr_force
            overall = st.progress(0, tt(f"Zpracovavam 0/{n_files} files…", f"Processing 0/{n_files} files…"))
            results = []
            for fi, uploaded in enumerate(uploaded_files):
                overall.progress(
                    int(fi / n_files * 100),
                    f"📄 {uploaded.name} ({fi+1}/{n_files})…")
                try:
                    dest = UPLOADS_DIR / uploaded.name
                    dest.write_bytes(uploaded.getbuffer())
                    pages = extract_pages_from_file(dest, _up_s)
                    char_count = sum(len(p.text) for p in pages)
                    # Auto-OCR: pokud je málo textu a OCR nebylo zapnuto, zkus znovu
                    avg_ch_pre = (char_count // len(pages)) if pages else 0
                    if avg_ch_pre < int(_up_s.get("pdf_min_chars", 80)) and not ocr_force and _HAS_ANY_OCR():
                        _ocr_s = dict(_up_s); _ocr_s["ocr_enabled"] = True
                        pages = extract_pages_from_file(dest, _ocr_s)
                        char_count = sum(len(p.text) for p in pages)
                    con = db()
                    cur = con.execute(
                        "INSERT INTO documents (filename,path,lang,page_count,char_count,notes,created_at) "
                        "VALUES (?,?,?,?,?,'',?)",
                        (uploaded.name, str(dest), lang_opt, len(pages), char_count,
                         datetime.now().isoformat()))
                    doc_id = cur.lastrowid
                    for pg in pages:
                        con.execute(
                            "INSERT INTO pages (document_id,page_number,text,method,ocr_note,layout_note) "
                            "VALUES (?,?,?,?,?,?)",
                            (doc_id, pg.page_number, pg.text, pg.method, pg.ocr_note, pg.layout_note))
                    con.commit(); con.close()
                    diag = detect_candidates(doc_id, pages, s)
                    avg_ch = (char_count // len(pages)) if pages else 0
                    ocr_warn = "⚠️ little text — consider OCR" if avg_ch < 60 else ""

                    # ── Post-indexační pipeline: remap + volitelný překlad ────
                    # Spouští se vždy (remap), překlad jen pokud toggle zapnut
                    # a jazyk Documentu není English.
                    _n_tr_total = 0
                    _n_remap_total = 0
                    _lang_is_en = lang_opt.lower() in {"en", "eng", "english"}
                    _need_translate = (
                        _do_translate_after
                        and s.get("llm_enabled")
                        and not _lang_is_en
                    )
                    _n_found = diag.get("Accepted", 0) + diag.get("Low-confidence", 0)

                    if _n_found > 0:
                        _con_tr = db()
                        _tr_cands = _con_tr.execute(
                            "SELECT id, block_text FROM taxon_candidates "
                            "WHERE document_id=? AND status IN ('pending','low_confidence')",
                            (doc_id,)).fetchall()
                        _con_tr.close()

                        if _tr_cands:
                            # Pokud indexace spustila llm_extract_fields_cjk_ru,
                            # pole jsou JIŽ v angličtině (method='llm').
                            # Překlad by byl redundantní → přeskočit celý Document.
                            _con_chk = db()
                            _doc_llm_fields = _con_chk.execute(
                                "SELECT COUNT(*) FROM occurrence_fields of "
                                "JOIN taxon_candidates tc ON of.candidate_id=tc.id "
                                "WHERE tc.document_id=? AND of.method='llm'",
                                (doc_id,)).fetchone()[0]
                            _con_chk.close()
                            _skip_translate = _need_translate and bool(_doc_llm_fields)

                            _tr_prog = st.progress(0, "Processing records…")
                            for _tci, _tc in enumerate(_tr_cands):
                                _cid = _tc["id"]
                                _blk = _tc["block_text"] or ""
                                try:
                                    if _need_translate and not _skip_translate:
                                        _n = auto_translate_candidate_fields(
                                            _cid, lang_opt, s,
                                            force=True, force_lang=lang_opt,
                                            block=_blk)
                                        _n_tr_total += _n
                                    # Remapování (vždy): Characteristics→DESCRIPTION…
                                    _n_remap_total += remap_fields_post_translation(_cid)
                                except Exception as _te:
                                    logging.warning(
                                        "Post-index pipeline cand %s: %s", _cid, _te)
                                _tr_prog.progress(
                                    int((_tci + 1) / len(_tr_cands) * 100),
                                    f"Records {_tci+1}/{len(_tr_cands)}…")
                            _tr_prog.empty()

                    results.append({
                        "File": uploaded.name, "ID": doc_id,
                        "Pages": len(pages), "Zn./str.": avg_ch,
                        "✅ Pending": diag["Accepted"],
                        "🔵 Low-conf": diag["Low-confidence"],
                        "❌ Rejected": diag["Rejected"],
                        "🌐 Translated": _n_tr_total or ("—" if not _need_translate else 0),
                        "♻️ Remap": _n_remap_total or 0,
                        "Stav": ocr_warn if ocr_warn else "✓ OK",
                    })
                except Exception as exc:
                    results.append({
                        "File": uploaded.name, "ID": "—",
                        "Pages": "—", "Chars./page.": "—", "✅ Pending": "—",
                        "🔵 Low-conf": "—", "❌ Rejected": "—",
                        "Stav": f"✗ Error: {exc}",
                    })

            overall.progress(100, "Hotovo ✓")
            ok_count = sum(1 for r in results if "✓ OK" in r.get("Stav",""))
            ocr_warn_count = sum(1 for r in results if "⚠" in r.get("Stav",""))
            # Zobrazit výsledky PŘED rerununem (toast přežije rerun, success ne)
            _toast_msg = f"✅ Indexed {ok_count}/{n_files} documents"
            if ocr_warn_count:
                _toast_msg += f" · ⚠️ {ocr_warn_count}× little text (consider OCR)"
            st.toast(_toast_msg, icon="✅")
            st.session_state["last_index_results"] = results  # pro zobrazení po rerunu
            st.session_state[_up_hash_key] = _up_hash
            st.session_state["uploader_key_counter"] += 1
            st.rerun()

    # ── Zobrazit výsledky posledního indexování (po rerunu) ──────────────────
    _last_results = st.session_state.pop("last_index_results", None)
    if _last_results:
        ok_c = sum(1 for r in _last_results if "✓ OK" in r.get("Stav",""))
        ocr_c = sum(1 for r in _last_results if "⚠" in r.get("Stav",""))
        err_c = sum(1 for r in _last_results if "✗" in r.get("Stav",""))
        st.success(
            f"✅ **Indexing complete** — {ok_c} OK"
            + (f", {ocr_c}× ⚠️ little text" if ocr_c else "")
            + (f", {err_c}× ✗ chyba" if err_c else ""))
        st.dataframe(pd.DataFrame(_last_results), use_container_width=True, hide_index=True)

    # ── Celkový souhrn ───────────────────────────────────────────────────────
    con = db()
    docs = [dict(r) for r in con.execute("SELECT * FROM documents ORDER BY created_at DESC").fetchall()]
    con.close()
    if not docs:
        st.info(t("no_docs"))
        return

    total_pages = sum(d["page_count"] for d in docs)
    con = db()
    tot_c  = con.execute("SELECT COUNT(*) FROM taxon_candidates").fetchone()[0]
    tot_ap = con.execute("SELECT COUNT(*) FROM taxon_candidates WHERE status='approved'").fetchone()[0]
    tot_fi = con.execute("SELECT COUNT(*) FROM occurrence_fields WHERE field_value!=''").fetchone()[0]
    con.close()

    st.divider()
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("📄 Documents",   len(docs))
    m2.metric("📃 Pages",       total_pages)
    m3.metric("🔍 Candidates",   tot_c)
    m4.metric("✅ Approved",   tot_ap)
    m5.metric("📝 Pole",        tot_fi)
    st.divider()

    # ── Tabulka documents s checkboxy ─────────────────────────────────────────
    st.markdown(
        "**Documents in the library** — check rows pro bulk actions.  "
        "Columns: **✅** = approved &nbsp;│&nbsp; **⏳** = awaiting review "
        "&nbsp;│&nbsp; **🔵** = low confidence (low-conf) &nbsp;│&nbsp; "
        "**❌** = rejected",
        unsafe_allow_html=False)

    con = db()
    status_by_doc: Dict[int, Dict[str,int]] = {}
    for did_row in docs:
        did = did_row["id"]
        cnt = {r["status"]:r["n"] for r in con.execute(
            "SELECT status, COUNT(*) as n FROM taxon_candidates WHERE document_id=? GROUP BY status",
            (did,)).fetchall()}
        status_by_doc[did] = cnt
    con.close()

    df_docs = pd.DataFrame([{
        "_id": d["id"],
        "✓": False,
        "File": d["filename"],
        "Pages": d["page_count"],
        "✅": status_by_doc.get(d["id"], {}).get("approved", 0),
        "⏳": status_by_doc.get(d["id"], {}).get("pending", 0),
        "🔵": status_by_doc.get(d["id"], {}).get("low_confidence", 0),
        "❌": status_by_doc.get(d["id"], {}).get("rejected", 0),
        "Language": d["lang"],
        "Uploaded": (d["created_at"] or "")[:16],
    } for d in docs])

    sa_col1, sa_col2 = st.columns([1,3])
    select_all = sa_col1.checkbox("☑️ Select all", key="lib_select_all_chk")
    if select_all != st.session_state.get("_lib_select_all_prev", False):
        st.session_state.pop("library_table", None)
        st.session_state["_lib_select_all_prev"] = select_all
    lib_search = sa_col2.text_input(
        "🔍 Hledat Document", key="lib_search", placeholder="filter by name…",
        label_visibility="collapsed")

    df_show = df_docs.drop(columns=["_id"])
    if lib_search:
        keep_mask = df_show["File"].str.lower().str.contains(lib_search.lower(), na=False)
        df_docs = df_docs[keep_mask].reset_index(drop=True)
        df_show = df_show[keep_mask].reset_index(drop=True)
    if select_all:
        df_show["✓"] = True
    edited_docs = st.data_editor(
        df_show,
        column_config={
            "✓":      st.column_config.CheckboxColumn("✓", width=40),
            "File": st.column_config.TextColumn("File", width=280, disabled=True),
            "Pages":  st.column_config.NumberColumn("Str.", width=55, disabled=True),
            "✅":      st.column_config.NumberColumn("✅", width=45, disabled=True),
            "⏳":      st.column_config.NumberColumn("⏳", width=45, disabled=True),
            "🔵":      st.column_config.NumberColumn("🔵", width=45, disabled=True),
            "❌":      st.column_config.NumberColumn("❌", width=45, disabled=True),
            "Language":  st.column_config.TextColumn("Language", width=60, disabled=True),
            "Uploaded":st.column_config.TextColumn("Uploaded", width=140, disabled=True),
        },
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="library_table",
    )

    selected_ids = [
        int(df_docs.iloc[i]["_id"])
        for i, row in edited_docs.iterrows()
        if row.get("✓", False)
    ]
    n_sel = len(selected_ids)

    # ── Hromadné akce ────────────────────────────────────────────────────────
    if n_sel:
        st.markdown(tt(f"**Selected: {n_sel} document(s)**", f"**Selected: {n_sel} document(s)**"))
        ba1, ba2, ba3, ba4 = st.columns(4)

        if ba1.button(f"🔄 Re-indexovat ({n_sel})", key="lib_batch_reindex"):
            prog = st.progress(0, "Re-indexuji…")
            results = []
            for i, did in enumerate(selected_ids):
                prog.progress(int((i)/n_sel*100), f"Re-indexuji {i+1}/{n_sel}…")
                con = db()
                row = con.execute("SELECT path FROM documents WHERE id=?", (did,)).fetchone()
                con.close()
                if row:
                    path = pathlib.Path(row["path"])
                    if path.exists():
                        diag = _reindex_document(did, path, s)
                        results.append((did, diag["Accepted"]))
            prog.empty()
            st.success(tt(f"Re-indexed {len(results)} documents.", f"Re-indexed {len(results)} documents."))
            st.rerun()

        if ba2.button(f"♻️ Reset detekce ({n_sel})", key="lib_batch_reset"):
            prog = st.progress(0, "Resetuji detekci…")
            for i, did in enumerate(selected_ids):
                prog.progress(int(i/n_sel*100), f"{i+1}/{n_sel}…")
                con = db()
                pgs = con.execute(
                    "SELECT page_number, text, method, ocr_note, layout_note "
                    "FROM pages WHERE document_id=?", (did,)).fetchall()
                _delete_candidates_for_document(did, con=con)
                con.commit(); con.close()
                pages_obj = [PageText(r["page_number"], r["text"], r["method"],
                                      r["ocr_note"], r["layout_note"]) for r in pgs]
                if pages_obj:
                    detect_candidates(did, pages_obj, s)
            prog.empty()
            st.success(tt(f"Detection reset for {n_sel} documents.", f"Detection reset for {n_sel} documents."))
            st.rerun()

        if ba3.button(f"📄 Export TXT ({n_sel})", key="lib_batch_export_txt"):
            con = db()
            lines = []
            for did in selected_ids:
                doc_row = con.execute("SELECT filename FROM documents WHERE id=?", (did,)).fetchone()
                pgs = con.execute(
                    "SELECT page_number, text, method, layout_note FROM pages "
                    "WHERE document_id=? ORDER BY page_number", (did,)).fetchall()
                lines.append(f"\n\n{'='*70}\nDocument: {doc_row['filename']}\n{'='*70}")
                for pg in pgs:
                    lines.append(f"\n--- Page {pg['page_number']} [{pg['method']}] "
                                 f"{pg['layout_note'] or ''} ---")
                    lines.append(pg["text"] or "")
            con.close()
            st.download_button(
                "📥 Download TXT (selected)",
                data="\n".join(lines).encode("utf-8"),
                file_name=f"paleon_library_export_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain", key="lib_dl_batch_txt")

        if ba4.button(f"🗑️ Delete ({n_sel})", key="lib_batch_delbtn"):
            st.session_state["lib_confirm_batch_del"] = True

        if st.session_state.get("lib_confirm_batch_del"):
            st.warning(tt(f"⚠️ Opravdu smazat {n_sel} selected documents and all their data?", f"⚠️ Really delete {n_sel} selected documents and all their data?"))
            yc, nc = st.columns(2)
            if yc.button("✅ Yes, delete all", key="lib_batch_del_yes", type="primary"):
                for did in selected_ids:
                    _delete_document(did)
                st.session_state.pop("lib_confirm_batch_del", None)
                st.success(tt(f"Deleted {n_sel} documents.", f"Deleted {n_sel} documents."))
                st.rerun()
            if nc.button("❌ Cancel", key="lib_batch_del_no"):
                st.session_state.pop("lib_confirm_batch_del", None)
                st.rerun()

    # ── Hromadné OCR ────────────────────────────────────────────────────────
    _ocr_col1, _ocr_col2, _ocr_col3 = st.columns([2, 2, 3])
    _ocr_low = [d for d in docs if (d.get("char_count") or 0) // max(d.get("page_count") or 1, 1) < int(s.get("pdf_min_chars", 80))]
    _engine_str = ", ".join(filter(None, [
        "easyocr" if HAS_EASYOCR else "",
        "fitz" if HAS_FITZ else "",
        "tesseract" if HAS_TESSERACT else ""])) or "žádný"
    if _ocr_col1.button(
            f"🔍 OCR — documents with little text ({len(_ocr_low)})",
            key="lib_batch_ocr_low",
            disabled=not _ocr_low,
            help="OCR jen na Documentech pod prahem zn./str."):
        _ocr_prog = st.progress(0, "OCR…")
        for _oi, _od in enumerate(_ocr_low):
            _ocr_prog.progress(int(_oi / max(len(_ocr_low),1) * 100),
                               f"OCR: {_od['filename']} ({_oi+1}/{len(_ocr_low)})…")
            _op = pathlib.Path(_od["path"])
            if _op.exists():
                _ocs = dict(s); _ocs["ocr_enabled"] = True; _ocs["pdf_min_chars"] = 99999
                _reindex_document(_od["id"], _op, _ocs)
        _ocr_prog.empty()
        st.success(tt(f"OCR complete for {len(_ocr_low)} documents.", f"OCR complete for {len(_ocr_low)} documents.")); st.rerun()
    if _ocr_col2.button(
            f"🔍 OCR — all documents ({len(docs)})",
            key="lib_batch_ocr_all",
            help="Forces OCR across all documents in the library regardless of text amount."):
        _ocr_prog2 = st.progress(0, "OCR…")
        for _oi, _od in enumerate(docs):
            _ocr_prog2.progress(int(_oi / max(len(docs),1) * 100),
                                f"OCR: {_od['filename']} ({_oi+1}/{len(docs)})…")
            _op = pathlib.Path(_od["path"])
            if _op.exists():
                _ocs = dict(s); _ocs["ocr_enabled"] = True; _ocs["pdf_min_chars"] = 99999
                _reindex_document(_od["id"], _op, _ocs)
        _ocr_prog2.empty()
        st.success(tt(f"OCR complete for {len(docs)} documents.", f"OCR complete for {len(docs)} documents.")); st.rerun()
    _ocr_col3.caption(f"Engine: {_engine_str}  |  Pod prahem: {len(_ocr_low)} / {len(docs)}")

    st.divider()

    # ── Detail jednoho Documentu ────────────────────────────────────────────
    st.markdown("**Document detail**")
    doc_options = {f"{d['filename']} (ID {d['id']})": d["id"] for d in docs}
    sel_label = st.selectbox("Select document for detail", list(doc_options.keys()),
                             key="lib_detail_sel")
    did = doc_options[sel_label]
    doc = next(d for d in docs if d["id"] == did)

    cnt = status_by_doc.get(did, {})
    ap = cnt.get("approved",0); pe = cnt.get("pending",0)
    lo = cnt.get("low_confidence",0); rj = cnt.get("rejected",0)

    r1,r2,r3,r4,r5 = st.columns(5)
    r1.metric("📄 Pages", doc["page_count"])
    r2.metric("✅ Approved", ap)
    r3.metric("⏳ Pending", pe)
    r4.metric("🔵 Low confidence", lo)
    r5.metric("❌ Rejected", rj)
    char_k = (doc["char_count"] or 0) // 1000
    st.caption(
        f"~{char_k}k characters |  "
        f"Uploaded: `{(doc['created_at'] or '')[:16]}`  |  "
        f"File: `{doc['path']}`")

    # Language Documentu a poznámka
    lang_options = ["en","cs","de","fr","ru","zh","mixed"]
    lc, nc = st.columns([1,3])
    cur_lang = doc["lang"] if doc["lang"] in lang_options else "en"
    new_lang = lc.selectbox(
        "🌐 Language", lang_options, index=lang_options.index(cur_lang),
        key=f"doclang_{did}")
    if new_lang != doc["lang"]:
        con = db(); con.execute("UPDATE documents SET lang=? WHERE id=?", (new_lang, did))
        con.commit(); con.close()
        st.rerun()

    notes_val = doc["notes"] or ""
    new_notes = nc.text_area(
        "📝 Note", value=notes_val, height=55,
        key=f"notes_{did}", placeholder="Optional note…")
    if new_notes != notes_val and st.button(t("doc_notes_save"), key=f"savenotes_{did}"):
        con = db(); con.execute("UPDATE documents SET notes=? WHERE id=?", (new_notes, did))
        con.commit(); con.close()
        st.rerun()

    # ── BOD 4: Metadata publikace ─────────────────────────────────────────
    with st.expander(t("biblio_metadata"), expanded=False):
        st.caption(t("biblio_caption"))
        _mp_c1, _mp_c2 = st.columns([3, 1])
        _mp_c3, _mp_c4, _mp_c5 = st.columns([3, 1, 1])
        _mp_doi     = _mp_c1.text_input("DOI", value=doc.get("doi","") or "",
                          key=f"doi_{did}", placeholder="10.xxxx/…")
        _mp_year    = _mp_c2.number_input("Rok", min_value=1700, max_value=2100,
                          value=int(doc["pub_year"]) if doc.get("pub_year") else 1900,
                          key=f"pyear_{did}", step=1)
        _mp_journal = _mp_c3.text_input("Journal / Proceedings",
                          value=doc.get("pub_journal","") or "",
                          key=f"pjournal_{did}")
        _mp_volume  = _mp_c4.text_input("Vol.", value=doc.get("pub_volume","") or "",
                          key=f"pvol_{did}")
        _mp_pages   = _mp_c5.text_input("Str.", value=doc.get("pub_pages","") or "",
                          key=f"ppages_{did}")
        _mp_authors = st.text_input("Authors",
                          value=doc.get("pub_authors","") or "",
                          key=f"pauthors_{did}",
                          placeholder="Novak O., Smith J.")
        _mp_title   = st.text_input("Paper title",
                          value=doc.get("pub_title","") or "",
                          key=f"ptitle_{did}")
        if st.button(t("save_metadata_btn"), key=f"savemeta_{did}", type="primary"):
            con = db()
            con.execute(
                "UPDATE documents SET doi=?,pub_authors=?,pub_journal=?,"
                "pub_year=?,pub_volume=?,pub_pages=?,pub_title=? WHERE id=?",
                (_mp_doi.strip(), _mp_authors.strip(), _mp_journal.strip(),
                 int(_mp_year) if _mp_year != 1900 else None,
                 _mp_volume.strip(), _mp_pages.strip(), _mp_title.strip(), did))
            con.commit(); con.close()
            st.success(t("metadata_saved"))
            st.rerun()
        # Formátovaná citace
        _cite_parts = []
        if doc.get("pub_authors"): _cite_parts.append(doc["pub_authors"])
        if doc.get("pub_year"):    _cite_parts.append(f"({doc['pub_year']})")
        if doc.get("pub_title"):   _cite_parts.append(doc["pub_title"] + ".")
        if doc.get("pub_journal"): _cite_parts.append(f"*{doc['pub_journal']}*")
        if doc.get("pub_volume"):  _cite_parts.append(doc["pub_volume"])
        if doc.get("pub_pages"):   _cite_parts.append(f"pp. {doc['pub_pages']}")
        if doc.get("doi"):         _cite_parts.append(
            f"https://doi.org/{doc['doi']}")
        if _cite_parts:
            st.markdown("**Citace:** " + " ".join(_cite_parts))

    # BOD 10: OCR quality indicator
    _ocr = _get_document_ocr_stats(did)
    if _ocr["n_pages"] > 0:
        with st.expander(
            f"📊 Text quality — {_ocr['n_pages']} pages, "
            f"{_ocr['avg_chars']} chars/page avg.",
            expanded=False):
            _oq1, _oq2, _oq3, _oq4 = st.columns(4)
            _oq1.metric("📄 Total pages",  _ocr["n_pages"])
            _oq2.metric("⚙️ Native text",  _ocr["n_native"],
                        help="fitz / pdfplumber / docx")
            _oq3.metric("🔍 OCR pages",      _ocr["n_ocr"],
                        help="Tesseract OCR — lower accuracy")
            _oq4.metric("⚠️ Short pages", _ocr["n_empty"],
                        help="< 100 characters — possibly empty or badly extracted")
            # Vizuální bar OCR vs nativní
            _native_pct = int(_ocr["n_native"] / max(_ocr["n_pages"],1) * 100)
            _ocr_pct    = int(_ocr["n_ocr"]    / max(_ocr["n_pages"],1) * 100)
            _unk_pct    = 100 - _native_pct - _ocr_pct
            st.markdown(
                f"<div style='margin:6px 0 2px;font-size:0.72rem;color:#94a3b8'>"
                f"Extraction share:</div>"
                f"<div style='display:flex;height:10px;border-radius:4px;overflow:hidden'>"
                f"<div style='width:{_native_pct}%;background:#10b981' "
                f"title='Native {_native_pct}%'></div>"
                f"<div style='width:{_ocr_pct}%;background:#f59e0b' "
                f"title='OCR {_ocr_pct}%'></div>"
                f"<div style='width:{_unk_pct}%;background:#475569' "
                f"title='Unknown {_unk_pct}%'></div>"
                f"</div>"
                f"<div style='font-size:0.68rem;margin:3px 0'>"
                f"<span style='color:#10b981'>■ Native {_native_pct}%</span> &nbsp;"
                f"<span style='color:#f59e0b'>■ OCR {_ocr_pct}%</span> &nbsp;"
                f"<span style='color:#94a3b8'>■ Unknown {_unk_pct}%</span></div>",
                unsafe_allow_html=True)
            if _ocr["low_quality_pages"]:
                st.caption(
                    f"⚠️ Pages with short text (<200 chars): "
                    + ", ".join(str(p) for p in _ocr["low_quality_pages"]))
            if _ocr["avg_chars"] < 200:
                st.warning(
                    "The average amount of text per page is very short. Consider re-indexing using OCR or another method of text extraction.")

    st.divider()
    a1, a2, a3, a4, a5 = st.columns(5)

    if a1.button("🔄 Re-index", key=f"reindex_{did}",
                 help="Rung again text extraction and detection (global settings OCR applie)."):
        path = pathlib.Path(doc["path"])
        if path.exists():
            with st.spinner("Re-indexuji…"):
                diag = _reindex_document(did, path, s)
            avg_ch = (doc.get("char_count", 0) or 0) // max(doc.get("page_count", 1), 1)
            st.success(
                f"Hotovo: {diag['Accepted']} pending, {diag['Low-confidence']} low-conf." +
                (" ⚠️ Little text — try OCR." if avg_ch < int(s.get("pdf_min_chars", 80)) else ""))
            st.rerun()
        else:
            st.error(f"File not found: {doc['path']}")

    if a2.button("🔍 OCR", key=f"ocr_{did}",
                 disabled=not _HAS_ANY_OCR(),
                 type="primary",
                 help="Forces Tesseract OCR on all pages and re-indexes."
                      + ("" if _HAS_ANY_OCR() else " ⚠️ No OCR engine — install: pip install easyocr")):
        path = pathlib.Path(doc["path"])
        if path.exists():
            _ocr_s = dict(s)
            _ocr_s["ocr_enabled"] = True
            _ocr_s["pdf_min_chars"] = 99999  # vynuť OCR na každé stránce
            with st.spinner(t("ocr_running")):
                diag = _reindex_document(did, path, _ocr_s)
            st.success(f"OCR hotovo: {diag['Accepted']} pending, {diag['Low-confidence']} low-conf.")
            st.rerun()
        else:
            st.error(f"File not found: {doc['path']}")

    if a3.button("♻️ Reset det.", key=f"resetdet_{did}",
                 help="Deletes candidates, preserves text - run again the detection."):
        con = db()
        pgs = con.execute(
            "SELECT page_number, text, method, ocr_note, layout_note "
            "FROM pages WHERE document_id=?", (did,)).fetchall()
        _delete_candidates_for_document(did, con=con)
        con.commit(); con.close()
        pages_obj = [PageText(r["page_number"], r["text"], r["method"],
                              r["ocr_note"], r["layout_note"]) for r in pgs]
        if pages_obj:
            with st.spinner("Detekce…"):
                diag = detect_candidates(did, pages_obj, s)
            st.success(f"Detekce: {diag['Accepted']} pending, {diag['Low-confidence']} low-conf.")
            st.rerun()
        else:
            st.warning(t("no_pages_reindex"))

    if a4.button("📄 TXT pages", key=f"exptxt_{did}"):
        con = db()
        pgs = con.execute(
            "SELECT page_number, text, method, layout_note FROM pages "
            "WHERE document_id=? ORDER BY page_number", (did,)).fetchall()
        con.close()
        lines = [f"Document: {doc['filename']}", "="*60]
        for pg in pgs:
            lines.append(f"\n--- Page {pg['page_number']} [{pg['method']}] "
                         f"{pg['layout_note'] or ''} ---")
            lines.append(pg["text"] or "")
        st.download_button(
            "📥 Download TXT pages",
            data="\n".join(lines).encode("utf-8"),
            file_name=f"{pathlib.Path(doc['filename']).stem}_pages.txt",
            mime="text/plain", key=f"dl_pages_{did}")

    if a5.button("🗑️ Delete", key=f"delbtn_{did}"):
        st.session_state[f"confirm_del_{did}"] = True
    if st.session_state.get(f"confirm_del_{did}"):
        st.warning(tt(f"⚠️ Opravdu smazat **{doc['filename']}** and all data?", f"⚠️ Really delete **{doc['filename']}** and all data?"))
        yc, nc = st.columns(2)
        if yc.button("✅ Yes, smazat", key=f"delyes_{did}", type="primary"):
            _delete_document(did)
            st.session_state.pop(f"confirm_del_{did}", None)
            st.rerun()
        if nc.button("❌ Cancel", key=f"delno_{did}"):
            st.session_state.pop(f"confirm_del_{did}", None); st.rerun()

    # Prohlížeč pages
    with st.expander(t("page_viewer")):
        con = db()
        pgs = con.execute(
            "SELECT page_number, text, method, ocr_note, layout_note "
            "FROM pages WHERE document_id=? ORDER BY page_number", (did,)).fetchall()
        cands_pg: Dict[int, list] = {}
        for c in con.execute(
            "SELECT page_start, taxon_name, status, confidence "
            "FROM taxon_candidates WHERE document_id=? ORDER BY page_start", (did,)).fetchall():
            cands_pg.setdefault(c["page_start"], []).append(c)
        con.close()

        if not pgs:
            st.info(t("no_pages"))
        else:
            page_nums = [pg["page_number"] for pg in pgs]
            sel_page = st.selectbox("Page", page_nums,
                                    format_func=lambda n: f"Page {n}",
                                    key=f"pgsel_{did}")
            pg_row = next((p for p in pgs if p["page_number"]==sel_page), None)
            if pg_row:
                st.caption(
                    f"Method: `{pg_row['method']}`  |  "
                    f"Layout: `{pg_row['layout_note'] or 'N/A'}`  |  "
                    f"OCR: `{pg_row['ocr_note'] or 'N/A'}`")
                page_c = cands_pg.get(sel_page, [])
                if page_c:
                    for c in page_c:
                        st.markdown(
                            f"{_conf_badge(c['confidence'])} "
                            f"[{c['confidence']:.2f}] "
                            f"**{c['taxon_name']}** "
                            f"{_status_badge(c['status'])}",
                            unsafe_allow_html=False)

                doc_path = pathlib.Path(doc["path"]) if doc["path"] else None
                is_pdf = bool(doc_path and doc_path.suffix.lower() == ".pdf")

                if is_pdf and HAS_FITZ:
                    img_col, txt_col = st.columns([1, 1])
                    with img_col:
                        st.caption(t("pdf_page_real"))
                        mtime = doc_path.stat().st_mtime if doc_path.exists() else 0.0
                        png = _render_pdf_page_image(str(doc_path), sel_page, mtime)
                        if png:
                            st.image(png, use_container_width=True)
                        else:
                            st.warning(t("page_preview_failed"))
                    with txt_col:
                        st.text_area(t("page_text_label"), value=pg_row["text"] or "",
                                     height=420, key=f"pgtext_{did}_{sel_page}",
                                     disabled=True)
                else:
                    if not HAS_FITZ:
                        st.caption(t("pdf_page_requires_pymupdf"))
                    elif not is_pdf:
                        st.caption(t("preview_pdf_docs_only"))
                    st.text_area(t("page_text_label"), value=pg_row["text"] or "",
                                 height=280, key=f"pgtext_{did}_{sel_page}",
                                 disabled=True)

    # Statistiky / dashboard napříč knihovnou
    _library_stats_dashboard()



def _get_document_ocr_stats(document_id: int) -> Dict[str, Any]:
    """
    Returns text quality statistics for the given document.
    Analyses the pages table and returns:
      - n_pages: celkový count pages
      - n_native: pagesy extrahované nativně (fitz, pdfplumber)
      - n_ocr: pagesy přes OCR (tesseract)
      - n_empty: pagesy bez textu (krátký text < 100 znaku)
      - avg_chars: pruměrná délka textu na pagesu
      - ocr_ratio: podíl OCR pages (0–1)
      - low_quality_pages: seznam čísel pages s krátkým textem
    """
    con = db()
    pages = con.execute(
        "SELECT page_number, method, ocr_note, "
        "       COALESCE(LENGTH(text), 0) AS text_len "
        "FROM pages WHERE document_id=? ORDER BY page_number",
        (document_id,),
    ).fetchall()
    con.close()

    if not pages:
        return {
            "n_pages": 0, "n_native": 0, "n_ocr": 0, "n_empty": 0,
            "avg_chars": 0, "ocr_ratio": 0.0, "low_quality_pages": [],
        }

    n_native = sum(1 for p in pages
                   if (p["method"] or "").lower() in ("fitz", "pdfplumber", "docx"))
    n_ocr    = sum(1 for p in pages
                   if (p["method"] or "").lower() in ("tesseract", "ocr"))
    n_empty  = sum(1 for p in pages if p["text_len"] < 100)
    avg_chars = int(sum(p["text_len"] for p in pages) / max(len(pages), 1))
    ocr_ratio = n_ocr / max(len(pages), 1)
    low_q = [p["page_number"] for p in pages
             if 0 < p["text_len"] < 200]  # neprázdné, ale velmi krátké

    return {
        "n_pages":          len(pages),
        "n_native":         n_native,
        "n_ocr":            n_ocr,
        "n_empty":          n_empty,
        "avg_chars":        avg_chars,
        "ocr_ratio":        round(ocr_ratio, 3),
        "low_quality_pages": low_q[:20],  # max 20
    }


def _library_stats_dashboard() -> None:
    """
    Statistics across the entire library — taxa by rank, by
    stratigraphic perioddobí (z term_matches) a timeline nahrávání
    documents v čase. Zobrazeno jako samostatný expander v Knihovně,
    aby nezatěžovalo hlavní pohled.
    """
    with st.expander(t("lib_stats"), expanded=False):
        con = db()
        rank_rows = con.execute("""
            SELECT COALESCE(NULLIF(TRIM(rank_guess),''), '(unknown)') AS rank, COUNT(*) AS n
            FROM taxon_candidates
            WHERE status != 'rejected'
            GROUP BY rank ORDER BY n DESC
        """).fetchall()

        strat_rows = con.execute("""
            SELECT tm.category AS category, COUNT(DISTINCT tm.candidate_id) AS n
            FROM term_matches tm
            JOIN taxon_candidates tc ON tm.candidate_id = tc.id
            WHERE tm.term_type = 'stratigraphy' AND tc.status != 'rejected'
            GROUP BY tm.category ORDER BY n DESC
        """).fetchall()

        upload_rows = con.execute("""
            SELECT DATE(created_at) AS day, COUNT(*) AS n
            FROM documents
            WHERE created_at IS NOT NULL AND created_at != ''
            GROUP BY day ORDER BY day
        """).fetchall()

        cand_rows = con.execute("""
            SELECT DATE(created_at) AS day, COUNT(*) AS n
            FROM taxon_candidates
            WHERE created_at IS NOT NULL AND created_at != ''
            GROUP BY day ORDER BY day
        """).fetchall()
        con.close()

        st.markdown("**Taxa by rank**")
        if rank_rows:
            df_rank = pd.DataFrame([{"rank": r["rank"], "count": r["n"]} for r in rank_rows])
            st.bar_chart(df_rank.set_index("rank")["count"])
        else:
            st.caption(t("no_candidates_yet"))

        st.markdown(t("taxa_by_strat"))
        strat_named = [r for r in strat_rows if r["category"]]
        if strat_named:
            df_strat = pd.DataFrame([{"period": r["category"], "count": r["n"]} for r in strat_named])
            st.bar_chart(df_strat.set_index("period")["count"])
        else:
            st.caption(
                "Zatim zadne spocitane stratigraficke terminy. Scalculated se automaticky when a record is approved, or prepocitejte v tab ⏳🔬 Morpho/Strat.")

        st.markdown(t("timeline_docs_candidates"))
        if upload_rows or cand_rows:
            tl1, tl2 = st.columns(2)
            with tl1:
                st.caption(t("docs_per_day"))
                if upload_rows:
                    df_up = pd.DataFrame([{"den": r["day"], "Documentu": r["n"]} for r in upload_rows])
                    st.bar_chart(df_up.set_index("den")["Documentu"])
                else:
                    st.caption("—")
            with tl2:
                st.caption(t("candidates_per_day"))
                if cand_rows:
                    df_cd = pd.DataFrame([{"den": r["day"], "kandidátu": r["n"]} for r in cand_rows])
                    st.bar_chart(df_cd.set_index("den")["kandidátu"])
                else:
                    st.caption("—")
        else:
            st.caption(t("no_timeline_data"))


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 – REVIEW  (data_editor + batch ops)
# ══════════════════════════════════════════════════════════════════════════════


STATUS_OPTIONS = ["pending","approved","rejected","needs_review","low_confidence"]
RANK_OPTIONS   = ["","Species","Genus","Family","Subfamily","Tribe","Superfamily",
                  "Order","Suborder","Class","Phylum","Subgenus","Subspecies",NOT_PROVIDED]


def compute_and_save_term_matches_batch(candidate_ids: List[int]) -> int:
    """
    Batch recalculation of term_matches for a list of candidates.

    For N candidates performs onlye 4 DB dotazy (nezávisle na N):
      1. SELECT candidates (block_text)
      2. SELECT occurrence_fields (allchna pole najednou)
      3. DELETE IN (...)  — jeden statement
      4. executemany INSERT

    Oproti sekvenčnímu volání compute_and_save_term_matches_for_candidate()
    v cyklu je to pro 192 kandidátu z ~1000 DB dotazu na 4.
    """
    if not candidate_ids:
        return 0

    ph = ",".join("?" * len(candidate_ids))
    con = db()
    cand_rows = con.execute(
        f"SELECT id, block_text FROM taxon_candidates WHERE id IN ({ph})",
        candidate_ids).fetchall()
    field_rows = con.execute(
        f"SELECT candidate_id, field_name, field_value FROM occurrence_fields "
        f"WHERE candidate_id IN ({ph})",
        candidate_ids).fetchall()
    con.close()

    block_by_cand: Dict[int, str] = {r["id"]: r["block_text"] or "" for r in cand_rows}
    fields_by_cand: Dict[int, Dict[str, str]] = {}
    for fr in field_rows:
        fields_by_cand.setdefault(fr["candidate_id"], {})[fr["field_name"]] = \
            fr["field_value"] or ""

    # Prewarm cached regexes (žádná re.compile() v cyklu)
    get_term_regex("morphology")
    get_term_regex("stratigraphy")

    all_rows: List[tuple] = []
    for cid in candidate_ids:
        fmap  = fields_by_cand.get(cid, {})
        block = block_by_cand.get(cid, "")
        for ttype in ("morphology", "stratigraphy"):
            for m in compute_term_matches(fmap, block, ttype):
                all_rows.append((cid, ttype, m["term"], m["canonical"],
                                 m["category"], m["source_field"]))

    con2 = db()
    con2.execute(f"DELETE FROM term_matches WHERE candidate_id IN ({ph})",
                 candidate_ids)
    if all_rows:
        con2.executemany(
            "INSERT INTO term_matches "
            "(candidate_id, term_type, term, canonical, category, source_field) "
            "VALUES (?,?,?,?,?,?)",
            all_rows)
    con2.commit()
    con2.close()
    return len(candidate_ids)


def _batch_update_status(ids: List[int], new_status: str) -> int:
    """
    Batch-changes the status of candidates. On approval:
      - extracts/fills block_text (jen pokud missing)
      - regex-mapuje pole (jen pokud kandidát nemá žádná)
      - přepočítá term_matches DÁVKOVĚ = 4 DB dotazy celkem

    Pro 192 kandidátu nahrazuje ~1 000 sekvenčních DB dotazu za ~10.
    """
    if not ids:
        return 0

    ph = ",".join("?" * len(ids))
    con = db()
    con.execute(f"UPDATE taxon_candidates SET status=? WHERE id IN ({ph})",
                [new_status] + ids)

    approved_ids: List[int] = []

    if new_status == "approved":
        # 1. Načíst allchny kandidáty najednou
        need_rows = con.execute(
            f"SELECT id, document_id, page_start, block_text "
            f"FROM taxon_candidates WHERE id IN ({ph})",
            ids).fetchall()

        # 2. Zjistit kteří už mají pole (1 dotaz místo N)
        existing_field_cids: set = set(
            r[0] for r in con.execute(
                f"SELECT DISTINCT candidate_id FROM occurrence_fields "
                f"WHERE candidate_id IN ({ph})",
                ids).fetchall())

        # 3. Zpracovat bloky + připravit hromadný INSERT fields
        field_inserts: List[tuple] = []
        for r in need_rows:
            cid   = r["id"]
            block = r["block_text"] or ""
            approved_ids.append(cid)

            if not block:
                block = extract_block_for_candidate(
                    cid, r["document_id"], r["page_start"])
                if block:
                    con.execute(
                        "UPDATE taxon_candidates SET block_text=? WHERE id=?",
                        (block, cid))

            if block and cid not in existing_field_cids:
                for fname, fval in (map_sections_from_block(block) or {}).items():
                    if fval:
                        field_inserts.append((cid, fname, fval, "regex_auto"))

        # 4. Hromadný INSERT fields (executemany = 1 operace)
        if field_inserts:
            con.executemany(
                "INSERT OR IGNORE INTO occurrence_fields "
                "(candidate_id, field_name, field_value, method) VALUES (?,?,?,?)",
                field_inserts)

    con.commit()
    con.close()

    # Term matching se záměrně NEPROVÁDÍ zde — bylo příčinou zmrazení UI.
    # Provede se odloženě přes "Přepočítat termíny" v tab ⏳🔬 Morpho/Strat.

    return len(ids)



def _save_table_edits(edited_df: pd.DataFrame, original_df: pd.DataFrame) -> int:
    """Compares changes in the data_editor and writes them to the DB."""
    changed = 0
    approved_ids: List[int] = []
    con = db()
    for i, (_, new_row) in enumerate(edited_df.iterrows()):
        orig_row = original_df.iloc[i]
        cid = int(new_row["_id"])
        updates = {}
        if str(new_row["name"]) != str(orig_row["name"]):
            updates["taxon_name"] = str(new_row["name"])
        if str(new_row["rank"]) != str(orig_row["rank"]):
            updates["rank_guess"] = str(new_row["rank"])
        if str(new_row["status"]) != str(orig_row["status"]):
            updates["status"] = str(new_row["status"])
        if updates:
            set_clause = ", ".join(f"{k}=?" for k in updates)
            con.execute(f"UPDATE taxon_candidates SET {set_clause} WHERE id=?",
                        list(updates.values()) + [cid])
            changed += 1
            # Extrahuj blok a auto-mapuj pole při schválení
            if updates.get("status") == "approved":
                approved_ids.append(cid)
                r = con.execute(
                    "SELECT document_id, page_start, block_text FROM taxon_candidates WHERE id=?",
                    (cid,)).fetchone()
                block = r["block_text"] if r else ""
                if r and not block:
                    block = extract_block_for_candidate(cid, r["document_id"], r["page_start"])
                    if block:
                        con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                    (block, cid))
                if block:
                    existing = con.execute(
                        "SELECT COUNT(*) FROM occurrence_fields WHERE candidate_id=?",
                        (cid,)).fetchone()[0]
                    if existing == 0:
                        mapped = map_sections_from_block(block)
                        for fname, fval in mapped.items():
                            if fval:
                                con.execute(
                                    "INSERT INTO occurrence_fields "
                                    "(candidate_id,field_name,field_value,method) VALUES (?,?,?,?)",
                                    (cid, fname, fval, "regex_auto"))
    con.commit(); con.close()

    for cid in approved_ids:
        compute_and_save_term_matches_for_candidate(cid)

    return changed


def _batch_automap(ids: List[int]) -> Tuple[int, int]:
    """Batch auto-mapping of sections for approved records. Returns (ok, skip)."""
    ok = skip = 0
    for cid in ids:
        cand = get_candidate(cid)
        if not cand:
            skip += 1; continue
        block = cand["block_text"] or ""
        if not block:
            block = extract_block_for_candidate(cid, cand["document_id"], cand["page_start"])
            if block:
                con = db()
                con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?", (block, cid))
                con.commit(); con.close()
        if block:
            mapped = map_sections_from_block(block)
            if mapped:
                save_fields(cid, mapped, method="regex")
                ok += 1
            else:
                skip += 1
        else:
            skip += 1
    return ok, skip


def _batch_llm_validate(ids: List[int], s: Dict, progress_cb=None) -> Dict[str, int]:
    """Batch LLM validation. Returns per-action counts."""
    counts: Dict[str, int] = {"keep":0,"reject":0,"needs_review":0,"error":0}
    for i, cid in enumerate(ids):
        if progress_cb:
            progress_cb(i, len(ids))
        cand = get_candidate(cid)
        if not cand: continue
        try:
            prompt = (f"Candidate: {cand['heading_text']}\n"
                      f"Rank: {cand['rank_guess']}\n"
                      f"Context:\n{cand['context_after'][:600]}")
            raw = lm_chat(s, s.get("llm_validation_prompt", LLM_VALIDATION_PROMPT), prompt)
            result = lm_parse_json(raw)
            if result:
                action = result.get("action","needs_review")
                new_status = {"keep":"approved","reject":"rejected"}.get(action,"needs_review")
                con = db()
                con.execute("UPDATE taxon_candidates SET status=?, llm_json=? WHERE id=?",
                            (new_status, json.dumps(result, ensure_ascii=False), cid))
                if new_status == "approved":
                    r = con.execute(
                        "SELECT document_id, page_start, block_text FROM taxon_candidates WHERE id=?",
                        (cid,)).fetchone()
                    block = r["block_text"] if r else ""
                    if r and not block:
                        block = extract_block_for_candidate(cid, r["document_id"], r["page_start"])
                        if block:
                            con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                        (block, cid))
                    if block:
                        existing = con.execute(
                            "SELECT COUNT(*) FROM occurrence_fields WHERE candidate_id=?",
                            (cid,)).fetchone()[0]
                        if existing == 0:
                            mapped = map_sections_from_block(block)
                            for fname, fval in mapped.items():
                                if fval:
                                    con.execute(
                                        "INSERT INTO occurrence_fields "
                                        "(candidate_id,field_name,field_value,method) VALUES (?,?,?,?)",
                                        (cid, fname, fval, "regex_auto"))
                con.commit(); con.close()
                if new_status == "approved":
                    compute_and_save_term_matches_for_candidate(cid)
                counts[action if action in counts else "needs_review"] += 1
            else:
                counts["error"] += 1
        except Exception:
            counts["error"] += 1
    return counts


def _suspicion_score(r: sqlite3.Row) -> float:
    """
    Returns the suspicion score of a candidate (0 = legitimate, 1 = probable FP).
    Uses se pro Smart Sort v Review — suspicious kandidáty jdou dolu a
    dostanou ikonu ⚠️, aby na ně uživatel soustředil pozornost při odmítání.

    Faktory:
    - Name začíná stopwordem (TAXON_STOPWORDS) → +0.70
    - Nízká confidence (< 0.65) → +0.15 × (0.65 − conf) / 0.65
    - Blok kratší než 50 znaku (prázdný or stub) → +0.30
    - Name obsahuje ':' or je příliš dlouhé (> 7 slov) → +0.25
    - Krátké jméno bez roku (1 slovo, ne rod/druh) → +0.15
    - Name obsahuje číslo → +0.10 (stránkové reference, obrázky)
    """
    name = (r["taxon_name"] or "").strip()
    conf = float(r["confidence"] or 0)
    block = (r["block_text"] or "")

    score = 0.0
    first_word = name.split()[0].lower() if name.split() else ""
    if first_word in TAXON_STOPWORDS:
        score += 0.70
    if conf < 0.65:
        score += 0.15 * (0.65 - conf) / 0.65
    if len(block) < 50:
        score += 0.30
    if ":" in name or len(name.split()) > 7:
        score += 0.25
    if len(name.split()) == 1 and not re.search(r"(idae|inae|oidea)$", name, re.I):
        score += 0.15
    if re.search(r"\d", name):
        score += 0.10
    return min(score, 1.0)


def _batch_approve_threshold(min_conf: float) -> Tuple[int, int]:
    """
    Approves ALL candidates in all documents with confidence ≥ min_conf
    and status pending/low_confidence (ne already approved/rejected).
    Vrací (approved, přeskočeno).
    """
    con = db()
    candidates = con.execute(
        "SELECT id FROM taxon_candidates "
        "WHERE confidence >= ? AND status IN ('pending','low_confidence')",
        (min_conf,)).fetchall()
    con.close()
    ids = [c["id"] for c in candidates]
    if not ids:
        return 0, 0
    approved = _batch_update_status(ids, "approved")
    return approved, 0


def tab_review():
    st.markdown(f"### {t('review_header')}", unsafe_allow_html=False)
    s = get_settings()

    con = db()
    docs = con.execute("SELECT id, filename FROM documents ORDER BY created_at DESC").fetchall()
    con.close()
    if not docs:
        st.info(t("upload_doc_first"))
        return

    ALL_DOCS_LABEL = "— All documents —"
    doc_options = {ALL_DOCS_LABEL: None}
    doc_options.update({f"{d['filename']} (ID {d['id']})": d["id"] for d in docs})
    d1, d2 = st.columns([3,1])
    sel_doc = d1.selectbox("Document", list(doc_options.keys()), key="rev_doc",
                           label_visibility="collapsed")
    doc_id = doc_options[sel_doc]
    all_docs_mode = doc_id is None

    # ── Filtry ────────────────────────────────────────────────────────────────
    with st.expander(t("filters_expander"), expanded=True):
        fc1, fc2, fc3, fc4 = st.columns([2,1,1,2])
        filter_status = fc1.multiselect(
            "Status", STATUS_OPTIONS,
            default=["pending","low_confidence"], key="rev_status")
        filter_conf_min = fc2.slider("Min. score", 0.0, 1.0, 0.0, 0.05, key="rev_conf")
        filter_rank = fc3.selectbox("Rank", ["— all —"]+RANK_OPTIONS[1:], key="rev_rank")
        filter_text = fc4.text_input("🔍 Search in title", key="rev_search",
                                     placeholder="e.g. Examplius")

    # ── Načti a filtruj ───────────────────────────────────────────────────────
    con = db()
    if all_docs_mode:
        all_rows = con.execute(
            """SELECT tc.*, d.filename FROM taxon_candidates tc
               JOIN documents d ON tc.document_id = d.id
               ORDER BY d.filename, tc.page_start, tc.confidence DESC"""
        ).fetchall()
    else:
        all_rows = con.execute(
            "SELECT tc.*, d.filename FROM taxon_candidates tc "
            "JOIN documents d ON tc.document_id = d.id "
            "WHERE tc.document_id=? ORDER BY tc.page_start, tc.confidence DESC",
            (doc_id,)).fetchall()
    con.close()

    filtered = [r for r in all_rows
                if r["status"] in filter_status
                and r["confidence"] >= filter_conf_min
                and (filter_rank == "— all —" or r["rank_guess"]==filter_rank)
                and (not filter_text or filter_text.lower() in (r["taxon_name"] or "").lower())]

    if not filtered:
        st.info(t("no_candidates_filter"))
        return

    # ── Smart Sort ───────────────────────────────────────────────────────────
    # Vypočítáme suspicion score pro každého kandidáta a seřadíme:
    # legitimní hyoliti (nízká podezřelost) nahoru, FP dolu.
    sort_col, threshold_col = st.columns([2, 3])
    _sort_opts_cs = ["Smart (legitimate first)", "Score ↓", "Page ↑"]
    _sort_opts_en = ["Smart (legitimate first)", "Confidence ↓", "Page ↑"]
    _lang_now = st.session_state.get("pn_lang","cs")
    _sort_opts = _sort_opts_cs if _lang_now=="cs" else _sort_opts_en
    sort_mode = sort_col.selectbox(
        t("sort_order") if t("sort_order")!="sort_order" else "Sorting",
        _sort_opts, key="rev_sort", label_visibility="collapsed")

    suspicion: Dict[int, float] = {r["id"]: _suspicion_score(r) for r in filtered}

    if sort_mode in (_sort_opts[0], "Smart (legitimate first)", "Smart (legitimate first)"):
        filtered = sorted(filtered, key=lambda r: (suspicion[r["id"]], -r["confidence"]))
    elif sort_mode in (_sort_opts[1], "Confidence ↓", "Score ↓"):
        filtered = sorted(filtered, key=lambda r: -r["confidence"])
    else:
        filtered = sorted(filtered, key=lambda r: r["page_start"])

    # ── Batch approve by threshold (allchny Documenty) ───────────────────────
    with threshold_col:
        tc1, tc2, tc3 = st.columns([2, 1, 2])
        thr = tc1.slider(
            "Threshold", 0.50, 1.00,
            float(s.get("review_batch_threshold", 0.85)), 0.05,
            key="rev_threshold", label_visibility="collapsed",
            format="%.2f")
        s["review_batch_threshold"] = thr
        # Spočítat z DB přímo kolik kandidátu by bylo approved
        con_t = db()
        n_above = con_t.execute(
            "SELECT COUNT(*) FROM taxon_candidates "
            "WHERE confidence >= ? AND status IN ('pending','low_confidence')",
            (thr,)).fetchone()[0]
        con_t.close()
        if tc2.button(f"✅ ≥{thr:.2f}", key="rev_threshold_btn", type="primary",
                      help=f"Approves {n_above} candidates across VSEMI Documenty s conf ≥ {thr:.2f}"):
            approved_n, _ = _batch_approve_threshold(thr)
            st.success(tt(f"Approved {approved_n} candidates across the library (conf ≥ {thr:.2f}).", f"Approved {approved_n} candidates library-wide (conf ≥ {thr:.2f})."))
            st.rerun()
        tc3.caption(f"→ {n_above} candidates across all docs.")

    # ── Souhrn (kompaktní jeden řádek) ──────────────────────────────────────
    _comp_data = []
    if filtered:
        _cids = [r["id"] for r in filtered]
        _ph   = ",".join("?" * len(_cids))
        _ccon = db()
        _all_flds = _ccon.execute(
            f"SELECT candidate_id, field_name, field_value FROM occurrence_fields "
            f"WHERE candidate_id IN ({_ph})", _cids).fetchall()
        _ccon.close()
        _flds_map: Dict[int, Dict[str, str]] = {}
        for _ff in _all_flds:
            _flds_map.setdefault(_ff["candidate_id"], {})[_ff["field_name"]] = _ff["field_value"]
        for _r in filtered:
            _cs, _ = _completeness_check(_r["rank_guess"] or "", _flds_map.get(_r["id"], {}))
            _comp_data.append(_cs)
    _avg_comp = round(sum(_comp_data)/len(_comp_data), 2) if _comp_data else 0.0
    _n_hi  = sum(1 for r in filtered if r["confidence"] >= 0.80)
    _n_mid = sum(1 for r in filtered if 0.60 <= r["confidence"] < 0.80)
    _n_lo  = sum(1 for r in filtered if r["confidence"] < 0.60)
    _n_sus = sum(1 for r in filtered if suspicion[r["id"]] >= 0.50)
    _dok_info = (f"dok: {len(set(r['document_id'] for r in filtered))}" if all_docs_mode
                 else f"str.max: {max((r['page_start'] for r in filtered), default=0)}")
    st.markdown(
        f"<div style='font-size:0.78rem;padding:3px 0;opacity:0.85'>"
        f"Zobrazeno <b>{len(filtered)}</b> &nbsp;│&nbsp; "
        f"🟢 ≥0.80: <b>{_n_hi}</b> &nbsp;│&nbsp; "
        f"🟡 0.6–0.8: <b>{_n_mid}</b> &nbsp;│&nbsp; "
        f"🔴 &lt;0.60: <b>{_n_lo}</b> &nbsp;│&nbsp; "
        f"⚠️ suspicious: <b>{_n_sus}</b> &nbsp;│&nbsp; "
        f"avg completeness: <b>{_avg_comp*100:.0f}%</b> &nbsp;│&nbsp; {_dok_info}"
        f"</div>", unsafe_allow_html=True)
    # ── Batch překlad celého Documentu ───────────────────────────────────
    if s.get("llm_enabled"):
        with st.expander(t("batch_translate_expander"), expanded=False):
            st.caption(
                "Prelozi allchna vyplnena fields schvalenych/pending records z selectedho Documentu do anglictiny pomoci LLM.")
            _bt_con = db()
            _bt_docs = _bt_con.execute(
                "SELECT DISTINCT d.id, d.filename, d.lang "
                "FROM documents d "
                "JOIN taxon_candidates tc ON tc.document_id=d.id "
                "WHERE tc.status IN ('approved','pending') "
                "AND (d.lang IS NULL OR d.lang NOT IN "
                "('en','eng','english')) "
                "ORDER BY d.filename"
            ).fetchall()
            _bt_con.close()
            if not _bt_docs:
                st.info(t("no_non_en_docs"))
            else:
                _bt_doc_opts = {
                    f"{d['filename']} [{d['lang'] or '?'}]": d
                    for d in _bt_docs}
                _bt_sel = st.selectbox(
                    "Document:", list(_bt_doc_opts.keys()),
                    key="batch_tr_doc_sel")
                _bt_doc = _bt_doc_opts[_bt_sel]
                _bt_fname = _bt_doc["filename"]
                _bt_lang  = _bt_doc["lang"] or ""
                if st.button(
                    f"🌐 Translate all from '{_bt_fname}' [{_bt_lang}→EN]",
                    key="batch_tr_run", type="primary"):
                    _bt_con2 = db()
                    _bt_cands = _bt_con2.execute(
                        "SELECT id FROM taxon_candidates "
                        "WHERE document_id=? AND status IN "
                        "('approved','pending')",
                        (_bt_doc["id"],)).fetchall()
                    _bt_con2.close()
                    _bt_prog  = st.progress(0, "Translating…")
                    _bt_total = len(_bt_cands)
                    _bt_done  = 0
                    _bt_tr    = 0
                    for _bt_c in _bt_cands:
                        _bt_n = auto_translate_candidate_fields(
                            _bt_c["id"], _bt_lang, s, force=True)
                        _bt_tr   += _bt_n
                        _bt_done += 1
                        _bt_prog.progress(
                            int(_bt_done / max(_bt_total, 1) * 100),
                            f"{_bt_done}/{_bt_total} ({_bt_tr} fields)…")
                    _bt_prog.empty()
                    st.success(
                        f"Done: translated {_bt_tr} fields "
                        f"v {_bt_done} records.")
                    st.rerun()


    # ── Klávesové zkratky (A/R) ───────────────────────────────────────────────
    _inject_keyboard_shortcuts("review", "Approve (A)", "Reject (R)")
    # ── Přehled napříč Documenty (jen v režimu "Allchny Documenty") ───────────
    if all_docs_mode:
        st.markdown(f"<div style='font-size:0.82rem;margin:3px 0;opacity:0.8'>{tt('Overview by document:', 'Overview by document:')}</div>", unsafe_allow_html=True)

        by_doc: Dict[int, List] = {}
        doc_names: Dict[int, str] = {}
        for r in filtered:
            by_doc.setdefault(r["document_id"], []).append(r)
            doc_names[r["document_id"]] = r["filename"]

        # Tlačítko schválit napříč VŠEMI filtrovanými Documenty najednou
        gc1, gc2 = st.columns([1,3])
        if gc1.button(f"✅ Approve all across all ({len(filtered)})",
                      type="primary", key="rev_approve_all_global"):
            with st.spinner(tt(f"Approving {len(filtered)} candidates…", f"Approving {len(filtered)} candidates…")):
                n = _batch_update_status([r["id"] for r in filtered], "approved")
            st.toast(tt(f"✅ Approved {n} candidates across {len(by_doc)} Documenty.", f"✅ Approved {n} candidates across {len(by_doc)} documents."), icon="✅")
            st.rerun()
        gc2.caption("Approves allchny kandidaty zobrazene v prehledu below, ve allch documents.")

        for did, rows in sorted(by_doc.items(), key=lambda kv: doc_names[kv[0]]):
            dc1, dc2, dc3 = st.columns([3,1,1])
            dc1.markdown(f"📄 **{doc_names[did]}** — {len(rows)} candidate(s)")
            if dc2.button("✅ Approve all", key=f"rev_approve_doc_{did}"):
                with st.spinner(tt(f"Approving {len(rows)} records…", f"Approving {len(rows)} records…")):
                    n = _batch_update_status([r["id"] for r in rows], "approved")
                st.toast(tt(f"✅ Approved {n} · {doc_names[did]}", f"✅ Approved {n} · {doc_names[did]}"), icon="✅")
                st.rerun()
            if dc3.button("❌ Reject all", key=f"rev_reject_doc_{did}"):
                n = _batch_update_status([r["id"] for r in rows], "rejected")
                st.toast(tt(f"❌ Rejected {n} · {doc_names[did]}", f"❌ Rejected {n} · {doc_names[did]}"), icon="❌")
                st.rerun()

    # ── Hromadné akce na filtrovaných ───────────────────────────────────────
    st.markdown(f"<div style='margin-top:4px;font-size:0.8rem;opacity:0.7'>{tt('Bulk action on filtered:', 'Bulk actions on filtered:')}</div>", unsafe_allow_html=True)
    ba1,ba2,ba3,ba4,ba5 = st.columns(5)
    if ba1.button("✅ Approve all", help="Schvali all filtered"):
        with st.spinner(tt(f"Approving {len(filtered)} records…", f"Approving {len(filtered)} records…")):
            n = _batch_update_status([r["id"] for r in filtered], "approved")
        st.toast(tt(f"✅ Approved {n}", f"✅ Approved {n}"), icon="✅"); st.rerun()
    if ba2.button("✅ ≥ 0.80", help="Approves jen Score ≥ 0.80"):
        ids = [r["id"] for r in filtered if r["confidence"]>=0.80]
        with st.spinner(tt(f"Approving {len(ids)} records se Score ≥ 0.80…", f"Approving {len(ids)} records with score ≥ 0.80…")):
            n = _batch_update_status(ids, "approved")
        st.toast(tt(f"✅ Approved {n}", f"✅ Approved {n}"), icon="✅"); st.rerun()
    if ba3.button("❌ Reject", help="Reject all filtered"):
        n = _batch_update_status([r["id"] for r in filtered], "rejected")
        st.toast(tt(f"❌ Rejected {n}", f"❌ Rejected {n}"), icon="❌"); st.rerun()
    if ba4.button("🔄 Auto-map", help="Auto-mapuje sekce pro schvalene z filtru"):
        ids = [r["id"] for r in filtered if r["status"]=="approved"]
        if ids:
            with st.spinner(tt(f"Auto-mapuji {len(ids)} records…", f"Auto-mapping {len(ids)} records…")):
                ok, skip = _batch_automap(ids)
            st.success(tt(f"Mapped {ok}, skipped {skip}", f"Mapped {ok}, skipped {skip}")); st.rerun()
        else:
            st.warning(t("no_approved_in_filter"))
    if ba5.button("🤖 LLM batch", disabled=not s.get("llm_enabled"),
                   help="Runs LLM validation na all filtered (may lasts long)"):
        ids = [r["id"] for r in filtered]
        prog = st.progress(0, f"LLM validace 0/{len(ids)}…")
        def _upd(i, total): prog.progress(int(i/total*100), f"LLM {i+1}/{total}…")
        with st.spinner("LLM validace…"):
            counts = _batch_llm_validate(ids, s, _upd)
        prog.empty()
        st.success(f"LLM: keep={counts['keep']}, reject={counts['reject']}, "
                   f"review={counts['needs_review']}, chyby={counts['error']}")
        st.rerun()

    # ── Data editor tabulka ───────────────────────────────────────────────────
    st.markdown(f"<div style='font-size:0.8rem;opacity:0.7;margin:2px 0'>{tt('Table of candidates — edit in cells, then Save', 'Candidate table — edit cells directly, then Save')}</div>", unsafe_allow_html=True)

    row_dicts = []
    for r in filtered:
        susp = suspicion.get(r["id"], 0.0)
        d = {
            "_id":       r["id"],
            "✓":         False,
            "⚠️":        "⚠️" if susp >= 0.50 else "",
            "Str.":      r["page_start"],
            "Score":     round(r["confidence"], 3),
            "name":     r["taxon_name"] or "",
            "rank":      r["rank_guess"] or "",
            "status":    r["status"],
            "Nadpis":    (r["heading_text"] or "")[:70],
        }
        if all_docs_mode:
            d["Document"] = r["filename"]
        row_dicts.append(d)
    df_orig = pd.DataFrame(row_dicts)

    df_show = df_orig.drop(columns=["_id"])
    col_order = (["✓","⚠️","Document","Str.","Score","name","rank","status","Nadpis"]
                if all_docs_mode else
                ["✓","⚠️","Str.","Score","name","rank","status","Nadpis"])
    df_show = df_show[col_order]

    col_config = {
        "✓":      st.column_config.CheckboxColumn("✓",    width=40),
        "⚠️":     st.column_config.TextColumn("⚠️",      width=40, disabled=True, help="Podezrely kandidat — suspicion score ≥ 0.50 (mimo systematiku, kratky name, ref-like)"),
        "Document": st.column_config.TextColumn("Document", width=180, disabled=True),
        "Str.":   st.column_config.NumberColumn("Str.",   width=60,  disabled=True),
        "Score":  st.column_config.NumberColumn("Score",  width=70,  disabled=True, format="%.3f"),
        "name":  st.column_config.TextColumn("Name taxon", width=240),
        "rank":   st.column_config.SelectboxColumn("Rank", width=120, options=RANK_OPTIONS),
        "status": st.column_config.SelectboxColumn("Status", width=130, options=STATUS_OPTIONS),
        "Nadpis": st.column_config.TextColumn("Nadpis (info)", width=280, disabled=True),
    }

    edited = st.data_editor(
        df_show,
        column_config=col_config,
        hide_index=True,
        use_container_width=True,
        num_rows="fixed",
        key="review_table",
    )

    # Přidat _id zpět pro saveení
    edited_with_id = edited.copy()
    edited_with_id["_id"] = df_orig["_id"].values

    # ── Akce na tabulce ───────────────────────────────────────────────────────
    ea1 = st.columns(1)[0]

    if ea1.button("💾 Save table changes", type="primary"):
        n = _save_table_edits(edited_with_id, df_orig)
        if n:
            st.success(tt(f"Saved {n} changes.", f"Saved {n} changes.")); st.rerun()
        else:
            st.info(t("no_changes_to_save"))

    # Akce na zaškrtnutých
    selected_ids = [
        int(df_orig.iloc[i]["_id"])
        for i, row in edited.iterrows()
        if row.get("✓", False)
    ]
    n_sel = len(selected_ids)

    if n_sel:
        st.markdown(tt(f"**Selected records: {n_sel}**", f"**Selected records: {n_sel}**"))
        sa1,sa2,sa3,sa4,sa5 = st.columns(5)
        if sa1.button(f"✅ Approve ({n_sel})", key="sel_approve"):
            with st.spinner(tt(f"Approving {n_sel} records…", f"Approving {n_sel} records…")):
                _batch_update_status(selected_ids, "approved")
            st.toast(tt(f"✅ Approved {n_sel}", f"✅ Approved {n_sel}"), icon="✅"); st.rerun()
        if sa2.button(f"❌ Reject ({n_sel})", key="sel_reject"):
            _batch_update_status(selected_ids, "rejected")
            st.toast(tt(f"❌ Rejected {n_sel}", f"❌ Rejected {n_sel}"), icon="❌"); st.rerun()
        if sa3.button(f"🔍 Needs review ({n_sel})", key="sel_review"):
            _batch_update_status(selected_ids, "needs_review")
            st.toast(tt(f"🔍 Marked {n_sel}", f"🔍 Marked {n_sel}"), icon="🔍"); st.rerun()
        if sa4.button(f"🔄 Auto-map ({n_sel})", key="sel_automap"):
            appr = [i for i in selected_ids
                    if any(r["id"]==i and r["status"]=="approved" for r in filtered)]
            if appr:
                with st.spinner("Auto-mapuji…"):
                    ok, skip = _batch_automap(appr)
                st.success(tt(f"Mapped {ok}, skipped {skip}", f"Mapped {ok}, skipped {skip}")); st.rerun()
            else:
                st.warning(t("no_approved_in_sel"))
        if sa5.button(f"🤖 LLM ({n_sel})", key="sel_llm",
                      disabled=not s.get("llm_enabled")):
            prog2 = st.progress(0, "LLM…")
            def _upd2(i, total): prog2.progress(int(i/total*100), f"LLM {i+1}/{total}…")
            counts = _batch_llm_validate(selected_ids, s, _upd2)
            prog2.empty()
            st.success(f"LLM: keep={counts['keep']}, reject={counts['reject']}, "
                       f"review={counts['needs_review']}, err={counts['error']}")
            st.rerun()

    # ── Detail kandidáta — navigace Prev/Noxt + klávesy ← → ────────────────
    # review_detail_idx: index do filteredanního seznamu (přežívá rerun)
    # Pokud je zaškrtnut přesně jeden checkbox → přejít na něj.
    if "review_detail_idx" not in st.session_state:
        st.session_state["review_detail_idx"] = 0

    if len(selected_ids) == 1:
        # Zaškrtnutý jeden checkbox → synchronizovat detail na ten záznam
        _sel_pos = next((i for i, r in enumerate(filtered) if r["id"] == selected_ids[0]), 0)
        if st.session_state["review_detail_idx"] != _sel_pos:
            st.session_state["review_detail_idx"] = _sel_pos

    _det_idx = st.session_state["review_detail_idx"]
    if _det_idx >= len(filtered):
        _det_idx = max(0, len(filtered) - 1)
        st.session_state["review_detail_idx"] = _det_idx

    detail_id = filtered[_det_idx]["id"] if filtered else None

    # ── Navigační lišta Prev/Noxt ────────────────────────────────────────────
    _nav_c1, _nav_c2, _nav_c3, _nav_c4 = st.columns([1, 1, 4, 1])
    if _nav_c1.button(tt("◀ Previous", "◀ Previous"), key="rev_det_prev",
                      disabled=(_det_idx == 0),
                      help=tt("Previous candidate (key ←)", "Previous candidate (key ←)")):
        st.session_state["review_detail_idx"] = max(0, _det_idx - 1)
        st.rerun()
    if _nav_c2.button(tt("Noxt ▶", "Noxt ▶"), key="rev_det_next",
                      disabled=(_det_idx >= len(filtered) - 1),
                      help=tt("Noxt candidate (key →)", "Noxt candidate (key →)")):
        st.session_state["review_detail_idx"] = min(len(filtered) - 1, _det_idx + 1)
        st.rerun()
    _nav_c3.markdown(
        f"<div style='padding-top:6px;font-size:0.82rem;opacity:0.85'>"
        f"{tt(f'Candidate <b>{_det_idx + 1}</b> / {len(filtered)}', f'Candidate <b>{_det_idx + 1}</b> / {len(filtered)}')}"
        f"</div>", unsafe_allow_html=True)
    # Skokový výběr — číslo záznamu v sérii
    _jump = _nav_c4.number_input(
        "→#", min_value=1, max_value=max(len(filtered), 1),
        value=_det_idx + 1, step=1,
        key="rev_det_jump", label_visibility="collapsed",
        help=tt("Go to record n.", "Jump to record #"))
    if int(_jump) - 1 != _det_idx:
        st.session_state["review_detail_idx"] = int(_jump) - 1
        st.rerun()

    # ── Keyboard navigace ← → v detail panelu ────────────────────────────────
    st.markdown("""<script> (function() { const _navKey = '_rev_nav_injected'; if (window[_navKey]) return; window[_navKey] = true; document.addEventListener('keydown', function(e) { if (['INPUT','TEXTAREA','SELECT'].includes(document.activeElement.tagName)) return; if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') { const btns = document.querySelectorAll('button'); const label = e.key === 'ArrowLeft' ? '◀ Previous': 'Noxt ▶'; for (const b of btns) { if (b.innerText.trim().startsWith(e.key === 'ArrowLeft' ? '◀': 'Noxt') || b.innerText.trim() === label) { b.click(); break; } } } }); })(); </script>""", unsafe_allow_html=True)

    with st.expander(t("candidate_detail_expander"), expanded=True):
        cand_detail = get_candidate(int(detail_id)) if detail_id else None
        if cand_detail:
            # Hlavička — rank badge s barvou podle species/genus/family
            _conf_d = cand_detail["confidence"] or 0
            _stat_d = cand_detail["status"] or ""
            _stat_icon = {"approved":"✅","rejected":"❌","pending":"⏳",
                          "low_confidence":"🔵","needs_review":"🔍"}.get(_stat_d,"❓")
            _rank_d = (cand_detail.get("rank_guess") or "").lower()
            _rank_color = {
                "species":"#1d4ed8","subspecies":"#2563eb",
                "genus":"#7c3aed","subgenus":"#8b5cf6",
                "family":"#b45309","subfamily":"#d97706",
                "order":"#065f46","class":"#1e3a5f","phylum":"#374151",
            }.get(_rank_d, "#475569")
            _rank_badge = (
                f"<span style='background:{_rank_color};color:#fff;"
                f"border-radius:3px;padding:1px 6px;font-size:0.78rem'>"
                f"{cand_detail['rank_guess'] or '?'}</span>"
            )
            # Rank-aware TYPE pole varování v detailu
            _det_fields_pre = get_candidate_fields(int(detail_id))
            _rk_d_norm = _RANK_ALIASES.get(_rank_d, _rank_d if _rank_d in _REQUIRED_FIELDS_BY_RANK else "")
            _is_sp_grade = _rk_d_norm in ("species", "subspecies")
            _is_hi_grade = _rk_d_norm in ("genus", "family", "order", "class", "phylum")
            _type_warn = ""
            if _is_sp_grade and (_det_fields_pre.get("TYPE TAXON","") or "") not in ("", NOT_PROVIDED):
                _type_warn = "⚠️ TYPE TAXON at species — check (should be TYPE SPECIMENS)"
            elif _is_hi_grade:
                _ts_pre = _det_fields_pre.get("TYPE SPECIMENS","") or ""
                if _ts_pre and _ts_pre != NOT_PROVIDED and not any(k in _ts_pre.lower() for k in TYPE_SPECIMEN_KEYWORDS):
                    _type_warn = "⚠️ TYPE SPECIMENS at genus/family — check (should be TYPE TAXON)"

            st.markdown(
                f"**{cand_detail['taxon_name']}** {_rank_badge} "
                f"— Score **{_conf_d:.3f}** {_stat_icon} {_stat_d} "
                f"| str. {cand_detail['page_start']} "
                f"| dok. *{cand_detail.get('filename','?')}*",
                unsafe_allow_html=True)
            if _type_warn:
                st.warning(_type_warn)

            dd1, dd2 = st.columns([3, 2])
            with dd1:
                if cand_detail["heading_text"]:
                    st.markdown(tt(f"**Nadpis:** `{cand_detail['heading_text']}`", f"**Heading:** `{cand_detail['heading_text']}`"))
                st.caption(t("context_before_caption"))
                st.text(cand_detail["context_before"][:600] if cand_detail["context_before"] else "–")
                st.caption(t("block_text_caption"))
                st.text(cand_detail["block_text"][:600] if cand_detail["block_text"] else "–")
                st.caption(t("context_after_caption"))
                st.text(cand_detail["context_after"][:400] if cand_detail["context_after"] else "–")

            with dd2:
                # Vyplněná pole — zvýrazni TYPE TAXON/SPECIMENS podle ranku
                _det_fields = _det_fields_pre
                if _det_fields:
                    st.caption(t("extracted_fields_caption"))
                    for _fn, _fv in sorted(_det_fields.items()):
                        if _fv and _fv != NOT_PROVIDED:
                            # Barevné označení problematických fields
                            _fld_warn = (
                                (_is_sp_grade and _fn == "TYPE TAXON") or
                                (_is_hi_grade and _fn == "TYPE SPECIMENS" and
                                 not any(k in (_fv or "").lower() for k in TYPE_SPECIMEN_KEYWORDS))
                            )
                            _lbl = f"{'⚠️' if _fld_warn else ''} {_fn}"
                            st.text_input(_lbl, value=_fv, key=f"det_fld_{detail_id}_{_fn}",
                                          disabled=True, label_visibility="visible")
                # Kompletnost
                _det_comp, _det_miss = _completeness_check(
                    cand_detail["rank_guess"] or "", _det_fields)
                _det_pct = int(_det_comp * 100)
                _col = "#10b981" if _det_pct>=70 else "#f59e0b" if _det_pct>=40 else "#ef4444"
                st.markdown(
                    f"<span style='color:{_col}'>**Completeness: {_det_pct}%**</span>"
                    + (f" — missing: {', '.join(_det_miss)}" if _det_miss else ""),
                    unsafe_allow_html=True)
                # Debug scoring
                try:
                    dbg = json.loads(cand_detail["debug_json"] or "{}")
                    st.caption("Scoring debug:")
                    st.json(dbg, expanded=False)
                except Exception:
                    pass
                llm_j = json.loads(cand_detail["llm_json"] or "{}") if cand_detail["llm_json"] else {}
                if llm_j:
                    action = llm_j.get("action","?")
                    color = {"keep":"success","reject":"error","needs_review":"warning"}.get(action,"info")
                    getattr(st, color)(
                        f"LLM: **{action}** [{llm_j.get('confidence','?')}] "
                        f"_{llm_j.get('reason','')}_")

            # Rychlé akce + navigace v jedné řadě
            _qa1, _qa2, _qa3, _qa4, _qa5 = st.columns([1, 1, 1, 1, 1])
            if _qa1.button("✅ Approve", key=f"det_appr_{detail_id}", type="primary"):
                _batch_update_status([detail_id], "approved")
                # Automatický posun na další po schválení
                if _det_idx < len(filtered) - 1:
                    st.session_state["review_detail_idx"] = _det_idx + 1
                st.rerun()
            if _qa2.button("❌ Reject", key=f"det_rej_{detail_id}"):
                _batch_update_status([detail_id], "rejected")
                if _det_idx < len(filtered) - 1:
                    st.session_state["review_detail_idx"] = _det_idx + 1
                st.rerun()
            if _qa3.button("🔍 Needs review", key=f"det_rev_{detail_id}"):
                _batch_update_status([detail_id], "needs_review"); st.rerun()
            if _qa4.button("◀", key=f"det_prev_{detail_id}",
                           disabled=(_det_idx == 0), help="Previous"):
                st.session_state["review_detail_idx"] = max(0, _det_idx - 1); st.rerun()
            if _qa5.button("▶", key=f"det_next_{detail_id}",
                           disabled=(_det_idx >= len(filtered) - 1), help="Noxt"):
                st.session_state["review_detail_idx"] = min(len(filtered)-1, _det_idx+1); st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 – EDITOR
# ══════════════════════════════════════════════════════════════════════════════

FIELD_GROUPS: Dict[str, List[str]] = {
    # TYPE TAXON (rod/Ďledě → typový druh/rod) patří do Identita
    "🏷️ Identita":          ["TAXONOMIC PLACEMENT","TAXON","NOMENCLATURAL ACTS",
                              "TYPE TAXON",
                              "INCLUDED TAXONS"],
    "📋 Nomenklatura":      ["AUTHOR","SYNONYMY",
                              "OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS"],
    # TYPE SPECIMENS (druh/poddruh → holotyp, paratypus…) patří do Typy & materiál
    "🔬 Types & material":   ["TYPE SPECIMENS",
                              "TYPE MATERIAL","MATERIAL EXAMINED"],
    "📍 Locality & strat.": ["LOCALITY","STRATIGRAPHY","OCCURRENCE","YEAR_OF_PUBLICATION"],
    "📝 Popis":              ["DIAGNOSIS","DESCRIPTION","SIZE","ETYMOLOGY"],
    "💬 Misc":               ["REMARKS","FIGURES","REFERENCE"],
}
# Prioritní pořadí fields pro ruzné ranky (používá ordered_field_groups_for_rank)
_SPECIES_GRADE_PRIORITY_FIELDS  = ["TYPE SPECIMENS","DIAGNOSIS","LOCALITY",
                                    "STRATIGRAPHY","DESCRIPTION","SYNONYMY","FIGURES"]
_GENUS_GRADE_PRIORITY_FIELDS    = ["TYPE TAXON","DIAGNOSIS","INCLUDED TAXONS",
                                    "OCCURRENCE","DESCRIPTION","SYNONYMY"]
_FAMILY_GRADE_PRIORITY_FIELDS   = ["TYPE TAXON","DIAGNOSIS","INCLUDED TAXONS",
                                    "OCCURRENCE","REMARKS"]


def get_field_priority_map() -> Dict[str, float]:
    """
    Returns {target_field: max_priority} from section_schema.tsv (column 'priority').
    Používá se k seřazení fields a skupin fields v Editoru — vyšší priorita
    znamená duležitější/spolehlivější pole, zobrazí se výš.
    """
    schema = load_schema()
    if "priority" not in schema.columns:
        return {}
    pr = schema.copy()
    pr["priority"] = pd.to_numeric(pr["priority"], errors="coerce").fillna(0)
    return pr.groupby("target_field")["priority"].max().to_dict()


def ordered_field_groups() -> Dict[str, List[str]]:
    """
    Returns FIELD_GROUPS with fields SORTED by schema priority (descending)
    and groupsinami seřazenými podle nejvyšší priority pole, které obsahují.
    Logické seskupení (Identita / Nomenklatura / …) zustává zachováno —
    mění se jen POŘADÍ uvnitř skupin a pořadí samotných skupin.
    """
    prio = get_field_priority_map()
    result: Dict[str, List[str]] = {}
    group_max_prio: Dict[str, float] = {}
    for gname, fields in FIELD_GROUPS.items():
        sorted_fields = sorted(fields, key=lambda f: -prio.get(f, 0))
        result[gname] = sorted_fields
        group_max_prio[gname] = max((prio.get(f, 0) for f in fields), default=0)
    # Seřadit skupiny podle nejvyšší priority pole uvnitř (sestupně)
    ordered_names = sorted(result.keys(), key=lambda g: -group_max_prio[g])
    return {g: result[g] for g in ordered_names}


def ordered_field_groups_for_rank(rank: str = "") -> Dict[str, List[str]]:
    """
    Returns FIELD_GROUPS sorted with respect to taxon rank:
    - species/subspecies: TYPE SPECIMENS tab se zobrazí první ve skupině Typy,
      TYPE TAXON je skryt/přesunut dozadu (nepotřebný pro druhy).
    - genus/family/…: TYPE TAXON tab se zobrazí první, TYPE SPECIMENS dozadu.
    - Jinak: použije standardní ordered_field_groups().
    """
    base = ordered_field_groups()
    if not rank:
        return base

    _rn = (rank or "").lower().split()[0]
    if _rn in _REQUIRED_FIELDS_BY_RANK:
        _rk = _rn
    else:
        _rk = _RANK_ALIASES.get(_rn, "")

    result: Dict[str, List[str]] = {}
    for gname, fields in base.items():
        ordered = list(fields)
        if _rk in ("species", "subspecies"):
            # Pro druhy: TYPE SPECIMENS na prvním místě v Typy,
            # TYPE TAXON přesunout na poslední v Identita (málo relevantní)
            if "TYPE SPECIMENS" in ordered:
                ordered = ["TYPE SPECIMENS"] + [f for f in ordered if f != "TYPE SPECIMENS"]
            if "TYPE TAXON" in ordered:
                ordered = [f for f in ordered if f != "TYPE TAXON"] + ["TYPE TAXON"]
        elif _rk in ("genus", "family", "order", "class", "phylum"):
            # Pro rody/čeledi: TYPE TAXON na prvním místě v Identita
            if "TYPE TAXON" in ordered:
                ordered = ["TYPE TAXON"] + [f for f in ordered if f != "TYPE TAXON"]
            # TYPE SPECIMENS na konec Typy (málo relevantní jako heading)
            if "TYPE SPECIMENS" in ordered:
                ordered = [f for f in ordered if f != "TYPE SPECIMENS"] + ["TYPE SPECIMENS"]
        result[gname] = ordered
    return result



def _unit_role_badges(unit) -> str:
    txt = (unit.text or "").strip()
    roles = []
    zf = unit.zone_flags or ""
    if "caption" in zf: roles.append("caption")
    if "references" in zf: roles.append("references")
    if "systematic" in zf: roles.append("systematic")
    if _TYPE_MATERIAL_SIGNAL_RE.search(txt): roles.append("type_material")
    if _TYPE_TAXON_LINE_RE.search(txt): roles.append("type_taxon")
    if _INCLUDED_LINE_RE.search(txt): roles.append("included_taxa")
    try:
        nm, rk, pat = _match_name(txt)
        if nm: roles.append(f"taxon_heading:{rk}")
    except Exception:
        pass
    try:
        if map_sections_from_block(txt): roles.append("field_label")
    except Exception:
        pass
    return ", ".join(dict.fromkeys(roles))


def tab_block_editor():
    st.markdown(tt("### 🧩 Block Editor", "### 🧩 Block Editor"))

    with st.expander(t("block_editor_help"), expanded=False):
        _be_cs = """
**Block Editor is used for manual inspection and correction of which text belongs to a given record.**

PaleoN při indexaci automaticky přiřadí každému nalezenému taxonu blok textu (= odstavce z PDF/DOCX,
ze kterých se poté extrahují pole jako Description, Occurrence, Type specimens atd.).
Pokud automatická detekce hranice bloku selhala or záznam zasahuje na více stránek,
Block Editor umožní to opravit.

---
**Tři panely:**

| Panel | Co dělá |
|---|---|
| **① Seznam records** (vlevo) | Přehled allch nalezených taxa v Documentu. Vyber záznam kliknutím v seznamu nahoře. Zobrazuje Score, rank a stav. |
| **② Hranice v textu** (uprostřed) | Text Documentu je rozdělen na číslované „jednotky" (odstavce). Uprav čísla **od** a **do** tak, aby zahrnovaly celý text záznamu — pak stiskni 💾 Save. |
| **③ Blok a pole** (vpravo) | Zobrazuje aktuální text bloku a z něj extrahovaná pole. Blok lze ručně přepsat or spustit automatické mapování fields. |

---
**Typický postup:**
1. Vyber Document a záznam nahoře
2. Ve středním panelu zkontroluj, zda jsou označeny správné odstavce (✓ = ano, prázdno = ne)
3. Uprav čísla **od / do** a stiskni 💾 Save hranice
4. V pravém panelu stiskni **🔄 Re-run automap** pro přecount fields
"""
        _be_en = """
**Block Editor is used for manually checking and fixing which text belongs to a given record.**

During indexing, PaleoN automatically assigns each found taxon a block of text (= paragraphs from PDF/DOCX,
from which fields like Description, Occurrence, Type specimens etc. are extracted).
If the automatic block boundary detection failed or the record spans multiple pages,
Block Editor lets you fix it.

---
**Three panels:**

| Panel | What it does |
|---|---|
| **① Record list** (left) | Overview of all taxa found in the document. Select a record by clicking in the list above. Shows score, rank and status. |
| **② Text boundaries** (centre) | The document text is split into numbered "units" (paragraphs). Adjust **from** and **to** numbers to cover the full record text — then press 💾 Save. |
| **③ Block and fields** (right) | Shows the current block text and extracted fields. The block can be manually overwritten or auto-mapping can be triggered. |

---
**Typical workflow:**
1. Select document and record above
2. In the middle panel check that the correct paragraphs are marked (✓ = yes, empty = no)
3. Adjust **from / to** numbers and press 💾 Save boundaries
4. In the right panel press **🔄 Re-run automap** to recompute fields
"""
        st.markdown(tt(_be_cs, _be_en))

    con = db()
    docs = con.execute("SELECT id, filename FROM documents ORDER BY created_at DESC").fetchall()
    con.close()
    if not docs:
        st.info("Nojprve nahrajte Document.")
        return

    doc_labels = {f"{d['filename']} (ID {d['id']})": d['id'] for d in docs}
    dcol, scol = st.columns([3, 2])
    doc_label = dcol.selectbox("Document", list(doc_labels.keys()), key="be_doc")
    doc_id = doc_labels[doc_label]
    statuses = scol.multiselect("Status", STATUS_OPTIONS,
                                default=["pending", "low_confidence", "needs_review", "approved"],
                                key="be_status")
    if not statuses:
        st.warning(t("select_at_least_one_status"))
        return

    ph = ",".join("?"*len(statuses))
    con = db()
    rows = con.execute(
        f"SELECT * FROM taxon_candidates WHERE document_id=? AND status IN ({ph}) ORDER BY page_start, confidence DESC",
        [doc_id] + statuses).fetchall()
    con.close()
    if not rows:
        st.info(t("no_candidates_filter"))
        return

    cand_labels = {f"#{r['id']} · str. {r['page_start']} · {r['taxon_name']} · {r['rank_guess']} · {r['status']}": r['id'] for r in rows}
    cand_label = st.selectbox(t("candidate_block_label"), list(cand_labels.keys()), key="be_cand")
    cid = int(cand_labels[cand_label])
    cand = get_candidate(cid)
    if not cand:
        st.error(t("candidate_not_found"))
        return

    active_block = get_active_block_text(cand)
    if not active_block:
        active_block = ensure_candidate_block(cid)
        cand = get_candidate(cid) or cand

    fields = get_candidate_fields(cid)
    comp, missing = _completeness_check(cand.get("rank_guess", ""), fields)
    h1, h2, h3, h4, h5 = st.columns(5)
    h1.metric("Rank", cand.get("rank_guess") or "?")
    h2.metric("Status", cand.get("status") or "?")
    h3.metric("Score", f"{(cand.get('confidence') or 0):.2f}")
    h4.metric("Field completeness", f"{int(comp*100)} %")
    _src = cand.get("active_block_source") or "parser"
    h5.metric("Block source", "✋ manual" if _src == "manual" else "🤖 auto")
    if missing:
        st.warning(tt("⚠ Missing field:", "⚠️ Missing fields: ") + ", ".join(missing))

    left, mid, right = st.columns([1.7, 2.8, 2.2])
    units = get_cached_text_units(doc_id)
    unit_index = int(cand.get("unit_index") if cand.get("unit_index") is not None else -1)
    if unit_index < 0 or unit_index >= len(units):
        unit_index = next((i for i, u in enumerate(units) if u.page_number == cand.get("page_start")), 0)
    start_default = max(0, int(cand.get("start_unit_id") or unit_index))
    end_default = int(cand.get("end_unit_id") if cand.get("end_unit_id") is not None else min(len(units)-1, unit_index+8))
    if end_default < start_default:
        end_default = start_default

    with left:
        st.subheader(t("records_in_doc"))
        st.caption(t("all_taxa_found"))
        st.dataframe(pd.DataFrame([{
            "str.": r["page_start"], "taxon": r["taxon_name"],
            "rank": r["rank_guess"], "stav": r["status"],
            "Score": round(r["confidence"] or 0, 2)
        } for r in rows]), use_container_width=True, hide_index=True)

        st.markdown("**Block boundary information**")
        _bm = cand.get("boundary_method") or "parser"
        _bc = cand.get("boundary_confidence")
        _br = cand.get("boundary_reason") or ""
        _bc_pct = f"{int(float(_bc)*100)} %" if _bc is not None else "?"
        st.caption(f"Method: `{_bm}` · Confidence: {_bc_pct}")
        if _br:
            st.caption(tt(f"Reason: {_br}", f"Reason: {_br}"))
        if st.button(t("suggest_boundaries"), key=f"be_suggest_{cid}"):
            st.session_state[f"be_boundary_suggestion_{cid}"] = suggest_boundary_for_candidate(cid)
        if st.session_state.get(f"be_boundary_suggestion_{cid}"):
            sug = st.session_state[f"be_boundary_suggestion_{cid}"]
            st.json(sug, expanded=False)
            if st.button(t("apply_boundary_suggestion"), key=f"be_apply_suggest_{cid}"):
                s_id = int(sug.get("suggested_start_unit_id", start_default))
                e_id = int(sug.get("suggested_end_unit_id", end_default))
                block = "\n\n".join((units[i].text or "").strip() for i in range(s_id, e_id+1) if (units[i].text or "").strip())
                con = db()
                con.execute("""UPDATE taxon_candidates
                               SET start_unit_id=?, end_unit_id=?, block_text=?, boundary_method='suggested',
                                   boundary_reason=?, boundary_confidence=?, block_version=COALESCE(block_version,1)+1
                               WHERE id=?""",
                            (s_id, e_id, block, sug.get("reason", "suggested"), float(sug.get("confidence", 0.0)), cid))
                con.commit(); con.close()
                st.success(t("boundary_suggestion_applied")); st.rerun()

    with mid:
        st.subheader(t("text_boundaries"))
        st.caption(
            "The text is divided into paragraphs (units). Set the **from / to** range to include the full text record, then save."
        )
        rng1, rng2 = st.columns(2)
        start_unit = rng1.number_input(
            "From (first paragraph)", 0, max(0, len(units)-1), start_default,
            key=f"be_start_{cid}",
            help="Index firs paragraph of this record.")
        end_unit = rng2.number_input(
            "To (last paragraph)", 0, max(0, len(units)-1), end_default,
            key=f"be_end_{cid}",
            help="Index of last paragraph of this record.")
        if end_unit < start_unit:
            st.warning(t("end_before_start"))
        view_from = max(0, int(start_unit)-5); view_to = min(len(units)-1, int(end_unit)+5)
        preview_rows = []
        for i in range(view_from, view_to+1):
            u = units[i]
            preview_rows.append({
                "in": bool(int(start_unit) <= i <= int(end_unit)),
                "unit": i, "page": u.page_number,
                "roles": _unit_role_badges(u),
                "text": re.sub(r"\s+", " ", (u.text or "").strip())[:220]
            })
        st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
        c1, c2, c3 = st.columns(3)
        if c1.button("💾 Save start/end", key=f"be_save_bounds_{cid}"):
            block = "\n\n".join((units[i].text or "").strip()
                                for i in range(int(start_unit), int(end_unit)+1)
                                if (units[i].text or "").strip())
            con = db()
            con.execute("""UPDATE taxon_candidates
                           SET start_unit_id=?, end_unit_id=?, block_text=?, boundary_method='manual_units',
                               boundary_reason='manual_unit_range', boundary_confidence=1.0,
                               block_version=COALESCE(block_version,1)+1
                           WHERE id=?""", (int(start_unit), int(end_unit), block, cid))
            con.commit(); con.close()
            st.success(t("boundaries_updated")); st.rerun()
        if c2.button("➕ Add next", key=f"be_next_{cid}"):
            st.session_state[f"be_end_{cid}"] = min(len(units)-1, int(end_unit)+1); st.rerun()
        if c3.button("➖ Remove last", key=f"be_prev_{cid}"):
            st.session_state[f"be_end_{cid}"] = max(int(start_unit), int(end_unit)-1); st.rerun()

    with right:
        st.subheader(t("block_text_fields"))
        source = cand.get("active_block_source") or "parser"
        st.caption(
            "**Block source** determines, which text is used extraction of fields. 🤖 auto = text from the parser (automatic), ✋ manual = text manually insterted below."
        )
        st.radio("Block source", ["parser", "manual"],
                 index=1 if source == "manual" else 0,
                 format_func=lambda x: "🤖 Auto (from the parser)" if x == "parser" else "✋ Manual (your text)",
                 key=f"be_source_{cid}", horizontal=True)
        if st.button(t("switch_to_source"), key=f"be_set_source_{cid}"):
            set_active_block_source(cid, st.session_state[f"be_source_{cid}"]); st.rerun()
        with st.expander(t("auto_block_expander"), expanded=False):
            st.text_area("Parser block_text", value=cand.get("block_text") or "", height=220,
                         disabled=True, key=f"be_parser_{cid}")
        manual_default = cand.get("manual_block_text") or active_block or cand.get("block_text") or ""
        manual_text = st.text_area(
            "✏️ Manual block — paste/edit the record text here",
            value=manual_default, height=230, key=f"be_manual_{cid}",
            help="If the parser block is incorrect, enter the correct text here and save.")
        m1, m2 = st.columns(2)
        if m1.button("💾 Save manual block", key=f"be_save_manual_{cid}"):
            save_manual_block(cid, manual_text, activate=True)
            st.success(t("manual_block_saved")); st.rerun()
        if m2.button("↩️ Back to auto block", key=f"be_parser_on_{cid}"):
            set_active_block_source(cid, "parser"); st.rerun()
        if st.button("🔄 Extract the array from the block again", key=f"be_pipeline_{cid}",
                     help="Run automatic mapping of sections → fields from the current block."):
            run_post_approval_pipeline(cid); st.success(t("field_extraction_done")); st.rerun()
        events = detect_field_events(get_active_block_text(cand) or active_block)
        event_fields, event_debug = segment_fields_by_events(get_active_block_text(cand) or active_block, events)
        with st.expander(t("detected_sections_debug"), expanded=False):
            st.caption(
                "PaleoN searches for \"sections\" labels (Characteristics, Diagnosis, Occurrence...) in the block and divides the text into fields. Here you can see where and with what degree of certainty."
            )
            st.dataframe(pd.DataFrame([{
                "type": e.get("kind"), "field": e.get("field", ""), "label in text": e.get("label", ""),
                "position": e.get("pos"), "certainty": round(e.get("confidence", 0), 2)
            } for e in events]), use_container_width=True, hide_index=True)
        st.markdown(t("extracted_fields_md"))
        st.caption(t("extracted_fields_hint"))
        fields = get_candidate_fields(cid)
        if fields:
            st.dataframe(pd.DataFrame([{"pole": k, "hodnota": v[:500]}
                                       for k, v in sorted(fields.items()) if v]),
                         use_container_width=True, hide_index=True)
        else:
            st.info(t("no_fields_yet"))



def render_manual_taxon_creator(default_doc_id: Optional[int] = None) -> None:
    """UI block for fully creating a new taxonomic record in the Editor."""
    con = db()
    docs = con.execute("SELECT id, filename FROM documents ORDER BY created_at DESC, id DESC").fetchall()
    con.close()
    doc_options = {"— Create / use virtual document Manual records —": None}
    for d in docs:
        doc_options[f"{d['filename']} (ID {d['id']})"] = int(d["id"])
    rank_opts = ["Species", "Subspecies", "Genus", "Subgenus", "Family", "Subfamily", "Superfamily", "Order", "Suborder", "Class", "Phylum", "Higher taxon"]
    status_opts = ["approved", "needs_review", "pending"]
    with st.expander(t("create_record_expander"), expanded=False):
        st.caption(t("create_record_caption"))
        with st.form("manual_taxon_creator_form", clear_on_submit=False):
            c1, c2, c3 = st.columns([3, 1.4, 1.2])
            taxon_name = c1.text_input("Taxon name / name taxon *", placeholder="e.g. Gracilitheca astronauta")
            rank = c2.selectbox("Rank", rank_opts, index=0)
            status = c3.selectbox("Status", status_opts, index=0)
            doc_labels = list(doc_options.keys())
            default_index = 0
            if default_doc_id:
                for i, label in enumerate(doc_labels):
                    if doc_options[label] == int(default_doc_id):
                        default_index = i; break
            doc_label = st.selectbox("Document", doc_labels, index=default_index)
            page_start = st.number_input("Page / locator", min_value=1, value=1, step=1)
            block_text = st.text_area("RAW taxonomicky block", height=220, placeholder=t("raw_block_placeholder"))
            st.markdown(t("optional_base_fields"))
            f1, f2 = st.columns(2)
            field_values = {
                "TAXONOMIC PLACEMENT": f1.text_area("Taxonomic placement", height=70),
                "TYPE TAXON": f2.text_area("Type taxon", height=70),
                "TYPE SPECIMENS": f1.text_area("Type specimens", height=90),
                "LOCALITY": f2.text_area("Locality", height=90),
                "STRATIGRAPHY": f1.text_area("Stratigraphy", height=90),
                "DIAGNOSIS": f2.text_area("Diagnosis", height=90),
                "DESCRIPTION": f1.text_area("Description", height=120),
                "REMARKS": f2.text_area("Remarks", height=120),
                "OCCURRENCE": f1.text_area("Occurrence", height=90),
                "FIGURES": f2.text_input("Figures"),
                "REFERENCE": f1.text_input("Reference"),
                "AUTHOR": f2.text_input("Author"),
            }
            submitted = st.form_submit_button("✅ Create new record", type="primary")
        if submitted:
            try:
                cid = create_manual_taxon_record(
                    taxon_name=taxon_name, rank=rank, block_text=block_text,
                    document_id=doc_options.get(doc_label), status=status,
                    page_start=int(page_start or 1), fields=field_values,
                    source_note="manual_editor_form")
                st.session_state["editor_search_prefill"] = taxon_name.strip()
                st.success(tt(f"✅ Now record created: ID {cid} — {taxon_name.strip()}.", f"✅ Now record created: ID {cid} — {taxon_name.strip()}."))
                st.info(tt("Record saved in DB, FTS and Morpho/Strat. If you don't see it, refresh or search by name."))
            except Exception as exc:
                st.error(tt(f"❌ Record se nepodarilo vytvorit: {exc}", f"❌ Failed to create record: {exc}"))

def tab_editor():
    st.header(t("editor_header"))
    render_manual_taxon_creator()
    s = get_settings()

    # ── Načti schválené záznamy ───────────────────────────────────────────────
    con = db()
    approved = con.execute("""
        SELECT tc.id, tc.taxon_name, tc.rank_guess, tc.page_start,
               tc.block_text, tc.document_id, tc.confidence, tc.status,
               d.filename
        FROM taxon_candidates tc
        JOIN documents d ON tc.document_id=d.id
        WHERE tc.status IN ('approved','needs_review')
        ORDER BY d.filename, tc.page_start
    """).fetchall()
    con.close()

    if not approved:
        st.info(t("no_approved"))
        return

    # ── Živé hledání — VŽDY viditelné nad navigací (ne v expanderu) ──────────
    fdocs  = sorted(set(r["filename"]  for r in approved))
    franks = sorted(set(r["rank_guess"] for r in approved if r["rank_guess"]))

    sf1, sf2, sf3 = st.columns([3, 1, 1])
    _ed_prefill = st.session_state.pop("editor_search_prefill", "") if st.session_state.get("editor_search_prefill") else ""
    f_search = sf1.text_input(
        t("search_taxon"), value=_ed_prefill, key="ed_search",
        placeholder="Gracilitheca*, Hyolith?, *theca…  (* = cokoliv, ? = 1 znak)",
        label_visibility="collapsed")
    f_doc  = sf2.selectbox("Document", ["— All —"] + fdocs,
                            key="ed_filter_doc", label_visibility="collapsed")
    f_rank = sf3.selectbox("Rank", ["— All —"] + franks,
                            key="ed_filter_rank", label_visibility="collapsed")

    # Wildcard podpora: * → .* , ? → . ; bez wildcard = substring hledání
    def _match_search(name: str, pattern: str) -> bool:
        if not pattern:
            return True
        _n = _norm_search(name)
        _p = _norm_search(pattern)
        if "*" in _p or "?" in _p:
            try:
                rx = re.compile(
                    _p.replace(".", r"\.").replace("*", ".*").replace("?", "."),
                    re.IGNORECASE)
                return bool(rx.search(_n))
            except re.error:
                return _p in _n
        return _p in _n

    filtered = [
        r for r in approved
        if _match_search(r["taxon_name"] or "", f_search)
        and (f_doc  == "— All —" or r["filename"]  == f_doc)
        and (f_rank == "— All —" or r["rank_guess"] == f_rank)
    ]

    if not filtered:
        st.warning(t("no_record_in_filter"))
        return

    # Pokud je aktivní vyhledávání (neprázdný f_search), zobraz tabulku výsledku
    if f_search and f_search.strip():
        st.caption(tt(f"🔎 Nalezeno {len(filtered)} results — click a row to edit", f"🔎 Found {len(filtered)} results — click a row to edit"))
        _search_df = pd.DataFrame([{
            "Status": {"approved": "✅", "needs_review": "🔍"}.get(r["status"], "⏳"),
            "Taxon": r["taxon_name"] or "?",
            "Rank": r["rank_guess"] or "?",
            "Str.": r["page_start"],
            "Document": r["filename"],
            "_i": i,
        } for i, r in enumerate(filtered)])
        _sr = st.dataframe(
            _search_df.drop(columns=["_i"]),
            use_container_width=True, hide_index=True,
            height=min(36 * len(_search_df) + 38, 320),
            on_select="rerun", selection_mode="single-row",
            key="ed_search_table")
        _sr_rows = getattr(_sr, "selection", {})
        if isinstance(_sr_rows, dict):
            _sr_rows = _sr_rows.get("rows", [])
        elif hasattr(_sr_rows, "rows"):
            _sr_rows = list(_sr_rows.rows)
        else:
            _sr_rows = []
        if _sr_rows and _sr_rows[0] < len(_search_df):
            _new_idx = int(_search_df.iloc[_sr_rows[0]]["_i"])
            if _new_idx != st.session_state.get("editor_idx", 0):
                st.session_state["editor_idx"] = _new_idx
                st.rerun()
        st.divider()

    id_list = [r["id"] for r in filtered]

    if "editor_idx" not in st.session_state:
        st.session_state["editor_idx"] = 0
    idx = st.session_state["editor_idx"]
    if idx >= len(filtered):
        idx = 0; st.session_state["editor_idx"] = 0

    # ── Dvoupanelový layout: levý = literatura+taxoni, pravý = editor ────────

    # ── Dvoupanelový layout: levý = taxony, pravý = editor ──────────────────
    ed_left, ed_right = st.columns([1, 3], gap="medium")

    with ed_left:
        # ── Seznam allch taxa (filtrovaný dle horní lišty) ─────────────────
        # Literatura se vybírá filtrem "Document" v horní liště (f_doc selectbox).
        # Když je vybraný konkrétní Document, zobrazí se jen jeho taxony.
        _left_label = (
            f"**🔬 Taxony** — {len(filtered)} records"
            + (f" · 📄 {f_doc}" if f_doc != "— All —" else "")
            + (f" · rank: {f_rank}" if f_rank != "— All —" else "")
        )
        st.caption(_left_label)

        if not filtered:
            st.info(t("no_record_in_filter"))
        else:
            _ed_tax_df = pd.DataFrame([{
                "✔":    {"approved": "✅", "needs_review": "🔍"}.get(r["status"], "⏳"),
                "Taxon": r["taxon_name"] or "?",
                "Rank":  r["rank_guess"] or "?",
                "Str.":  r["page_start"],
                "_ri":   filtered.index(r),
            } for r in filtered])

            _ed_tax_sel = st.dataframe(
                _ed_tax_df.drop(columns=["_ri"]),
                column_config={
                    "✔":    st.column_config.TextColumn("", width="small"),
                    "Taxon": st.column_config.TextColumn("Taxon"),
                    "Rank":  st.column_config.TextColumn("Rank", width="small"),
                    "Str.":  st.column_config.NumberColumn("Str.", width="small"),
                },
                hide_index=True,
                use_container_width=True,
                height=min(36 * len(_ed_tax_df) + 38, 560),
                on_select="rerun",
                selection_mode="single-row",
                key="ed_tax_table",
            )
            _ed_rows = getattr(_ed_tax_sel, "selection", {})
            if isinstance(_ed_rows, dict):
                _ed_rows = _ed_rows.get("rows", [])
            elif hasattr(_ed_rows, "rows"):
                _ed_rows = list(_ed_rows.rows)
            else:
                _ed_rows = []
            if _ed_rows and _ed_rows[0] < len(filtered):
                _new_ri = int(_ed_tax_df.iloc[_ed_rows[0]]["_ri"])
                if _new_ri != idx:
                    st.session_state["editor_idx"] = _new_ri
                    idx = _new_ri
                    st.rerun()
        st.caption(tt(f"Record {idx+1}/{len(filtered)}", f"Record {idx+1}/{len(filtered)}"))

    # ── Pravý panel: vlastní editor záznamu ──────────────────────────────
    with ed_right:

        cand = get_candidate(id_list[idx])
        if not cand:
            st.error(t("record_not_found")); return
        cand_id = cand["id"]

        # ── Akce na aktuálním záznamu (klávesy A = schválit, R = odmítnout) ──────
        act1, act2, act3, act4 = st.columns([1,1,1,3])
        if act1.button(t("approve"), key=f"ed_approve_{cand_id}", type="primary"):
            _batch_update_status([cand_id], "approved"); st.rerun()
        if act2.button(t("reject"), key=f"ed_reject_{cand_id}"):
            _batch_update_status([cand_id], "rejected"); st.rerun()
        if act3.button(t("needs_review"), key=f"ed_needsreview_{cand_id}"):
            _batch_update_status([cand_id], "needs_review"); st.rerun()
        act4.markdown(
            f"<div style='padding-top:6px'>{_badge_html(cand['status'])}</div>",
            unsafe_allow_html=True)

        _inject_keyboard_shortcuts("editor", "Approve (A)", "Reject (R)")

        # ── Header záznamu ────────────────────────────────────────────────────────
        filled, total = _count_filled_fields(cand_id)
        fill_pct = int(filled/total*100) if total else 0
        conf_color = "#047857" if cand["confidence"]>=0.80 else "#b45309" if cand["confidence"]>=0.60 else "#b91c1c"

        st.markdown(
            f'<div class="taxon-header">'
            f'🦕 {cand["taxon_name"]}'
            f'<span style="float:right;font-size:0.75rem;opacity:0.85">'
            f'{cand["rank_guess"] or "?"} | p.{cand["page_start"]} | '
            f'Score: <span style="color:#93c5fd">{cand["confidence"]:.3f}</span>'
            f'</span></div>',
            unsafe_allow_html=True)

        # ── Editace jména taxonu ──────────────────────────────────────────────
        with st.expander(t("edit_name_assign"), expanded=False):
            _en1, _en2 = st.columns([3, 1])
            _new_name = _en1.text_input(
                "Name taxon", value=cand["taxon_name"] or "",
                key=f"ed_rename_{cand_id}",
                help="Opravte preklep or OCR chybu v jmenu taxonu.")
            _new_rank = _en2.selectbox(
                "Rank", [""] + RANK_OPTIONS,
                index=(RANK_OPTIONS.index(cand["rank_guess"]) + 1
                       if cand.get("rank_guess") in RANK_OPTIONS else 0),
                key=f"ed_rerank_{cand_id}",
                label_visibility="collapsed")
            if st.button(t("save_name_rank_btn"), key=f"ed_save_name_{cand_id}"):
                _saved_name  = _new_name.strip()
                _saved_rank  = _new_rank.strip() or cand.get("rank_guess") or ""
                if _saved_name:
                    _con_rn = db()
                    _con_rn.execute(
                        "UPDATE taxon_candidates SET taxon_name=?, rank_guess=? WHERE id=?",
                        (_saved_name, _saved_rank, cand_id))
                    _con_rn.commit(); _con_rn.close()
                    st.success(tt(f"Saved: {_saved_name} [{_saved_rank}]", f"Saved: {_saved_name} [{_saved_rank}]")); st.rerun()

            # Přiřazení ke správnému taxonu (synonymum / přesun)
            _con_all = db()
            _all_taxa = [r[0] for r in _con_all.execute(
                "SELECT DISTINCT taxon_name FROM taxon_candidates "
                "WHERE taxon_name IS NOT NULL AND taxon_name != '' "
                "AND id != ? ORDER BY taxon_name",
                (cand_id,)).fetchall()]
            _con_all.close()
            if _all_taxa:
                st.caption(t("assign_to_taxon_caption"))
                _link_target = st.selectbox(
                    "Correct taxon", ["— not selected —"] + _all_taxa,
                    key=f"ed_link_{cand_id}",
                    label_visibility="collapsed",
                    help="Select cilovy taxon ze seznamu allch records v databazi.")
                if _link_target != "— not selected —":
                    if st.button(tt(f"🔗 Assign to '{_link_target}'", f"🔗 Assign to '{_link_target}'"),
                                 key=f"ed_do_link_{cand_id}"):
                        _con_lnk = db()
                        _con_lnk.execute(
                            "UPDATE taxon_candidates SET taxon_name=? WHERE id=?",
                            (_link_target, cand_id))
                        # Přidat original jméno jako synonymum
                        if cand.get("taxon_name") and cand["taxon_name"] != _link_target:
                            _cur_syn = _con_lnk.execute(
                                "SELECT field_value FROM occurrence_fields "
                                "WHERE candidate_id=? AND field_name='SYNONYMY'",
                                (cand_id,)).fetchone()
                            _syn_val = (_cur_syn[0] + "; " if _cur_syn else "") + \
                                       f"{cand['taxon_name']} [assigned to {_link_target}]"
                            _con_lnk.execute(
                                "INSERT OR REPLACE INTO occurrence_fields "
                                "(candidate_id, field_name, field_value, method) "
                                "VALUES (?, 'SYNONYMY', ?, 'manual')",
                                (cand_id, _syn_val))
                        _con_lnk.commit(); _con_lnk.close()
                        st.success(tt(f"Assigned to '{_link_target}'.", f"Assigned to '{_link_target}'.")); st.rerun()

        bar_color = "#ef4444" if fill_pct<30 else "#f59e0b" if fill_pct<70 else "#10b981"
        st.markdown(
            f'<div class="fill-wrap">'
            f'<div class="fill-bar" style="width:{fill_pct}%;background:{bar_color}"></div>'
            f'</div><p style="font-size:0.75rem;color:#666;margin:0">'
            f'Filled: {filled}/{total} fields ({fill_pct} %)</p>',
            unsafe_allow_html=True)

        st.markdown("")  # spacer

        # ── Načti/extrahuj blok ───────────────────────────────────────────────────
        block = cand["block_text"] or ""
        if not block:
            with st.spinner("Extrahuji blok textu…"):
                block = extract_block_for_candidate(cand_id, cand["document_id"], cand["page_start"])
            if block:
                con = db()
                # Cap na 200 KB — extrémně dlouhé bloky by zpusobily problémy v DB/UI
                con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                            (block[:200_000], cand_id))
                con.commit(); con.close()

        # ── Hlavní layout: blok vlevo, pole vpravo ────────────────────────────────
        left_col, right_col = st.columns([2, 3])

        with left_col:
            hdr_col, toggle_col = st.columns([3,2])
            hdr_col.subheader("📜 RAW blok (verbatim)")
            show_annotated = toggle_col.toggle(
                t("annotate_toggle"), value=False, key=f"ann_{cand_id}",
                help="Shows before every paragraph label category.")

            # ── PDF inline vedle RAW bloku ────────────────────────────────────
            _ed_doc_path_key = f"docpath_{cand['document_id']}"
            if _ed_doc_path_key not in st.session_state:
                _ed_dcon = db()
                _ed_dr = _ed_dcon.execute(
                    "SELECT path FROM documents WHERE id=?",
                    (cand["document_id"],)).fetchone()
                _ed_dcon.close()
                st.session_state[_ed_doc_path_key] = _ed_dr["path"] if _ed_dr else None
            _ed_doc_p = st.session_state.get(_ed_doc_path_key)
            _ed_is_pdf = bool(_ed_doc_p and pathlib.Path(_ed_doc_p).suffix.lower() == ".pdf"
                               and pathlib.Path(_ed_doc_p).exists() and HAS_FITZ)
            _pg_s = int(dict(cand).get("page_start") or 1)
            _pg_e = int(dict(cand).get("block_end_page") or _pg_s)
            _pg_lbl = f"str. {_pg_s}–{_pg_e}" if _pg_e > _pg_s else f"str. {_pg_s}"
            if _ed_is_pdf:
                _ed_pdf_show = st.toggle(
                    f"📄 View PDF ({_pg_lbl})",
                    value=False,
                    key=f"ed_pdf_tog_{cand_id}",
                    help="Shows page(s) of original in the app.")
                if _ed_pdf_show:
                    _n_ed = _pg_e - _pg_s + 1
                    if _n_ed <= 1:
                        _show_pdf_page_inline(
                            cand["document_id"], _pg_s, key=f"ed_pdf_{cand_id}")
                    else:
                        _ed_pcols = st.columns(min(_n_ed, 3))
                        for _epi, _epg in enumerate(range(_pg_s, min(_pg_e+1, _pg_s+6))):
                            with _ed_pcols[_epi % 3]:
                                st.caption(f"str. {_epg}")
                                _show_pdf_page_inline(
                                    cand["document_id"], _epg,
                                    key=f"ed_pdf_{cand_id}_{_epg}")

            # ── Boundary tlačítka ──────────────────────────────────────────────
            # Přidej/Odeber jeden odstavec ze začátku or konce bloku.
            bnd1, bnd2, bnd3, bnd4 = st.columns(4)
            if bnd1.button("⬆ Add before", key=f"bnd_add_before_{cand_id}",
                           help="Adds text unit before start block"):
                units = get_cached_text_units(cand["document_id"])
                idx_u = cand["unit_index"] if cand["unit_index"] is not None else -1
                if idx_u > 0:
                    prev_text = units[idx_u - 1].text
                    new_block = prev_text + "\n\n" + (block or "")
                    con = db()
                    con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                (new_block, cand_id))
                    con.commit(); con.close()
                    st.success(t("prev_unit_added")); st.rerun()
                else:
                    st.warning(t("no_prev_unit"))
            if bnd2.button("⬇ Remove before", key=f"bnd_rm_before_{cand_id}",
                           help="Removes first paragraph block"):
                paragraphs = [p for p in (block or "").split("\n\n") if p.strip()]
                if len(paragraphs) > 1:
                    new_block = "\n\n".join(paragraphs[1:])
                    con = db()
                    con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                (new_block, cand_id))
                    con.commit(); con.close()
                    st.success(t("first_para_removed")); st.rerun()
                else:
                    st.warning(t("only_one_para"))
            if bnd3.button("⬆ Odeber po", key=f"bnd_rm_after_{cand_id}",
                           help="Removes posledni paragraph block"):
                paragraphs = [p for p in (block or "").split("\n\n") if p.strip()]
                if len(paragraphs) > 1:
                    new_block = "\n\n".join(paragraphs[:-1])
                    con = db()
                    con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                (new_block, cand_id))
                    con.commit(); con.close()
                    st.success(t("last_para_removed")); st.rerun()
                else:
                    st.warning(t("only_one_para"))
            if bnd4.button("⬇ Add after", key=f"bnd_add_after_{cand_id}",
                           help="Adds text aunit after end of block"):
                units = get_cached_text_units(cand["document_id"])
                con = db()
                all_cands_u = con.execute(
                    "SELECT unit_index FROM taxon_candidates "
                    "WHERE document_id=? AND status!='rejected' ORDER BY unit_index",
                    (cand["document_id"],)).fetchall()
                con.close()
                my_ui = cand["unit_index"] if cand["unit_index"] is not None else -1
                next_boundary = next(
                    (c["unit_index"] for c in all_cands_u
                     if c["unit_index"] is not None and c["unit_index"] > my_ui),
                    len(units))
                if next_boundary < len(units):
                    next_text = units[next_boundary].text
                    new_block = (block or "") + "\n\n" + next_text
                    con = db()
                    con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                (new_block, cand_id))
                    con.commit(); con.close()
                    st.success(t("next_unit_added")); st.rerun()
                else:
                    st.warning(t("no_next_unit"))

            if show_annotated:
                st.text_area(
                    "raw_annotated", value=annotate_raw_block(block), height=340,
                    key=f"raw_ann_view_{cand_id}", label_visibility="collapsed",
                    disabled=True)
                edited_block = block
            else:
                edited_block = st.text_area(
                    "raw", value=block, height=340,
                    key=f"raw_{cand_id}", label_visibility="collapsed")

            b1, b2, b3 = st.columns(3)
            if b1.button(t("save_block"), key=f"sblk_{cand_id}"):
                if edited_block != block:
                    con = db()
                    con.execute("UPDATE taxon_candidates SET block_text=? WHERE id=?",
                                (edited_block, cand_id))
                    con.commit(); con.close()
                    block = edited_block
                    st.success(t("block_saved_ok2"))

            if b2.button(t("automap"), key=f"amap_{cand_id}",
                         help="Maps sections using regex from schema TSV"):
                mapped = map_sections_from_block(edited_block or block, rank=cand.get("rank_guess",""))
                if mapped:
                    save_fields(cand_id, mapped, method="regex")
                    st.success(tt(f"{len(mapped)} fields: {list(mapped.keys())}", f"{len(mapped)} fields: {list(mapped.keys())}"))
                    st.rerun()
                else:
                    st.warning(t("no_labels_found2"))

            if b3.button(t("llm_fields"), key=f"llmf_{cand_id}",
                         disabled=not s.get("llm_enabled")):
                try:
                    with st.spinner(t("llm_assigning_fields")):
                        raw_resp = lm_chat(
                            s, s.get("llm_field_prompt", LLM_FIELD_PROMPT),
                            f"Taxon: {cand['taxon_name']}\n\nBlok:\n{(edited_block or block)[:4000]}")
                    result = lm_parse_json(raw_resp)
                    if result and "fields" in result:
                        llm_fields = {k:v for k,v in result["fields"].items()
                                      if k in SECTION_FIELDS}
                        save_fields(cand_id, llm_fields, method="llm")
                        st.success(tt(f"LLM: {len(llm_fields)} fields", f"LLM: {len(llm_fields)} fields"))
                        st.rerun()
                    else:
                        st.warning(t("llm_no_valid_json"))
                except Exception as exc:
                    st.error(f"LLM chyba: {exc}")

            # ── Ruční překlad fields (LM Studio) ────────────────────────────
            with st.expander(t("translate_fields_expander"), expanded=False):
                _doc_lang_tr_r = db().execute(
                    "SELECT lang FROM documents WHERE id=?",
                    (cand["document_id"],)).fetchone()
                _doc_lang_tr = (_doc_lang_tr_r["lang"] if _doc_lang_tr_r else "") or ""
                _llm_ok = s.get("llm_enabled", False)

                if not _llm_ok:
                    st.warning(
                        "⚠ LM Studio neni zapnute. Aktivujte ho v Settings → ⚙ LM Studio → Enable LLM.",
                        icon="⚠️")
                else:
                    # Zobrazit detekovaný jazyk; nechat uživatele přepsat
                    _lang_detect_display = _doc_lang_tr.upper() if _doc_lang_tr else "?"
                    _tr_c1, _tr_c2 = st.columns([2, 3])
                    _tr_c1.caption(f"Detected document language: **{_lang_detect_display}**")
                    _force_lang_override = _tr_c2.text_input(
                        "Language (overwrite)", value=_doc_lang_tr,
                        placeholder="cs / ru / de / fr / zh …",
                        key=f"tr_lang_{cand_id}",
                        help="Zadejte kód jazyka pokud auto-detekce selhala (cs, ru, de, fr, zh…)",
                        label_visibility="collapsed")
                    _eff_lang = _force_lang_override.strip() or _doc_lang_tr
                    _is_expl_en = _eff_lang.lower() in {"en","eng","english","en-gb","en-us"}
                    if _is_expl_en:
                        st.info(t("doc_is_english") + " " + tt(
                                "For translation remove 'en' from field language above.",
                                " To force translation, remove 'en' from the language field above."))
                    _tr_btn = st.button(
                        f"🌐 Translate → EN{f' (z {_eff_lang.upper()})' if _eff_lang else ''}",
                        key=f"translate_now_{cand_id}",
                        disabled=not _llm_ok,
                        help="Translates all fields to English (format: translation (original))")
                    if _tr_btn:
                        try:
                            with st.spinner(t("lmstudio_translating")):
                                _n_tr = auto_translate_candidate_fields(
                                    cand_id, _eff_lang, s,
                                    force=True,
                                    force_lang=_eff_lang)
                            if _n_tr:
                                st.success(tt(f"✅ Translated {_n_tr} fields to English.", f"✅ Translated {_n_tr} fields to English."))
                                st.rerun()
                            else:
                                st.info(
                                    "No fields to translate — are fields filled in and is the language set correctly?")
                        except RuntimeError as _tr_err:
                            st.error(tt(f"❌ Translation error: {_tr_err}", f"❌ Translation error: {_tr_err}"))
                        except Exception as _tr_exc:
                            st.error(tt(f"❌ Unexpected error: {_tr_exc}", f"❌ Unexpected error: {_tr_exc}"))

            # ── Field Mapper Panel ─────────────────────────────────────────────
            # Rozbalitelný panel: seznam allch labels detekovaných v bloku →
            # cílové pole + možnost přemapovat na jiné pole kliknutím.
            with st.expander(t("field_mapper_expander"), expanded=False):
                current_block = edited_block or block
                if current_block:
                    detected_sections = map_sections_from_block(current_block, rank=cand.get("rank_guess",""))
                    if detected_sections:
                        st.caption(
                            f"Detected {len(detected_sections)} sections. "
                            "Click 'Save' to remap fields.")
                        existing_fields = get_candidate_fields(cand_id)
                        remap_changes: Dict[str, str] = {}
                        for sec_field, sec_val in detected_sections.items():
                            fm_c1, fm_c2, fm_c3 = st.columns([2, 2, 3])
                            fm_c1.markdown(f"**{sec_field}**")
                            new_field = fm_c2.selectbox(
                                "→", ["(skip)"] + list(SECTION_FIELDS),
                                index=(["(skip)"] + list(SECTION_FIELDS)).index(sec_field)
                                      if sec_field in SECTION_FIELDS else 0,
                                key=f"remap_{cand_id}_{sec_field}",
                                label_visibility="collapsed")
                            fm_c3.caption(sec_val[:60] + ("…" if len(sec_val) > 60 else ""))
                            if new_field and new_field != "(skip)":
                                remap_changes[new_field] = sec_val
                        if st.button(t("save_remap_btn"), key=f"remap_save_{cand_id}"):
                            if remap_changes:
                                save_fields(cand_id, remap_changes, method="manual")
                                st.success(tt(f"Saved {len(remap_changes)} fields.", f"Saved {len(remap_changes)} fields.")); st.rerun()
                            else:
                                st.warning(t("no_changes_remap"))
                    else:
                        st.caption(t("no_sections_in_block") +
                                   tt("Zkuste Auto-map or upravte text bloku.",
                                      "Try Auto-map or edit the block text."))
                else:
                    st.caption(t("load_block_first"))



            st.divider()
            st.subheader(t("context_subheader"))
            with st.expander(t("context_before_after"), expanded=False):
                st.text_area(t("context_before_ta"), value=cand["context_before"] or "–",
                             height=100, disabled=True, key=f"ctx_b_{cand_id}",
                             label_visibility="collapsed")
                st.text_area(t("context_after_ta"), value=cand["context_after"] or "–",
                             height=100, disabled=True, key=f"ctx_a_{cand_id}",
                             label_visibility="collapsed")
            try:
                dbg = json.loads(cand["debug_json"] or "{}")
                with st.expander("🔬 Scoring debug", expanded=False):
                    st.json(dbg, expanded=True)
            except Exception:
                pass
            llm_j = json.loads(cand["llm_json"] or "{}") if cand["llm_json"] else {}
            if llm_j:
                action = llm_j.get("action","?")
                color  = {"keep":"success","reject":"error","needs_review":"warning"}.get(action,"info")
                getattr(st, color)(
                    f"LLM: **{action}** [{llm_j.get('confidence','?')}] "
                    f"_{llm_j.get('reason','')}_")

        with right_col:
            st.subheader(t("fields_header"))
            fields = get_candidate_fields(cand_id)
            schema = load_schema()

            # BOD 9: Auto-save indikátor
            st.markdown(
                "<div style='font-size:0.72rem;color:#10b981;margin-bottom:4px'>💾 Auto-save aktivni — fields se savei pri odchodu z bunky</div>", unsafe_allow_html=True)
            # ── Rank-aware pole varování ────────────────────────────────────────
            # Upozornit pokud jsou vyplněna pole neodpovídající ranku:
            #   • species/subspecies: TYPE TAXON by neměl být vyplněn
            #   • genus/family/…: TYPE SPECIMENS by neměl být vyplněn (jako heading)
            _cand_rank_l = (cand.get("rank_guess","") or "").lower().split()
            _cand_rank_l = _cand_rank_l[0] if _cand_rank_l else ""
            if _cand_rank_l in _REQUIRED_FIELDS_BY_RANK:
                _rk_eff = _cand_rank_l
            else:
                _rk_eff = _RANK_ALIASES.get(_cand_rank_l, "")
            if _rk_eff in ("species", "subspecies"):
                if fields.get("TYPE TAXON","") and fields["TYPE TAXON"] != NOT_PROVIDED:
                    st.warning(
                        "⚠ **TYPE TAXON** je vyplnen u species/subspecies — zkontrolujte, zda nejde o TYPE SPECIMENS (holotyp/paratypus). Species maji TYPE SPECIMENS, ne TYPE TAXON.",
                        icon="⚠️")
            elif _rk_eff in ("genus","family","order","class","phylum"):
                _ts_val = fields.get("TYPE SPECIMENS","") or ""
                if _ts_val and _ts_val != NOT_PROVIDED:
                    if not any(kw in _ts_val.lower() for kw in TYPE_SPECIMEN_KEYWORDS):
                        st.warning(
                            "⚠ **TYPE SPECIMENS** je vyplnen u genus/family/… — pravdepodobne jde o TYPE TAXON (typovy druh/rod). Genus a vyssi maji TYPE TAXON, ne TYPE SPECIMENS.",
                            icon="⚠️")
            _cand_rank_for_groups = cand.get("rank_guess", "") or ""
            _fgroups_ranked = ordered_field_groups_for_rank(_cand_rank_for_groups)
            group_names = list(_fgroups_ranked.keys())
            group_tabs  = st.tabs(group_names)
            for gi, (grp_name, grp_fields) in enumerate(_fgroups_ranked.items()):
                with group_tabs[gi]:
                    for fname in grp_fields:
                        cur_val = fields.get(fname, "")
                        is_filled = bool(cur_val and cur_val != NOT_PROVIDED)
                        label = f"{'✅' if is_filled else '⬜'} {fname}"
                        # BOD 9: Auto-save — saveí se při odchodu z pole (on_blur).
                        # Bez st.rerun() → bez blikání. "✅" v labelu signalizuje saveení.
                        _saved_key = f"saved_{cand_id}_{fname}"
                        _was_saved = st.session_state.get(_saved_key, False)
                        _label_icon = "✅" if is_filled else ("💾" if _was_saved else "⬜")
                        _label_auto = f"{_label_icon} {fname}"
                        new_val = st.text_area(
                            _label_auto, value=cur_val or "", height=90,
                            key=f"field_{cand_id}_{fname}",
                            help=f"{fname} — saved automatically when leaving the field",
                            placeholder=f"Zadejte {fname}…")
                        if new_val != cur_val:
                            save_fields(cand_id, {fname: new_val}, method="manual")
                            st.session_state[_saved_key] = True
                            # Tichý save bez rerun → žádný flash stránky.
                            # Fill bar se aktualizuje při příští interakci uživatele.

            st.divider()
            # Taxon-name rename
            with st.expander(tt("✏ Rename taxon", "✏️ Rename taxon"), expanded=False):
                new_name = st.text_input(
                    "Now taxon name", value=cand["taxon_name"],
                    key=f"rename_{cand_id}")
                new_rank = st.selectbox(
                    "Rank", ["Genus","Species","Family","Subfamily","Order","Class",
                              "Phylum","Subspecies","Subgenus","Variety","Other"],
                    index=["Genus","Species","Family","Subfamily","Order","Class",
                           "Phylum","Subspecies","Subgenus","Variety","Other"
                           ].index(cand["rank_guess"])
                          if cand["rank_guess"] in
                             ["Genus","Species","Family","Subfamily","Order","Class",
                              "Phylum","Subspecies","Subgenus","Variety","Other"] else 0,
                    key=f"rank_{cand_id}")
                if st.button(t("save_name"), key=f"save_name_{cand_id}"):
                    con = db()
                    con.execute(
                        "UPDATE taxon_candidates SET taxon_name=?, rank_guess=? WHERE id=?",
                        (new_name.strip(), new_rank, cand_id))
                    con.commit(); con.close()
                    st.success(tt("Saved.", "Saved.")); st.rerun()

            # Rychlé výskytové čítadlo
            con = db()
            occ_count = con.execute(
                "SELECT COUNT(*) FROM taxon_candidates WHERE taxon_name=? AND id!=?",
                (cand["taxon_name"], cand_id)).fetchone()[0]
            con.close()
            if occ_count:
                st.caption(tt(f"📚 This taxon appears {occ_count}× more in the library.", f"📚 This taxon occurs {occ_count}× more in the library."))

            filled2, _ = _count_filled_fields(cand_id)
            st.caption(tt(f"Filled: {filled2}/{total} fields", f"Filled: {filled2}/{total} fields"))



    # ══════════════════════════════════════════════════════════════════════════════
    # BOD 6: CSV export (UTF-8 BOM pro Excel)
    # ══════════════════════════════════════════════════════════════════════════════

    def export_to_csv(rows: List[Dict]) -> bytes:
        """
        Exportuje záznamy do CSV (UTF-8 BOM).
        UTF-8 BOM zajišťuje správné zobrazení diakritiky v Excel na Windows.
        """
        import csv, io as _io
        out = _io.StringIO()
        if not rows:
            return b"\xef\xbb\xbf"
        writer = csv.DictWriter(
            out,
            fieldnames=list(rows[0].keys()),
            extrasaction="ignore",
            lineterminator="\r\n",
            quoting=csv.QUOTE_MINIMAL,
        )
        writer.writeheader()
        writer.writerows(rows)
        # Přidat BOM na začátek pro Excel
        return ("\ufeff" + out.getvalue()).encode("utf-8")


    # ══════════════════════════════════════════════════════════════════════════════
    # BOD 8: Darwin Core Archive (DwC-A) export
    # Formát GBIF — ZIP s occurrence.csv + meta.xml + eml.xml
    # ══════════════════════════════════════════════════════════════════════════════

    # Mapování PaleoN fields na Darwin Core termíny (URI)
    _DWC_FIELD_MAP: Dict[str, str] = {
        "RECORD_ID":                "occurrenceID",
        "TAXON_NAME_VERBATIM":      "scientificName",
        "TAXON_RANK_AS_WRITTEN":    "taxonRank",
        "AUTHOR":                   "scientificNameAuthorship",
        "DIAGNOSIS":                "taxonRemarks",
        "DESCRIPTION":              "occurrenceRemarks",
        "TYPE SPECIMENS":           "typeStatus",
        "TYPE_SPECIMEN_KIND":       "typeStatus",
        "INSTITUTION_CODE":         "institutionCode",
        "CATALOG_NUMBER":           "catalogNumber",
        "LOCALITY":                 "verbatimLocality",
        "STRATIGRAPHY":             "verbatimEventDate",
        "OCCURRENCE":               "habitat",
        "SIZE":                     "measurementValue",
        "ETYMOLOGY":                "taxonRemarks",
        "REMARKS":                  "occurrenceRemarks",
        "SYNONYMY":                 "namePublishedIn",
        "FIGURES":                  "associatedMedia",
        "REFERENCE":                "namePublishedIn",
        "SOURCE PAGES":             "verbatimCoordinates",
        "SOURCE_DOCUMENT":          "datasetName",
    }

    _DWC_META_XML = """<?xml version="1.0" encoding="UTF-8"?>
    <archive xmlns="http://rs.tdwg.org/dwc/text/"
             xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
             xsi:schemaLocation="http://rs.tdwg.org/dwc/text/ http://rs.tdwg.org/dwc/text/tdwg_dwc_text.xsd">
      <core encoding="UTF-8" fieldsTerminatedBy="," linesTerminatedBy="\r\n"
            fieldsEnclosedBy="\"" ignoreHeaderLines="1"
            rowType="http://rs.tdwg.org/dwc/terms/Taxon">
        <files><location>occurrence.csv</location></files>
        <id index="0"/>
        <field index="0" term="http://rs.tdwg.org/dwc/terms/occurrenceID"/>
        <field index="1" term="http://rs.tdwg.org/dwc/terms/scientificName"/>
        <field index="2" term="http://rs.tdwg.org/dwc/terms/taxonRank"/>
        <field index="3" term="http://rs.tdwg.org/dwc/terms/scientificNameAuthorship"/>
        <field index="4" term="http://rs.tdwg.org/dwc/terms/institutionCode"/>
        <field index="5" term="http://rs.tdwg.org/dwc/terms/catalogNumber"/>
        <field index="6" term="http://rs.tdwg.org/dwc/terms/typeStatus"/>
        <field index="7" term="http://rs.tdwg.org/dwc/terms/verbatimLocality"/>
        <field index="8" term="http://rs.tdwg.org/dwc/terms/verbatimEventDate"/>
        <field index="9" term="http://rs.tdwg.org/dwc/terms/occurrenceRemarks"/>
        <field index="10" term="http://rs.tdwg.org/dwc/terms/taxonRemarks"/>
        <field index="11" term="http://rs.tdwg.org/dwc/terms/namePublishedIn"/>
        <field index="12" term="http://rs.tdwg.org/dwc/terms/datasetName"/>
        <field index="13" term="http://rs.tdwg.org/dwc/terms/basisOfRecord"/>
        <field index="14" term="http://rs.tdwg.org/dwc/terms/kingdom"/>
        <field index="15" term="http://rs.tdwg.org/dwc/terms/phylum"/>
      </core>
    </archive>
    """

    def _dwc_eml_xml(n_records: int, dataset_name: str, creator: str) -> str:
        from xml.sax.saxutils import escape as _esc
        return (
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            '<eml:eml xmlns:eml="eml://ecoinformatics.org/eml-2.1.1"\n'
            '         xmlns:dc="http://purl.org/dc/terms/"\n'
            '         packageId="paleon-export" system="paleon">\n'
            '  <dataset>\n'
            f'    <title>{_esc(dataset_name)}</title>\n'
            f'    <creator><individualName><givenName>{_esc(creator)}</givenName>'
            f'</individualName></creator>\n'
            '    <abstract><para>Taxonomic data extracted by PaleoN.</para></abstract>\n'
            f'    <numberOfRecords>{n_records}</numberOfRecords>\n'
            '  </dataset>\n'
            '</eml:eml>\n'
        )


    def export_to_dwca(
        rows: List[Dict],
        dataset_name: str = "PaleoN Export",
        creator: str = "PaleoN",
    ) -> bytes:
        """
        Exportuje záznamy do Darwin Core Archive (DwC-A) — standardní formát GBIF.
        Vrátí ZIP v paměti (bytes).

        Struktura ZIP:
            occurrence.csv  — hlavní data (Darwin Core termíny)
            meta.xml        — popis schématu
            eml.xml         — metadata datasetu
        """
        import csv, io as _io

        # Sestavit occurrence.csv
        occ_cols = [
            "occurrenceID", "scientificName", "taxonRank",
            "scientificNameAuthorship", "institutionCode", "catalogNumber",
            "typeStatus", "verbatimLocality", "verbatimEventDate",
            "occurrenceRemarks", "taxonRemarks", "namePublishedIn",
            "datasetName", "basisOfRecord", "kingdom", "phylum",
        ]
        occ_buf = _io.StringIO()
        writer = csv.DictWriter(occ_buf, fieldnames=occ_cols,
                                extrasaction="ignore",
                                lineterminator="\r\n",
                                quoting=csv.QUOTE_ALL)
        writer.writeheader()

        for row in rows:
            def _get(*paleon_fields: str) -> str:
                for pf in paleon_fields:
                    v = row.get(pf, "") or ""
                    if v and v != NOT_PROVIDED:
                        return v.replace("\n", " ").replace("\r", "")[:500]
                return ""

            dwc_row = {
                "occurrenceID":               _get("RECORD_ID"),
                "scientificName":             _get("TAXON_NAME_VERBATIM"),
                "taxonRank":                  _get("TAXON_RANK_AS_WRITTEN"),
                "scientificNameAuthorship":   _get("AUTHOR"),
                "institutionCode":            _get("INSTITUTION_CODE"),
                "catalogNumber":              _get("CATALOG_NUMBER"),
                "typeStatus":                 _get("TYPE_SPECIMEN_KIND", "TYPE SPECIMENS"),
                "verbatimLocality":           _get("LOCALITY"),
                "verbatimEventDate":          _get("STRATIGRAPHY"),
                "occurrenceRemarks":          _get("DESCRIPTION", "REMARKS"),
                "taxonRemarks":               _get("DIAGNOSIS"),
                "namePublishedIn":            _get("REFERENCE"),
                "datasetName":                _get("SOURCE_DOCUMENT"),
                "basisOfRecord":              "FossilSpecimen",
                "kingdom":                    "Animalia",
                "phylum":                     "Mollusca",
            }
            writer.writerow(dwc_row)

        # Zabalit do ZIP
        zip_buf = _io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("occurrence.csv",
                        ("\ufeff" + occ_buf.getvalue()).encode("utf-8"))
            zf.writestr("meta.xml", _DWC_META_XML.encode("utf-8"))
            zf.writestr("eml.xml",
                        _dwc_eml_xml(len(rows), dataset_name, creator).encode("utf-8"))
        return zip_buf.getvalue()



# ══════════════════════════════════════════════════════════════════════════════
# ROUND-TRIP XLSX EXPORT / IMPORT
# Export schválených záznamů do editovatelného XLSX + import zpět.
# ══════════════════════════════════════════════════════════════════════════════

# Pole metadat záznamu – editovatelná přímo v taxon_candidates
_RT_META_EDITABLE   = ("taxon_name", "rank_guess", "status")
_RT_META_LOCKED     = ("id", "document", "page_start")
# Pole obsahu – editovatelná přes occurrence_fields (vynecháme odvozená)
_RT_DERIVED_SKIP    = {"COMPLETENESS_SCORE", "MISSING_REQUIRED_FIELDS"}

def _rt_content_fields() -> List[str]:
    """Vrátí SECTION_FIELDS vhodná pro round-trip (bez odvozených sloupců)."""
    return [f for f in SECTION_FIELDS if f not in _RT_DERIVED_SKIP]


def export_records_editable_xlsx(candidate_ids: Tuple[int, ...]) -> bytes:
    """
    Exports selected candidate records to a round-trip XLSX file.
    - Gray columns (id, document, page_start) are locked/read-only.
    - White columns (taxon_name, rank_guess, status + content fields) are editable.
    - Import via import_records_from_xlsx() matches rows by the 'id' column.
    """
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl is not installed — run: pip install openpyxl")

    content_cols = _rt_content_fields()

    # ── sloupce: (header_text, is_locked) ─────────────────────────────────
    ALL_COLS: List[Tuple[str, bool]] = (
        [("id", True), ("taxon_name", False), ("rank_guess", False),
         ("status", False), ("document", True), ("page_start", True)]
        + [(f, False) for f in content_cols]
    )
    headers = [c for c, _ in ALL_COLS]

    # ── data z DB ─────────────────────────────────────────────────────────
    con = db()
    ph  = ",".join("?" * len(candidate_ids))
    cands = {
        r["id"]: dict(r)
        for r in con.execute(
            f"SELECT tc.id, tc.taxon_name, tc.rank_guess, tc.status, "
            f"tc.page_start, d.filename AS document "
            f"FROM taxon_candidates tc "
            f"JOIN documents d ON tc.document_id=d.id "
            f"WHERE tc.id IN ({ph})",
            candidate_ids).fetchall()
    }
    field_rows = con.execute(
        f"SELECT candidate_id, field_name, field_value "
        f"FROM occurrence_fields WHERE candidate_id IN ({ph})",
        candidate_ids).fetchall()
    con.close()

    fmap_all: Dict[int, Dict[str, str]] = {}
    for fr in field_rows:
        fmap_all.setdefault(fr["candidate_id"], {})[fr["field_name"]] = fr["field_value"] or ""

    # ── styles ────────────────────────────────────────────────────────────
    h_locked = PatternFill("solid", fgColor="1A3A5C")   # dark navy  – locked header
    h_edit   = PatternFill("solid", fgColor="1F6F8B")   # steel blue – editable header
    h_font   = Font(bold=True, color="FFFFFF", size=9)
    h_aln    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    d_locked = PatternFill("solid", fgColor="ECECEC")   # light gray  – locked data
    d_edit   = PatternFill("solid", fgColor="FFFFFF")   # white       – editable data
    d_aln    = Alignment(vertical="top", wrap_text=True)
    thin_s   = Side(style="thin", color="CCCCCC")
    brd      = Border(left=thin_s, right=thin_s, top=thin_s, bottom=thin_s)

    # ── workbook ──────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Records"

    # Řádek 1: záhlaví
    ws.append(headers)
    ws.row_dimensions[1].height = 36
    for i, (_, locked) in enumerate(ALL_COLS, 1):
        c = ws.cell(row=1, column=i)
        c.fill = h_locked if locked else h_edit
        c.font = h_font
        c.alignment = h_aln

    # Datové řádky
    for cid in candidate_ids:
        row_obj = cands.get(cid)
        if not row_obj:
            continue
        fmap = fmap_all.get(cid, {})
        vals = []
        for col_name, _ in ALL_COLS:
            if col_name in ("id", "taxon_name", "rank_guess", "status",
                            "document", "page_start"):
                v = row_obj.get(col_name, "")
                vals.append("" if v is None else v)
            else:
                v = fmap.get(col_name, "")
                vals.append("" if (not v or v == NOT_PROVIDED) else v)
        ws.append(vals)
        r = ws.max_row
        for i, (_, locked) in enumerate(ALL_COLS, 1):
            cell = ws.cell(row=r, column=i)
            cell.fill = d_locked if locked else d_edit
            cell.alignment = d_aln
            cell.border = brd

    # Šířky sloupců
    _wide = {"DESCRIPTION", "DIAGNOSIS", "SYNONYMY", "REMARKS",
             "MATERIAL EXAMINED", "TYPE SPECIMENS", "OCCURRENCE"}
    for i, (col_name, _) in enumerate(ALL_COLS, 1):
        if col_name == "id":               w = 8
        elif col_name == "taxon_name":     w = 30
        elif col_name == "document":       w = 34
        elif col_name in ("rank_guess",
                          "status",
                          "page_start"):   w = 14
        elif col_name in _wide:            w = 52
        else:                              w = 26
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    # Sheet 2: pokyny
    ws2 = wb.create_sheet("Instructions")
    ws2.column_dimensions["A"].width = 110
    ws2["A1"] = "PaleoN — Round-trip XLSX Editor"
    ws2["A1"].font = Font(bold=True, size=14, color="1A3A5C")
    ws2["A2"] = (f"Exported: {datetime.now().strftime('%Y-%m-%d %H:%M')}  │  "
                 f"Records: {len(candidate_ids)}")
    ws2["A2"].font = Font(italic=True, size=9)
    ws2["A4"] = "HOW TO USE"
    ws2["A4"].font = Font(bold=True, size=11)
    for row_n, line in enumerate([
        "1. Edit cells in the 'Records' sheet.",
        "2. GRAY columns (id, document, page_start) are read-only identifiers — do NOT modify them.",
        "3. WHITE/BLUE columns are editable: taxon_name, rank_guess, status, and all content fields.",
        "4. Valid values for 'status': approved | needs_review | rejected | pending",
        "5. Leave a cell EMPTY to keep the existing database value (empty = no change on import).",
        "6. Save the file, then upload it via PaleoN → Export tab → Round-trip XLSX editor → Import.",
        "7. After import, FTS5 index and morpho/strat term matches are recomputed automatically.",
    ], start=5):
        ws2[f"A{row_n}"] = line
        ws2[f"A{row_n}"].alignment = Alignment(wrap_text=True)
    ws2["A13"] = "TIP: Use Excel AutoFilter (already active) to sort/filter by taxon name, rank, etc."
    ws2["A13"].font = Font(italic=True, color="555555")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def import_records_from_xlsx(file_bytes: bytes) -> Dict[str, Any]:
    """
    Imports a XLSX file produced by export_records_editable_xlsx and updates
    the database records in-place.

    Returns:
        {
          'updated':        int,   # records with at least one change
          'skipped':        int,   # rows with no changes or missing from DB
          'fields_changed': int,   # total individual field updates
          'errors':         list,  # non-fatal warnings
        }
    """
    if not HAS_OPENPYXL:
        raise RuntimeError("openpyxl is not installed.")

    buf = io.BytesIO(file_bytes)
    wb = openpyxl.load_workbook(buf, data_only=True)
    if "Records" not in wb.sheetnames:
        raise ValueError("Sheet 'Records' not found. Upload a file exported by PaleoN.")

    ws = wb["Records"]
    raw_headers = [cell.value for cell in ws[1]]
    headers: List[Optional[str]] = [str(h).strip() if h is not None else None
                                     for h in raw_headers]

    if "id" not in headers:
        raise ValueError("Column 'id' not found. This file was not exported by PaleoN.")

    id_col   = headers.index("id")
    valid_sf = set(SECTION_FIELDS)
    stats: Dict[str, Any] = {
        "updated": 0, "skipped": 0, "fields_changed": 0, "errors": []
    }
    updated_cids: List[int] = []

    con = db()

    for row_vals in ws.iter_rows(min_row=2, values_only=True):
        # Prázdný řádek
        if all(v is None for v in row_vals):
            continue

        # Získat candidate id
        raw_id = row_vals[id_col] if id_col < len(row_vals) else None
        if raw_id is None:
            continue
        try:
            cid = int(raw_id)
        except (ValueError, TypeError):
            stats["errors"].append(f"Non-integer id {raw_id!r} — row skipped")
            continue

        # Kandidát v DB?
        existing = con.execute(
            "SELECT id, taxon_name, rank_guess, status "
            "FROM taxon_candidates WHERE id=?", (cid,)).fetchone()
        if not existing:
            stats["skipped"] += 1
            continue

        row_changed = False

        # ── 1) Aktualizace polí v taxon_candidates ────────────────────────
        meta_sets: Dict[str, str] = {}
        for col_name in _RT_META_EDITABLE:
            if col_name not in headers:
                continue
            col_i = headers.index(col_name)
            if col_i >= len(row_vals):
                continue
            new_v = row_vals[col_i]
            if new_v is None:
                continue
            new_s = str(new_v).strip()
            if not new_s:
                continue  # prázdné = nezměnit
            # Validace status
            if col_name == "status" and new_s not in (
                    "approved", "needs_review", "rejected", "pending"):
                stats["errors"].append(
                    f"id={cid}: invalid status '{new_s}' (allowed: approved|needs_review|rejected|pending)")
                continue
            old_s = str(existing[col_name] or "")
            if new_s != old_s:
                meta_sets[col_name] = new_s

        if meta_sets:
            set_sql = ", ".join(f"{k}=?" for k in meta_sets)
            con.execute(
                f"UPDATE taxon_candidates SET {set_sql} WHERE id=?",
                list(meta_sets.values()) + [cid])
            stats["fields_changed"] += len(meta_sets)
            row_changed = True

        # ── 2) Aktualizace occurrence_fields ──────────────────────────────
        existing_fields: Dict[str, int] = {
            r["field_name"]: r["id"]
            for r in con.execute(
                "SELECT id, field_name FROM occurrence_fields WHERE candidate_id=?",
                (cid,)).fetchall()
        }

        for col_i, col_name in enumerate(headers):
            if col_name not in valid_sf:
                continue
            if col_i >= len(row_vals):
                continue
            new_v = row_vals[col_i]
            if new_v is None:
                continue
            new_s = str(new_v).strip()
            if not new_s or new_s == NOT_PROVIDED:
                continue  # prázdné = nezměnit
            # Porovnat se stávající hodnotou
            if col_name in existing_fields:
                old_row = con.execute(
                    "SELECT field_value FROM occurrence_fields WHERE id=?",
                    (existing_fields[col_name],)).fetchone()
                if old_row and str(old_row["field_value"] or "").strip() == new_s:
                    continue  # beze změny
                con.execute(
                    "UPDATE occurrence_fields SET field_value=?, method=? WHERE id=?",
                    (new_s, "xlsx_import", existing_fields[col_name]))
            else:
                con.execute(
                    "INSERT INTO occurrence_fields (candidate_id,field_name,field_value,method) "
                    "VALUES (?,?,?,?)",
                    (cid, col_name, new_s, "xlsx_import"))
            stats["fields_changed"] += 1
            row_changed = True

        if row_changed:
            stats["updated"] += 1
            updated_cids.append(cid)
        else:
            stats["skipped"] += 1

    con.commit()
    con.close()

    # ── FTS5 + term_matches pro aktualizované záznamy ─────────────────────
    for cid in updated_cids:
        try:
            _fts_update_candidate(cid)
        except Exception as exc:
            logging.debug(f"FTS update failed for {cid}: {exc}")

    if updated_cids:
        try:
            compute_and_save_term_matches_batch(tuple(updated_cids))
        except Exception as exc:
            logging.warning(f"Term-match recompute failed after XLSX import: {exc}")

    return stats


def tab_export():
    st.header(t("export_header"))

    con = db()
    docs = con.execute("SELECT id, filename FROM documents ORDER BY created_at DESC").fetchall()
    all_approved = con.execute(
        "SELECT tc.*, d.filename, d.lang FROM taxon_candidates tc "
        "JOIN documents d ON tc.document_id=d.id WHERE tc.status='approved'"
    ).fetchall()
    con.close()

    if not docs:
        st.info(t("no_docs_dossier"))
        return

    # ── Výběr documents (zaškrtávací seznam s vyhledáváním) ────────────────────
    st.markdown("**Documents to export**")
    doc_counts: Dict[int, int] = {}
    for r in all_approved:
        doc_counts[r["document_id"]] = doc_counts.get(r["document_id"], 0) + 1

    sf1, sf2 = st.columns([3,1])
    doc_search = sf1.text_input(
        "🔍 Hledat Document", key="exp_doc_search",
        placeholder="filter by file name…", label_visibility="collapsed")
    select_all_docs = sf2.checkbox("☑️ All", value=True, key="exp_select_all_docs")
    if select_all_docs != st.session_state.get("_exp_select_all_prev", True):
        st.session_state.pop("exp_doc_table", None)
        st.session_state["_exp_select_all_prev"] = select_all_docs

    visible_docs = [
        d for d in docs
        if not doc_search or doc_search.lower() in d["filename"].lower()
    ]
    df_docs_exp = pd.DataFrame([{
        "_id": d["id"], "✓": select_all_docs,
        "File": d["filename"],
        "Approved": doc_counts.get(d["id"], 0),
    } for d in visible_docs])

    if df_docs_exp.empty:
        st.warning(t("no_doc_matches"))
        return

    edited_doc_table = st.data_editor(
        df_docs_exp.drop(columns=["_id"]),
        column_config={
            "✓": st.column_config.CheckboxColumn("✓", width=40),
            "File": st.column_config.TextColumn("File", width=320, disabled=True),
            "Approved": st.column_config.NumberColumn("Approved", width=90, disabled=True),
        },
        hide_index=True, use_container_width=True, num_rows="fixed",
        key="exp_doc_table",
    )
    selected_doc_ids = set(
        int(df_docs_exp.iloc[i]["_id"])
        for i, row in edited_doc_table.iterrows() if row.get("✓", False)
    )

    if not selected_doc_ids:
        st.warning(t("check_at_least_one"))
        return

    # ── Filtry kvality ───────────────────────────────────────────────────────
    e2, e3 = st.columns(2)
    min_conf_exp = e2.slider("Min. score", 0.0, 1.0, 0.0, 0.05, key="exp_conf")
    filter_rank_exp = e3.selectbox("Rank", ["— all —"]+RANK_OPTIONS[1:], key="exp_rank")

    candidates = [
        r for r in all_approved
        if r["document_id"] in selected_doc_ids
        and r["confidence"] >= min_conf_exp
        and (filter_rank_exp=="— all —" or r["rank_guess"]==filter_rank_exp)
    ]

    st.metric("Records to export", len(candidates))
    if not candidates:
        st.info(t("no_records_sel_filter"))
        return

    cand_ids_tuple = tuple(r["id"] for r in candidates)

    # ── Připravit řádky ───────────────────────────────────────────────────────
    @st.cache_data(ttl=30, show_spinner="Preparing data…")
    def _build_rows(cand_ids: Tuple[int, ...]) -> Tuple[List[Dict], List[Dict]]:
        rows, review_rows = [], []
        con2 = db()
        for cid in cand_ids:
            c = con2.execute("SELECT * FROM taxon_candidates WHERE id=?", (cid,)).fetchone()
            if not c: continue
            doc_row = con2.execute("SELECT * FROM documents WHERE id=?",
                                   (c["document_id"],)).fetchone()
            fields = get_candidate_fields(cid)
            if not fields and c["block_text"]:
                auto_mapped = map_sections_from_block(c["block_text"])
                if auto_mapped:
                    save_fields(cid, auto_mapped, method="regex_auto")
                    fields = get_candidate_fields(cid)
            rows.append(build_output_row(c, fields, doc_row))
        low_cands = con2.execute(
            "SELECT tc.*, d.filename FROM taxon_candidates tc "
            "JOIN documents d ON tc.document_id=d.id WHERE tc.status='low_confidence'"
        ).fetchall()
        for r in low_cands:
            if r["document_id"] in selected_doc_ids:
                review_rows.append({
                    "ID":r["id"],"Document":r["filename"],"Taxon":r["taxon_name"],
                    "Page":r["page_start"],"Confidence":r["confidence"],
                    "Status":r["status"],"Heading":r["heading_text"],"Notes":"",
                })
        con2.close()
        return rows, review_rows

    # ── Literatura / reference (zaškrtávací výběr) ────────────────────────────
    rows_preview, _ = _build_rows(cand_ids_tuple)
    unique_refs = sorted(set(
        r.get("REFERENCE", NOT_PROVIDED) for r in rows_preview
        if r.get("REFERENCE", NOT_PROVIDED) != NOT_PROVIDED
    ))
    selected_refs: List[str] = []
    if unique_refs:
        with st.expander(tt(f"📚 Literatura / reference ({len(unique_refs)} nalezeno)", f"📚 References ({len(unique_refs)} found)"), expanded=False):
            rf1, rf2 = st.columns([3,1])
            ref_search = rf1.text_input(
                "🔍 Search in literature", key="exp_ref_search",
                placeholder="filtrovat reference…", label_visibility="collapsed")
            ref_select_all = rf2.checkbox("☑️ All", value=False, key="exp_ref_select_all")
            if ref_select_all != st.session_state.get("_exp_ref_select_all_prev", False):
                st.session_state.pop("exp_ref_table", None)
                st.session_state["_exp_ref_select_all_prev"] = ref_select_all

            visible_refs = [r for r in unique_refs
                           if not ref_search or ref_search.lower() in r.lower()]
            df_refs = pd.DataFrame([
                {"✓": ref_select_all, "Reference": r[:200]} for r in visible_refs
            ])
            if not df_refs.empty:
                edited_refs = st.data_editor(
                    df_refs,
                    column_config={
                        "✓": st.column_config.CheckboxColumn("✓", width=40),
                        "Reference": st.column_config.TextColumn("Reference", disabled=True),
                    },
                    hide_index=True, use_container_width=True, num_rows="fixed",
                    key="exp_ref_table",
                )
                selected_refs = [
                    visible_refs[i] for i, row in edited_refs.iterrows()
                    if row.get("✓", False)
                ]
            else:
                st.caption(t("no_refs_match"))

    ts = datetime.now().strftime("%Y%m%d_%H%M")

    # ── Tlačítka exportu ──────────────────────────────────────────────────────
    st.markdown(t("formats_label"))
    annotate_opt = st.checkbox(
        "🏷 Category at the start of paragraph in RAW block (TXT export)",
        value=False, key="exp_annotate_raw")
    fc1,fc2,fc3,fc4,fc5,fc6,fc7 = st.columns(7)

    if fc1.button("⬇️ XLSX", type="primary"):
        rows, rrows = _build_rows(cand_ids_tuple)
        data = export_to_xlsx(rows, rrows, references=selected_refs or None)
        st.download_button(
            "📥 Download XLSX", data=data, file_name=f"paleon_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_xlsx")

    if fc2.button("⬇️ DOCX"):
        rows, _ = _build_rows(cand_ids_tuple)
        data = export_to_docx(rows)
        st.download_button(
            "📥 Download DOCX", data=data, file_name=f"paleon_{ts}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_docx")

    if fc3.button("⬇️ JSON"):
        rows, _ = _build_rows(cand_ids_tuple)
        data = export_to_json(rows)
        st.download_button(
            "📥 Download JSON", data=data, file_name=f"paleon_{ts}.json",
            mime="application/json", key="dl_json")

    if fc4.button("⬇️ TXT"):
        rows, _ = _build_rows(cand_ids_tuple)
        data = export_to_txt(rows, annotate_raw=annotate_opt, references=selected_refs or None)
        st.download_button(
            "📥 Download TXT", data=data, file_name=f"paleon_{ts}.txt",
            mime="text/plain", key="dl_txt")

    pdf_lbl  = "⬇️ PDF" if HAS_FPDF else "⬇️ PDF (HTML)"
    pdf_help = "Export via fpdf2" if HAS_FPDF else \
               "fpdf2 is not installed – will generate HTML for printing"
    if fc5.button(pdf_lbl, help=pdf_help):
        rows, _ = _build_rows(cand_ids_tuple)
        data = export_to_pdf(rows)
        if HAS_FPDF:
            st.download_button("📥 Download PDF", data=data,
                file_name=f"paleon_{ts}.pdf", mime="application/pdf", key="dl_pdf")
        else:
            st.download_button("📥 Download HTML (tisk→PDF)", data=data,
                file_name=f"paleon_{ts}_print.html", mime="text/html", key="dl_html")
            st.info(t("open_print_save"))

    # BOD 6: CSV export
    if fc6.button("⬇️ CSV",
                  help="Comma-Separated Values s UTF-8 BOM (Excel kompatibilni)"):
        rows, _ = _build_rows(cand_ids_tuple)
        data = export_to_csv(rows)
        st.download_button(
            "📥 Download CSV", data=data,
            file_name=f"paleon_{ts}.csv",
            mime="text/csv; charset=utf-8-sig",
            key="dl_csv")

    # BOD 8: Darwin Core Archive
    if fc7.button("⬇️ DwC-A",
                  help="Darwin Core Archive — ZIP for primy import to GBIF"):
        rows, _ = _build_rows(cand_ids_tuple)
        con_dwc = db()
        _dwc_docs = con_dwc.execute("SELECT filename FROM documents LIMIT 1").fetchone()
        con_dwc.close()
        _ds_name = f"PaleoN export {ts}"
        data = export_to_dwca(rows, dataset_name=_ds_name)
        st.download_button(
            "📥 Download DwC-A (ZIP)", data=data,
            file_name=f"paleon_{ts}_dwca.zip",
            mime="application/zip",
            key="dl_dwca")
        st.caption("ZIP obsahuje occurrence.csv + meta.xml + eml.xml kompatibilni s GBIF IPT.")

    # ── Round-trip XLSX editor ───────────────────────────────────────────────
    st.divider()
    st.subheader(t("roundtrip_sub"))
    st.caption(t("roundtrip_caption"))

    rt1, rt2 = st.columns([1, 2])
    with rt1:
        if st.button(t("export_edit_xlsx_btn"), key="rt_export_btn", type="primary",
                     help="Exports the currently selected & filtered records to an editable XLSX."):
            with st.spinner("Preparing XLSX…"):
                rt_data = export_records_editable_xlsx(cand_ids_tuple)
            st.download_button(
                "📥 Download editable XLSX",
                data=rt_data,
                file_name=f"paleon_edit_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="rt_dl_xlsx")
        st.caption(
            f"Will export **{len(candidates)}** records "
            f"(same filter as the export buttons above).")

    with rt2:
        rt_upload = st.file_uploader(
            t("import_xlsx_label"),
            type=["xlsx"], key="rt_import_up",
            help="Upload a XLSX file that was previously exported by the 'Export for editing' button.")
        if rt_upload:
            if st.button(t("import_xlsx_btn"), key="rt_import_btn", type="primary"):
                with st.spinner("Importing records…"):
                    try:
                        rt_result = import_records_from_xlsx(rt_upload.getbuffer())
                        n = rt_result["updated"]
                        f_ch = rt_result["fields_changed"]
                        s    = rt_result["skipped"]
                        errs = rt_result.get("errors", [])
                        if n > 0:
                            st.success(
                                t("import_ok").format(n=n, f=f_ch) + "  " +
                                t("import_skipped").format(s=s))
                        else:
                            st.info(t("import_nothing") + " " +
                                    t("import_skipped").format(s=s))
                        if errs:
                            st.warning(t("import_errors").format(e=len(errs)) +
                                       "\n" + "\n".join(f"• {e}" for e in errs[:10]))
                    except ValueError as ve:
                        st.error(str(ve))
                    except Exception as exc:
                        st.error(f"Import failed: {exc}")
                        logging.exception("XLSX import error")

        # ── Náhled ───────────────────────────────────────────────────────────────
    if st.checkbox(t("show_table_preview")):
        rows, _ = _build_rows(cand_ids_tuple)
        preview = ["RECORD_ID","TAXON_NAME_VERBATIM","TAXON_RANK_AS_WRITTEN",
                   "SOURCE PAGES","DIAGNOSIS","DESCRIPTION",
                   "LOCALITY","STRATIGRAPHY"]
        df_prev = pd.DataFrame(rows)[preview]
        st.dataframe(df_prev, use_container_width=True)

    if st.checkbox(t("show_field_stats")):
        rows, _ = _build_rows(cand_ids_tuple)
        if rows:
            fill_data = {}
            for f in SECTION_FIELDS:
                fill_data[f] = sum(1 for r in rows if r.get(f,NOT_PROVIDED) != NOT_PROVIDED)
            df_fill = pd.DataFrame({"pole":list(fill_data.keys()),
                                    "filled":list(fill_data.values())}).sort_values(
                                        "filled", ascending=True)
            st.bar_chart(df_fill.set_index("pole")["filled"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 – TAXON DOSSIER
# ══════════════════════════════════════════════════════════════════════════════

def _render_hierarchy_badge(cand_row) -> str:
    """Returns an HTML string with the hierarchical breadcrumb for a candidate."""
    parts = []
    if "parent_taxon_name" in cand_row.keys() and cand_row["parent_taxon_name"]:
        pr = cand_row["parent_rank"] or "?"
        pn = cand_row["parent_taxon_name"]
        parts.append(
            f"<span style='font-size:0.70rem;opacity:0.7'>"
            f"{pr.title()}: <i>{pn}</i> ›</span> "
        )
    return "".join(parts)


def tab_taxon_dossier():
    # ── CSS pro dense 4K layout ───────────────────────────────────────────────
    st.markdown("""
    <style>
    .pn-dos-wrap{font-family:inherit}
    .pn-field-grid{display:grid;grid-template-columns:155px 1fr;
        column-gap:10px;row-gap:1px;margin:0}
    .pn-fl{font-size:0.73rem;font-weight:700;
        color:var(--pn-label-color,#374151);
        text-transform:uppercase;letter-spacing:.04em;
        padding:3px 0;line-height:1.2;word-break:break-word}
    .pn-fv{font-size:0.87rem;line-height:1.35;padding:3px 0;
        color:var(--pn-value-color,#111827);
        border-bottom:1px solid rgba(100,116,139,0.12)}
    .pn-taxhead{font-size:1.05rem;font-weight:700;margin:0 0 2px;padding:0}
    .pn-meta{font-size:0.77rem;color:var(--pn-meta-color,#374151);
        line-height:1.4;margin:0 0 4px}
    .pn-badge{display:inline-block;padding:2px 7px;border-radius:3px;
        font-size:0.74rem;font-weight:700;margin-right:4px}
    .pn-crumb{font-size:0.77rem;color:var(--pn-meta-color,#374151);
        margin:0 0 6px;border-left:2px solid #64748b;padding-left:6px}
    .pn-sec-label{font-size:0.80rem;font-weight:700;
        color:var(--pn-label-color,#374151);
        margin:8px 0 2px;text-transform:uppercase;letter-spacing:.05em}
    .pn-occ-sep{border:none;border-top:1px solid rgba(100,116,139,0.18);
        margin:8px 0}
    .pn-fill-bar-wrap{background:rgba(100,116,139,0.2);border-radius:3px;
        height:5px;width:100%;margin:4px 0}
    .pn-fill-bar{height:5px;border-radius:3px;transition:width .3s}
    /* theme detection via forced-colors or prefers-color-scheme */
    @media (prefers-color-scheme:light){
        .pn-fl{color:#1e293b}
        .pn-fv{color:#0f172a}
        .pn-meta{color:#334155}
        .pn-crumb{color:#334155}
    }
    div[data-testid="stTextInput"]>div>div>input{font-size:0.87rem!important}
    /* Radio list compact */
    div[data-testid="stRadio"] label{font-size:0.82rem!important;padding:2px 0!important}
    div[data-testid="stRadio"] div[role="radiogroup"]{gap:1px!important}
    </style>
    """, unsafe_allow_html=True)

    s = get_settings()
    con = db()
    total_docs = con.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
    con.close()
    if total_docs == 0:
        st.info(t("empty_library"))
        return

    # ── Horní lišta: vyhledávání + filtry ────────────────────────────────────
    hc1, hc2, hc3, hc4, hc5 = st.columns([5, 1, 1, 1, 1])
    query = hc1.text_input(
        "🔍", key="td_query",
        placeholder="Search taxon or fulltext (min. 1 znak)…",
        label_visibility="collapsed")
    fulltext_mode = hc2.toggle("Fulltext", value=False, key="td_fulltext",
                                help="Search in all fields too (slower)")
    include_low  = hc3.toggle("Low conf", value=False, key="td_low",
                               help="Include low-confidence records")
    annotate_all = hc4.toggle("Annotate", value=False, key="td_ann",
                               help="[CATEGORY] labels in block")
    show_raw     = hc5.toggle("Raw blok", value=True, key="td_raw")

    # ── Načíst allchny kandidáty (jednou, pak filtrovat v Pythonu) ────────────
    statuses = (["approved","pending","needs_review","low_confidence"]
                if include_low else ["approved","pending","needs_review"])
    sp = ",".join("?"*len(statuses))
    con = db()
    all_cands = con.execute(
        f"SELECT tc.*, d.filename, d.lang "
        f"FROM taxon_candidates tc "
        f"JOIN documents d ON tc.document_id=d.id "
        f"WHERE tc.status IN ({sp}) "
        f"ORDER BY tc.taxon_name, d.filename, tc.page_start",
        statuses).fetchall()
    con.close()

    # ── Build: seznam unikátních jmen + počty (abecedně) ────────────────────
    needle = _norm_search(query.strip()) if query and len(query.strip()) >= 1 else ""

    # Wildcard podpora v hledání jmen taxa: * → .* , ? → .
    def _dos_match_name(taxon_name: str) -> bool:
        if not needle:
            return True
        _n = _norm_search(taxon_name or "")
        if "*" in needle or "?" in needle:
            try:
                rx = re.compile(
                    needle.replace(".", r"\.").replace("*", ".*").replace("?", "."),
                    re.IGNORECASE)
                return bool(rx.search(_n))
            except re.error:
                return needle in _n
        return needle in _n

    if needle and fulltext_mode:
        # Fulltext: zkusit FTS5, fallback na LIKE
        _fts_ids = set(fts_search(query.strip(), limit=1000))
        if _fts_ids:
            # FTS5 uspělo — kombinovat s name matchem
            filtered = [
                c for c in all_cands
                if c["id"] in _fts_ids
                or _dos_match_name(c["taxon_name"] or "")
            ]
            _fts_method = "FTS5"
        else:
            # Fallback: LIKE scan (pomalejší, ale vždy funkční)
            _ft_con = db()
            _ft_sql = f"%{query.strip()}%"
            _ft_field_ids = set(
                r[0] for r in _ft_con.execute(
                    "SELECT DISTINCT candidate_id FROM occurrence_fields "
                    "WHERE field_value LIKE ? AND field_value != ?",
                    (_ft_sql, NOT_PROVIDED)).fetchall())
            _ft_block_ids = set(
                r[0] for r in _ft_con.execute(
                    "SELECT id FROM taxon_candidates WHERE block_text LIKE ?",
                    (_ft_sql,)).fetchall())
            _ft_con.close()
            _fts_ids = _ft_field_ids | _ft_block_ids
            filtered = [
                c for c in all_cands
                if _dos_match_name(c["taxon_name"] or "")
                or c["id"] in _fts_ids
            ]
            _fts_method = f"LIKE ({len(_ft_field_ids)} fields + {len(_ft_block_ids)} blocks)"
        if filtered:
            st.caption(
                f"Fulltext \"{query.strip()}\": "
                f"{len(_fts_ids)} shod — metoda: {_fts_method}")
    elif needle:
        filtered = [c for c in all_cands if _dos_match_name(c["taxon_name"] or "")]
    else:
        filtered = all_cands

    # Seskupit podle normalizovaného jména
    groups: Dict[str, list] = {}
    for c in filtered:
        key = re.sub(r"\s+", " ", (c["taxon_name"] or "").strip())
        groups.setdefault(key, []).append(c)

    # Seřadit abecedně
    sorted_names = sorted(groups.keys())

    if not sorted_names:
        if query:
            st.warning(tt(f"No taxon matches \"{query}\".", f"No taxon matches \"{query}\"."))
        else:
            con2 = db()
            sample = con2.execute(
                "SELECT taxon_name FROM taxon_candidates "
                "WHERE status IN ('approved','pending') "
                "ORDER BY taxon_name LIMIT 10").fetchall()
            con2.close()
            if sample:
                st.caption(tt("Example taxa: ", "Sample taxa: ") +
                           "  ·  ".join(f"`{r['taxon_name']}`" for r in sample))
            st.info(tt("Insert part name of taxon for searching.", "Enter part of the taxon name to search."))
        return

    # ── Metriky celkem ────────────────────────────────────────────────────────
    m1, m2, m3, m4 = st.columns(4)
    m1.metric(tt("Nalezeno taxa", "Taxa found"), len(sorted_names))
    m2.metric(tt("Total occurrences", "Total occurrences"), len(filtered))
    m3.metric(t("documents"), len(set(c["filename"] for c in filtered)))
    avg_conf = round(sum(c["confidence"] for c in filtered)/max(len(filtered),1), 3)
    m4.metric("Average confidence", f"{avg_conf:.3f}")

    # ── FIELD_ORDER pro zobrazení fields ───────────────────────────────────────
    FIELD_ORDER = [
        "TAXONOMIC PLACEMENT","AUTHOR","SYNONYMY",
        "NOMENCLATURAL ACTS","OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS",
        "DIAGNOSIS","DESCRIPTION","SIZE","SIZE_PARSED","ETYMOLOGY",
        "TYPE TAXON","TYPE SPECIMENS","TYPE_SPECIMEN_KIND",
        "INSTITUTION_CODE","CATALOG_NUMBER","TYPE MATERIAL",
        "MATERIAL EXAMINED","LOCALITY","STRATIGRAPHY",
        "OCCURRENCE","FIGURES","REMARKS","REFERENCE",
        "INCLUDED TAXONS","TAXON","PARENT_TAXON_NAME",
    ]

    # ── Výběr aktuálního taxonu (session state) ──────────────────────────────
    if "td_selected" not in st.session_state or \
       st.session_state["td_selected"] not in sorted_names:
        st.session_state["td_selected"] = sorted_names[0]

    # ── Dvousloupcový layout: seznam vlevo | detail vpravo ───────────────────
    list_col, detail_col = st.columns([1, 3], gap="large")

    # ── Levý sloupec: abecední seznam taxa ─────────────────────────────────
    with list_col:
        st.markdown(
            f"<div style='font-size:0.70rem;color:#64748b;margin-bottom:4px'>"
            f"<b>{len(sorted_names)}</b> taxa — click a row for detail</div>",
            unsafe_allow_html=True)

        # Sestavit DataFrame pro přehlednou tabulku s výběrem řádku
        _td_df = pd.DataFrame([{
            "Taxon": n,
            "n": len(groups[n]),
            "rank": (groups[n][0]["rank_guess"] or "?") if groups[n] else "?",
        } for n in sorted_names])

        _td_cur_idx = sorted_names.index(st.session_state["td_selected"])
        _td_sel = st.dataframe(
            _td_df,
            column_config={
                "Taxon": st.column_config.TextColumn("Taxon", width="large"),
                "n":     st.column_config.NumberColumn("n", width=40),
                "rank":  st.column_config.TextColumn("Rank"),
            },
            hide_index=True,
            use_container_width=True,
            height=min(38 * len(sorted_names) + 38, 560),
            on_select="rerun",
            selection_mode="single-row",
            key="td_taxon_table",
        )
        # Přečíst výběr — pokud uživatel klikl na řádek, přejít na daný taxon
        _sel_rows = getattr(_td_sel, "selection", {})
        if isinstance(_sel_rows, dict):
            _sel_rows = _sel_rows.get("rows", [])
        elif hasattr(_sel_rows, "rows"):
            _sel_rows = list(_sel_rows.rows)
        else:
            _sel_rows = []
        if _sel_rows:
            selected_name = sorted_names[_sel_rows[0]]
            st.session_state["td_selected"] = selected_name
        else:
            selected_name = st.session_state["td_selected"]

    # ── Pravý sloupec: detail vybraného taxonu ────────────────────────────────
    with detail_col:
        occs = groups.get(selected_name, [])
        if not occs:
            st.info("Select taxon ze seznamu.")
            st.stop()

        # ── Záhlaví ──────────────────────────────────────────────────────────
        all_fields_across_docs = {}
        for occ in occs:
            f = get_candidate_fields(occ["id"])
            for k, v in f.items():
                if v and v != NOT_PROVIDED and k not in all_fields_across_docs:
                    all_fields_across_docs[k] = v

        comp_score, comp_missing = _completeness_check(
            occs[0]["rank_guess"] or "", all_fields_across_docs)
        comp_pct = int(comp_score * 100)
        comp_color = "#10b981" if comp_pct>=70 else "#f59e0b" if comp_pct>=40 else "#ef4444"

        # Parent breadcrumb
        parent_name = ""
        parent_rank = ""
        if "parent_taxon_name" in occs[0].keys():
            parent_name = occs[0]["parent_taxon_name"] or ""
            parent_rank = occs[0]["parent_rank"] or ""

        crumb_html = ""
        if parent_name:
            crumb_html = (
                f"<div class='pn-crumb'>"
                f"{parent_rank.title() if parent_rank else '?'}: "
                f"<i>{parent_name}</i> ›</div>")

        rank_badge_color = {
            "species":"#1d4ed8","genus":"#7c3aed","family":"#b45309",
            "order":"#166534","class":"#831843",
        }.get((occs[0]["rank_guess"] or "").lower(), "#475569")

        statuses_all = sorted(set(o["status"] for o in occs))
        _STAT_COLORS = {"approved":"#065f46","pending":"#1e3a5f",
                        "needs_review":"#78350f"}
        status_html = "".join(
            f"<span class='pn-badge' style='background:{_STAT_COLORS.get(s,'#374151')};color:#e2e8f0'>{s}</span>"
            for s in statuses_all)

        st.markdown(
            f"<div class='pn-dos-wrap'>{crumb_html}"
            f"<p class='pn-taxhead'>"
            f"<span class='pn-badge' style='background:{rank_badge_color};color:#fff'>"
            f"{occs[0]['rank_guess'] or '?'}</span>"
            f" {selected_name}</p>"
            f"<p class='pn-meta'>"
            f"{status_html}"
            f"&nbsp;{len(occs)}× &nbsp;·&nbsp; "
            f"{len(set(o['filename'] for o in occs))} dok."
            f"&nbsp;·&nbsp; avg.conf {avg_conf:.3f}"
            f"&nbsp;·&nbsp; <span style='color:{comp_color}'>completeness {comp_pct}%</span>"
            + (f"&nbsp;·&nbsp;<span style='color:#f59e0b'>missing: {', '.join(comp_missing[:3])}</span>"
               if comp_missing else "")
            + "</p>"
            f"<div class='pn-fill-bar-wrap'><div class='pn-fill-bar' "
            f"style='width:{comp_pct}%;background:{comp_color}'></div></div>"
            f"</div>",
            unsafe_allow_html=True)

        # ── PDF odkaz (první výskyt) vedle záložek ─────────────────────────
        _first_occ = sorted(occs, key=lambda x: (x["filename"], x["page_start"]))[0]
        _tdck = f"docpath_{_first_occ['document_id']}"
        if _tdck not in st.session_state:
            _ccon = db(); _cr = _ccon.execute(
                "SELECT path FROM documents WHERE id=?",
                (_first_occ["document_id"],)).fetchone()
            _ccon.close()
            st.session_state[_tdck] = (_cr["path"] if _cr else None)
        # (_pdf_link_html odpagesěn — file:// URLs nefungují z Streamlit HTTP Contextu;
        #  PDF se nyní zobrazuje inline v tab 📜 Raw + PDF)
        # ── Záložky detail panelu ─────────────────────────────────────────────
        has_multi = len(set(o["filename"] for o in occs)) > 1
        _tab_labels = ["📋 Pole", "🧬 Systematika", "📜 Raw + PDF"]
        if has_multi:
            _tab_labels.append("🔀 Mezi Documenty")
        detail_tabs = st.tabs(_tab_labels)
        tab_fields   = detail_tabs[0]
        tab_syst     = detail_tabs[1]
        tab_block    = detail_tabs[2]
        tab_compare  = detail_tabs[3] if has_multi else None

        with tab_fields:
            # Banner po úspěšném překladu (zobrazí se jednou po rerunu)
            _just_tr = st.session_state.pop(f"dos_just_translated_{selected_name}", None)
            if _just_tr:
                st.success(
                    f"✅ Translated {_just_tr} fields — zde jsou aktualizovane hodnoty",
                    icon="🌐")

            # Sloučená pole ze allch occurrences, zobrazeno v dense gridu
            all_fv_display = []
            for fname in FIELD_ORDER:
                if fname in all_fields_across_docs:
                    all_fv_display.append((fname, all_fields_across_docs[fname]))
            extra = [(k,v) for k,v in all_fields_across_docs.items()
                     if k not in FIELD_ORDER]
            all_fv_display += extra

            if all_fv_display:
                # ── Tlačítka: Auto-map + Přeložit do AJ ──────────────────────
                _dos_btn_c1, _dos_btn_c2, _dos_btn_c3 = st.columns([2, 2, 3])
                if _dos_btn_c1.button("🔄 Auto-map all",
                                      key=f"td_am2_{selected_name}",
                                      help="Map again fields from RAW block."):
                    n_total = 0
                    for occ in occs:
                        _blk = occ["block_text"] or extract_block_for_candidate(
                            occ["id"], occ["document_id"], occ["page_start"])
                        if _blk:
                            n_total += auto_map_candidate_fields(occ["id"], _blk)
                    st.success(tt(f"Mapped {n_total} fields.", f"Mapped {n_total} fields."))
                    st.rerun()

                _dos_s = get_settings()
                _dos_llm_ok = _dos_s.get("llm_enabled")

                # ── Hromadný překlad: VŠECHNY výskyty taxonu ──────────────
                # Sbíráme jazyky ze allch zdrojových documents
                _dos_langs: Dict[int, str] = {}  # document_id → lang
                for _oc in occs:
                    if _oc["document_id"] not in _dos_langs:
                        _dl_r = db().execute(
                            "SELECT lang FROM documents WHERE id=?",
                            (_oc["document_id"],)).fetchone()
                        _dos_langs[_oc["document_id"]] = (
                            (_dl_r["lang"] if _dl_r else "") or "")
                _dos_unique_langs = sorted(set(
                    v for v in _dos_langs.values() if v))
                _dos_lang_badge = (
                    ", ".join(l.upper() for l in _dos_unique_langs)
                    if _dos_unique_langs else "?"
                )
                _dos_tr_help = (
                    f"Translates fields of ALL {len(occs)} výskytu "
                    f"(jazyky: {_dos_lang_badge}) do angličtiny. "
                    "Každý výskyt se přeloží zvlášť."
                    if _dos_llm_ok else
                    "Requires LM Studio — enable it in Settings → LLM.")
                if _dos_btn_c2.button(
                        f"🌐 Translate →EN{f' ({_dos_lang_badge})' if _dos_lang_badge != '?' else ''}",
                        key=f"td_tr_{selected_name}",
                        disabled=not _dos_llm_ok,
                        help=_dos_tr_help):
                    _tr_total = 0
                    _tr_errors: List[str] = []
                    with st.spinner(
                        f"Translating {len(occs)} occurrences using LM Studio…"):
                        for _oc in occs:
                            _oc_lang = _dos_langs.get(_oc["document_id"], "")
                            try:
                                _n = auto_translate_candidate_fields(
                                    _oc["id"], _oc_lang, _dos_s, force=True)
                                _tr_total += _n
                                remap_fields_post_translation(_oc["id"])
                            except RuntimeError as _e:
                                _tr_errors.append(str(_e))
                                break
                            except Exception as _e:
                                _tr_errors.append(str(_e)[:100])
                    if _tr_errors:
                        st.error(tt(f"❌ Translation error: {_tr_errors[0]}", f"❌ Translation error: {_tr_errors[0]}"))
                    elif _tr_total:
                        # Po překladu přepnout na záložku Pole →
                        # saveíme flag, který po rerunu zobrazí banner + zvýrazní pole
                        st.session_state[f"dos_just_translated_{selected_name}"] = _tr_total
                        st.rerun()
                    else:
                        st.info(
                            "Zadna fields k prekladu — jsou vyskyty namapovane? Languages spravne nastaveny v documents?")

                # Rozdělit do dvou sloupcu pro 4K využití
                mid = (len(all_fv_display)+1)//2
                col_a, col_b = st.columns(2, gap="large")
                for ci, (col, chunk) in enumerate([(col_a, all_fv_display[:mid]),
                                                    (col_b, all_fv_display[mid:])]):
                    with col:
                        rows_html = ""
                        for fname, fval in chunk:
                            safe_val = str(fval).replace("<","&lt;").replace(">","&gt;")
                            # Ikony pro vybraná pole
                            _ficons = {
                                "STRATIGRAPHY": "🪨", "OCCURRENCE": "🌍",
                                "LOCALITY": "📍", "TYPE SPECIMENS": "🔬",
                                "TYPE MATERIAL": "🔬", "TYPE TAXON": "🎯",
                                "DIAGNOSIS": "📝", "DESCRIPTION": "📝",
                                "AUTHOR": "✍️", "SYNONYMY": "📜",
                                "FIGURES": "🖼️", "ETYMOLOGY": "🔤",
                                "SIZE": "📐", "REMARKS": "💬",
                                "YEAR_OF_PUBLICATION": "📅",
                                "INCLUDED TAXONS": "📋",
                                "MATERIAL EXAMINED": "🧪",
                            }
                            _ficon = _ficons.get(fname, "")
                            _flabel = f"{_ficon}&nbsp;{fname}" if _ficon else fname
                            rows_html += (
                                f"<div class='pn-field-grid'>"
                                f"<div class='pn-fl'>{_flabel}</div>"
                                f"<div class='pn-fv'>{safe_val}</div>"
                                f"</div>")
                        st.markdown(rows_html, unsafe_allow_html=True)
            else:
                st.caption(tt("No filled fields. Run Auto-map or fill in manually.", "No filled fields. Run Auto-map or fill manually."))
                if st.button(t("auto_map_all"), key=f"td_am_{selected_name}"):
                    n_total = 0
                    for occ in occs:
                        block = occ["block_text"] or extract_block_for_candidate(
                            occ["id"], occ["document_id"], occ["page_start"])
                        if block:
                            n_total += auto_map_candidate_fields(occ["id"], block)
                    st.success(tt(f"Mapped {n_total} fields.", f"Mapped {n_total} fields."))
                    st.rerun()

        if tab_compare:
            with tab_compare:
                # Tabulka: řádky = pole, sloupce = Documenty
                doc_names = sorted(set(o["filename"] for o in occs))
                # Batch načíst allchna pole najednou (1 dotaz místo N)
                _bulk_fields = get_candidates_fields_bulk([o["id"] for o in occs])
                doc_short = [pathlib.Path(dn).stem[:30] for dn in doc_names]
                occ_by_doc: Dict[str, list] = {}
                for o in occs:
                    occ_by_doc.setdefault(o["filename"], []).append(o)

                # Nojduležitější pole pro porovnání
                CMP_FIELDS = ["DIAGNOSIS","DESCRIPTION","TYPE SPECIMENS",
                              "LOCALITY","STRATIGRAPHY","OCCURRENCE",
                              "SIZE","SYNONYMY","AUTHOR","FIGURES"]
                tbl_html = (
                    "<table style='width:100%;border-collapse:collapse;"
                    "font-size:0.76rem'><thead><tr>"
                    "<th style='text-align:left;padding:3px 6px;border-bottom:"
                    "2px solid #475569;font-size:0.68rem;color:#94a3b8;min-width:120px'>"
                    "POLE</th>")
                for dn in doc_short:
                    tbl_html += (
                        f"<th style='text-align:left;padding:3px 6px;"
                        f"border-bottom:2px solid #475569;font-size:0.68rem;"
                        f"color:#94a3b8'>{dn}</th>")
                tbl_html += "</tr></thead><tbody>"
                for fi, fn in enumerate(CMP_FIELDS):
                    row_bg = "rgba(71,85,105,0.08)" if fi%2==0 else "transparent"
                    tbl_html += f"<tr style='background:{row_bg}'>"
                    tbl_html += (
                        f"<td style='padding:3px 6px;font-weight:700;"
                        f"color:#64748b;font-size:0.66rem;text-transform:uppercase;"
                        f"vertical-align:top'>{fn}</td>")
                    for doc_name in doc_names:
                        doc_occs = occ_by_doc.get(doc_name, [])
                        cell_val = ""
                        for do in doc_occs:
                            # Použít bulk cache (vypočítán před tabulkou)
                            flds = _bulk_fields.get(do["id"], {})
                            v = flds.get(fn,"")
                            if v and v != NOT_PROVIDED:
                                cell_val = v[:200]
                                break
                        safe_cell = cell_val.replace("<","&lt;").replace(">","&gt;")
                        color = "#e2e8f0" if safe_cell else "#4b5563"
                        tbl_html += (
                            f"<td style='padding:3px 6px;color:{color};"
                            f"vertical-align:top;border-left:1px solid "
                            f"rgba(71,85,105,0.15)'>"
                            f"{'<i style=color:#4b5563>—</i>' if not safe_cell else safe_cell}"
                            f"{'…' if len(cell_val)>=200 else ''}</td>")
                    tbl_html += "</tr>"
                tbl_html += "</tbody></table>"
                st.markdown(tbl_html, unsafe_allow_html=True)

        # ══════════════════════════════════════════════════════════════════════
        # 🧬 SYSTEMATIKA
        # ══════════════════════════════════════════════════════════════════════
        with tab_syst:
            # ── Systematické zařazení ─────────────────────────────────────────
            _sys_fields = all_fields_across_docs
            _tax_placement = _sys_fields.get("TAXONOMIC PLACEMENT", "")
            _type_taxon    = _sys_fields.get("TYPE TAXON", "")
            _incl_taxa     = _sys_fields.get("INCLUDED TAXONS", "")
            _synonymy      = _sys_fields.get("SYNONYMY", "")
            _author_val    = _sys_fields.get("AUTHOR", "")
            _year_val      = _sys_fields.get("YEAR_OF_PUBLICATION", "")
            _status_val    = _sys_fields.get("OPEN NOMENCLATURE / IDENTIFICATION QUALIFIERS", "")
            _occ_rank      = (occs[0]["rank_guess"] or "") if occs else ""

            # ── Systematické zařazení / breadcrumb ───────────────────────────
            st.markdown(tt("**📍 Systematicke zarazeni**", "**📍 Systematic placement**"))
            if _tax_placement and _tax_placement != NOT_PROVIDED:
                st.markdown(
                    f"<div style='font-size:0.88rem;padding:6px 10px;"
                    f"background:rgba(100,116,139,0.08);border-radius:5px;"
                    f"border-left:3px solid #64748b;margin-bottom:8px'>"
                    f"{_tax_placement}</div>",
                    unsafe_allow_html=True)
            else:
                # Zkus sestavit z dat DB: najdi záznamy se stejným rankem o řád výše
                _rank_order = ["Species","Genus","Family","Order","Class","Phylum"]
                _cur_rank_i = _rank_order.index(_occ_rank) if _occ_rank in _rank_order else -1
                _breadcrumb_parts = []
                if _cur_rank_i > 0:
                    _parent_ranks = _rank_order[_cur_rank_i:]
                    _con_bc = db()
                    for _pr in _parent_ranks:
                        # SQLite nepodporuje != ALL(subquery) — použijeme taxon_name != ?
                        _bc_r = _con_bc.execute(
                            "SELECT DISTINCT taxon_name FROM taxon_candidates "
                            "WHERE rank_guess=? AND status IN ('approved','pending') "
                            "AND taxon_name != ? "
                            "LIMIT 1",
                            (_pr, selected_name)).fetchone()
                        if _bc_r:
                            _breadcrumb_parts.append(f"*{_bc_r[0]}* ({_pr})")
                    _con_bc.close()
                if _breadcrumb_parts:
                    st.caption(" › ".join(reversed(_breadcrumb_parts)) + f" › **{selected_name}**")
                else:
                    st.caption(t("no_placement"))

            # ── Autor a rok ───────────────────────────────────────────────────
            _auth_display = " ".join(filter(None, [_author_val, _year_val])).strip()
            if not _auth_display or _auth_display == NOT_PROVIDED:
                # Zkus extrahovat z jména taxonu (pokud obsahuje autora)
                _am = AUTHOR_YEAR_RE.search(selected_name)
                if _am:
                    _auth_display = _am.group(0)
            if _auth_display and _auth_display != NOT_PROVIDED:
                st.markdown(
                    f"<div style='font-size:0.9rem;margin:4px 0'>"
                    f"<b>Autor:</b> {_auth_display}</div>",
                    unsafe_allow_html=True)

            # ── Status ────────────────────────────────────────────────────────
            if _status_val and _status_val != NOT_PROVIDED:
                st.markdown(
                    f"<div style='font-size:0.88rem;color:#f59e0b;margin:2px 0'>"
                    f"⚠️ <b>Nomenklatura:</b> {_status_val}</div>",
                    unsafe_allow_html=True)

            st.divider()

            # ── Typové informace ─────────────────────────────────────────────
            _s1, _s2 = st.columns(2)
            with _s1:
                if _type_taxon and _type_taxon != NOT_PROVIDED:
                    st.markdown(tt("**🎯 Type taxon**", "**🎯 Type taxon**"))
                    st.info(_type_taxon)
                _type_spec = _sys_fields.get("TYPE SPECIMENS", "") or _sys_fields.get("TYPE MATERIAL", "")
                if _type_spec and _type_spec != NOT_PROVIDED:
                    st.markdown("**🔬 Type specimens**")
                    st.text_area("ts", value=_type_spec, height=90,
                                 disabled=True, label_visibility="collapsed",
                                 key=f"dos_ts_{selected_name}")
            with _s2:
                if _incl_taxa and _incl_taxa != NOT_PROVIDED:
                    st.markdown(tt("**📋 Included taxa / Type species**", "**📋 Included taxa / Type species**"))
                    st.text_area("it", value=_incl_taxa, height=90,
                                 disabled=True, label_visibility="collapsed",
                                 key=f"dos_it_{selected_name}")
                _locality = _sys_fields.get("LOCALITY", "")
                _strat    = _sys_fields.get("STRATIGRAPHY", "")
                if _strat and _strat != NOT_PROVIDED:
                    st.markdown(tt("**⏳ Stratigraphy**", "**⏳ Stratigraphy**"))
                    st.caption(_strat[:300])
                if _locality and _locality != NOT_PROVIDED:
                    st.markdown(tt("**📍 Locality**", "**📍 Locality**"))
                    st.caption(_locality[:200])

            # ── Synonymika ────────────────────────────────────────────────────
            if _synonymy and _synonymy != NOT_PROVIDED:
                st.divider()
                with st.expander(tt("📜 Synonymika", "📜 Synonymy"), expanded=False):
                    st.text(_synonymy)

            # ── Příbuzné taxony — Context-dependentní dotaz ───────────────────
            st.divider()
            _rb1, _rb2 = st.columns([3, 1])
            _rb1.markdown("**🔗 Relatives taxa in DB**")
            if _rb2.button("🔄 Recalculate relationships", key="dos_rebuild_hier",
                           help="Prepocita parent-child vztahy allch taxa v DB. Spustte po pridani novych documents."):
                with st.spinner(tt("Recalculating hierarchy…", "Recomputing hierarchy…")):
                    _rh = rebuild_taxon_hierarchy()
                st.success(
                    f"Hotovo: {_rh['updated']} updated, "
                    f"{_rh['already_ok']} unchanged, "
                    f"{_rh['skipped']} preskoceno (unknown rank).")
                st.rerun()

            _kin_rank_l = _occ_rank.lower()
            _taxon_first_word = (selected_name or "").split()[0]

            if _kin_rank_l == "species":
                # Pro druh: sourozenci = ostatní druhy se stejným rodem
                # (tj. stejné první slovo jména, or stejný parent_taxon_name)
                _kin_genus = (
                    parent_name if parent_name
                    else _taxon_first_word)
                _con_kin = db()
                _kin_rows = _con_kin.execute(
                    "SELECT DISTINCT taxon_name FROM taxon_candidates "
                    "WHERE rank_guess='Species' "
                    "AND status IN ('approved','pending','needs_review') "
                    "AND taxon_name != ? "
                    "AND (parent_taxon_name=? OR taxon_name LIKE ?) "
                    "ORDER BY taxon_name LIMIT 20",
                    (selected_name, _kin_genus, f"{_kin_genus} %")).fetchall()
                _con_kin.close()
                if _kin_rows:
                    st.caption(f"Druhy rodu *{_kin_genus}*:")
                    _kin_cols = st.columns(min(len(_kin_rows), 3))
                    for _ki, _kr in enumerate(_kin_rows[:12]):
                        with _kin_cols[_ki % 3]:
                            if st.button(
                                f"*{_kr['taxon_name']}*",
                                key=f"kin_sp_{_ki}_{selected_name}",
                                use_container_width=True):
                                st.session_state["td_selected"] = _kr["taxon_name"]
                                st.rerun()
                else:
                    st.caption(tt(f"No siblings for genus *{_kin_genus}* v DB.", f"No siblings for genus *{_kin_genus}* in DB."))

            elif _kin_rank_l in ("genus", "subgenus"):
                # Pro rod: sourozenci = ostatní rody stejné čeledi
                # Čeleď zjistíme z parent_taxon_name or z INCLUDED TAXONS nadřazené čeledi
                _kin_family = parent_name or ""
                if not _kin_family:
                    # Zkus najít čeleď, jejíž INCLUDED TAXONS obsahuje tento rod
                    _con_kin2 = db()
                    _fam_r = _con_kin2.execute(
                        "SELECT tc.taxon_name FROM taxon_candidates tc "
                        "JOIN occurrence_fields of ON tc.id=of.candidate_id "
                        "WHERE tc.rank_guess IN ('Family','Subfamily') "
                        "AND tc.status IN ('approved','pending','needs_review') "
                        "AND of.field_name='INCLUDED TAXONS' "
                        "AND of.field_value LIKE ? LIMIT 1",
                        (f"%{_taxon_first_word}%",)).fetchone()
                    _con_kin2.close()
                    if _fam_r:
                        _kin_family = _fam_r["taxon_name"]

                _con_kin = db()
                if _kin_family:
                    _kin_rows = _con_kin.execute(
                        "SELECT DISTINCT taxon_name FROM taxon_candidates "
                        "WHERE rank_guess IN ('Genus','Subgenus') "
                        "AND status IN ('approved','pending','needs_review') "
                        "AND taxon_name != ? AND parent_taxon_name=? "
                        "ORDER BY taxon_name LIMIT 20",
                        (selected_name, _kin_family)).fetchall()
                    _kin_label = tt(f"Rody celedi *{_kin_family}*", f"Genera of family *{_kin_family}*")
                else:
                    # Fallback: allchny rody ze stejného Documentu
                    _doc_ids = [o["document_id"] for o in occs]
                    _ph = ",".join("?"*len(_doc_ids))
                    _kin_rows = _con_kin.execute(
                        f"SELECT DISTINCT taxon_name FROM taxon_candidates "
                        f"WHERE rank_guess IN ('Genus','Subgenus') "
                        f"AND status IN ('approved','pending','needs_review') "
                        f"AND taxon_name != ? AND document_id IN ({_ph}) "
                        f"ORDER BY taxon_name LIMIT 12",
                        [selected_name] + _doc_ids).fetchall()
                    _kin_label = tt("Rody z techto documents", "Genera from these documents")
                _con_kin.close()
                if _kin_rows:
                    st.caption(f"{_kin_label}:")
                    _kin_cols = st.columns(min(len(_kin_rows), 3))
                    for _ki, _kr in enumerate(_kin_rows[:12]):
                        with _kin_cols[_ki % 3]:
                            if st.button(
                                _kr["taxon_name"],
                                key=f"kin_gn_{_ki}_{selected_name}",
                                use_container_width=True):
                                st.session_state["td_selected"] = _kr["taxon_name"]
                                st.rerun()
                else:
                    st.caption(tt("No related genera in DB (try Recalculate relationships).", "No related genera in DB (run Recompute relationships)."))

            elif _kin_rank_l in ("family", "subfamily", "superfamily"):
                # Pro čeleď: zobrazit zařazené rody (z INCLUDED TAXONS or z parent)
                _incl_source = _incl_taxa or ""
                _con_kin = db()
                _child_rows = _con_kin.execute(
                    "SELECT DISTINCT taxon_name FROM taxon_candidates "
                    "WHERE rank_guess IN ('Genus','Subgenus') "
                    "AND status IN ('approved','pending','needs_review') "
                    "AND parent_taxon_name=? "
                    "ORDER BY taxon_name LIMIT 30",
                    (selected_name,)).fetchall()
                _con_kin.close()
                _child_display = [r["taxon_name"] for r in _child_rows]
                # Doplnit z textového pole INCLUDED TAXONS
                if _incl_source and _incl_source != NOT_PROVIDED:
                    for _iw in re.split(r"[;,\n]", _incl_source):
                        _iw = _iw.strip().strip("*_")
                        if _iw and _iw not in _child_display:
                            _child_display.append(_iw)
                if _child_display:
                    st.caption(tt(f"Included genera ({len(_child_display)}):", f"Assigned genera ({len(_child_display)}):"))
                    _kin_cols = st.columns(min(len(_child_display), 3))
                    for _ki, _cn in enumerate(_child_display[:18]):
                        with _kin_cols[_ki % 3]:
                            if st.button(
                                _cn, key=f"kin_fm_{_ki}_{selected_name}",
                                use_container_width=True):
                                st.session_state["td_selected"] = _cn
                                st.rerun()
                else:
                    st.caption(
                        "Zadne rody prirazeny k teto celedi. Spustte Prepocitat vztahy or doplnte fields INCLUDED TAXONS.")

            else:
                # Vyšší ranky (Order, Class, Phylum): zobrazit podřazené čeledi
                _con_kin = db()
                _child_rows = _con_kin.execute(
                    "SELECT DISTINCT taxon_name, rank_guess FROM taxon_candidates "
                    "WHERE rank_guess IN ('Family','Subfamily','Order','Suborder') "
                    "AND status IN ('approved','pending','needs_review') "
                    "AND parent_taxon_name=? "
                    "ORDER BY rank_guess, taxon_name LIMIT 20",
                    (selected_name,)).fetchall()
                _con_kin.close()
                if _child_rows:
                    st.caption(tt(f"Subordinate taxa ({len(_child_rows)}):", f"Subordinate taxa ({len(_child_rows)}):"))
                    for _cr in _child_rows[:20]:
                        if st.button(
                            f"{_cr['taxon_name']} ({_cr['rank_guess']})",
                            key=f"kin_hi_{_cr['taxon_name']}_{selected_name}",
                            use_container_width=True):
                            st.session_state["td_selected"] = _cr["taxon_name"]
                            st.rerun()
                else:
                    st.caption(
                        "Zadne podrazene taxony. Spustte Prepocitat vztahy po pridani documents.")

        # ══════════════════════════════════════════════════════════════════════
        # 📜 RAW + PDF
        # ══════════════════════════════════════════════════════════════════════
        with tab_block:
            # Allchny výskyty chronologicky s raw blokem + inline PDF
            for occ in sorted(occs, key=lambda x: (x["filename"], x["page_start"])):
                block = occ["block_text"] or ""
                if not block:
                    block = extract_block_for_candidate(
                        occ["id"], occ["document_id"], occ["page_start"])

                # Načíst cestu k fileu
                _td_doc_key = f"docpath_{occ['document_id']}"
                if _td_doc_key not in st.session_state:
                    _tdcon = db()
                    _tdoc_r = _tdcon.execute(
                        "SELECT path FROM documents WHERE id=?",
                        (occ["document_id"],)).fetchone()
                    _tdcon.close()
                    st.session_state[_td_doc_key] = (_tdoc_r["path"] if _tdoc_r else None)
                _tddoc_path = st.session_state[_td_doc_key]
                _occ_d = dict(occ)
                _raw_pg_s = int(_occ_d.get("page_start") or 1)
                _raw_pg_e = int(_occ_d.get("block_end_page") or _raw_pg_s)
                _raw_is_pdf = bool(
                    _tddoc_path
                    and pathlib.Path(_tddoc_path).suffix.lower() == ".pdf"
                    and pathlib.Path(_tddoc_path).exists()
                    and HAS_FITZ)

                # Záhlaví výskytu
                conf_color = ("22c55e" if occ["confidence"] >= 0.80
                              else "f59e0b" if occ["confidence"] >= 0.60 else "ef4444")
                _pg_range_lbl = (f"str. {_raw_pg_s}–{_raw_pg_e}"
                                 if _raw_pg_e > _raw_pg_s else f"str. {_raw_pg_s}")
                st.markdown(
                    f"<div style='font-size:0.79rem;margin:6px 0 2px;color:#334155'>"
                    f"<span style='background:#{conf_color};color:#fff;"
                    f"padding:1px 5px;border-radius:3px;font-weight:700;"
                    f"font-size:0.74rem'>{occ['confidence']:.3f}</span>&nbsp;"
                    f"<b>{occ['filename']}</b>&nbsp;"
                    f"<span style='color:#64748b'>· {_pg_range_lbl}</span>&nbsp;"
                    f"{_badge_html(occ['status'])}</div>",
                    unsafe_allow_html=True)

                # ── PDF toggle: zobrazí stránky inline ────────────────────────
                if _raw_is_pdf:
                    _raw_pdf_tog = st.toggle(
                        f"📄 View PDF ({_pg_range_lbl})",
                        value=False,
                        key=f"raw_pdf_{occ['id']}",
                        help="Shows original page PDF in app.")
                    if _raw_pdf_tog:
                        _n_raw_pgs = _raw_pg_e - _raw_pg_s + 1
                        if _n_raw_pgs <= 1:
                            _show_pdf_page_inline(
                                occ["document_id"], _raw_pg_s,
                                key=f"rpdf_{occ['id']}")
                        else:
                            _raw_pcols = st.columns(min(_n_raw_pgs, 3))
                            for _rpi, _rpg in enumerate(range(
                                    _raw_pg_s, min(_raw_pg_e + 1, _raw_pg_s + 6))):
                                with _raw_pcols[_rpi % 3]:
                                    st.caption(f"str. {_rpg}")
                                    _show_pdf_page_inline(
                                        occ["document_id"], _rpg,
                                        key=f"rpdf_{occ['id']}_{_rpg}")

                # ── Raw text ──────────────────────────────────────────────────
                display_text = annotate_raw_block(block) if annotate_all else block
                st.text_area(
                    "##blk", value=display_text or "–", height=200,
                    key=f"td_blk_{occ['id']}",
                    label_visibility="collapsed", disabled=True)

                # ── Akce na výskytu: auto-map / přeložit / editor ────────
                _ac_s = get_settings()
                _ac_llm = _ac_s.get("llm_enabled", False)
                # Zjistit jazyk Documentu pro tento výskyt
                _ac_doc_key = f"lang_{occ['document_id']}"
                if _ac_doc_key not in st.session_state:
                    _ac_lr = db().execute(
                        "SELECT lang FROM documents WHERE id=?",
                        (occ["document_id"],)).fetchone()
                    st.session_state[_ac_doc_key] = (
                        (_ac_lr["lang"] if _ac_lr else "") or "")
                _ac_lang = st.session_state[_ac_doc_key]
                _ac_lang_lbl = _ac_lang.upper() if _ac_lang else "?"

                ac1, ac2, ac3 = st.columns(3)
                if ac1.button("🔄 Auto-map", key=f"td_am2_{occ['id']}",
                              help="Extrahovat pole z bloku"):
                    n = auto_map_candidate_fields(occ["id"], block)
                    st.success(tt(f"Mapped {n} fields.", f"Mapped {n} fields.")); st.rerun()

                # ── Výběr jazyka pokud není znám ────────────────────────────
                _lang_picker_key = f"td_lang_pick_{occ['id']}"
                _lang_confirmed_key = f"td_lang_ok_{occ['id']}"
                if not _ac_lang or _ac_lang.lower() in {"?", ""}:
                    _lang_opts = {
                        "cs": "🇨🇿 Czech",
                        "ru": "🇷🇺 Russian",
                        "de": "🇩🇪 German",
                        "fr": "🇫🇷 French",
                        "zh": "🇨🇳 Chinese",
                        "la": "latina",
                        "pl": "Polish",
                        "sv": "Swedish",
                        "sk": "Slovak",
                        "it": "Italian",
                        "es": "Spanish",
                        "hu": "Hungarian",
                    }
                    _picked = ac2.selectbox(
                        "Language Documentu",
                        options=list(_lang_opts.keys()),
                        format_func=lambda k: _lang_opts[k],
                        key=_lang_picker_key,
                        label_visibility="collapsed",
                    )
                    if ac2.button("✅", key=_lang_confirmed_key,
                                  help="Confir language and translate"):
                        # Save jazyk do DB + session_state
                        _con_l = db()
                        _con_l.execute(
                            "UPDATE documents SET lang=? WHERE id=?",
                            (_picked, occ["document_id"]))
                        _con_l.commit(); _con_l.close()
                        st.session_state[_ac_doc_key] = _picked
                        _ac_lang = _picked
                        st.rerun()
                    _tr_do = False   # tlačítko překladu zobrazíme až po potvrzení
                else:
                    _tr_do = True

                _tr_lbl = (f"🌐 →EN ({_ac_lang_lbl})"
                           if _ac_lang and _ac_lang.lower() not in
                              {"en","eng","english"}
                           else "🌐 →EN")
                if _tr_do and ac2.button(
                        _tr_lbl,
                        key=f"td_tr_occ_{occ['id']}",
                        disabled=not _ac_llm,
                        help=("Translates the fields of this occurrence to English using LM Studio."
                              if _ac_llm else
                              "Requires LM Studio — enable it in Settings.")):
                    try:
                        with st.spinner(tt("Translating…", "Translating…")):
                            _n_occ = auto_translate_candidate_fields(
                                occ["id"], _ac_lang, _ac_s,
                                force=True, force_lang=_ac_lang,
                                block=block)
                        if _n_occ:
                            st.success(tt(f"✅ Translated / mapped {_n_occ} fields.", f"✅ Translated / mapped {_n_occ} fields."))
                            st.rerun()
                        else:
                            st.info(tt("No fields to translate — try Auto-map first.", "No fields to translate — try Auto-map first."))
                    except RuntimeError as _e:
                        st.error(tt(f"❌ {_e}", f"❌ {_e}"))
                    except Exception as _e:
                        st.error(tt(f"❌ Unexpected error: {_e}", f"❌ Unexpected error: {_e}"))

                if ac3.button("✏️ Editor", key=f"td_ed_{occ['id']}"):
                    st.session_state["editor_candidate_id"] = occ["id"]
                    st.session_state["active_tab"] = "editor"
                    st.rerun()
                st.markdown("<hr class='pn-occ-sep'>", unsafe_allow_html=True)

    # ── Export TXT ────────────────────────────────────────────────────────────
    if st.button(tt("📥 Export TXT pro filtr", "📥 Export TXT for filter"), key="td_export"):
        lines = [f"PaleoN Taxon Dossier – {query or '(all)'}",
                 f"{len(filtered)} occurrences, {len(sorted_names)} taxa", "="*70]
        for name in sorted_names:
            lines.append(f"\n{'='*70}\n{name} ({len(groups[name])}×)\n{'='*70}")
            for occ in groups[name]:
                lines.append(f"\n--- {occ['filename']}, str.{occ['page_start']} ---")
                flds = get_candidate_fields(occ["id"])
                for fn in FIELD_ORDER:
                    if fn in flds and flds[fn] and flds[fn] != NOT_PROVIDED:
                        lines.append(f"{fn}: {flds[fn]}")
                block = occ["block_text"] or ""
                lines.append(f"\nRAW:\n{block}")
        st.download_button(
            "📥 Download TXT",
            data="\n".join(lines).encode("utf-8"),
            file_name=f"taxon_dossier_{re.sub(r'[^a-zA-Z0-9]+','_',query or 'all')}.txt",
            mime="text/plain", key="td_dl")



def tab_morphostrat_dossier():
    # ── CSS ───────────────────────────────────────────────────────────────────
    st.markdown("""
    <style>
    /* Morphostrat-specific — shared pn-fl/pn-fv/pn-meta injected by tab_taxon_dossier */
    .pn-term-table{width:100%;border-collapse:collapse;font-size:0.83rem}
    .pn-term-table th{text-align:left;padding:3px 8px;border-bottom:2px solid #64748b;
        font-size:0.74rem;color:#374151;text-transform:uppercase;letter-spacing:.04em}
    .pn-term-table td{padding:3px 8px;border-bottom:1px solid rgba(100,116,139,0.15);
        vertical-align:top;color:#1e293b}
    .pn-cat-chip{display:inline-block;padding:2px 8px;border-radius:10px;
        font-size:0.75rem;cursor:pointer;margin:2px 2px;
        border:1px solid rgba(100,116,139,0.3)}
    .pn-result-row{padding:4px 0;border-bottom:1px solid rgba(71,85,105,0.15)}
    .pn-result-name{font-size:0.90rem;font-weight:700}
    .pn-result-meta{font-size:0.75rem;color:#374151}
    @media (prefers-color-scheme:light){
        .pn-term-table td,.pn-result-name{color:#0f172a}
        .pn-result-meta,.pn-term-table th{color:#334155}
    }
    div[data-testid="stRadio"] label p{font-size:0.82rem!important}
    </style>
    """, unsafe_allow_html=True)

    s = get_settings()
    con = db()
    n_morph = con.execute(
        "SELECT COUNT(DISTINCT candidate_id) FROM term_matches WHERE term_type='morphology'"
    ).fetchone()[0]
    n_strat = con.execute(
        "SELECT COUNT(DISTINCT candidate_id) FROM term_matches WHERE term_type='stratigraphy'"
    ).fetchone()[0]
    n_approved = con.execute(
        "SELECT COUNT(*) FROM taxon_candidates WHERE status='approved'"
    ).fetchone()[0]
    con.close()

    # ── Header lišta ─────────────────────────────────────────────────────────
    hh1, hh2, hh3, hh4 = st.columns([2, 2, 2, 1])
    hh1.metric("Approved records", n_approved)
    hh2.metric("With morphological terms",   n_morph)
    hh3.metric("With stratigraphic terms",n_strat)

    # ── Přecount termínu — přímý callback progress ───────────────────────────
    _just_recomputed = False

    if hh4.button("🔄 Recompute", key="ms_recompute",
                  help="Recompute morphological and stratigraphic terms"):
        # Vyčistit starý chunk-state, aby starý globální loop nepřebil UI.
        st.session_state.pop("terms_recompute_ids", None)
        st.session_state.pop("terms_recompute_done", None)

        status_text = st.empty()
        progress_bar = st.progress(0.0)

        def progress_callback(current: int, total: int):
            # st.progress bere float 0.0–1.0; text držíme odděleně, protože je
            # spolehlivější napříč verzemi Streamlitu než kombinovaný text parametr.
            total = max(int(total or 0), 1)
            current = max(0, min(int(current or 0), total))
            fraction = current / total
            progress_bar.progress(float(fraction))
            status_text.text(
                f"Recalculating morphology and stratigraphy: {current} / {total} "
                f"({int(fraction * 100)} %) ..."
            )

        try:
            status_text.text("Preparing batch recalculation of morphology and stratigraphy…")
            processed_count = recompute_all_term_matches(progress_cb=progress_callback)
            progress_bar.empty()
            status_text.success(
                f"✅ Recalculation complete! Records processed: {processed_count}."
            )
            _just_recomputed = True
        except Exception as e:
            progress_bar.empty()
            status_text.error(f"❌ An error occurred during recalculation: {e}")
            raise

    if n_morph == 0 and n_strat == 0 and not _just_recomputed:
        st.info(t("no_terms_computed2"))
        return

    # ── Typ + hledání v termínech ─────────────────────────────────────────────
    rc1, rc2, rc3, rc4 = st.columns([2, 4, 2, 1])
    ttype = "stratigraphy" if rc1.radio(
        "##t", [t("stratigraphy"), t("morphology")],
        horizontal=True, key="ms_type",
        label_visibility="collapsed").startswith(t("stratigraphy")[:2]) \
        else "morphology"

    term_search = rc2.text_input(
        "🔍 Hledat pojem", key="ms_term_search",
        placeholder="filter terms alphabetically…",
        label_visibility="collapsed")

    sort_alpha = rc3.toggle("A–Z sorting", value=True, key="ms_sort_alpha",
                            help="Alphabetically (otherwise by number of matches ↓)")
    show_ctx  = rc4.toggle("Context", value=False, key="ms_ctx")

    # ── Načíst kategorie a termíny ────────────────────────────────────────────
    con = db()
    cat_counts = con.execute(
        "SELECT category, COUNT(DISTINCT candidate_id) n FROM term_matches "
        "WHERE term_type=? GROUP BY category ORDER BY n DESC",
        (ttype,)).fetchall()
    con.close()

    if not cat_counts:
        st.warning(tt(f"Pro '{ttype}' — no matches. First approve records and recalculate.", f"No matches for '{ttype}' — approve records and recompute first."))
        return

    # ── Kategorie jako horizontální filtrovací pill ───────────────────────────
    all_cats = ["— All —"] + [r["category"] for r in cat_counts if r["category"]]
    sel_cat_label = st.radio(
        "Kategorie", all_cats,
        index=min(st.session_state.get("ms_cat_idx", 0), len(all_cats)-1),
        horizontal=True, key="ms_cat_radio", label_visibility="collapsed")
    st.session_state["ms_cat_idx"] = all_cats.index(sel_cat_label)
    sel_cat = None if sel_cat_label == "— All —" else sel_cat_label

    # ── Načíst termíny pro tuto kategorii ─────────────────────────────────────
    con = db()
    if sel_cat:
        term_rows = con.execute(
            "SELECT canonical, COUNT(DISTINCT candidate_id) n FROM term_matches "
            "WHERE term_type=? AND category=? GROUP BY canonical",
            (ttype, sel_cat)).fetchall()
    else:
        term_rows = con.execute(
            "SELECT canonical, COUNT(DISTINCT candidate_id) n FROM term_matches "
            "WHERE term_type=? GROUP BY canonical",
            (ttype,)).fetchall()
    con.close()

    # Filtrovat podle hledání
    term_needle = _norm_search(term_search.strip()) if term_search.strip() else ""
    if term_needle:
        term_rows = [r for r in term_rows if term_needle in _norm_search(r["canonical"])]

    # Seřadit
    if sort_alpha:
        term_rows = sorted(term_rows, key=lambda r: r["canonical"].lower())
    else:
        term_rows = sorted(term_rows, key=lambda r: -r["n"])

    if not term_rows:
        st.info(t("no_term_matches"))
        return

    # ── Dvousloupcový layout: seznam termínu | výsledky ───────────────────────
    left_col, right_col = st.columns([1, 3], gap="small")

    with left_col:
        st.markdown(
            f"<div style='font-size:0.70rem;color:#64748b;margin-bottom:4px'>"
            f"<b>{len(term_rows)}</b> terms"
            + (" (A–Z)" if sort_alpha else " (↓ frequency)")
            + "</div>", unsafe_allow_html=True)

        # Klikací seznam termínu (st.radio = přímý výběr kliknutím)
        term_options = [r["canonical"] for r in term_rows[:300]]
        if not term_options:
            st.stop()

        if "ms_sel_term" not in st.session_state or \
           st.session_state["ms_sel_term"] not in term_options:
            st.session_state["ms_sel_term"] = term_options[0]

        _ms_df = pd.DataFrame([{
            "Term": r["canonical"],
            "n": r["n"],
        } for r in term_rows[:300]])

        _ms_cur_idx = 0
        if st.session_state["ms_sel_term"] in term_options:
            _ms_cur_idx = term_options.index(st.session_state["ms_sel_term"])

        _ms_sel = st.dataframe(
            _ms_df,
            column_config={
                "Term": st.column_config.TextColumn("Term", width="large"),
                "n":      st.column_config.NumberColumn("n", width=50),
            },
            hide_index=True,
            use_container_width=True,
            height=min(36 * len(_ms_df) + 38, 520),
            on_select="rerun",
            selection_mode="single-row",
            key="ms_term_table",
        )
        _ms_rows = getattr(_ms_sel, "selection", {})
        if isinstance(_ms_rows, dict):
            _ms_rows = _ms_rows.get("rows", [])
        elif hasattr(_ms_rows, "rows"):
            _ms_rows = list(_ms_rows.rows)
        else:
            _ms_rows = []
        if _ms_rows and _ms_rows[0] < len(term_options):
            sel_canonical = term_options[_ms_rows[0]]
            st.session_state["ms_sel_term"] = sel_canonical
        else:
            # Guard: stored selection may be out of range after filter change
            if st.session_state.get("ms_sel_term") not in term_options:
                st.session_state["ms_sel_term"] = term_options[0] if term_options else ""
            sel_canonical = st.session_state["ms_sel_term"]
        if len(term_rows) > 300:
            st.caption(tt(f"… a {len(term_rows)-300} more", f"… and {len(term_rows)-300} more"))

    with right_col:
        # ── Výsledky pro vybraný termín ───────────────────────────────────────
        con = db()
        matches = con.execute(
            """SELECT tm.term, tm.category, tm.source_field,
                      tc.id cid, tc.taxon_name, tc.page_start,
                      tc.document_id, tc.block_text, tc.status,
                      tc.confidence, tc.context_before, tc.context_after,
                      tc.rank_guess, tc.debug_json, d.filename
               FROM term_matches tm
               JOIN taxon_candidates tc ON tm.candidate_id=tc.id
               JOIN documents d ON tc.document_id=d.id
               WHERE tm.term_type=? AND tm.canonical=?
               ORDER BY tc.taxon_name, d.filename, tc.page_start""",
            (ttype, sel_canonical)).fetchall()
        con.close()

        # Live filtr uvnitř výsledku
        res_search = st.text_input(
            "🔍 Filtr v taxonech", key="ms_res_filter",
            placeholder="Search in taxon names…",
            label_visibility="collapsed")
        if res_search.strip():
            rn = _norm_search(res_search.strip())
            matches = [m for m in matches if rn in _norm_search(m["taxon_name"] or "")]

        st.markdown(
            f"<div style='font-size:0.75rem;margin:2px 0 6px;"
            f"color:#94a3b8'>Term <b style='color:#e2e8f0'>"
            f"\"{sel_canonical}\"</b> — "
            f"<b style='color:#e2e8f0'>{len(matches)}</b> taxa</div>",
            unsafe_allow_html=True)

        if not matches:
            st.info(t("no_match_for_term"))
        else:
            # ── Klikací tabulka výsledku (st.dataframe s výběrem řádku) ──────
            _ms_match_df = pd.DataFrame([{
                "Taxon":    m["taxon_name"],
                "Rank":     m["rank_guess"] or "—",
                "Dok.":     pathlib.Path(m["filename"]).stem[:30],
                "Str.":     m["page_start"],
                "Conf.":    round(m["confidence"], 3),
                "Pole":     m["source_field"],
                "_cid":     m["cid"],
            } for m in matches[:200]])

            _ms_tax_sel = st.dataframe(
                _ms_match_df.drop(columns=["_cid"]),
                column_config={
                    "Taxon": st.column_config.TextColumn("Taxon", width="large"),
                    "Rank":  st.column_config.TextColumn("Rank",  width=80),
                    "Dok.":  st.column_config.TextColumn("Dok.",  width=150),
                    "Str.":  st.column_config.NumberColumn("Str.", width=55),
                    "Conf.": st.column_config.NumberColumn("Conf.", width=65, format="%.3f"),
                    "Pole":  st.column_config.TextColumn("Pole",  width=120),
                },
                hide_index=True,
                use_container_width=True,
                height=min(36 * len(_ms_match_df) + 38, 380),
                on_select="rerun",
                selection_mode="single-row",
                key="ms_tax_table",
            )
            if len(matches) > 200:
                st.caption(tt(f"… a {len(matches)-200} more", f"… and {len(matches)-200} more"))

            _ms_tax_rows = getattr(_ms_tax_sel, "selection", {})
            if isinstance(_ms_tax_rows, dict):
                _ms_tax_rows = _ms_tax_rows.get("rows", [])
            elif hasattr(_ms_tax_rows, "rows"):
                _ms_tax_rows = list(_ms_tax_rows.rows)
            else:
                _ms_tax_rows = []

            if _ms_tax_rows:
                _ms_sel_row = _ms_match_df.iloc[_ms_tax_rows[0]]
                _ms_sel_cid = int(_ms_sel_row["_cid"])
                _ms_sel_name = _ms_sel_row["Taxon"]

                st.markdown("---")
                _ms_str = int(_ms_sel_row["Str."])
                _ms_dok = _ms_sel_row["Dok."]
                st.markdown(
                    f"<div style='font-size:0.80rem;margin:2px 0 4px'>"
                    f"Detail: <b>{_ms_sel_name}</b> "
                    f"<span style='color:#94a3b8'>str.{_ms_str}"
                    f" · {_ms_dok}</span></div>",
                    unsafe_allow_html=True)

                # Načíst pole a blok pro vybraného kandidáta
                _ms_fields = get_candidate_fields(_ms_sel_cid)
                _ms_cand = get_candidate(_ms_sel_cid)
                if _ms_fields:
                    _ms_fhtml = ""
                    for _fn in ["DIAGNOSIS","DESCRIPTION","STRATIGRAPHY",
                                "LOCALITY","OCCURRENCE","SIZE","MATERIAL EXAMINED"]:
                        _fv = _ms_fields.get(_fn,"")
                        if _fv and _fv != NOT_PROVIDED:
                            _safe = str(_fv).replace("<","&lt;").replace(">","&gt;")
                            _ms_fhtml += (
                                f"<div class='pn-field-grid'>"
                                f"<div class='pn-fl'>{_fn}</div>"
                                f"<div class='pn-fv'>{_safe}</div></div>")
                    if _ms_fhtml:
                        st.markdown(_ms_fhtml, unsafe_allow_html=True)

                # PDF odkaz
                if _ms_cand:
                    _doc_id = _ms_cand.get("document_id")
                    _pg     = _ms_cand.get("page_start", 0)
                    _pck = f"docpath_{_doc_id}"
                    if _pck not in st.session_state:
                        _pc = db(); _pr = _pc.execute(
                            "SELECT path FROM documents WHERE id=?",(_doc_id,)).fetchone()
                        _pc.close()
                        st.session_state[_pck] = (_pr["path"] if _pr else None)
                    _msp = st.session_state[_pck]
                    if _msp and pathlib.Path(_msp).exists() and _msp.lower().endswith(".pdf"):
                        _msu = pathlib.Path(_msp).resolve().as_uri() + f"#page={_pg}"
                        st.markdown(
                            f"<a href='{_msu}' target='_blank' "
                            f"style='color:#60a5fa;font-size:0.76rem'>"
                            f"📄 Open PDF page&nbsp;{_pg}</a>",
                            unsafe_allow_html=True)
                    # Toggle PDF náhled
                    if HAS_FITZ and _doc_id and st.toggle(
                            "📄 Preview of pages", value=False, key=f"ms_pdf_{_ms_sel_cid}"):
                        _show_pdf_page_inline(_doc_id, _pg, key=f"ms_pdfv_{_ms_sel_cid}")

            # Second detail panel removed — using the dataframe selection above only
            _show_ctx_detail = show_ctx and _ms_tax_rows and _ms_tax_rows[0] < len(matches)
            if _ms_tax_rows and _ms_tax_rows[0] < len(matches):
                m = matches[_ms_tax_rows[0]]
            elif matches:
                m = matches[0]
            if _show_ctx_detail:
                block = m["block_text"] or extract_block_for_candidate(
                    m["cid"], m["document_id"], m["page_start"])
                cand_fields = get_candidate_fields(m["cid"])

                dc1, dc2 = st.columns([3, 2], gap="medium")
                with dc1:
                    _con_lnk2 = db()
                    _doc_lnk2 = _con_lnk2.execute(
                        "SELECT path FROM documents WHERE id=?",
                        (m["document_id"],)).fetchone()
                    _con_lnk2.close()
                    _lnk2 = ""
                    if _doc_lnk2 and _doc_lnk2["path"]:
                        _fp2 = pathlib.Path(_doc_lnk2["path"])
                        if _fp2.exists():
                            _fu2 = _fp2.resolve().as_uri()
                            if _fp2.suffix.lower() == ".pdf":
                                _fu2 += f"#page={m['page_start']}"
                            _lnk2 = (
                                f" &nbsp;<a href='{_fu2}' target='_blank'"
                                f" style='color:#3b82f6;font-size:0.75rem'>"
                                f"[str.{m['page_start']}]</a>")
                    st.markdown(
                        f"<p style='font-size:0.90rem;font-weight:700;margin:0'>"
                        f"{m['taxon_name']}</p>"
                        f"<p style='font-size:0.77rem;color:#334155;margin:2px 0 6px'>"
                        f"{m['filename']}, str.{m['page_start']}{_lnk2} · "
                        f"conf.{m['confidence']:.3f}</p>",
                        unsafe_allow_html=True)

                    if cand_fields:
                        DISP_FIELDS = ["DIAGNOSIS","DESCRIPTION","LOCALITY",
                                       "STRATIGRAPHY","TYPE SPECIMENS","OCCURRENCE",
                                       "SIZE","REMARKS","AUTHOR","SYNONYMY"]
                        rows_html = ""
                        for fn in DISP_FIELDS:
                            fv = cand_fields.get(fn,"")
                            if fv and fv != NOT_PROVIDED:
                                safe_fv = fv.replace("<","&lt;").replace(">","&gt;")
                                rows_html += (
                                    f"<div style='display:grid;"
                                    f"grid-template-columns:130px 1fr;"
                                    f"gap:2px 8px;margin:1px 0;"
                                    f"border-bottom:1px solid rgba(71,85,105,.1)'>"
                                    f"<div style='font-size:0.64rem;font-weight:700;"
                                    f"color:#64748b;text-transform:uppercase;"
                                    f"letter-spacing:.04em;padding:2px 0'>{fn}</div>"
                                    f"<div style='font-size:0.85rem;line-height:1.35;"
                                    f"padding:2px 0'>{safe_fv}</div></div>")
                        if rows_html:
                            st.markdown(rows_html, unsafe_allow_html=True)
                        else:
                            st.caption(tt("Field are no filled.", "No fields filled."))

                with dc2:
                    _ms_pdf = st.toggle(
                        "📄 PDF Preview",
                        value=False,
                        key=f"ms_pdf_{m['cid']}_{sel_canonical}",
                        help="Show original page PDF.")
                    if _ms_pdf and HAS_FITZ:
                        _show_pdf_page_inline(
                            m["document_id"], m["page_start"],
                            key=f"ms_pdfv_{m['cid']}")
                    st.caption("Raw blok:")
                    st.text_area(
                        "##ms_det_blk", value=block or "–", height=200,
                        key=f"ms_db_{m['cid']}_{sel_canonical}",
                        label_visibility="collapsed", disabled=True)
                    if show_ctx:
                        ctx_b = m["context_before"] or ""
                        ctx_a = m["context_after"] or ""
                        if ctx_b:
                            st.caption(t("context_before_caption"))
                            st.text(ctx_b[:200])
                        if ctx_a:
                            st.caption("Context po:")
                            st.text(ctx_a[:200])



# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 – NASTAVENÍ
# ══════════════════════════════════════════════════════════════════════════════

def tab_settings():
    st.header(t("settings_header"))
    s = get_settings()

    stab1, stab1b, stab1c, stab2, stab3, stab4, stab5, stab_users, stab_dup = st.tabs(
        [t("schema_tab"), t("gazetteer_tab"), t("terms_tab"), t("llm_tab"),
         t("data_tab"), t("quality_tab"), "🏆 Gold Set", "👥 Users",
         "🔀 Duplicates"])

    # ── Schéma ────────────────────────────────────────────────────────────────
    with stab1:
        schema = load_schema()

        # ── Metriky ──────────────────────────────────────────────────────────
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.metric("Number of labels", len(schema))
        sc2.metric("Fields (target_field)", schema["target_field"].nunique())
        sc3.metric("Languages", schema["language"].nunique())
        sc4.metric("Active", (schema["enabled"] == "1").sum())

        schema_src = ("✅ File: `" + str(SCHEMA_FILE) + "`"
                      if SCHEMA_FILE.exists() else "⚠️ Built-in fallback")
        st.caption(schema_src)

        # ── Upload + Download ─────────────────────────────────────────────────
        up_col, dl_col = st.columns(2)
        with up_col:
            new_schema = st.file_uploader(
                "📂 Upload new schema TSV", type=["tsv"], key="schema_up",
                help="TSV se sloupci: enabled, label, label_regex, canonical_label, "
                     "target_field, is_strong, can_start_treatment, "
                     "type_species_standalone, priority, language")
            if new_schema and st.button(t("schema_tab") + " " + t("save_caption")[3:].strip(),
                                         key="schema_upload_btn", type="primary"):
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                SCHEMA_FILE.write_bytes(new_schema.getbuffer())
                reload_schema()
                st.success(tt(f"✅ Schema saved ({new_schema.size:,} B).", f"✅ Schema saved ({new_schema.size:,} B)."))
                st.rerun()
        with dl_col:
            if SCHEMA_FILE.exists():
                st.download_button(
                    "📥 Download aktualni schema (TSV)",
                    data=SCHEMA_FILE.read_bytes(),
                    file_name="section_schema.tsv",
                    mime="text/tab-separated-values",
                    key="schema_dl",
                    help="Save actuall TSV file.")
            else:
                st.download_button(
                    "📥 Download vestavene schema (TSV)",
                    data=_MINIMAL_SCHEMA_TSV.encode("utf-8"),
                    file_name="section_schema_minimal.tsv",
                    mime="text/tab-separated-values",
                    key="schema_dl_fallback")

        # ── Distribuce ───────────────────────────────────────────────────────
        with st.expander(t("label_dist_expander"), expanded=False):
            ch1, ch2 = st.columns(2)
            with ch1:
                st.caption(t("labels_by_field"))
                dist = (schema.groupby("target_field").size()
                        .reset_index(name="count")
                        .sort_values("count", ascending=True))
                st.bar_chart(dist.set_index("target_field")["count"])
            with ch2:
                st.caption(t("labels_by_lang"))
                ldist = (schema.groupby("language").size()
                         .reset_index(name="count")
                         .sort_values("count", ascending=False))
                st.bar_chart(ldist.set_index("language")["count"])

        st.divider()
        st.subheader(t("label_editor_sub"))
        st.caption(
            "Filter labels (fields, language, status), edit the values directly in the table, and save. Add a new label using the **＋ Add Label** button below. The *label_regex* and *canonical_label* columns will be filled in automatically upon saving if you leave them blank.")

        all_fields = sorted(schema["target_field"].dropna().unique().tolist())
        all_langs  = sorted(schema["language"].dropna().unique().tolist())

        # ── Filtry ───────────────────────────────────────────────────────────
        ef1, ef2, ef3, ef4 = st.columns([3, 2, 1, 2])
        sel_field   = ef1.selectbox("Target field", ["— all —"] + all_fields,
                                     key="schema_edit_field")
        sel_lang    = ef2.selectbox("Language",       ["— all —"] + all_langs,
                                     key="schema_edit_lang")
        sel_enabled = ef3.selectbox("State", ["all", "1 – active", "0 – inactive"],
                                     key="schema_edit_enabled")
        label_srch  = ef4.text_input("🔍 Search in label", key="schema_srch",
                                      placeholder="diagnosis, 诊断, Диагноз…")

        view = schema.copy()
        if sel_field != "— all —":
            view = view[view["target_field"] == sel_field]
        if sel_lang != "— all —":
            view = view[view["language"] == sel_lang]
        if sel_enabled.startswith("1"):
            view = view[view["enabled"] == "1"]
        elif sel_enabled.startswith("0"):
            view = view[view["enabled"] == "0"]
        if label_srch:
            view = view[view["label"].str.contains(label_srch, case=False, na=False)]
        view = view.reset_index(drop=True)

        # Save original label pro merge při saveení
        view["_orig_label"] = view["label"]

        st.caption(tt(f"Zobrazeno **{len(view)}** z {len(schema)} labels", f"Showing **{len(view)}** of {len(schema)} labels")
                   + (f" — pole: `{sel_field}`" if sel_field != "— all —" else ""))

        # ── Hromadné akce na pole ─────────────────────────────────────────────
        if sel_field != "— all —":
            ba1, ba2, _ = st.columns([1, 1, 4])
            if ba1.button("✅ Activate all", key="schema_enable_all"):
                schema.loc[schema["target_field"] == sel_field, "enabled"] = "1"
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                schema.to_csv(SCHEMA_FILE, sep="\t", index=False)
                reload_schema()
                st.success(tt("Allchny labely tohoto fields aktivovany.", "All labels for this field activated.")); st.rerun()
            if ba2.button("⛔ Deactivate all", key="schema_disable_all"):
                schema.loc[schema["target_field"] == sel_field, "enabled"] = "0"
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                schema.to_csv(SCHEMA_FILE, sep="\t", index=False)
                reload_schema()
                st.success(tt("Allchny labely tohoto fields deaktivovany.", "All labels for this field deactivated.")); st.rerun()

        # ── Tabulka editoru ───────────────────────────────────────────────────
        EDIT_COLS_BASIC = ["enabled", "label", "target_field",
                           "is_strong", "can_start_treatment", "priority", "language"]
        EDIT_COLS_ADV   = ["label_regex", "canonical_label", "type_species_standalone"]

        show_adv = st.toggle(
            "Show advanced  columns (label_regex, canonical_label)",
            value=False, key="schema_show_adv")
        display_cols = EDIT_COLS_BASIC + (EDIT_COLS_ADV if show_adv else [])

        for col in EDIT_COLS_ADV + ["_orig_label"]:
            if col not in view.columns:
                view[col] = ""

        view_display = view[display_cols].copy()

        col_cfg_full: Dict = {
            "enabled": st.column_config.SelectboxColumn(
                "Active", options=["1", "0"], width=70),
            "label": st.column_config.TextColumn(
                "Label (text k detekci)", width=280),
            "target_field": st.column_config.SelectboxColumn(
                "Cilove field", options=all_fields, width=180),
            "is_strong": st.column_config.SelectboxColumn(
                "Strong", options=["1", "0"], width=70,
                help="1 = trusty context signal."),
            "can_start_treatment": st.column_config.SelectboxColumn(
                "Starts a record?", options=["1", "0"], width=110,
                help="1 = label muze byt prvnim radkem taxonomickeho record."),
            "priority": st.column_config.NumberColumn(
                "Priorita", width=80, min_value=0, max_value=200,
                help="Higher number = higher reliability (0–200)."),
            "language": st.column_config.TextColumn(
                "Language", width=90,
                help="en/mixed, cs, de, fr, ru, zh, la/abbr, dwc/db"),
            "label_regex": st.column_config.TextColumn(
                "label_regex", width=240,
                help="Regex varianta (prazdne = pouzije se label)."),
            "canonical_label": st.column_config.TextColumn(
                "canonical_label", width=240,
                help="Kanonicka forma (prazdne = pouzije se label)."),
            "type_species_standalone": st.column_config.SelectboxColumn(
                "Type sp. standalone", options=["0", "1"], width=130,
                help="1 = 'Type species' may be as separate record."),
        }

        editor_key = (f"schema_editor_{sel_field}_{sel_lang}_"
                      f"{sel_enabled}_{label_srch}_{show_adv}")
        edited_view = st.data_editor(
            view_display,
            column_config={k: v for k, v in col_cfg_full.items()
                           if k in display_cols},
            num_rows="dynamic",
            hide_index=True,
            use_container_width=True,
            key=editor_key,
        )

        # ── Uložení změn ─────────────────────────────────────────────────────
        sv1, sv2 = st.columns(2)
        if sv1.button("💾 Save displayed changes", type="primary",
                       key="schema_save_view"):
            BASE_DIR.mkdir(parents=True, exist_ok=True)

            # Odpagesit zobrazené řádky ze stávajícího schématu podle
            # jejich originálních labels (před editací)
            displayed_labels = set(view["_orig_label"].tolist())
            displayed_field  = sel_field if sel_field != "— all —" else None
            if displayed_field:
                mask_remove = ((schema["target_field"] == displayed_field) &
                               (schema["label"].isin(displayed_labels)))
            else:
                mask_remove = schema["label"].isin(displayed_labels)
            full = schema[~mask_remove].copy()

            # Zpracovat editované řádky
            new_rows = edited_view.copy()

            # label_regex a canonical_label: prázdné → doplnit z label
            for col_auto in ["label_regex", "canonical_label"]:
                if col_auto not in new_rows.columns:
                    new_rows[col_auto] = new_rows["label"]
                else:
                    mask_e = new_rows[col_auto].isna() | (new_rows[col_auto] == "")
                    new_rows.loc[mask_e, col_auto] = new_rows.loc[mask_e, "label"]

            if "type_species_standalone" not in new_rows.columns:
                new_rows["type_species_standalone"] = "0"
            else:
                new_rows["type_species_standalone"] = (
                    new_rows["type_species_standalone"].fillna("0"))

            TSV_COLS = ["enabled", "label", "label_regex", "canonical_label",
                        "target_field", "is_strong", "can_start_treatment",
                        "type_species_standalone", "priority", "language"]
            for col in TSV_COLS:
                if col not in full.columns:
                    full[col] = ""
                if col not in new_rows.columns:
                    new_rows[col] = ""

            full = pd.concat([full[TSV_COLS], new_rows[TSV_COLS]],
                             ignore_index=True, sort=False)
            full.to_csv(SCHEMA_FILE, sep="\t", index=False)
            reload_schema()
            st.success(tt(f"✅ Saved {len(new_rows)} labels ", f"✅ Saved {len(new_rows)} labels ") +
                       tt(f"(celkem {len(full)} in files).", f"({len(full)} total in file)."))
            st.rerun()

        if sv2.button("↩ Discard changes", key="schema_discard_view"):
            st.rerun()

        # ── Rychlé přidání nového labelu ──────────────────────────────────────
        st.divider()
        with st.expander(t("add_label_expander"), expanded=False):
            st.caption(
                "Vyplnte below a kliknete Pridat. Label se ihned savei do TSV.")
            na1, na2, na3 = st.columns([3, 2, 1])
            nb1, nb2, nb3 = st.columns([1, 1, 1])
            nc1, nc2      = st.columns([1, 1])

            new_label    = na1.text_input(
                "Label *", key="new_lbl_text",
                placeholder="Type locality, Locus typicus…")
            new_tfield   = na2.selectbox(
                "Target field *", all_fields, key="new_lbl_field")
            new_enabled  = na3.selectbox(
                "Active", ["1", "0"], key="new_lbl_enabled")
            new_strong   = nb1.selectbox(
                "Strong", ["1", "0"], key="new_lbl_strong")
            new_cst      = nb2.selectbox(
                "Starts a record?", ["0", "1"], key="new_lbl_cst")
            new_priority = nb3.number_input(
                "Priorita", 0, 200, 80, key="new_lbl_priority")
            new_lang     = nc1.text_input(
                "Language", value="en/mixed", key="new_lbl_lang")
            new_regex    = nc2.text_input(
                "label_regex (empty = label)", key="new_lbl_regex",
                placeholder="optional")

            if st.button(t("add_label_btn"), type="primary", key="new_lbl_add"):
                if not new_label.strip():
                    st.warning(t("label_text_required"))
                else:
                    BASE_DIR.mkdir(parents=True, exist_ok=True)
                    regex_val = new_regex.strip() or new_label.strip()
                    new_row = pd.DataFrame([{
                        "enabled":                new_enabled,
                        "label":                  new_label.strip(),
                        "label_regex":            regex_val,
                        "canonical_label":        new_label.strip(),
                        "target_field":           new_tfield,
                        "is_strong":              new_strong,
                        "can_start_treatment":    new_cst,
                        "type_species_standalone": "0",
                        "priority":               str(int(new_priority)),
                        "language":               new_lang.strip(),
                    }])
                    if SCHEMA_FILE.exists():
                        cur = pd.read_csv(
                            SCHEMA_FILE, sep="\t", dtype=str).fillna("")
                    else:
                        cur = pd.read_csv(
                            io.StringIO(_MINIMAL_SCHEMA_TSV),
                            sep="\t", dtype=str).fillna("")
                    updated = pd.concat(
                        [cur, new_row], ignore_index=True, sort=False)
                    updated.to_csv(SCHEMA_FILE, sep="\t", index=False)
                    reload_schema()
                    st.success(
                        f"✅ Label '{new_label.strip()}' -> '{new_tfield}' added.")
                    st.rerun()

    # ── Gazetteer (seznam známých taxa) ──────────────────────────────────────
    with stab1b:
        st.caption(
            "List of known taxa (taxons.txt) — used as a SOFT bonus in detection scoring, not as a hard filter. A candidate whose name (or genus) is on the list receives a confidence bonus; a missing match is not penalized in any way (newly described species in the current article are, of course, not included in the list).")

        genera, binomials = load_gazetteer()
        gz1, gz2 = st.columns(2)
        gz1.metric("Genera", len(genera))
        gz2.metric("Binomials", len(binomials))

        gaz_src = ("✅ File: `" + str(GAZETTEER_FILE) + "`"
                   if GAZETTEER_FILE.exists() else "⚠️ No gazetteer loaded")
        st.caption(gaz_src)

        s_gaz = get_settings()
        s_gaz["gazetteer_bonus"] = st.slider(
            "Bonus strength", 0.0, 0.30, float(s_gaz.get("gazetteer_bonus", 0.10)), 0.02,
            help="How many points are added to the score for a match with the gazetteer (binomial match = 1.5× this value)).")
        st.session_state["paleon_settings"] = s_gaz

        new_gaz = st.file_uploader(
            "Upload a new taxon list (TXT, one name per row)",
            type=["txt"], key="gazetteer_up")
        if new_gaz and st.button(t("gaz_save_btn")):
            BASE_DIR.mkdir(parents=True, exist_ok=True)
            GAZETTEER_FILE.write_bytes(new_gaz.getbuffer())
            reload_gazetteer()
            st.success(tt(f"Gazetteer saved ({new_gaz.size} B).", f"Gazetteer saved ({new_gaz.size} B)."))
            st.rerun()

        with st.expander(t("gaz_preview")):
            st.write(sorted(genera)[:60])

    # ── Morfologické a stratigrafické termíny ──────────────────────────────────
    with stab1c:
        st.caption(tt(
            'Termíny se používají v záložce „Morpho/Strat.“ pro filtrování taxonů'
            ' podle morfologických znaků a stratigrafických jednotek.'
            ' Shody jsou vypočítány automaticky při schválení záznamu — nebo ručně tlačítkem 🔄 Recompute níže.',
            'Terms are used in tab „Morpho/Strat.“ to filter taxa by morphological characters'
            ' and stratigraphic units.'
            ' Matches are computed automatically on approval — or manually with 🔄 Recompute below.'))

        # ── Resolve file paths once ────────────────────────────────────────────
        _morph_path = _upath("paleon_morphology_terms.tsv")
        if not _morph_path.exists():
            _morph_path = MORPHOLOGY_FILE
        _strat_path = _upath("stratigraphy_terms.tsv")
        if not _strat_path.exists():
            _strat_path = STRATIGRAPHY_FILE
        _syst_path = _upath("systematic_sections.tsv")
        if not _syst_path.exists():
            _syst_path = SYSTEMATIC_SECTIONS_FILE

        # ── Load raw files (all rows, including enabled=0) ─────────────────────
        _morph_raw = _load_raw_tsv(_morph_path, ["term", "canonical", "category", "language", "enabled"])
        _strat_raw = _load_raw_tsv(_strat_path, ["enabled", "term", "canonical", "category", "language", "match_mode"])
        _syst_raw  = _load_raw_tsv(_syst_path,  ["enabled", "language", "heading", "priority", "weight", "pattern", "note"])

        # ── Compact metrics strip ──────────────────────────────────────────────
        _mc1, _mc2, _mc3 = st.columns(3)
        with _mc1:
            st.caption(f"**{t('morpho_sub')}**")
            _mm1, _mm2, _mm3 = st.columns(3)
            _mm1.metric("Total", len(_morph_raw))
            _morph_active = int((_morph_raw["enabled"] == "1").sum()) if "enabled" in _morph_raw.columns else len(_morph_raw)
            _mm2.metric("Active", _morph_active)
            _mm3.metric("Cat.", _morph_raw["category"].nunique() if "category" in _morph_raw.columns else 0)
            st.caption(("✅ `" + _morph_path.name + "`") if _morph_path.exists() else "⚠️ No file")
        with _mc2:
            st.caption(f"**{t('strat_sub')}**")
            _sm1, _sm2, _sm3 = st.columns(3)
            _sm1.metric("Total", len(_strat_raw))
            _strat_active = int((_strat_raw["enabled"] == "1").sum()) if "enabled" in _strat_raw.columns else len(_strat_raw)
            _sm2.metric("Active", _strat_active)
            _sm3.metric("Cat.", _strat_raw["category"].nunique() if "category" in _strat_raw.columns else 0)
            st.caption(("✅ `" + _strat_path.name + "`") if _strat_path.exists() else "⚠️ No file")
        with _mc3:
            st.caption(f"**{t('sys_sections_sub')}**")
            _sy1, _sy2, _sy3 = st.columns(3)
            _sy1.metric("Total", len(_syst_raw))
            _sy2.metric("Lang.", _syst_raw["language"].nunique() if "language" in _syst_raw.columns else 0)
            _sy3.metric("Core", int((_syst_raw["priority"] == "core").sum()) if "priority" in _syst_raw.columns else 0)
            st.caption(("✅ `" + _syst_path.name + "`") if _syst_path.exists() else "⚠️ No file (fallback)")

        # ── Recompute hint banner ──────────────────────────────────────────────
        if st.session_state.get("terms_recompute_after_tsv"):
            st.info(tt(
                "⚠️ Byl nahrán nebo upraven soubor termínů. Pro aktualizaci shod v databázi spusť 🔄 Recompute níže.",
                "⚠️ Term file was uploaded or edited. Run 🔄 Recompute below to update matches in the database."),
                icon="⚠️")

        st.divider()

        # ── Tabs: Morphology / Stratigraphy / Systematic sections ─────────────
        _term_tabs = st.tabs([t("morpho_sub"), t("strat_sub"), t("sys_sections_sub")])

        # ════════════════════════════════════════════════════════════════════════
        # TAB: MORPHOLOGY
        # ════════════════════════════════════════════════════════════════════════
        with _term_tabs[0]:

            _mup_col, _mdl_col = st.columns(2)
            with _mup_col:
                new_morph = st.file_uploader(
                    tt("📂 Nahrát nový soubor morfologických termínů (TSV)",
                       "📂 Upload new morphological terms file (TSV)"),
                    type=["tsv"], key="morph_up")
                if new_morph and st.button(t("save_morpho_btn"), key="morph_up_save", type="primary"):
                    BASE_DIR.mkdir(parents=True, exist_ok=True)
                    _morph_write = _upath("paleon_morphology_terms.tsv")
                    _morph_write.parent.mkdir(parents=True, exist_ok=True)
                    _morph_write.write_bytes(new_morph.getbuffer())
                    MORPHOLOGY_FILE.write_bytes(new_morph.getbuffer())
                    reload_morphology_terms()
                    st.session_state["terms_recompute_after_tsv"] = True
                    st.success(tt(f"✅ Uloženo ({new_morph.size:,} B). Spusť Recompute níže.",
                                  f"✅ Saved ({new_morph.size:,} B). Run Recompute below."))
                    st.rerun()
            with _mdl_col:
                if _morph_path.exists():
                    st.download_button(
                        tt("📥 Stáhnout aktuální TSV", "📥 Download current TSV"),
                        data=_morph_path.read_bytes(),
                        file_name="paleon_morphology_terms.tsv",
                        mime="text/tab-separated-values",
                        key="morph_dl")

            # ── Full-width editor ─────────────────────────────────────────────
            st.caption(tt(
                f"Edituj přímo v tabulce — **{len(_morph_raw)}** termínů (včetně disabled). "
                "Sloupcové filtry jsou dostupné přes ikonu 🔍 v záhlaví sloupce. "
                "Po uložení spusť 🔄 Recompute.",
                f"Edit directly in the table — **{len(_morph_raw)}** terms (including disabled). "
                "Column filters available via 🔍 in the column header. "
                "After saving, run 🔄 Recompute."))

            _morph_col_cfg: Dict = {}
            if "enabled" in _morph_raw.columns:
                _morph_col_cfg["enabled"] = st.column_config.SelectboxColumn(
                    "Active", options=["1", "0"], width=70,
                    help="1 = použit při výpočtu shod / used for match computation")
            if "term" in _morph_raw.columns:
                _morph_col_cfg["term"] = st.column_config.TextColumn("Term", width=240)
            if "canonical" in _morph_raw.columns:
                _morph_col_cfg["canonical"] = st.column_config.TextColumn("Canonical", width=240)
            if "category" in _morph_raw.columns:
                _morph_col_cfg["category"] = st.column_config.TextColumn("Category", width=150)
            if "language" in _morph_raw.columns:
                _morph_col_cfg["language"] = st.column_config.TextColumn("Lang.", width=70)

            _morph_edited = st.data_editor(
                _morph_raw,
                column_config=_morph_col_cfg,
                hide_index=True,
                use_container_width=True,
                num_rows="dynamic",
                key="morph_data_editor",
            )
            _msave1, _msave2 = st.columns([1, 5])
            if _msave1.button(t("save_morpho_btn"), key="morph_edit_save", type="primary"):
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                _morph_write2 = _upath("paleon_morphology_terms.tsv")
                _morph_write2.parent.mkdir(parents=True, exist_ok=True)
                _morph_edited.to_csv(str(_morph_write2), sep="\t", index=False, encoding="utf-8")
                MORPHOLOGY_FILE.write_text(
                    _morph_edited.to_csv(sep="\t", index=False), encoding="utf-8")
                reload_morphology_terms()
                st.session_state["terms_recompute_after_tsv"] = True
                st.success(tt(f"✅ Uloženo {len(_morph_edited)} termínů. Spusť Recompute níže.",
                              f"✅ Saved {len(_morph_edited)} terms. Run Recompute below."))
                st.rerun()

        # ════════════════════════════════════════════════════════════════════════
        # TAB: STRATIGRAPHY
        # ════════════════════════════════════════════════════════════════════════
        with _term_tabs[1]:

            _sup_col, _sdl_col = st.columns(2)
            with _sup_col:
                new_strat = st.file_uploader(
                    tt("📂 Nahrát nový soubor stratigrafických termínů (TSV)",
                       "📂 Upload new stratigraphical terms file (TSV)"),
                    type=["tsv"], key="strat_up")
                if new_strat and st.button(t("save_strat_btn"), key="strat_up_save", type="primary"):
                    BASE_DIR.mkdir(parents=True, exist_ok=True)
                    _strat_write = _upath("stratigraphy_terms.tsv")
                    _strat_write.parent.mkdir(parents=True, exist_ok=True)
                    _strat_write.write_bytes(new_strat.getbuffer())
                    STRATIGRAPHY_FILE.write_bytes(new_strat.getbuffer())
                    reload_stratigraphy_terms()
                    st.session_state["terms_recompute_after_tsv"] = True
                    st.success(tt(f"✅ Uloženo ({new_strat.size:,} B). Spusť Recompute níže.",
                                  f"✅ Saved ({new_strat.size:,} B). Run Recompute below."))
                    st.rerun()
            with _sdl_col:
                if _strat_path.exists():
                    st.download_button(
                        tt("📥 Stáhnout aktuální TSV", "📥 Download current TSV"),
                        data=_strat_path.read_bytes(),
                        file_name="stratigraphy_terms.tsv",
                        mime="text/tab-separated-values",
                        key="strat_dl")

            st.caption(tt(
                f"Edituj přímo v tabulce — **{len(_strat_raw)}** termínů (včetně disabled). "
                "Sloupcové filtry jsou dostupné přes ikonu 🔍 v záhlaví sloupce. "
                "Po uložení spusť 🔄 Recompute.",
                f"Edit directly in the table — **{len(_strat_raw)}** terms (including disabled). "
                "Column filters available via 🔍 in the column header. "
                "After saving, run 🔄 Recompute."))

            _strat_col_cfg: Dict = {}
            if "enabled" in _strat_raw.columns:
                _strat_col_cfg["enabled"] = st.column_config.SelectboxColumn(
                    "Active", options=["1", "0"], width=70)
            if "term" in _strat_raw.columns:
                _strat_col_cfg["term"] = st.column_config.TextColumn("Term", width=220)
            if "canonical" in _strat_raw.columns:
                _strat_col_cfg["canonical"] = st.column_config.TextColumn("Canonical", width=220)
            if "category" in _strat_raw.columns:
                _strat_col_cfg["category"] = st.column_config.TextColumn("Category", width=140)
            if "language" in _strat_raw.columns:
                _strat_col_cfg["language"] = st.column_config.TextColumn("Lang.", width=70)
            if "match_mode" in _strat_raw.columns:
                _strat_col_cfg["match_mode"] = st.column_config.SelectboxColumn(
                    "Match mode", options=["exact", "prefix", "fuzzy"], width=100)
            if "source_term_en" in _strat_raw.columns:
                _strat_col_cfg["source_term_en"] = st.column_config.TextColumn("Source (EN)", width=180)

            _strat_edited = st.data_editor(
                _strat_raw,
                column_config=_strat_col_cfg,
                hide_index=True,
                use_container_width=True,
                num_rows="dynamic",
                key="strat_data_editor",
            )
            _ssave1, _ssave2 = st.columns([1, 5])
            if _ssave1.button(t("save_strat_btn"), key="strat_edit_save", type="primary"):
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                _strat_write2 = _upath("stratigraphy_terms.tsv")
                _strat_write2.parent.mkdir(parents=True, exist_ok=True)
                _strat_edited.to_csv(str(_strat_write2), sep="\t", index=False, encoding="utf-8")
                STRATIGRAPHY_FILE.write_text(
                    _strat_edited.to_csv(sep="\t", index=False), encoding="utf-8")
                reload_stratigraphy_terms()
                st.session_state["terms_recompute_after_tsv"] = True
                st.success(tt(f"✅ Uloženo {len(_strat_edited)} termínů. Spusť Recompute níže.",
                              f"✅ Saved {len(_strat_edited)} terms. Run Recompute below."))
                st.rerun()

        # ════════════════════════════════════════════════════════════════════════
        # TAB: SYSTEMATIC SECTIONS (přesunuto do tabu, raw load → vidíme i disabled)
        # ════════════════════════════════════════════════════════════════════════
        with _term_tabs[2]:
            st.caption(
                "Headings for the systematic/taxonomic section of the article in 12 languages. "
                "They determine where the `systematic` zone begins — blocks within it receive a +0.15 bonus; "
                "blocks outside it receive a penalty.  "
                "File: `systematic_sections.tsv`  |  "
                "columns: `enabled, language, heading, priority, weight, pattern, note`")

            with st.expander(t("sys_sections_edit"), expanded=True):
                _sys_edit = st.data_editor(
                    _syst_raw,
                    column_config={
                        "enabled":   st.column_config.SelectboxColumn("✓", width=50, options=["1", "0"]),
                        "language":  st.column_config.TextColumn("Language", width=110),
                        "heading":   st.column_config.TextColumn("Heading", width=220),
                        "priority":  st.column_config.SelectboxColumn("Priority", width=90,
                                        options=["core", "support", "weak"]),
                        "weight":    st.column_config.NumberColumn("Weight", width=60, min_value=0, max_value=100),
                        "pattern":   st.column_config.TextColumn("Regex", width=300),
                        "note":      st.column_config.TextColumn("Note", width=200),
                    },
                    hide_index=True,
                    use_container_width=True,
                    num_rows="dynamic",
                    key="sys_sec_editor",
                )
                if st.button(t("save_sys_sections_btn"), key="sys_sec_save", type="primary"):
                    _sys_path_w = _upath("systematic_sections.tsv")
                    BASE_DIR.mkdir(parents=True, exist_ok=True)
                    _sys_path_w.parent.mkdir(parents=True, exist_ok=True)
                    _sys_edit.to_csv(str(_sys_path_w), sep="\t", index=False, encoding="utf-8")
                    SYSTEMATIC_SECTIONS_FILE.write_text(
                        _sys_edit.to_csv(sep="\t", index=False, encoding="utf-8"), encoding="utf-8")
                    reload_systematic_sections()
                    st.success(tt(
                        f"Uloženo {len(_sys_edit)} řádků. Změny se projeví při příštím indexování.",
                        f"Saved {len(_sys_edit)} rows. Changes take effect at next indexing."))
                    st.rerun()

            _sys_up_col, _sys_dl_col = st.columns(2)
            with _sys_up_col:
                new_sys = st.file_uploader(
                    "Upload TSV (systematic_sections.tsv)", type=["tsv"], key="sys_sec_up")
                if new_sys and st.button(t("upload_file_btn"), key="sys_sec_file_save"):
                    _sys_path_up = _upath("systematic_sections.tsv")
                    _sys_path_up.parent.mkdir(parents=True, exist_ok=True)
                    _sys_path_up.write_bytes(new_sys.getbuffer())
                    SYSTEMATIC_SECTIONS_FILE.write_bytes(new_sys.getbuffer())
                    reload_systematic_sections()
                    st.success(tt(f"Uloženo ({new_sys.size} B).", f"Saved ({new_sys.size} B)."))
                    st.rerun()
            with _sys_dl_col:
                if _syst_path.exists():
                    st.download_button(
                        tt("📥 Stáhnout aktuální TSV", "📥 Download current TSV"),
                        data=_syst_path.read_bytes(),
                        file_name="systematic_sections.tsv",
                        mime="text/tab-separated-values",
                        key="sys_sec_dl")

        # ── Recompute (zvýrazněn po uploadu/editaci) ───────────────────────────
        st.divider()
        _recomp_pending = bool(st.session_state.get("terms_recompute_after_tsv"))
        _recomp_type    = "primary" if _recomp_pending else "secondary"
        _recomp_label   = (
            tt("🔄 Přepočítat shody — doporučeno po uploadu / editaci termínů",
               "🔄 Recompute matches — recommended after uploading / editing terms")
            if _recomp_pending else
            tt("🔄 Přepočítat termíny pro všechny schválené záznamy",
               "🔄 Recompute terms for all approved records")
        )
        if st.button(_recomp_label, key="settings_recompute_terms", type=_recomp_type):
            st.session_state.pop("terms_recompute_after_tsv", None)
            _rc_status = st.empty()
            _rc_bar    = st.progress(0.0)

            def _rc_cb(current: int, total: int):
                _tot = max(int(total or 0), 1)
                _cur = max(0, min(int(current or 0), _tot))
                _rc_bar.progress(_cur / _tot)
                _rc_status.text(tt(
                    f"Přepočítávám termíny: {_cur} / {_tot} ({int(_cur/_tot*100)} %) …",
                    f"Recomputing terms: {_cur} / {_tot} ({int(_cur/_tot*100)} %) …"))

            try:
                _rc_status.text(tt(
                    "Připravuji přepočet morfologických a stratigrafických termínů…",
                    "Preparing recompute of morphological and stratigraphic terms…"))
                _rc_n = recompute_all_term_matches(progress_cb=_rc_cb,
                                                   include_needs_review=True)
                _rc_bar.empty()
                _rc_status.success(tt(
                    f"✅ Hotovo! Přepočítáno {_rc_n} záznamů.",
                    f"✅ Done! Recomputed {_rc_n} records."))
            except Exception as _rc_exc:
                _rc_bar.empty()
                _rc_status.error(tt(
                    f"❌ Přepočet selhal: {_rc_exc}",
                    f"❌ Recompute failed: {_rc_exc}"))
                raise

    # ── LLM prompty ───────────────────────────────────────────────────────────
    with stab2:
        st.caption(f"File: `{PROMPTS_FILE}`  "
                   + ("✅ saved" if PROMPTS_FILE.exists() else "⚠️ nonexist"))

        PRESETS = {
            "Default (Czech)": (LLM_VALIDATION_PROMPT, LLM_BOUNDARY_PROMPT, LLM_FIELD_PROMPT),
            "English – strict": (
                "You are PaleoN validation assistant. Return ONLY valid JSON:\n"
                '{"action":"keep|reject|needs_review","reason":"brief","suggested_rank":"","confidence":"High|Medium|Low"}\n'
                "NEVER create records. NEVER invent data.",
                "You are PaleoN boundary assistant. Return ONLY valid JSON:\n"
                '{"assessment":"correct|too_short|too_long","reason":"brief","suggested_end_marker":""}\n'
                "NEVER invent missing text.",
                "You are PaleoN field assignment assistant.\n"
                'Return ONLY valid JSON: {"fields":{"FIELD":"verbatim"},"reasons":{},"confidence":"High|Medium|Low"}\n'
                "Use ONLY text from the input. Nover summarize.",
            ),
        }
        preset = st.selectbox(t("preset_load"), list(PRESETS.keys()), key="preset_sel")
        if st.button("Apply preset"):
            vp, bp, fp = PRESETS[preset]
            s.update({"llm_validation_prompt":vp,
                      "llm_boundary_prompt":bp,
                      "llm_field_prompt":fp})
            st.success(tt(f"Preset '{preset}' loaded.", f"Preset '{preset}' loaded."))

        s["llm_validation_prompt"] = st.text_area(
            "1️⃣ Validation prompt (keep / reject / needs_review)",
            value=s.get("llm_validation_prompt", LLM_VALIDATION_PROMPT), height=160)
        s["llm_boundary_prompt"] = st.text_area(
            "2️⃣ Boundary prompt (block boundary correctness)",
            value=s.get("llm_boundary_prompt", LLM_BOUNDARY_PROMPT), height=130)
        s["llm_field_prompt"] = st.text_area(
            "3️⃣ Field assignment prompt (field assignment)",
            value=s.get("llm_field_prompt", LLM_FIELD_PROMPT), height=160)
        s["llm_translation_prompt"] = st.text_area(
            "4️⃣ Translation prompt (translation of non-English fields → English)",
            value=s.get("llm_translation_prompt", LLM_TRANSLATION_PROMPT), height=200,
            help="Prompt pro automaticky preklad pri indexaci ne-anglickych documents. LLM vrati JSON {\"fields\":{\"FIELD\":\"English translation\"}}. PaleoN pak savei: anglicky text (puvodni text v zavorce).")

        pb1,pb2,pb3 = st.columns(3)
        if pb1.button("💾 Save prompty", type="primary", key="set_prompts_save"):
            save_prompts(s)
            st.success(tt(f"Saved → `{PROMPTS_FILE}`", f"Saved → `{PROMPTS_FILE}`"))
        if pb2.button("📂 Load from disc", key="set_prompts_load"):
            s.update(load_prompts()); st.success(tt("Loaded.", "Loaded.")); st.rerun()
        if pb3.button("⚠ Restore default", key="set_prompts_reset"):
            s.update({"llm_validation_prompt": LLM_VALIDATION_PROMPT,
                      "llm_boundary_prompt":   LLM_BOUNDARY_PROMPT,
                      "llm_field_prompt":      LLM_FIELD_PROMPT,
                      "llm_translation_prompt": LLM_TRANSLATION_PROMPT})
            st.success("Obnoveno.")

        st.session_state["paleon_settings"] = s

    # ── Data & persistence ────────────────────────────────────────────────────
    with stab3:
        # FTS5 rebuild
        st.markdown(t("fts_label"))
        st.caption(
            "The full-text index is updated automatically when fields are saved. Run rebuild after bulk import or pokud full-text search does not return the expected results.")
        if st.button("🔄 Rebuild FTS5 index", key="fts_rebuild"):
            with st.spinner(t("indexing_records")):
                n_fts = fts_rebuild_all()
            st.success(tt(f"FTS5: {n_fts} records indexovano.", f"FTS5: {n_fts} records indexed."))
        st.divider()
        ds1,ds2,ds3 = st.columns(3)
        if ds1.button("💾 Save all", type="primary", key="set_data_save"):
            save_settings_to_disc(s); save_prompts(s)
            st.success(tt(f"Saved → `{SETTINGS_FILE}` + `{PROMPTS_FILE}`", f"Saved → `{SETTINGS_FILE}` + `{PROMPTS_FILE}`"))
        if ds2.button("📂 Load from disc", key="set_data_load"):
            st.session_state["paleon_settings"] = load_settings_from_disc()
            st.success(tt("Loaded.", "Loaded.")); st.rerun()
        if ds3.button("⚠ Restore default", key="set_data_reset"):
            st.session_state["paleon_settings"] = DEFAULT_SETTINGS.copy()
            st.success("Obnoveno."); st.rerun()

        st.divider()
        st.subheader(t("db_sub"))
        if DB_FILE.exists():
            size_mb = DB_FILE.stat().st_size / 1024 / 1024
            st.caption(f"`{DB_FILE}`  —  {size_mb:.2f} MB")
        con = db()
        nd = con.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
        nc = con.execute("SELECT COUNT(*) FROM taxon_candidates").fetchone()[0]
        nf = con.execute("SELECT COUNT(*) FROM occurrence_fields").fetchone()[0]
        con.close()
        dc1,dc2,dc3 = st.columns(3)
        dc1.metric("Documents", nd)
        dc2.metric("Candidates", nc)
        dc3.metric("Saved fields", nf)

        st.divider()
        st.subheader(t("backup_sub"))
        st.caption(
            "Export/import the ENTIRE file `paleon.db` (Documents, candidates, fields) by one click. Settings (settings.json) and schema is backed up separately above.")

        bk1, bk2 = st.columns(2)
        with bk1:
            st.markdown("**Export**")
            if DB_FILE.exists():
                try:
                    # WAL checkpoint zajistí, že jsou v hlavním fileu
                    # propsané i nejnovější zápisy, které by jinak mohly
                    # zustat jen v -wal fileu.
                    con = db()
                    con.execute("PRAGMA wal_checkpoint(FULL)")
                    con.close()
                    db_bytes = DB_FILE.read_bytes()
                    st.download_button(
                        "📥 Download backup (.db)",
                        data=db_bytes,
                        file_name=f"paleon_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.db",
                        mime="application/octet-stream",
                        key="db_backup_download")
                except Exception as exc:
                    st.error(tt(f"Failed to prepare backup: {exc}", f"Failed to prepare backup: {exc}"))
            else:
                st.caption(t("db_not_exist"))

        with bk2:
            st.markdown("**Restore**")
            restore_file = st.file_uploader(
                "Upload backup (.db)", type=["db"], key="db_restore_upload")
            if restore_file is not None:
                st.warning(
                    "⚠ Restore PREPISE celou aktualni databazi. Doporuceno nejprve stahnout zalohu aktualniho stavu (vlevo).")
                if st.button(t("restore_db_btn"), key="db_restore_confirm_btn"):
                    st.session_state["db_confirm_restore"] = True

            if st.session_state.get("db_confirm_restore") and restore_file is not None:
                st.error(t("restore_confirm_warn"))
                ry, rn = st.columns(2)
                if ry.button("✅ Yes, overwrite", key="db_restore_yes", type="primary"):
                    try:
                        BASE_DIR.mkdir(parents=True, exist_ok=True)
                        DB_FILE.write_bytes(restore_file.getbuffer())
                        # Zahodit stopy WAL ze staré databáze, ať SQLite
                        # nezkouší replayovat žurnál patřící jinému fileu.
                        for suffix in ("-wal", "-shm"):
                            side = pathlib.Path(str(DB_FILE) + suffix)
                            if side.exists():
                                side.unlink()
                        st.session_state.pop("db_confirm_restore", None)
                        st.success(t("db_restored"))
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Restore selhala: {exc}")
                if rn.button("❌ Cancel", key="db_restore_no"):
                    st.session_state.pop("db_confirm_restore", None)
                    st.rerun()

        st.divider()
        files_info = [
            {"File": str(f), "Exists": "✅" if f.exists() else "❌",
             "Size": f"{f.stat().st_size/1024:.1f} KB" if f.exists() else "–"}
            for f in [SETTINGS_FILE, PROMPTS_FILE, SCHEMA_FILE, DB_FILE]
        ]
        st.dataframe(pd.DataFrame(files_info), use_container_width=True,
                     hide_index=True)

    # ── Kvalita dat ──────────────────────────────────────────────────────────
    with stab4:
        st.subheader(t("fuzzy_dup_sub"))
        st.caption(
            "It finds pairs of candidates with similar but not identical taxon names (typically a minor typo or OCR error in the taxon name in another paper). Exact matches are not displayed—these are expected repeated occurrences of the same taxon and do not constitute a problem.")

        fd1, fd2 = st.columns([3,1])
        fd_threshold = fd1.slider(
            "Min. name similarity", 0.70, 0.98, 0.86, 0.01, key="qc_fuzzy_thresh")
        run_fuzzy = fd2.button("🔍 Run detection", key="qc_fuzzy_run", type="primary")

        if run_fuzzy:
            with st.spinner(t("searching_similar")):
                dup_pairs = find_fuzzy_taxon_duplicates(fd_threshold)
            st.session_state["qc_fuzzy_results"] = dup_pairs

        dup_pairs = st.session_state.get("qc_fuzzy_results")
        if dup_pairs is not None:
            if not dup_pairs:
                st.success(t("no_dup_pairs"))
            else:
                st.warning(tt(f"Nalezeno {len(dup_pairs)} suspicious pairs.", f"Found {len(dup_pairs)} suspicious pairs."))
                df_dup = pd.DataFrame([{
                    "Similarity": p["podobnost"],
                    "ID A": p["id_a"], "Name A": p["name_a"],
                    "Document A": p["doc_a"], "Str. A": p["page_a"], "Status A": p["status_a"],
                    "ID B": p["id_b"], "Name B": p["name_b"],
                    "Document B": p["doc_b"], "Str. B": p["page_b"], "Status B": p["status_b"],
                } for p in dup_pairs])
                st.dataframe(df_dup, use_container_width=True, hide_index=True)
                st.caption(
                    "💡 To verify a specific entry, go to the Review tab—in the “Detail of selected record” expander, enter the ID.")

        st.divider()
        st.subheader(t("ocr_typos_sub"))
        st.caption(
            "Compare unique names in fields AUTHOR and find pair with high text similarity, that are not identical — a typical OCR-error signature (e.g. „Sysoey“ vs. „Sysoev“). These are suggestions for manual verification, nothing is edited automatically.")

        ad1, ad2 = st.columns([3,1])
        ad_threshold = ad1.slider(
            "Min. name similarity", 0.70, 0.95, 0.82, 0.01, key="qc_author_thresh")
        run_author = ad2.button("🔍 Run detection", key="qc_author_run", type="primary")

        if run_author:
            with st.spinner(t("comparing_authors")):
                typo_suggestions = find_author_ocr_typos(ad_threshold)
            st.session_state["qc_author_results"] = typo_suggestions

        typo_suggestions = st.session_state.get("qc_author_results")
        if typo_suggestions is not None:
            if not typo_suggestions:
                st.success(t("no_author_typos"))
            else:
                st.warning(tt(f"Nalezeno {len(typo_suggestions)} suspicious name pairs.", f"Found {len(typo_suggestions)} suspicious name pairs."))
                df_typo = pd.DataFrame([{
                    "Similarity": t["podobnost"],
                    "Probably correct": t["pravděpodobně_správně"],
                    "Occurrence": t["výskytu_správně"],
                    "Possible typo": t["možný_překlep"],
                    "Typo occurrences": t["výskytu_překlepu"],
                    "Example (taxon)": t["ukázka"]["taxon_name"],
                    "Document": t["ukázka"]["filename"],
                    "Page.": t["ukázka"]["page_start"],
                } for t in typo_suggestions])
                st.dataframe(df_typo, use_container_width=True, hide_index=True)
                st.caption(
                    "💡 Make the correction manually in the Editor for the specific record (AUTHOR field)—automatic correction could corrupt correctly spelled names.")

    # ── Gold Set Evaluátor ───────────────────────────────────────────────────
    with stab5:
        st.subheader(t("gold_set_sub"))
        st.caption(
            "Upload gold set DOCX (format: `Record N -- genus/species TaxonName` or `Record N TaxonName`), pair it with the Documentem in the library a run the benchmark. The evaluator compares PaleoN detection with the gold standard.")

        # ── Nahrání gold set DOCX ───────────────────────────────────────────
        gs_col1, gs_col2 = st.columns([3, 2])
        with gs_col1:
            st.markdown(t("upload_gold_docx"))
            gs_upload = st.file_uploader(
                "Gold set DOCX", type=["docx"], key="gs_upload",
                label_visibility="collapsed")

            if gs_upload:
                import tempfile, os
                tmp_path = pathlib.Path(tempfile.mktemp(suffix=".docx"))
                tmp_path.write_bytes(gs_upload.read())
                try:
                    preview_recs = parse_goldset_docx(str(tmp_path))
                    st.success(tt(f"Parsed {len(preview_recs)} records z DOCX.", f"Parsed {len(preview_recs)} records from DOCX."))
                    prev_df = pd.DataFrame([{
                        "#": r["n"], "Taxon": r["taxon_name"],
                        "Rank": r.get("rank",""),
                        "Pole": ", ".join(r.get("fields",{}).keys()),
                    } for r in preview_recs])
                    st.dataframe(prev_df, use_container_width=True, hide_index=True)
                    st.session_state["gs_tmp_path"] = str(tmp_path)
                    st.session_state["gs_preview"] = preview_recs
                except Exception as exc:
                    st.error(tt(f"Parse error: {exc}", f"Parsing error: {exc}"))

        with gs_col2:
            st.markdown(t("pair_with_lib"))
            con = db()
            docs = con.execute(
                "SELECT id, filename FROM documents ORDER BY filename").fetchall()
            existing_gs = con.execute(
                "SELECT id, filename, linked_document_id FROM goldset_documents "
                "ORDER BY created_at DESC").fetchall()
            con.close()

            doc_options = ["— Select —"] + [d["filename"] for d in docs]
            sel_doc_label = st.selectbox(
                "Document", doc_options, key="gs_doc_sel",
                label_visibility="collapsed")
            sel_doc_id = next(
                (d["id"] for d in docs if d["filename"] == sel_doc_label), None)

            st.markdown(t("save_evaluate"))
            import_btn = st.button(
                "💾 Import gold set into DB", key="gs_import_btn",
                disabled=not (st.session_state.get("gs_tmp_path") and sel_doc_id),
                type="primary")
            if import_btn:
                tmp = st.session_state.get("gs_tmp_path", "")
                if tmp and pathlib.Path(tmp).exists():
                    try:
                        gid = import_goldset_to_db(tmp, sel_doc_id)
                        st.session_state["gs_last_gid"] = gid
                        st.success(tt(f"Gold set saved (ID={gid}).", f"Gold set saved (ID={gid})."))
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Import selhal: {exc}")

        st.divider()

        # ── Výběr saveeného gold setu a spuštění benchmarku ─────────────────
        if existing_gs:
            st.markdown(t("gold_sets_saved_hdr"))
            gs_options = {
                f"#{g['id']} — {g['filename']} → dok. {g['linked_document_id'] or '—'}": g
                for g in existing_gs
            }
            sel_gs_label = st.selectbox(
                "Gold set", list(gs_options.keys()), key="gs_run_sel",
                label_visibility="collapsed")
            sel_gs = gs_options.get(sel_gs_label)

            run_col, del_col = st.columns([3, 1])
            run_btn = run_col.button(
                "▶️ Spustit benchmark", key="gs_run_btn", type="primary",
                disabled=not (sel_gs and sel_gs["linked_document_id"]))
            if del_col.button("🗑️ Delete gold set", key="gs_del_btn"):
                if sel_gs:
                    con = db()
                    con.execute("DELETE FROM goldset_records WHERE goldset_doc_id=?",
                                (sel_gs["id"],))
                    con.execute("DELETE FROM goldset_documents WHERE id=?",
                                (sel_gs["id"],))
                    con.commit(); con.close()
                    st.success(tt("Deleted.", "Deleted.")); st.rerun()

            if run_btn and sel_gs:
                with st.spinner(t("run_benchmark")):
                    results = evaluate_goldset(
                        sel_gs["id"], sel_gs["linked_document_id"])
                st.session_state["gs_results"] = results
                st.session_state["gs_results_meta"] = {
                    "gs_name": sel_gs["filename"],
                    "doc_id": sel_gs["linked_document_id"],
                }

        # ── Zobrazení výsledku ───────────────────────────────────────────────
        results = st.session_state.get("gs_results", [])
        if results:
            meta = st.session_state.get("gs_results_meta", {})
            n_total = len(results)
            n_det   = sum(1 for r in results if r.detected)
            n_cont  = sum(1 for r in results if r.block_contains_gold)
            n_short = sum(1 for r in results if r.block_too_short)
            n_long  = sum(1 for r in results if r.block_too_long)
            recall  = n_det / n_total if n_total else 0

            # Přesnost fields
            all_fields_gs: List[str] = []
            for r in results:
                for k in r.field_results:
                    if k not in all_fields_gs:
                        all_fields_gs.append(k)
            field_acc = {}
            for f in all_fields_gs:
                vals = [r.field_results[f] for r in results if f in r.field_results]
                field_acc[f] = sum(vals)/len(vals) if vals else 0.0

            # Metriky
            mc1, mc2, mc3, mc4, mc5 = st.columns(5)
            mc1.metric("Records gold setu", n_total)
            mc2.metric("Recall detekce", f"{recall:.0%}",
                       delta=f"{n_det}/{n_total}")
            mc3.metric("Blok OK (obsahuje text)", n_cont)
            mc4.metric("⚠ Block is too short", n_short)
            mc5.metric("⚠ Block is too long", n_long)

            # Přesnost fields
            with st.expander(t("mapping_accuracy"), expanded=True):
                field_df = pd.DataFrame([
                    {"Pole": f,
                     "Correct": f"{v:.0%}",
                     "Correct (#)": int(sum(r.field_results.get(f, False)
                                            for r in results)),
                     "Celkem":  sum(1 for r in results if f in r.field_results)}
                    for f, v in sorted(field_acc.items(), key=lambda x: -x[1])
                ])
                st.dataframe(field_df, use_container_width=True, hide_index=True)

            # Detail records
            with st.expander(t("record_detail_exp"), expanded=False):
                for r in results:
                    icon = "✅" if r.detected and r.block_contains_gold else (
                           "⚠️" if r.detected else "❌")
                    st.markdown(
                        f"**{icon} #{r.gold_n} {r.gold_name}** "
                        f"[{r.gold_rank or '?'}]  →  "
                        f"{'`' + (r.matched_name or '') + '`' if r.detected else '*nenalezeno*'}  "
                        f"shoda: {r.name_score:.2f}")
                    if r.detected:
                        flags = []
                        if r.block_too_short: flags.append("⚠️ block too short")
                        if r.block_too_long:  flags.append("⚠️ block too long")
                        if not r.block_contains_gold: flags.append("❌ blok neobsahuje gold text")
                        if flags:
                            st.caption("  ".join(flags))
                        if r.field_results:
                            field_summary = "  ".join(
                                f"{'✅' if ok else '❌'} {lbl}"
                                for lbl, ok in r.field_results.items())
                            st.caption(f"Pole: {field_summary}")
                    st.markdown("---")

            # Export
            con = db()
            doc_name_r = con.execute(
                "SELECT filename FROM documents WHERE id=?",
                (meta.get("doc_id"),)).fetchone()
            con.close()
            doc_name = doc_name_r["filename"] if doc_name_r else "unknown"

            xlsx_bytes = export_goldset_report_xlsx(
                results, meta.get("gs_name", ""), doc_name)
            st.download_button(
                "📥 Download benchmark report (XLSX)",
                data=xlsx_bytes,
                file_name=f"paleon_goldset_report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="gs_xlsx_dl")
        elif not existing_gs:
            st.info(t("no_gold_sets"))


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════


    # ── Uživatelé ─────────────────────────────────────────────────────────────
    with stab_users:
        st.subheader(t("user_mgmt_sub"))
        _cur = st.session_state.get("pn_user", "")

        # ── Admin přihlášení ──────────────────────────────────────────────────
        if not st.session_state.get("admin_auth"):
            st.info(t("admin_required"))
            ap1, ap2 = st.columns([3, 1])
            _ap = ap1.text_input("Admin heslo", type="password", key="admin_pw_input",
                                  label_visibility="collapsed", placeholder="Admin heslo…")
            if ap2.button("Odemknout", type="primary", key="admin_pw_btn"):
                if _check_admin_password(_ap):
                    st.session_state["admin_auth"] = True
                    st.rerun()
                else:
                    st.error(t("wrong_password_admin"))
            st.stop()

        st.success(t("admin_unlocked"))
        if st.button("🔒 Zamknout", key="admin_lock"):
            st.session_state.pop("admin_auth", None)
            st.rerun()

        st.divider()

        # ── Seznam uživatelu ──────────────────────────────────────────────────
        all_users = _get_all_users()
        st.markdown(tt(f"**Total users: {len(all_users)}**", f"**Total users: {len(all_users)}**"))

        for usr in all_users:
            uname = usr["username"]
            disp  = usr["display_name"] or uname
            is_adm = bool(usr["is_admin"])
            has_pw = bool(usr["password_hash"])
            udir   = USERS_DIR / _sanitize_username(uname)

            # Počet documents v DB
            n_docs = 0
            udb = udir / "paleon.db"
            if udb.exists():
                try:
                    _uc = sqlite3.connect(str(udb))
                    n_docs = _uc.execute("SELECT COUNT(*) FROM documents").fetchone()[0]
                    _uc.close()
                except Exception:
                    pass

            with st.expander(
                f"{'👑' if is_adm else '👤'} **{disp}** (`{uname}`) — "
                f"{n_docs} documents {'🔑' if has_pw else ''}",
                expanded=False):

                uc1, uc2, uc3 = st.columns(3)
                uc1.markdown(f"**Created:** {(usr['created_at'] or '')[:10]}")
                uc2.markdown(f"**Naposledy:** {(usr['last_seen'] or '')[:10]}")
                uc3.markdown(f"**Admin:** {'Yes' if is_adm else 'No'}")

                # Nastavit / změnit heslo
                with st.expander("🔑 Heslo", expanded=False):
                    new_pw = st.text_input(
                        "Now password (empty = without hesla)",
                        type="password", key=f"pw_{uname}",
                        label_visibility="collapsed",
                        placeholder="Now password or leave blank…")
                    if st.button(t("save_password_btn"), key=f"pw_save_{uname}"):
                        pw_hash = hashlib.sha256(new_pw.encode()).hexdigest() if new_pw else None
                        _uc2 = _users_db()
                        _uc2.execute("UPDATE users SET password_hash=? WHERE username=? COLLATE NOCASE",
                                     (pw_hash, uname))
                        _uc2.commit(); _uc2.close()
                        st.success(tt("Password saved.", "Password saved.") if pw_hash else tt("Password removed.", "Password removed."))
                        st.rerun()

                # Admin toggle
                if uname != _cur:
                    new_admin = st.checkbox(
                        "Admin rights", value=is_adm, key=f"adm_{uname}")
                    if new_admin != is_adm:
                        _uc3 = _users_db()
                        _uc3.execute("UPDATE users SET is_admin=? WHERE username=? COLLATE NOCASE",
                                     (1 if new_admin else 0, uname))
                        _uc3.commit(); _uc3.close()
                        st.rerun()

                # Slovníky pro tohoto uživatele
                with st.expander(t("dicts_expander"), expanded=False):
                    for dict_file, dict_label in [
                        ("section_schema.tsv", "Schema"),
                        ("paleon_morphology_terms.tsv", "Morphology"),
                        ("stratigraphy_terms.tsv", "Stratigraphy"),
                        ("taxons.txt", "Gazetteer"),
                    ]:
                        df_path = udir / dict_file
                        sys_path = SYSTEM_DIR / dict_file
                        exists = df_path.exists()
                        dc1, dc2, dc3 = st.columns([2, 1, 1])
                        dc1.markdown(
                            f"{'✅' if exists else '❌'} {dict_label} "
                            f"({'exists' if exists else 'missing'})")
                        if dc2.button("↺ From system", key=f"sys_{uname}_{dict_file}",
                                      help="Prepise slovnik verze ze system/"):
                            ok = _copy_system_dict_to_user(uname, dict_file)
                            st.success(tt(f"{dict_label} copied.", f"{dict_label} copied.") if ok else tt("System dictionary not found.", "System dictionary not found."))
                            st.rerun()
                        if exists and dc3.button("⬇", key=f"dl_{uname}_{dict_file}",
                                                  help="Download dictionary"):
                            st.download_button(
                                f"📥 {dict_label}",
                                data=df_path.read_bytes(),
                                file_name=dict_file, key=f"dlb_{uname}_{dict_file}")

                # Delete uživatele
                if uname != _cur:
                    st.divider()
                    del_key = f"confirm_del_{uname}"
                    if not st.session_state.get(del_key):
                        if st.button(tt(f"🗑️ Delete user {uname}", f"🗑️ Delete user {uname}"),
                                     type="secondary", key=f"del_{uname}"):
                            st.session_state[del_key] = True
                            st.rerun()
                    else:
                        st.warning(tt(f"⚠️ Opravdu smazat **{uname}** and all their data?", f"⚠️ Really delete **{uname}** and all their data?"))
                        c1, c2 = st.columns(2)
                        if c1.button("Yes, smazat", type="primary", key=f"del_yes_{uname}"):
                            _delete_user(uname)
                            st.session_state.pop(del_key, None)
                            st.success(tt(f"User {uname} deleted.", f"User {uname} deleted."))
                            st.rerun()
                        if c2.button("Cancel", key=f"del_no_{uname}"):
                            st.session_state.pop(del_key, None)
                            st.rerun()

        st.divider()

        # ── Vytvořit nového uživatele (admin panel) ───────────────────────────
        st.subheader(t("add_user_sub"))
        na1, na2, na3 = st.columns([2, 2, 2])
        adm_new_name = na1.text_input("Username", key="adm_new_name")
        adm_new_disp = na2.text_input("Display name", key="adm_new_disp")
        adm_new_pw   = na3.text_input("Password (optional)", type="password", key="adm_new_pw")
        adm_is_admin = st.checkbox(t("admin_rights"), key="adm_new_is_admin")
        if st.button(t("create_btn2"), type="primary", key="adm_create_btn"):
            if not adm_new_name.strip():
                st.warning(t("name_required"))
            elif not re.match(r"^[a-zA-Z0-9_\-]{2,40}$", adm_new_name.strip()):
                st.warning(t("username_invalid"))
            else:
                ok = _create_user(
                    adm_new_name.strip().lower(),
                    adm_new_disp.strip() or adm_new_name.strip(),
                    adm_new_pw.strip() or None,
                    is_admin=adm_is_admin)
                if ok:
                    st.success(tt(f"User '{adm_new_name}' created.", f"User '{adm_new_name}' created."))
                    st.rerun()
                else:
                    st.error(tt("User s timto jmenem jiz exists.", "A user with this name already exists."))

        st.divider()
        st.subheader(t("update_sys_dicts"))
        st.caption(
            "System dictionaries serve as a template for new users. Existing users can update their dictionary manually above.")
        if st.button(t("publish_dicts_btn"),
                     key="pub_sys_dicts"):
            import shutil
            n_pub = 0
            for fname, src in [
                ("section_schema.tsv",          _upath("section_schema.tsv")),
                ("paleon_morphology_terms.tsv",  _upath("paleon_morphology_terms.tsv")),
                ("stratigraphy_terms.tsv",       _upath("stratigraphy_terms.tsv")),
                ("taxons.txt",                   _upath("taxons.txt")),
            ]:
                if src.exists():
                    shutil.copy2(str(src), str(SYSTEM_DIR / fname))
                    n_pub += 1
            st.success(tt(f"Published {n_pub} dictionaries.", f"Published {n_pub} dictionaries."))







    # ── Duplikáty napříč Documenty ────────────────────────────────────────────
    with stab_dup:
        st.subheader(t("similar_taxa_sub"))
        st.caption(
            "Hleda pary kandidatu se stejnym or velmi podobnym nazvem taxonu v ruznych documents. Umozni sloucit chybejici fields or propojit zaznamy jako synonyma.")

        _dup_c1, _dup_c2, _dup_c3 = st.columns([1, 1, 2])
        _dup_threshold = _dup_c1.slider(
            "Minimum similarity", 0.70, 1.00, 0.86, 0.01,
            key="dup_threshold")
        _dup_cross_only = _dup_c2.toggle(
            "Jen z ruznych documents", value=True, key="dup_cross_only",
            help="Turn off for searching also in one document.")

        if _dup_c3.button("🔍 Run analyzu", key="dup_run", type="primary"):
            with st.spinner(t("comparing_taxa")):
                pairs = find_fuzzy_taxon_duplicates(threshold=_dup_threshold)
            if _dup_cross_only:
                pairs = [p for p in pairs if p["doc_a"] != p["doc_b"]]
            st.session_state["dup_pairs"] = pairs
            st.success(tt(f"Nalezeno {len(pairs)} pairs.", f"Found {len(pairs)} pairs."))

        pairs = st.session_state.get("dup_pairs", [])
        if not pairs:
            st.info(tt("Run analysis above.", "Run the analysis above."))
            st.stop()

        st.markdown(tt(f"**{len(pairs)} pairs** (sorted ↓ by similarity):", f"**{len(pairs)} pairs** (sorted ↓ by similarity):"))
        for idx_p, pair in enumerate(pairs[:100]):
            sim_pct = int(pair["podobnost"] * 100)
            sim_color = ("#22c55e" if sim_pct >= 95 else
                         "#f59e0b" if sim_pct >= 86 else "#94a3b8")
            with st.expander(
                f"[{sim_pct}%] **{pair['name_a']}** ↔ **{pair['name_b']}**",
                expanded=False):

                dc1, dc2 = st.columns(2)
                with dc1:
                    _mda = (f"**A:** `{pair['name_a']}`\n"
                            f"📄 {pair['doc_a']}, str.{pair['page_a']}\n"
                            f"Status: {pair['status_a']}")
                    st.markdown(_mda)
                    # Ukázka fields A
                    _fa = get_candidate_fields(pair["id_a"])
                    for fn in ["DIAGNOSIS","LOCALITY","STRATIGRAPHY","TYPE SPECIMENS"]:
                        if _fa.get(fn):
                            st.caption(f"**{fn}:** {_fa[fn][:120]}…"
                                       if len(_fa.get(fn,""))>120 else
                                       f"**{fn}:** {_fa.get(fn,'')}")

                with dc2:
                    _mdb = (f"**B:** `{pair['name_b']}`\n"
                            f"📄 {pair['doc_b']}, str.{pair['page_b']}\n"
                            f"Status: {pair['status_b']}")
                    st.markdown(_mdb)
                    _fb = get_candidate_fields(pair["id_b"])
                    for fn in ["DIAGNOSIS","LOCALITY","STRATIGRAPHY","TYPE SPECIMENS"]:
                        if _fb.get(fn):
                            st.caption(f"**{fn}:** {_fb[fn][:120]}…"
                                       if len(_fb.get(fn,""))>120 else
                                       f"**{fn}:** {_fb.get(fn,'')}")

                # Akce
                act1, act2, act3, act4 = st.columns(4)

                if act1.button("A ← B (doplnit)", key=f"merge_ab_{idx_p}",
                               help="Copies missing field from B to A"):
                    n = merge_candidate_fields(pair["id_a"], pair["id_b"],
                                              overwrite=False)
                    st.success(tt(f"Copied {n} fields do A.", f"Copied {n} fields to A."))
                    st.rerun()

                if act2.button("B ← A (doplnit)", key=f"merge_ba_{idx_p}",
                               help="Copies missing field from A to B"):
                    n = merge_candidate_fields(pair["id_b"], pair["id_a"],
                                              overwrite=False)
                    st.success(tt(f"Copied {n} fields do B.", f"Copied {n} fields to B."))
                    st.rerun()

                if act3.button("🔗 Synonyma", key=f"link_syn_{idx_p}",
                               help="Connects records as synonyms (RELATED_RECORD_ID)"):
                    link_as_synonym(pair["id_a"], pair["id_b"])
                    st.success(tt("Records propojeny jako synonyma.", "Records linked as synonyms."))
                    st.rerun()

                if act4.button("✏️ Editor →", key=f"dup_edit_{idx_p}",
                               help="Open record A in Editoru"):
                    st.session_state["editor_candidate_id"] = pair["id_a"]
                    st.session_state["active_tab"] = "editor"
                    st.rerun()

        if len(pairs) > 100:
            st.caption(tt(f"… a {len(pairs)-100} more paru", f"… and {len(pairs)-100} more pairs"))


def _show_user_selection() -> None:
    """
    Login screen — displayed before the main application.
    The user clicks on their jméno or vytvoří nový účet.
    Heslo je optional; pokud uživatel heslo má, musí ho zadat.
    """
    st.set_page_config(
        page_title=f"{APP_NAME} — Login",
        page_icon="🦕", layout="centered",
    )
    st.markdown(
        "<div style='text-align:center;padding:40px 0 20px'>"
        "<span style='font-size:3rem'>🦕</span>"
        f"<h1 style='margin:8px 0 4px'>{APP_NAME}</h1>"
        f"<p style='color:#64748b;margin:0'>v{APP_VERSION} </p>"
        "</div>",
        unsafe_allow_html=True)

    users = _get_all_users()

    if users:
        st.markdown(t("select_user_md"))
        # Grid tlačítek — max 4 per row
        cols_per_row = 4
        for row_start in range(0, len(users), cols_per_row):
            row_users = users[row_start:row_start+cols_per_row]
            cols = st.columns(len(row_users))
            for col, user in zip(cols, row_users):
                with col:
                    label = (
                        f"{'👑 ' if user['is_admin'] else '👤 '}"
                        f"{user['display_name'] or user['username']}")
                    if col.button(label, key=f"sel_{user['username']}",
                                  use_container_width=True):
                        # Zkontrolovat heslo (pokud uživatel má)
                        if user["password_hash"]:
                            st.session_state["pn_pending_user"] = user["username"]
                        else:
                            _switch_user(user["username"])
                            st.rerun()

    # Zadání hesla pro uživatele s heslem
    pending = st.session_state.get("pn_pending_user", "")
    if pending:
        st.divider()
        st.markdown(tt(f"**Password for user `{pending}`:**", f"**Password for user `{pending}`:**"))
        entered_pw = st.text_input(
            "Heslo", type="password", key="login_pw",
            label_visibility="collapsed")
        lc1, lc2 = st.columns(2)
        if lc1.button("Log in", type="primary", key="login_btn"):
            pw_hash = hashlib.sha256(entered_pw.encode()).hexdigest()
            con = _users_db()
            user_row = con.execute(
                "SELECT password_hash FROM users WHERE username=? COLLATE NOCASE",
                (pending,)).fetchone()
            con.close()
            if user_row and hmac.compare_digest(
                    user_row["password_hash"] or "", pw_hash):
                st.session_state.pop("pn_pending_user", None)
                _switch_user(pending)
                st.rerun()
            else:
                st.error(t("wrong_password_simple"))
        if lc2.button("Back", key="login_back"):
            st.session_state.pop("pn_pending_user", None)
            st.rerun()

    st.divider()

    # Vytvoření nového účtu
    with st.expander(t("new_user_expander"), expanded=not users):
        st.caption(t("unique_username"))
        nc1, nc2 = st.columns(2)
        new_name = nc1.text_input(
            "Name", key="new_user_name",
            placeholder="e.g. martin, honza…")
        new_disp = nc2.text_input(
            "Display name (optional)", key="new_user_disp",
            placeholder="Martin Novak")
        new_pw   = st.text_input(
            "Heslo (optional — leave empty for sing in without password)",
            type="password", key="new_user_pw")
        if st.button(t("create_account_btn"), key="new_user_btn", type="primary"):
            if not new_name.strip():
                st.warning(t("name_required"))
            elif not re.match(r"^[a-zA-Z0-9_\-]{2,40}$", new_name.strip()):
                st.warning(t("name_letters_only"))
            else:
                ok = _create_user(
                    new_name.strip().lower(),
                    new_disp.strip() or new_name.strip(),
                    new_pw.strip() or None,
                    is_admin=False)
                if ok:
                    _switch_user(new_name.strip().lower())
                    st.success(tt(f"Account '{new_name}' created!", f"Account '{new_name}' created!"))
                    st.rerun()
                else:
                    st.error(tt("User s timto jmenem jiz exists.", "A user with this name already exists."))


def main():
    if st is None:
        print("Streamlit is not installed. Run: pip install streamlit")
        return

    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon="🦕",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # ── Multi-user init ──────────────────────────────────────────────────
    _init_user_registry()
    if not st.session_state.get("pn_user"):
        _show_user_selection()
        return

    s = get_settings()
    # ── Auto-detekce LM Studio (jednou za sezení) ────────────────────────────
    if "lm_studio_detected" not in st.session_state:
        try:
            _detected_models = lm_models(s)
            if _detected_models:
                st.session_state["lm_studio_detected"] = True
                st.session_state["lm_models"] = _detected_models
                _changed = False
                if not s.get("llm_enabled"):
                    s["llm_enabled"] = True; _changed = True
                if not s.get("lmstudio_model") and _detected_models:
                    s["lmstudio_model"] = _detected_models[0]; _changed = True
                if _changed:
                    save_settings_to_disc(s)
            else:
                st.session_state["lm_studio_detected"] = False
        except Exception:
            st.session_state["lm_studio_detected"] = False
    st.session_state["app_lang"] = "en"
    inject_css(s.get("theme", "dark"))
    init_db()

    # Auto-kopírovat schema TSV z aktuálního adresáře při prvním spuštění
    if not SCHEMA_FILE.exists():
        import shutil
        for candidate in [
            pathlib.Path("section_schema.tsv"),
            pathlib.Path("section_schema_from_prompt_no_notes.tsv"),
        ]:
            if candidate.exists():
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                shutil.copy(str(candidate), str(SCHEMA_FILE))
                break

    # Auto-kopírovat gazetteer (seznam známých taxa) z aktuálního adresáře
    if not GAZETTEER_FILE.exists():
        import shutil
        for candidate in [
            pathlib.Path("taxons.txt"),
            pathlib.Path("taxons.tsv"),
        ]:
            if candidate.exists():
                BASE_DIR.mkdir(parents=True, exist_ok=True)
                shutil.copy(str(candidate), str(GAZETTEER_FILE))
                break

    # Auto-kopírovat morfologické a stratigrafické termíny
    if not MORPHOLOGY_FILE.exists():
        import shutil
        candidate = pathlib.Path("paleon_morphology_terms.tsv")
        if candidate.exists():
            BASE_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy(str(candidate), str(MORPHOLOGY_FILE))
    if not STRATIGRAPHY_FILE.exists():
        import shutil
        candidate = pathlib.Path("stratigraphy_terms.tsv")
        if candidate.exists():
            BASE_DIR.mkdir(parents=True, exist_ok=True)
            shutil.copy(str(candidate), str(STRATIGRAPHY_FILE))

    sidebar_ui()

    # ── Starý globální session-state chunk recompute deaktivován ───────────────
    # Přecount termínu nyní běží přímo v Morpho/Strat přes progress_callback.
    st.session_state.pop("terms_recompute_ids", None)
    st.session_state.pop("terms_recompute_done", None)

    tabs = st.tabs([
        t("tab_library"),
        t("tab_review"),
        "🧩 Block Editor",
        t("tab_editor"),
        t("tab_dossier"),
        t("tab_morpho"),
        t("tab_export"),
        t("tab_settings"),
    ])
    with tabs[0]: tab_library()
    with tabs[1]: tab_review()
    with tabs[2]: tab_block_editor()
    with tabs[3]: tab_editor()
    with tabs[4]: tab_taxon_dossier()
    with tabs[5]: tab_morphostrat_dossier()
    with tabs[6]: tab_export()
    with tabs[7]: tab_settings()



if __name__ == "__main__":
    main()
