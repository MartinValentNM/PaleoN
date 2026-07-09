# Contributing to PaleoN

Thank you for considering a contribution to PaleoN.

## Preferred workflow

1. Open an issue describing the problem, proposed feature, or dataset-specific parsing case.
2. Create a focused branch.
3. Keep changes small and documented.
4. Test the app locally with:

```bash
pip install -r requirements.txt
streamlit run paleon_st.py
```

## Parser changes

Taxonomic parsing should remain conservative. Prefer changes that reduce false positives, expose diagnostics, and preserve verbatim source text. Do not add rules that silently infer data not present in the source document.

## Data and examples

Do not commit copyrighted PDFs, local databases, user exports, or private research data. Use small synthetic examples or clearly licensed sample text.

## Acknowledgement

This work was supported by the Ministry of Culture of the Czech Republic through project DKRVO 2024–2028/2.I.c (NM 00023272).
