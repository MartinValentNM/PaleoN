# PaleoN

**PaleoN** je lokální Streamlit aplikace pro konzervativní vícejazyčnou extrakci, kontrolu, editaci a export paleontologických taxonomických záznamů z dokumentů PDF, DOCX a TXT.

- **Autor:** Martin Valent, Národní Muzeum Praha — Oddělení Paleontologie
- **Projekt:** DKRVO 2024–2028/2.I.c (NM 00023272)
- **Licence:** MIT

## Hlavní funkce

- Lokální Streamlit rozhraní se SQLite backendem.
- Víceuživatelský režim s oddělenými datovými složkami a slovníky.
- Konzervativní detekce kandidátů taxonů s ručním schválením před exportem.
- Taxon Dossier pro vyhledávání taxonu napříč celou lokální knihovnou.
- Morfologický a stratigrafický dossier založený na editovatelných TSV slovnících.
- Podporované vstupy: PDF, DOCX, TXT.
- Volitelná OCR podpora přes Tesseract/easyOCR/PyMuPDF podle lokální instalace.
- Volitelná integrace LM Studio pro asistovanou validaci, mapování polí a překlady.
- Exporty strukturovaných tabulek i lidsky čitelných reportů.

## Rychlé spuštění

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run paleon_st.py
```

Aplikace si při spuštění automaticky vytvoří pracovní složku `paleon_data/`.

## Volitelné OCR

Pro OCR nainstalujte balíčky uvedené v `requirements-optional.txt`. Tesseract OCR může podle operačního systému vyžadovat také systémovou instalaci a jazyková data.

## Integrace LM Studio

LLM asistence je volitelná a ve výchozím nastavení vypnutá. Pokud ji chcete použít, spusťte LM Studio lokálně a nastavte v aplikaci base URL a model. PaleoN je navržen tak, aby výstupy LLM byly pouze asistivní; taxonomické záznamy má vždy kontrolovat a schvalovat uživatel.

## Obsah repozitáře

- `paleon_st.py` — hlavní Streamlit aplikace.
- `requirements.txt` — základní Python závislosti.
- `requirements-optional.txt` — volitelné OCR a doplňkové exportní závislosti.
- `README.md` — dokumentace v angličtině.
- `README.cs.md` — dokumentace v češtině.
- `LICENSE` — licence MIT.
- `CITATION.cff` — citační metadata.
- `CONTRIBUTING.md` — pravidla pro přispívání.
- `SECURITY.md` — bezpečnostní politika.
- `.gitignore` — ignoruje lokální data aplikace a generované soubory.

## Data a soukromí

PaleoN je zamýšlen jako lokální výzkumný nástroj. Nahrané dokumenty, vytvořené databáze, nastavení, exporty a uživatelské slovníky se ukládají lokálně do `paleon_data/`, pokud je uživatel výslovně nepřesune nebo nezveřejní. Složku `paleon_data/` **necommitujte** do GitHub repozitáře.

## Poděkování

Tato práce vznikla s podporou Ministerstva kultury České republiky v rámci projektu DKRVO 2024–2028/2.I.c (NM 00023272).

## Licence

Projekt je zveřejněn pod licencí MIT. Viz soubor [`LICENSE`](LICENSE).
