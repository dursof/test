"""
Estrattore DDT (Documento di Trasporto) da PDF
Legge PDF da una cartella, estrae articoli e dati intestazione, salva in CSV.
File già processati vengono saltati.

Requisiti Windows:
  - Python 3.8+
  - Tesseract OCR installato (https://github.com/UB-Mannheim/tesseract/wiki)
    con lingua italiana (Tessdata Italian)
  - pip install pdf2image pytesseract pillow

Uso:
  extract_ddt.exe [cartella_pdf] [--csv output.csv] [--reset]

  cartella_pdf : cartella con i PDF da processare (default: cartella corrente)
  --csv        : nome file CSV di output (default: ddt_articoli.csv)
  --reset      : forza ri-elaborazione di tutti i file
"""

import os
import re
import csv
import sys
import json
import argparse
import traceback
from pathlib import Path
from datetime import datetime

# ── Tesseract path auto-detection per Windows ──────────────────────────────
def setup_tesseract():
    import pytesseract
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        r"C:\Users\{}\AppData\Local\Programs\Tesseract-OCR\tesseract.exe".format(
            os.environ.get("USERNAME", "")
        ),
    ]
    for path in candidates:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return True
    # fallback: si assume che tesseract sia nel PATH
    return True


# ── Costanti ────────────────────────────────────────────────────────────────
CSV_HEADERS = [
    "file_pdf", "numero_ddt", "data_ddt", "fornitore",
    "n_ordine_cliente", "n_ordine_fornitore",
    "codice_articolo", "codice_articolo_cliente",
    "descrizione_articolo", "um", "quantita", "hu",
]

UOM_VARIANTS = r"(?:NR|KG|LT|MT|PZ|CF|SC|BL|ROT|PAL|N\b)"

FIELD_PATTERNS = {
    "numero_ddt":       r"NUMERO\s+(\d{7,12})\b",
    "data_ddt":         r"DATA\s+USCITA\s+MERCI\s+[\-–—]?\s*(\d{2}/\d{2}/\d{4})",
    "n_ordine_cliente": [
        r"N\.?\s*ORD\.?\s*ACQ\.?\s*CLIENTE\s+(\d+)",
        r"ORD(?:INE)?\s+CLIENTE\s+(\d+)",
        r"N\.?\s*ORD\.?\s*CLIENT[EI]\s*[:\.]?\s*(\d+)",
    ],
    "n_ordine_fornitore": r"NUMERO\s+ORDINE\s+(\d+)",
    "fornitore": r"([A-Z][A-Za-zÀ-ÿ\s\.\,\']+(?:S\.p\.A|S\.r\.l|SpA|Srl|S\.P\.A|SPA|SRL)[^\n]{0,30})",
}


# ── Gestione file processati ─────────────────────────────────────────────────
def load_processed(processed_file: Path) -> set:
    if processed_file.exists():
        return set(processed_file.read_text(encoding="utf-8").splitlines())
    return set()


def mark_processed(processed_file: Path, filename: str):
    with open(processed_file, "a", encoding="utf-8") as f:
        f.write(filename + "\n")


# ── OCR ──────────────────────────────────────────────────────────────────────
def ocr_pdf(pdf_path: Path) -> str:
    from pdf2image import convert_from_path
    import pytesseract

    print(f"  OCR: {pdf_path.name} ...", flush=True)
    pages = convert_from_path(str(pdf_path), dpi=300)
    full_text = ""
    for i, page in enumerate(pages):
        # Multipli PSM per massimizzare la cattura dei campi
        t1 = pytesseract.image_to_string(page, lang="ita", config="--psm 3")
        t2 = pytesseract.image_to_string(page, lang="ita", config="--psm 6")
        # Unisci i due testi (t1 di solito è più completo)
        full_text += t1 + "\n" + t2 + "\n"
    return full_text


# ── Estrazione campi intestazione ─────────────────────────────────────────────
def extract_field(text: str, pattern, default: str = "") -> str:
    if isinstance(pattern, list):
        for p in pattern:
            m = re.search(p, text, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return default
    m = re.search(pattern, text, re.IGNORECASE)
    return m.group(1).strip() if m else default


def clean_fornitore(raw: str) -> str:
    """Rimuove testo spurio dopo il nome azienda."""
    # Tronca a parole chiave che non fanno parte del nome
    stopwords = ["DATA", "SEDE", "TEL", "FAX", "CAP", "VIA", "P.IVA", "SOCIETA"]
    parts = raw.split()
    out = []
    for w in parts:
        if w.upper() in stopwords:
            break
        out.append(w)
    return " ".join(out).strip(" -,")


# ── Estrazione articoli ───────────────────────────────────────────────────────
def normalize_article_text(raw_table: str) -> str:
    """
    Corregge artefatti OCR frequenti nelle righe articolo:
    - "264 2\nR" → "264 22"  (HU e NR spezzati)
    - "NR 17622"  → "NR 176 22"  (QTY e HU senza spazio)
    """
    # Unisci riga che contiene solo "R" o "NR" come continuazione della riga precedente
    text = re.sub(r"(\d+)\s*\n+\s*R\b", r"\g<1>2", raw_table)
    text = re.sub(r"\n+R\b\s*\n", "\n", text)

    # Fix "NR 17622" → se dopo NR c'è un numero >3 cifre, prova a splittare
    def fix_merged_qty_hu(m):
        uom = m.group(1)
        digits = m.group(2)
        if len(digits) >= 4:
            # Prova split: ultime 2 cifre come HU se plausibili
            qty = digits[:-2]
            hu = digits[-2:]
            return f"{uom} {qty} {hu}"
        return m.group(0)

    text = re.sub(r"\b(" + UOM_VARIANTS + r")\s+(\d{4,})\b", fix_merged_qty_hu, text)
    return text


def parse_articles(text: str) -> list:
    """
    Estrae le righe articolo dalla sezione tabella del DDT.
    Formato atteso: codice  [cod_cliente]  descrizione  UOM  quantità  HU
    """
    # Trova inizio tabella
    m_start = re.search(
        r"CODICE\s+ARTICOLO.*?(?:QUANTITA|QUANTITÀ).*?HU",
        text, re.IGNORECASE | re.DOTALL
    )
    if not m_start:
        return []

    table_text = text[m_start.end():]

    # Trova fine tabella
    m_end = re.search(
        r"ORARIO\s+DI\s+SCARICO|TOTALE\s+(?:VOLUME|PESO)|DATA\s+CARICO|"
        r"DICHIARAZIONE\s+DI|VETTORE\b",
        table_text, re.IGNORECASE
    )
    if m_end:
        table_text = table_text[:m_end.start()]

    table_text = normalize_article_text(table_text)

    articles = []
    # Pattern principale: codice numerico + descrizione + UOM + quantità + HU opzionale
    art_re = re.compile(
        r"^(\d{1,10})\s+"                        # codice articolo
        r"((?:[A-Z0-9][^\n]{2,80}?)\s+)"         # descrizione
        r"(" + UOM_VARIANTS + r")\s+"             # unità di misura
        r"(\d+(?:[.,]\d+)?)"                      # quantità
        r"(?:\s+(\d+))?"                          # HU (opzionale)
        r"\s*$",
        re.MULTILINE | re.IGNORECASE
    )

    for m in art_re.finditer(table_text):
        descr = m.group(2).strip()
        # Salta righe che sono chiaramente header/noise
        if re.match(r"CODICE|DESCRIZIONE|ARTICOLO", descr, re.IGNORECASE):
            continue
        uom = m.group(3).upper()
        if uom == "N":
            uom = "NR"
        articles.append({
            "codice_articolo":         m.group(1),
            "codice_articolo_cliente": "",
            "descrizione_articolo":    descr,
            "um":                      uom,
            "quantita":                m.group(4).replace(",", "."),
            "hu":                      m.group(5) or "",
        })

    return articles


# ── Pipeline principale per un PDF ───────────────────────────────────────────
def process_pdf(pdf_path: Path) -> list:
    text = ocr_pdf(pdf_path)

    numero_ddt      = extract_field(text, FIELD_PATTERNS["numero_ddt"])
    data_ddt        = extract_field(text, FIELD_PATTERNS["data_ddt"])
    n_ord_cliente   = extract_field(text, FIELD_PATTERNS["n_ordine_cliente"])
    n_ord_fornitore = extract_field(text, FIELD_PATTERNS["n_ordine_fornitore"])
    fornitore_raw   = extract_field(text, FIELD_PATTERNS["fornitore"])
    fornitore       = clean_fornitore(fornitore_raw)

    articles = parse_articles(text)

    if not articles:
        # Nessun articolo trovato: inserisce comunque una riga con i dati header
        articles = [{
            "codice_articolo": "", "codice_articolo_cliente": "",
            "descrizione_articolo": "*** ARTICOLI NON RILEVATI - VERIFICARE MANUALMENTE ***",
            "um": "", "quantita": "", "hu": "",
        }]

    rows = []
    for art in articles:
        rows.append({
            "file_pdf":             pdf_path.name,
            "numero_ddt":           numero_ddt,
            "data_ddt":             data_ddt,
            "fornitore":            fornitore,
            "n_ordine_cliente":     n_ord_cliente,
            "n_ordine_fornitore":   n_ord_fornitore,
            **art,
        })
    return rows


# ── Scrittura CSV ─────────────────────────────────────────────────────────────
def write_csv(csv_path: Path, rows: list, write_header: bool):
    with open(csv_path, "a", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS, delimiter=";")
        if write_header:
            writer.writeheader()
        writer.writerows(rows)


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Estrae articoli da DDT in PDF e salva in CSV"
    )
    parser.add_argument(
        "cartella", nargs="?", default=".",
        help="Cartella contenente i PDF (default: cartella corrente)"
    )
    parser.add_argument(
        "--csv", default="ddt_articoli.csv",
        help="File CSV di output (default: ddt_articoli.csv)"
    )
    parser.add_argument(
        "--reset", action="store_true",
        help="Forza ri-elaborazione di tutti i file"
    )
    args = parser.parse_args()

    setup_tesseract()

    folder = Path(args.cartella).resolve()
    csv_path = folder / args.csv
    processed_file = folder / "ddt_processed.txt"

    if not folder.is_dir():
        print(f"ERRORE: cartella non trovata: {folder}")
        sys.exit(1)

    pdf_files = sorted(folder.glob("*.pdf"))
    if not pdf_files:
        print(f"Nessun file PDF trovato in: {folder}")
        sys.exit(0)

    if args.reset and processed_file.exists():
        processed_file.unlink()
        print("Reset: tutti i file saranno ri-elaborati.")

    processed = load_processed(processed_file)
    write_header = not csv_path.exists()

    pending = [p for p in pdf_files if p.name not in processed]
    print(f"PDF trovati: {len(pdf_files)}  |  Da elaborare: {len(pending)}  |  Già processati: {len(processed)}")

    if not pending:
        print("Nessun nuovo file da elaborare.")
        sys.exit(0)

    total_rows = 0
    errors = []

    for pdf_path in pending:
        print(f"\n[{pdf_path.name}]")
        try:
            rows = process_pdf(pdf_path)
            write_csv(csv_path, rows, write_header)
            write_header = False
            mark_processed(processed_file, pdf_path.name)
            total_rows += len(rows)
            print(f"  → {len(rows)} riga/e scritte nel CSV")
            for r in rows:
                print(f"     {r['codice_articolo']:10} {r['descrizione_articolo'][:45]:45} "
                      f"{r['um']:4} {r['quantita']:>8} {r['hu']:>4}")
        except Exception as e:
            err_msg = f"ERRORE su {pdf_path.name}: {e}"
            print(err_msg)
            traceback.print_exc()
            errors.append(err_msg)

    print(f"\n{'='*60}")
    print(f"Elaborazione completata.")
    print(f"  CSV:       {csv_path}")
    print(f"  Righe:     {total_rows}")
    print(f"  Errori:    {len(errors)}")
    if errors:
        for e in errors:
            print(f"  - {e}")

    input("\nPremi INVIO per chiudere...")


if __name__ == "__main__":
    main()
