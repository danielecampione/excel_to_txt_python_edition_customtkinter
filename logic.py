"""
logic.py
--------
Logica di elaborazione: lettura Excel e scrittura file di testo.
"""

import os
import pandas as pd


def get_first_non_bkp_sheet(xls: pd.ExcelFile) -> str:
    for name in xls.sheet_names:
        if not name.lower().endswith("bkp"):
            return name
    raise ValueError("Tutti i fogli hanno il suffisso 'bkp': nessun foglio utilizzabile.")


def process_excel(path: str, columns: list[str]) -> str:
    """
    Legge il primo foglio non-bkp e restituisce il testo formattato.

    :param path:    percorso del file Excel
    :param columns: nomi delle colonne da estrarre, nell'ordine desiderato.
                    I valori multi-riga vengono automaticamente racchiusi tra virgolette.
    """
    path = path.strip()
    xls = pd.ExcelFile(path)
    sheet_name = get_first_non_bkp_sheet(xls)

    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(
            f"Colonne mancanti nel foglio '{sheet_name}': {', '.join(missing)}"
        )

    blocks = []
    for _, row in df.iterrows():
        lines = []
        for col in columns:
            v = row[col]
            value = str(v).strip() if pd.notna(v) else ""
            # Se il valore contiene a-capo lo racchiude tra virgolette
            if "\n" in value:
                value = f'"{value}"'
            lines.append(value)
        blocks.append("\n\n".join(lines))

    return "\n\n---\n\n".join(blocks) + "\n\n---"


def save_output(excel_path: str, content: str) -> str:
    base = os.path.splitext(excel_path)[0]
    out_path = base + "_output.txt"

    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)
    except OSError:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        out_path = os.path.join(script_dir, os.path.basename(base) + "_output.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)

    return out_path
