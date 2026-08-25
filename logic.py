"""
logic.py
--------
Logica di elaborazione: lettura del primo foglio Excel e scrittura file di testo.
"""

import os

import pandas as pd


def process_excel(path: str, columns_to_extract: list) -> str:
    path = path.strip()

    # Il primo foglio da sinistra e sempre quello da convertire.
    df = pd.read_excel(path, sheet_name=0, dtype=str)

    # Pulizia nomi colonne
    df.columns = [c.strip() for c in df.columns]

    # Verifica che le colonne richieste dalla textarea esistano nel primo foglio
    missing = set(columns_to_extract) - set(df.columns)
    if missing:
        raise ValueError(f"Colonne mancanti nel primo foglio: {missing}")

    blocks = []
    for _, row in df.iterrows():
        def cell(col):
            value = row[col]
            return str(value).strip() if pd.notna(value) else ""

        # Estraiamo dinamicamente i valori per ogni colonna richiesta
        parts = [cell(col) for col in columns_to_extract]

        # Uniamo le parti del record con il doppio a capo
        block = "\n\n".join(parts) + "\n"
        blocks.append(block)

    # Uniamo i record con il separatore
    return "\n---\n\n".join(blocks) + "\n---\n"


def save_output(excel_path: str, content: str) -> str:
    base = os.path.splitext(excel_path)[0]
    out_path = base + "_output.txt"

    try:
        with open(out_path, "w", encoding="utf-8") as file:
            file.write(content)
    except OSError:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        out_path = os.path.join(
            script_dir,
            os.path.basename(base) + "_output.txt",
        )
        with open(out_path, "w", encoding="utf-8") as file:
            file.write(content)

    return out_path
