"""Logica Excel -> TXT con join opzionale e fallback controllato."""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from typing import Iterable

import pandas as pd

_KEY_SEPARATOR = "\x1f"
_warnings: list[str] = []


def _clean(value) -> str:
    return "" if pd.isna(value) else str(value).strip()


def _normalize(value) -> str:
    return re.sub(r"\s+", " ", _clean(value)).strip().casefold()


def _read_first_sheet(path: str, label: str) -> pd.DataFrame:
    path = path.strip()
    if not path:
        raise ValueError(f"Percorso del {label} non valido.")
    try:
        df = pd.read_excel(path, sheet_name=0, dtype=str, keep_default_na=False)
    except (ValueError, OSError, ImportError) as exc:
        raise ValueError(f"Impossibile leggere il primo foglio del {label}: {exc}") from exc
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _require_columns(df: pd.DataFrame, columns: Iterable[str], label: str) -> None:
    missing: list[str] = []
    for col in columns:
        if col not in df.columns and col not in missing:
            missing.append(col)
    if missing:
        raise ValueError(f"Colonne mancanti nel primo foglio del {label}: {missing}")


def _key(row: pd.Series, columns: list[str]) -> str:
    parts = [_normalize(row[col]) for col in columns]
    return "" if any(not part for part in parts) else _KEY_SEPARATOR.join(parts)


@dataclass
class SecondaryIndexes:
    exact: dict[str, list[str]]
    fallback: dict[str, list[str]]
    ambiguous: set[str]
    fallback_enabled: bool


def _build_indexes(path: str, keys: list[str], outputs: list[str], fallback: bool) -> SecondaryIndexes:
    secondary = _read_first_sheet(path, "documento secondario")
    _require_columns(secondary, [*keys, *outputs], "documento secondario")
    enabled = fallback and len(keys) >= 3
    indexes = SecondaryIndexes({}, {}, set(), enabled)
    source_rows: dict[str, int] = {}

    for pos, (_, row) in enumerate(secondary.iterrows(), start=2):
        values = [_clean(row[col]) for col in outputs]
        exact_key = _key(row, keys)
        if exact_key:
            if exact_key in indexes.exact:
                _warnings.append(
                    f"Documento secondario, riga {pos} ignorata: chiave completa duplicata "
                    f"della riga {source_rows[exact_key]}. Mantenuto il primo record dall'alto."
                )
            else:
                indexes.exact[exact_key] = values
                source_rows[exact_key] = pos

        if enabled:
            fallback_key = _key(row, keys[:-1])
            if fallback_key:
                if fallback_key in indexes.fallback:
                    indexes.ambiguous.add(fallback_key)
                    indexes.fallback.pop(fallback_key, None)
                elif fallback_key not in indexes.ambiguous:
                    indexes.fallback[fallback_key] = values
    return indexes


def _validate_join(path, main_keys, secondary_keys, outputs) -> bool:
    join = bool(path and path.strip())
    if not join:
        return False
    if not main_keys or not secondary_keys:
        raise ValueError("Per la join indica almeno una colonna chiave per entrambi i documenti.")
    if len(main_keys) != len(secondary_keys):
        raise ValueError("Le chiavi di join devono essere in coppia: stesso numero di righe nei due documenti.")
    if not outputs:
        raise ValueError("Indica almeno una colonna del documento secondario da aggiungere all'output.")
    return True


def process_excel(path: str, columns: list[str], secondary_path: str | None = None,
                  main_keys: list[str] | None = None, secondary_keys: list[str] | None = None,
                  secondary_output: list[str] | None = None, fallback: bool = False) -> str:
    """Converte il primo foglio e, se richiesto, applica una join con il file secondario."""
    _warnings.clear()
    main_keys = main_keys or []
    secondary_keys = secondary_keys or []
    secondary_output = secondary_output or []
    join = _validate_join(secondary_path, main_keys, secondary_keys, secondary_output)
    indexes = _build_indexes(secondary_path.strip(), secondary_keys, secondary_output, fallback) if join else None

    main = _read_first_sheet(path, "documento principale")
    _require_columns(main, [*columns, *(main_keys if join else [])], "documento principale")

    blocks: list[str] = []
    for pos, (_, row) in enumerate(main.iterrows(), start=2):
        values = [_clean(row[col]) for col in columns]
        if not any(values):
            continue

        joined: list[str] = []
        note: str | None = None
        if join and indexes is not None:
            exact_key = _key(row, main_keys)
            match = indexes.exact.get(exact_key) if exact_key else None
            if match is None and indexes.fallback_enabled:
                fallback_key = _key(row, main_keys[:-1])
                if fallback_key in indexes.ambiguous:
                    _warnings.append(
                        f"Documento principale, riga {pos}: fallback non applicato perche le chiavi residue "
                        "non identificano un solo record nel documento secondario."
                    )
                elif fallback_key and fallback_key in indexes.fallback:
                    match = indexes.fallback[fallback_key]
                    ignored = main_keys[-1]
                    _warnings.append(
                        f"Documento principale, riga {pos}: join risolta con fallback ignorando la chiave '{ignored}'."
                    )
                    note = ("[NOTA JOIN] Dati ottenuti tramite fallback: corrispondenza trovata senza "
                            f"considerare la chiave '{ignored}'.")
            if match is None:
                joined = [""] * len(secondary_output)
            else:
                if note:
                    joined.append(note)
                joined.extend(match)

        if join:
            block = "\n\n".join(values) + "\n\n" + "\n".join(joined) + "\n"
        else:
            block = "\n\n".join(values) + "\n"
        blocks.append(block)

    return "".join(block + "\n---\n" + ("\n" if i < len(blocks) - 1 else "")
                   for i, block in enumerate(blocks))


def consume_warnings() -> str:
    if not _warnings:
        return ""
    shown = _warnings[:8]
    text = f"Elaborazione completata con {len(_warnings)} warning." + "".join(f"\n- {w}" for w in shown)
    if len(_warnings) > 8:
        text += f"\n- ... e altri {len(_warnings) - 8} warning."
    _warnings.clear()
    return text


def save_output(excel_path: str, content: str) -> str:
    base = os.path.splitext(excel_path)[0]
    out_path = base + "_output.txt"
    try:
        with open(out_path, "w", encoding="utf-8", newline="") as file:
            file.write(content)
    except OSError:
        out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), os.path.basename(base) + "_output.txt")
        with open(out_path, "w", encoding="utf-8", newline="") as file:
            file.write(content)
    return os.path.abspath(out_path)
