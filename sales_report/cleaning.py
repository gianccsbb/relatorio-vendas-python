"""Load messy sales exports from different sources and turn them into one clean table.

Every change made to the data is recorded in a log, so the owner can see exactly
what was fixed and why (nothing is silently altered or dropped).
"""
from __future__ import annotations

import csv
import difflib
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, time
from pathlib import Path

import pandas as pd

from .catalog import ALIASES, CATALOG, PAYMENT_ALIASES

UNKNOWN_PREFIX = "Não identificado: "

# Header names used by each source, mapped to our standard column names.
COLUMN_ALIASES = {
    "data": "date", "date": "date", "data do pedido": "datetime",
    "hora": "time", "time": "time",
    "produto": "product", "item": "product",
    "qtd": "qty", "quantidade": "qty", "quantity": "qty",
    "valor unit.": "unit_price", "valor unit": "unit_price", "unit_price": "unit_price", "preco": "unit_price",
    "total": "line_total",
    "pagamento": "payment", "payment": "payment", "forma de pagamento": "payment",
    "sale_id": "sale_id", "no venda": "sale_id", "n venda": "sale_id", "id venda": "sale_id",
}


@dataclass
class CleaningLog:
    entries: list[dict] = field(default_factory=list)

    def add(self, source: str, row: int | str, field_name: str, original, fixed, action: str):
        self.entries.append({
            "Fonte": source, "Linha": row, "Campo": field_name,
            "Valor original": "" if original is None else str(original),
            "Valor corrigido": "" if fixed is None else str(fixed),
            "Ação": action,
        })

    def to_frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.entries, columns=["Fonte", "Linha", "Campo", "Valor original", "Valor corrigido", "Ação"])


# ---------------------------------------------------------------- parsers

def normalize_text(value: str) -> str:
    """Lowercase, remove accents and extra spaces: 'CAFÉ  Coado ' -> 'cafe coado'."""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", text).strip().lower()


def parse_price(value) -> int:
    """Return a price in cents from 'R$ 6,00', '6.00', '6', '1.234,50' or a number."""
    if isinstance(value, (int, float)) and not pd.isna(value):
        return round(float(value) * 100)
    text = str(value).strip().replace("R$", "").replace(" ", "").replace(" ", "")
    if not text:
        raise ValueError("empty price")
    if "," in text:  # Brazilian format: thousands with '.', decimals with ','
        text = text.replace(".", "").replace(",", ".")
    return round(float(text) * 100)


DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y")


def parse_date(value) -> date:
    """Accept dd/mm/yyyy, yyyy-mm-dd, dd-mm-yyyy, dd/mm/yy, or datetime objects."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"unrecognized date: {value!r}")


def parse_time(value) -> time:
    if isinstance(value, time):
        return value
    text = str(value).strip()
    for fmt in ("%H:%M:%S", "%H:%M", "%Hh%M", "%Hh"):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            continue
    raise ValueError(f"unrecognized time: {value!r}")


_CATALOG_KEYS = {normalize_text(name): name for name in CATALOG}


def normalize_product(value) -> tuple[str, str]:
    """Return (canonical name, how it was matched)."""
    key = normalize_text(value)
    if key in _CATALOG_KEYS:
        return _CATALOG_KEYS[key], "exact"
    if key in ALIASES:
        return ALIASES[key], "alias"
    match = difflib.get_close_matches(key, list(_CATALOG_KEYS) + list(ALIASES), n=1, cutoff=0.82)
    if match:
        hit = match[0]
        return _CATALOG_KEYS.get(hit) or ALIASES[hit], "fuzzy"
    return UNKNOWN_PREFIX + str(value).strip(), "unknown"


def normalize_payment(value) -> str:
    key = normalize_text(value)
    return PAYMENT_ALIASES.get(key, str(value).strip().title() or "Não informado")


# ---------------------------------------------------------------- loading

def _read_rows(path: Path) -> list[dict]:
    """Read a CSV (any delimiter) or Excel file into a list of dicts with standard column names."""
    if path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.read_excel(path, dtype=object)
        rows = df.to_dict("records")
        header = list(df.columns)
    else:
        raw = path.read_text(encoding="utf-8-sig")
        dialect = csv.Sniffer().sniff(raw.splitlines()[0], delimiters=";,\t|")
        reader = csv.reader(raw.splitlines(), dialect)
        lines = list(reader)
        header, data = lines[0], lines[1:]
        rows = [dict(zip(header, line)) for line in data]
    mapping = {col: COLUMN_ALIASES.get(normalize_text(col)) for col in header}
    out = []
    for i, row in enumerate(rows, start=2):  # row 1 is the header
        std = {mapping[k]: v for k, v in row.items() if mapping.get(k)}
        std["_row"] = i
        std["_header_values"] = [normalize_text(v) for v in row.values() if v is not None]
        std["_header_names"] = [normalize_text(c) for c in header]
        out.append(std)
    return out


def load_source(path: str | Path, log: CleaningLog) -> pd.DataFrame:
    path = Path(path)
    source = path.name
    records = []
    for row in _read_rows(path):
        n = row["_row"]
        values = [v for v in row.values() if isinstance(v, str)]
        # Blank lines exported by the POS
        if all((v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "")
               for k, v in row.items() if not k.startswith("_")):
            log.add(source, n, "linha", "(vazia)", None, "Linha vazia removida")
            continue
        # Header repeated in the middle of the file (export glitch)
        if row["_header_values"] == row["_header_names"]:
            log.add(source, n, "linha", "(cabeçalho repetido)", None, "Cabeçalho duplicado removido")
            continue
        try:
            if "datetime" in row:
                dt_val = row["datetime"]
                dt = dt_val if isinstance(dt_val, datetime) else pd.to_datetime(dt_val, dayfirst=True).to_pydatetime()
            else:
                dt = datetime.combine(parse_date(row["date"]), parse_time(row["time"]))
            qty = int(float(str(row["qty"]).replace(",", ".")))
            if "unit_price" in row:
                unit = parse_price(row["unit_price"])
            else:
                total = parse_price(row["line_total"])
                unit = round(total / qty) if qty else 0
                log.add(source, n, "preço unitário", f"total {row['line_total']} / qtd {qty}", f"{unit/100:.2f}",
                        "Preço unitário calculado a partir do total")
        except (ValueError, KeyError, TypeError) as exc:
            log.add(source, n, "linha", " | ".join(map(str, values))[:120], None, f"Linha ignorada: {exc}")
            continue

        product, how = normalize_product(row.get("product", ""))
        if how in ("alias", "fuzzy"):
            log.add(source, n, "produto", row.get("product"), product, "Nome padronizado" + (" (aproximação)" if how == "fuzzy" else ""))
        elif how == "unknown":
            log.add(source, n, "produto", row.get("product"), product, "Produto não encontrado no catálogo, mantido para revisão")
        elif str(row.get("product", "")).strip() != product:
            log.add(source, n, "produto", row.get("product"), product, "Maiúsculas/acentos/espaços corrigidos")

        payment = normalize_payment(row.get("payment", "App de entrega" if "datetime" in row else ""))
        if qty < 0:
            log.add(source, n, "quantidade", qty, qty, "Estorno mantido como valor negativo")

        records.append({
            "Data e hora": dt, "Produto": product,
            "Categoria": CATALOG.get(product, ("Outros", 0))[0],
            "Quantidade": qty, "Preço unitário": unit / 100,
            "Pagamento": payment,
            "Canal": "Delivery" if "datetime" in row else "Balcão",
            "Fonte": source, "_row": n, "_sale_id": str(row.get("sale_id") or ""),
        })
    return pd.DataFrame.from_records(records)


def clean(paths: list[str | Path]) -> tuple[pd.DataFrame, CleaningLog]:
    """Load every export, standardize it, remove duplicates and return (sales, log)."""
    log = CleaningLog()
    frames = [load_source(p, log) for p in paths]
    frames = [f for f in frames if not f.empty]
    if not frames:
        raise ValueError("no valid sales rows found in the input files")
    df = pd.concat(frames, ignore_index=True)

    # Duplicates inside the same file: the POS exports some sales twice.
    # When the export has a sale ID, that decides; otherwise every field must match.
    has_id = df["_sale_id"] != ""
    full_key = ["Fonte", "Data e hora", "Produto", "Quantidade", "Preço unitário", "Pagamento"]
    dup_mask = (has_id & df.duplicated(subset=["Fonte", "_sale_id"], keep="first")) | \
               (~has_id & df.duplicated(subset=full_key, keep="first"))
    for _, r in df[dup_mask].iterrows():
        log.add(r["Fonte"], r["_row"], "linha", f"{r['Data e hora']:%d/%m %H:%M} {r['Produto']}", None, "Venda duplicada removida")
    df = df[~dup_mask].copy()

    df["Total"] = (df["Quantidade"] * df["Preço unitário"]).round(2)
    df["Data"] = df["Data e hora"].dt.date
    df["Hora"] = df["Data e hora"].dt.hour
    dias = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta", "Sábado", "Domingo"]
    df["Dia da semana"] = df["Data e hora"].dt.weekday.map(lambda i: dias[i])
    df = df.sort_values("Data e hora", kind="stable").drop(columns=["_row", "_sale_id"]).reset_index(drop=True)
    cols = ["Data", "Hora", "Dia da semana", "Data e hora", "Produto", "Categoria", "Quantidade",
            "Preço unitário", "Total", "Pagamento", "Canal", "Fonte"]
    return df[cols], log
