import hashlib
import os
import sqlite3
import urllib.request
from datetime import UTC, date, datetime
from pathlib import Path

from loguru import logger
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

SOURCE_URL = "https://downloads.fipe.org.br/indices/fipezap/fipezap-serieshistoricas.xlsx"
SKIP_SHEETS = frozenset({"Resumo", "Aux"})
AGGREGATE_CITIES = {"Índice FipeZAP": "Brasil"}
KINDS = (
    ("venda", "venda"),
    ("locação", "locacao"),
    ("preços de venda", "venda"),
    ("preços de locação", "locacao"),
    ("rentabilidade do aluguel", "locacao"),
)
METRICS = {
    "número-índice": "indice",
    "var. mensal": "variacao_mensal",
    "var. em 12 meses": "variacao_12m",
    "preço médio": "preco_m2",
    "(% - mensalizada": "rentabilidade_aluguel",
}
MISSING = frozenset({"", ".", "não disponível"})
DATA_COL = 1
BODY_ROW = 5
ALL_TIPOS = frozenset({"venda", "locacao"})

SCHEMA = """
CREATE TABLE fipezap (
    cidade TEXT NOT NULL,
    segmento TEXT NOT NULL CHECK (segmento IN ('residencial', 'comercial')),
    data TEXT NOT NULL CHECK (
        data GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
        AND data = strftime('%Y-%m-%d', data)
        AND CAST(substr(data, 6, 2) AS INTEGER) BETWEEN 1 AND 12
        AND CAST(substr(data, 9, 2) AS INTEGER) BETWEEN 1 AND 31
    ),
    tipo TEXT NOT NULL CHECK (tipo IN ('venda', 'locacao')),
    categoria TEXT NOT NULL CHECK (categoria IN ('total', '1d', '2d', '3d', '4d')),
    metrica TEXT NOT NULL CHECK (
        metrica IN (
            'indice', 'variacao_mensal', 'variacao_12m', 'preco_m2', 'rentabilidade_aluguel'
        )
    ),
    valor REAL,
    PRIMARY KEY (cidade, segmento, metrica, categoria, data, tipo)
) STRICT, WITHOUT ROWID;

CREATE VIEW preco_m2 AS
    SELECT cidade, segmento, data, tipo, categoria, valor
    FROM fipezap
    WHERE metrica = 'preco_m2';

CREATE VIEW atual AS
    SELECT f.cidade, f.segmento, f.data, f.tipo, f.categoria, f.metrica, f.valor
    FROM fipezap f
    WHERE f.data = (
        SELECT max(g.data) FROM fipezap g
        WHERE g.cidade = f.cidade AND g.segmento = f.segmento
    );

CREATE VIEW preco_m2_atual AS
    SELECT f.cidade, f.segmento, f.data, f.tipo, f.valor
    FROM fipezap f
    WHERE f.metrica = 'preco_m2'
      AND f.categoria = 'total'
      AND f.valor IS NOT NULL
      AND f.data = (
          SELECT max(g.data) FROM fipezap g
          WHERE g.metrica = 'preco_m2'
            AND g.categoria = 'total'
            AND g.valor IS NOT NULL
            AND g.cidade = f.cidade
            AND g.segmento = f.segmento
            AND g.tipo = f.tipo
      );

CREATE TABLE metadados (
    fonte_url TEXT NOT NULL,
    fonte_sha256 TEXT NOT NULL,
    extraido_em TEXT NOT NULL
) STRICT;
"""


def _normalize(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).casefold().split())


def _kinda(label: str) -> str | None:
    for prefix, kind in KINDS:
        if label.startswith(prefix):
            return kind
    return None


def _metrica(label: str) -> str:
    for prefix, name in METRICS.items():
        if label.startswith(prefix):
            return name
    msg = f"unrecognized metric label: {label!r}"
    raise ValueError(msg)


def _segments(row: tuple[object, ...], col: int, current: str | None) -> str | None:
    label = _normalize(row[col] if col < len(row) else None)
    if label == "imóveis residenciais":
        return "residencial"
    if label == "imóveis comerciais":
        return "comercial"
    return current if not label else None


def _columns(
    ws: Worksheet, tipos: frozenset[str]
) -> list[tuple[int, str, str, str, str]]:
    header = list(ws.iter_rows(min_row=1, max_row=4, values_only=True))
    segment_row, kind_row, metric_row, category_row = header
    if _normalize(category_row[DATA_COL]) != "data":
        msg = f"sheet {ws.title!r}: missing 'Data' header, layout changed"
        raise ValueError(msg)
    segment: str | None = None
    kind: str | None = None
    metrica: str | None = None
    skipped: set[str] = set()
    mapping: list[tuple[int, str, str, str, str]] = []
    for col, category in enumerate(category_row):
        segment = _segments(segment_row, col, segment)
        block = _normalize(kind_row[col] if col < len(kind_row) else None)
        if block:
            kind = _kinda(block)
            if kind is None:
                skipped.add(block)
        label = _normalize(metric_row[col] if col < len(metric_row) else None)
        if label and kind is not None:
            metrica = _metrica(label)
        if col == DATA_COL or kind is None or metrica is None or kind not in tipos:
            continue
        categoria = _normalize(category)
        if categoria and segment is not None:
            mapping.append((col, segment, kind, categoria, metrica))
    if not mapping:
        msg = f"sheet {ws.title!r}: no data columns"
        raise ValueError(msg)
    for block in sorted(skipped):
        logger.debug("sheet {}: skipped block {!r}", ws.title, block)
    return mapping


def _valor(value: object) -> float | None:
    if value is None:
        return None
    if isinstance(value, int | float):
        return float(value)
    text = str(value).strip()
    if text.casefold() in MISSING:
        return None
    try:
        return float(text.replace(",", "."))
    except ValueError:
        logger.warning("dropped unparseable value: {!r}", value)
        return None


def _data(value: object) -> str | None:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return None


def _records(
    city: str,
    ws: Worksheet,
    columns: list[tuple[int, str, str, str, str]],
) -> list[tuple[str, str, str, str, str, str, float | None]]:
    records = []
    for row in ws.iter_rows(min_row=BODY_ROW, values_only=True):
        day = _data(row[DATA_COL] if len(row) > DATA_COL else None)
        if day is None:
            break
        for col, segmento, tipo, categoria, metrica in columns:
            value = _valor(row[col] if len(row) > col else None)
            records.append((city, segmento, day, tipo, categoria, metrica, value))
    return records


def extract(
    xlsx: Path,
    tipos: frozenset[str] = ALL_TIPOS,
) -> list[tuple[str, str, str, str, str, str, float | None]]:
    wb = load_workbook(xlsx, read_only=True, data_only=True)
    try:
        records: list[tuple[str, str, str, str, str, str, float | None]] = []
        for name in wb.sheetnames:
            if name in SKIP_SHEETS:
                continue
            ws = wb[name]
            rows = _records(AGGREGATE_CITIES.get(name, name), ws, _columns(ws, tipos))
            logger.info("{}: {} rows", name, len(rows))
            records.extend(rows)
    finally:
        wb.close()
    logger.info("total {} rows from {}", len(records), xlsx.name)
    return records


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def download(dest: Path) -> Path:
    logger.info("downloading {}", SOURCE_URL)
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_name(dest.name + ".tmp")
    req = urllib.request.Request(SOURCE_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp, tmp.open("wb") as f:
        f.write(resp.read())
    os.replace(tmp, dest)
    logger.info("saved {} ({} bytes)", dest, dest.stat().st_size)
    return dest


def build(xlsx: Path, output: Path, tipos: frozenset[str] = ALL_TIPOS) -> None:
    records = extract(xlsx, tipos)
    tmp = output.with_name(output.name + ".tmp")
    tmp.unlink(missing_ok=True)
    conn = sqlite3.connect(tmp)
    try:
        with conn:
            conn.executescript(SCHEMA)
            conn.executemany("INSERT INTO fipezap VALUES (?, ?, ?, ?, ?, ?, ?)", records)
            conn.execute(
                "INSERT INTO metadados VALUES (?, ?, ?)",
                (SOURCE_URL, _sha256(xlsx), datetime.now(UTC).isoformat(timespec="seconds")),
            )
    finally:
        conn.close()
    os.replace(tmp, output)
    logger.info("wrote {} ({} rows)", output, len(records))
