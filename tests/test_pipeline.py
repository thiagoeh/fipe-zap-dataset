import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from fipezap_sqlite.pipeline import ALL_TIPOS, BODY_ROW, SOURCE_URL, build, extract

DATES = [datetime(2008, 1, 1), datetime(2008, 2, 1), datetime(2008, 3, 1)]
ND = "não disponível"
SP_ROWS = [
    [61.59, 57.57, ".", 2610.3, ".", 0.0, 71.02, 999.0],
    [62.08, 58.43, 0.008, 2631.13, 72.8, 0.025, 72.8, 999.0],
    [63.0, 59.0, 0.01, 2650.0, ".", ".", ".", 999.0],
]
DIADEMA_ROWS = [
    [150.28, 148.0, ".", 6973.8, ND, ND, ND, ND],
    [151.1, 149.2, 0.005, 6980.4, ND, ND, ND, ND],
]
CUIABA_ROWS = [
    [200.0, 198.0, ".", 5000.0, ND, ND, ND, ND],
    [201.0, 199.0, 0.005, 5010.0, ND, ND, ND, ND],
    [202.0, 200.0, 0.005, 5020.0, ND, ND, ND, ND],
]


def _write_sheet(ws: Worksheet, name: str, rows: list[list[object]], labels: bool = True) -> None:
    sale_anchor = "Venda" if labels else "Preços de venda"
    rent_anchor = "Locação" if labels else "Preços de locação"
    ws.cell(row=1, column=2, value=name)
    ws.cell(row=1, column=3, value="Imóveis residenciais")
    ws.cell(row=1, column=12, value="Imóveis comerciais")
    ws.cell(row=2, column=3, value=sale_anchor)
    ws.cell(row=2, column=7, value=rent_anchor)
    ws.cell(row=2, column=10, value="Rentabilidade do aluguel (rental yield)")
    ws.cell(row=2, column=12, value="Venda")
    ws.cell(row=3, column=12, value="Número-Índice")
    ws.cell(row=4, column=12, value="Total")
    for col, label in [
        (3, "Número-Índice"),
        (5, "Var. mensal (%)"),
        (6, "Preço médio (R$/m²)"),
        (7, "Número-Índice"),
        (8, "Var. mensal (%)"),
        (9, "Preço médio (R$/m²)"),
    ]:
        ws.cell(row=3, column=col, value=label)
    ws.cell(row=4, column=2, value="Data")
    for col in range(3, 11):
        if col == 4:
            ws.cell(row=4, column=col, value="1D")
        else:
            ws.cell(row=4, column=col, value="Total")
    for i, values in enumerate(rows):
        ws.cell(row=BODY_ROW + i, column=2, value=DATES[i])
        for j, value in enumerate(values):
            ws.cell(row=BODY_ROW + i, column=3 + j, value=value)
        ws.cell(row=BODY_ROW + i, column=12, value=123.0)


@pytest.fixture()
def xlsx(tmp_path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "São Paulo"
    _write_sheet(ws, "São Paulo", SP_ROWS)
    _write_sheet(wb.create_sheet("Diadema"), "Diadema", DIADEMA_ROWS)
    _write_sheet(wb.create_sheet("Cuiabá"), "Cuiabá", CUIABA_ROWS, labels=False)
    path = tmp_path / "fipezap.xlsx"
    wb.save(path)
    return path


def _index(
    xlsx_path: Path, tipos: frozenset[str] = ALL_TIPOS
) -> dict[tuple[str, str, str, str, str], float | None]:
    return {(r[0], r[1], r[2], r[3], r[4]): r[5] for r in extract(xlsx_path, tipos)}


def test_extract_counts(xlsx: Path) -> None:
    assert len(extract(xlsx)) == 56
    assert len(extract(xlsx, frozenset({"venda"}))) == 32


def test_nulls_and_values(xlsx: Path) -> None:
    records = _index(xlsx)
    assert records[("São Paulo", "2008-01-01", "venda", "1d", "indice")] == 57.57
    assert records[("São Paulo", "2008-01-01", "venda", "total", "variacao_mensal")] is None
    assert records[("Diadema", "2008-02-01", "locacao", "total", "preco_m2")] is None


def test_skipped_blocks(xlsx: Path) -> None:
    records = extract(xlsx)
    assert {r[4] for r in records} == {"indice", "variacao_mensal", "preco_m2"}
    assert all(r[5] != 123.0 for r in records)


def test_critical_question(tmp_path: Path, xlsx: Path) -> None:
    db = tmp_path / "fipezap.sqlite3"
    build(xlsx, db)
    con = sqlite3.connect(db)
    latest = con.execute(
        "SELECT data, tipo, valor FROM preco_m2 "
        "WHERE cidade = ? AND categoria = 'total' ORDER BY data DESC",
        ("São Paulo",),
    ).fetchall()
    assert latest[0] == ("2008-03-01", "venda", 2650.0)
    assert latest[1] == ("2008-03-01", "locacao", None)
    atual = dict(
        con.execute(
            "SELECT tipo, valor FROM preco_m2_atual WHERE cidade = 'São Paulo'",
        ).fetchall()
    )
    assert atual == {"venda": 2650.0, "locacao": 72.8}
    series = con.execute(
        "SELECT data, valor FROM preco_m2 WHERE cidade = 'São Paulo' AND tipo = 'venda' "
        "AND categoria = 'total' ORDER BY data",
    ).fetchall()
    assert series == [("2008-01-01", 2610.3), ("2008-02-01", 2631.13), ("2008-03-01", 2650.0)]


def test_atual_per_city(tmp_path: Path, xlsx: Path) -> None:
    db = tmp_path / "fipezap.sqlite3"
    build(xlsx, db)
    con = sqlite3.connect(db)
    assert con.execute(
        "SELECT DISTINCT data FROM atual WHERE cidade = 'São Paulo'",
    ).fetchall() == [("2008-03-01",)]
    assert con.execute(
        "SELECT DISTINCT data FROM atual WHERE cidade = 'Diadema'",
    ).fetchall() == [("2008-02-01",)]
    sp_latest = con.execute(
        "SELECT tipo, categoria, metrica, valor FROM atual "
        "WHERE cidade = 'São Paulo' AND metrica = 'preco_m2' AND categoria = 'total'",
    ).fetchall()
    assert set(sp_latest) == {
        ("venda", "total", "preco_m2", 2650.0),
        ("locacao", "total", "preco_m2", None),
    }


def test_metadata(tmp_path: Path, xlsx: Path) -> None:
    db = tmp_path / "fipezap.sqlite3"
    build(xlsx, db)
    con = sqlite3.connect(db)
    url, sha, _ = con.execute(
        "SELECT fonte_url, fonte_sha256, extraido_em FROM metadados"
    ).fetchone()
    assert url == SOURCE_URL
    assert sha == hashlib.sha256(xlsx.read_bytes()).hexdigest()


def test_constraints(tmp_path: Path, xlsx: Path) -> None:
    db = tmp_path / "fipezap.sqlite3"
    build(xlsx, db)
    con = sqlite3.connect(db)
    bad_inserts = [
        "INSERT INTO fipezap VALUES ('Diadema', '2008-01-01', 'venda', 'total', 'indice', 1.0)",
        "INSERT INTO fipezap VALUES ('X', '2008-01-01', 'aluguel', 'total', 'indice', 1.0)",
        "INSERT INTO fipezap VALUES ('X', '2008-01-01', 'venda', 'total', 'indice', 'oops')",
        "INSERT INTO fipezap VALUES ('X', 'nope', 'venda', 'total', 'indice', 1.0)",
        "INSERT INTO fipezap VALUES ('X', '2008-13-01', 'venda', 'total', 'indice', 1.0)",
    ]
    for sql in bad_inserts:
        with pytest.raises(sqlite3.IntegrityError):
            con.execute(sql)
