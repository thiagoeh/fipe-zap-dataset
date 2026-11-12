import tempfile
from enum import StrEnum
from pathlib import Path
from typing import Annotated

import typer

from fipezap_sqlite import pipeline


class Tipo(StrEnum):
    VENDA = "venda"
    LOCACAO = "locacao"
    AMBOS = "ambos"


TIPOS = {
    Tipo.VENDA: frozenset({"venda"}),
    Tipo.LOCACAO: frozenset({"locacao"}),
    Tipo.AMBOS: pipeline.ALL_TIPOS,
}

app = typer.Typer(no_args_is_help=True)


@app.command()
def download(
    output: Annotated[
        Path,
        typer.Option(help="Where to save the XLSX"),
    ] = Path("fipezap-serieshistoricas.xlsx"),
) -> None:
    """Download the FIPE ZAP historical series XLSX."""
    pipeline.download(output)


@app.command()
def build(
    input_path: Annotated[Path | None, typer.Option("--input", help="Local XLSX path")] = None,
    download: Annotated[bool, typer.Option(help="Download the XLSX from FIPE first")] = False,
    output: Annotated[
        Path,
        typer.Option(help="Output SQLite database path"),
    ] = Path("fipezap.sqlite3"),
    tipo: Annotated[Tipo, typer.Option(help="Which series to load")] = Tipo.AMBOS,
) -> None:
    """Build the SQLite database from a local XLSX or a fresh download."""
    if input_path is None and not download:
        raise typer.BadParameter("provide --input or --download")
    if download:
        with tempfile.TemporaryDirectory() as tmp:
            xlsx = Path(tmp) / "fipezap-serieshistoricas.xlsx"
            pipeline.download(xlsx)
            pipeline.build(xlsx, output, TIPOS[tipo])
    else:
        pipeline.build(input_path, output, TIPOS[tipo])
