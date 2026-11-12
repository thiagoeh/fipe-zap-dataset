# fipe-zap-dataset

Converts the [FIPE ZAP](https://downloads.fipe.org.br/indices/fipezap/fipezap-serieshistoricas.xlsx)
residential price index XLSX into a SQLite database.

Requires [uv](https://docs.astral.sh/uv/).

## Run (uvx)

```
uvx --from git+https://github.com/thiagoeh/fipe-zap-dataset fipezap-sqlite build --download
```

## Install (uv)

```
uv tool install git+https://github.com/thiagoeh/fipe-zap-dataset
fipezap-sqlite build --download
```

## Query

"qual o preço médio do m² na cidade X?"

```sql
SELECT data, tipo, valor FROM preco_m2_atual WHERE cidade = 'São Paulo';
```

Table: `fipezap (cidade, data, tipo, categoria, metrica, valor)`. Views: `atual`
(latest snapshot per city), `preco_m2_atual` (latest non-null m² price per tipo).

## Development

```
mise run          # list tasks
mise run check    # lint + tests
mise run build    # download + rebuild data/fipezap.sqlite3
```

AI agents: read `AGENTS.md` before changing anything.
