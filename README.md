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
SELECT segmento, tipo, data, valor FROM preco_m2_atual WHERE cidade = 'São Paulo';
```

Filter `segmento = 'residencial' | 'comercial'` (or omit for both).

Table: `fipezap (cidade, segmento, data, tipo, categoria, metrica, valor)`. Views: `atual`
(latest snapshot per cidade+segmento), `preco_m2_atual` (latest non-null m² price per tipo).
Cities include `Brasil` (the national aggregate); `metrica` covers `rentabilidade_aluguel`
(rental yield).

## Development

```
mise run          # list tasks
mise run check    # lint + tests
mise run build    # download + rebuild data/fipezap.sqlite3
```

AI agents: read `AGENTS.md` before changing anything.
