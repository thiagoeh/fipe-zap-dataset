## Rules

- Run tasks via mise; never call `pytest`, `ruff`, `pip`, or `uv sync` directly.
- The Python version lives only in `.python-version`; never redeclare it elsewhere.
- No git write operations (commit, stage, branch) unless explicitly asked.
- No code comments. Fix the code, never widen ruff ignores or add `noqa`.

## Whys (not visible in code)

- The DB's #1 job: "qual o preço médio do m² na cidade X?" — `preco_m2_atual` returns the latest non-null m² price per tipo, segmento included; `atual` returns the latest snapshot per (cidade, segmento), NULLs included. Mixed months across cities, segmentos and tipos are expected (rent data lags sale).
- The strict date CHECK exists because SQLite silently rolls over invalid dates: `2008-13-01` passes a plain `strftime` equality test.
- The composite PK starts with `(cidade, segmento, metrica, categoria)` so the critical query is a prefix scan.
- Only the `Resumo` and `Aux` sheet skips are deliberate — don't add them back. Everything else parses: `Índice FipeZAP` sheet becomes cidade `Brasil`, the `Rentabilidade do aluguel` block becomes metrica `rentabilidade_aluguel` (tipo `locacao`, label prefix `(% - mensalizada`), and `Imóveis comerciais` columns land in segmento `comercial`.
- Downloads require a browser User-Agent because FIPE 403s the default Python one.
- The `tests/` fixture mirrors the real XLSX header layout; update both together.
