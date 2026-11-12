## Rules

- Run tasks via mise; never call `pytest`, `ruff`, `pip`, or `uv sync` directly.
- The Python version lives only in `.python-version`; never redeclare it elsewhere.
- No git write operations (commit, stage, branch) unless explicitly asked.
- No code comments. Fix the code, never widen ruff ignores or add `noqa`.

## Whys (not visible in code)

- The DB's #1 job: "qual o preço médio do m² na cidade X?" — the `atual` view returns the latest snapshot per city, NULLs included; mixed months across cities and tipos are expected (rent data lags sale).
- The strict date CHECK exists because SQLite silently rolls over invalid dates: `2008-13-01` passes a plain `strftime` equality test.
- The composite PK starts with `(cidade, metrica, categoria)` so the critical query is a prefix scan.
- All sheet and block skips are deliberate — don't add them back.
- Downloads require a browser User-Agent because FIPE 403s the default Python one.
- The `tests/` fixture mirrors the real XLSX header layout; update both together.
