# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

Single-file Streamlit app (`validar.py`) that validates Excel files field by field before they are sent to a telecom operator (**CLARO** or **ETB**, Colombia). The user picks the operator, uploads an `.xlsx`/`.xls`, and gets a per-field / per-row error report plus a downloadable Excel report (`Errores Detalle` + `Resumen por Campo` sheets). All UI text, messages and identifiers are in Spanish — keep new code consistent with that.

## Commands

```powershell
# Windows, using the existing local venv
.\venv\Scripts\activate
pip install -r requirements.txt
streamlit run validar.py
```

There are no tests, linter, or build step in the repo. Importing `validar.py` executes the Streamlit UI at module level, so test it headlessly with `streamlit.testing.v1.AppTest.from_file("validar.py")`: `at.file_uploader[0].set_value((name, bytes, mime)).run()` simulates an upload, and the operator is switched with `at.session_state["operador"] = "ETB"; at.run()` (AppTest has no accessor for `st.segmented_control`). Build test workbooks with openpyxl; assigning `"#REF!"` to a cell produces a real error-typed cell.

Theme (light palette, Archivo + Atkinson Hyperlegible fonts from Google Fonts) lives in `.streamlit/config.toml`; the CSS block at the top of `validar.py` styles the app's own HTML blocks and targets keyed containers via their `.st-key-<key>` class (`barra`, `titulo`, `veredicto_ok`/`veredicto_err`, `acciones`, `campo_sel`). `.claude/skills/frontend-design` (symlink to `.agents/skills/`, installed via `npx skills`, tracked in `skills-lock.json`) was used for the visual design.

## Architecture (validar.py)

The file is organized in banner-delimited sections:

1. **FUNCIONES COMUNES** – reusable cell validators. Convention: most return `None` for empty, `True`/`False` otherwise; validators that need a reason (`valid_latlon`, `valid_puerto_olt`, `valid_estrato`, `valid_codigo_mun`) return `(ok, msg)`. `not_empty` treats `NaN`, `""`, `"nan"`, `"None"` as empty.
   This section also holds the Excel formula-error handling (`ERRORES_EXCEL`, `error_excel_en`, `celdas_con_error_excel`, `validar_distribucion`), see below.
2. **VALIDACIÓN CLARO** / 3. **VALIDACIÓN ETB** – each operator defines things that must stay in sync:
   - `COLUMNAS_<OP>`: required columns (drive the column chips and missing-column warning).
   - `OPCIONALES_<OP>`: optional columns (currently empty for both); renamed like required ones but never reported as missing.
   - `REGLAS_<OP>`: human-readable rule per column ("Reglas" popover, and the rule line above the error table when a field is selected). Display-only — it does not enforce anything; rules starting with "Opcional" render muted.
   - `validar_fila_<op>(row, idx)`: the actual enforcement; returns a list of `{"Fila", "Campo", "Valor", "Error"}` dicts.
   When changing a rule, update both the `REGLAS_*` text and the `validar_fila_*` logic.
4. **MOTOR GENÉRICO** – `OPERADORES` dict groups each operator's config (columns, optional columns, rules, brand color, description, `sin_errores_excel` columns) for the UI. `validar_dataframe` dispatches to the per-row function, for CLARO also runs `validar_duplicados_handle` (cross-row duplicate check on `HANDLE`), and appends `errores_extra`. `normalizar_columnas` maps headers to canonical names ignoring case, accents and surrounding spaces (`clave_columna`).
5. **UI** – runs top-level (no `main()`) and uses `st.stop()` for early exits (no upload, no rows, no errors). Layout is horizontal: a header bar (`st.container(horizontal=True)`) with title, operator `st.segmented_control` and a "Reglas" popover; a verdict strip (status | counts | download + "Columnas" popover) whose "ready" state requires no errors *and* no missing columns; then a master-detail explorer (`st.radio` key `campo_sel` styled as a list, plus a `buscar` text filter over the error table).

Key details:
- Excel is read with `pd.read_excel(..., dtype=str)` so validators receive strings; `valid_puerto_olt` still guards against `Timestamp` values because Excel often auto-converts `1/3/3` into a date.
- Reported row number is `idx + 2` (DataFrame index + header row + 1-based) so it matches the Excel row.
- **Excel error cells (`#REF!`, `#N/A`, ...)**: pandas' openpyxl reader turns error-typed cells into `NaN`, so they look empty. For the operator's `sin_errores_excel` columns (CLARO: `DISTRIBUCION`), `celdas_con_error_excel` re-opens the uploaded bytes with openpyxl (`data_only=True`, first sheet) to catch error-typed cells, while `validar_distribucion` catches error codes stored as text (English and Spanish variants). In every other column an error cell is still seen as empty.
- `DISTRIBUCION` is a required CLARO column (not ETB). Its cells may be empty while `DISTRIBUCION_PERMITE_VACIOS = True`; the business owner was unsure whether values are mandatory, so flipping it to `False` makes empty cells a "Campo vacío" error.
- Missing required columns are not skipped: `row.get` returns `None`, so every row gets "Campo vacío" errors for them.
- Adding a new operator requires: new `COLUMNAS_*`/`OPCIONALES_*`/`REGLAS_*`/`validar_fila_*`, an `OPERADORES` entry (plus a `--<op>` color variable in the CSS), and a branch in `validar_dataframe`.
- ETB format checks for `CTO` and `ID_POLIGON` currently only require five hyphen-separated alphanumeric groups — looser than the patterns described in `REGLAS_ETB`.

## Repo notes

- `README.md` and `requirements.txt` are UTF-16 encoded; preserve the encoding (or convert deliberately) when editing them.
- `venv/` and `__pycache__/` exist in the working tree; `__pycache__/` is tracked in git despite being a build artifact.
