# OMSA Test Automation v2

A local Windows + Streamlit application for automating OMSA Excel testing while keeping the existing Excel workbooks as the source of truth for business logic.

## Correct operating model

**Fixed master template + named run-specific upload fields → automated population → native Microsoft Excel recalculation → completed Excel workbook.**

The user does **not** upload the master template on every run. Every Loader, View and Control has a fixed active template. The user only uploads the files needed to populate that template for the reporting date.

Loaders, Views and Controls do not silently reuse files from previous runs. If a View needs a Loader file, the View page shows that Loader file as a dedicated upload field.

## Included configuration

- 17 Loader templates/configurations
- 10 View templates/configurations
- 9 Control templates/configurations
- OMSA-style horizontal navigation and full-list → Launch → detail workflow
- Named upload field for every currently identified input
- Template Management
- Relations / data-lineage table
- Run History
- Native Excel COM engine for `.xlsx`, `.xlsm` and `.xlsb`

## First run on Windows

1. Extract the READY package completely. Do not run it from inside the ZIP.
2. Microsoft Excel Desktop must be installed.
3. Double-click `SETUP_AND_RUN.bat`.
4. When Streamlit starts, open `http://localhost:8501` if the browser does not open automatically.
5. Start with **Loaders → Anagrafica strumenti** or **Loaders → NAV Details**.

Subsequent runs: double-click `RUN.bat`.

## Git / GitHub

`templates_local/` is intentionally ignored by Git. The READY package contains the templates for local testing, but the large business Excel files should normally remain local rather than being pushed to GitHub.

Double-click `INIT_GIT.bat` to initialize a local repository, then commit the code/configuration. If a private GitHub remote is later created, add the remote and push normally.

## Business-logic boundary

Python does not recreate the reconciliation formulas, XLOOKUPs, tolerances, thresholds or TRUE/FALSE rules. Those remain in Excel. Python performs orchestration: copy template, clear configured input areas, import run data, apply explicit operational transformations, extend existing formulas, recalculate in Microsoft Excel and return the workbook.

## Components marked Review / Mapping required

Some source workbooks contain contradictory or incomplete instructions. Those components are visibly flagged in the UI. Most are still runnable with the best-supported mapping; `MTB / FBL003BIS` is intentionally blocked until its transpose/column mapping is confirmed.

See `docs/OMSA_Test_Automation_English_Upload_Field_Specification.docx` for the detailed upload-field specification.
