# Webfleet Tools Script Guide

This document explains what the scripts in this project do, what information they need, what they create, and how the pieces fit together. It is written for a reader who uses the tool or manages the process, not for a programmer.

## Big Picture

This project is a Streamlit web app called **Outils Webfleet**. It brings several internal data jobs into one browser interface:

| App section | Main purpose | Typical result |
|---|---|---|
| Accueil | Starting page with instructions and template downloads | User guidance |
| Telechargement Webfleet | Download trip logs from Webfleet | CSV and Excel trip report |
| Fusionner des fichiers | Combine several files with the same columns | One merged CSV or Excel file |
| Transferts RDA | Prepare RDA files for Nexus, including the 61010 15-minute rule | ZIP folder with Nexus CSVs, batch files, maps, and checks |
| Controle LTR | Check labor-time/rest rules from RDA data | Multi-sheet Excel audit workbook |
| Audit Webfleet-RDA | Compare Webfleet trips, RDA work entries, mapping, and planning | Excel audit report plus optional Gantt PDFs and RDA correction files |

Generated reports are temporary app-session files. The app expects the user to download the CSV, Excel, or ZIP files from the buttons shown in the browser.

## Main App Files

| File | What it does | Who normally touches it |
|---|---|---|
| `app.py` | Starts the Streamlit app, shows the side menu, and sends the user to the selected section. | Usually only developers |
| `app_config.py` | Stores shared settings: API URL, accepted column names, output folder names, RDA codes, audit thresholds, and menu labels. | Developers or process owner when business rules change |
| `ui_common.py` | Shared helper functions for reading files, merging data, cleaning Excel text, and showing download buttons. | Developers |
| `webfleet.py` | Webfleet API downloader and Webfleet dashboard. | Users through the app |
| `merge_files.py` | File merge tool and merged-file dashboard. | Users through the app |
| `rda_transfers.py` | RDA 15-minute adjustment and RDA-to-Nexus export builder. | Users through the app |
| `nexus_batch_runner.py` | Validates prepared Nexus data and can run the Nexus import client. | Users with Nexus access |
| `ltr_checks.py` | Runs LTR checks by loading logic from the LTR notebook. | Users through the app |
| `audit_webfleet_rda.py` | Full Webfleet/RDA/planning audit, dashboards, PDFs, and optional RDA cutting/manual review outputs. | Users through the app |
| `desktop_launcher.py` | Starts the app locally, finds a free browser port, and opens the browser. | Used when packaging as a desktop app |
| `launch_streamlit.cmd` | Simple Windows command file to launch Streamlit on port 8501. | Local Windows users |
| `requirements.txt` | Python packages needed by the app. | Setup/deployment |
| `runtime.txt` | Python runtime version for hosting. | Deployment |
| `WebfleetTools.spec` | PyInstaller packaging recipe for building a desktop executable. | Developers/release owner |

## How A User Starts The App

There are two common ways:

| Method | What happens |
|---|---|
| `streamlit run app.py` | Starts the browser app directly from the code folder. |
| `launch_streamlit.cmd` | Windows shortcut-style launcher. It runs `app.py` and writes logs to `streamlit_current.out.log` and `streamlit_current.err.log`. |
| Packaged desktop app | `desktop_launcher.py` is used by PyInstaller. It finds an open port, starts Streamlit, and opens the browser automatically. |

## Section 1: `app.py` - Home And Navigation

**Goal:** provide one simple entry point for all tools.

What it does:

1. Loads all main modules.
2. Creates a side menu with the available tasks.
3. Shows the home page when the app opens.
4. Calls the correct section script when a user selects a task.
5. Adds shared styling so the app has the same look across pages.

Important user-facing information on the home page:

| Topic | Why it matters |
|---|---|
| Download generated files during the session | Reports are temporary and should not be assumed to stay forever on the server. |
| Nexus `nx-spi-client` instructions | RDA batch files need the Nexus client executable in the expected folder. |
| Webfleet API instructions | The app cannot create the Webfleet API key; it must be requested and enabled in Webfleet. |
| Section overview | Helps users pick the correct tool. |

Inputs: none directly, except the user choosing a menu item.

Outputs: no data file; it shows instructions and links the user to the other tools.

## Section 2: `webfleet.py` - Webfleet Trip Download

**Goal:** download trip logs from Webfleet safely, even for long date ranges.

### Inputs Needed

| Input | Meaning |
|---|---|
| Compte | Webfleet account name. |
| Utilisateur | Webfleet user name. |
| Mot de passe | Webfleet password. |
| Cle API | API key provided by Webfleet. |
| Date de debut | First day to download. |
| Date de fin | Last day to download. |
| Jours max par requete | How many days are requested at once. Default is 7. |
| Secondes entre requetes | Waiting time between API calls. Default is 61 seconds because Webfleet is rate-limited. |
| Nombre max de tentatives | How many times to retry a failing request. |
| Timeout requete | How long to wait before a Webfleet request is considered stuck. |

### How It Works

The downloader does not ask Webfleet for the whole period in one call. Instead, it splits the selected period into smaller blocks, downloads each block, and saves a checkpoint file for each successful block.

If a block fails, the script tries again. If a large block still fails, it splits that block into smaller pieces. This makes the download more recoverable.

### Important Safeguards

| Safeguard | Plain-language explanation |
|---|---|
| API wait time | Prevents sending requests too quickly. |
| Retries | Temporary Webfleet errors do not immediately stop the whole job. |
| Checkpoints | Already downloaded blocks are reused if the same job is restarted. |
| Empty markers | Days with no trips are recorded as intentionally empty. |
| Trip ID deduplication | If Webfleet returns overlapping data, repeated `tripid` values are kept only once. |
| Download audit | Checks whether all requested days were covered and whether rows were lost. |

### Outputs

| Output | Description |
|---|---|
| `webfleet_ALL_TRIPS_...csv` | Main trip report. |
| `webfleet_ALL_TRIPS_...xlsx` | Excel version, if the row count fits inside one Excel sheet. |
| `checkpoint_chunks/` | One CSV per downloaded period plus empty markers. |
| `download_manifest.csv` | Log of downloaded, empty, and failed periods. |
| Dashboard tab | Lets the user filter trips by driver, object, trip mode, date, distance, duration, and search text. |

## Section 3: `merge_files.py` And `ui_common.py` - File Merge

**Goal:** combine several CSV or Excel files into one file.

### Inputs Needed

| Input | Meaning |
|---|---|
| At least two files | CSV, XLSX, or XLS. |
| Same header names | The files must have the same column names. The order may differ. |
| Output format | XLSX or CSV. |
| Output base name | Optional name for the generated file. |

### How It Works

1. Reads each uploaded file.
2. Removes unnamed Excel helper columns.
3. Uses the first file as the reference for column names and order.
4. Checks that every later file has the same column names.
5. Stacks the rows into one table.
6. Checks that the final row count equals the sum of all input rows.
7. Writes the merged file.

### Outputs

| Output | Description |
|---|---|
| Merged CSV or XLSX | The combined file. |
| Merge summary | Number of files, input rows, output rows, and schema information. |
| Dashboard | Search, filter, preview, and download a filtered CSV. |

## Section 4: `rda_transfers.py` - RDA Adjustment And Nexus Export

**Goal:** prepare RDA data for Nexus imports, especially where prestation `61010` must be limited to 15 minutes.

### Main Tasks In This Section

| Tab | What it does |
|---|---|
| Ajustement 15 minutes | Reduces `61010` entries above 15 minutes and moves the extra minutes to allowed prestations. |
| Transfert UO vers UO | Prepares a transfer from one organizational unit to another. |
| Transfert vers Nexus | Uses `nexus_batch_runner.py` to validate and optionally run a Nexus import. |

### Inputs Needed For 15-Minute Adjustment

| Input | Meaning |
|---|---|
| Fichier RDA a ajuster | Raw RDA file in CSV, XLSX, or XLS format. |
| UO du fichier RDA | The organizational unit of that RDA file, such as `NE 301`, `SARL 201`, or `SA 101`. |
| Nom du dossier | Name used for the output package. |
| Prestations autorisees | Codes allowed to receive minutes removed from `61010`. Default: `11000,11100,11200,14000,14100,14200`. |

### Required RDA Columns

The script accepts several possible names for the same idea:

| Needed information | Accepted examples |
|---|---|
| Date | `Date Debut`, `Date`, `Jour`, `Date de prestation` |
| Start time | `Debut`, `Heure debut`, `Von` |
| End time | `Fin`, `Heure fin`, `Bis` |
| Prestation code | `N Prestation`, `No prestation`, `Code prestation`, `Prestation`, `Code` |
| Duration | `Duree`, `Duree (min)`, `Dauer_verrechnet` |
| Client | `N du client`, `No client`, `KD-Nr` |
| Collaborator | `No collaborateur`, `Collaborateur`, `Mitarbeiter-ID` |

Note: column names in the real files may contain accents, such as `Debut`/`Début` or `Duree`/`Durée`. The code includes the accented versions.

### How The 61010 Rule Works

For each collaborator and day:

1. The script finds rows with prestation `61010`.
2. If a `61010` row is longer than 15 minutes, it keeps only 15 minutes on that row.
3. The removed minutes are added to nearby rows whose prestation code is authorized.
4. The total minutes must stay the same. The script stops if minutes are lost or added incorrectly.
5. Start and end times are rebuilt so the adjusted durations still match the times.
6. The script checks for overlaps after the adjustment.

### Outputs

| Output | Description |
|---|---|
| ZIP folder | Main download containing all generated RDA/Nexus files. |
| Adjusted RDA file | RDA after the 61010 adjustment. |
| `HAS_map_main.csv` | Mapping file for Nexus prestation codes. |
| `RDA_duree_check.csv` | Duration totals used for checking. |
| `01_All_Collabs_One_CSV/` | One import CSV and batch for all collaborators. |
| `02_Collabs_With_61010_One_CSV/` | One import CSV and batch only for collaborators who had 61010 entries. |
| `03_Per_Collab_Separate/` | Separate import folders per collaborator. |
| QA tables | Checks for total minutes, generated CSVs, and overlaps. |

## Section 5: `nexus_batch_runner.py` - Nexus Validation And Import

**Goal:** check a Nexus import file before transfer, generate the import files, and optionally run the Nexus client.

### Inputs Needed

| Input | Meaning |
|---|---|
| Type de fichier d'entree | Either raw RDA or already prepared Nexus file. |
| Fichier RDA brut | RDA file to convert to Nexus format. |
| Fichier Nexus prepare | File already containing Nexus columns. |
| Chemin local du dossier `nx-spi-client` | Folder containing `Asebis.Client.StarterCommand.exe`. |
| UO cible | Target organizational unit when starting from raw RDA. |
| Utilisateur Nexus | Nexus username, only when running the transfer. |
| Mot de passe Nexus | Nexus password, only when running the transfer. |
| Confirmation checkbox | User must confirm before the transfer can run. |

### Expected Nexus Columns

| Column | Meaning |
|---|---|
| `Datum` | Date |
| `Von` | Start time |
| `Bis` | End time |
| `Leistungscode` | Prestation code |
| `Dauer_verrechnet` | Duration charged |
| `OE` | Organizational entity |
| `KD-Nr` | Client number |
| `Klient` | Client flag/value used by Nexus |
| `Einsatzgrund` | Reason/value used by Nexus |
| `Mitarbeiter-ID` | Collaborator ID |

### Checks Before Transfer

The script blocks the transfer if serious errors are found:

| Check | Why it matters |
|---|---|
| File is not empty | Nexus should not receive an empty import. |
| Dates and times are readable | Bad dates or times would import incorrectly. |
| Duration is numeric and not negative | Nexus needs clean durations. |
| Duration matches start/end | Prevents inconsistent records. |
| Prestation code exists | Required for the import. |
| Client/collaborator IDs are valid | Prevents unmapped or invalid people/clients. |
| OE is unique | A prepared Nexus file must target one OE. |
| Duplicate rows | Shown as a warning for review. |

### Outputs

| Output | Description |
|---|---|
| `RDA_Nexus.csv` | CSV used for Nexus import. |
| `HAS_map.csv` | Code mapping used during import. |
| `RDA_Nexus_batch.bat` | Batch file that calls the Nexus client. |
| `RDA_Nexus_transfer.log` | Transfer result, command without password, return code, and stdout/stderr. |

## Section 6: `ltr_checks.py` - LTR Controls

**Goal:** check work/rest rules using RDA data and a collaborator matching workbook.

This script loads several functions from `Scripts/Check LTR.ipynb`. The notebook contains the detailed rule logic; the Streamlit script wraps it in the app, handles file upload/download, and displays dashboards.

### Inputs Needed

| Input | Meaning |
|---|---|
| Classeur collaborateurs matches | Excel workbook that maps/identifies collaborators. |
| Fichier RDA fusionne | RDA data to check. CSV, XLSX, or XLS. |
| Nom du dossier de sortie | Optional folder name for the result. |

### Main Rules Checked

| Rule | Plain-language meaning |
|---|---|
| `OVER_50H_WEEK` | Work time above the weekly limit. |
| `STREAK_7DAYS` | Too many consecutive worked days. |
| `SPAN_OVER_14H` | A service span longer than the configured limit. |
| `REST_UNDER_11H` | Not enough rest between services. |
| `PAUSE_INSUFF` | Insufficient pauses/breaks. |

### Outputs

The main output is a multi-sheet Excel workbook. Important sheets include:

| Sheet | What it contains |
|---|---|
| `SUMMARY_BY_MONTH` | Month-by-month counts. |
| `ALL_INFRACTIONS` | All detected issues in one sheet. |
| `OVER_50H_WEEK` | Details for weekly-hour issues. |
| `STREAK_7DAYS` | Details for consecutive-day issues. |
| `SPAN_OVER_14H` | Details for long service spans. |
| `REST_UNDER_11H` | Details for rest issues. |
| `REST_REVIEW_ALLOWED` | Cases that may need review rather than automatic failure. |
| `PAUSE_INSUFF` | Break/pause issues. |
| `SERVICES_AUDIT` | How services were built from rows. |
| `CALENDAR_HOUR_SLICES` | Calendar-based time slices used for weekly checks. |
| `DATA_QUALITY` | Bad or unclear input rows. |

The app also shows charts and filters for months, rules, and collaborators.

## Section 7: `audit_webfleet_rda.py` - Webfleet/RDA/Planning Audit

**Goal:** compare Webfleet vehicle trips with RDA work entries and planning data to find suspicious or unexplained vehicle use.

### Inputs Needed

| Input | Meaning |
|---|---|
| Fichier RDA | RDA entries with collaborator, date, time, duration, and prestation information. |
| Fichier Webfleet | Trip data downloaded from Webfleet. |
| Fichier Mapping | Workbook used to match collaborator IDs across systems. |
| Fichier Planning | Planning file with scheduled work periods. |

### What It Compares

| Data source | Used for |
|---|---|
| RDA | What work was recorded. |
| Webfleet | When and how vehicles were used. |
| Mapping | Which people match between systems. |
| Planning | What was scheduled. |

### Main Ideas In The Audit

The audit builds a day-by-day picture for each collaborator:

1. Normalize dates, times, collaborator IDs, prestation codes, and trip distances.
2. Match RDA, Webfleet, mapping, and planning data.
3. Build daily RDA work blocks.
4. Apply buffers around work time, such as before-shift and after-shift windows.
5. Mark Webfleet trips that look normal, outside service, before shift, after buffer, or otherwise suspicious.
6. Summarize kilometers, trip modes, flags, and suspect private kilometers.
7. Create an Excel report and optional Gantt-style day charts.

### Outputs

| Output | Description |
|---|---|
| Excel audit report | Main report with collaborator summaries, daily aggregates, trips, RDA entries, planning, mapping, flags, and 61010 checks. |
| Dashboard | Metrics, filtered tables, charts, and one-day Gantt viewer. |
| PDF ZIP | Gantt charts comparing Webfleet, original RDA, and planning lanes. |
| RDA cutting package | Optional output that suggests or applies time cuts based on Webfleet overlaps. |
| Manual review package | Optional worklist for cases that should be reviewed by a person. |

### Important Audit Settings

These are configured in `app_config.py`:

| Setting | Meaning |
|---|---|
| `AUDIT_WORK_END_BUFFER_MIN = 30` | Buffer after work end. |
| `AUDIT_PRE_SHIFT_BUFFER_MIN = 30` | Buffer before shift. |
| `AUDIT_INTERNAL_BLOCK_GAP_MIN = 180` | Gap threshold used to group internal work blocks. |
| `AUDIT_MAX_REASONABLE_SPEED_KMH = 160` | Speed above this may be treated as unreasonable. |
| `AUDIT_PRESTATION_61010_CODE = "61010"` | Special prestation code checked in the audit. |
| `AUDIT_TZ_NAME = "Europe/Zurich"` | Time zone used for local comparisons. |

## Supporting Configuration: `app_config.py`

This file is the central place for shared constants.

| Area | Examples |
|---|---|
| Webfleet API | API URL, action name, output format, language. |
| App menu | Labels for Home, Webfleet, Merge, RDA, LTR, Audit. |
| Output folders | `WebfleetReports`, `MergedReports`, `RDAReports`, `LTRReports`, `AuditReports`. |
| RDA business rules | OE values, prestation `61010`, allowed target codes, accepted column names. |
| Audit rules | Time buffers, planning colors, speed threshold, time zone. |
| Session output | Creates a temporary output folder per app session. |

If business rules change, this is often the first file to inspect.

## Shared Helpers: `ui_common.py`

This file contains common actions used by several sections:

| Helper area | What it does |
|---|---|
| File reading | Reads CSV, XLSX, and XLS files with flexible separators and sheets. |
| Merge logic | Checks matching headers and combines rows. |
| Excel cleanup | Removes characters Excel cannot write. |
| Download buttons | Shows a real download button only when a file exists. |
| Template downloads | Shows ZIP templates from a `Templates` folder if present. |
| Folder names | Cleans unsafe characters from names used for output folders. |

## Notebooks In `Scripts/`

These notebooks are older or reference workflows. Some logic has been moved into the Streamlit app, while some is still used as a source of functions.

| Notebook | Purpose |
|---|---|
| `Scripts/Copy_of_Download_webfleet_logs.ipynb` | Original Google Colab workflow for downloading Webfleet trips to Google Drive. The app version is now mainly `webfleet.py`. |
| `Scripts/Merge.ipynb` | Original notebook-style merge workflow. The app version is now mainly `merge_files.py` and `ui_common.py`. |
| `Scripts/Adjustment_RDA_15min.ipynb` | Original RDA 61010 adjustment workflow. The app version is now mainly `rda_transfers.py`. |
| `Scripts/Check LTR.ipynb` | LTR rule logic. This is still loaded by `ltr_checks.py` for the core calculations. |
| `Scripts/planning_webfleet_rda_charts.ipynb` | Original Webfleet/RDA/planning audit and chart workflow. The app version is now mainly `audit_webfleet_rda.py`. |

There are also root-level notebook copies:

| Notebook | Notes |
|---|---|
| `Copy_of_Download_webfleet_logs.ipynb` | Root copy of the Webfleet download notebook. |
| `new_Copy_of_Adjustment_RDA_15min.ipynb` | Root copy/newer copy of the RDA 15-minute notebook. |

These notebooks are useful for understanding history, testing logic in Colab, or comparing old output. For normal use, the Streamlit app is the cleaner entry point.

## File And Folder Outputs

| Tool | Output location behavior |
|---|---|
| Webfleet download | Temporary session folder under the system temp directory. User downloads CSV/XLSX from the app. |
| Merge | Temporary session folder. User downloads merged file from the app. |
| RDA transfers | Temporary session folder. User downloads ZIP and/or selected files. |
| Nexus runner | Temporary session folder. User downloads generated CSV/map/batch/log. |
| LTR checks | Temporary session folder. User downloads Excel workbook. |
| Audit | Temporary session folder. User downloads Excel report and optional ZIP packages. |

The app does not automatically save finished reports to the user's Desktop.

## Common User Mistakes And What The App Checks

| Situation | What happens |
|---|---|
| Webfleet credentials are missing | The download button stays disabled. |
| Webfleet date start is after date end | The app shows an error. |
| Files to merge have different headers | The merge stops and lists missing/extra columns. |
| RDA required columns are missing | RDA processing stops with a clear missing-column message. |
| RDA duration does not match start/end | RDA/Nexus validation can stop the transfer. |
| Nexus client path is wrong | The Nexus transfer cannot run. |
| Nexus file has invalid dates/times/durations | The transfer is blocked. |
| LTR collaborator matching is incomplete | The LTR report includes unrecognized rows. |
| Audit mapping does not match people correctly | Audit reports and dashboards show unmapped or suspicious results that need review. |

## Dependencies

The app needs these Python packages from `requirements.txt`:

| Package | Used for |
|---|---|
| `streamlit` | Browser app interface. |
| `pandas` | Reading, cleaning, and writing tabular data. |
| `requests` | Webfleet API calls. |
| `openpyxl` | Excel file reading/writing. |
| `numpy` | Calculations. |
| `pytz` | Time zone compatibility for notebook logic. |
| `xlrd` | Older Excel file support. |
| `matplotlib` | Audit Gantt/PDF charts. |
| `altair` | Some dashboard charts. |

## Practical Order Of Use

For a typical monthly workflow:

1. Download Webfleet trips in **Telechargement Webfleet**.
2. Merge source files if needed in **Fusionner des fichiers**.
3. Prepare or adjust RDA in **Transferts RDA**.
4. Run **Controle LTR** if labor/rest checks are required.
5. Run **Audit Webfleet-RDA** with RDA, Webfleet, mapping, and planning.
6. Download all generated reports before ending the session.

## Quick Glossary

| Term | Meaning |
|---|---|
| RDA | Work/service data file used by the organization. |
| Webfleet | Vehicle trip tracking system. |
| Nexus | Target system that receives prepared import CSV files. |
| OE / UO | Organizational unit/entity code used by Nexus/RDA. |
| Prestation | Service/activity code. |
| `61010` | Special prestation code that often needs the 15-minute cap rule. |
| HAS map | CSV mapping external prestation codes to Nexus tariff numbers. |
| Batch file / `.bat` | Windows script used to launch a Nexus import. |
| Checkpoint | Saved partial download so a long job can resume or be audited. |
| Gantt chart | Timeline-style chart showing Webfleet trips, RDA entries, and planning on the same day. |
