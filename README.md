# Webfleet Log Downloader

Streamlit app for downloading Webfleet trip reports and generating Webfleet, RDA, LTR, merge, and audit reports.

## Run locally

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub.
2. In Streamlit Community Cloud, create a new app from the repository.
3. Select `app.py` as the entry point.
4. Deploy from the branch you want to share.
5. For a private app, deploy from a private GitHub repository and invite coworkers by email from the Streamlit Cloud sharing settings.

Generated files are temporary server artifacts for the current app session. Users must download generated CSV, XLSX, and ZIP files from the app; reports are not saved to the user's desktop automatically.

After a Webfleet CSV exists in the session, open the Dashboard tab to filter and inspect trip data. The dashboard focuses on:

- `tripmode`
- `start_time`
- `end_time`
- `duration`
- `distance`
- `drivername`
- `driverno`
- `objectname`

## Notes

- The Webfleet `showTripReportExtern` endpoint is rate limited. The default app setting waits 61 seconds between requests.
- Webfleet credentials are entered in the app and are not persisted by the app.
- Generated Nexus batch files prompt for Nexus credentials when run. Nexus credentials are not hardcoded in the app.
- Excel output is skipped when the CSV has more rows than one Excel sheet can hold.
- No Docker setup is required for hosting.
# Planning / RDA comparison workbook

The PDF/Webfleet/RDA/planning workflow adds a `Planning_RDA_Comparison` folder
to the complete ZIP package, containing `planning_rda_comparison.xlsx`.
Run the audit again to include it in a new download.

The workbook includes overall, daily, collaborator, day/collaborator, client-time,
active-client-time, client totals and day/collaborator/client comparisons, plus
exceptions, activity breakdowns, absences, data-quality issues and source entries.
Its `Read_Me` sheet explains all calculation and matching rules.

Time differences are original recorded RDA minus planning, in minutes and decimal
hours. Client counts use the main key in `Matched Clients` when aliases are mapped
unambiguously, otherwise source IDs (name fallback when no ID is present), with
separate checks for equal counts and matching client identities. Names are retained
for reading the report; anonymized names do not merge different IDs. All source
dates are included independently of PDF filters. RDA cutting does not change these
totals. Missing mappings and invalid values remain visible for investigation.

Entity detail sheets include `collab_all_ids` and `client_all_ids`, formatted as
`(2343, 3432, 4322)`, alongside the main mapping keys. `collab_record_id` and
`client_record_id` preserve the IDs actually used on each source entry. Entity
summary sheets list contributing Planning and RDA IDs separately, since a summary
can combine records with different IDs. Mapping aliases are included even when
they have no records in the selected files; shared names do not create ID links.
