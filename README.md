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
