# Webfleet Log Downloader

Streamlit app for downloading Webfleet trip reports through the CSV API to a local folder.

## Run

```powershell
pip install -r requirements.txt
streamlit run app.py
```

The app writes resumable checkpoint files into the selected local output folder. Rerunning the same date range skips completed chunks and combines the available checkpoints into a local CSV for the dashboard.

Browser downloads only happen when you click the export buttons after the trips have been downloaded. Excel files are created only after clicking **Create Excel file**.

After a CSV exists, open the Dashboard tab to filter and inspect trip data. The dashboard focuses on:

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
- Credentials are entered in the Streamlit sidebar and are not written to disk by the app.
- Excel output is skipped when the CSV has more rows than one Excel sheet can hold.
- No Google Drive or cloud output is used.
