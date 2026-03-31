# Liseo Dashboard Live
Live dashboard for Liseo production data.

## Refresh paths

- GitHub Actions can still be run manually from `workflow_dispatch` if you want a one-off cloud refresh.
- `scripts/refresh_dashboard_data.py` can now auto-detect the local workbook and falls back to an isolated Excel snapshot if the workbook is open and locked.
- `scripts/publish_dashboard_data.py` refreshes the JSON from the local synced workbook and pushes changes to GitHub.
- `scripts/run_local_autopublish.pyw` runs the local publish silently and can fall back to auto-detect the workbook if the configured path is missing.
- `scripts/register_local_autopublish.ps1` creates a Windows scheduled task that runs the local publisher every minute on this PC and can auto-detect the workbook path.
- `scripts/save_excel_snapshot.ps1` creates a safe hidden Excel copy when the live workbook is open.
- The dashboard page itself checks for a fresher `dashboard_data.json` every 60 seconds and also has a manual `Refresh Now` button.
