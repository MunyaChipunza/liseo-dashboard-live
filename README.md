# Liseo Dashboard Live
Live dashboard for Liseo production data.

## Refresh paths

- GitHub Actions can still be run manually from `workflow_dispatch` if you want a one-off cloud refresh.
- `scripts/publish_dashboard_data.py` refreshes the JSON from the local synced workbook and pushes changes to GitHub.
- `scripts/register_local_autopublish.ps1` creates a Windows scheduled task that runs the local publisher every minute on this PC.
- The dashboard page itself checks for a fresher `dashboard_data.json` every 60 seconds and also has a manual `Refresh Now` button.
