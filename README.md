# Sales Rep Strengths

Every morning (07:30 Cairo) GitHub Actions reads the master reservations sheet (tab FY26),
copies it into the working sheet, builds the dashboard and publishes it to GitHub Pages.

* `rep_strengths.py` — the Colab code, unchanged
* `run_github.py` — lets that code use the Google service account instead of the Colab sign-in
* `.github/workflows/rep_strengths.yml` — the daily schedule + publishing

The Excel file is kept on each run's page (Artifacts), not published.
