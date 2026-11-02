# SEC Financial & Sentiment Analyzer — Web Application

This folder (`docs/`) serves as the root web directory for **GitHub Pages**.

## 📁 Folder Structure
* `index.html`: Responsive single-page application dashboard.
* `css/style.css`: Financial dashboard dark theme styling.
* `js/app.js`: State management, controllers, and Chart.js dual-axis rendering logic.
* `data/`:
  * `aapl_data.json`: 10-year audited Apple metrics, financial statements, and ratios.
  * `amd_data.json`: 10-year audited AMD metrics, financial statements, and ratios.
  * `msft_precomputed.json`: 10-year audited Microsoft peer benchmark metrics and ratios.
  * `comparison_data.json`: Cross-company comparison metrics and baseline figures.
  * `data_bundle.js`: Pre-bundled script version of datasets for zero-latency, offline execution.
* `downloads/`: Exact copies of verified Excel workbooks and `sec_downloader.py` for direct user download.

## 🚀 GitHub Pages Configuration
To host this directory on GitHub Pages:
1. Go to repository **Settings** -> **Pages**.
2. Under **Branch**, select `main` and set the folder to `/docs`.
3. Click **Save**.
