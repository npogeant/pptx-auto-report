# From csv to slides

A small demonstration project showing how to generate PowerPoint slides programmatically with
`python-pptx`. This repository contains a minimal pipeline that loads a CSV dataset, builds a
line chart and summary cards, and writes a `.pptx` report.

**Prerequisites**
- Python 3.8+ installed on your system.
- `uv` CLI available locally (used here to run commands inside the repository environment).
	- You can install `uv` with `pip`:

```bash
# with pip (global or inside a venv)
pip install uv
```

Installation and quick start
----------------------------

1. Clone the repository and move into it:

```bash
git clone https://github.com/npogeant/pptx-auto-report.git
cd pptx-auto-report
```

2. Install the Python dependencies using `uv` and the provided `requirements.txt`:

```bash
uv run pip install -r requirements.txt
```

3. Run the example report generator:

```bash
# run from project root
uv run python src/main.py
```

This will generate a report at:

```bash
output/report.pptx
```


What this demo does
-------------------
- Loads sample data: `data/sample_data.csv` (monthly values for two series).
- Builds a `CategoryChartData` line chart (via `src/chart_builder.py`) and formats category labels
	as short month/year.
- Creates a slide with the chart on the left and two summary cards on the right (via
	`src/slide_builder.py`). The cards display percentage evolution between the first and last
	value for each series.
- Saves a `report.pptx` to `output/`.

Project layout
--------------

```bash
src/
 ├── main.py            # Script that generates the complete report
 ├── chart_builder.py   # Builds and styles line charts
 ├── slide_builder.py   # Creates slides, summary cards, overall layout
 └── data_loader.py     # Helper to load CSV data (Pandas)
data/
 └── sample_data.csv     # Demo dataset
output/
 └── report.pptx         # Generated report (ignored by .gitignore)
requirements.txt         # Project dependencies
```

Enjoy !
