from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
import pandas as pd

def build_line_chart_data(df, x_col, y_cols):
    chart_data = CategoryChartData()
    # try to parse date-like category (format YYYY-MM) and format as 'Oct 24'
    try:
        dates = pd.to_datetime(df[x_col], format="%Y-%m", errors="coerce")
        if dates.isnull().any():
            chart_data.categories = df[x_col].astype(str).tolist()
        else:
            chart_data.categories = dates.dt.strftime('%b %y').tolist()
    except Exception:
        chart_data.categories = df[x_col].astype(str).tolist()
    #chart_data.add_series(y_col, df[y_col].tolist())
    for col in y_cols:
        chart_data.add_series(col, df[col].tolist())

    # attach series names and evolution (start -> end) so callers can compute summaries
    chart_data.series_names = list(y_cols)
    evolutions = []
    try:
        for col in y_cols:
            series = df[col].dropna().astype(float)
            if len(series) == 0:
                evolutions.append((None, None))
                continue
            start = float(series.iloc[0])
            end = float(series.iloc[-1])
            delta = end - start
            try:
                pct = (delta / start) * 100.0 if start != 0 else None
            except Exception:
                pct = None
            evolutions.append((float(delta), None if pct is None else float(pct)))
    except Exception:
        evolutions = []

    chart_data.evolution = evolutions
    return chart_data

def build_bar_chart_data(df, x_col, y_cols):
    chart_data = CategoryChartData()
    chart_data.categories = df[x_col].tolist()
    for col in y_cols:
        chart_data.add_series(col, df[col].tolist())
    return chart_data
