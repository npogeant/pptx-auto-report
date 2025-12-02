from data_loader import load_data
from chart_builder import build_line_chart_data
from slide_builder import (
    create_presentation,
    add_title_slide,
    add_chart_slide,
    save_presentation
)

def main():
    df = load_data("../data/sample_data.csv")

    # Create presentation
    prs = create_presentation()
    add_title_slide(prs, "Automated Report", "Generated with Python")

    # Build chart data
    x_col = df.columns[0]
    y_cols = df.columns[1:]
    chart_data = build_line_chart_data(df, x_col, y_cols)

    add_chart_slide(prs, chart_data, title="Monthly Stock Prices")

    save_presentation(prs)

if __name__ == "__main__":
    main()
