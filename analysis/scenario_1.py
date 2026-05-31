from pathlib import Path

import plotly.graph_objects as go
import polars as pl
import xlsxwriter

from analysis.model_3 import Model3

START_DATE = "2026-04-01"
END_DATE = "2026-05-01"

CONFIGS = [
    (2, 2),
    (2, 4),
    (2, 8),
]

BREAKDOWN_ROWS = [
    ("Day-ahead Revenue", "da_revenue"),
    ("Production Tariffs", "prod_tariff"),
    ("Day-ahead Cost", "da_cost"),
    ("Consumption Tariffs", "cons_tariff"),
    ("Degradation Cost", "degradation"),
    ("FFR Revenue", "ffr_revenue"),
    ("FCR-D up Revenue (early)", "fcrd_up_E_revenue"),
    ("FCR-D up Revenue (late)", "fcrd_up_L_revenue"),
    ("FCR-D down Revenue (early)", "fcrd_down_E_revenue"),
    ("FCR-D down Revenue (late)", "fcrd_down_L_revenue"),
    ("FCR-N Revenue (early)", "fcrn_E_revenue"),
    ("FCR-N Revenue (late)", "fcrn_L_revenue"),
    ("aFRR up Revenue", "afrr_up_revenue"),
    ("aFRR down Revenue", "afrr_down_revenue"),
    ("mFRR up Revenue", "mfrr_up_revenue"),
    ("mFRR down Revenue", "mfrr_down_revenue"),
    ("Profit", "profit"),
]


def _row_value(breakdown: dict, key: str) -> float:
    return breakdown[key]


PARETO_WEIGHT_PAIRS = (
    [(0.9999, 0.0001)]
    + [(round(1.0 - i * 0.01, 2), round(i * 0.01, 2)) for i in range(1, 100)]
    + [(0.0001, 0.9999)]
)


def config_labels(configs: list[tuple[int, int]]) -> list[str]:
    return [f"{mw} MW / {mwh} MWh" for mw, mwh in configs]


def get_extreme_breakdowns(m: Model3) -> tuple[dict, dict]:
    """Solve at the first and last Pareto weight pairs and return breakdowns."""
    first_breakdown = last_breakdown = None
    for i, (lp, lc) in enumerate((PARETO_WEIGHT_PAIRS[0], PARETO_WEIGHT_PAIRS[-1])):
        m.lambda_profit = lp
        m.lambda_co2 = lc
        solved = m.solve()
        _, _, breakdown = m._extract_objectives(solved)
        if i == 0:
            first_breakdown = breakdown
        else:
            last_breakdown = breakdown
    return first_breakdown, last_breakdown


def print_breakdown_table(
    breakdowns: list[dict],
    column_labels: list[str],
    title: str,
) -> None:
    col_w = 16
    label_w = 32
    width = label_w + len(column_labels) * (col_w + 2)
    sep = "-" * width

    print(f"\n{title}")
    print(sep)
    header = f"{'':<{label_w}}"
    for col in column_labels:
        header += f"{col:>{col_w}}  "
    print(header)
    print(sep)

    for row_label, key in BREAKDOWN_ROWS:
        row = f"{row_label:<{label_w}}"
        for bd in breakdowns:
            row += f"{_row_value(bd, key):>{col_w},.2f}  "
        print(row)

    print(sep)
    print("Note: All values in DKK.")


def save_breakdown_excel(
    sheets: list[tuple[str, list[dict]]],
    configs: list[tuple[int, int]],
    out_path: str,
) -> None:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(str(out))

    header_fmt = wb.add_format(
        {"bold": True, "align": "center", "border": 1, "bg_color": "#D9D9D9"}
    )
    label_fmt = wb.add_format({"border": 1})
    num_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1})
    bold_num_fmt = wb.add_format(
        {"bold": True, "num_format": "#,##0.00", "border": 1, "bg_color": "#F2F2F2"}
    )
    bold_label_fmt = wb.add_format({"bold": True, "border": 1, "bg_color": "#F2F2F2"})

    for sheet_name, breakdowns in sheets:
        ws = wb.add_worksheet(sheet_name[:31])

        ws.write(0, 0, "", header_fmt)
        for col, (mw, mwh) in enumerate(configs, start=1):
            ws.write(0, col, f"{mw} MW / {mwh} MWh", header_fmt)

        for row_idx, (label, key) in enumerate(BREAKDOWN_ROWS, start=1):
            is_total = label == "Profit"
            lf = bold_label_fmt if is_total else label_fmt
            nf = bold_num_fmt if is_total else num_fmt
            ws.write(row_idx, 0, label, lf)
            for col, breakdown in enumerate(breakdowns, start=1):
                value = _row_value(breakdown, key)
                ws.write(row_idx, col, value, nf)

        ws.set_column(0, 0, 28)
        ws.set_column(1, len(configs), 18)

    wb.close()
    print(f"Breakdown saved to {out}")


def save_pareto_excel(results: list[dict], out_path: str) -> None:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    df = pl.DataFrame(
        {
            "lambda_profit": [r["lambda_profit"] for r in results],
            "lambda_co2": [r["lambda_co2"] for r in results],
            "profit_dkk": [r["profit"] for r in results],
            "co2_kg": [r["co2"] for r in results],
        }
    )
    df.write_excel(out)
    print(f"Pareto data saved to {out}")


def run_pareto(m: Model3) -> tuple[list[dict], dict, dict]:
    """Run 101 weight pairs and collect profit + CO2 and extreme breakdowns."""
    results = []
    first_breakdown = last_breakdown = None
    for i, (lp, lc) in enumerate(PARETO_WEIGHT_PAIRS):
        m.lambda_profit = lp
        m.lambda_co2 = lc
        print(f"  λ_profit={lp:.2f}, λ_co2={lc:.2f}")
        solved = m.solve()
        profit, co2, breakdown = m._extract_objectives(solved)
        if i == 0:
            first_breakdown = breakdown
        if i == len(PARETO_WEIGHT_PAIRS) - 1:
            last_breakdown = breakdown
        results.append(
            {"lambda_profit": lp, "lambda_co2": lc, "profit": profit, "co2": co2}
        )
    return results, first_breakdown, last_breakdown


PARETO_COLORS = ["steelblue", "seagreen", "darkorange"]


def load_pareto_excel(path: str) -> list[dict]:
    df = pl.read_excel(path)
    return [
        {
            "lambda_profit": r["lambda_profit"],
            "lambda_co2": r["lambda_co2"],
            "profit": r["profit_dkk"],
            "co2": r["co2_kg"],
        }
        for r in df.iter_rows(named=True)
    ]


def visualize_asset_pareto(
    all_results: list[list[dict]],
    configs: list[tuple[int, int]],
) -> None:
    all_profits = [r["profit"] for results in all_results for r in results]
    all_co2s = [r["co2"] for results in all_results for r in results]

    x_max = max(all_profits) * 1.05
    y_min = min(all_co2s) * 1.05 if min(all_co2s) < 0 else -1

    fig = go.Figure()

    fig.add_shape(
        type="rect",
        xref="x",
        yref="y",
        x0=0,
        x1=x_max,
        y0=y_min,
        y1=0,
        fillcolor="rgba(0, 200, 100, 0.15)",
        line_width=0,
        layer="below",
    )
    fig.add_annotation(
        x=x_max,
        y=y_min,
        text="Profit > 0 & CO₂ < 0",
        showarrow=False,
        font=dict(size=10, color="green"),
        xanchor="right",
        yanchor="bottom",
    )

    for results, (mw, mwh), color in zip(all_results, configs, PARETO_COLORS):
        profits = [r["profit"] for r in results]
        co2s = [r["co2"] for r in results]
        labels = [
            f"λ=({r['lambda_profit']:.2f}, {r['lambda_co2']:.2f})" for r in results
        ]
        fig.add_trace(
            go.Scatter(
                x=profits,
                y=co2s,
                mode="lines+markers",
                name=f"{mw} MW / {mwh} MWh",
                text=labels,
                hovertemplate="%{text}<br>Profit: %{x:.2f} DKK<br>CO₂: %{y:.4f} kg<extra></extra>",
                marker=dict(size=6, color=color),
                line=dict(color=color, width=1.5),
            )
        )

    fig.update_layout(
        xaxis_title="Profit (DKK)",
        yaxis_title="CO₂ Emissions (kg)",
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.075,
            xanchor="center",
            x=0.5,
        ),
        margin=dict(l=0, r=0, t=40, b=10),
    )

    out = Path("results/scenario_1/pareto.png")
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.show()
    fig.write_image(str(out))
    print(f"Pareto plot saved to {out}")


if __name__ == "__main__":
    profit_breakdowns: list[dict] = []
    co2_breakdowns: list[dict] = []
    all_pareto_results = []

    breakdown_path = Path("results/scenario_1/breakdown.xlsx")
    column_labels = config_labels(CONFIGS)

    for bat_mw, bat_mwh in CONFIGS:
        label = f"{bat_mw}mw_{bat_mwh}mwh"
        pareto_path = f"results/scenario_1/pareto_{label}.xlsx"

        print(f"\n{'=' * 60}")
        print(f"Asset: {bat_mw} MW / {bat_mwh} MWh")
        print(f"{'=' * 60}")

        m = Model3(
            start_date=START_DATE,
            end_date=END_DATE,
            bat_mw=bat_mw,
            bat_mwh=bat_mwh,
        )

        if Path(pareto_path).exists():
            print(f"  Loading cached Pareto data from {pareto_path}")
            pareto_results = load_pareto_excel(pareto_path)
            print("\n  Breakdown at Pareto extremes")
            first_bd, last_bd = get_extreme_breakdowns(m)
        else:
            print("\n  Pareto sweep (101 solves)")
            pareto_results, first_bd, last_bd = run_pareto(m)
            save_pareto_excel(pareto_results, pareto_path)

        profit_breakdowns.append(first_bd)
        co2_breakdowns.append(last_bd)
        all_pareto_results.append(pareto_results)

    print_breakdown_table(
        profit_breakdowns,
        column_labels,
        "Profit extreme (λ_profit=0.9999, λ_co2=0.0001)",
    )
    print_breakdown_table(
        co2_breakdowns,
        column_labels,
        "CO₂ extreme (λ_profit=0.0001, λ_co2=0.9999)",
    )

    save_breakdown_excel(
        [
            ("Profit extreme", profit_breakdowns),
            ("CO2 extreme", co2_breakdowns),
        ],
        CONFIGS,
        str(breakdown_path),
    )

    visualize_asset_pareto(all_pareto_results, CONFIGS)
