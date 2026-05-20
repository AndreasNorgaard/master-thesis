from pathlib import Path

import plotly.graph_objects as go
import polars as pl
import xlsxwriter

from models.model_3 import Model3

BAT_MW = 2
BAT_MWH = 4

SCENARIOS = [
    ("January", "2026-01-01", "2026-01-30"),
    ("April", "2026-04-01", "2026-04-30"),
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


def run_breakdown_solve(m: Model3) -> dict:
    """Solve at λ_profit=1.0, λ_co2=0.0 and return the breakdown dict."""
    m.lambda_profit = 1.0
    m.lambda_co2 = 0.0
    solved = m.solve()
    _, _, breakdown = m._extract_objectives(solved)
    return breakdown


def save_breakdown_excel(
    breakdowns: list[dict],
    scenario_labels: list[str],
    out_path: str,
) -> None:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    wb = xlsxwriter.Workbook(str(out))
    ws = wb.add_worksheet("Profit Breakdown")

    header_fmt = wb.add_format(
        {"bold": True, "align": "center", "border": 1, "bg_color": "#D9D9D9"}
    )
    label_fmt = wb.add_format({"border": 1})
    num_fmt = wb.add_format({"num_format": "#,##0.00", "border": 1})
    bold_num_fmt = wb.add_format(
        {"bold": True, "num_format": "#,##0.00", "border": 1, "bg_color": "#F2F2F2"}
    )
    bold_label_fmt = wb.add_format({"bold": True, "border": 1, "bg_color": "#F2F2F2"})

    # Header row
    ws.write(0, 0, "", header_fmt)
    for col, label in enumerate(scenario_labels, start=1):
        ws.write(0, col, label, header_fmt)

    # Data rows
    for row_idx, (label, key) in enumerate(BREAKDOWN_ROWS, start=1):
        is_total = label == "Profit"
        lf = bold_label_fmt if is_total else label_fmt
        nf = bold_num_fmt if is_total else num_fmt
        ws.write(row_idx, 0, label, lf)
        for col, breakdown in enumerate(breakdowns, start=1):
            value = _row_value(breakdown, key)
            ws.write(row_idx, col, value, nf)

    ws.set_column(0, 0, 28)
    ws.set_column(1, len(scenario_labels), 18)

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


def run_pareto(m: Model3) -> list[dict]:
    """Run 101 evenly-spaced weight pairs and collect profit + CO2."""
    weight_pairs = [(round(1.0 - i * 0.01, 2), round(i * 0.01, 2)) for i in range(101)]
    results = []
    for lp, lc in weight_pairs:
        m.lambda_profit = lp
        m.lambda_co2 = lc
        print(f"  λ_profit={lp:.2f}, λ_co2={lc:.2f}")
        solved = m.solve()
        profit, co2, _ = m._extract_objectives(solved)
        results.append(
            {"lambda_profit": lp, "lambda_co2": lc, "profit": profit, "co2": co2}
        )
    return results


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


PARETO_COLORS = ["steelblue", "darkorange"]


def visualize_price_pareto(
    all_results: list[list[dict]],
    scenario_labels: list[str],
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

    for results, label, color in zip(all_results, scenario_labels, PARETO_COLORS):
        profits = [r["profit"] for r in results]
        co2s = [r["co2"] for r in results]
        hover_labels = [
            f"λ=({r['lambda_profit']:.2f}, {r['lambda_co2']:.2f})" for r in results
        ]
        fig.add_trace(
            go.Scatter(
                x=profits,
                y=co2s,
                mode="lines+markers",
                name=label,
                text=hover_labels,
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

    out = Path("results/price_scenarios_pareto.png")
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.show()
    fig.write_image(str(out))
    print(f"Pareto plot saved to {out}")


if __name__ == "__main__":
    breakdowns = []
    all_pareto_results = []

    breakdown_path = Path("results/price_scenarios_breakdown.xlsx")

    for name, start_date, end_date in SCENARIOS:
        label = name.lower()
        pareto_path = f"results/price_scenarios_pareto_{label}.xlsx"

        print(f"\n{'=' * 60}")
        print(f"Scenario: {name}  ({start_date} → {end_date})")
        print(f"Battery:  {BAT_MW} MW / {BAT_MWH} MWh")
        print(f"{'=' * 60}")

        if Path(pareto_path).exists():
            print(f"  Loading cached Pareto data from {pareto_path}")
            pareto_results = load_pareto_excel(pareto_path)
        else:
            m = Model3(
                start_date=start_date,
                end_date=end_date,
                bat_mw=BAT_MW,
                bat_mwh=BAT_MWH,
            )
            if not breakdown_path.exists():
                print("\n[1/2] Breakdown solve (λ_profit=1.0, λ_co2=0.0)")
                breakdown = run_breakdown_solve(m)
                breakdowns.append(breakdown)

            print("\n[2/2] Pareto sweep (101 solves)")
            pareto_results = run_pareto(m)
            save_pareto_excel(pareto_results, pareto_path)

        all_pareto_results.append(pareto_results)

    scenario_labels = [name for name, _, _ in SCENARIOS]

    if not breakdown_path.exists() and breakdowns:
        save_breakdown_excel(breakdowns, scenario_labels, str(breakdown_path))

    visualize_price_pareto(all_pareto_results, scenario_labels)
