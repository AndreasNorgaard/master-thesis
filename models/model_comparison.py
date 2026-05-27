from pathlib import Path

import plotly.graph_objects as go
import polars as pl


def load_results(path: str) -> pl.DataFrame:
    return pl.read_excel(path)


def compare_pareto_frontiers():
    df2 = load_results("results/model_2.xlsx")
    df3 = load_results("results/model_3.xlsx")

    profits2 = df2["profit_dkk"].to_list()
    co2s2 = df2["co2_kg"].to_list()
    labels2 = [
        f"λ=({r['lambda_profit']:.2f}, {r['lambda_co2']:.2f})"
        for r in df2.iter_rows(named=True)
    ]

    profits3 = df3["profit_dkk"].to_list()
    co2s3 = df3["co2_kg"].to_list()
    labels3 = [
        f"λ=({r['lambda_profit']:.2f}, {r['lambda_co2']:.2f})"
        for r in df3.iter_rows(named=True)
    ]

    all_profits = profits2 + profits3
    all_co2s = co2s2 + co2s3

    x_max = max(all_profits) * 1.05
    y_min = min(all_co2s) * 1.05 if min(all_co2s) < 0 else -1

    fig = go.Figure()

    # Green shaded region: profit > 0 and CO2 < 0
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

    # Model 2 curve (steelblue)
    fig.add_trace(
        go.Scatter(
            x=profits2,
            y=co2s2,
            mode="lines+markers",
            name="Model 2",
            text=labels2,
            hovertemplate="%{text}<br>Profit: %{x:.2f} DKK<br>CO₂: %{y:.4f} kg<extra></extra>",
            marker=dict(size=6, color="steelblue"),
            line=dict(color="steelblue", width=1.5),
        )
    )

    # Model 3 curve (seagreen)
    fig.add_trace(
        go.Scatter(
            x=profits3,
            y=co2s3,
            mode="lines+markers",
            name="Model 3",
            text=labels3,
            hovertemplate="%{text}<br>Profit: %{x:.2f} DKK<br>CO₂: %{y:.4f} kg<extra></extra>",
            marker=dict(size=6, color="seagreen"),
            line=dict(color="seagreen", width=1.5),
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

    out = Path("results/model_comparison_pareto.png")
    out.parent.mkdir(parents=True, exist_ok=True)
    fig.show()
    fig.write_image(str(out))
    print(f"Saved to {out}")


if __name__ == "__main__":
    compare_pareto_frontiers()
