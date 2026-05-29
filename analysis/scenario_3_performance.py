"""Performance evaluation: Model 3 (perfect foresight) vs Model 4
(sequential, naive forecast prices) on a per-weight basis.

Reads the two cached Pareto frontiers, joins them on (lambda_profit,
lambda_co2), and computes profit and CO2 losses in both absolute and
relative terms. Outputs:
  * `results/scenario_3/performance/performance.xlsx` (per-weight table)
  * `results/scenario_3/performance/summary.txt` (aggregate metrics)
  * `results/scenario_3/performance/profit_loss.png`
  * `results/scenario_3/performance/co2_loss.png`
  * `results/scenario_3/performance/overview.png` (4-panel comparison)
"""

from pathlib import Path

import plotly.graph_objects as go
import polars as pl
from plotly.subplots import make_subplots

BASELINE_PATH = "results/model_3/model_3.xlsx"
FORECAST_PATH = "results/scenario_3/forecast_la2_101pts.xlsx"
OUT_DIR = Path("results/scenario_3/performance")


def _safe_rel(numerator: pl.Expr, denominator: pl.Expr) -> pl.Expr:
    """Relative change with sign-aware denominator. Returns null when the
    baseline is exactly zero (avoids division-by-zero and meaningless
    infinities)."""
    return (
        pl.when(denominator.abs() < 1e-12)
        .then(None)
        .otherwise(numerator / denominator.abs())
    )


def load_and_join() -> pl.DataFrame:
    base = pl.read_excel(BASELINE_PATH).rename(
        {"profit_dkk": "profit_baseline", "co2_kg": "co2_baseline"}
    )
    seq = pl.read_excel(FORECAST_PATH).rename(
        {"profit_dkk": "profit_seq", "co2_kg": "co2_seq"}
    )
    df = base.join(seq, on=["lambda_profit", "lambda_co2"], how="inner")

    df = df.with_columns(
        # Profit loss: how much profit the sequential model leaves on the
        # table relative to the perfect-foresight benchmark. Positive ==
        # sequential is worse.
        (pl.col("profit_baseline") - pl.col("profit_seq")).alias("profit_loss_abs"),
        # CO2 loss: extra emissions caused by sequential commitment.
        # Positive == sequential emits more.
        (pl.col("co2_seq") - pl.col("co2_baseline")).alias("co2_loss_abs"),
    ).with_columns(
        _safe_rel(pl.col("profit_loss_abs"), pl.col("profit_baseline")).alias(
            "profit_loss_rel"
        ),
        _safe_rel(pl.col("co2_loss_abs"), pl.col("co2_baseline")).alias("co2_loss_rel"),
    )
    return df


def write_summary(df: pl.DataFrame, path: Path) -> str:
    def stat(col: str) -> dict[str, float | None]:
        s = df[col].drop_nulls()
        if s.len() == 0:
            return {"mean": None, "median": None, "min": None, "max": None}
        return {
            "mean": float(s.mean()),
            "median": float(s.median()),
            "min": float(s.min()),
            "max": float(s.max()),
        }

    lines = ["Sequential (forecast) vs Model 3 (perfect foresight)", "=" * 60, ""]
    for metric, unit in [
        ("profit_loss_abs", "DKK"),
        ("profit_loss_rel", "fraction"),
        ("co2_loss_abs", "kg"),
        ("co2_loss_rel", "fraction"),
    ]:
        s = stat(metric)
        lines.append(f"{metric} ({unit})")
        for k in ("mean", "median", "min", "max"):
            v = s[k]
            if v is None:
                lines.append(f"  {k:>6}:    n/a")
            elif "rel" in metric:
                lines.append(f"  {k:>6}: {v:>10.4%}")
            else:
                lines.append(f"  {k:>6}: {v:>12.2f}")
        lines.append("")

    text = "\n".join(lines)
    path.write_text(text)
    return text


def plot_loss_panel(df: pl.DataFrame, out_path: Path) -> None:
    """4-panel comparison: profit & CO2 levels (top row) and absolute &
    relative losses (bottom row), all against lambda_profit."""
    x = df["lambda_profit"].to_list()

    fig = make_subplots(
        rows=2,
        cols=2,
        subplot_titles=(
            "Profit (DKK)",
            "CO₂ (kg)",
            "Absolute loss (sequential vs baseline)",
            "Relative loss (sequential vs baseline)",
        ),
        vertical_spacing=0.12,
        horizontal_spacing=0.10,
    )

    fig.add_trace(
        go.Scatter(
            x=x,
            y=df["profit_baseline"].to_list(),
            mode="lines+markers",
            name="Model 3 (perfect foresight)",
            line=dict(color="seagreen", width=1.5),
            marker=dict(size=5, color="seagreen"),
            legendgroup="baseline",
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=df["profit_seq"].to_list(),
            mode="lines+markers",
            name="Model 4 (forecast prices)",
            line=dict(color="darkorange", width=1.5),
            marker=dict(size=5, color="darkorange"),
            legendgroup="forecast",
        ),
        row=1,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=df["co2_baseline"].to_list(),
            mode="lines+markers",
            line=dict(color="seagreen", width=1.5),
            marker=dict(size=5, color="seagreen"),
            legendgroup="baseline",
            showlegend=False,
        ),
        row=1,
        col=2,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=df["co2_seq"].to_list(),
            mode="lines+markers",
            line=dict(color="darkorange", width=1.5),
            marker=dict(size=5, color="darkorange"),
            legendgroup="forecast",
            showlegend=False,
        ),
        row=1,
        col=2,
    )

    fig.add_trace(
        go.Scatter(
            x=x,
            y=df["profit_loss_abs"].to_list(),
            mode="lines+markers",
            name="Profit loss (DKK)",
            line=dict(color="steelblue", width=1.5),
            marker=dict(size=5, color="steelblue"),
        ),
        row=2,
        col=1,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=df["co2_loss_abs"].to_list(),
            mode="lines+markers",
            name="CO₂ loss (kg)",
            line=dict(color="crimson", width=1.5),
            marker=dict(size=5, color="crimson"),
            yaxis="y3",
        ),
        row=2,
        col=1,
    )

    fig.add_trace(
        go.Scatter(
            x=x,
            y=[None if v is None else 100 * v for v in df["profit_loss_rel"].to_list()],
            mode="lines+markers",
            name="Profit loss (%)",
            line=dict(color="steelblue", width=1.5, dash="dot"),
            marker=dict(size=5, color="steelblue"),
        ),
        row=2,
        col=2,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=[None if v is None else 100 * v for v in df["co2_loss_rel"].to_list()],
            mode="lines+markers",
            name="CO₂ loss (%)",
            line=dict(color="crimson", width=1.5, dash="dot"),
            marker=dict(size=5, color="crimson"),
        ),
        row=2,
        col=2,
    )

    fig.update_xaxes(title_text="λ_profit", autorange="reversed")
    fig.update_yaxes(title_text="DKK", row=1, col=1)
    fig.update_yaxes(title_text="kg", row=1, col=2)
    fig.update_yaxes(title_text="Absolute loss", row=2, col=1)
    fig.update_yaxes(title_text="Relative loss [%]", row=2, col=2)
    fig.update_layout(
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.08,
            xanchor="center",
            x=0.5,
        ),
        margin=dict(l=0, r=0, t=40, b=10),
        height=700,
    )
    fig.write_image(str(out_path))
    print(f"Overview saved to {out_path}")


def plot_single_loss(
    df: pl.DataFrame,
    abs_col: str,
    rel_col: str,
    abs_title: str,
    rel_title: str,
    color: str,
    out_path: Path,
) -> None:
    """Dual-y plot of absolute vs relative loss for one objective."""
    x = df["lambda_profit"].to_list()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Scatter(
            x=x,
            y=df[abs_col].to_list(),
            mode="lines+markers",
            name=abs_title,
            line=dict(color=color, width=1.5),
            marker=dict(size=5, color=color),
        ),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=x,
            y=[None if v is None else 100 * v for v in df[rel_col].to_list()],
            mode="lines+markers",
            name=rel_title,
            line=dict(color=color, width=1.5, dash="dot"),
            marker=dict(size=5, color=color, symbol="square"),
        ),
        secondary_y=True,
    )
    fig.update_xaxes(title_text="λ_profit", autorange="reversed")
    fig.update_yaxes(title_text=abs_title, secondary_y=False)
    fig.update_yaxes(title_text=f"{rel_title} [%]", secondary_y=True)
    fig.update_layout(
        template="plotly_white",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.1,
            xanchor="center",
            x=0.5,
        ),
        margin=dict(l=0, r=0, t=30, b=10),
    )
    fig.write_image(str(out_path))
    print(f"Saved {out_path}")


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    df = load_and_join()
    table_path = OUT_DIR / "performance.xlsx"
    df.write_excel(table_path)
    print(f"Per-weight table saved to {table_path}")

    summary = write_summary(df, OUT_DIR / "summary.txt")
    print()
    print(summary)

    plot_single_loss(
        df,
        abs_col="profit_loss_abs",
        rel_col="profit_loss_rel",
        abs_title="Profit loss (DKK)",
        rel_title="Profit loss",
        color="steelblue",
        out_path=OUT_DIR / "profit_loss.png",
    )
    plot_single_loss(
        df,
        abs_col="co2_loss_abs",
        rel_col="co2_loss_rel",
        abs_title="CO₂ loss (kg)",
        rel_title="CO₂ loss",
        color="crimson",
        out_path=OUT_DIR / "co2_loss.png",
    )
    plot_loss_panel(df, OUT_DIR / "overview.png")


if __name__ == "__main__":
    main()
