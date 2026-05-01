import holidays
import plotly.graph_objects as go
import polars as pl
import pyomo.environ as pyo

from data.energi_data_service import EnergiDataServiceAPIClient


class Model2:
    def __init__(
        self,
        start_date: str,
        end_date: str,
        lambda_profit: float = 0.5,
        lambda_co2: float = 0.5,
    ):
        self.results_file_path = "results/model_2.xlsx"

        # Set dates
        self.start_date = start_date
        self.end_date = end_date

        # Set BESS configuration
        self.bat_mw = 2  # Max charging/discharging power (MW)
        self.bat_mwh = 4  # Energy storage capacity (MWh)
        self.bat_charge_eff = 0.99  # Charging efficiency (η_c)
        self.bat_discharge_eff = 0.87  # Discharging efficiency (η_d)
        self.soc_initial = 0.50  # Initial state of charge (fraction of bat_mwh)
        self.soc_min = 0.05 * self.bat_mwh  # Minimum state of charge (MWh)
        self.soc_max = 0.95 * self.bat_mwh  # Maximum state of charge (MWh)
        self.soc_quarterly_loss = 0.025 / (30 * 24 * 4)  # 0.025% per month
        self.delta_t = 0.25  # Length of each time interval (hours)
        self.cycle_cost = 13.0 * 7.44  # Degradation cost (DKK/MWh)
        self.n_cycles = 2  # Max full cycles per day (contractual limit)

        # Tariffs (Appendix B of thesis, DKK/MWh)
        # Production tariff (discharging): TSO (5.0 + 5.3) + DSO (5.2) = 15.5 DKK/MWh
        self.tariff_prod = 15.5

        # Consumption tariff (charging): time-varying, computed in create_dataset()
        # τ_c_q = systemtarif + nettabstarif_q + DSO_q
        #   systemtarif   = 72 DKK/MWh (fixed)
        #   nettabstarif  = 1.42% × (P_spot_q + 26 DKK/MWh)
        #   DSO_q         = 30.4 DKK/MWh (00:00–06:00 weekdays and all weekends)
        #                 = 91.1 DKK/MWh (06:00–24:00 weekdays)

        # Multi-objective weights (must sum to 1)
        self.lambda_profit = lambda_profit
        self.lambda_co2 = lambda_co2

        # Load data
        self.load_data()

    def load_data(self):
        client = EnergiDataServiceAPIClient(
            start_date=self.start_date,
            end_date=self.end_date,
            price_area="DK2",
        )
        df_da = client.day_ahead_prices(write_to_file=False)
        df_co2 = client.co2_emissions(write_to_file=False)
        self.df = self.create_dataset(df_da, df_co2)

    def create_dataset(self, df_da: pl.DataFrame, df_co2: pl.DataFrame) -> pl.DataFrame:
        """
        Aligns CO2 emissions (5-min) to the 15-min DA price resolution and
        computes the time-varying consumption tariff.
        """
        df_quarters = df_da.select(
            ["TimeDK", "TimeUTC", "DayAheadPriceDKK"]
        ).with_columns(
            pl.col("TimeDK").str.to_datetime("%Y-%m-%dT%H:%M:%S", strict=False),
            pl.col("TimeUTC").str.to_datetime("%Y-%m-%dT%H:%M:%S", strict=False),
        )
        df_co2 = df_co2.with_columns(
            pl.col("Minutes5UTC").str.to_datetime(
                format="%Y-%m-%dT%H:%M:%S", strict=False
            )
        )

        # Resample CO2 from 5-min to 15-min by averaging
        df_co2_15min = (
            df_co2.sort("Minutes5UTC")
            .group_by_dynamic("Minutes5UTC", every="15m")
            .agg(pl.col("CO2Emission").mean())
            .rename({"Minutes5UTC": "TimeUTC"})
        )

        df = (
            df_quarters.join(df_co2_15min, on="TimeUTC", how="left")
            .sort("TimeUTC")
            .with_columns(pl.col("CO2Emission").forward_fill().backward_fill())
        )

        # Compute time-varying consumption tariff per quarter
        # DSO B-høj: 91.1 DKK/MWh weekdays 06:00–24:00; 30.4 DKK/MWh otherwise
        # Public holidays are treated as weekends (30.4 DKK/MWh all day)
        dk_holidays = holidays.Denmark(
            years=range(int(self.start_date[:4]), int(self.end_date[:4]) + 1)
        )
        holiday_dates = set(dk_holidays.keys())
        df = (
            df.with_columns(
                pl.col("TimeDK").dt.hour().alias("hour"),
                pl.col("TimeDK").dt.weekday().alias("weekday"),  # 0=Mon … 6=Sun
                pl.col("TimeDK")
                .dt.date()
                .is_in(list(holiday_dates))
                .alias("is_holiday"),
            )
            .with_columns(
                pl.when(
                    (pl.col("weekday") < 5)
                    & (pl.col("hour") >= 6)
                    & (~pl.col("is_holiday"))
                )
                .then(91.1)
                .otherwise(30.4)
                .alias("dso_tariff")
            )
            .with_columns(
                (
                    72.0
                    + 0.0142 * (pl.col("DayAheadPriceDKK") + 26.0)
                    + pl.col("dso_tariff")
                ).alias("tariff_cons")
            )
        )

        return df

    def equation_1(self, model):
        """
        Objective Function: Weighted sum of normalized profit and CO2 emissions. (Eq. 1)
        """
        profit = sum(
            self.delta_t
            * (
                model.da_sell[q] * (model.da_price[q] - model.tariff_prod)
                - model.da_buy[q] * (model.da_price[q] + model.tariff_cons[q])
                - model.cycle_cost * (model.da_buy[q] + model.da_sell[q])
            )
            for q in model.quarters
        )
        co2 = sum(
            self.delta_t * model.gamma[q] * (model.da_buy[q] - model.da_sell[q])
            for q in model.quarters
        )
        return model.lambda_profit * profit - model.lambda_co2 * co2

    def equation_2(self, model, q):
        """
        Constraint: Charging power limit (Eq. 2)
        0 ≤ DA_buy_q ≤ BatMW
        """
        return (0, model.bat_mw)

    def equation_3(self, model, q):
        """
        Constraint: Discharging power limit (Eq. 3)
        0 ≤ DA_sell_q ≤ η_d * BatMW
        """
        return (0, model.bat_discharge_eff * model.bat_mw)

    def equation_4(self, model, q):
        """
        Constraint: State of Charge bounds (Eq. 4)
        SoC_min ≤ SoC_q ≤ SoC_max
        """
        return (model.soc_min, model.soc_max)

    def equation_5(self, model, q):
        """
        Constraint: State of Charge dynamics (Eq. 5)
        SoC_q = (1 - λ) * SoC_{q-1} + Δt * (η_c * DA_buy_q - DA_sell_q / η_d)
        """
        if q == model.quarters.first():
            return pyo.Constraint.Skip
        return model.soc[q] == (
            (1 - model.lam) * model.soc[q - 1]
            + self.delta_t
            * (
                model.bat_charge_eff * model.da_buy[q]
                - model.da_sell[q] / model.bat_discharge_eff
            )
        )

    def equation_6(self, model, q):
        """
        Constraint: Initial State of Charge (Eq. 6)
        SoC_1 = (1 - λ) * SoC_init + Δt * (η_c * DA_buy_1 - DA_sell_1 / η_d)
        """
        if q != model.quarters.first():
            return pyo.Constraint.Skip
        return model.soc[q] == (
            (1 - model.lam) * model.soc_initial
            + self.delta_t
            * (
                model.bat_charge_eff * model.da_buy[q]
                - model.da_sell[q] / model.bat_discharge_eff
            )
        )

    def equation_7(self, model, q):
        """
        Constraint: Terminal State of Charge equals initial SoC (Eq. 7)
        SoC_Q = SoC_init
        """
        if q != model.quarters.last():
            return pyo.Constraint.Skip
        return model.soc[q] == model.soc_initial

    def equation_8(self, model, d):
        """
        Constraint: Daily cycle limit (Eq. 8)
        Σ_{q ∈ Q_D_d} Δt * DA_buy_q ≤ N_cycles * BatMWh
        """
        quarters_in_day = range(96 * (d - 1) + 1, 96 * d + 1)
        return (
            sum(self.delta_t * model.da_buy[q] for q in quarters_in_day)
            <= model.n_cycles * model.bat_mwh
        )

    def solve(self):
        model = pyo.ConcreteModel()
        Q = len(self.df)
        D = Q // 96
        model.quarters = pyo.RangeSet(1, Q)
        model.days = pyo.RangeSet(1, D)

        # Parameters
        model.bat_mw = pyo.Param(initialize=self.bat_mw)
        model.bat_mwh = pyo.Param(initialize=self.bat_mwh)
        model.bat_charge_eff = pyo.Param(initialize=self.bat_charge_eff)
        model.bat_discharge_eff = pyo.Param(initialize=self.bat_discharge_eff)
        model.soc_initial = pyo.Param(initialize=self.soc_initial * self.bat_mwh)
        model.soc_min = pyo.Param(initialize=self.soc_min)
        model.soc_max = pyo.Param(initialize=self.soc_max)
        model.lam = pyo.Param(initialize=self.soc_quarterly_loss)
        model.tariff_prod = pyo.Param(initialize=self.tariff_prod)
        model.tariff_cons = pyo.Param(
            model.quarters,
            initialize={q: self.df["tariff_cons"][q - 1] for q in range(1, Q + 1)},
        )
        model.cycle_cost = pyo.Param(initialize=self.cycle_cost)
        model.n_cycles = pyo.Param(initialize=self.n_cycles)
        model.lambda_profit = pyo.Param(initialize=self.lambda_profit)
        model.lambda_co2 = pyo.Param(initialize=self.lambda_co2)
        model.da_price = pyo.Param(
            model.quarters,
            initialize={q: self.df["DayAheadPriceDKK"][q - 1] for q in range(1, Q + 1)},
        )
        model.gamma = pyo.Param(
            model.quarters,
            initialize={q: self.df["CO2Emission"][q - 1] for q in range(1, Q + 1)},
        )

        # Decision Variables: bounds defined by equations (2), (3), (4)
        model.da_buy = pyo.Var(model.quarters, bounds=self.equation_2)
        model.da_sell = pyo.Var(model.quarters, bounds=self.equation_3)
        model.soc = pyo.Var(model.quarters, bounds=self.equation_4)

        # Objective Function: Equation (1)
        model.objective = pyo.Objective(expr=self.equation_1(model), sense=pyo.maximize)

        # Constraints: Equations (5)–(8)
        model.equation_5 = pyo.Constraint(model.quarters, rule=self.equation_5)
        model.equation_6 = pyo.Constraint(model.quarters, rule=self.equation_6)
        model.equation_7 = pyo.Constraint(model.quarters, rule=self.equation_7)
        model.equation_8 = pyo.Constraint(model.days, rule=self.equation_8)

        # Solve
        solver = pyo.SolverFactory("glpk")
        solver.solve(model)

        return model

    def _extract_objectives(self, model) -> tuple[float, float]:
        """Extract actual (unweighted) profit and CO2 from a solved model, printing a breakdown."""
        Q = len(self.df)
        da_sell_vals = [pyo.value(model.da_sell[q]) for q in range(1, Q + 1)]
        da_buy_vals = [pyo.value(model.da_buy[q]) for q in range(1, Q + 1)]
        da_price_vals = [pyo.value(model.da_price[q]) for q in range(1, Q + 1)]
        tariff_cons_vals = self.df["tariff_cons"].to_list()

        da_revenue = self.delta_t * sum(
            s * p for s, p in zip(da_sell_vals, da_price_vals)
        )
        prod_tariff = self.delta_t * sum(s * self.tariff_prod for s in da_sell_vals)
        da_cost = self.delta_t * sum(b * p for b, p in zip(da_buy_vals, da_price_vals))
        cons_tariff = self.delta_t * sum(
            b * t for b, t in zip(da_buy_vals, tariff_cons_vals)
        )
        degradation = (
            self.delta_t
            * self.cycle_cost
            * sum(b + s for b, s in zip(da_buy_vals, da_sell_vals))
        )
        profit = da_revenue - prod_tariff - da_cost - cons_tariff - degradation

        co2 = sum(
            self.delta_t
            * pyo.value(model.gamma[q])
            * (pyo.value(model.da_buy[q]) - pyo.value(model.da_sell[q]))
            for q in model.quarters
        )

        print(f"  Day-ahead Revenue:   {da_revenue:>10.2f} DKK")
        print(f"  Production Tariffs:  {-prod_tariff:>10.2f} DKK")
        print(f"  Day-ahead Cost:      {-da_cost:>10.2f} DKK")
        print(f"  Consumption Tariffs: {-cons_tariff:>10.2f} DKK")
        print(f"  Degradation Cost:    {-degradation:>10.2f} DKK")
        print(f"  Net Profit:          {profit:>10.2f} DKK")
        print(f"  CO2 Emissions:       {co2:>10.4f} kg")

        return profit, co2

    def pareto_frontier(self) -> list[dict]:
        """Solve the model 11 times across evenly spaced weight combinations and return results."""
        weight_pairs = [(round(1.0 - i * 0.1, 1), round(i * 0.1, 1)) for i in range(11)]
        results = []

        for lp, lc in weight_pairs:
            self.lambda_profit = lp
            self.lambda_co2 = lc
            print(f"\nλ_profit={lp:.1f}, λ_co2={lc:.1f}")
            solved = self.solve()
            profit, co2 = self._extract_objectives(solved)
            results.append(
                {"lambda_profit": lp, "lambda_co2": lc, "profit": profit, "co2": co2}
            )

        return results

    def visualize_pareto_frontier(self, results: list[dict]):
        """Plot the Pareto frontier from the results of pareto_frontier()."""
        profits = [r["profit"] for r in results]
        co2s = [r["co2"] for r in results]
        labels = [
            f"λ=({r['lambda_profit']:.1f}, {r['lambda_co2']:.1f})" for r in results
        ]

        fig = go.Figure(
            go.Scatter(
                x=profits,
                y=co2s,
                mode="lines+markers+text",
                text=labels,
                textposition="top right",
                textfont=dict(size=10),
                marker=dict(size=8, color="steelblue"),
                line=dict(color="steelblue", width=1.5),
            )
        )
        fig.update_layout(
            title="Pareto Frontier — Profit vs. CO₂ Emissions",
            xaxis_title="Profit (DKK)",
            yaxis_title="CO₂ Emissions (kg)",
            template="plotly_white",
        )
        fig.show()
        fig.write_image("results/model_2_pareto_frontier.png")


if __name__ == "__main__":
    m = Model2(
        start_date="2026-04-01",
        end_date="2026-04-30",
    )
    pareto_results = m.pareto_frontier()
    m.visualize_pareto_frontier(pareto_results)
