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
        self.bat_mw = 1  # Max charging/discharging power (MW)
        self.bat_mwh = 2  # Energy storage capacity (MWh)
        self.bat_charge_eff = 0.99  # Charging efficiency
        self.bat_discharge_eff = 0.86  # Discharging efficiency
        self.soc_initial = 0.50  # Initial state of charge (fraction of Bat^MWh)
        self.soc_min = 0.10 * self.bat_mwh  # Minimum state of charge (MWh)
        self.soc_max = 0.90 * self.bat_mwh  # Maximum state of charge (MWh)
        self.soc_quarterly_loss = 0.0002083  # Self-discharge rate per period
        self.delta_t = 0.25  # Length of each time interval (hours)
        self.cycle_cost = 0.0  # Degradation cost (EUR/MWh)

        # Set tariffs
        self.tariff_cons = 129.3  # Consumption tariff (EUR/MWh)
        self.tariff_prod = 1.04  # Production tariff (EUR/MWh)

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
        Aligns day-ahead prices (15-min) and CO2 emissions (5-min) to a common
        15-min resolution and joins them on TimeUTC.
        """
        df_da = df_da.with_columns(
            pl.col("TimeUTC").str.to_datetime(format="%Y-%m-%dT%H:%M:%S", strict=False)
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

        return (
            df_da.join(df_co2_15min, on="TimeUTC", how="left")
            .sort("TimeUTC")
            .with_columns(pl.col("CO2Emission").forward_fill().backward_fill())
        )

    def equation_1(self, model):
        """
        Objective Function: Weighted sum of profit and CO2 emissions.
        """
        profit = sum(
            self.delta_t
            * (
                model.da_sell[q] * (model.da_price[q] - model.tariff_prod)
                - model.da_buy[q] * (model.da_price[q] + model.tariff_cons)
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
        Constraint: Charging limit
        """
        return (0, model.bat_mw)

    def equation_3(self, model, q):
        """
        Constraint: Discharging limit
        """
        return (0, model.bat_discharge_eff * model.bat_mw)

    def equation_4(self, model, q):
        """
        Constraint: State of Charge Limits
        """
        return (model.soc_min, model.soc_max)

    def equation_5(self, model, q):
        """
        Constraint: State of Charge Dynamics
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
        Constraint: Initial State of Charge
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
        Constraint: Terminal State of Charge
        """
        if q != model.quarters.last():
            return pyo.Constraint.Skip
        return model.soc[q] == model.soc_initial

    def solve(self):
        model = pyo.ConcreteModel()
        Q = len(self.df)
        model.quarters = pyo.RangeSet(1, Q)

        # Parameters
        model.bat_mw = pyo.Param(initialize=self.bat_mw)
        model.bat_mwh = pyo.Param(initialize=self.bat_mwh)
        model.bat_charge_eff = pyo.Param(initialize=self.bat_charge_eff)
        model.bat_discharge_eff = pyo.Param(initialize=self.bat_discharge_eff)
        model.soc_initial = pyo.Param(initialize=self.soc_initial * self.bat_mwh)
        model.soc_min = pyo.Param(initialize=self.soc_min)
        model.soc_max = pyo.Param(initialize=self.soc_max)
        model.lam = pyo.Param(initialize=self.soc_quarterly_loss)
        model.tariff_cons = pyo.Param(initialize=self.tariff_cons)
        model.tariff_prod = pyo.Param(initialize=self.tariff_prod)
        model.cycle_cost = pyo.Param(initialize=self.cycle_cost)
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

        # Decision Variables: Bounds defined by equations (2), (3), (4)
        model.da_buy = pyo.Var(model.quarters, bounds=self.equation_2)
        model.da_sell = pyo.Var(model.quarters, bounds=self.equation_3)
        model.soc = pyo.Var(model.quarters, bounds=self.equation_4)

        # Objective Function: Equation (1)
        model.objective = pyo.Objective(expr=self.equation_1(model), sense=pyo.maximize)

        # Constraints: Equations (5), (6), (7)
        model.equation_5 = pyo.Constraint(model.quarters, rule=self.equation_5)
        model.equation_6 = pyo.Constraint(model.quarters, rule=self.equation_6)
        model.equation_7 = pyo.Constraint(model.quarters, rule=self.equation_7)

        # Solve
        solver = pyo.SolverFactory("glpk")
        solver.solve(model)

        return model

    def _extract_objectives(self, model) -> tuple[float, float]:
        """Extract actual (unweighted) profit and CO2 from a solved model."""
        profit = sum(
            self.delta_t
            * (
                pyo.value(model.da_sell[q])
                * (pyo.value(model.da_price[q]) - self.tariff_prod)
                - pyo.value(model.da_buy[q])
                * (pyo.value(model.da_price[q]) + self.tariff_cons)
                - self.cycle_cost
                * (pyo.value(model.da_buy[q]) + pyo.value(model.da_sell[q]))
            )
            for q in model.quarters
        )
        co2 = sum(
            self.delta_t
            * pyo.value(model.gamma[q])
            * (pyo.value(model.da_buy[q]) - pyo.value(model.da_sell[q]))
            for q in model.quarters
        )
        return profit, co2

    def pareto_frontier(self) -> list[dict]:
        """Solve the model 11 times across evenly spaced weight combinations and return results."""
        weight_pairs = [(round(1.0 - i * 0.1, 1), round(i * 0.1, 1)) for i in range(11)]
        results = []

        for lp, lc in weight_pairs:
            self.lambda_profit = lp
            self.lambda_co2 = lc
            solved = self.solve()
            profit, co2 = self._extract_objectives(solved)
            results.append(
                {"lambda_profit": lp, "lambda_co2": lc, "profit": profit, "co2": co2}
            )
            print(
                f"λ_profit={lp:.1f}, λ_co2={lc:.1f} → Profit={profit:.2f} DKK, CO2={co2:.4f} kg"
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

    def calculate_profit(self):
        pass

    def visualize_profit(self):
        pass

    def visualize_schedule(self):
        pass


if __name__ == "__main__":
    m = Model2(
        start_date="2026-01-01",
        end_date="2026-01-31",
    )
    pareto_results = m.pareto_frontier()
    m.visualize_pareto_frontier(pareto_results)
