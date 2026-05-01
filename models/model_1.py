import holidays
import plotly.graph_objects as go
import polars as pl
import pyomo.environ as pyo
from plotly.subplots import make_subplots

from data.energi_data_service import EnergiDataServiceAPIClient


class Model1:
    def __init__(self, start_date: str, end_date: str):
        self.results_file_path = "results/model_1.xlsx"

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
        self.cycle_cost = 13.0  # Degradation cost (EUR/MWh)
        self.n_cycles = 2  # Max full cycles per day (contractual limit)

        # Tariffs (Appendix B of thesis, DKK/MWh)
        # Production tariff (discharging): TSO (5.0 + 5.3) + DSO (5.2) = 15.5 DKK/MWh
        # Convert production tariff from DKK to EUR (using 1 EUR = 7.45 DKK)
        self.tariff_prod = 15.5 / 7.44  # EUR/MWh

        # Consumption tariff (charging): time-varying, computed in load_data()
        # τ_c_q = systemtarif + nettabstarif_q + DSO_q
        #   systemtarif   = 72 DKK/MWh (fixed)
        #   nettabstarif  = 1.42% × (P_spot_q + 26 DKK/MWh)
        #   DSO_q         = 30.4 DKK/MWh (00:00–06:00 weekdays and all weekends)
        #                 = 91.1 DKK/MWh (06:00–24:00 weekdays)

        # Load data
        self.load_data()

    def load_data(self):
        df_hourly = EnergiDataServiceAPIClient(
            start_date=self.start_date,
            end_date=self.end_date,
            price_area="DK2",
        ).day_ahead_prices(write_to_file=False)

        # Expand each hourly row into 4 identical 15-min quarters
        df_hourly = df_hourly.with_columns(
            pl.col("TimeDK").str.to_datetime("%Y-%m-%dT%H:%M:%S", strict=False)
        )
        df_quarters = df_hourly.select(["TimeDK", "DayAheadPriceDKK"]).with_columns(
            pl.lit(4).alias("repeat_count")
        )
        # Repeat each row 4 times to get 15-min resolution
        df_quarters = df_quarters.select(
            pl.col("TimeDK").repeat_by("repeat_count").explode(),
            pl.col("DayAheadPriceDKK").repeat_by("repeat_count").explode(),
        )

        # Compute time-varying consumption tariff per quarter
        # DSO B-høj: 91.1 DKK/MWh weekdays 06:00–24:00; 30.4 DKK/MWh otherwise
        # Public holidays are treated as weekends (30.4 DKK/MWh all day)
        dk_holidays = holidays.Denmark(
            years=range(int(self.start_date[:4]), int(self.end_date[:4]) + 1)
        )
        holiday_dates = set(dk_holidays.keys())
        df_quarters = (
            df_quarters.with_columns(
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
                    72.0 / 7.44
                    + 0.0142 * (pl.col("DayAheadPriceDKK") + (26.0 / 7.44))
                    + pl.col("dso_tariff")
                ).alias("tariff_cons")
            )
        )

        self.df = df_quarters

    def equation_1(self, model):
        """
        Objective Function: Maximize profit over the optimization period. (Eq. 1)
        """
        return sum(
            self.delta_t
            * (
                model.da_sell[q] * (model.da_price[q] - model.tariff_prod)
                - model.da_buy[q] * (model.da_price[q] + model.tariff_cons[q])
                - model.cycle_cost * (model.da_buy[q] + model.da_sell[q])
            )
            for q in model.quarters
        )

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
        model.da_price = pyo.Param(
            model.quarters,
            initialize={q: self.df["DayAheadPriceDKK"][q - 1] for q in range(1, Q + 1)},
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

    def calculate_profit(self, model):
        Q = len(self.df)
        revenue = sum(
            self.delta_t
            * pyo.value(model.da_sell[q])
            * (pyo.value(model.da_price[q]) - self.tariff_prod)
            for q in range(1, Q + 1)
        )
        cost_buy = sum(
            self.delta_t
            * pyo.value(model.da_buy[q])
            * (pyo.value(model.da_price[q]) + self.df["tariff_cons"][q - 1])
            for q in range(1, Q + 1)
        )
        degradation = sum(
            self.delta_t
            * self.cycle_cost
            * (pyo.value(model.da_buy[q]) + pyo.value(model.da_sell[q]))
            for q in range(1, Q + 1)
        )
        profit = revenue - cost_buy - degradation

        print(f"Revenue from selling:    {revenue:>10.2f} EUR")
        print(f"Cost of buying:          {cost_buy:>10.2f} EUR")
        print(f"Degradation cost:        {degradation:>10.2f} EUR")
        print(f"Net profit:              {profit:>10.2f} EUR")
        return profit

    def visualize_profit(self):
        pass

    def visualize_schedule(self, model):
        Q = len(self.df)
        times = self.df["TimeDK"].to_list()

        soc = [pyo.value(model.soc[q]) / self.bat_mwh for q in range(1, Q + 1)]
        da_buy = [pyo.value(model.da_buy[q]) for q in range(1, Q + 1)]
        da_sell = [-pyo.value(model.da_sell[q]) for q in range(1, Q + 1)]

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        # Shaded SoC feasible band
        soc_min_frac = self.soc_min / self.bat_mwh
        soc_max_frac = self.soc_max / self.bat_mwh
        fig.add_hrect(
            y0=soc_min_frac,
            y1=soc_max_frac,
            fillcolor="steelblue",
            opacity=0.15,
            layer="below",
            line_width=0,
            secondary_y=False,
        )

        # Charging bars (positive, red)
        fig.add_trace(
            go.Bar(
                x=times,
                y=da_buy,
                name="Day-ahead buy (Charge)",
                marker_color="red",
                opacity=0.85,
            ),
            secondary_y=True,
        )

        # Discharging bars (negative, green)
        fig.add_trace(
            go.Bar(
                x=times,
                y=da_sell,
                name="Day-ahead sell (Discharge)",
                marker_color="green",
                opacity=0.85,
            ),
            secondary_y=True,
        )

        # State of Charge line
        fig.add_trace(
            go.Scatter(
                x=times,
                y=soc,
                name="State of Charge",
                mode="lines",
                line=dict(color="black", width=1.5),
            ),
            secondary_y=False,
        )

        fig.update_layout(
            title="Visualization of Production Schedule from Model 1",
            barmode="overlay",
            plot_bgcolor="white",
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.15,
                xanchor="center",
                x=0.5,
            ),
            xaxis=dict(showgrid=False),
        )
        fig.update_yaxes(
            title_text="SoC [%]",
            range=[0, 1],
            secondary_y=False,
            showgrid=True,
            gridcolor="lightgrey",
        )
        fig.update_yaxes(
            title_text="Effect [MW]",
            range=[-(self.bat_mw + 0.2), self.bat_mw + 0.2],
            secondary_y=True,
            showgrid=False,
        )

        fig.show()


if __name__ == "__main__":
    m = Model1(start_date="2026-04-01", end_date="2026-04-30")
    solved = m.solve()
    m.calculate_profit(solved)
    m.visualize_schedule(solved)
