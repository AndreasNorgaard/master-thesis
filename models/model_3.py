from pathlib import Path

import holidays
import plotly.graph_objects as go
import polars as pl
import pyomo.environ as pyo
import xlsxwriter

from data.energi_data_service import EnergiDataServiceAPIClient


class Model3:
    # EUR -> DKK conversion for datasets that only publish EUR prices
    # (matches the implicit factor in AfrrReservesNordic: 7.4588).
    EUR_TO_DKK = 7.4588

    def __init__(
        self,
        start_date: str,
        end_date: str,
        lambda_profit: float = 0.5,
        lambda_co2: float = 0.5,
    ):
        self.results_file_path = "results/model_3.xlsx"

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
        self.tariff_prod = 15.5

        # Endurance requirements (hours)
        self.E_FCRD = 1.0 / 3.0  # 20 minutes
        self.E_FCRN = 1.0
        self.E_aFRR = 4.0
        self.E_mFRR = 0.25  # 15 minutes

        # Multi-objective weights (must sum to 1)
        self.lambda_profit = lambda_profit
        self.lambda_co2 = lambda_co2

        # Load data
        self.load_data()

    def load_data(self, write_to_file: bool = True):
        client = EnergiDataServiceAPIClient(
            start_date=self.start_date,
            end_date=self.end_date,
            price_area="DK2",
        )
        df_da = client.day_ahead_prices(write_to_file=False)
        df_co2 = client.co2_emissions(write_to_file=False)
        df_ffr = client.ffr_capacity(write_to_file=False)
        df_fcr_nd = client.fcr_nd_capacity(write_to_file=False)
        df_afrr = client.afrr_capacity(write_to_file=False)
        df_mfrr = client.mfrr_capacity(write_to_file=False)
        self.df, self.df_hourly, self.df_block = self.create_dataset(
            df_da, df_co2, df_ffr, df_fcr_nd, df_afrr, df_mfrr
        )
        if write_to_file:
            out = Path("data/prepared/model_3.xlsx")
            out.parent.mkdir(parents=True, exist_ok=True)
            wb = xlsxwriter.Workbook(str(out))
            self.df.write_excel(wb, worksheet="quarterly")
            self.df_hourly.write_excel(wb, worksheet="hourly")
            self.df_block.write_excel(wb, worksheet="blocks")
            wb.close()

    def create_dataset(
        self,
        df_da: pl.DataFrame,
        df_co2: pl.DataFrame,
        df_ffr: pl.DataFrame,
        df_fcr_nd: pl.DataFrame,
        df_afrr: pl.DataFrame,
        df_mfrr: pl.DataFrame,
    ) -> tuple[pl.DataFrame, pl.DataFrame, pl.DataFrame]:
        """
        Builds the quarterly base dataset (DA prices, CO2, tariffs) plus the
        hourly and 4-hour-block reserve capacity price tables aligned to the
        same horizon as the quarterly index.
        """
        # ---------- Quarterly base (identical to Model 2) ----------
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

        dk_holidays = holidays.Denmark(
            years=range(int(self.start_date[:4]), int(self.end_date[:4]) + 1)
        )
        holiday_dates = set(dk_holidays.keys())
        df = (
            df.with_columns(
                pl.col("TimeDK").dt.hour().alias("hour"),
                pl.col("TimeDK").dt.weekday().alias("weekday"),
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

        # ---------- Hourly skeleton ----------
        # Anchor hours to every 4th quarter so model.hour=h <-> quarters 4(h-1)+1..4h.
        # Capacity prices are looked up by floor-rounding the anchor to the
        # market's natural hourly UTC boundary.
        df_hours = (
            df.with_row_index("__qidx")
            .filter(pl.col("__qidx") % 4 == 0)
            .select(
                pl.col("TimeUTC").alias("TimeUTC_anchor"),
                pl.col("TimeDK"),
                pl.col("TimeUTC").dt.truncate("1h").alias("TimeUTC"),
            )
        )

        # FFR (DK2-only dataset, hourly, DKK price already provided)
        if df_ffr.height > 0:
            ffr = df_ffr.with_columns(
                pl.col("HourUTC").str.to_datetime("%Y-%m-%dT%H:%M:%S", strict=False),
                pl.col("FFR_PriceDKK").cast(pl.Float64),
            ).select(
                pl.col("HourUTC").alias("TimeUTC"),
                pl.col("FFR_PriceDKK").alias("P_FFR"),
            )
            df_hours = df_hours.join(ffr, on="TimeUTC", how="left")
        else:
            df_hours = df_hours.with_columns(pl.lit(0.0).alias("P_FFR"))

        # FCR-N/D: pivot to one column per product, take AuctionType = "Total"
        if df_fcr_nd.height > 0:
            fcr = (
                df_fcr_nd.filter(
                    (pl.col("PriceArea") == "DK2") & (pl.col("AuctionType") == "Total")
                )
                .with_columns(
                    pl.col("HourUTC").str.to_datetime(
                        "%Y-%m-%dT%H:%M:%S", strict=False
                    ),
                    (pl.col("PriceTotalEUR").cast(pl.Float64) * self.EUR_TO_DKK).alias(
                        "PriceDKK"
                    ),
                )
                .select(["HourUTC", "ProductName", "PriceDKK"])
                .pivot(values="PriceDKK", index="HourUTC", on="ProductName")
                .rename({"HourUTC": "TimeUTC"})
            )
            for src, dst in [
                ("FCR-D upp", "P_FCRD_up"),
                ("FCR-D ned", "P_FCRD_down"),
                ("FCR-N", "P_FCRN"),
            ]:
                if src in fcr.columns:
                    fcr = fcr.rename({src: dst})
                else:
                    fcr = fcr.with_columns(pl.lit(0.0).alias(dst))
            df_hours = df_hours.join(
                fcr.select(["TimeUTC", "P_FCRD_up", "P_FCRD_down", "P_FCRN"]),
                on="TimeUTC",
                how="left",
            )
        else:
            df_hours = df_hours.with_columns(
                pl.lit(0.0).alias("P_FCRD_up"),
                pl.lit(0.0).alias("P_FCRD_down"),
                pl.lit(0.0).alias("P_FCRN"),
            )

        # mFRR (hourly, DKK prices)
        if df_mfrr.height > 0:
            mfrr = df_mfrr.with_columns(
                pl.col("TimeUTC").str.to_datetime("%Y-%m-%dT%H:%M:%S", strict=False),
                pl.col("UpPriceDKK").cast(pl.Float64),
                pl.col("DownPriceDKK").cast(pl.Float64),
            ).select(
                "TimeUTC",
                pl.col("UpPriceDKK").alias("P_mFRR_up"),
                pl.col("DownPriceDKK").alias("P_mFRR_down"),
            )
            df_hours = df_hours.join(mfrr, on="TimeUTC", how="left")
        else:
            df_hours = df_hours.with_columns(
                pl.lit(0.0).alias("P_mFRR_up"),
                pl.lit(0.0).alias("P_mFRR_down"),
            )

        # Fill gaps with 0 (no clearing -> zero revenue, never blocks bidding)
        for col in [
            "P_FFR",
            "P_FCRD_up",
            "P_FCRD_down",
            "P_FCRN",
            "P_mFRR_up",
            "P_mFRR_down",
        ]:
            df_hours = df_hours.with_columns(pl.col(col).fill_null(0.0))

        # ---------- 4-hour block skeleton ----------
        # Anchor blocks to every 16th quarter so model.block=b <-> quarters
        # 16(b-1)+1..16b. The aFRR price for block b is the price of the market
        # block containing the anchor's start time.
        df_blocks = (
            df.with_row_index("__qidx")
            .filter(pl.col("__qidx") % 16 == 0)
            .select(
                pl.col("TimeUTC").alias("TimeUTC_anchor"),
                pl.col("TimeDK"),
                pl.col("TimeUTC").dt.truncate("4h").alias("TimeUTC"),
            )
        )

        if df_afrr.height > 0:
            afrr = (
                df_afrr.filter(pl.col("PriceArea") == "DK2")
                .with_columns(
                    pl.col("TimeUTC").str.to_datetime(
                        "%Y-%m-%dT%H:%M:%S", strict=False
                    ),
                    pl.col("UpPriceDKK").cast(pl.Float64),
                    pl.col("DownPriceDKK").cast(pl.Float64),
                )
                .select(
                    "TimeUTC",
                    pl.col("UpPriceDKK").alias("P_aFRR_up"),
                    pl.col("DownPriceDKK").alias("P_aFRR_down"),
                )
            )
            df_blocks = df_blocks.join(afrr, on="TimeUTC", how="left")
        else:
            df_blocks = df_blocks.with_columns(
                pl.lit(0.0).alias("P_aFRR_up"),
                pl.lit(0.0).alias("P_aFRR_down"),
            )

        for col in ["P_aFRR_up", "P_aFRR_down"]:
            df_blocks = df_blocks.with_columns(pl.col(col).fill_null(0.0))

        return df, df_hours, df_blocks

    @staticmethod
    def quarter_to_hour(q: int) -> int:
        return (q - 1) // 4 + 1

    @staticmethod
    def quarter_to_block(q: int) -> int:
        return (q - 1) // 16 + 1

    # ----------------------- Objective -----------------------

    def equation_1(self, model):
        """
        Objective: weighted sum of (day-ahead profit + reserve capacity revenue)
        and CO2 emissions.
        """
        da_profit = sum(
            self.delta_t
            * (
                model.da_sell[q] * (model.da_price[q] - model.tariff_prod)
                - model.da_buy[q] * (model.da_price[q] + model.tariff_cons[q])
                - model.cycle_cost * (model.da_buy[q] + model.da_sell[q])
            )
            for q in model.quarters
        )
        hourly_rev = sum(
            model.ffr[h] * model.p_ffr[h]
            + model.fcrd_up[h] * model.p_fcrd_up[h]
            + model.fcrd_down[h] * model.p_fcrd_down[h]
            + model.fcrn[h] * model.p_fcrn[h]
            + model.mfrr_up[h] * model.p_mfrr_up[h]
            + model.mfrr_down[h] * model.p_mfrr_down[h]
            for h in model.hours
        )
        block_rev = sum(
            model.afrr_up[b] * model.p_afrr_up[b]
            + model.afrr_down[b] * model.p_afrr_down[b]
            for b in model.blocks
        )
        co2 = sum(
            self.delta_t * model.gamma[q] * (model.da_buy[q] - model.da_sell[q])
            for q in model.quarters
        )
        return (
            model.lambda_profit * (da_profit + hourly_rev + block_rev)
            - model.lambda_co2 * co2
        )

    # --------------- Constraints inherited from Model 2 ---------------

    def equation_2(self, model, q):
        return (0, model.bat_mw)

    def equation_3(self, model, q):
        return (0, model.bat_discharge_eff * model.bat_mw)

    def equation_4(self, model, q):
        return (model.soc_min, model.soc_max)

    def equation_5(self, model, q):
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
        if q != model.quarters.last():
            return pyo.Constraint.Skip
        return model.soc[q] == model.soc_initial

    def equation_8(self, model, d):
        quarters_in_day = range(96 * (d - 1) + 1, 96 * d + 1)
        return (
            sum(self.delta_t * model.da_buy[q] for q in quarters_in_day)
            <= model.n_cycles * model.bat_mwh
        )

    # ----------------- New Model 3 constraints -----------------

    def equation_ler_fcrd_up(self, model, q):
        """LER buffer with FCR-D up + aFRR up + mFRR up."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_min
            + model.fcrd_up[h] * self.E_FCRD
            + model.afrr_up[b] * self.E_aFRR
            + model.mfrr_up[h] * self.E_mFRR
            <= model.soc[q]
        )

    def equation_ler_fcrd_down(self, model, q):
        """LER buffer with FCR-D down + aFRR down + mFRR down."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_max
            - model.fcrd_down[h] * self.E_FCRD
            - model.afrr_down[b] * self.E_aFRR
            - model.mfrr_down[h] * self.E_mFRR
            >= model.soc[q]
        )

    def equation_ler_fcrn_up(self, model, q):
        """LER buffer with FCR-N + aFRR up + mFRR up."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_min
            + model.fcrn[h] * self.E_FCRN
            + model.afrr_up[b] * self.E_aFRR
            + model.mfrr_up[h] * self.E_mFRR
            <= model.soc[q]
        )

    def equation_ler_fcrn_down(self, model, q):
        """LER buffer with FCR-N + aFRR down + mFRR down."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_max
            - model.fcrn[h] * self.E_FCRN
            - model.afrr_down[b] * self.E_aFRR
            - model.mfrr_down[h] * self.E_mFRR
            >= model.soc[q]
        )

    def equation_power_discharge(self, model, q):
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.da_sell[q]
            + model.ffr[h]
            + model.fcrd_up[h]
            + model.fcrn[h]
            + model.afrr_up[b]
            + model.mfrr_up[h]
            <= model.bat_discharge_eff * model.bat_mw
        )

    def equation_power_charge(self, model, q):
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.da_buy[q]
            + model.fcrd_down[h]
            + model.fcrn[h]
            + model.afrr_down[b]
            + model.mfrr_down[h]
            <= model.bat_mw
        )

    # ----------------------- Solve -----------------------

    def solve(self):
        model = pyo.ConcreteModel()
        Q = len(self.df)
        D = Q // 96
        H = Q // 4
        B = Q // 16
        model.quarters = pyo.RangeSet(1, Q)
        model.days = pyo.RangeSet(1, D)
        model.hours = pyo.RangeSet(1, H)
        model.blocks = pyo.RangeSet(1, B)

        # Scalar parameters
        model.bat_mw = pyo.Param(initialize=self.bat_mw)
        model.bat_mwh = pyo.Param(initialize=self.bat_mwh)
        model.bat_charge_eff = pyo.Param(initialize=self.bat_charge_eff)
        model.bat_discharge_eff = pyo.Param(initialize=self.bat_discharge_eff)
        model.soc_initial = pyo.Param(initialize=self.soc_initial * self.bat_mwh)
        model.soc_min = pyo.Param(initialize=self.soc_min)
        model.soc_max = pyo.Param(initialize=self.soc_max)
        model.lam = pyo.Param(initialize=self.soc_quarterly_loss)
        model.tariff_prod = pyo.Param(initialize=self.tariff_prod)
        model.cycle_cost = pyo.Param(initialize=self.cycle_cost)
        model.n_cycles = pyo.Param(initialize=self.n_cycles)
        model.lambda_profit = pyo.Param(initialize=self.lambda_profit)
        model.lambda_co2 = pyo.Param(initialize=self.lambda_co2)

        # Quarterly time series
        model.tariff_cons = pyo.Param(
            model.quarters,
            initialize={q: self.df["tariff_cons"][q - 1] for q in range(1, Q + 1)},
        )
        model.da_price = pyo.Param(
            model.quarters,
            initialize={q: self.df["DayAheadPriceDKK"][q - 1] for q in range(1, Q + 1)},
        )
        model.gamma = pyo.Param(
            model.quarters,
            initialize={q: self.df["CO2Emission"][q - 1] for q in range(1, Q + 1)},
        )

        # Hourly capacity prices (truncated/extended to H rows)
        H_avail = self.df_hourly.height

        def _hourly(col, h):
            idx = h - 1
            if idx < H_avail:
                v = self.df_hourly[col][idx]
                return float(v) if v is not None else 0.0
            return 0.0

        model.p_ffr = pyo.Param(
            model.hours, initialize={h: _hourly("P_FFR", h) for h in range(1, H + 1)}
        )
        model.p_fcrd_up = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRD_up", h) for h in range(1, H + 1)},
        )
        model.p_fcrd_down = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRD_down", h) for h in range(1, H + 1)},
        )
        model.p_fcrn = pyo.Param(
            model.hours, initialize={h: _hourly("P_FCRN", h) for h in range(1, H + 1)}
        )
        model.p_mfrr_up = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_mFRR_up", h) for h in range(1, H + 1)},
        )
        model.p_mfrr_down = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_mFRR_down", h) for h in range(1, H + 1)},
        )

        # Block capacity prices
        B_avail = self.df_block.height

        def _block(col, b):
            idx = b - 1
            if idx < B_avail:
                v = self.df_block[col][idx]
                return float(v) if v is not None else 0.0
            return 0.0

        model.p_afrr_up = pyo.Param(
            model.blocks,
            initialize={b: _block("P_aFRR_up", b) for b in range(1, B + 1)},
        )
        model.p_afrr_down = pyo.Param(
            model.blocks,
            initialize={b: _block("P_aFRR_down", b) for b in range(1, B + 1)},
        )

        # Decision variables
        model.da_buy = pyo.Var(model.quarters, bounds=self.equation_2)
        model.da_sell = pyo.Var(model.quarters, bounds=self.equation_3)
        model.soc = pyo.Var(model.quarters, bounds=self.equation_4)

        model.ffr = pyo.Var(model.hours, bounds=(0, None))
        model.fcrd_up = pyo.Var(model.hours, bounds=(0, None))
        model.fcrd_down = pyo.Var(model.hours, bounds=(0, None))
        model.fcrn = pyo.Var(model.hours, bounds=(0, None))
        model.mfrr_up = pyo.Var(model.hours, bounds=(0, None))
        model.mfrr_down = pyo.Var(model.hours, bounds=(0, None))
        model.afrr_up = pyo.Var(model.blocks, bounds=(0, None))
        model.afrr_down = pyo.Var(model.blocks, bounds=(0, None))

        # Objective
        model.objective = pyo.Objective(expr=self.equation_1(model), sense=pyo.maximize)

        # Constraints (Model 2 inherited)
        model.equation_5 = pyo.Constraint(model.quarters, rule=self.equation_5)
        model.equation_6 = pyo.Constraint(model.quarters, rule=self.equation_6)
        model.equation_7 = pyo.Constraint(model.quarters, rule=self.equation_7)
        model.equation_8 = pyo.Constraint(model.days, rule=self.equation_8)

        # New Model 3 constraints
        model.eq_ler_fcrd_up = pyo.Constraint(
            model.quarters, rule=self.equation_ler_fcrd_up
        )
        model.eq_ler_fcrd_down = pyo.Constraint(
            model.quarters, rule=self.equation_ler_fcrd_down
        )
        model.eq_ler_fcrn_up = pyo.Constraint(
            model.quarters, rule=self.equation_ler_fcrn_up
        )
        model.eq_ler_fcrn_down = pyo.Constraint(
            model.quarters, rule=self.equation_ler_fcrn_down
        )
        model.eq_power_discharge = pyo.Constraint(
            model.quarters, rule=self.equation_power_discharge
        )
        model.eq_power_charge = pyo.Constraint(
            model.quarters, rule=self.equation_power_charge
        )

        solver = pyo.SolverFactory("glpk")
        solver.solve(model)
        return model

    # ----------------------- Reporting -----------------------

    def _extract_objectives(self, model) -> tuple[float, float]:
        Q = len(self.df)
        H = Q // 4
        B = Q // 16

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
        da_profit = da_revenue - prod_tariff - da_cost - cons_tariff - degradation

        ffr_rev = sum(
            pyo.value(model.ffr[h]) * pyo.value(model.p_ffr[h]) for h in range(1, H + 1)
        )
        fcrd_up_rev = sum(
            pyo.value(model.fcrd_up[h]) * pyo.value(model.p_fcrd_up[h])
            for h in range(1, H + 1)
        )
        fcrd_down_rev = sum(
            pyo.value(model.fcrd_down[h]) * pyo.value(model.p_fcrd_down[h])
            for h in range(1, H + 1)
        )
        fcrn_rev = sum(
            pyo.value(model.fcrn[h]) * pyo.value(model.p_fcrn[h])
            for h in range(1, H + 1)
        )
        mfrr_up_rev = sum(
            pyo.value(model.mfrr_up[h]) * pyo.value(model.p_mfrr_up[h])
            for h in range(1, H + 1)
        )
        mfrr_down_rev = sum(
            pyo.value(model.mfrr_down[h]) * pyo.value(model.p_mfrr_down[h])
            for h in range(1, H + 1)
        )
        afrr_up_rev = sum(
            pyo.value(model.afrr_up[b]) * pyo.value(model.p_afrr_up[b])
            for b in range(1, B + 1)
        )
        afrr_down_rev = sum(
            pyo.value(model.afrr_down[b]) * pyo.value(model.p_afrr_down[b])
            for b in range(1, B + 1)
        )
        reserve_rev = (
            ffr_rev
            + fcrd_up_rev
            + fcrd_down_rev
            + fcrn_rev
            + mfrr_up_rev
            + mfrr_down_rev
            + afrr_up_rev
            + afrr_down_rev
        )

        profit = da_profit + reserve_rev

        co2 = sum(
            self.delta_t
            * pyo.value(model.gamma[q])
            * (pyo.value(model.da_buy[q]) - pyo.value(model.da_sell[q]))
            for q in model.quarters
        )

        print(f"  Day-ahead Revenue:   {da_revenue:>12.2f} DKK")
        print(f"  Production Tariffs:  {-prod_tariff:>12.2f} DKK")
        print(f"  Day-ahead Cost:      {-da_cost:>12.2f} DKK")
        print(f"  Consumption Tariffs: {-cons_tariff:>12.2f} DKK")
        print(f"  Degradation Cost:    {-degradation:>12.2f} DKK")
        print(f"  DA Net Profit:       {da_profit:>12.2f} DKK")
        print(f"  FFR Revenue:         {ffr_rev:>12.2f} DKK")
        print(f"  FCR-D up Revenue:    {fcrd_up_rev:>12.2f} DKK")
        print(f"  FCR-D down Revenue:  {fcrd_down_rev:>12.2f} DKK")
        print(f"  FCR-N Revenue:       {fcrn_rev:>12.2f} DKK")
        print(f"  aFRR up Revenue:     {afrr_up_rev:>12.2f} DKK")
        print(f"  aFRR down Revenue:   {afrr_down_rev:>12.2f} DKK")
        print(f"  mFRR up Revenue:     {mfrr_up_rev:>12.2f} DKK")
        print(f"  mFRR down Revenue:   {mfrr_down_rev:>12.2f} DKK")
        print(f"  Reserve Revenue:     {reserve_rev:>12.2f} DKK")
        print(f"  Total Profit:        {profit:>12.2f} DKK")
        print(f"  CO2 Emissions:       {co2:>12.4f} kg")

        breakdown = {
            "da_revenue": da_revenue,
            "prod_tariff": -prod_tariff,
            "da_cost": -da_cost,
            "cons_tariff": -cons_tariff,
            "degradation": -degradation,
            "da_profit": da_profit,
            "ffr_revenue": ffr_rev,
            "fcrd_up_revenue": fcrd_up_rev,
            "fcrd_down_revenue": fcrd_down_rev,
            "fcrn_revenue": fcrn_rev,
            "afrr_up_revenue": afrr_up_rev,
            "afrr_down_revenue": afrr_down_rev,
            "mfrr_up_revenue": mfrr_up_rev,
            "mfrr_down_revenue": mfrr_down_rev,
            "reserve_revenue": reserve_rev,
            "profit": profit,
            "co2": co2,
        }
        return profit, co2, breakdown

    def visualize_profit_distribution(
        self,
        breakdown: dict,
        lambda_profit: float,
        lambda_co2: float,
    ) -> None:
        """Plot a waterfall chart of the profit breakdown for one weight pair."""
        labels = [
            "Day-ahead Revenue",
            "Production Tariffs",
            "Day-ahead Cost",
            "Consumption Tariffs",
            "Degradation Cost",
            "FFR Revenue",
            "FCR-D up Revenue",
            "FCR-D down Revenue",
            "FCR-N Revenue",
            "aFRR up Revenue",
            "aFRR down Revenue",
            "mFRR up Revenue",
            "mFRR down Revenue",
            "Profit",
        ]
        values = [
            breakdown["da_revenue"],
            breakdown["prod_tariff"],
            breakdown["da_cost"],
            breakdown["cons_tariff"],
            breakdown["degradation"],
            breakdown["ffr_revenue"],
            breakdown["fcrd_up_revenue"],
            breakdown["fcrd_down_revenue"],
            breakdown["fcrn_revenue"],
            breakdown["afrr_up_revenue"],
            breakdown["afrr_down_revenue"],
            breakdown["mfrr_up_revenue"],
            breakdown["mfrr_down_revenue"],
            None,
        ]
        measures = ["relative"] * (len(labels) - 1) + ["total"]

        fig = go.Figure(
            go.Waterfall(
                orientation="v",
                measure=measures,
                x=labels,
                y=values,
                text=[
                    f"{v:.2f}" if v is not None else f"{breakdown['profit']:.2f}"
                    for v in values
                ],
                textposition="outside",
                increasing={"marker": {"color": "green"}},
                decreasing={"marker": {"color": "red"}},
                totals={"marker": {"color": "steelblue"}},
                connector={"line": {"color": "grey", "width": 1}},
            )
        )
        fig.update_layout(
            title=f"Profit Distribution - Model 3 (λ_profit={lambda_profit}, λ_co2={lambda_co2})",
            yaxis_title="DKK",
            plot_bgcolor="aliceblue",
            showlegend=False,
        )
        fig.show()
        fig.write_image("results/model_3_profit.png")

    def pareto_frontier(self) -> list[dict]:
        """Solve the model 11 times across evenly spaced weight combinations and return results."""
        weight_pairs = [(round(1.0 - i * 0.1, 1), round(i * 0.1, 1)) for i in range(11)]
        results = []

        for i, (lp, lc) in enumerate(weight_pairs):
            self.lambda_profit = lp
            self.lambda_co2 = lc
            print(f"\nλ_profit={lp:.1f}, λ_co2={lc:.1f}")
            solved = self.solve()
            profit, co2, breakdown = self._extract_objectives(solved)
            results.append(
                {"lambda_profit": lp, "lambda_co2": lc, "profit": profit, "co2": co2}
            )
            if i == 0:
                self.visualize_profit_distribution(breakdown, lp, lc)

        return results

    def visualize_pareto_frontier(self, results: list[dict]):
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
                marker=dict(size=8, color="seagreen"),
                line=dict(color="seagreen", width=1.5),
            )
        )
        fig.update_layout(
            title="Model 3 Pareto Frontier — Profit vs. CO₂ Emissions",
            xaxis_title="Profit (DKK)",
            yaxis_title="CO₂ Emissions (kg)",
            template="plotly_white",
        )
        fig.show()
        fig.write_image("results/model_3_pareto_frontier.png")


if __name__ == "__main__":
    m = Model3(
        start_date="2026-04-01",
        end_date="2026-04-30",
    )
    pareto_results = m.pareto_frontier()
    m.visualize_pareto_frontier(pareto_results)
