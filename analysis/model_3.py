from pathlib import Path

import holidays
import plotly.graph_objects as go
import polars as pl
import pyomo.environ as pyo
import xlsxwriter
from plotly.subplots import make_subplots

from data.energi_data_service import EnergiDataServiceAPIClient


class Model3:
    # EUR -> DKK conversion for datasets that only publish EUR prices
    # (matches the implicit factor in AfrrReservesNordic: 7.4588).
    EUR_TO_DKK = 7.4588

    def __init__(
        self,
        start_date: str,
        end_date: str,
        bat_mw: float = 2,
        bat_mwh: float = 4,
    ):
        self.results_file_path = "results/model_3.xlsx"

        # Set dates
        self.start_date = start_date
        self.end_date = end_date

        # Set BESS configuration
        self.bat_mw = bat_mw  # Max charging/discharging power (MW)
        self.bat_mwh = bat_mwh  # Energy storage capacity (MWh)
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
        if self.start_date == "2026-01-01":
            self.tariff_prod = 16.2  # Winter tariff
        else:
            self.tariff_prod = 15.5  # Summer tariff

        # Endurance requirements (hours)
        self.E_FCRD = 1.0 / 3.0  # 20 minutes
        self.E_FCRN = 1.0
        self.E_aFRR = 4.0
        self.E_mFRR = 0.25  # 15 minutes

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
        if self.start_date == "2026-01-01":
            dso_tariff_expr = (
                pl.when((pl.col("weekday") < 5) & (~pl.col("is_holiday")))
                .then(
                    pl.when(pl.col("hour") < 6)
                    .then(30.9)
                    .when(pl.col("hour") < 21)
                    .then(185.3)
                    .otherwise(92.6)
                )
                .otherwise(pl.when(pl.col("hour") < 6).then(30.9).otherwise(92.6))
            )
        else:
            dso_tariff_expr = (
                pl.when(
                    (pl.col("weekday") < 5)
                    & (pl.col("hour") >= 6)
                    & (~pl.col("is_holiday"))
                )
                .then(91.1)
                .otherwise(30.4)
            )

        df = (
            df.with_columns(
                pl.col("TimeDK").dt.hour().alias("hour"),
                pl.col("TimeDK").dt.weekday().alias("weekday"),
                pl.col("TimeDK")
                .dt.date()
                .is_in(list(holiday_dates))
                .alias("is_holiday"),
            )
            .with_columns(dso_tariff_expr.alias("dso_tariff"))
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

        # FCR-N/D: pivot to one column per (product, auction).
        # Auctions: "D-1 early" (E) and "D-1 late" (L).
        fcr_cols = [
            "P_FCRD_up_E",
            "P_FCRD_up_L",
            "P_FCRD_down_E",
            "P_FCRD_down_L",
            "P_FCRN_E",
            "P_FCRN_L",
        ]
        if df_fcr_nd.height > 0:
            fcr = (
                df_fcr_nd.filter(
                    (pl.col("PriceArea") == "DK2")
                    & (pl.col("AuctionType").is_in(["D-1 early", "D-1 late"]))
                )
                .with_columns(
                    pl.col("HourUTC").str.to_datetime(
                        "%Y-%m-%dT%H:%M:%S", strict=False
                    ),
                    (pl.col("PriceTotalEUR").cast(pl.Float64) * self.EUR_TO_DKK).alias(
                        "PriceDKK"
                    ),
                    pl.when(pl.col("AuctionType") == "D-1 early")
                    .then(pl.lit("E"))
                    .otherwise(pl.lit("L"))
                    .alias("AuctionTag"),
                )
                .with_columns(
                    (pl.col("ProductName") + "__" + pl.col("AuctionTag")).alias("Key")
                )
                .select(["HourUTC", "Key", "PriceDKK"])
                .pivot(values="PriceDKK", index="HourUTC", on="Key")
                .rename({"HourUTC": "TimeUTC"})
            )
            rename_map = {
                "FCR-D upp__E": "P_FCRD_up_E",
                "FCR-D upp__L": "P_FCRD_up_L",
                "FCR-D ned__E": "P_FCRD_down_E",
                "FCR-D ned__L": "P_FCRD_down_L",
                "FCR-N__E": "P_FCRN_E",
                "FCR-N__L": "P_FCRN_L",
            }
            for src, dst in rename_map.items():
                if src in fcr.columns:
                    fcr = fcr.rename({src: dst})
                else:
                    fcr = fcr.with_columns(pl.lit(0.0).alias(dst))
            df_hours = df_hours.join(
                fcr.select(["TimeUTC", *fcr_cols]),
                on="TimeUTC",
                how="left",
            )
        else:
            df_hours = df_hours.with_columns(*[pl.lit(0.0).alias(c) for c in fcr_cols])

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
            *fcr_cols,
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
            + model.fcrd_up_E[h] * model.p_fcrd_up_E[h]
            + model.fcrd_up_L[h] * model.p_fcrd_up_L[h]
            + model.fcrd_down_E[h] * model.p_fcrd_down_E[h]
            + model.fcrd_down_L[h] * model.p_fcrd_down_L[h]
            + model.fcrn_E[h] * model.p_fcrn_E[h]
            + model.fcrn_L[h] * model.p_fcrn_L[h]
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
        """LER buffer with FCR-D up (E+L) + aFRR up + mFRR up."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_min
            + (model.fcrd_up_E[h] + model.fcrd_up_L[h]) * self.E_FCRD
            + model.afrr_up[b] * self.E_aFRR
            + model.mfrr_up[h] * self.E_mFRR
            <= model.soc[q]
        )

    def equation_ler_fcrd_down(self, model, q):
        """LER buffer with FCR-D down (E+L) + aFRR down + mFRR down."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_max
            - (model.fcrd_down_E[h] + model.fcrd_down_L[h]) * self.E_FCRD
            - model.afrr_down[b] * self.E_aFRR
            - model.mfrr_down[h] * self.E_mFRR
            >= model.soc[q]
        )

    def equation_ler_fcrn_up(self, model, q):
        """LER buffer with FCR-N (E+L) + aFRR up + mFRR up."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_min
            + (model.fcrn_E[h] + model.fcrn_L[h]) * self.E_FCRN
            + model.afrr_up[b] * self.E_aFRR
            + model.mfrr_up[h] * self.E_mFRR
            <= model.soc[q]
        )

    def equation_ler_fcrn_down(self, model, q):
        """LER buffer with FCR-N (E+L) + aFRR down + mFRR down."""
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.soc_max
            - (model.fcrn_E[h] + model.fcrn_L[h]) * self.E_FCRN
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
            + model.fcrd_up_E[h]
            + model.fcrd_up_L[h]
            + model.fcrn_E[h]
            + model.fcrn_L[h]
            + model.afrr_up[b]
            + model.mfrr_up[h]
            <= model.bat_discharge_eff * model.bat_mw
        )

    def equation_power_charge(self, model, q):
        h = self.quarter_to_hour(q)
        b = self.quarter_to_block(q)
        return (
            model.da_buy[q]
            + model.fcrd_down_E[h]
            + model.fcrd_down_L[h]
            + model.fcrn_E[h]
            + model.fcrn_L[h]
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
        model.p_fcrd_up_E = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRD_up_E", h) for h in range(1, H + 1)},
        )
        model.p_fcrd_up_L = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRD_up_L", h) for h in range(1, H + 1)},
        )
        model.p_fcrd_down_E = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRD_down_E", h) for h in range(1, H + 1)},
        )
        model.p_fcrd_down_L = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRD_down_L", h) for h in range(1, H + 1)},
        )
        model.p_fcrn_E = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRN_E", h) for h in range(1, H + 1)},
        )
        model.p_fcrn_L = pyo.Param(
            model.hours,
            initialize={h: _hourly("P_FCRN_L", h) for h in range(1, H + 1)},
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
        model.fcrd_up_E = pyo.Var(model.hours, bounds=(0, None))
        model.fcrd_up_L = pyo.Var(model.hours, bounds=(0, None))
        model.fcrd_down_E = pyo.Var(model.hours, bounds=(0, None))
        model.fcrd_down_L = pyo.Var(model.hours, bounds=(0, None))
        model.fcrn_E = pyo.Var(model.hours, bounds=(0, None))
        model.fcrn_L = pyo.Var(model.hours, bounds=(0, None))
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
        fcrd_up_E_rev = sum(
            pyo.value(model.fcrd_up_E[h]) * pyo.value(model.p_fcrd_up_E[h])
            for h in range(1, H + 1)
        )
        fcrd_up_L_rev = sum(
            pyo.value(model.fcrd_up_L[h]) * pyo.value(model.p_fcrd_up_L[h])
            for h in range(1, H + 1)
        )
        fcrd_down_E_rev = sum(
            pyo.value(model.fcrd_down_E[h]) * pyo.value(model.p_fcrd_down_E[h])
            for h in range(1, H + 1)
        )
        fcrd_down_L_rev = sum(
            pyo.value(model.fcrd_down_L[h]) * pyo.value(model.p_fcrd_down_L[h])
            for h in range(1, H + 1)
        )
        fcrn_E_rev = sum(
            pyo.value(model.fcrn_E[h]) * pyo.value(model.p_fcrn_E[h])
            for h in range(1, H + 1)
        )
        fcrn_L_rev = sum(
            pyo.value(model.fcrn_L[h]) * pyo.value(model.p_fcrn_L[h])
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
            + fcrd_up_E_rev
            + fcrd_up_L_rev
            + fcrd_down_E_rev
            + fcrd_down_L_rev
            + fcrn_E_rev
            + fcrn_L_rev
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

        # print(f"  Day-ahead Revenue:   {da_revenue:>12.2f} DKK")
        # print(f"  Production Tariffs:  {-prod_tariff:>12.2f} DKK")
        # print(f"  Day-ahead Cost:      {-da_cost:>12.2f} DKK")
        # print(f"  Consumption Tariffs: {-cons_tariff:>12.2f} DKK")
        # print(f"  Degradation Cost:    {-degradation:>12.2f} DKK")
        # print(f"  DA Net Profit:       {da_profit:>12.2f} DKK")
        # print(f"  FFR Revenue:           {ffr_rev:>12.2f} DKK")
        # print(f"  FCR-D up early Rev:    {fcrd_up_E_rev:>12.2f} DKK")
        # print(f"  FCR-D up late Rev:     {fcrd_up_L_rev:>12.2f} DKK")
        # print(f"  FCR-D down early Rev:  {fcrd_down_E_rev:>12.2f} DKK")
        # print(f"  FCR-D down late Rev:   {fcrd_down_L_rev:>12.2f} DKK")
        # print(f"  FCR-N early Revenue:   {fcrn_E_rev:>12.2f} DKK")
        # print(f"  FCR-N late Revenue:    {fcrn_L_rev:>12.2f} DKK")
        # print(f"  aFRR up Revenue:       {afrr_up_rev:>12.2f} DKK")
        # print(f"  aFRR down Revenue:     {afrr_down_rev:>12.2f} DKK")
        # print(f"  mFRR up Revenue:       {mfrr_up_rev:>12.2f} DKK")
        # print(f"  mFRR down Revenue:     {mfrr_down_rev:>12.2f} DKK")
        # print(f"  Reserve Revenue:       {reserve_rev:>12.2f} DKK")
        print(f"  Total Profit:          {profit:>12.2f} DKK")
        print(f"  CO2 Emissions:         {co2:>12.2f} kg")

        breakdown = {
            "da_revenue": da_revenue,
            "prod_tariff": -prod_tariff,
            "da_cost": -da_cost,
            "cons_tariff": -cons_tariff,
            "degradation": -degradation,
            "da_profit": da_profit,
            "ffr_revenue": ffr_rev,
            "fcrd_up_E_revenue": fcrd_up_E_rev,
            "fcrd_up_L_revenue": fcrd_up_L_rev,
            "fcrd_down_E_revenue": fcrd_down_E_rev,
            "fcrd_down_L_revenue": fcrd_down_L_rev,
            "fcrn_E_revenue": fcrn_E_rev,
            "fcrn_L_revenue": fcrn_L_rev,
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
        out_file: str = "results/model_3_profit.png",
    ) -> None:
        """Plot a waterfall chart of the profit breakdown for one weight pair."""
        labels = [
            "Day-ahead Revenue",
            "Production Tariffs",
            "Day-ahead Cost",
            "Consumption Tariffs",
            "Degradation Cost",
            "FFR Revenue",
            "FCR-D up early Revenue",
            "FCR-D up late Revenue",
            "FCR-D down early Revenue",
            "FCR-D down late Revenue",
            "FCR-N early Revenue",
            "FCR-N late Revenue",
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
            breakdown["fcrd_up_E_revenue"],
            breakdown["fcrd_up_L_revenue"],
            breakdown["fcrd_down_E_revenue"],
            breakdown["fcrd_down_L_revenue"],
            breakdown["fcrn_E_revenue"],
            breakdown["fcrn_L_revenue"],
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
            yaxis_title="DKK",
            plot_bgcolor="aliceblue",
            showlegend=False,
            margin=dict(l=0, r=0, t=20, b=10),
        )
        fig.show()
        fig.write_image(out_file)

    def visualize_schedule(
        self,
        model,
        out_file: str = "results/model_3_schedule.png",
    ) -> None:
        """Stacked production schedule: DA + all capacity-auction allocations + SoC."""
        Q = len(self.df)
        times = self.df["TimeDK"].to_list()

        def q_to_h(q: int) -> int:
            return (q - 1) // 4 + 1

        def q_to_b(q: int) -> int:
            return (q - 1) // 16 + 1

        soc = [pyo.value(model.soc[q]) / self.bat_mwh for q in range(1, Q + 1)]

        # Day-ahead (quarterly) — sign convention: charge positive, discharge negative
        da_buy = [pyo.value(model.da_buy[q]) for q in range(1, Q + 1)]
        da_sell = [-pyo.value(model.da_sell[q]) for q in range(1, Q + 1)]

        # Broadcast hourly reserves to quarterly index
        def hourly_q(var):
            return [pyo.value(var[q_to_h(q)]) for q in range(1, Q + 1)]

        def block_q(var):
            return [pyo.value(var[q_to_b(q)]) for q in range(1, Q + 1)]

        # Discharge-direction (negative bars)
        ffr_q = [-v for v in hourly_q(model.ffr)]
        fcrd_up_E_q = [-v for v in hourly_q(model.fcrd_up_E)]
        fcrd_up_L_q = [-v for v in hourly_q(model.fcrd_up_L)]
        fcrn_E_up_q = [-v for v in hourly_q(model.fcrn_E)]
        fcrn_L_up_q = [-v for v in hourly_q(model.fcrn_L)]
        mfrr_up_q = [-v for v in hourly_q(model.mfrr_up)]
        afrr_up_q = [-v for v in block_q(model.afrr_up)]

        # Charge-direction (positive bars)
        fcrd_down_E_q = hourly_q(model.fcrd_down_E)
        fcrd_down_L_q = hourly_q(model.fcrd_down_L)
        fcrn_E_down_q = hourly_q(model.fcrn_E)
        fcrn_L_down_q = hourly_q(model.fcrn_L)
        mfrr_down_q = hourly_q(model.mfrr_down)
        afrr_down_q = block_q(model.afrr_down)

        fig = make_subplots(specs=[[{"secondary_y": True}]])

        soc_min_frac = self.soc_min / self.bat_mwh
        soc_max_frac = self.soc_max / self.bat_mwh
        fig.add_hrect(
            y0=soc_min_frac,
            y1=soc_max_frac,
            fillcolor="steelblue",
            opacity=0.10,
            layer="below",
            line_width=0,
            secondary_y=False,
        )

        # Color palette — one hue per product family, shades per auction/direction
        discharge_traces = [
            ("Day-ahead sell", da_sell, "#525252"),
            ("FFR", ffr_q, "#d73027"),
            ("FCR-D up (early)", fcrd_up_E_q, "#f4a582"),
            ("FCR-D up (late)", fcrd_up_L_q, "#d6604d"),
            ("FCR-N up (early)", fcrn_E_up_q, "#fee08b"),
            ("FCR-N up (late)", fcrn_L_up_q, "#fdae61"),
            ("aFRR up", afrr_up_q, "#762a83"),
            ("mFRR up", mfrr_up_q, "#1b7837"),
        ]
        charge_traces = [
            ("Day-ahead buy", da_buy, "#969696"),
            ("FCR-D down (early)", fcrd_down_E_q, "#fdb863"),
            ("FCR-D down (late)", fcrd_down_L_q, "#e08214"),
            ("FCR-N down (early)", fcrn_E_down_q, "#ffffbf"),
            ("FCR-N down (late)", fcrn_L_down_q, "#fee090"),
            ("aFRR down", afrr_down_q, "#9970ab"),
            ("mFRR down", mfrr_down_q, "#7fbc41"),
        ]

        for name, y, color in discharge_traces + charge_traces:
            fig.add_trace(
                go.Bar(
                    x=times,
                    y=y,
                    name=name,
                    marker_color=color,
                    marker_line_width=0,
                    opacity=0.65,
                ),
                secondary_y=True,
            )

        fig.add_trace(
            go.Scatter(
                x=times,
                y=soc,
                name="State of Charge",
                mode="lines",
                line=dict(color="black", width=2.5),
                cliponaxis=False,
            ),
            secondary_y=False,
        )

        fig.update_layout(
            barmode="relative",
            plot_bgcolor="white",
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.05,
                xanchor="center",
                x=0.5,
            ),
            xaxis=dict(showgrid=False),
            margin=dict(l=0, r=0, t=20, b=10),
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
        fig.write_image(out_file)

    def pareto_frontier(self) -> list[dict]:
        """Solve the model 101 times across evenly spaced weight combinations and return results."""
        weight_pairs = (
            [(0.9999, 0.0001)]
            + [(round(1.0 - i * 0.01, 2), round(i * 0.01, 2)) for i in range(1, 100)]
            + [(0.0001, 0.9999)]
        )
        results = []

        for i, (lp, lc) in enumerate(weight_pairs):
            self.lambda_profit = lp
            self.lambda_co2 = lc
            print(f"\nλ_profit={lp:.2f}, λ_co2={lc:.2f}")
            solved = self.solve()
            profit, co2, breakdown = self._extract_objectives(solved)
            results.append(
                {"lambda_profit": lp, "lambda_co2": lc, "profit": profit, "co2": co2}
            )
            if i == 0:
                self.visualize_profit_distribution(breakdown, lp, lc)
                self.visualize_schedule(solved)
            if i == len(weight_pairs) - 1:
                self.visualize_profit_distribution(
                    breakdown,
                    lp,
                    lc,
                    out_file="results/model_3_profit_co2_saving_extreme.png",
                )
                self.visualize_schedule(
                    solved,
                    out_file="results/model_3_schedule_co2_saving_extreme.png",
                )

        return results

    def save_results(self, results: list[dict]):
        """Save pareto frontier results (weights, profit, CO2) to an Excel file."""
        out = Path(self.results_file_path)
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
        print(f"\nResults saved to {out}")

    def visualize_pareto_frontier(self, results: list[dict]):
        profits = [r["profit"] for r in results]
        co2s = [r["co2"] for r in results]

        fig = go.Figure()

        # Shaded region: profit > 0 and CO2 < 0
        x_max = max(profits) * 1.05
        y_min = min(co2s) * 1.05 if min(co2s) < 0 else -1
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

        fig.add_trace(
            go.Scatter(
                x=profits,
                y=co2s,
                mode="lines+markers",
                marker=dict(size=8, color="seagreen"),
                line=dict(color="seagreen", width=1.5),
            )
        )
        fig.update_layout(
            xaxis_title="Profit (DKK)",
            yaxis_title="CO₂ Emissions (kg)",
            template="plotly_white",
            margin=dict(l=0, r=0, t=20, b=10),
        )
        fig.show()
        fig.write_image("results/model_3_pareto_frontier.png")


if __name__ == "__main__":
    m = Model3(
        start_date="2026-04-01",
        end_date="2026-04-30",
    )
    pareto_results = m.pareto_frontier()
    m.save_results(pareto_results)
    m.visualize_pareto_frontier(pareto_results)
