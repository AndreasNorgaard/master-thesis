"""
Sequential / rolling-horizon model for Analysis 3 (Section 6.3).

The model simulates the actual Danish bidding sequence by solving a 3-day
look-ahead LP five times per delivery day, once at each auction closure
(D-2 15:00 FCR early; D-1 07:30 aFRR up/down and mFRR up/down -- treated
as one combined pool because they clear simultaneously; D-1 12:00
day-ahead; D-1 15:00 FFR; D-1 18:00 FCR late). Capacity bids committed in
earlier pools for the delivery day are fixed before solving the next pool.

Two variants are produced: one with forecasted prices (24h-lagged realized
prices, naive forecast) and one with realized prices. The forecast variant
captures both the structural cost of sequential commitment and the cost of
forecast error; the realized variant isolates only the structural cost.

The file is intentionally self-contained: it duplicates the data-loading and
constraint logic from ``models.model_3`` so it can be read in isolation.
"""

from datetime import datetime, timedelta
from pathlib import Path

import holidays
import plotly.graph_objects as go
import polars as pl
import pyomo.environ as pyo
import xlsxwriter

from data.energi_data_service import EnergiDataServiceAPIClient


class SequentialModel:
    EUR_TO_DKK = 7.4588

    # Pool index -> list of (product, idx_type). idx_type is 'q' (quarterly),
    # 'h' (hourly) or 'b' (4-hour block). The order is the chronological order
    # of auction closures around delivery day D.
    POOL_PRODUCTS: dict[int, list[tuple[str, str]]] = {
        1: [("fcrn_E", "h"), ("fcrd_up_E", "h"), ("fcrd_down_E", "h")],
        2: [
            ("afrr_up", "b"),
            ("afrr_down", "b"),
            ("mfrr_up", "h"),
            ("mfrr_down", "h"),
        ],
        3: [("da_buy", "q"), ("da_sell", "q")],
        4: [("ffr", "h")],
        5: [("fcrn_L", "h"), ("fcrd_up_L", "h"), ("fcrd_down_L", "h")],
    }

    HOURLY_PRODUCTS = [
        "ffr",
        "fcrd_up_E",
        "fcrd_up_L",
        "fcrd_down_E",
        "fcrd_down_L",
        "fcrn_E",
        "fcrn_L",
        "mfrr_up",
        "mfrr_down",
    ]
    BLOCK_PRODUCTS = ["afrr_up", "afrr_down"]
    QUARTERLY_PRODUCTS = ["da_buy", "da_sell"]

    HOURLY_PRICE_COLS = {
        "ffr": "P_FFR",
        "fcrd_up_E": "P_FCRD_up_E",
        "fcrd_up_L": "P_FCRD_up_L",
        "fcrd_down_E": "P_FCRD_down_E",
        "fcrd_down_L": "P_FCRD_down_L",
        "fcrn_E": "P_FCRN_E",
        "fcrn_L": "P_FCRN_L",
        "mfrr_up": "P_mFRR_up",
        "mfrr_down": "P_mFRR_down",
    }
    BLOCK_PRICE_COLS = {
        "afrr_up": "P_aFRR_up",
        "afrr_down": "P_aFRR_down",
    }

    def __init__(
        self,
        start_date: str,
        end_date: str,
        bat_mw: float = 2,
        bat_mwh: float = 4,
    ):
        self.results_file_path = "results/price_uncertainty.xlsx"
        self.start_date = start_date
        self.end_date = end_date

        # BESS configuration (identical to Model 3)
        self.bat_mw = bat_mw
        self.bat_mwh = bat_mwh
        self.bat_charge_eff = 0.99
        self.bat_discharge_eff = 0.87
        self.soc_initial_frac = 0.50
        self.soc_initial = self.soc_initial_frac * self.bat_mwh
        self.soc_min = 0.05 * self.bat_mwh
        self.soc_max = 0.95 * self.bat_mwh
        self.soc_quarterly_loss = 0.025 / (30 * 24 * 4)
        self.delta_t = 0.25
        self.cycle_cost = 13.0 * 7.44
        self.n_cycles = 2

        self.tariff_prod = 15.5

        self.E_FCRD = 1.0 / 3.0
        self.E_FCRN = 1.0
        self.E_aFRR = 4.0
        self.E_mFRR = 0.25

        # Weight parameters are reset before each solve.
        self.lambda_profit = 1.0
        self.lambda_co2 = 0.0

        self.load_data()

    # ---------------------------------------------------------------- data

    def load_data(self, write_to_file: bool = False) -> None:
        """Load realized series over [start - 1 day, end + 2 days].

        The extra day at the front provides a 24-hour-lagged forecast for the
        first delivery day; the extra two days at the back provide the
        look-ahead window for the last delivery day.
        """
        sd = (
            (datetime.fromisoformat(self.start_date) - timedelta(days=1))
            .date()
            .isoformat()
        )
        # +3: the API's `end` is exclusive, so to include `end_date + 2`
        # (the last look-ahead day) we must request one further day.
        ed = (
            (datetime.fromisoformat(self.end_date) + timedelta(days=3))
            .date()
            .isoformat()
        )
        self._data_start = sd
        self._data_end = ed

        client = EnergiDataServiceAPIClient(
            start_date=sd, end_date=ed, price_area="DK2"
        )
        df_da = client.day_ahead_prices(write_to_file=False)
        df_co2 = client.co2_emissions(write_to_file=False)
        df_ffr = client.ffr_capacity(write_to_file=False)
        df_fcr_nd = client.fcr_nd_capacity(write_to_file=False)
        df_afrr = client.afrr_capacity(write_to_file=False)
        df_mfrr = client.mfrr_capacity(write_to_file=False)

        self.df, self.df_hourly, self.df_block = self._create_dataset(
            df_da, df_co2, df_ffr, df_fcr_nd, df_afrr, df_mfrr
        )

        if write_to_file:
            out = Path("data/prepared/price_uncertainty.xlsx")
            out.parent.mkdir(parents=True, exist_ok=True)
            wb = xlsxwriter.Workbook(str(out))
            self.df.write_excel(wb, worksheet="quarterly")
            self.df_hourly.write_excel(wb, worksheet="hourly")
            self.df_block.write_excel(wb, worksheet="blocks")
            wb.close()

    def _create_dataset(
        self,
        df_da: pl.DataFrame,
        df_co2: pl.DataFrame,
        df_ffr: pl.DataFrame,
        df_fcr_nd: pl.DataFrame,
        df_afrr: pl.DataFrame,
        df_mfrr: pl.DataFrame,
    ) -> tuple[pl.DataFrame, pl.DataFrame, pl.DataFrame]:
        """Duplicated from ``Model3.create_dataset`` so this file is standalone."""
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
            years=range(int(self._data_start[:4]), int(self._data_end[:4]) + 1)
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

        df_hours = (
            df.with_row_index("__qidx")
            .filter(pl.col("__qidx") % 4 == 0)
            .select(
                pl.col("TimeUTC").alias("TimeUTC_anchor"),
                pl.col("TimeDK"),
                pl.col("TimeUTC").dt.truncate("1h").alias("TimeUTC"),
            )
        )

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

        for col in ["P_FFR", *fcr_cols, "P_mFRR_up", "P_mFRR_down"]:
            df_hours = df_hours.with_columns(pl.col(col).fill_null(0.0))

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

    # ----------------------------------------------------- price slicing

    @staticmethod
    def _quarter_to_hour(q: int) -> int:
        return (q - 1) // 4 + 1

    @staticmethod
    def _quarter_to_block(q: int) -> int:
        return (q - 1) // 16 + 1

    def _slice_prices(
        self, day_offset: int, use_forecast: bool
    ) -> dict[str, list[float]]:
        """Return the 3-day window of price/CO2 series the operator sees at the
        time of bidding.

        ``day_offset`` is the index of delivery day D in the loaded ``df``
        (counted in days from ``self._data_start``). With the loader pulling
        ``start_date - 1 day``, the first delivery day (``start_date``) has
        ``day_offset = 1``.

        With ``use_forecast=True`` the prices are 24h-lagged (the operator's
        forecast = previous day's realized prices, applied to D, D+1 and D+2).
        """
        if use_forecast:
            p_day = day_offset - 1
        else:
            p_day = day_offset

        q_start = 96 * p_day
        h_start = 24 * p_day
        b_start = 6 * p_day

        slice_q = self.df.slice(q_start, 288)
        slice_h = self.df_hourly.slice(h_start, 72)
        slice_b = self.df_block.slice(b_start, 18)

        return {
            "da_price": slice_q["DayAheadPriceDKK"].to_list(),
            "co2": slice_q["CO2Emission"].to_list(),
            "tariff_cons": slice_q["tariff_cons"].to_list(),
            **{
                f"p_{prod}": slice_h[col].to_list()
                for prod, col in self.HOURLY_PRICE_COLS.items()
            },
            **{
                f"p_{prod}": slice_b[col].to_list()
                for prod, col in self.BLOCK_PRICE_COLS.items()
            },
        }

    # ------------------------------------------------------ model build

    def _build_window_model(
        self,
        prices: dict[str, list[float]],
        fixed_bids: dict[str, dict[int, float]],
        lambda_profit: float,
        lambda_co2: float,
    ) -> pyo.ConcreteModel:
        """Build a 3-day LP using the supplied price dictionary.

        Variables listed in ``fixed_bids`` are pinned (these are the
        delivery-day commitments accumulated from earlier auction pools).
        """
        m = pyo.ConcreteModel()
        Q, D, H, B = 288, 3, 72, 18
        m.quarters = pyo.RangeSet(1, Q)
        m.days = pyo.RangeSet(1, D)
        m.hours = pyo.RangeSet(1, H)
        m.blocks = pyo.RangeSet(1, B)

        # Scalar parameters
        m.bat_mw = pyo.Param(initialize=self.bat_mw)
        m.bat_mwh = pyo.Param(initialize=self.bat_mwh)
        m.bat_charge_eff = pyo.Param(initialize=self.bat_charge_eff)
        m.bat_discharge_eff = pyo.Param(initialize=self.bat_discharge_eff)
        m.soc_initial = pyo.Param(initialize=self.soc_initial)
        m.soc_min = pyo.Param(initialize=self.soc_min)
        m.soc_max = pyo.Param(initialize=self.soc_max)
        m.lam = pyo.Param(initialize=self.soc_quarterly_loss)
        m.tariff_prod = pyo.Param(initialize=self.tariff_prod)
        m.cycle_cost = pyo.Param(initialize=self.cycle_cost)
        m.n_cycles = pyo.Param(initialize=self.n_cycles)

        def q_param(name: str) -> pyo.Param:
            vals = prices[name]
            return pyo.Param(
                m.quarters,
                initialize={
                    q: float(vals[q - 1]) if vals[q - 1] is not None else 0.0
                    for q in range(1, Q + 1)
                },
            )

        def h_param(name: str) -> pyo.Param:
            vals = prices[name]
            return pyo.Param(
                m.hours,
                initialize={
                    h: float(vals[h - 1]) if vals[h - 1] is not None else 0.0
                    for h in range(1, H + 1)
                },
            )

        def b_param(name: str) -> pyo.Param:
            vals = prices[name]
            return pyo.Param(
                m.blocks,
                initialize={
                    b: float(vals[b - 1]) if vals[b - 1] is not None else 0.0
                    for b in range(1, B + 1)
                },
            )

        m.da_price = q_param("da_price")
        m.gamma = q_param("co2")
        m.tariff_cons = q_param("tariff_cons")
        for prod in self.HOURLY_PRODUCTS:
            setattr(m, f"p_{prod}", h_param(f"p_{prod}"))
        for prod in self.BLOCK_PRODUCTS:
            setattr(m, f"p_{prod}", b_param(f"p_{prod}"))

        # Decision variables
        m.da_buy = pyo.Var(m.quarters, bounds=(0, self.bat_mw))
        m.da_sell = pyo.Var(
            m.quarters, bounds=(0, self.bat_discharge_eff * self.bat_mw)
        )
        m.soc = pyo.Var(m.quarters, bounds=(self.soc_min, self.soc_max))
        for prod in self.HOURLY_PRODUCTS:
            setattr(m, prod, pyo.Var(m.hours, bounds=(0, None)))
        for prod in self.BLOCK_PRODUCTS:
            setattr(m, prod, pyo.Var(m.blocks, bounds=(0, None)))

        dt = self.delta_t

        # --- Objective ---
        da_profit = sum(
            dt
            * (
                m.da_sell[q] * (m.da_price[q] - m.tariff_prod)
                - m.da_buy[q] * (m.da_price[q] + m.tariff_cons[q])
                - m.cycle_cost * (m.da_buy[q] + m.da_sell[q])
            )
            for q in m.quarters
        )
        hourly_rev = sum(
            m.ffr[h] * m.p_ffr[h]
            + m.fcrd_up_E[h] * m.p_fcrd_up_E[h]
            + m.fcrd_up_L[h] * m.p_fcrd_up_L[h]
            + m.fcrd_down_E[h] * m.p_fcrd_down_E[h]
            + m.fcrd_down_L[h] * m.p_fcrd_down_L[h]
            + m.fcrn_E[h] * m.p_fcrn_E[h]
            + m.fcrn_L[h] * m.p_fcrn_L[h]
            + m.mfrr_up[h] * m.p_mfrr_up[h]
            + m.mfrr_down[h] * m.p_mfrr_down[h]
            for h in m.hours
        )
        block_rev = sum(
            m.afrr_up[b] * m.p_afrr_up[b] + m.afrr_down[b] * m.p_afrr_down[b]
            for b in m.blocks
        )
        co2 = sum(dt * m.gamma[q] * (m.da_buy[q] - m.da_sell[q]) for q in m.quarters)
        m.objective = pyo.Objective(
            expr=lambda_profit * (da_profit + hourly_rev + block_rev)
            - lambda_co2 * co2,
            sense=pyo.maximize,
        )

        # --- SoC dynamics ---
        def soc_dyn_rule(m, q):
            if q == 1:
                prev = m.soc_initial
            else:
                prev = m.soc[q - 1]
            return m.soc[q] == (
                (1 - m.lam) * prev
                + dt
                * (m.bat_charge_eff * m.da_buy[q] - m.da_sell[q] / m.bat_discharge_eff)
            )

        m.eq_soc_dynamics = pyo.Constraint(m.quarters, rule=soc_dyn_rule)

        def soc_final_rule(m):
            return m.soc[Q] == m.soc_initial

        m.eq_soc_final = pyo.Constraint(rule=soc_final_rule)

        # --- Cycle constraint per day ---
        def cycle_rule(m, d):
            qs = range(96 * (d - 1) + 1, 96 * d + 1)
            return sum(dt * m.da_buy[q] for q in qs) <= m.n_cycles * m.bat_mwh

        m.eq_cycle = pyo.Constraint(m.days, rule=cycle_rule)

        # --- LER buffer constraints ---
        def ler_fcrd_up(m, q):
            h = self._quarter_to_hour(q)
            b = self._quarter_to_block(q)
            return (
                m.soc_min
                + (m.fcrd_up_E[h] + m.fcrd_up_L[h]) * self.E_FCRD
                + m.afrr_up[b] * self.E_aFRR
                + m.mfrr_up[h] * self.E_mFRR
                <= m.soc[q]
            )

        def ler_fcrd_down(m, q):
            h = self._quarter_to_hour(q)
            b = self._quarter_to_block(q)
            return (
                m.soc_max
                - (m.fcrd_down_E[h] + m.fcrd_down_L[h]) * self.E_FCRD
                - m.afrr_down[b] * self.E_aFRR
                - m.mfrr_down[h] * self.E_mFRR
                >= m.soc[q]
            )

        def ler_fcrn_up(m, q):
            h = self._quarter_to_hour(q)
            b = self._quarter_to_block(q)
            return (
                m.soc_min
                + (m.fcrn_E[h] + m.fcrn_L[h]) * self.E_FCRN
                + m.afrr_up[b] * self.E_aFRR
                + m.mfrr_up[h] * self.E_mFRR
                <= m.soc[q]
            )

        def ler_fcrn_down(m, q):
            h = self._quarter_to_hour(q)
            b = self._quarter_to_block(q)
            return (
                m.soc_max
                - (m.fcrn_E[h] + m.fcrn_L[h]) * self.E_FCRN
                - m.afrr_down[b] * self.E_aFRR
                - m.mfrr_down[h] * self.E_mFRR
                >= m.soc[q]
            )

        m.eq_ler_fcrd_up = pyo.Constraint(m.quarters, rule=ler_fcrd_up)
        m.eq_ler_fcrd_down = pyo.Constraint(m.quarters, rule=ler_fcrd_down)
        m.eq_ler_fcrn_up = pyo.Constraint(m.quarters, rule=ler_fcrn_up)
        m.eq_ler_fcrn_down = pyo.Constraint(m.quarters, rule=ler_fcrn_down)

        # --- Power-flow constraints ---
        def power_discharge(m, q):
            h = self._quarter_to_hour(q)
            b = self._quarter_to_block(q)
            return (
                m.da_sell[q]
                + m.ffr[h]
                + m.fcrd_up_E[h]
                + m.fcrd_up_L[h]
                + m.fcrn_E[h]
                + m.fcrn_L[h]
                + m.afrr_up[b]
                + m.mfrr_up[h]
                <= m.bat_discharge_eff * m.bat_mw
            )

        def power_charge(m, q):
            h = self._quarter_to_hour(q)
            b = self._quarter_to_block(q)
            return (
                m.da_buy[q]
                + m.fcrd_down_E[h]
                + m.fcrd_down_L[h]
                + m.fcrn_E[h]
                + m.fcrn_L[h]
                + m.afrr_down[b]
                + m.mfrr_down[h]
                <= m.bat_mw
            )

        m.eq_power_discharge = pyo.Constraint(m.quarters, rule=power_discharge)
        m.eq_power_charge = pyo.Constraint(m.quarters, rule=power_charge)

        # --- Fix already-committed delivery-day bids ---
        # Delivery day occupies q=1..96, h=1..24, b=1..6 within the window.
        for product, idx_to_val in fixed_bids.items():
            var = getattr(m, product)
            for idx, val in idx_to_val.items():
                var[idx].fix(float(val))

        return m

    # ----------------------------------------------------- pool / day

    def _solve_pool(
        self,
        day_offset: int,
        pool: int,
        fixed_bids: dict[str, dict[int, float]],
        use_forecast: bool,
        lambda_profit: float,
        lambda_co2: float,
    ) -> dict[str, dict[int, float]]:
        """Solve one auction pool. Returns the updated ``fixed_bids`` dict with
        the products auctioned in this pool added for the delivery day."""
        prices = self._slice_prices(day_offset, use_forecast)
        model = self._build_window_model(prices, fixed_bids, lambda_profit, lambda_co2)
        solver = pyo.SolverFactory("glpk")
        solver.solve(model)

        for product, idx_type in self.POOL_PRODUCTS[pool]:
            if idx_type == "q":
                indices = range(1, 97)
            elif idx_type == "h":
                indices = range(1, 25)
            else:
                indices = range(1, 7)
            var = getattr(model, product)
            fixed_bids.setdefault(product, {})
            for idx in indices:
                v = float(pyo.value(var[idx]))
                # Clamp tiny LP solver numerical noise back inside the
                # variable's bounds so it can be safely re-fixed next pool.
                lb = var[idx].lb
                ub = var[idx].ub
                if lb is not None and v < lb:
                    v = lb
                if ub is not None and v > ub:
                    v = ub
                fixed_bids[product][idx] = v

        return fixed_bids

    def _simulate_day(
        self,
        day_offset: int,
        use_forecast: bool,
        lambda_profit: float,
        lambda_co2: float,
    ) -> dict[str, dict[int, float]]:
        fixed_bids: dict[str, dict[int, float]] = {}
        for pool in range(1, 6):
            fixed_bids = self._solve_pool(
                day_offset,
                pool,
                fixed_bids,
                use_forecast,
                lambda_profit,
                lambda_co2,
            )
        return fixed_bids

    # --------------------------------------------------- realized eval

    def _evaluate_realized(
        self,
        committed_per_day: list[dict[str, dict[int, float]]],
        day_offsets: list[int],
    ) -> tuple[float, float]:
        """Compute realized profit and CO2 from committed bids using the
        realized series."""
        total_profit = 0.0
        total_co2 = 0.0
        dt = self.delta_t

        for d_off, committed in zip(day_offsets, committed_per_day):
            q0 = 96 * d_off
            h0 = 24 * d_off
            b0 = 6 * d_off

            da_p = self.df["DayAheadPriceDKK"][q0 : q0 + 96].to_list()
            tariff_c = self.df["tariff_cons"][q0 : q0 + 96].to_list()
            gamma = self.df["CO2Emission"][q0 : q0 + 96].to_list()

            da_buy = committed["da_buy"]
            da_sell = committed["da_sell"]

            da_revenue = dt * sum(da_sell[q + 1] * da_p[q] for q in range(96))
            prod_tariff = dt * sum(da_sell[q + 1] * self.tariff_prod for q in range(96))
            da_cost = dt * sum(da_buy[q + 1] * da_p[q] for q in range(96))
            cons_tariff = dt * sum(da_buy[q + 1] * tariff_c[q] for q in range(96))
            degradation = (
                dt
                * self.cycle_cost
                * sum(da_buy[q + 1] + da_sell[q + 1] for q in range(96))
            )
            da_profit = da_revenue - prod_tariff - da_cost - cons_tariff - degradation

            hourly_rev = 0.0
            for prod, col in self.HOURLY_PRICE_COLS.items():
                prices = self.df_hourly[col][h0 : h0 + 24].to_list()
                bids = committed[prod]
                hourly_rev += sum(bids[h + 1] * prices[h] for h in range(24))

            block_rev = 0.0
            for prod, col in self.BLOCK_PRICE_COLS.items():
                prices = self.df_block[col][b0 : b0 + 6].to_list()
                bids = committed[prod]
                block_rev += sum(bids[b + 1] * prices[b] for b in range(6))

            co2 = dt * sum(
                gamma[q] * (da_buy[q + 1] - da_sell[q + 1]) for q in range(96)
            )

            total_profit += da_profit + hourly_rev + block_rev
            total_co2 += co2

        return total_profit, total_co2

    # -------------------------------------------------------- pareto

    def _delivery_day_offsets(self) -> list[int]:
        """List of day offsets (relative to ``self._data_start``) covering
        every delivery day in [start_date, end_date]."""
        sd = datetime.fromisoformat(self.start_date).date()
        ed = datetime.fromisoformat(self.end_date).date()
        ds = datetime.fromisoformat(self._data_start).date()
        n_days = (ed - sd).days + 1
        first_offset = (sd - ds).days
        return [first_offset + i for i in range(n_days)]

    def pareto_frontier(self, use_forecast: bool) -> list[dict]:
        """101 weight pairs: extremes plus 99 evenly spaced pairs in between."""
        weight_pairs = (
            [(0.9999, 0.0001)]
            + [(round(1.0 - i * 0.01, 2), round(i * 0.01, 2)) for i in range(1, 100)]
            + [(0.0001, 0.9999)]
        )
        offsets = self._delivery_day_offsets()
        label = "forecast" if use_forecast else "realized"
        results: list[dict] = []
        for lp, lc in weight_pairs:
            print(
                f"\n[{label}] lambda_profit={lp:.2f} lambda_co2={lc:.2f}"
                f" -- simulating {len(offsets)} delivery days"
            )
            committed_per_day: list[dict[str, dict[int, float]]] = []
            for d_off in offsets:
                committed = self._simulate_day(d_off, use_forecast, lp, lc)
                committed_per_day.append(committed)
                print(f"  day_offset={d_off} done")
            profit, co2 = self._evaluate_realized(committed_per_day, offsets)
            print(f"  total realized profit = {profit:>12.2f} DKK")
            print(f"  total realized CO2    = {co2:>12.4f} kg")
            results.append(
                {
                    "lambda_profit": lp,
                    "lambda_co2": lc,
                    "profit": profit,
                    "co2": co2,
                }
            )
        return results

    # ------------------------------------------------------- output

    def save_results(
        self,
        forecast_results: list[dict],
        realized_results: list[dict],
    ) -> None:
        out = Path(self.results_file_path)
        out.parent.mkdir(parents=True, exist_ok=True)
        wb = xlsxwriter.Workbook(str(out))
        for name, results in [
            ("forecast", forecast_results),
            ("realized", realized_results),
        ]:
            df = pl.DataFrame(
                {
                    "lambda_profit": [r["lambda_profit"] for r in results],
                    "lambda_co2": [r["lambda_co2"] for r in results],
                    "profit_dkk": [r["profit"] for r in results],
                    "co2_kg": [r["co2"] for r in results],
                }
            )
            df.write_excel(wb, worksheet=name)
        wb.close()
        print(f"\nResults saved to {out}")

    def visualize_pareto_frontier(
        self,
        forecast_results: list[dict],
        realized_results: list[dict],
        model3_results: list[dict] | None = None,
    ) -> None:
        fig = go.Figure()

        if model3_results:
            fig.add_trace(
                go.Scatter(
                    x=[r["profit"] for r in model3_results],
                    y=[r["co2"] for r in model3_results],
                    mode="lines+markers",
                    name="Model 3 (perfect foresight)",
                    marker=dict(size=7, color="seagreen"),
                    line=dict(color="seagreen", width=1.5),
                )
            )

        fig.add_trace(
            go.Scatter(
                x=[r["profit"] for r in realized_results],
                y=[r["co2"] for r in realized_results],
                mode="lines+markers",
                name="Sequential (realized prices)",
                marker=dict(size=7, color="steelblue"),
                line=dict(color="steelblue", width=1.5),
            )
        )
        fig.add_trace(
            go.Scatter(
                x=[r["profit"] for r in forecast_results],
                y=[r["co2"] for r in forecast_results],
                mode="lines+markers",
                name="Sequential (forecasted prices)",
                marker=dict(size=7, color="indianred"),
                line=dict(color="indianred", width=1.5),
            )
        )
        fig.update_layout(
            xaxis_title="Profit (DKK)",
            yaxis_title="CO2 Emissions (kg)",
            template="plotly_white",
            margin=dict(l=0, r=0, t=20, b=10),
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.1,
                xanchor="center",
                x=0.5,
            ),
        )
        fig.show()
        fig.write_image("results/price_uncertainty_pareto.png")

    @staticmethod
    def _load_model3_results() -> list[dict] | None:
        """Best-effort load of the Model 3 frontier from results/model_3.xlsx."""
        path = Path("results/model_3.xlsx")
        if not path.exists():
            return None
        df = pl.read_excel(path)
        return [
            {
                "lambda_profit": float(r["lambda_profit"]),
                "lambda_co2": float(r["lambda_co2"]),
                "profit": float(r["profit_dkk"]),
                "co2": float(r["co2_kg"]),
            }
            for r in df.iter_rows(named=True)
        ]

    def run(self) -> None:
        print("\n========== Sequential model: REALIZED prices ==========")
        realized_results = self.pareto_frontier(use_forecast=False)
        print("\n========== Sequential model: FORECASTED prices ==========")
        forecast_results = self.pareto_frontier(use_forecast=True)
        self.save_results(forecast_results, realized_results)
        model3_results = self._load_model3_results()
        self.visualize_pareto_frontier(
            forecast_results, realized_results, model3_results
        )


if __name__ == "__main__":
    m = SequentialModel(
        start_date="2026-04-01",
        end_date="2026-04-30",
    )
    m.run()
