from datetime import datetime, timedelta
from pathlib import Path

import plotly.graph_objects as go
import polars as pl
import pyomo.environ as pyo
from plotly.subplots import make_subplots

from analysis.model_3 import Model3
from data.energi_data_service import EnergiDataServiceAPIClient


class Model4(Model3):
    """Sequential five-auction bidding extension of Model 3.

    The model solves Model 3 five times per delivery day (once per auction
    closure). At each auction, the day-of-delivery bids for that auction's
    products are captured as fixed parameters and excluded from subsequent
    solves. Optimization at each auction looks ahead two days beyond the
    delivery day; days outside this window are not modelled.

    The model can be run with realized prices (perfect information within
    the sequential structure) or with a naive forecast where each delivery
    day's window is priced using the realized values of the prior day (D-1).
    """

    AUCTIONS = [
        ("A1_FCR_early", ["fcrn_E", "fcrd_up_E", "fcrd_down_E"]),
        ("A2_aFRR_mFRR", ["afrr_up", "afrr_down", "mfrr_up", "mfrr_down"]),
        ("A3_DA", ["da_buy", "da_sell"]),
        ("A4_FFR", ["ffr"]),
        ("A5_FCR_late", ["fcrn_L", "fcrd_up_L", "fcrd_down_L"]),
    ]

    QUARTERLY_VARS = {"da_buy", "da_sell"}
    HOURLY_VARS = {
        "ffr",
        "fcrn_E",
        "fcrn_L",
        "fcrd_up_E",
        "fcrd_up_L",
        "fcrd_down_E",
        "fcrd_down_L",
        "afrr_up",
        "afrr_down",
        "mfrr_up",
        "mfrr_down",
    }

    PARETO_WEIGHTS = (
        [(0.9999, 0.0001)]
        + [(round(1.0 - i * 0.01, 2), round(i * 0.01, 2)) for i in range(1, 100)]
        + [(0.0001, 0.9999)]
    )

    # Numerical slack (in MWh / MW) applied to constraints that bind on
    # values pinned from a prior LP. GLPK's reported optimal values can
    # drift up to ~1e-5 from the exact constraint boundary; without this
    # tolerance the next rolling solve sees a slightly-violated cycle cap
    # or LER constraint and refuses to start. The slack is well below the
    # physical resolution of the model and does not affect economics.
    NUM_TOL = 1e-4
    # Half-width of the box added around each fixed bid in subsequent
    # rolling solves. Big enough to swallow GLPK's solution tolerance,
    # small enough that the bid is still effectively committed.
    PIN_SLACK = 1e-4

    HOURLY_PRICE_COLS = {
        "ffr": "P_FFR",
        "fcrd_up_E": "P_FCRD_up_E",
        "fcrd_up_L": "P_FCRD_up_L",
        "fcrd_down_E": "P_FCRD_down_E",
        "fcrd_down_L": "P_FCRD_down_L",
        "fcrn_E": "P_FCRN_E",
        "fcrn_L": "P_FCRN_L",
        "afrr_up": "P_aFRR_up",
        "afrr_down": "P_aFRR_down",
        "mfrr_up": "P_mFRR_up",
        "mfrr_down": "P_mFRR_down",
    }

    def __init__(
        self,
        start_date: str,
        end_date: str,
        bat_mw: float = 2,
        bat_mwh: float = 4,
        lookahead_days: int = 2,
    ):
        # Energi Data Service treats `end` as exclusive, so passing
        # 2026-05-01 yields 30 delivery days (Apr 1 .. Apr 30). We mirror
        # Model 3's convention here so the two horizons are identical and
        # the frontiers are directly comparable.
        super().__init__(start_date, end_date, bat_mw, bat_mwh)
        # Number of extra delivery days included in each rolling LP window.
        # 0 = optimize the delivery day in isolation; 2 = the original
        # behaviour (delivery day plus a two-day look-ahead).
        self.lookahead_days = lookahead_days
        self.results_file_path = "results/scenario_3/results.xlsx"

    # --------------------- Data loading ---------------------

    def load_data(self, write_to_file: bool = False) -> None:
        """Fetch the full dataset over [start_date - 1 day, end_date] in a
        single round-trip per endpoint, then split it into the main horizon
        and the single prior day used by the naive forecast. Overrides the
        parent's two-call behaviour to avoid hitting each Energi Data Service
        endpoint twice.
        """
        sd = datetime.strptime(self.start_date, "%Y-%m-%d")
        prior_start = (sd - timedelta(days=1)).strftime("%Y-%m-%d")

        client = EnergiDataServiceAPIClient(
            start_date=prior_start,
            end_date=self.end_date,
            price_area="DK2",
        )
        df_da = client.day_ahead_prices(write_to_file=False)
        df_co2 = client.co2_emissions(write_to_file=False)
        df_ffr = client.ffr_capacity(write_to_file=False)
        df_fcr_nd = client.fcr_nd_capacity(write_to_file=False)
        df_afrr = client.afrr_capacity(write_to_file=False)
        df_mfrr = client.mfrr_capacity(write_to_file=False)
        df_full, df_hourly_full = self.create_dataset(
            df_da, df_co2, df_ffr, df_fcr_nd, df_afrr, df_mfrr
        )

        # First 96 quarters / 24 hours belong to the prior day.
        self.df_prior = df_full.slice(0, 96)
        self.df_hourly_prior = df_hourly_full.slice(0, 24)
        # The remaining rows are the optimisation horizon proper.
        self.df = df_full.slice(96, df_full.height - 96)
        self.df_hourly = df_hourly_full.slice(24, df_hourly_full.height - 24)

    def _realized_day(self, d: int) -> tuple[pl.DataFrame, pl.DataFrame]:
        """Return realized (quarterly, hourly) frames for delivery day d.

        d == 0 returns the prior day (loaded by `_load_prior_day`).
        """
        if d == 0:
            return self.df_prior, self.df_hourly_prior
        return (
            self.df.slice(96 * (d - 1), 96),
            self.df_hourly.slice(24 * (d - 1), 24),
        )

    @staticmethod
    def _opt_float(v) -> float:
        return float(v) if v is not None else 0.0

    # ----------- Constraint overrides with numerical tolerance -----------
    # The bodies are identical to Model 3 except a small NUM_TOL slack is
    # added on the right-hand side so that values forwarded between rolling
    # solves do not produce spurious infeasibility from floating-point drift
    # in the LP solver's reported optimum.

    def equation_8(self, model, d):
        quarters_in_day = range(96 * (d - 1) + 1, 96 * d + 1)
        return (
            sum(self.delta_t * model.da_buy[q] for q in quarters_in_day)
            <= model.n_cycles * model.bat_mwh + self.NUM_TOL
        )

    def equation_ler_fcrd_up(self, model, q):
        h = self.quarter_to_hour(q)
        return (
            model.soc_min
            + (model.fcrd_up_E[h] + model.fcrd_up_L[h]) * self.E_FCRD
            + model.afrr_up[h] * self.E_aFRR
            + model.mfrr_up[h] * self.E_mFRR
            - self.NUM_TOL
            <= model.soc[q]
        )

    def equation_ler_fcrd_down(self, model, q):
        h = self.quarter_to_hour(q)
        return (
            model.soc_max
            - (model.fcrd_down_E[h] + model.fcrd_down_L[h]) * self.E_FCRD
            - model.afrr_down[h] * self.E_aFRR
            - model.mfrr_down[h] * self.E_mFRR
            + self.NUM_TOL
            >= model.soc[q]
        )

    def equation_ler_fcrn_up(self, model, q):
        h = self.quarter_to_hour(q)
        return (
            model.soc_min
            + (model.fcrn_E[h] + model.fcrn_L[h]) * self.E_FCRN
            + model.afrr_up[h] * self.E_aFRR
            + model.mfrr_up[h] * self.E_mFRR
            - self.NUM_TOL
            <= model.soc[q]
        )

    def equation_ler_fcrn_down(self, model, q):
        h = self.quarter_to_hour(q)
        return (
            model.soc_max
            - (model.fcrn_E[h] + model.fcrn_L[h]) * self.E_FCRN
            - model.afrr_down[h] * self.E_aFRR
            - model.mfrr_down[h] * self.E_mFRR
            + self.NUM_TOL
            >= model.soc[q]
        )

    def equation_power_discharge(self, model, q):
        h = self.quarter_to_hour(q)
        return (
            model.da_sell[q]
            + model.ffr[h]
            + model.fcrd_up_E[h]
            + model.fcrd_up_L[h]
            + model.fcrn_E[h]
            + model.fcrn_L[h]
            + model.afrr_up[h]
            + model.mfrr_up[h]
            <= model.bat_discharge_eff * model.bat_mw + self.NUM_TOL
        )

    def equation_power_charge(self, model, q):
        h = self.quarter_to_hour(q)
        return (
            model.da_buy[q]
            + model.fcrd_down_E[h]
            + model.fcrd_down_L[h]
            + model.fcrn_E[h]
            + model.fcrn_L[h]
            + model.afrr_down[h]
            + model.mfrr_down[h]
            <= model.bat_mw + self.NUM_TOL
        )

    # --------------------- Sequential solve ---------------------

    def _solve_window(
        self,
        delivery_day: int,
        current_soc_initial: float,
        per_day_fixed: dict,
        use_forecast: bool,
        auction_name: str = "?",
    ):
        """Solve one auction's LP window starting at `delivery_day`.

        The window covers days [delivery_day, delivery_day+2], clipped at the
        end of the horizon. All variables already in `per_day_fixed` (the
        delivery day's bids fixed by earlier auctions) are pinned via
        Pyomo's `.fix()`. SoC continuity across delivery days is handled by
        passing `current_soc_initial` (= end-of-prior-day SoC).
        """
        Q_total = len(self.df)

        q_start = 96 * (delivery_day - 1) + 1
        q_end = min(96 * (delivery_day + self.lookahead_days), Q_total)
        h_start = (q_start - 1) // 4 + 1
        h_end = q_end // 4
        d_start = delivery_day
        d_end = (q_end - 1) // 96 + 1

        end_at_horizon = q_end == Q_total

        prior_q, prior_h = (
            self._realized_day(delivery_day - 1) if use_forecast else (None, None)
        )

        model = pyo.ConcreteModel()
        model.quarters = pyo.RangeSet(q_start, q_end)
        model.hours = pyo.RangeSet(h_start, h_end)
        model.days = pyo.RangeSet(d_start, d_end)

        # Scalar parameters
        model.bat_mw = pyo.Param(initialize=self.bat_mw)
        model.bat_mwh = pyo.Param(initialize=self.bat_mwh)
        model.bat_charge_eff = pyo.Param(initialize=self.bat_charge_eff)
        model.bat_discharge_eff = pyo.Param(initialize=self.bat_discharge_eff)
        model.soc_initial = pyo.Param(initialize=current_soc_initial, mutable=True)
        model.soc_min = pyo.Param(initialize=self.soc_min)
        model.soc_max = pyo.Param(initialize=self.soc_max)
        model.lam = pyo.Param(initialize=self.soc_quarterly_loss)
        model.tariff_prod = pyo.Param(initialize=self.tariff_prod)
        model.cycle_cost = pyo.Param(initialize=self.cycle_cost)
        model.n_cycles = pyo.Param(initialize=self.n_cycles)
        model.lambda_profit = pyo.Param(initialize=self.lambda_profit)
        model.lambda_co2 = pyo.Param(initialize=self.lambda_co2)

        # Quarterly series
        tariff_cons_init: dict[int, float] = {}
        da_price_init: dict[int, float] = {}
        gamma_init: dict[int, float] = {}
        for q in range(q_start, q_end + 1):
            tod_q = (q - 1) % 96
            tariff_cons_init[q] = float(self.df["tariff_cons"][q - 1])
            if use_forecast:
                da_price_init[q] = float(prior_q["DayAheadPriceDKK"][tod_q])
                gamma_init[q] = float(prior_q["CO2Emission"][tod_q])
            else:
                da_price_init[q] = float(self.df["DayAheadPriceDKK"][q - 1])
                gamma_init[q] = float(self.df["CO2Emission"][q - 1])
        model.tariff_cons = pyo.Param(model.quarters, initialize=tariff_cons_init)
        model.da_price = pyo.Param(model.quarters, initialize=da_price_init)
        model.gamma = pyo.Param(model.quarters, initialize=gamma_init)

        # Hourly capacity prices
        for var_name, col in self.HOURLY_PRICE_COLS.items():
            attr = f"p_{var_name}" if not var_name.startswith("ffr") else "p_ffr"
            # Map var name -> parameter attribute name used inside Model 3's
            # rules (e.g. p_fcrd_up_E, p_fcrn_E, p_mfrr_up, p_ffr).
            init = {}
            for h in range(h_start, h_end + 1):
                tod_h = (h - 1) % 24
                if use_forecast:
                    init[h] = self._opt_float(prior_h[col][tod_h])
                else:
                    init[h] = self._opt_float(self.df_hourly[col][h - 1])
            setattr(model, attr, pyo.Param(model.hours, initialize=init))

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
        model.afrr_up = pyo.Var(model.hours, bounds=(0, None))
        model.afrr_down = pyo.Var(model.hours, bounds=(0, None))

        # Pin already-decided bids for this delivery day. Instead of a hard
        # .fix(value), we narrow the variable's bounds to a small interval
        # [value - PIN_SLACK, value + PIN_SLACK]. The LP can then resolve
        # tiny inconsistencies between successive solves (GLPK's reported
        # optimum is accurate only within its primal feasibility tolerance,
        # roughly 1e-7, but the SoC recurrence amplifies that drift), while
        # the bid is still effectively committed at the auction's chosen
        # level.
        for (var_name, idx), value in per_day_fixed.items():
            var = getattr(model, var_name)
            if idx not in var:
                continue
            v = var[idx]
            lo = max(0.0, value - self.PIN_SLACK)
            hi = value + self.PIN_SLACK
            v.setlb(lo)
            v.setub(hi)

        # Objective
        model.objective = pyo.Objective(expr=self.equation_1(model), sense=pyo.maximize)

        # Inherited Model 3 constraints
        model.eq5 = pyo.Constraint(model.quarters, rule=self.equation_5)
        model.eq6 = pyo.Constraint(model.quarters, rule=self.equation_6)
        if end_at_horizon:
            # Close the SoC cycle only at the very end of the full horizon.
            soc_target = self.soc_initial * self.bat_mwh

            def _eq7_rule(model_, q):
                if q != Q_total:
                    return pyo.Constraint.Skip
                return model_.soc[q] == soc_target

            model.eq7 = pyo.Constraint(model.quarters, rule=_eq7_rule)
        model.eq8 = pyo.Constraint(model.days, rule=self.equation_8)
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
        results = solver.solve(model)
        status = results.solver.status
        condition = results.solver.termination_condition
        if (
            status != pyo.SolverStatus.ok
            or condition != pyo.TerminationCondition.optimal
        ):
            debug_path = Path(
                f"results/scenario_3/debug/d{delivery_day:02d}_{auction_name}.lp"
            )
            debug_path.parent.mkdir(parents=True, exist_ok=True)
            try:
                model.write(
                    str(debug_path),
                    io_options={"symbolic_solver_labels": True},
                )
            except Exception as exc:  # noqa: BLE001
                print(f"  (could not write LP file: {exc})")
            raise RuntimeError(
                f"GLPK failed for delivery day {delivery_day}, auction "
                f"{auction_name} (status={status}, termination={condition}, "
                f"soc_initial={current_soc_initial:.6f}, "
                f"#fixed_bids={len(per_day_fixed)}, LP written to {debug_path})."
            )
        return model

    def _capture_day_bids(
        self,
        model,
        delivery_day: int,
        products: list[str],
        per_day_fixed: dict,
        global_bids: dict,
    ) -> None:
        """Pull this auction's day-of-delivery bids out of the solved model
        and append them to both the per-day fixing dict and the global
        accumulator used for objective evaluation.
        """
        day_q_start = 96 * (delivery_day - 1) + 1
        day_q_end = 96 * delivery_day
        day_h_start = 24 * (delivery_day - 1) + 1
        day_h_end = 24 * delivery_day

        # Strict upper bounds on the LP variables. These are only used to
        # clean up bid values reported externally (objective evaluation,
        # Excel output); the values pinned to subsequent solves keep their
        # raw LP magnitude to avoid breaking the SoC / LER cascade.
        upper_bounds = {
            "da_buy": self.bat_mw,
            "da_sell": self.bat_discharge_eff * self.bat_mw,
        }

        for product in products:
            var = getattr(model, product)
            if product in self.QUARTERLY_VARS:
                rng = range(day_q_start, day_q_end + 1)
            elif product in self.HOURLY_VARS:
                rng = range(day_h_start, day_h_end + 1)
            else:
                raise ValueError(f"Unknown product {product}")
            upper = upper_bounds.get(product)
            for idx in rng:
                v_raw = pyo.value(var[idx])
                # Sanitised copy used for reporting only.
                v_clean = max(0.0, v_raw)
                if upper is not None and v_clean > upper:
                    v_clean = upper
                # Forward the exact LP value to keep downstream LPs feasible.
                per_day_fixed[(product, idx)] = v_raw
                global_bids[(product, idx)] = v_clean

    def solve_sequential(self, use_forecast: bool) -> dict:
        """Run the sequential five-auction model over the full horizon.

        Returns a dict keyed by (product_name, global_index) -> bid value.
        """
        Q_total = len(self.df)
        D_total = Q_total // 96

        global_bids: dict = {}
        current_soc_initial = self.soc_initial * self.bat_mwh

        for d in range(1, D_total + 1):
            per_day_fixed: dict = {}
            last_model = None
            for auction_name, products in self.AUCTIONS:
                model = self._solve_window(
                    d,
                    current_soc_initial,
                    per_day_fixed,
                    use_forecast,
                    auction_name=auction_name,
                )
                self._capture_day_bids(model, d, products, per_day_fixed, global_bids)
                last_model = model
            # Hand off SoC: end-of-day-d SoC becomes start of day d+1, clipped
            # to the feasible SoC band to absorb GLPK numerical noise.
            end_q = 96 * d
            current_soc_initial = pyo.value(last_model.soc[end_q])
            current_soc_initial = min(
                max(current_soc_initial, self.soc_min), self.soc_max
            )
            print(
                f"  Day {d:02d}/{D_total} done. End-of-day SoC = "
                f"{current_soc_initial:.3f} MWh"
            )
        return global_bids

    # --------------------- Objective evaluation ---------------------

    def _extract_objectives_from_bids(self, bids: dict) -> tuple[float, float, dict]:
        """Aggregate profit and CO2 over the full horizon using realized
        prices and CO2 intensities, regardless of which prices were used
        during optimization. This makes the three frontiers directly
        comparable in objective terms.
        """
        Q = len(self.df)
        H = Q // 4
        dt = self.delta_t

        def q_bid(name, q):
            return bids.get((name, q), 0.0)

        def h_bid(name, h):
            return bids.get((name, h), 0.0)

        da_buy = [q_bid("da_buy", q) for q in range(1, Q + 1)]
        da_sell = [q_bid("da_sell", q) for q in range(1, Q + 1)]
        da_price = self.df["DayAheadPriceDKK"].to_list()
        tariff_cons = self.df["tariff_cons"].to_list()
        gamma = self.df["CO2Emission"].to_list()

        da_revenue = dt * sum(s * p for s, p in zip(da_sell, da_price))
        prod_tariff = dt * sum(s * self.tariff_prod for s in da_sell)
        da_cost = dt * sum(b * p for b, p in zip(da_buy, da_price))
        cons_tariff = dt * sum(b * t for b, t in zip(da_buy, tariff_cons))
        degradation = dt * self.cycle_cost * sum(b + s for b, s in zip(da_buy, da_sell))
        da_profit = da_revenue - prod_tariff - da_cost - cons_tariff - degradation

        def hourly_rev(name, col):
            return sum(
                h_bid(name, h) * self._opt_float(self.df_hourly[col][h - 1])
                for h in range(1, H + 1)
            )

        ffr_rev = hourly_rev("ffr", "P_FFR")
        fcrd_up_E_rev = hourly_rev("fcrd_up_E", "P_FCRD_up_E")
        fcrd_up_L_rev = hourly_rev("fcrd_up_L", "P_FCRD_up_L")
        fcrd_down_E_rev = hourly_rev("fcrd_down_E", "P_FCRD_down_E")
        fcrd_down_L_rev = hourly_rev("fcrd_down_L", "P_FCRD_down_L")
        fcrn_E_rev = hourly_rev("fcrn_E", "P_FCRN_E")
        fcrn_L_rev = hourly_rev("fcrn_L", "P_FCRN_L")
        mfrr_up_rev = hourly_rev("mfrr_up", "P_mFRR_up")
        mfrr_down_rev = hourly_rev("mfrr_down", "P_mFRR_down")
        afrr_up_rev = hourly_rev("afrr_up", "P_aFRR_up")
        afrr_down_rev = hourly_rev("afrr_down", "P_aFRR_down")

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
        co2 = sum(dt * g * (b - s) for g, b, s in zip(gamma, da_buy, da_sell))

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

    # --------------------- Frontier sweeps ---------------------

    def pareto_frontier_sequential(self, use_forecast: bool) -> list[dict]:
        mode = "forecast" if use_forecast else "realized"
        results: list[dict] = []
        for lp, lc in self.PARETO_WEIGHTS:
            self.lambda_profit = lp
            self.lambda_co2 = lc
            print(f"\n[{mode}] λ_profit={lp:.4f}, λ_co2={lc:.4f}")
            bids = self.solve_sequential(use_forecast=use_forecast)
            profit, co2, _ = self._extract_objectives_from_bids(bids)
            print(f"  Total profit: {profit:>12.2f} DKK")
            print(f"  CO2:          {co2:>12.2f} kg")
            results.append(
                {
                    "lambda_profit": lp,
                    "lambda_co2": lc,
                    "profit": profit,
                    "co2": co2,
                }
            )
        return results

    def run_model3_baseline(self) -> list[dict]:
        """Re-run Model 3 with the same weight pairs as Model 3 (Sequential) to build
        an apples-to-apples perfect-foresight baseline.
        """
        baseline = Model3(
            start_date=self.start_date,
            end_date=self.end_date,
            bat_mw=self.bat_mw,
            bat_mwh=self.bat_mwh,
            df=self.df,
            df_hourly=self.df_hourly,
        )
        results: list[dict] = []
        for lp, lc in self.PARETO_WEIGHTS:
            baseline.lambda_profit = lp
            baseline.lambda_co2 = lc
            print(f"\n[model3] λ_profit={lp:.4f}, λ_co2={lc:.4f}")
            solved = baseline.solve()
            profit, co2, _ = baseline._extract_objectives(solved)
            results.append(
                {
                    "lambda_profit": lp,
                    "lambda_co2": lc,
                    "profit": profit,
                    "co2": co2,
                }
            )
        return results

    @staticmethod
    def save_results(results: list[dict], path: str) -> None:
        out = Path(path)
        out.parent.mkdir(parents=True, exist_ok=True)
        pl.DataFrame(
            {
                "lambda_profit": [r["lambda_profit"] for r in results],
                "lambda_co2": [r["lambda_co2"] for r in results],
                "profit_dkk": [r["profit"] for r in results],
                "co2_kg": [r["co2"] for r in results],
            }
        ).write_excel(out)
        print(f"Saved to {out}")

    @staticmethod
    def load_results(path: str) -> list[dict]:
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

    # --------------------- Visualisation ---------------------

    @staticmethod
    def visualize_three_frontiers(
        baseline: list[dict],
        realized: list[dict],
        forecast: list[dict],
        out_file: str = "results/scenario_3/pareto_comparison.png",
    ) -> None:
        all_profits = [
            r["profit"] for results in (baseline, realized, forecast) for r in results
        ]
        all_co2s = [
            r["co2"] for results in (baseline, realized, forecast) for r in results
        ]
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

        for results, name, color in [
            (baseline, "Model 3 (perfect foresight)", "seagreen"),
            (realized, "Model 3 (Sequential) (realized prices)", "steelblue"),
            (forecast, "Model 3 (Sequential) (forecast prices)", "darkorange"),
        ]:
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
                    name=name,
                    text=labels,
                    hovertemplate=(
                        "%{text}<br>Profit: %{x:.2f} DKK"
                        "<br>CO₂: %{y:.4f} kg<extra></extra>"
                    ),
                    marker=dict(size=7, color=color),
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

        out = Path(out_file)
        out.parent.mkdir(parents=True, exist_ok=True)
        fig.show()
        fig.write_image(str(out))
        print(f"Pareto comparison saved to {out}")

    @staticmethod
    def visualize_vpi(
        baseline: list[dict],
        realized: list[dict],
        forecast: list[dict],
        out_file: str = "results/scenario_3/vpi.png",
    ) -> None:
        """Compare profit and CO₂ across the 11 weight pairs for all three
        setups. The vertical gap between curves at any λ_profit gives the
        value of perfect information (between Model 3 (Sequential) forecast and Model 3 (Sequential)
        realized) and the cost of sequential commitment (between Model 3 (Sequential)
        realized and Model 3 perfect foresight).
        """
        fig = make_subplots(
            rows=2,
            cols=1,
            shared_xaxes=True,
            vertical_spacing=0.08,
            subplot_titles=("Profit (DKK)", "CO₂ Emissions (kg)"),
        )

        x_labels = [f"{r['lambda_profit']:.2f}" for r in baseline]

        for results, name, color in [
            (baseline, "Model 3 (perfect foresight)", "seagreen"),
            (realized, "Model 3 (Sequential) (realized prices)", "steelblue"),
            (forecast, "Model 3 (Sequential) (forecast prices)", "darkorange"),
        ]:
            profits = [r["profit"] for r in results]
            co2s = [r["co2"] for r in results]
            fig.add_trace(
                go.Scatter(
                    x=x_labels,
                    y=profits,
                    mode="lines+markers",
                    name=name,
                    legendgroup=name,
                    marker=dict(size=7, color=color),
                    line=dict(color=color, width=1.5),
                ),
                row=1,
                col=1,
            )
            fig.add_trace(
                go.Scatter(
                    x=x_labels,
                    y=co2s,
                    mode="lines+markers",
                    name=name,
                    legendgroup=name,
                    showlegend=False,
                    marker=dict(size=7, color=color),
                    line=dict(color=color, width=1.5),
                ),
                row=2,
                col=1,
            )

        fig.update_xaxes(title_text="λ_profit", row=2, col=1)
        fig.update_yaxes(title_text="Profit (DKK)", row=1, col=1)
        fig.update_yaxes(title_text="CO₂ (kg)", row=2, col=1)
        fig.update_layout(
            template="plotly_white",
            legend=dict(
                orientation="h",
                yanchor="top",
                y=-0.1,
                xanchor="center",
                x=0.5,
            ),
            margin=dict(l=0, r=0, t=40, b=10),
        )

        out = Path(out_file)
        out.parent.mkdir(parents=True, exist_ok=True)
        fig.show()
        fig.write_image(str(out))
        print(f"VPI figure saved to {out}")

    @staticmethod
    def visualize_two_frontiers(
        baseline: list[dict],
        forecast: list[dict],
        out_file: str = "results/scenario_3/pareto_comparison.png",
    ) -> None:
        """Compare the Model 3 perfect-foresight frontier and the Model 3 (Sequential)
        forecast-price sequential frontier on one Pareto chart.
        """
        all_profits = [r["profit"] for results in (baseline, forecast) for r in results]
        all_co2s = [r["co2"] for results in (baseline, forecast) for r in results]
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

        for results, name, color in [
            (baseline, "Model 3 (perfect foresight)", "seagreen"),
            (forecast, "Model 3 - Sequential (forecast prices)", "darkorange"),
        ]:
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
                    name=name,
                    text=labels,
                    hovertemplate=(
                        "%{text}<br>Profit: %{x:.2f} DKK"
                        "<br>CO₂: %{y:.4f} kg<extra></extra>"
                    ),
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

        out = Path(out_file)
        out.parent.mkdir(parents=True, exist_ok=True)
        fig.show()
        fig.write_image(str(out))
        print(f"Pareto comparison saved to {out}")

    @staticmethod
    def visualize_lookahead_sweep(
        baseline: list[dict],
        frontiers: list[tuple[int, list[dict]]],
        title: str,
        out_file: str,
    ) -> None:
        """Plot the Model 3 perfect-foresight frontier together with one
        Model 3 (Sequential) frontier per look-ahead horizon, on a single Pareto chart.
        """
        all_profits = [r["profit"] for r in baseline] + [
            r["profit"] for _, results in frontiers for r in results
        ]
        all_co2s = [r["co2"] for r in baseline] + [
            r["co2"] for _, results in frontiers for r in results
        ]
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

        palette = ["steelblue", "darkorange", "crimson", "purple", "teal"]
        curves: list[tuple[list[dict], str, str]] = [
            (baseline, "Model 3 (perfect foresight)", "seagreen")
        ]
        for i, (la, results) in enumerate(frontiers):
            label = (
                f"Model 3 (Sequential) (look-ahead = {la} day{'s' if la != 1 else ''})"
            )
            curves.append((results, label, palette[i % len(palette)]))

        for results, name, color in curves:
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
                    name=name,
                    text=labels,
                    hovertemplate=(
                        "%{text}<br>Profit: %{x:.2f} DKK"
                        "<br>CO₂: %{y:.4f} kg<extra></extra>"
                    ),
                    marker=dict(size=7, color=color),
                    line=dict(color=color, width=1.5),
                )
            )

        fig.update_layout(
            title=title,
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

        out = Path(out_file)
        out.parent.mkdir(parents=True, exist_ok=True)
        fig.show()
        fig.write_image(str(out))
        print(f"Look-ahead comparison saved to {out}")


def run_debug_schedules(start: str = "2026-04-01", end: str = "2026-05-01") -> None:
    """Re-solve the Model 3 (Sequential) realized-price model (look-ahead = 2) at the
    (0.9999, 0.0001) weight pair and dump per-day production schedules, plus
    the Model 3 perfect-foresight schedule for the same weights.
    """
    m4 = Model4(start, end, lookahead_days=2)
    m4.lambda_profit = 0.9999
    m4.lambda_co2 = 0.0001
    debug_dir = "results/scenario_3/debug_schedules/realized_w0.9999_0.0001"
    Path(debug_dir).mkdir(parents=True, exist_ok=True)
    print(f"\n[debug] Sequential realized, λ=(0.9999, 0.0001). Writing to {debug_dir}")
    m4.solve_sequential(use_forecast=False, debug_schedule_dir=debug_dir)

    m3 = Model3(start_date=start, end_date=end)
    m3.lambda_profit = 0.9999
    m3.lambda_co2 = 0.0001
    print("\n[debug] Model 3 perfect-foresight, λ=(0.9999, 0.0001)")
    solved = m3.solve()
    bids, soc, q_max = m3.bids_and_soc_from_model(solved)
    m3_path = "results/scenario_3/debug_schedules/model3_w0.9999_0.0001.xlsx"
    m3.write_schedule_excel(bids, soc, q_max, m3_path)
    print(f"Model 3 schedule saved to {m3_path}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1 and sys.argv[1] == "debug":
        run_debug_schedules()
        sys.exit(0)

    START = "2026-04-01"
    END = "2026-05-01"
    LOOKAHEAD = 2

    # Prefer Model 3's own 101-point sweep if it has already been generated;
    # otherwise fall back to a scenario_3-local copy that we generate here.
    model3_path = "results/model_3/model_3.xlsx"
    baseline_path = "results/scenario_3/model_3_101pts.xlsx"
    forecast_path = f"results/scenario_3/forecast_la{LOOKAHEAD}_101pts.xlsx"

    baseline_cached = Path(model3_path).exists() or Path(baseline_path).exists()
    needs_model = (not baseline_cached) or (not Path(forecast_path).exists())

    m = Model4(START, END, lookahead_days=LOOKAHEAD) if needs_model else None

    if Path(model3_path).exists():
        print(f"Loading cached Model 3 baseline from {model3_path}")
        baseline_results = Model4.load_results(model3_path)
    elif Path(baseline_path).exists():
        print(f"Loading cached Model 3 baseline from {baseline_path}")
        baseline_results = Model4.load_results(baseline_path)
    else:
        assert m is not None
        baseline_results = m.run_model3_baseline()
        Model4.save_results(baseline_results, baseline_path)

    if Path(forecast_path).exists():
        forecast_results = Model4.load_results(forecast_path)
        print(f"Loaded cached forecast frontier from {forecast_path}")
    else:
        assert m is not None
        forecast_results = m.pareto_frontier_sequential(use_forecast=True)
        Model4.save_results(forecast_results, forecast_path)

    Model4.visualize_two_frontiers(
        baseline_results,
        forecast_results,
        out_file="results/scenario_3/pareto_comparison.png",
    )
