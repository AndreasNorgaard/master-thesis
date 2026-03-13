import pyomo.environ as pyo

from data.energi_data_service import EnergiDataServiceAPIClient


class Model1:
    def __init__(self, start_date: str, end_date: str):
        self.results_file_path = "results/model_1.xlsx"

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

        # Load data
        self.load_data()

    def load_data(self):
        self.df = EnergiDataServiceAPIClient(
            start_date=self.start_date,
            end_date=self.end_date,
            price_area="DK2",
        ).day_ahead_prices(write_to_file=False)

    def equation_1(self, model):
        """
        Objective Function: Maximize profit over the optimization period.
        """
        return sum(
            self.delta_t
            * (
                model.da_sell[q] * (model.da_price[q] - model.tariff_prod)
                - model.da_buy[q] * (model.da_price[q] + model.tariff_cons)
                - model.cycle_cost * (model.da_buy[q] + model.da_sell[q])
            )
            for q in model.quarters
        )

    def equation_2(self, model, q):
        """
        Constraint: Charging limit
        """
        return (0, model.bat_mw)

    def equation_3(self, model, q):
        """
        Constraint: Discharging limit
        """
        return (0, model.bat_mw)

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
        model.da_price = pyo.Param(
            model.quarters,
            initialize={q: self.df["DayAheadPriceDKK"][q - 1] for q in range(1, Q + 1)},
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

    def calculate_profit(self):
        pass

    def visualize_profit(self):
        pass

    def visualize_schedule(self):
        pass


if __name__ == "__main__":
    model = Model1(start_date="2026-01-01", end_date="2026-01-31")
    model.solve()
