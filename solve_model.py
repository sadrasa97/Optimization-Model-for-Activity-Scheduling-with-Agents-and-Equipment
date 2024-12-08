import pandas as pd
import numpy as np
from pyomo.environ import *

def load_input_data(filename):
    data = {}
    data["Dh"] = pd.read_excel(filename, sheet_name="Dh")["Dh"].values
    data["Tmin"] = pd.read_excel(filename, sheet_name="Tmin")["Tmin"].values
    data["Tmax"] = pd.read_excel(filename, sheet_name="Tmax")["Tmax"].values
    data["Kh"] = pd.read_excel(filename, sheet_name="Kh")["Kh"].values
    data["Bhj"] = pd.read_excel(filename, sheet_name="Bhj").values
    data["Rhe"] = pd.read_excel(filename, sheet_name="Rhe").values
    data["Wij"] = pd.read_excel(filename, sheet_name="Wij").values

    data["G"] = []
    data["Beta"] = []
    for sheet in pd.ExcelFile(filename).sheet_names:
        if sheet.startswith("G_Scenario"):
            data["G"].append(pd.read_excel(filename, sheet_name=sheet).values)
        elif sheet.startswith("Beta_Scenario"):
            data["Beta"].append(pd.read_excel(filename, sheet_name=sheet).values)
    data["G"] = np.stack(data["G"], axis=-1)
    data["Beta"] = np.stack(data["Beta"], axis=-1)
    return data

def solve_model(data, num_h=10, num_j=25, num_e=15, num_s=4):
    model = ConcreteModel()
    model.H = RangeSet(0, num_h - 1)
    model.J = RangeSet(0, num_j - 1)
    model.E = RangeSet(0, num_e - 1)
    model.S = RangeSet(0, num_s - 1)

    model.X = Var(model.H, model.J, model.S, within=Binary)
    model.Y = Var(model.H, model.E, model.S, within=Binary)
    model.T = Var(model.H, model.S, bounds=lambda model, h, s: (data["Tmin"][h], data["Tmax"][h]))
    model.Z = Var(model.H, model.J, model.J, model.S, within=Binary)
    model.Aux = Var(model.H, model.E, model.S, within=NonNegativeReals)

    def objective_rule(model):
        quality = sum(model.X[h, j, s] * data["Bhj"][h, j] for h in model.H for j in model.J for s in model.S)
        cost = sum(
            model.Y[h, e, s] * data["G"][h, e, s] + model.Aux[h, e, s] * data["Beta"][h, e, s]
            for h in model.H for e in model.E for s in model.S
        )
        collaboration = sum(
            model.Z[h, i, j, s] * data["Wij"][i, j]
            for h in model.H for i in model.J for j in model.J for s in model.S
        )
        return quality - cost + collaboration

    model.obj = Objective(rule=objective_rule, sense=maximize)

    def agent_constraint_rule(model, h, s):
        return sum(model.X[h, j, s] for j in model.J) == 1
    model.agent_constraint = Constraint(model.H, model.S, rule=agent_constraint_rule)

    def equipment_constraint_rule(model, h, s):
        return sum(model.Y[h, e, s] for e in model.E) >= 1
    model.equipment_constraint = Constraint(model.H, model.S, rule=equipment_constraint_rule)

    solver = SolverFactory("cbc")
    result = solver.solve(model, tee=True)

    x_results = {(h, j, s): model.X[h, j, s].value for h in model.H for j in model.J for s in model.S}
    y_results = {(h, e, s): model.Y[h, e, s].value for h in model.H for e in model.E for s in model.S}
    t_results = {(h, s): model.T[h, s].value for h in model.H for s in model.S}
    return x_results, y_results, t_results

if __name__ == "__main__":
    data = load_input_data("Generated_data.xlsx")
    x_res, y_res, t_res = solve_model(data)
    print("Agent Assignments:", x_res)
    print("Equipment Assignments:", y_res)
    print("Time Allocations:", t_res)
