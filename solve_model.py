!pip install pyomo xlsxwriter
!apt-get install -y -qq glpk-utils
!apt-get install -y -qq coinor-cbc
# Install Ipopt solver dependencies
!pip install -q condacolab
import condacolab
condacolab.install()

!conda install -c conda-forge ipopt pyomo -y
!which ipopt
from pyomo.environ import SolverFactory

solver = SolverFactory('ipopt', executable='/usr/local/bin/ipopt')
print(solver.available())
# ====== Step 2: Load Input Data ======
def load_input_data(filename):
    """
    Loads input data from an Excel file and organizes it into a dictionary.
    :param filename: The file containing input data.
    :return: A dictionary containing all data.
    """
    data = {}
    # Load scalar data (1D arrays)
    data["Dh"] = pd.read_excel(filename, sheet_name="Dh")["Dh"].values
    data["Tmin"] = pd.read_excel(filename, sheet_name="Tmin")["Tmin"].values
    data["Tmax"] = pd.read_excel(filename, sheet_name="Tmax")["Tmax"].values
    data["Kh"] = pd.read_excel(filename, sheet_name="Kh")["Kh"].values
    data["Bhj"] = pd.read_excel(filename, sheet_name="Bhj").values  # 2D matrix for agents
    data["Rhe"] = pd.read_excel(filename, sheet_name="Rhe").values  # 2D matrix for equipment
    data["Wij"] = pd.read_excel(filename, sheet_name="Wij").values  # 2D collaboration matrix

    # Load multi-scenario data (3D arrays for equipment costs and variable factors)
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

# Load data from the generated file
data = load_input_data("Generated_data.xlsx")


# ====== Step 3: Optimization Model ======
def solve_model(data, num_h=10, num_j=25, num_e=15, num_s=4):
    """
    Defines and solves an optimization model using Pyomo.
    :param data: Input data dictionary.
    :param num_h: Number of activities.
    :param num_j: Number of agents.
    :param num_e: Number of equipment types.
    :param num_s: Number of scenarios.
    :return: Optimization results including variables X, Y, and T.
    """
    # Create a concrete Pyomo model
    model = ConcreteModel()
    H = range(num_h) 
    J = range(num_j)  
    E = range(num_e) 
    S = range(num_s) 

    # ====== Sets ======
    model.H = RangeSet(0, num_h - 1)  # Set of activities
    model.J = RangeSet(0, num_j - 1)  # Set of agents
    model.E = RangeSet(0, num_e - 1)  # Set of equipment types
    model.S = RangeSet(0, num_s - 1)  # Set of scenarios

        # ====== Variables ======
    model.X = Var(model.H, model.J, model.S,bounds=(0,1))  
    model.Y = Var(model.H, model.E, model.S,bounds=(0,1))  

    model.T = Var(model.H, model.S, bounds=lambda model, h, s: (data["Tmin"][h], data["Tmax"][h]))
    model.Z = Var(model.H, model.J, model.J, model.S, within=Binary)  # Collaboration binary variable
    model.U = Var(model.H, model.S, within=NonNegativeReals)

    model.Aux = Var(model.H, model.E, model.S, within=NonNegativeReals)  # Auxiliary variable for cost calculation

    model.P = Param(model.S, initialize=lambda model, s: 1 / len(data["G"][0][0]))
    model.M = Param(initialize=1000)  
    model.B = Param(initialize=50)  
    model.R = Param(initialize=100)  
    model.G = Param(model.H, model.E, model.S, initialize=lambda model, h, e, s: data["G"][h, e, s])
    model.Beta = Param(model.H, model.E, model.S, initialize=lambda model, h, e, s: data["Beta"][h, e, s])
    model.O = Param(model.J, model.H, model.S, initialize=lambda model, j, h, s: np.random.randint(50, 100))
    model.Alpha = Param(model.J, model.H, model.S, initialize=lambda model, j, h, s: np.random.uniform(5, 10))
    model.Delta = Param(model.J, model.H, model.S, initialize=lambda model, j, h, s: np.random.uniform(0.1, 0.5))



    # ====== Objective ======

    # Combine the objectives using weights
    weights = {"cost": 0.1, "quality": 0.6, "collaboration": 0.3}

    def combined_objective_rule(model):
        cost_component = sum(
            model.P[s] * (
                sum(model.Y[h, e, s] * model.G[h, e, s] + model.T[h, s] * model.Beta[h, e, s] for h in model.H for e in model.E) +
               
                sum(model.O[j, h, s] + model.Alpha[j, h, s] * model.T[h, s] + model.Delta[j, h, s] / model.T[h, s]
                    for j in model.J for h in model.H)
            )
            for s in model.S
        )
        print(f"Cost Component: {cost_component}")

        quality_component = sum(
            model.P[s] * sum(model.X[h, j, s] * data["Bhj"][h, j] for h in model.H for j in model.J)
            for s in model.S
        )
        print(f"Quality Component: {quality_component}")

        collaboration_component = sum(
            model.P[s] * sum(data["Wij"][i, j] * model.X[h, i, s] * model.X[h, j, s]
                            for h in model.H for i in model.J for j in model.J)
            for s in model.S
        )
        print(f"Collaboration Component: {collaboration_component}")

        # Weighted sum of the three objectives
        return (
             weights["cost"] *cost_component +
             weights["quality"] *quality_component +
             weights["collaboration"] *collaboration_component
        )

    model.obj = Objective(rule=combined_objective_rule, sense=minimize)
    # ====== Constraints ======
    # Each activity must be assigned to exactly one agent
    def agent_constraint_rule(model, h, s):
        return sum(model.X[h, j, s] for j in model.J) == 1
    model.agent_constraint = Constraint(model.H, model.S, rule=agent_constraint_rule)

    def single_agent_per_activity_rule(model, h, s):
        return sum(model.X[h, j, s] for j in model.J) <= 1
    model.single_agent_per_activity = Constraint(model.H, model.S, rule=single_agent_per_activity_rule)
    # Each activity must have at least one equipment assigned
    def equipment_constraint_rule(model, h, s):
        return sum(model.Y[h, e, s] for e in model.E) >= 1
    model.equipment_constraint = Constraint(model.H, model.S, rule=equipment_constraint_rule)

    # Collaboration constraints for Z variables
    def z_constraint_rule_1(model, h, i, j, s):
        return model.Z[h, i, j, s] <= model.X[h, i, s]
    def z_constraint_rule_2(model, h, i, j, s):
        return model.Z[h, i, j, s] <= model.X[h, j, s]
    def z_constraint_rule_3(model, h, i, j, s):
        return model.Z[h, i, j, s] >= model.X[h, i, s] + model.X[h, j, s] - 1

    model.z_constraint_1 = Constraint(model.H, model.J, model.J, model.S, rule=z_constraint_rule_1)
    model.z_constraint_2 = Constraint(model.H, model.J, model.J, model.S, rule=z_constraint_rule_2)
    model.z_constraint_3 = Constraint(model.H, model.J, model.J, model.S, rule=z_constraint_rule_3)

    # Auxiliary variable constraints for cost modeling
    def aux_constraint_1(model, h, e, s):
        return model.Aux[h, e, s] <= data["Tmax"][h] * model.Y[h, e, s]
    def aux_constraint_2(model, h, e, s):
        return model.Aux[h, e, s] >= model.T[h, s] - (1 - model.Y[h, e, s]) * data["Tmax"][h]

    model.aux_constraint_1 = Constraint(model.H, model.E, model.S, rule=aux_constraint_1)
    model.aux_constraint_2 = Constraint(model.H, model.E, model.S, rule=aux_constraint_2)

    def time_min_rule(model, h, s):
        return model.T[h, s] >= data["Tmin"][h]
    def time_max_rule(model, h, s):
        return model.T[h, s] <= data["Tmax"][h]
    model.time_min_constraint = Constraint(model.H, model.S, rule=time_min_rule)
    model.time_max_constraint = Constraint(model.H, model.S, rule=time_max_rule)

    # Ù…Ø­Ø¯ÙˆØ¯ÛŒØª ØªØ®ØµÛŒØµ
    def single_agent_rule(model, h, s):
        return sum(model.X[h, j, s] for j in model.J) == 1
    model.single_agent_constraint = Constraint(model.H, model.S, rule=single_agent_rule)

    # Ù…Ø­Ø¯ÙˆØ¯ÛŒØª Ú©ÛŒÙÛŒØª
    def quality_constraint_rule(model, s):
        return sum(data["Dh"][h] * data["Bhj"][h, j] * model.X[h, j, s] for h in model.H for j in model.J) >= model.B
    def equipment_quality_constraint_rule(model, s):
        return sum(data["Rhe"][h, e] * model.Y[h, e, s] for h in model.H for e in model.E) >= model.R
    model.quality_constraint = Constraint(model.S, rule=quality_constraint_rule)
    model.equipment_quality_constraint = Constraint(model.S, rule=equipment_quality_constraint_rule)

    def collaboration_constraint_rule(model, h, i, j, s):
        if data["Wij"][i, j] < 0.3:
            return model.X[h, i, s] + model.X[h, j, s] <= 1
        return Constraint.Skip
    model.collaboration_constraint = Constraint(model.H, model.J, model.J, model.S, rule=collaboration_constraint_rule)



    M = 100000  
    model.budget_constraint = Constraint(expr=sum(model.Y[h, e, s] * model.G[h, e, s] for h in model.H for e in model.E for s in model.S) +
                                          sum(model.O[j, h, s] + model.Alpha[j, h, s] * model.T[h, s] for j in model.J for h in model.H for s in model.S) <= M)

    model.min_time_constraint = ConstraintList()
    for h in model.H:
        for s in model.S:
            model.min_time_constraint.add(model.T[h, s] >= data["Tmin"][h])

    model.max_time_constraint = ConstraintList()
    for h in model.H:
        for s in model.S:
            model.max_time_constraint.add(model.T[h, s] <= data['Tmax'][h])

    model.single_agent_constraint = ConstraintList()
    for h in model.H:
        for s in model.S:
            model.single_agent_constraint.add(sum(model.X[h, j, s] for j in J) == 1)

    Bmin = 5
    model.quality_agent_constraint = ConstraintList()
    for h in model.H:
        for s in model.S:
            model.quality_agent_constraint.add(sum(data["Bhj"][h, j] * model.X[h, j, s] for j in J) >= Bmin)

    Rmin = 2
    model.quality_equipment_constraint = ConstraintList()
    for h in model.H:
        for s in model.S:
            model.quality_equipment_constraint.add(sum(data['Rhe'][h, e] * model.Y[h, e, s] for e in E) >= Rmin)

    model.collaboration_constraint = ConstraintList()
    for h in model.H:
        for i in model.J:
            for j in model.J:
                for s in model.S:
                    if data["Wij"][i, j] > 0.3:
                        model.collaboration_constraint.add(model.X[h, i, s] + model.X[h, j, s] <= 2)

    # ====== Solve the Model ======

    solver =SolverFactory('ipopt', executable='/usr/local/bin/ipopt')



    results = solver.solve(model, tee=True)

    # Extract results and round values
    # ====== Retrieve Results ======
    x_results = {(h, j, s): round(model.X[h, j, s](), 2) for h in model.H for j in model.J for s in model.S}
    y_results = {(h, e, s): round(model.Y[h, e, s](), 2) for h in model.H for e in model.E for s in model.S}

    t_results = {(h, s): round(model.T[h, s].value) for h in model.H for s in model.S}
    model.Positive_Utilization = Constraint(
    model.H, model.E, model.S,
    rule=lambda model, h, e, s: model.Y[h, e, s] >= 0
    )
    return x_results, y_results, t_results
# Solve the optimization model
x_res, y_res, t_res = solve_model(data)

# Print or save results
print("Agent Assignment (X):", x_res)
print("Equipment Assignment (Y):", y_res)
print("Time Allocations:", t_res)
