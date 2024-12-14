import pandas as pd
import numpy as np
from pyomo.environ import *

# ====== Step 1: Generate Random Input Data ======
def generate_input_data(filename, num_h=10, num_j=25, num_e=15, num_s=4):
    """
    Generates random input data and saves it to an Excel file.
    :param filename: Name of the file to save data.
    :param num_h: Number of activities.
    :param num_j: Number of agents.
    :param num_e: Number of equipment types.
    :param num_s: Number of scenarios.
    """
    # Randomly generate input parameters
    data = {
        "Dh": np.random.randint(1, 10, size=num_h),  # Importance of each activity
        "Tmin": np.random.randint(1, 5, size=num_h),  # Minimum time required for activities
        "Tmax": np.random.randint(5, 10, size=num_h),  # Maximum time available for activities
        "Kh": np.random.randint(50, 100, size=num_h),  # Fixed daily cost for activities
        "Bhj": np.random.randint(10, 100, size=(num_h, num_j)),  # Quality scores for agents
        "Rhe": np.random.randint(1, 10, size=(num_h, num_e)),  # Quality scores for equipment
        "Wij": np.random.uniform(0, 1, size=(num_j, num_j)),  # Collaboration factor between agents
        "G": np.random.randint(100, 200, size=(num_h, num_e, num_s)),  # Fixed cost for equipment
        "Beta": np.random.randint(10, 20, size=(num_h, num_e, num_s)),  # Variable cost factor
    }

    # Save generated data into an Excel file with multiple sheets
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    pd.DataFrame(data["Dh"], columns=["Dh"]).to_excel(writer, sheet_name="Dh", index=False)
    pd.DataFrame(data["Tmin"], columns=["Tmin"]).to_excel(writer, sheet_name="Tmin", index=False)
    pd.DataFrame(data["Tmax"], columns=["Tmax"]).to_excel(writer, sheet_name="Tmax", index=False)
    pd.DataFrame(data["Kh"], columns=["Kh"]).to_excel(writer, sheet_name="Kh", index=False)
    pd.DataFrame(data["Bhj"]).to_excel(writer, sheet_name="Bhj", index=False)
    pd.DataFrame(data["Rhe"]).to_excel(writer, sheet_name="Rhe", index=False)
    pd.DataFrame(data["Wij"]).to_excel(writer, sheet_name="Wij", index=False)
    for s in range(num_s):
        pd.DataFrame(data["G"][:, :, s]).to_excel(writer, sheet_name=f"G_Scenario_{s+1}", index=False)
        pd.DataFrame(data["Beta"][:, :, s]).to_excel(writer, sheet_name=f"Beta_Scenario_{s+1}", index=False)
    writer.close()

# Generate the input data
generate_input_data("Generated_data.xlsx")
