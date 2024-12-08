import pandas as pd
import numpy as np

def generate_input_data(filename, num_h=10, num_j=25, num_e=15, num_s=4):
    """
    Generates random input data and saves it to an Excel file.
    :param filename: Name of the file to save data.
    :param num_h: Number of activities.
    :param num_j: Number of agents.
    :param num_e: Number of equipment types.
    :param num_s: Number of scenarios.
    """
    data = {
        "Dh": np.random.randint(1, 10, size=num_h),
        "Tmin": np.random.randint(1, 5, size=num_h),
        "Tmax": np.random.randint(5, 10, size=num_h),
        "Kh": np.random.randint(50, 100, size=num_h),
        "Bhj": np.random.randint(1, 10, size=(num_h, num_j)),
        "Rhe": np.random.randint(1, 10, size=(num_h, num_e)),
        "Wij": np.random.uniform(0, 1, size=(num_j, num_j)),
        "G": np.random.randint(100, 200, size=(num_h, num_e, num_s)),
        "Beta": np.random.randint(10, 20, size=(num_h, num_e, num_s)),
    }

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

if __name__ == "__main__":
    generate_input_data("Generated_data.xlsx")
