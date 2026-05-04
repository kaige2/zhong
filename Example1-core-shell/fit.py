# Nomenclature
# A pre-exponential factor, min−1
# C concentration of gas A inside the particle, mol/m3
# C_e equilibrium concentration of gas A, mol/m3
# C_ion concentration of solid reactant ions on the surface of solid reactant B, mol/m3
# D_P product layer diffusion coefficient, m2/s
# D_ion ion diffusivity, m2/s
# D_S surface diffusion coefficient, m/s
# E apparent activation energy, J/mol
# f(X) rate-limiting model function,
# g_"D"  (X) geometric model function
# K_s ratio of the chemical reaction rate to the surface diffusion rate
# k_s chemical reaction rate constant, m4/mol/s
# M_CaO molar mass of CaO, 56 g/mol
# M_(CO_2 ) molar mass of CO2, 44 g/mol
# m mass of the carbonated slag, g
# m_0 the initial mass of the slag, g
# p_D (δ,X) geometric model function
# R The universal gas constant, 8.314 J·mol−1·K−1
# r0 initial pore radius, m
# r1 shell radius, m
# r2 core radius or pore radius at time t, m
# S_0 initial surface area of CaO per unit volume of the solid particle, m2/m3
# T thermodynamic temperature, K
# ν_B volume of solid reactant B, m3/m3
# V_B^M molar volume of solid reactant B, m3/mol
# ν_P stoichiometric coefficient of product P
# V_P^M molar volume of solid product, m3/mol
# X fractional CaO conversion
# Z ratio of the molar volume of the solid product to that of the solid reactant
# δ fraction of unoccupied CaO area
# ε_0 initial porosity
# β model parameter
# β^' constant heating rate, β′ = dT/dt = constant, K/min
# ζ_c critical thickness of the product layer, m
# κ_"S"  (X) geometric model function
# ψ pore structure parameter，ψ=4πL_0 (1-V_0 )/S_0^2

import numpy as np
import time
import pandas as pd
from scipy.integrate import solve_ivp
from scipy.optimize import curve_fit
import matplotlib.pyplot as plt
from sklearn.metrics import r2_score, mean_squared_error
import math
import scipy.stats as stats


# === Read experimental data (from Excel file) ===
file_path = 'data.xlsx'  # Excel file path
sheet_name = 'Sheet1'     # Worksheet name
col_time = 't'            # Time column name
col_conversion = 'X'      # Conversion rate column name

# === Define auxiliary functions ===
def calculate_Ce(T):
    """Calculate the equilibrium concentration Ce """
    R = 8.314  # Gas constant (J/mol·K)
    return (1.826e6 / R) * np.exp(-19680 / T)

# === Fixed parameter settings ===
# xxx needs to be replaced with the actual numerical value
Z = xxx          # Fixed volume ratio parameter
S_0 = xxx      # Fixed specific surface area 
epsilon_0 = xxx   # Fixed porosity
T = xxx          # Fixed reaction temperature (K)
C = xxx          # Fixed CO₂ concentration
C_e = calculate_Ce(T)

# === Initial guess for fitting parameters ===
a= 1e-13  # Initial guess for k_s
b= 1e-10  # Initial guess for D_s
c= 1e-20 # Initial guess for D_p
beta= a*(1-epsilon_0)/(26.20e-6*c*S_0)      # Initial guess for beta
print("The beta value is:",beta)
params_initial = [a,b,beta]  # [k_s, D_s, beta]


def model(t, y, k_s, D_s, beta, Z, S_0, epsilon_0, T, C):
    """Differential equation system: d δ/dt and dX/dt"""
    X = y  # State variable
    C_e = calculate_Ce(T)
    # Calculate dδ/dt
    term1 = (C - C_e) / (1 + k_s*Z/D_s* (C - C_e))
    delta=math.exp(- (k_s * Z / D_s) * term1 * t)
    # Calculate dX/dt
    kappa_S_val = (1 - X)**(2/3)
    g_D_val = (1 - X)**(2/3) / (1 + (Z / (1 - delta) - 1) * X)**(2/3)
    p_D_val = 3 * (1 - X)**(1/3) * ((1 + (Z / (1 - delta) - 1) * X)**(1/3) - (1 - X)**(1/3)) / (1 + (Z / (1 - delta) - 1) * X)**(1/3)
    denominator = (delta / (1 + k_s*Z/D_s* (C - C_e))) + ((1 - delta) / (g_D_val + beta * (C - C_e) * p_D_val))
    dX_dt = (k_s * S_0 / (1 - epsilon_0)) * kappa_S_val * (C - C_e) * denominator
    # Return the derivative of a system of differential equations 
    return dX_dt

def calculate_delta(k_s, D_s, Z, t):
    delta=[]
    for i in t:
        term1 = (C - C_e) / (1 + k_s*Z/D_s * (C - C_e))
        delta.append(math.exp(- (k_s * Z / D_s) * term1 * i))
    return delta

# === Define fitting function ===
def fit_model(t_data, X_data, params_initial, Z, S_0, epsilon_0, T, C):
    """Fit k_s, D_s, β, with other parameters fixed"""
    print("Step 1")
    # After solve_ivp receives the passed t_data and X_data, it performs integration operations,
    # and the results of the differential operations [d/dt, dX/dt] are used for fitting
    popt, pcov = curve_fit(
        lambda t, *params: solve_ivp(
            lambda t_inner, y_inner: model(t_inner, y_inner, *params, Z, S_0, epsilon_0, T, C),
            [t.min(), t.max()], [0], t_eval=t
        ).y[0],
        t_data, X_data, p0=params_initial
    )
    print("Step 2")
    return popt, pcov

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    t_data = df[col_time].values
    X_data = df[col_conversion].values
    print("Data read successfully!")
    print("Preview of the first 5 rows:")
    print(df.head())
except FileNotFoundError:
    print(f"Error: File '{file_path}' not found. Please check the path.")
    exit()
except KeyError as e:
    print(f"Error: Column '{e.args[0]}' not found in Excel file. Please check the column names.")
    exit()

# === Write to the third column of Excel ===
def update_excel_column(file_path,table_name,column_index, header, data):
    """
    Write the array data to the specified column of the specified worksheet in an Excel file, overwrite the original content, and set a new header.

    Parameters:
    -File_cath (str): Excel file path.
    -Column_index (int): Target column index (starting from 0).
    -Header (str): The new column header.
    -Data (list): A list of data to be written.
    """
    # Read Sheet1 from Excel
    df = pd.read_excel(file_path, sheet_name=table_name)
    # Delete the original third column (by column index)
    old_column_name = df.columns[column_index]
    df = df.drop(columns=[old_column_name])
    # Construct a new DataFrame for the new column
    new_col_df = pd.DataFrame({header: data})
    # Split the original DataFrame into left and right parts
    left_df = df.iloc[:, :column_index]     # Original first two columns
    right_df = df.iloc[:, column_index:]    # Original fourth column and beyond
    # Merge into a new DataFrame: left + new_col + right
    df_new = pd.concat([left_df, new_col_df, right_df], axis=1)
    # Write back to Excel file (only update Sheet1)
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_new.to_excel(writer, sheet_name='Sheet1', index=False)


# === Perform fitting ===
popt, pcov = fit_model(t_data, X_data, params_initial, Z, S_0, epsilon_0, T, C)
D_p=popt[0]*(1-epsilon_0)/(26.20e-6*popt[2]*S_0)
print("Optimize the parameters k_s, D_s, and beta respectively:", popt)
print("Optimize pD parameters to",D_p)
print("Covariance matrix:\n", pcov)

# === Use optimized parameters to predict conversion rate ===
k_s_opt, D_s_opt, beta_opt = popt
sol = solve_ivp(
    lambda t_inner, y_inner: model(t_inner, y_inner, k_s_opt, D_s_opt, beta_opt, Z, S_0, epsilon_0, T, C),
    [t_data[0], t_data[-1]], [0],
    t_eval=t_data,
    method='LSODA'
)
delta2=calculate_delta(popt[0], popt[1], Z, t_data)#计算theta

# === Plot fitting results ===
plt.figure(figsize=(10, 6))
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus']=False #Used to display negative signs normally
plt.scatter(t_data, X_data, label='experimental data', color='blue', s=60)
plt.plot(sol.t, sol.y[0], 'r-', label='fitted curve', linewidth=2) 
plt.xlabel('Time (t)')
plt.ylabel('Conversion Rate (X)')
plt.title('CaO Carbonation Reaction Kinetics Fitting (Fitting k_s, D_s, β Only)')
plt.legend()
plt.grid(True)
plt.show()

# === Evaluate fitting results ===
if __name__ == "__main__":
    X_pred = sol.y[0]
    r2 = r2_score(X_data, X_pred)
    mse = mean_squared_error(X_data, X_pred)
    print(f"R2 value: {r2:.6f}")
    print(f"MSE: {mse:.6f}")
