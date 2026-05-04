import os
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from scipy import stats


root = tk.Tk()
root.withdraw()
print("Please select the folder containing the .dat file...")
dat_folder = filedialog.askdirectory(title="Select the folder containing the .dat file.")

if not dat_folder:
    print("No folder selected, program exits.")
    exit()

output_excel = os.path.join(dat_folder, "merged_data_with_log.xlsx")
print(f"\nSelected folder: {dat_folder}")
print(f"Results will be saved to: {output_excel}\n")


dat_files = sorted([f for f in os.listdir(dat_folder) if f.lower().endswith(".dat")])
if not dat_files:
    print("No .dat files found!")
    exit()
print(f"Found {len(dat_files)} .dat files")


bg_file = "0.dat"
bg_path = os.path.join(dat_folder, bg_file)

if not os.path.exists(bg_path):
    print(f"❌ Error: Background file {bg_file} not found, exiting program")
    exit()

print(f"✅ Background file found: 0.dat, will be used to subtract background from all data\n")

df_bg = pd.read_csv(
    bg_path,
    skiprows=5,
    sep=r'\s+',
    header=None,
    engine="python"
)
bg_second_col = df_bg.iloc[:, 1] 

all_columns = []

first_file = dat_files[0]
first_path = os.path.join(dat_folder, first_file)
df_first = pd.read_csv(
    first_path,
    skiprows=5,
    sep=r'\s+',
    header=None,
    engine="python"
)
col_first = df_first.iloc[:, 0]
all_columns.append(col_first)

for i, filename in enumerate(dat_files[1:], start=2):
    file_path = os.path.join(dat_folder, filename)
    print(f"Processing:{filename}(Background 0.dat has been removed)")

    df = pd.read_csv(
        file_path,
        skiprows=5,
        sep=r'\s+',
        header=None,
        engine="python"
    )
    

    data_corrected = df.iloc[:, 1] - bg_second_col
    all_columns.append(data_corrected)


sheet1_df = pd.concat(all_columns, axis=1)
sheet1_df.columns = [str(i) for i in range(sheet1_df.shape[1])]


print("\nCalculating log10 values...")
sheet1_float = sheet1_df.astype(float)
sheet1_float[sheet1_float <= 0] = np.nan
sheet2_log_df = np.log10(sheet1_float)


print("\nCalculating linear fit...")
#The range of values ​​for q needs to be changed.
start_row = XXX
end_row = XXX

x_all = sheet2_log_df.iloc[start_row:end_row, 0]

slopes = []
r_values = []
r2_values = []
fit_indices = []

for col_idx in range(1, sheet2_log_df.shape[1]):
    y_all = sheet2_log_df.iloc[start_row:end_row, col_idx]
    mask = ~(np.isnan(x_all) | np.isnan(y_all))
    x_data = x_all[mask]
    y_data = y_all[mask]

    if len(x_data) < 2:
        slopes.append(np.nan)
        r_values.append(np.nan)
        r2_values.append(np.nan)
        fit_indices.append(len(slopes))
        continue

    reg = stats.linregress(x_data, y_data)
    slopes.append(reg.slope)
    r_values.append(reg.rvalue)
    r2_values.append(reg.rvalue ** 2)
    fit_indices.append(len(slopes))

# Sheet3
sheet3_df = pd.DataFrame({
    "ID": fit_indices,
    "Slope": slopes,
    "Correlation Coefficient r": r_values,
    "Determination Coefficient R²": r2_values
})


with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
    sheet1_df.to_excel(writer, sheet_name="Sheet1", index=False)
    sheet2_log_df.to_excel(writer, sheet_name="Sheet2", index=False)
    sheet3_df.to_excel(writer, sheet_name="Sheet3", index=False)

print("\n✅ All processes completed!")
print(f"📊 Sheet1: Original data after removing the 0.dat background data.")
print(f"📊 Sheet2: log10 (data after background subtraction)")
print(f"📊 Sheet3：Linear fit slopes + correlation coefficients + R²")
print(f"📁 Save location:{output_excel}")
