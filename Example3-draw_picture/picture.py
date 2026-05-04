import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from tkinter import Tk, filedialog
import matplotlib

matplotlib.use('TkAgg')


plt.rcParams['axes.linewidth'] = 1.0
plt.rcParams['xtick.direction'] = 'out'
plt.rcParams['ytick.direction'] = 'out'
plt.rcParams['xtick.major.size'] = 8
plt.rcParams['ytick.major.size'] = 8
plt.rcParams['xtick.major.width'] = 1.0
plt.rcParams['ytick.major.width'] = 1.0


root = Tk()
root.withdraw()
file_path = filedialog.askopenfilename(
    title="Please select an xlsx file.",
    filetypes=[("Excel files", "*.xlsx")]
)

if not file_path:
    print("No file selected, program exits.")
    exit()


df = pd.read_excel(file_path, sheet_name=0)  
data = df.iloc[:, 1:].values


data_selected = data[:-1, :]
n_rows, n_cols = data_selected.shape
print(f"Data after truncation:{n_rows} row × {n_cols} column")


width = 6
height = 14
fig, ax = plt.subplots(figsize=(width, height), dpi=150)


vmin_val = np.percentile(data_selected, 1)
vmax_val = np.percentile(data_selected, 99)

im = ax.imshow(
    data_selected.T,
    cmap='coolwarm',
    aspect='auto',
    rasterized=True,
    vmin=vmin_val,
    vmax=vmax_val
)
ax.invert_yaxis()


x_ticks = np.linspace(0, n_rows-1, 3)
ax.set_xticks(x_ticks)
ax.set_xticklabels([f'{int(t)}' for t in x_ticks])

y_ticks = np.linspace(0, n_cols-1, 5)
ax.set_yticks(y_ticks)
ax.set_yticklabels([])


ax.set_position([0.05, 0.05, 0.90, 0.90])


cbar = fig.colorbar(
    im, ax=ax,
    location="right",
    anchor=(1.3, 0.5),
    shrink=0.75,
    pad=0.25
)
cbar.set_ticks([])
cbar.ax.set_yticklabels([])
cbar.set_label('')


def hover(event):
    if event.inaxes == ax:
        x = int(round(event.xdata)) if event.xdata is not None else -1
        y = int(round(event.ydata)) if event.ydata is not None else -1
        if 0 <= x < n_rows and 0 <= y < n_cols:
            val = data_selected[x, y]
            fig.canvas.toolbar.set_message(f"X={x} | Y={y} | Value={val:.4f}")

fig.canvas.mpl_connect("motion_notify_event", hover)

plt.show()
