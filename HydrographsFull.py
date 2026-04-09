import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import os

# --- 1. SETUP ---
folder = r"C:\Users\ccexe\Downloads"

files = {
    "100%": ("Hello_100.xlsx", "#1f77b4"),  # Solid Blue
    "120%": ("Hello_120.xlsx", "#d62728")   # Solid Red
}
actual_sheets = ["L2", "L2_Burn", "L3", "L3_Burn", "L4", "L4_Burn"]

plot_data = {sheet: {} for sheet in actual_sheets}
observed_data = {sheet: None for sheet in actual_sheets}
datetime_col = {sheet: None for sheet in actual_sheets}

print("Extracting full time-series data...")

# --- 2. DATA EXTRACTION ---
for scenario, (filename, color) in files.items():
    file_path = os.path.join(folder, filename)
    if not os.path.exists(file_path): 
        print(f"Skipping {filename}: File not found.")
        continue
        
    for sheet in actual_sheets:
        try:
            df = pd.read_excel(file_path, sheet_name=sheet)
            
            dates = df.iloc[:, 0].astype(str).str.split().str[0] 
            times = df.iloc[:, 1].astype(str)
            df['Datetime'] = pd.to_datetime(dates + ' ' + times, errors='coerce')
            
            outflow = pd.to_numeric(df.iloc[:, 2], errors='coerce')
            plot_data[sheet][scenario] = outflow
            
            if datetime_col[sheet] is None:
                datetime_col[sheet] = df['Datetime']
            
            # Explicitly wipe observed data for L2/J2
            if "L2" in sheet:
                observed_data[sheet] = None
            elif observed_data[sheet] is None and df.shape[1] > 3:
                obs = pd.to_numeric(df.iloc[:, 3], errors='coerce')
                if not obs.isna().all():
                    observed_data[sheet] = obs
                    
        except Exception as e:
            print(f"Error on {filename}, sheet {sheet}: {e}")

# --- 3. PLOTTING THE HYDROGRAPHS ---
print("Generating Hydrographs with Triple Annotations...")
plt.style.use('ggplot')
fig, axes = plt.subplots(3, 2, figsize=(18, 22), constrained_layout=True)
axes = axes.flatten()

for i, sheet in enumerate(actual_sheets):
    ax = axes[i]
    display_name = sheet.replace("L", "J")
    
    x_data = datetime_col[sheet]
    
    if x_data is None or x_data.isna().all():
        ax.set_title(f"{display_name} - No Data Found", fontweight='normal')
        continue

    # 1. Plot 120% Scenario (Back, thin line)
    if "120%" in plot_data[sheet]:
        ax.plot(x_data, plot_data[sheet]["120%"], color=files["120%"][1], linewidth=0.8, 
                label='Simulated 120%', zorder=3)

    # 2. Plot 100% Scenario (Middle, medium line)
    if "100%" in plot_data[sheet]:
        ax.plot(x_data, plot_data[sheet]["100%"], color=files["100%"][1], linewidth=1.0, 
                label='Simulated 100%', zorder=4)

    # 3. Plot Observed Flow (Front, dashed black line)
    if observed_data[sheet] is not None:
        ax.plot(x_data, observed_data[sheet], color='black', linewidth=1.2, 
                label='Observed Flow', linestyle='--', zorder=5)

    # --- THE TRIPLE PEAK ANNOTATIONS ---
    
    # Label 1: 120% Peak (Points UP)
    if "120%" in plot_data[sheet]:
        max_val_120 = plot_data[sheet]["120%"].max()
        max_idx_120 = plot_data[sheet]["120%"].idxmax()
        max_time_120 = x_data.iloc[max_idx_120]
        
        ax.annotate(f"120% Max: {max_val_120:.1f} cms", 
                    xy=(max_time_120, max_val_120),
                    xytext=(0, 25), # 25 points straight UP
                    textcoords='offset points',
                    ha='center', va='bottom',
                    fontsize=8, fontweight='normal', color='#660000',
                    arrowprops=dict(arrowstyle='->', color='#660000', lw=1.0))

    # Label 2: 100% Peak (NOW POINTS RIGHT)
    if "100%" in plot_data[sheet]:
        max_val_100 = plot_data[sheet]["100%"].max()
        max_idx_100 = plot_data[sheet]["100%"].idxmax()
        max_time_100 = x_data.iloc[max_idx_100]
        
        ax.annotate(f"100% Max: {max_val_100:.1f} cms", 
                    xy=(max_time_100, max_val_100),
                    xytext=(45, 0), # 45 points RIGHT
                    textcoords='offset points',
                    ha='left', va='center',
                    fontsize=8, fontweight='normal', color='#1f77b4',
                    arrowprops=dict(arrowstyle='->', color='#1f77b4', lw=1.0))

    # Label 3: Observed Peak (Points RIGHT) - Skips J2 automatically
    if observed_data[sheet] is not None:
        max_val_obs = observed_data[sheet].max()
        max_idx_obs = observed_data[sheet].idxmax()
        max_time_obs = x_data.iloc[max_idx_obs]
        
        ax.annotate(f"Obs Max: {max_val_obs:.1f} cms", 
                    xy=(max_time_obs, max_val_obs),
                    xytext=(45, 0), # 45 points RIGHT
                    textcoords='offset points',
                    ha='left', va='center',
                    fontsize=8, fontweight='normal', color='black',
                    arrowprops=dict(arrowstyle='->', color='black', lw=1.0))

    # Standardize Y-Axis (Bumped to 60 to ensure the UP arrow has plenty of sky)
    ax.set_ylim(0, 60)
    
    # Format the X-Axis
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d %b'))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=14)) 
    
    # Text and Labels
    ax.set_title(f"Hydrograph: {display_name}", fontsize=15, fontweight='normal', pad=15)
    ax.set_ylabel("Outflow (cms)", fontweight='normal')
    
    # Legend Sorting (Observed on top, 100% middle, 120% bottom)
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles[::-1], labels[::-1], loc='upper right', fontsize='small', framealpha=0.9)
    
    ax.grid(True, linestyle=':', alpha=0.7)


output_path = os.path.join(folder, "Final_Triple_Annotated_Hydrographs.png")
plt.savefig(output_path, dpi=300, bbox_inches='tight')
plt.show()

print(f"Success! Triple-Annotated hydrographs saved to: {output_path}")
