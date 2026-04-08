import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import os

# --- 1. SETUP ---
folder = r"C:\Users\ccexe\Downloads"

# Stripped down to just the 100% baseline and 120% extreme
files = {
    "100%": ("Hello_100.xlsx", "#1f77b4"),  # Solid Blue
    "120%": ("Hello_120.xlsx", "#d62728")   # Solid Red
}
actual_sheets = ["L2", "L2_Burn", "L3", "L3_Burn", "L4", "L4_Burn"]

# Storage dictionaries
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
            
            # Combine Date and Time
            dates = df.iloc[:, 0].astype(str).str.split().str[0] 
            times = df.iloc[:, 1].astype(str)
            df['Datetime'] = pd.to_datetime(dates + ' ' + times, errors='coerce')
            
            outflow = pd.to_numeric(df.iloc[:, 2], errors='coerce')
            plot_data[sheet][scenario] = outflow
            
            if datetime_col[sheet] is None:
                datetime_col[sheet] = df['Datetime']
            
            # --- FIX 1: Explicitly wipe observed data for L2/J2 ---
            if "L2" in sheet:
                observed_data[sheet] = None
            elif observed_data[sheet] is None and df.shape[1] > 3:
                obs = pd.to_numeric(df.iloc[:, 3], errors='coerce')
                if not obs.isna().all():
                    observed_data[sheet] = obs
                    
        except Exception as e:
            print(f"Error on {filename}, sheet {sheet}: {e}")

# --- 3. PLOTTING THE HYDROGRAPHS ---
print("Generating Hydrographs...")
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

    # Plot Observed Flow first 
    if observed_data[sheet] is not None:
        ax.plot(x_data, observed_data[sheet], color='black', linewidth=1.5, 
                label='Observed Flow', linestyle='--', zorder=5)

    # Plot Simulated Flows 
    for scenario, (filename, color) in files.items():
        if scenario in plot_data[sheet]:
            ax.plot(x_data, plot_data[sheet][scenario], color=color, linewidth=1.5, 
                    label=f'Simulated {scenario}', zorder=4)

    # --- FIX 2: PEAK LABELS ---
    if "120%" in plot_data[sheet]:
        # Find the absolute maximum on the 120% curve
        max_val = plot_data[sheet]["120%"].max()
        max_idx = plot_data[sheet]["120%"].idxmax()
        max_time = x_data.iloc[max_idx]
        
        # Add the annotation arrow (Strictly normal font)
        ax.annotate(f"Max: {max_val:.1f} cms\n{max_time.strftime('%d %b')}", 
                    xy=(max_time, max_val),
                    xytext=(0, 20),
                    textcoords='offset points',
                    ha='center', va='bottom',
                    fontsize=9, fontweight='normal', color='#660000',
                    arrowprops=dict(arrowstyle='->', color='#660000', lw=1.0))

    # Standardize Y-Axis to 55 to leave room for the new annotations
    ax.set_ylim(0, 55)
    
    # Format the X-Axis
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d %b'))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=14)) 
    
    # Text and Labels (Strictly NO bold)
    ax.set_title(f"Hydrograph: {display_name}", fontsize=15, fontweight='normal', pad=15)
    ax.set_ylabel("Outflow (cms)", fontweight='normal')
    ax.legend(loc='upper right', fontsize='small', framealpha=0.9)
    
    ax.grid(True, linestyle=':', alpha=0.7)


# Save the final image
output_path = os.path.join(folder, "Final_Clean_Hydrographs.png")
plt.savefig(output_path, dpi=300, bbox_inches='tight')
plt.show()

print(f"Success! Hydrographs saved to: {output_path}")