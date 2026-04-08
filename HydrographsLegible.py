import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import timedelta
import os

# --- 1. SETUP ---
folder = r"C:\Users\ccexe\Downloads"

files = {
    "100%": "Hello_100.xlsx",
    "120%": "Hello_120.xlsx"
}
actual_sheets = ["L2", "L2_Burn", "L3", "L3_Burn", "L4", "L4_Burn"]

plot_data = {sheet: {} for sheet in actual_sheets}
observed_data = {sheet: None for sheet in actual_sheets}
datetime_col = {sheet: None for sheet in actual_sheets}
peak_times = {sheet: None for sheet in actual_sheets}

print("Extracting data...")

# --- 2. DATA EXTRACTION ---
for scenario, filename in files.items():
    file_path = os.path.join(folder, filename)
    if not os.path.exists(file_path): continue
        
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
                
            # Find the global peak time for the 120% scenario to set our baseline zoom
            if scenario == "120%":
                peak_idx = outflow.idxmax()
                peak_times[sheet] = df['Datetime'].iloc[peak_idx]
            
            # Wipe observed data for L2/J2
            if "L2" in sheet:
                observed_data[sheet] = None
            elif observed_data[sheet] is None and df.shape[1] > 3:
                obs = pd.to_numeric(df.iloc[:, 3], errors='coerce')
                if not obs.isna().all():
                    observed_data[sheet] = obs
                    
        except Exception as e:
            print(f"Error on {sheet}: {e}")

# --- SYNCHRONIZE THE ZOOM WINDOWS ---
for sheet in actual_sheets:
    if "_Burn" in sheet:
        normal_version = sheet.replace("_Burn", "")
        if peak_times[normal_version] is not pd.NaT:
            peak_times[sheet] = peak_times[normal_version]

# --- 3. PLOTTING THE ENVELOPES ---
print("Generating Locally Annotated Hydrographs...")
plt.style.use('ggplot')
fig, axes = plt.subplots(3, 2, figsize=(18, 22), constrained_layout=True)
axes = axes.flatten()

for i, sheet in enumerate(actual_sheets):
    ax = axes[i]
    display_name = sheet.replace("L", "J")
    
    x_data = datetime_col[sheet]
    
    if x_data is None or x_data.isna().all() or pd.isna(peak_times[sheet]):
        ax.set_title(f"{display_name} - No Data", fontweight='normal')
        continue

    # 1. Plot Observed Flow
    if observed_data[sheet] is not None:
        ax.plot(x_data, observed_data[sheet], color='black', linewidth=1.5, 
                label='Observed Flow', linestyle='--', zorder=5)

    # 2. Plot Baseline 100% 
    if "100%" in plot_data[sheet]:
        ax.plot(x_data, plot_data[sheet]["100%"], color='#1f77b4', linewidth=1.5, 
                label='Baseline (100%)', zorder=4)
        
    # 3. Plot Worst-Case 120% 
    if "120%" in plot_data[sheet]:
        ax.plot(x_data, plot_data[sheet]["120%"], color='#d62728', linewidth=1.5, 
                label='Extreme Climate (120%)', zorder=4)

    # 4. Fill the "Climate Envelope" 
    if "100%" in plot_data[sheet] and "120%" in plot_data[sheet]:
        ax.fill_between(x_data, plot_data[sheet]["100%"], plot_data[sheet]["120%"], 
                        color='#d62728', alpha=0.15, label='Climate Impact Zone', zorder=3)

    # --- CALCULATE THE ZOOM WINDOW FIRST ---
    peak = peak_times[sheet]
    start_zoom = peak - timedelta(days=2)
    end_zoom = peak + timedelta(days=2)
    ax.set_xlim(start_zoom, end_zoom)

    # --- THE LOCAL PEAK ANNOTATION FIX ---
    if "120%" in plot_data[sheet]:
        # Filter the data to ONLY look inside the start_zoom and end_zoom window
        mask = (x_data >= start_zoom) & (x_data <= end_zoom)
        local_120_data = plot_data[sheet]["120%"][mask]
        
        if not local_120_data.empty:
            local_max_val = local_120_data.max()
            local_max_idx = local_120_data.idxmax()
            local_max_time = x_data.loc[local_max_idx]
            
            # Draw the annotation pointing to the LOCAL peak
            ax.annotate(f"Max: {local_max_val:.1f} cms\n{local_max_time.strftime('%d %b %H:00')}", 
                        xy=(local_max_time, local_max_val),
                        xytext=(0, 20),
                        textcoords='offset points',
                        ha='center', va='bottom',
                        fontsize=9, fontweight='normal', color='#660000',
                        arrowprops=dict(arrowstyle='->', color='#660000', lw=1.0))
    
    # Standardize Y-Axis to 55
    ax.set_ylim(0, 55)
    
    # Clean X-Axis formatting 
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d %b\n%H:00'))
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=12)) 
    
    # Labels 
    ax.set_title(f"Hydrograph: {display_name}", fontsize=15, fontweight='normal', pad=15)
    ax.set_ylabel("Outflow (cms)", fontweight='normal')
    
    # Legend
    handles, labels = ax.get_legend_handles_labels()
    ax.legend(handles, labels, loc='upper right', fontsize='small', framealpha=0.9)
    
    ax.grid(True, linestyle=':', alpha=0.7)


output_path = os.path.join(folder, "Final_Local_Peak_Envelopes.png")
plt.savefig(output_path, dpi=300, bbox_inches='tight')
plt.show()

print(f"Success! Local Peak Envelopes saved to: {output_path}")