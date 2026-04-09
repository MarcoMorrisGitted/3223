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

# The exact sheet groupings you requested (4 per location)
target_locations = {
    "L2": ["L2", "L2_Burn79", "L2_Burn41", "L2_Burn50Imp"],
    "L3": ["L3", "L3_Burn79", "L3_Burn41", "L3_Burn50Imp"],
    "L4": ["L4", "L4_Burn79", "L4_Burn41", "L4_Burn50Imp"]
}

# Flatten the list so the extraction engine grabs everything in one pass
actual_sheets = [sheet for sheets in target_locations.values() for sheet in sheets]

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

# --- 3. PLOTTING THE HYDROGRAPHS (3 PAGES, 4 GRAPHS PER PAGE) ---
print("Generating separate pages for J2, J3, and J4...")
plt.style.use('ggplot')

# Loop through each location to create its own page
for loc, loc_sheets in target_locations.items():
    display_loc = loc.replace("L", "J") # Turns "L3" into "J3" for the main page title
    
    # FORCE 4 GRAPHS PER PAGE (2x2 Grid)
    fig, axes = plt.subplots(2, 2, figsize=(16, 12), constrained_layout=True)
    axes = axes.flatten()

    for i, sheet in enumerate(loc_sheets):
        ax = axes[i]
        
        # --- NEW CUSTOM TITLE LOGIC ---
        base_node = loc.replace("L", "J_") # Turns "L3" into "J_3" for the subplot titles
        
        if "Burn79" in sheet:
            custom_title = f"Post-Fire {base_node} (CN79)"
        elif "Burn41" in sheet:
            custom_title = f"Post-Fire {base_node} (CN41)"
        elif "50Imp" in sheet:
            custom_title = f"Post-Fire {base_node} (50% Imp)"
        else:
            custom_title = f"Baseline {base_node}"
        
        x_data = datetime_col[sheet]
        
        if x_data is None or x_data.isna().all():
            ax.set_title(f"{custom_title} - No Data Found", fontweight='normal')
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
                        xytext=(0, 25), 
                        textcoords='offset points',
                        ha='center', va='bottom',
                        fontsize=8, fontweight='normal', color='#660000',
                        arrowprops=dict(arrowstyle='->', color='#660000', lw=1.0))

        # Label 2: 100% Peak (Points RIGHT)
        if "100%" in plot_data[sheet]:
            max_val_100 = plot_data[sheet]["100%"].max()
            max_idx_100 = plot_data[sheet]["100%"].idxmax()
            max_time_100 = x_data.iloc[max_idx_100]
            
            ax.annotate(f"100% Max: {max_val_100:.1f} cms", 
                        xy=(max_time_100, max_val_100),
                        xytext=(45, 0), 
                        textcoords='offset points',
                        ha='left', va='center',
                        fontsize=8, fontweight='normal', color='#1f77b4',
                        arrowprops=dict(arrowstyle='->', color='#1f77b4', lw=1.0))

        # Label 3: Observed Peak (Points RIGHT)
        if observed_data[sheet] is not None:
            max_val_obs = observed_data[sheet].max()
            max_idx_obs = observed_data[sheet].idxmax()
            max_time_obs = x_data.iloc[max_idx_obs]
            
            ax.annotate(f"Obs Max: {max_val_obs:.1f} cms", 
                        xy=(max_time_obs, max_val_obs),
                        xytext=(45, 0), 
                        textcoords='offset points',
                        ha='left', va='center',
                        fontsize=8, fontweight='normal', color='black',
                        arrowprops=dict(arrowstyle='->', color='black', lw=1.0))

        # Standardize Y-Axis
        ax.set_ylim(0, 60)
        
        # Format the X-Axis
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%d %b'))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=14)) 
        
        # Apply the new custom title
        ax.set_title(custom_title, fontsize=15, fontweight='normal', pad=15)
        ax.set_ylabel("Outflow (cms)", fontweight='normal')
        
        # Legend Sorting 
        handles, labels = ax.get_legend_handles_labels()
        ax.legend(handles[::-1], labels[::-1], loc='upper right', fontsize='small', framealpha=0.9)
        
        ax.grid(True, linestyle=':', alpha=0.7)

    # Master Title for the specific page
    plt.suptitle(f"Watershed Hydrographs: {display_loc} Scenarios", fontsize=22, fontweight='normal', y=1.03)

    # Save the specific page
    output_path = os.path.join(folder, f"Final_Annotated_4Grid_{display_loc}.png")
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.close(fig)

print("Success! 3 separate images with updated titles saved to your Downloads folder.")
