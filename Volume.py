import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# --- 1. SETUP ---
folder = r"C:\Users\ccexe\Downloads"
file_path = os.path.join(folder, "globalsums.xlsx")

scenarios = ["100", "120"]
conditions = [
    {"suffix": "", "label": "Pre-Fire (Baseline)", "color": "#1f77b4"},
    {"suffix": "", "label": "Post-Fire (CN79)", "color": "#ff7f0e"},
    {"suffix": "41", "label": "Post-Fire (CN 41)", "color": "#d62728"},
    {"suffix": "Imp50", "label": "Post-Fire (Imp 50%)", "color": "#9467bd"}
]

data_matrix = [] 

print("Extracting Sink-1 (Total Watershed) True Volume...")
print("-" * 50)

# --- 2. DATA EXTRACTION ---
for sc in scenarios:
    sc_results = []
    for cond in conditions:
        prefix = "Pre" if "Pre-Fire" in cond["label"] else "Post"
        sheet_name = f"{prefix}{sc}{cond['suffix']}"
        
        try:
            # header=None ensures Python reads the first row properly
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
            
            # Find ONLY the Sink-1 outfall
            sink_row = df[df.iloc[:, 0].astype(str).str.strip() == 'Sink-1']
            
            if not sink_row.empty:
                area_km2 = pd.to_numeric(sink_row.iloc[0, 1], errors='coerce')
                depth_mm = pd.to_numeric(sink_row.iloc[0, 4], errors='coerce')
                
                # Convert to True Cubic Meters
                if pd.notna(area_km2) and pd.notna(depth_mm):
                    total_vol_m3 = area_km2 * depth_mm * 1000
                else:
                    total_vol_m3 = 0
                    
                print(f"[{sheet_name}] Sink-1 Found -> Depth: {depth_mm}mm | Total Volume: {total_vol_m3:,.0f} m³")
            else:
                total_vol_m3 = 0
                print(f"[{sheet_name}] ERROR: 'Sink-1' not found in sheet!")
                
            sc_results.append(total_vol_m3)
            
        except Exception as e:
            print(f"[{sheet_name}] Warning: Could not read sheet. Error: {e}")
            sc_results.append(0)
            
    data_matrix.append(sc_results)

print("-" * 50)

# --- 3. PLOTTING ---
max_val = max([max(row) for row in data_matrix]) if data_matrix else 0
if max_val == 0:
    print("CRITICAL ERROR: All extracted volumes are 0. Plot will not generate.")
else:
    print("Generating comparison chart...")
    plt.style.use('ggplot')
    fig, ax = plt.subplots(figsize=(14, 8), constrained_layout=True)

    x = np.arange(len(scenarios))
    width = 0.2 

    for i, cond in enumerate(conditions):
        bar_heights = [sc_data[i] for sc_data in data_matrix]
        offset = (i - 1.5) * width
        bars = ax.bar(x + offset, bar_heights, width, label=cond["label"], color=cond["color"])
        
        for bar in bars:
            yval = bar.get_height()
            if yval > 0:
                ax.text(bar.get_x() + bar.get_width()/2, yval + (max_val * 0.01),
                        f'{yval:,.0f}', ha='center', va='bottom', 
                        fontsize=9, fontweight='normal')

    ax.set_ylabel("Total Sink-1 Watershed Outfall (m³)", fontsize=12, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels([f"{sc}% Climate Scenario" for sc in scenarios], fontsize=11)
    
    # Adjust Y-limit based on actual max value
    ax.set_ylim(0, max_val * 1.15)
    
    ax.legend(loc='upper left', fontsize=10, framealpha=0.9)
    ax.grid(axis='y', linestyle='--', alpha=0.7)

    output_path = os.path.join(folder, "Sensitivity_Volume_Analysis_Sink1.png")
    plt.savefig(output_path, dpi=300, bbox_inches='tight')
    plt.show()

    print(f"Success! Sink-1 chart saved to: {output_path}")
