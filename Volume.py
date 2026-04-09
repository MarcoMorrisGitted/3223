import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# --- 1. SETUP ---
folder = r"C:\Users\ccexe\Downloads"
file_path = os.path.join(folder, "globalsums.xlsx")

# Defining the clusters and the specific conditions within them
scenarios = ["100", "120"]
conditions = [
    {"suffix": "", "label": "Pre-Fire (Baseline)", "color": "#1f77b4"},
    {"suffix": "", "label": "Post-Fire (CN79)", "color": "#ff7f0e"},
    {"suffix": "41", "label": "Post-Fire (CN 41)", "color": "#d62728"},
    {"suffix": "Imp50", "label": "Post-Fire (Imp 50%)", "color": "#9467bd"}
]

data_matrix = [] # Will store sums for each bar

print("Aggregating volumes for sensitivity analysis...")

# --- 2. DATA EXTRACTION ---
for sc in scenarios:
    sc_results = []
    for cond in conditions:
        # Determine the sheet name
        # Pre100, Post100, Post10041, Post100Imp50, etc.
        prefix = "Pre" if "Pre-Fire" in cond["label"] else "Post"
        sheet_name = f"{prefix}{sc}{cond['suffix']}"
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
            # Sum column index 4 (Volume mm)
            total_vol = pd.to_numeric(df.iloc[:, 4], errors='coerce').sum()
            sc_results.append(total_vol)
        except Exception as e:
            print(f"Warning: Could not read sheet {sheet_name}. Error: {e}")
            sc_results.append(0)
    data_matrix.append(sc_results)

# --- 3. PLOTTING ---
print("Generating comparison chart...")
plt.style.use('ggplot')
fig, ax = plt.subplots(figsize=(14, 8), constrained_layout=True)

# Transpose matrix for plotting groups
# row 0 = 100% results, row 1 = 120% results
x = np.arange(len(scenarios))
width = 0.2 # Thinner bars to fit 4 in a group

# Create the bars for each condition
for i, cond in enumerate(conditions):
    # Extract the i-th result from every scenario list
    bar_heights = [sc_data[i] for sc_data in data_matrix]
    # Offset the bars: -1.5, -0.5, 0.5, 1.5 times the width
    offset = (i - 1.5) * width
    
    bars = ax.bar(x + offset, bar_heights, width, label=cond["label"], color=cond["color"])
    
    # Add text labels (Strictly normal font)
    for bar in bars:
        yval = bar.get_height()
        if yval > 0:
            ax.text(bar.get_x() + bar.get_width()/2, yval + (max(max(data_matrix)) * 0.01),
                    f'{yval:,.0f}', ha='center', va='bottom', 
                    fontsize=9, fontweight='normal')

# Formatting
ax.set_ylabel("Total Cumulative Volume (mm)", fontsize=12, fontweight='normal')


ax.set_xticks(x)
ax.set_xticklabels([f"{sc}% Climate Scenario" for sc in scenarios], fontsize=11)

# Set Y-limit with room for labels
ax.set_ylim(0, max(max(data_matrix)) * 1.2)

ax.legend(loc='upper left', fontsize=10, framealpha=0.9)
ax.grid(axis='y', linestyle='--', alpha=0.7)

# Save the final image
output_path = os.path.join(folder, "Sensitivity_Volume_Analysis.png")
plt.savefig(output_path, dpi=300, bbox_inches='tight')
plt.show()

print(f"Success! Sensitivity chart saved to: {output_path}")
