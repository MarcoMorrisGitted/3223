import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os

# --- 1. SETUP ---
folder = r"C:\Users\ccexe\Downloads"
file_path = os.path.join(folder, "globalsums.xlsx")

scenarios = ["100", "110", "115", "120"]

pre_totals = []
post_totals = []

print("Calculating total watershed volumes...")

# --- 2. DATA EXTRACTION & AGGREGATION ---
for sc in scenarios:
    try:
        # Read Pre-Fire sheet
        df_pre = pd.read_excel(file_path, sheet_name=f"Pre{sc}", engine='openpyxl')
        # We assume you only want Subbasins. If Subbasins have specific names (like 'W'), 
        # you could filter here. For now, we sum the entire Volume column (Index 4).
        pre_sum = pd.to_numeric(df_pre.iloc[:, 4], errors='coerce').sum()
        pre_totals.append(pre_sum)
        
        # Read Post-Fire sheet
        df_post = pd.read_excel(file_path, sheet_name=f"Post{sc}", engine='openpyxl')
        post_sum = pd.to_numeric(df_post.iloc[:, 4], errors='coerce').sum()
        post_totals.append(post_sum)
        
    except Exception as e:
        print(f"Error on scenario {sc}: {e}. (Is the Excel file open?)")
        pre_totals.append(0)
        post_totals.append(0)

# --- 3. PLOTTING ---
print("Generating the Total Volume chart...")
plt.style.use('ggplot')
fig, ax = plt.subplots(figsize=(12, 8), constrained_layout=True)

x = np.arange(len(scenarios))
width = 0.35

# Plotting the grouped bars
bars1 = ax.bar(x - width/2, pre_totals, width, label='Pre-Fire (Baseline)', color='#1f77b4')
bars2 = ax.bar(x + width/2, post_totals, width, label='Post-Fire (Burn)', color='#d62728')

# --- CLEAN ANNOTATIONS (STRICTLY NORMAL FONT) ---
# Add text labels on top of the bars
for bar in bars1:
    yval = bar.get_height()
    if yval > 0:
        ax.text(bar.get_x() + bar.get_width()/2, yval + (max(post_totals) * 0.02), 
                f'{yval:,.0f}', ha='center', va='bottom', 
                fontsize=10, fontweight='normal', color='#333333')

for bar in bars2:
    yval = bar.get_height()
    if yval > 0:
        ax.text(bar.get_x() + bar.get_width()/2, yval + (max(post_totals) * 0.02), 
                f'{yval:,.0f}', ha='center', va='bottom', 
                fontsize=10, fontweight='normal', color='#660000')

# Formatting the Chart

ax.set_ylabel("Total Volume (mm)", fontsize=12, fontweight='normal')
ax.set_xlabel("Climate Scenario Multiplier", fontsize=12, fontweight='normal', labelpad=10)

# Label the X-Axis with the scenario percentages
ax.set_xticks(x)
ax.set_xticklabels([f"{sc}%" for sc in scenarios], fontsize=11)

# Dynamic Y-Axis scale to leave room for the text labels
ax.set_ylim(0, max(post_totals) * 1.15)

ax.legend(loc='upper left', fontsize=11, framealpha=0.9)
ax.grid(axis='y', linestyle='--', alpha=0.7)

# Save the final image
output_path = os.path.join(folder, "Total_Watershed_Volume.png")
plt.savefig(output_path, dpi=300, bbox_inches='tight')
plt.show()

print(f"Success! Executive summary chart saved to: {output_path}")