"""
visualize.py

Produces one chart showing material distribution across the building,
from output/passport_filled.xlsx or passport.json. 
Grouped by Material Category.
"""

import json
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

INPUT_EXCEL = Path("output/passport_filled.xlsx")
OUTPUT_PNG = Path("output/visualization.png")

# Set aesthetic theme for the plot
sns.set_theme(style="whitegrid")

def make_chart():
    if not INPUT_EXCEL.exists():
        print(f"Error: {INPUT_EXCEL} not found. Please run build_passport.py first.")
        return

    print("Reading data for visualization...")
    df = pd.read_excel(INPUT_EXCEL, sheet_name="Material Passport", header=2)
    
    # Filter out empty rows where Material Category might be missing
    if 'Material Category' in df.columns:
        df = df.dropna(subset=['Material Category'])
        
        # Group by Material Category and count frequencies
        cat_counts = df['Material Category'].value_counts().reset_index()
        cat_counts.columns = ['Material Category', 'Count']
        
        # Create the bar chart
        plt.figure(figsize=(10, 6))
        ax = sns.barplot(
            data=cat_counts, 
            x='Count', 
            y='Material Category', 
            palette='viridis', 
            hue='Material Category', 
            legend=False
        )
        
        plt.title('Material Distribution Across Building (by Category)', fontsize=14, fontweight='bold', pad=15)
        plt.xlabel('Number of BoQ Items', fontsize=12, fontweight='bold')
        plt.ylabel('Material Category', fontsize=12, fontweight='bold')
        
        # Add value labels on the bars for extra clarity
        for p in ax.patches:
            width = p.get_width()
            ax.annotate(f'{int(width)}',
                        (width + 0.2, p.get_y() + p.get_height() / 2.),
                        ha='left', va='center',
                        fontsize=10, color='black', fontweight='semibold')

        plt.tight_layout()
        
        # Save output
        OUTPUT_PNG.parent.mkdir(parents=True, exist_ok=True)
        plt.savefig(OUTPUT_PNG, dpi=300)
        plt.close()
        print(f"Successfully wrote visualization chart to {OUTPUT_PNG}")
    else:
        print("Error: 'Material Category' column not found in Excel.")

if __name__ == "__main__":
    make_chart()