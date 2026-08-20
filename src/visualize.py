"""
visualize.py

Produces one chart showing material distribution across the building,
from output/passport.json. 
Grouped by Material Category.
"""

import json
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

INPUT_JSON = Path("output/passport.json")
OUTPUT_PNG = Path("output/visualization.png")

sns.set_theme(style="whitegrid")

def make_chart():
    if not INPUT_JSON.exists():
        print(f"Error: {INPUT_JSON} not found. Please run build_passport.py first.")
        return

    print("Reading data for visualization from JSON...")
    with open(INPUT_JSON, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    df = pd.DataFrame(data)
    
    if 'material_category' in df.columns:
        df['material_category'] = df['material_category'].replace('', 'Not Specified').fillna('Not Specified')
        
        cat_counts = df['material_category'].value_counts().reset_index()
        cat_counts.columns = ['Material Category', 'Count']
        
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
        
        for p in ax.patches:
            width = p.get_width()
            if width > 0:
                ax.annotate(f'{int(width)}',
                            (width + 0.15, p.get_y() + p.get_height() / 2.),
                            ha='left', va='center',
                            fontsize=10, color='black', fontweight='semibold')

        plt.tight_layout()
        OUTPUT_PNG.parent.mkdir(parents=True, exist_ok=True)
        plt.savefig(OUTPUT_PNG, dpi=300, bbox_inches='tight')
        plt.close()
        print(f"✅ Successfully wrote visualization chart to {OUTPUT_PNG}")
    else:
        print("Error: 'material_category' field not found in JSON.")

if __name__ == "__main__":
    make_chart()