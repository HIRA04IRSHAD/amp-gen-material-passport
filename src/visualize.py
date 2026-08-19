"""
visualize.py

Produces one chart showing material distribution across the building,
from output/passport.json. Choice of grouping: Discipline, Material
Category, or Floor / Section (TODO: pick one).
"""

import json
from pathlib import Path

import matplotlib.pyplot as plt
import seaborn as sns

INPUT_JSON = Path("output/passport.json")
OUTPUT_PNG = Path("output/visualization.png")

sns.set_theme(style="whitegrid")


def load_records():
    with open(INPUT_JSON) as f:
        return json.load(f)


def make_chart(records):
    # TODO: aggregate records by chosen grouping (e.g. Discipline) and
    # plot as a bar chart.
    raise NotImplementedError


if __name__ == "__main__":
    records = load_records()
    make_chart(records)
    print(f"Wrote {OUTPUT_PNG}")
