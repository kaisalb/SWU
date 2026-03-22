# SWU Data Analysis

A collection of tools for analyzing and visualizing match data from Star Wars Unlimited.

## Main Tool (Full Version): `kabastCSVprocessing.py`

This script processes match data exported from Karabast (in CSV format) to generate comprehensive win rate reports and visualizations. It uses a GUI for user-friendly file selection and meta-analysis configuration.

### Features

- **Multi-File Aggregation**: Select multiple CSV files (e.g., from different sessions or decks) to combine data.
- **Robust Leader/Base Recognition**: Automatically normalizes names, handling subtitles and accents using data from `all_cards.json`.
- **Visual Reports**: Generates detailed visualizations in the `plots/` directory:
  - **Full Dataset Report**: Analysis of all matches in the selected CSVs.
  - **Meta Report**: A focused analysis on a user-selected subset of leaders, with persistent selection across runs (saved in `meta_leaders.json`).
- **Aspect-Based Analysis**: Provides deep dives into win rates by leader, base, and deck aspect combinations.
- **Custom Styling**: Visualizations use Star Wars inspired color schemes and include unique "hatching" patterns for aspect identification.

### Generated Plots

For both the Full and Meta reports, the following visualizations are produced:
- `win_rate_by_leader.png`: Win rate for each leader.
- `win_rate_by_leader_aspect.png`: Win rate broken down by leader and their base aspect.
- `win_rate_by_deck_aspect.png`: Win rate by the combination of leader and base aspects.
- `aspect_matchup_heatmap.png`: Heatmap showing performance of different aspect combinations against each other.
- `popularity_vs_winrate.png`: A bubble chart showing the relationship between how often a deck is played and its success rate.

## Light Version: `karabastCSVprocessing_light.py`

This is a portable version of the main tool designed to be fast and dependency-light. It provides the same core functionality and visualization quality as the full version but without the need for `pandas`.

### Key Differences
- **No Pandas Dependency**: Uses native Python `csv` and `json` modules for data processing, making it easier to run in environments where installing large libraries is difficult.
- **Improved Dialog Logic**: Includes more robust window management for the Meta Analysis selection dialog.
- **Consistent Visuals**: Generates identical plots to the full version, including aspect-based hatching and Star Wars-themed color schemes.

## Standalone Executable: `SWU_Deck_Stats.exe`

For users who do not have Python installed, a pre-compiled standalone executable is available in the `dist/` directory.

### Features
- **Zero Installation**: Run the tool directly without installing Python or any libraries.
- **All-in-One**: Bundles all necessary resources (like `all_cards_min.json` and `leader_subtitles_min.json`) into a single executable file.
- **Portability**: Perfect for sharing with other players or running on different machines.

### Usage (Executable)
1. Navigate to the `dist/` folder.
2. Run `SWU_Deck_Stats.exe`.
3. Follow the same steps as the Python scripts for file selection and meta report generation.

## Requirements

### Full Version (`kabastCSVprocessing.py`)
- Python 3.x
- `pandas`, `openpyxl` (for Excel processing)
- `matplotlib`, `seaborn`
- `tkinter` (usually included with Python)

### Light Version (`karabastCSVprocessing_light.py`)
- Python 3.x
- `matplotlib`
- `tkinter` (usually included with Python)
- *Note: `pandas` and `seaborn` are NOT required.*

## Usage (Scripts)

1. Run either the full or light script:
   ```bash
   python kabastCSVprocessing.py
   # OR
   python karabastCSVprocessing_light.py
   ```
2. A file selection dialog will appear. Select the CSV(s) you wish to analyze.
3. Once the full report is generated, a second dialog will appear asking you to select which leaders to include in the **Meta Analysis**.
4. Results will be saved in the `plots/` folder, organized by the filename(s) of the source data.

## Utility Scripts

- `xlsx_to_json.py`: Processes the `all_sets.xlsx` file (manually compiled from official data) to generate `all_cards.json`. It handles subtitle extraction, aspect identification, and card numbering for all sets including Spark of Rebellion (SOR) up to A Lawless Time (LAW).
- `list_leaders.py`: A simple utility to list all unique leaders found in `all_cards.json`.
- `debug_hatching.py`: A visualization script to test and verify the aspect-based hatching patterns used in the main reports.

## Legal Disclaimer

This tool is in no way affiliated with Disney or Fantasy Flight Games. Star Wars characters, cards, logos, and art are property of Disney and/or Fantasy Flight Games.
