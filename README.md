# SWU Data Analysis

A collection of tools for analyzing and visualizing match data from Star Wars Unlimited.

## Main Tool: `kabastCSVprocessing.py`

This script processes match data exported from Karabast (in CSV format) to generate comprehensive win rate reports and visualizations.

### Features

- **Multi-File Aggregation**: Select multiple CSV files to combine data across different sessions or decks.
- **Robust Leader/Base Recognition**: Automatically normalizes leader and base names, including handling of subtitles and accents.
- **Visual Reports**: Generates two types of reports in the `plots/` directory:
  - **Full Dataset Report**: Analysis of every match found in the selected CSVs.
  - **Meta Report**: A focused analysis on a user-selected subset of leaders, allowing for deep dives into specific competitive environments.
- **Interactive Selection**: Uses a GUI dialog for file selection and meta-analysis configuration.

### Generated Plots

For both the Full and Meta reports, the following visualizations are produced:
- `win_rate_by_leader.png`: Win rate for each leader.
- `win_rate_by_leader_aspect.png`: Win rate broken down by leader and their base aspect.
- `win_rate_by_deck_aspect.png`: Win rate by the combination of leader and base aspects.
- `aspect_matchup_heatmap.png`: Heatmap showing performance of different aspect combinations against each other.
- `popularity_vs_winrate.png`: A bubble chart showing the relationship between how often a deck is played and its success rate.

## Requirements

- Python 3.x
- `pandas`
- `matplotlib`
- `seaborn`
- `tkinter` (usually included with Python)

## Usage

1. Run the main script:
   ```bash
   python kabastCSVprocessing.py
   ```
2. A file selection dialog will appear. Select the CSV(s) you wish to analyze.
3. Once the full report is generated, a second dialog will appear asking you to select which leaders to include in the **Meta Analysis**.
4. Results will be saved in the `plots/` folder, organized by the filename(s) of the source data.

## Utility Scripts

- `xlsx_to_json.py`: Converts Excel-based card data (`all_sets.xlsx`) into the `all_cards.json` format used by the main processing script.
- `list_leaders.py`: A simple utility to list all leaders found in the source Excel file.

## Legal Disclaimer

This tool is in no way affiliated with Disney or Fantasy Flight Games. Star Wars characters, cards, logos, and art are property of Disney and/or Fantasy Flight Games.
