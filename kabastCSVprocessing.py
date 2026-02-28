import pandas as pd
import json
import seaborn as sns
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import os

# --- STYLING ---
# Star Wars inspired colors (from SWU aspects and general themes)
SW_COLORS = {
    "Vigilance": "#2c5da0",   # Blue
    "Command": "#3d7d44",     # Green
    "Aggression": "#b02a2a",  # Red
    "Cunning": "#cc9900",     # Yellow/Gold
    "Heroism": "#e0e0e0",     # White/Light Gray
    "Villainy": "#1a1a1a",    # Black/Dark Gray
    "Neutral": "#808080",     # Gray
    "Highlight": "#ff00ff",   # Magenta for highlights
    "Background": "#1a1a1a",  # Deep Space Black
    "Text": "#FFD700",        # Gold
    "Grid": "#404040"         # Dark Gray
}

# HATCH PATTERNS for aspects
ASPECT_HATCH = {
    "Vigilance": "//",
    "Command": "\\\\",
    "Aggression": "xx",
    "Cunning": "..",
    "Heroism": "++",
    "Villainy": "--",
    "Neutral": ""
}

# Set aesthetic parameters
plt.rcParams['figure.facecolor'] = SW_COLORS["Background"]
plt.rcParams['axes.facecolor'] = SW_COLORS["Background"]
plt.rcParams['axes.edgecolor'] = SW_COLORS["Text"]
plt.rcParams['axes.labelcolor'] = SW_COLORS["Text"]
plt.rcParams['axes.titlecolor'] = SW_COLORS["Text"]
plt.rcParams['xtick.color'] = SW_COLORS["Text"]
plt.rcParams['ytick.color'] = SW_COLORS["Text"]
plt.rcParams['grid.color'] = SW_COLORS["Grid"]
plt.rcParams['axes.titlesize'] = 18
plt.rcParams['axes.labelsize'] = 14
plt.rcParams['xtick.labelsize'] = 12
plt.rcParams['ytick.labelsize'] = 12
plt.rcParams['legend.fontsize'] = 12
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['legend.facecolor'] = SW_COLORS["Background"]
plt.rcParams['legend.edgecolor'] = SW_COLORS["Text"]
plt.rcParams['legend.title_fontsize'] = 13
plt.rcParams['text.color'] = SW_COLORS["Text"]

def get_aspect_color(aspects):
    if not aspects or not isinstance(aspects, list):
        return SW_COLORS["Neutral"]
    
    # Priority for coloring: Aggression, Command, Cunning, Vigilance
    # Then alignment Heroism/Villainy
    primaries = [a.strip() for a in aspects if a.strip() in ["Aggression", "Command", "Cunning", "Vigilance"]]
    alignments = [a.strip() for a in aspects if a.strip() in ["Heroism", "Villainy"]]
    
    if primaries:
        base_color = SW_COLORS.get(primaries[0], SW_COLORS["Neutral"])
        # If there's an alignment, we could potentially adjust the shade, 
        # but for now let's just use the primary.
        # However, to be "nuanced", if it's Villainy, we might want it darker?
        # Actually, let's keep it simple and recognizable by primary.
        return base_color
    
    if alignments:
        return SW_COLORS.get(alignments[0], SW_COLORS["Neutral"])
        
    return SW_COLORS["Neutral"]

# Set up GUI for file selection
root = tk.Tk()
root.withdraw()
root.configure(bg=SW_COLORS["Background"])
root.attributes("-topmost", True)

# --- GUI TOOLS ---
def multi_select_listbox(title, prompt, options, initial_selections=None):
    if initial_selections is None:
        initial_selections = []
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("400x500")
    win.configure(bg=SW_COLORS["Background"])
    win.attributes("-topmost", True)
    
    # We'll use a local loop to handle the window instead of just wait_window
    # to be more certain about event handling.
    # Also we'll deiconify root briefly if it's withdrawn to help some OS.
    # Actually let's try root.deiconify() for the duration of the dialog.
    # But user wanted it hidden. Let's try root.update_idletasks() first.
    root.update_idletasks()
    
    label = tk.Label(win, text=prompt, pady=10, bg=SW_COLORS["Background"], fg=SW_COLORS["Text"], font=("sans-serif", 10, "bold"))
    label.pack()
    
    frame = tk.Frame(win, bg=SW_COLORS["Background"])
    frame.pack(expand=True, fill='both', padx=20, pady=10)
    
    scrollbar = tk.Scrollbar(frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set, 
                         bg=SW_COLORS["Grid"], fg=SW_COLORS["Text"], 
                         selectbackground=SW_COLORS["Cunning"], selectforeground=SW_COLORS["Background"])
    sorted_options = sorted(options)
    for i, opt in enumerate(sorted_options):
        listbox.insert(tk.END, opt)
        if opt in initial_selections:
            listbox.selection_set(i)
    listbox.pack(expand=True, fill='both', side=tk.LEFT)
    scrollbar.config(command=listbox.yview)
    
    selection_state = {"result": [], "done": False}
    
    def on_confirm():
        indices = listbox.curselection()
        selection_state["result"] = [listbox.get(i) for i in indices]
        print(f"DEBUG: Confirmed {len(selection_state['result'])} leaders.")
        selection_state["done"] = True
        win.destroy()

    def on_cancel():
        print("DEBUG: Selection dialog cancelled or closed.")
        selection_state["done"] = True
        win.destroy()
        
    win.protocol("WM_DELETE_WINDOW", on_cancel)
        
    btn = tk.Button(win, text="Confirm Selection", command=on_confirm, pady=5,
                    bg=SW_COLORS["Cunning"], fg=SW_COLORS["Background"], font=("sans-serif", 10, "bold"))
    btn.pack(pady=10)
    
    win.transient(root)
    win.grab_set()
    win.focus_force() 
    
    # Use a local loop to wait for the dialog to close, which is often more reliable
    # than just wait_window in scripts that aren't primarily GUI-driven.
    print(f"DEBUG: Entering dialog event loop for '{title}'...")
    while not selection_state["done"]:
        try:
            root.update()
        except tk.TclError:
            # Window was destroyed
            break
        import time
        time.sleep(0.01) # Small sleep to avoid high CPU
    print("DEBUG: Exited dialog event loop.")
    
    return selection_state["result"]

# --- DATA PROCESSING ---
selected_files = filedialog.askopenfilenames(
    title="Select one or more CSV files you want to process and aggregate",
    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
    initialdir="."
)
root.attributes("-topmost", False)

if not selected_files:
    print("No files selected. Exiting.")
    exit()

# Load and aggregate match results
dataframes = []
for file_path in selected_files:
    try:
        temp_df = pd.read_csv(file_path)
        dataframes.append(temp_df)
    except Exception as e:
        print(f"Error loading {file_path}: {e}")

if not dataframes:
    print("No valid data found in selected files. Exiting.")
    exit()

df = pd.concat(dataframes, ignore_index=True)

# Determine output directory name based on selected files
if len(selected_files) == 1:
    filename_stem = Path(selected_files[0]).stem
else:
    filename_stem = f"Aggregated_{len(selected_files)}_files_{Path(selected_files[0]).stem}"

# Load card metadata
# The user's code tried cards.json then all_cards.json. 
# We'll check all_cards.json first as it's known to exist.
cards_file = "all_cards.json" if os.path.exists("all_cards.json") else "cards.json"
if not os.path.exists(cards_file):
    print(f"Error: Card metadata file {cards_file} not found.")
    exit()

with open(cards_file, "r", encoding="utf-8") as f:
    cards = json.load(f)

leader_lookup = {}
base_lookup = {}

for card in cards:
    # Use standard SWU-DB field names: "Name", "Type", "Aspects"
    name = card.get("Name")
    subtitle = card.get("Subtitle")
    card_type = card.get("Type")
    if card_type == "Leader":
        # Aspects are usually a list, e.g., ["Vigilance", "Villainy"]
        aspects = card.get("Aspects", [])
        entry = {"aspects": sorted(aspects), "subtitle": subtitle}
        # Store by Name and also by "Name | Subtitle" to match CSV
        leader_lookup[name] = entry
        if subtitle:
            leader_lookup[f"{name} | {subtitle}"] = entry
    elif card_type == "Base":
        # Bases have a single list of aspects, usually just one
        aspects = card.get("Aspects", [])
        base_lookup[name] = {
            "aspect": aspects[0] if aspects else None,
            "hp": card.get("HP", None)
        }

def normalize_leader(name):
    # Some exports might use pipes, like "Darth Vader | Unstoppable"
    return name.strip()

base_map = {
    "30hp-cunning-base": "Cunning Base",
    "30hp-aggression-base": "Aggression Base",
    "30hp-vigilance-base": "Vigilance Base",
    "30hp-command-base": "Command Base",
    "28hp-cunning-force-base": "Force Base",
    "28hp-aggression-force-base": "Force Base",
    "28hp-vigilance-force-base": "Force Base",
    "28hp-command-force-base": "Force Base",
    # Common SWU-DB export names
    "Aggression Base": "Aggression Base",
    "Cunning Base": "Cunning Base",
    "Vigilance Base": "Vigilance Base",
    "Command Base": "Command Base",
}

def normalize_base(name):
    if not name or not isinstance(name, str):
        return name
    return base_map.get(name, name.strip())

df["LeaderNorm"] = df["OpponentLeader"].apply(normalize_leader)
df["BaseNorm"] = df["OpponentBase"].apply(normalize_base)

# Map aspects and HP from lookup tables
df["LeaderAspects"] = df["LeaderNorm"].apply(lambda x: leader_lookup.get(x, {}).get("aspects"))
df["LeaderSubtitle"] = df["LeaderNorm"].apply(lambda x: leader_lookup.get(x, {}).get("subtitle"))

def get_base_aspect(name):
    # Try official lookup
    aspect = base_lookup.get(name, {}).get("aspect")
    if aspect:
        return aspect
    # Heuristic based on name
    name_lower = name.lower()
    if "aggression" in name_lower: return "Aggression"
    if "cunning" in name_lower: return "Cunning"
    if "vigilance" in name_lower: return "Vigilance"
    if "command" in name_lower: return "Command"
    return None

df["BaseAspect"] = df["BaseNorm"].apply(get_base_aspect)
df["BaseHP"] = df["BaseNorm"].apply(lambda x: base_lookup.get(x, {}).get("hp"))

# Ensure Wins and Losses are numeric
df["Wins"] = pd.to_numeric(df["Wins"], errors='coerce').fillna(0)
df["Losses"] = pd.to_numeric(df["Losses"], errors='coerce').fillna(0)

# --- PLOTTING FUNCTIONS ---
def generate_plots(data, output_dir, prefix="", highlighted=None):
    if highlighted is None:
        highlighted = []
    os.makedirs(output_dir, exist_ok=True)
    
    # Helper for stats
    def get_stats_local(d, group_cols):
        stats = d.groupby(group_cols).agg(
            Wins=("Wins", "sum"),
            Losses=("Losses", "sum"),
            Entries=("Wins", "count")
        ).reset_index()
        stats["TotalGames"] = stats["Wins"] + stats["Losses"]
        # Use 100 as base for percentage
        stats["WinRate"] = (stats["Wins"] / stats["TotalGames"].apply(lambda x: max(x, 1))) * 100
        return stats

    # 1. Win Rate by Leader Aspect Combination
    aspect_stats = get_stats_local(data, "LeaderAspectsStr")
    aspect_stats = aspect_stats.sort_values("WinRate", ascending=False)
    plt.figure(figsize=(12, 7))
    aspect_stats["Color"] = aspect_stats["LeaderAspectsStr"].apply(lambda x: get_aspect_color(x.split(", ")))
    
    # We use barplot but will apply hatches manually for consistency
    bars = sns.barplot(data=aspect_stats, 
                x="LeaderAspectsStr", y="WinRate", hue="LeaderAspectsStr", palette=dict(zip(aspect_stats["LeaderAspectsStr"], aspect_stats["Color"])), legend=True)
    
    # Apply hatches to the first plot as well
    # bars.patches contains all the bars. We need to match them to the data.
    for i, p in enumerate(bars.patches):
        if i < len(aspect_stats):
            row = aspect_stats.iloc[i]
            aspects = row["LeaderAspectsStr"].split(", ")
            # Priority for hatch: Aggression, Command, Cunning, Vigilance, then alignment
            primaries = [a.strip() for a in aspects if a.strip() in ["Aggression", "Command", "Cunning", "Vigilance"]]
            alignments = [a.strip() for a in aspects if a.strip() in ["Heroism", "Villainy"]]
            hatch_aspect = primaries[0] if primaries else (alignments[0] if alignments else "Neutral")
            hatch = ASPECT_HATCH.get(hatch_aspect, "")
            p.set_hatch(hatch)
            p.set_edgecolor(SW_COLORS["Text"])
            p.set_linewidth(1)

    # Add text labels on top of bars
    for i, p in enumerate(bars.patches):
        height = p.get_height()
        bars.annotate(f'{height:.1f}%', 
                    (p.get_x() + p.get_width() / 2., height), 
                    ha='center', va='center', 
                    xytext=(0, 10), 
                    textcoords='offset points',
                    color=SW_COLORS["Text"],
                    fontweight='bold',
                    fontsize=10)

    plt.title(f"{prefix}Win Rate by Opponent Leader Aspects (%)")
    plt.xticks(rotation=45, ha='right')
    plt.ylabel("Win Rate (%)")
    plt.ylim(0, 105)
    plt.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title="Aspects")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader_aspect.png"), dpi=300)
    plt.close()

    # 1b. Win Rate by Leader grouped by Aspect Pairing
    # Show win rate by leader, with them instead grouped positionally by the aspect pairings. 
    # The bar colors should be diagonal slashes of the aspect colors
    deck_stats = get_stats_local(data, ["LeaderAspectsStr", "LeaderNorm"])
    # If it's a Meta report, we might want to prioritize those specific leaders
    # But generally we group by Aspect pairing, then by WinRate within that pairing.
    deck_stats = deck_stats.sort_values(["LeaderAspectsStr", "WinRate"], ascending=[True, False])
    
    plt.figure(figsize=(14, 10))
    # We'll use a standard bar plot but manually apply hatches
    unique_aspect_pairs = deck_stats["LeaderAspectsStr"].unique()
    
    # Calculate x positions with extra spacing between aspect pairs
    x_positions = []
    current_x = 0
    last_aspect = None
    
    for i, row in deck_stats.iterrows():
        if last_aspect is not None and row["LeaderAspectsStr"] != last_aspect:
            current_x += 1.5 # Extra gap between groups
        else:
            current_x += 1.0
        x_positions.append(current_x)
        last_aspect = row["LeaderAspectsStr"]

    for i, (idx, row) in enumerate(deck_stats.iterrows()):
        aspects = row["LeaderAspectsStr"].split(", ")
        color = get_aspect_color(aspects)
        # Find primary aspect for hatch
        primaries = [a.strip() for a in aspects if a.strip() in ["Aggression", "Command", "Cunning", "Vigilance"]]
        alignments = [a.strip() for a in aspects if a.strip() in ["Heroism", "Villainy"]]
        hatch_aspect = primaries[0] if primaries else (alignments[0] if alignments else "Neutral")
        hatch = ASPECT_HATCH.get(hatch_aspect, "")
        
        bar_x = x_positions[i]
        plt.bar(bar_x, row["WinRate"], color=color, hatch=hatch, edgecolor=SW_COLORS["Text"], linewidth=1)
        
        # Add vertical text labels inside bars (centered)
        text_y = row["WinRate"] / 2
        plt.text(bar_x, text_y, f'{row["WinRate"]:.1f}%', 
                 ha='center', va='center', color=SW_COLORS["Text"], fontweight='bold', 
                 fontsize=10, rotation=90)

    plt.xticks(x_positions, deck_stats["LeaderNorm"], rotation=45, ha='right')
    plt.title(f"{prefix}Win Rate by Leader (%) (Grouped by Aspect Pairings)")
    plt.ylabel("Win Rate (%)")
    plt.ylim(0, 105)
    plt.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    
    # Custom legend for aspects
    from matplotlib.patches import Patch
    legend_elements = [Patch(facecolor=get_aspect_color([a]), hatch=ASPECT_HATCH.get(a, ""), 
                             edgecolor=SW_COLORS["Text"], label=a) for a in ASPECT_HATCH if a != "Neutral" and a != "Background" and a != "Text" and a != "Grid"]
    plt.legend(handles=legend_elements, bbox_to_anchor=(1.05, 1), loc='upper left', title="Aspect Patterns")
    
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "win_rate_by_deck_aspect.png"), dpi=300)
    plt.close()

    # 2. Top Opponent Leaders (prioritize most games played)
    leader_stats_full = get_stats_local(data, "LeaderNorm")
    
    # Identify truncation and add 'Other' category
    if len(leader_stats_full) > 20:
        # Sort by TotalGames first to identify top 19 (save one spot for Other)
        # Actually user said truncate to top 20, usually that means 20 entries total including Other? 
        # Or 20 specific entries + Other? 
        # "When truncating data to top 20, include a data series that is an aggregat of all other data"
        # I'll keep the top 19 and add 1 'Other' = 20 total.
        leader_stats_sorted = leader_stats_full.sort_values(["TotalGames", "WinRate"], ascending=[False, False])
        leader_stats = leader_stats_sorted.head(19).copy()
        other_data = leader_stats_sorted.iloc[19:]
        
        other_row = pd.DataFrame({
            "LeaderNorm": ["Other"],
            "Wins": [other_data["Wins"].sum()],
            "Losses": [other_data["Losses"].sum()],
            "Entries": [other_data["Entries"].sum()],
            "TotalGames": [other_data["TotalGames"].sum()]
        })
        other_row["WinRate"] = (other_row["Wins"] / other_row["TotalGames"].apply(lambda x: max(x, 1))) * 100
        leader_stats = pd.concat([leader_stats, other_row], ignore_index=True)
    else:
        # Sort by TotalGames first, then WinRate
        leader_stats = leader_stats_full.sort_values(["TotalGames", "WinRate"], ascending=[False, False])
    
    # Enrich labels with aspects for nuance
    def get_enriched_label(l):
        if l == "Other":
            count = leader_stats[leader_stats["LeaderNorm"] == "Other"]["Entries"].iloc[0]
            total_games = leader_stats[leader_stats["LeaderNorm"] == "Other"]["TotalGames"].iloc[0]
            return f"Other ({int(count)} leaders)\n({int(total_games)} games)"
        
        asp = leader_lookup.get(l, {}).get("aspects", [])
        total_games = leader_stats_full[leader_stats_full["LeaderNorm"] == l]["TotalGames"].iloc[0]
        return f"{l} ({int(total_games)} games)\n({', '.join(asp)})" if asp else f"{l} ({int(total_games)} games)"

    leader_stats["EnrichedLabel"] = leader_stats["LeaderNorm"].apply(get_enriched_label)
    
    plt.figure(figsize=(14, 10))
    colors = [SW_COLORS["Highlight"] if l in highlighted else get_aspect_color(leader_lookup.get(l, {}).get("aspects")) for l in leader_stats["LeaderNorm"]]
    bars = sns.barplot(data=leader_stats, x="WinRate", y="EnrichedLabel", hue="EnrichedLabel", palette=dict(zip(leader_stats["EnrichedLabel"], colors)), legend=False)
    
    # Add text labels on the bars
    for i, p in enumerate(bars.patches):
        width = p.get_width()
        # Avoid redundant 0% labels if any
        label = f'{width:.1f}%'
        bars.annotate(label, 
                    (width + 1.5, p.get_y() + p.get_height() / 2.), 
                    ha='left', va='center', 
                    color=SW_COLORS["Text"],
                    fontweight='bold',
                    fontsize=10)

    plt.title(f"{prefix}Top 20 Opponent Leaders (by Games Played)")
    if highlighted:
        plt.suptitle(f"(Highlighted: {', '.join(highlighted)})", fontsize=10, y=0.92, color=SW_COLORS["Text"])
    plt.xlabel("Win Rate (%)")
    plt.xlim(0, 115) # Extended X-axis
    plt.axvline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    plt.grid(axis='x', linestyle=':', alpha=0.3)
    # plt.legend removed as redundant with Y-axis labels
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader.png"), dpi=300)
    plt.close()

    # 3. Win Rate Heatmap
    matchup = get_stats_local(data, ["LeaderAspectsStr", "BaseAspect"])
    # Filter out Unknowns to make it cleaner
    matchup = matchup[(matchup["LeaderAspectsStr"] != "Unknown") & (matchup["BaseAspect"].notna())]
    
    if not matchup.empty:
        pivot = matchup.pivot(index="LeaderAspectsStr", columns="BaseAspect", values="WinRate")
        plt.figure(figsize=(12, 10))
        sns.heatmap(pivot, annot=True, cmap="RdYlGn", vmin=0, vmax=100, center=50, fmt=".1f", linewidths=.5,
                    cbar_kws={'label': 'Win Rate (%)'})
        plt.title(f"{prefix}Win Rate Heatmap (%): Leader Aspects vs Base Aspect")
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, "aspect_matchup_heatmap.png"), dpi=300)
        plt.close()
    else:
        print(f"DEBUG: Skipping heatmap for {prefix} because data is empty.")

    # 4. Scatter Plot (Replacing Bubble Chart)
    plt.figure(figsize=(16, 12)) # Slightly larger figure
    # For scatter plot, we'll use the top 30 most played leaders to avoid overcrowding.
    # When truncating, add an 'Other' category
    if len(leader_stats_full) > 30:
        bubble_stats_sorted = leader_stats_full.sort_values("TotalGames", ascending=False)
        bubble_stats = bubble_stats_sorted.head(29).copy()
        other_data = bubble_stats_sorted.iloc[29:]
        
        other_row = pd.DataFrame({
            "LeaderNorm": ["Other"],
            "Wins": [other_data["Wins"].sum()],
            "Losses": [other_data["Losses"].sum()],
            "TotalGames": [other_data["TotalGames"].sum()]
        })
        other_row["WinRate"] = (other_row["Wins"] / other_row["TotalGames"].apply(lambda x: max(x, 1))) * 100
        bubble_stats = pd.concat([bubble_stats, other_row], ignore_index=True)
    else:
        bubble_stats = leader_stats_full.sort_values("TotalGames", ascending=False)
    
    # Ensure points are visible even at 0% win rate
    # Use a color map from Red to Green based on WinRate
    sc = plt.scatter(bubble_stats["TotalGames"], bubble_stats["WinRate"], 
                        c=bubble_stats["WinRate"], cmap="RdYlGn", alpha=0.9, 
                        edgecolor=SW_COLORS["Text"], s=150) # Larger points
    
    # Add a colorbar to show the win rate scale
    cbar = plt.colorbar(sc)
    cbar.set_label("Win Rate (%)", color=SW_COLORS["Text"])
    cbar.ax.yaxis.set_tick_params(color=SW_COLORS["Text"], labelcolor=SW_COLORS["Text"])
    
    plt.title(f"{prefix}Leader Popularity (Games Played) vs Win Rate (%)")
    plt.ylabel("Win Rate (%)")
    plt.xlabel("Total Games Played")
    
    for i, row in bubble_stats.iterrows():
        # Clean up labels to prevent them from being too lengthy
        label = row["LeaderNorm"]
        if len(label) > 25:
            label = label[:22] + "..."
            
        # Offset labels to ensure they don't overlap with the point
        # Use a smaller rotation and adjust xytext to keep labels within bounds
        plt.annotate(label, (row["TotalGames"], row["WinRate"]), 
                     xytext=(8, 8), textcoords='offset points', fontsize=10, fontweight='bold',
                     rotation=10, color=SW_COLORS["Text"])
    
    plt.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    plt.ylim(-10, 115) # Increased margins to prevent label clipping at 0% and 100%
    max_games = bubble_stats["TotalGames"].max()
    plt.xlim(-max_games * 0.05, max_games * 1.25) # Increased padding for labels on the right
    plt.grid(True, linestyle=':', alpha=0.3)
    # plt.legend removed as it's now redundant with colorbar and labels
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "popularity_vs_winrate.png"), dpi=300)
    plt.close()

# --- MAIN EXECUTION ---
# filename_stem is determined during data processing
full_plots_dir = os.path.join("plots", filename_stem, "full_report")
meta_plots_dir = os.path.join("plots", filename_stem, "meta_report")

# Meta leaders persistence
META_FILE = "meta_leaders.json"
DEFAULT_META_LEADERS = [
    "Darth Vader", "Luke Skywalker", "Boba Fett", "Sabine Wren", 
    "Iden Versio", "Han Solo", "Leia Organa", "Grand Admiral Thrawn"
]

def load_meta_leaders():
    if os.path.exists(META_FILE):
        try:
            with open(META_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading {META_FILE}: {e}")
    return DEFAULT_META_LEADERS

def save_meta_leaders(leaders):
    try:
        with open(META_FILE, "w", encoding="utf-8") as f:
            json.dump(leaders, f, indent=2)
    except Exception as e:
        print(f"Error saving {META_FILE}: {e}")

meta_list = load_meta_leaders()

# Finalize data
def combine_deck_aspects(row):
    leader_aspects = row["LeaderAspects"]
    base_aspect = row["BaseAspect"]
    
    if not isinstance(leader_aspects, list):
        return "Unknown"
    
    # Ensure all elements are strings to avoid TypeError in sorted()
    combined = [str(a) for a in leader_aspects if a is not None]
    if base_aspect and pd.notna(base_aspect):
        combined.append(str(base_aspect))
    
    return ", ".join(sorted(combined))

df["LeaderAspectsStr"] = df["LeaderAspects"].apply(lambda x: ", ".join(x) if isinstance(x, list) else "Unknown")
df["DeckAspects"] = df.apply(combine_deck_aspects, axis=1)

# 1. Process Full Dataset
print("Generating full report...")
generate_plots(df, full_plots_dir, prefix="Full Dataset: ")
print(f"Full report generated in {full_plots_dir}")

# 2. Meta Selection
available_leaders = df["LeaderNorm"].unique().tolist()
print(f"Found {len(available_leaders)} available leaders.")
print("Opening leader meta selection dialog...")
root.deiconify() # Briefly show root if it's hidden to ensure children work correctly
root.title("SWU Deck Stats")
root.geometry("1x1+0+0") # Make it tiny and out of the way
root.attributes("-topmost", True)
root.update()
    
# Check if there are actually leaders to select
if not available_leaders:
    print("Warning: No leaders found in data to select.")
    selected_leaders = []
else:
    selected_leaders = multi_select_listbox("Leader Meta Analysis", 
                                          "Select leaders to include in the meta analysis.\nOnly these leaders will be used for the 'Meta' plots.", 
                                          available_leaders,
                                          initial_selections=meta_list)
root.withdraw()

if selected_leaders:
    # Update the persistent meta list: 
    # add new selections, and remove only those currently available that were deselected
    new_meta_set = set(meta_list)
    available_set = set(available_leaders)
    
    # Selected ones should definitely be in meta
    for sl in selected_leaders:
        new_meta_set.add(sl)
    
    # Available but NOT selected should be removed from meta
    for al in available_leaders:
        if al not in selected_leaders and al in new_meta_set:
            new_meta_set.remove(al)
            
    save_meta_leaders(sorted(list(new_meta_set)))
    
    print(f"Generating meta report for {len(selected_leaders)} leaders: {selected_leaders}")
    meta_df = df[df["LeaderNorm"].isin(selected_leaders)]
    generate_plots(meta_df, meta_plots_dir, prefix="Meta: ", highlighted=selected_leaders)
    print(f"Meta report generated in {meta_plots_dir}")
    messagebox.showinfo("Success", f"Reports generated!\n\nFull report: {full_plots_dir}\nMeta report: {meta_plots_dir}")
else:
    print("No leaders selected for meta. Skipping meta analysis.")
    messagebox.showinfo("Success", f"Full report generated in:\n{full_plots_dir}")

print("Processing complete.")
root.destroy()
