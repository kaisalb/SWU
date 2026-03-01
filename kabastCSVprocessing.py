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
    "Vigilance": "blue",
    "Command": "green",
    "Aggression": "red",
    "Cunning": "yellow",
    "Heroism": "white",
    "Villainy": "black",
    "Neutral": "#808080",
    "Highlight": "#ff00ff",
    "Background": "#1a1a1a",
    "Text": "#FFD700",
    "Grid": "#404040"
}

# HATCH PATTERNS for aspects logic
# Heroism: Slashes down to the right (\)
# Villainy: Slashes down to the left (/)
# No Heroism/Villainy: Vertical lines (|)
ASPECT_HATCH = {
    "Heroism": "\\",
    "Villainy": "/",
    "Neutral": "|"
}

def get_hatch(aspects):
    if not aspects or not isinstance(aspects, list):
        return ""
    
    # Heroism and Villany should continue to drive the hatch direction, 
    # note that this is currently only sourced from the leader aspects.
    # Convert to strings and handle potential case issues
    aspect_list = [str(a).strip().title() for a in aspects]
    
    if "Heroism" in aspect_list:
        return ASPECT_HATCH["Heroism"]
    elif "Villainy" in aspect_list:
        return ASPECT_HATCH["Villainy"]
    else:
        return ASPECT_HATCH["Neutral"]

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
    aspect_list = [str(a).strip().title() for a in aspects]
    primaries = [a for a in aspect_list if a in ["Aggression", "Command", "Cunning", "Vigilance"]]
    alignments = [a for a in aspect_list if a in ["Heroism", "Villainy"]]
    
    if primaries:
        base_color = SW_COLORS.get(primaries[0], SW_COLORS["Neutral"])
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
        if subtitle:
            leader_lookup[f"{name} | {subtitle}"] = entry
            
        # If multiple leaders have the same name, we might overwrite here
        # but the specific "Name | Subtitle" key will still be unique.
        # To handle cases where ONLY the name is provided, we should store a list of possibilities
        # or just the most recent one. 
        if name not in leader_lookup:
            leader_lookup[name] = entry
        else:
            # If it's already there, and we have a new one with a subtitle, 
            # we prefer the one that is likely more unique or we can just skip overwriting
            # to avoid ambiguity. For now, let's keep the first one found or handle it.
            pass
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
    "28hp-cunning-force-base": "Cunning Base",
    "28hp-aggression-force-base": "Aggression Base",
    "28hp-vigilance-force-base": "Vigilance Base",
    "28hp-command-force-base": "Command Base",
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
    # We must not group by 'LeaderAspects' because it contains unhashable lists.
    # Instead, we'll group by the string and map the aspects back for plotting.
    aspect_stats = get_stats_local(data, ["LeaderAspectsStr"])
    aspect_stats = aspect_stats.sort_values("WinRate", ascending=False)
    
    # Mapping back aspects for coloring
    aspect_map = data.drop_duplicates("LeaderAspectsStr").set_index("LeaderAspectsStr")["LeaderAspects"].to_dict()
    
    plt.figure(figsize=(12, 7))
    
    # Using plt.bar for more direct control over hatches and colors
    for i, (idx, row) in enumerate(aspect_stats.iterrows()):
        leader_aspects = aspect_map.get(row["LeaderAspectsStr"], [])
        color = get_aspect_color(leader_aspects)
        hatch = get_hatch_robust(row)
        
        bar_x = i
        plt.bar(bar_x, row["WinRate"], color=color, hatch=hatch, edgecolor=SW_COLORS["Text"], 
                linewidth=1.5, label=row["LeaderAspectsStr"])
        
        # Add text labels on top of bars
        plt.annotate(f'{row["WinRate"]:.1f}%', 
                    (bar_x, row["WinRate"]), 
                    ha='center', va='bottom', 
                    xytext=(0, 5), 
                    textcoords='offset points',
                    color=SW_COLORS["Text"],
                    fontweight='bold',
                    fontsize=10)

    plt.xticks(range(len(aspect_stats)), aspect_stats["LeaderAspectsStr"], rotation=45, ha='right')

    plt.title(f"{prefix}Win Rate by Opponent Leader Aspects (%)")
    plt.xticks(rotation=45, ha='right')
    plt.ylabel("Win Rate (%)")
    plt.ylim(0, 105)
    plt.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title="Aspects")
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader_aspect.png"), dpi=300)
    plt.close()

    # 1c. Win Rate by Deck Aspect Combination (Leader + Base)
    # Grouping by DeckAspects and LeaderNorm as requested
    # We remove LeaderAspects from groupby because it's a list (unhashable)
    deck_leader_stats = get_stats_local(data, ["DeckAspects", "LeaderNorm", "BaseAspect"])
    deck_leader_stats = deck_leader_stats.sort_values(["DeckAspects", "WinRate"], ascending=[True, True])
    
    # Mapping back leader aspects for coloring
    leader_aspect_map = data.drop_duplicates("LeaderNorm").set_index("LeaderNorm")["LeaderAspects"].to_dict()

    plt.figure(figsize=(16, 12))
    y_positions = []
    current_y = 0
    last_deck_aspect = None
    tick_positions = []
    tick_labels = []
    
    for i, row in deck_leader_stats.iterrows():
        if last_deck_aspect is not None and row["DeckAspects"] != last_deck_aspect:
            current_y += 1.5
        else:
            current_y += 1.0
        y_positions.append(current_y)
        last_deck_aspect = row["DeckAspects"]

    for i, (idx, row) in enumerate(deck_leader_stats.iterrows()):
        leader_norm = row["LeaderNorm"]
        leader_aspects = leader_aspect_map.get(leader_norm, [])
        base_aspect = row["BaseAspect"]
        
        # Logic for alternating colors:
        # 1. Leader aspect color(s)
        # 2. Base aspect color
        all_primaries = []
        if isinstance(leader_aspects, list):
            all_primaries.extend([a for a in leader_aspects if a in ["Aggression", "Command", "Cunning", "Vigilance"]])
        if base_aspect and base_aspect in ["Aggression", "Command", "Cunning", "Vigilance"]:
            if base_aspect not in all_primaries:
                all_primaries.append(base_aspect)
        
        # Calculate color based on index i and all available primary aspects
        # The user says "The colors should alternate between the leaders aspect color(s), and the base aspect color."
        # This implementation ensures that for each bar, we cycle through the aspects available for THAT deck.
        # Since each bar represents ONE deck, and we have many bars, 
        # using 'i % len(all_primaries)' makes the color depend on the bar's position and the number of aspects.
        if all_primaries:
            color = SW_COLORS.get(all_primaries[i % len(all_primaries)], SW_COLORS["Neutral"])
        else:
            color = SW_COLORS["Neutral"]

        hatch = get_hatch_robust(row)
        
        bar_y = y_positions[i]
        plt.barh(bar_y, row["WinRate"], color=color, hatch=hatch, edgecolor=SW_COLORS["Text"], linewidth=1.5)
        
        # Add horizontal text labels next to bars
        plt.text(row["WinRate"] + 1, bar_y, f'{row["WinRate"]:.1f}%', 
                 ha='left', va='center', color=SW_COLORS["Text"], fontweight='bold', 
                 fontsize=10)
        
        tick_positions.append(bar_y)
        # Label is Leader + Deck Aspects (shorthand if possible or just leader)
        tick_labels.append(f"{row['LeaderNorm']} ({row['DeckAspects']})")

    plt.title(f"{prefix}Win Rate by Leader & Deck Aspect Combination (%)")
    plt.yticks(tick_positions, tick_labels, fontsize=10)
    plt.xlabel("Win Rate (%)")
    plt.xlim(0, 115)
    plt.axvline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, "win_rate_by_deck_aspect.png"), dpi=300)
    plt.close()

    # 2. Top Opponent Leaders (prioritize most games played)
    leader_stats_full = get_stats_local(data, "LeaderNorm")
    
    # To handle split colors by base aspect, we need stats by (Leader, BaseAspect)
    leader_base_stats = get_stats_local(data, ["LeaderNorm", "BaseAspect"])
    
    # Identify truncation and add 'Other' category
    if len(leader_stats_full) > 20:
        leader_stats_sorted = leader_stats_full.sort_values(["TotalGames", "WinRate"], ascending=[False, False])
        leader_stats = leader_stats_sorted.head(19).copy()
        top_leaders = leader_stats["LeaderNorm"].tolist()
        
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
        leader_stats = leader_stats_full.sort_values(["TotalGames", "WinRate"], ascending=[False, False])
        top_leaders = leader_stats["LeaderNorm"].tolist()
    
    # Enrich labels
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
    for i, (idx, row) in enumerate(leader_stats.iterrows()):
        leader_name = row["LeaderNorm"]
        bar_y = i
        
        if leader_name == "Other":
            plt.barh(bar_y, row["WinRate"], color=SW_COLORS["Neutral"], edgecolor=SW_COLORS["Text"], linewidth=1.5)
        else:
            # Split bar logic
            l_aspects = leader_lookup.get(leader_name, {}).get("aspects", [])
            hatch = get_hatch_robust(row)
            
            # Get base aspect breakdown for this leader
            bases = leader_base_stats[leader_base_stats["LeaderNorm"] == leader_name].copy()
            total_l_games = bases["TotalGames"].sum()
            
            if total_l_games > 0:
                left = 0
                # Sort bases to have consistent coloring order if possible
                bases = bases.sort_values("BaseAspect")
                for _, b_row in bases.iterrows():
                    b_aspect = b_row["BaseAspect"]
                    b_games = b_row["TotalGames"]
                    # Percentage of this leader's games that were with this base aspect
                    share = b_games / total_l_games
                    # Width of this segment in the bar (share of the win rate)
                    width = row["WinRate"] * share
                    
                    # Color selection: use BaseAspect color if valid, otherwise fallback to highlight/leader color
                    if leader_name in highlighted:
                        color = SW_COLORS["Highlight"]
                    else:
                        color = SW_COLORS.get(b_aspect, get_aspect_color(l_aspects))
                    
                    plt.barh(bar_y, width, left=left, color=color, hatch=hatch, 
                             edgecolor=SW_COLORS["Text"], linewidth=1.5)
                    left += width
            else:
                # Fallback if no game data (shouldn't happen)
                color = SW_COLORS["Highlight"] if leader_name in highlighted else get_aspect_color(l_aspects)
                hatch = get_hatch_robust(row)
                plt.barh(bar_y, row["WinRate"], color=color, hatch=hatch, edgecolor=SW_COLORS["Text"], linewidth=1.5)
        
        # Win rate text annotation
        plt.annotate(f'{row["WinRate"]:.1f}%', 
                    (row["WinRate"] + 1.5, bar_y), 
                    ha='left', va='center', 
                    color=SW_COLORS["Text"],
                    fontweight='bold',
                    fontsize=10)

    plt.yticks(range(len(leader_stats)), leader_stats["EnrichedLabel"])
    plt.gca().invert_yaxis() # Top leaders first

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
    leader_norm = row["LeaderNorm"]
    
    if not isinstance(leader_aspects, list):
        # Fallback for leaders not in lookup
        # Attempt to find aspects if we have a name but the lookup failed
        # because of subtitle mismatch in CSV vs card DB
        if " | " in leader_norm:
            base_name = leader_norm.split(" | ")[0].strip()
            found_aspects = leader_lookup.get(base_name, {}).get("aspects")
            if found_aspects:
                leader_aspects = found_aspects
                # We can't easily update the row here, but we can return the correct combined aspects
    
    if not isinstance(leader_aspects, list):
        # Even if leader aspects are unknown, we might have base aspect
        if base_aspect and pd.notna(base_aspect):
            return str(base_aspect)
        return "Unknown"
    
    # Ensure all elements are strings to avoid TypeError in sorted()
    combined = [str(a) for a in leader_aspects if a is not None]
    if base_aspect and pd.notna(base_aspect):
        combined.append(str(base_aspect))
    
    # Remove duplicates if any (e.g. Heroism on both leader and base)
    combined = list(set(combined))
    
    return ", ".join(sorted(combined))

def get_hatch_robust(row):
    # Try to get from LeaderAspects first
    aspects = None
    if "LeaderAspects" in row:
        aspects = row["LeaderAspects"]
    elif "LeaderAspectsStr" in row and isinstance(row["LeaderAspectsStr"], str):
        # Fallback to splitting LeaderAspectsStr if it exists
        aspects = [a.strip() for a in row["LeaderAspectsStr"].split(",") if a.strip()]
        
    if isinstance(aspects, list) and aspects:
        return get_hatch(aspects)
    
    # Try finding aspects by base name if "Name | Subtitle" lookup failed
    # or if we are looking at a row from an aggregated stats table
    leader_norm = None
    if "LeaderNorm" in row:
        leader_norm = row["LeaderNorm"]
    
    if leader_norm and " | " in leader_norm:
        base_name = leader_norm.split(" | ")[0].strip()
        found_aspects = leader_lookup.get(base_name, {}).get("aspects")
        if found_aspects:
            return get_hatch(found_aspects)
            
    return ""

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
