import pandas as pd
import json
import seaborn as sns
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

import os

import unicodedata
from matplotlib.patches import Wedge
from matplotlib.offsetbox import AnnotationBbox, DrawingArea

# --- GLOBAL LOOKUPS ---
leader_lookup = {}
base_lookup = {}

def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                  if unicodedata.category(c) != 'Mn')

def get_leader_data(name):
    if not name:
        return {}
    
    name_lower = name.lower()
    # Try direct lookup
    data = leader_lookup.get(name_lower)
    if data:
        return data
    
    # Try name without accents
    name_stripped = strip_accents(name_lower)
    data = leader_lookup.get(name_stripped)
    if data:
        return data
    
    # Fallback for "Name | Subtitle" failure
    if " | " in name_lower:
        base_name = name_lower.split(" | ")[0].strip()
        data = leader_lookup.get(base_name)
        if data:
            return data
        
        # Try base name without accents
        base_name_stripped = strip_accents(base_name)
        data = leader_lookup.get(base_name_stripped)
        if data:
            return data
    
    # Very loose fallback for common typos/variations (e.g. "Vel Sarhta" -> "Vel Sartha")
    # Only if name is relatively short and we find a close match
    for k in leader_lookup:
        if len(k) > 5 and len(name_lower) > 5:
            # Check if one is a substring of the other or they are very similar
            if k in name_lower or name_lower in k:
                return leader_lookup[k]
            
    return {}

SET_ORDER = {
    "SOR": 0,
    "SHD": 1,
    "TWI": 2,
    "JTL": 3,
    "LOF": 4,
    "SEC": 5,
    "LAW": 6,
    "PRM": 7
}

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
    "Grid": "#404040",
    "WinRateHigh": "#90EE90",
    "WinRateLow": "#FFB6C1"
}

def format_aspects(aspect_list):
    """
    Sorts aspects so that primary colors come before Heroism/Villainy.
    Vigilance, Command, Aggression, Cunning, THEN Heroism, Villainy, Neutral.
    """
    if not aspect_list or not isinstance(aspect_list, list):
        return ""
    
    priority = {
        "Vigilance": 0, "Command": 1, "Aggression": 2, "Cunning": 3,
        "Heroism": 4, "Villainy": 5, "Neutral": 6
    }
    
    # Standardize names (Title Case)
    clean_list = list(set([str(a).strip().title() for a in aspect_list if a is not None]))
    
    # Sort by priority, then alphabetically for unknown ones
    sorted_list = sorted(clean_list, key=lambda x: (priority.get(x, 99), x))
    
    return ", ".join(sorted_list)

def get_leader_sort_info(name):
    """
    Returns sorting information for a leader based on set rank, alignment rank, and aspect rank.
    """
    data = get_leader_data(name)
    l_set = data.get("set", "Unknown")
    l_aspects = data.get("aspects", [])
    
    set_rank = SET_ORDER.get(l_set, 99)
    
    # Leader aspect (Heroism --> Villainy)
    alignment_priority = {"Heroism": 0, "Villainy": 1}
    alignments = [alignment_priority[a] for a in l_aspects if a in alignment_priority]
    alignment_rank = min(alignments) if alignments else 2
    
    # Leader Aspect color (Cunning --> Aggression --> Command --> Vigilance)
    aspect_priority = {"Cunning": 0, "Aggression": 1, "Command": 2, "Vigilance": 3}
    # Find the highest priority primary aspect
    primaries = [aspect_priority[a] for a in l_aspects if a in aspect_priority]
    aspect_rank = min(primaries) if primaries else 4
    
    return pd.Series([set_rank, alignment_rank, aspect_rank])

# Disclaimer for Star Wars property
DISCLAIMER_TEXT = ("Output is in no way affiliated with Disney or Fantasy Flight Games. "
                   "Star Wars characters, cards, logos, and art are property of Disney and/or Fantasy Flight Games.")

# Adjusted colors for text to ensure legibility on the Background (#1a1a1a)
# Black (Villainy) becomes light gray, White (Heroism) remains white, etc.
TEXT_ASPECT_COLORS = {
    "Vigilance": "#4169E1", # Royal Blue
    "Command": "#32CD32",   # Lime Green
    "Aggression": "#FF4500", # Orange Red
    "Cunning": "#FFD700",   # Gold
    "Heroism": "#FFFFFF",    # White
    "Villainy": "#D3D3D3",  # Light Gray
    "Neutral": "#A9A9A9"    # Dark Gray
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
    
    # Heroism and Villany drive the hatch direction. 
    # Currently sourced from the leader aspects.
    # Convert to strings and handle potential case issues
    aspect_list = [str(a).strip().title() for a in aspects]
    
    if "Heroism" in aspect_list:
        return ASPECT_HATCH["Heroism"]
    elif "Villainy" in aspect_list:
        return ASPECT_HATCH["Villainy"]
    else:
        return ASPECT_HATCH["Neutral"]

def get_alignment_color(aspects):
    if not aspects or not isinstance(aspects, list):
        return SW_COLORS["Neutral"]
    
    aspect_list = [str(a).strip().title() for a in aspects]
    if "Heroism" in aspect_list:
        return SW_COLORS["Heroism"]
    elif "Villainy" in aspect_list:
        return SW_COLORS["Villainy"]
    else:
        return SW_COLORS["Neutral"]

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


def get_aspect_color(aspects, use_text_colors=False):
    if not aspects or not isinstance(aspects, list):
        return TEXT_ASPECT_COLORS["Neutral"] if use_text_colors else SW_COLORS["Neutral"]
    
    # Priority for coloring: Vigilance, Command, Aggression, Cunning
    priority = ["Vigilance", "Command", "Aggression", "Cunning"]
    
    aspect_list = [str(a).strip().title() for a in aspects]
    primaries = [a for a in priority if a in aspect_list]
    alignments = [a for a in ["Heroism", "Villainy"] if a in aspect_list]
    
    palette = TEXT_ASPECT_COLORS if use_text_colors else SW_COLORS
    
    if primaries:
        # Use the highest-priority primary aspect for coloring.
        return palette.get(primaries[0], palette["Neutral"])
    
    if alignments:
        return palette.get(alignments[0], palette["Neutral"])
        
    return palette["Neutral"]

# Set up GUI for file selection
root = tk.Tk()
root.withdraw()
root.configure(bg=SW_COLORS["Background"])
root.attributes("-topmost", True)

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

# --- DATA PROCESSING HELPERS ---
def multi_select_listbox(title, prompt, options, initial_selections=None):
    if initial_selections is None:
        initial_selections = []
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("400x500")
    win.configure(bg=SW_COLORS["Background"])
    win.attributes("-topmost", True)
    
    # Use a local loop to handle the window instead of wait_window
    # to ensure event handling.
    # root.update_idletasks() ensures correct window state.
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

# --- PLOTTING FUNCTIONS ---
def generate_plots(data, output_dir, prefix="", filename_stem="", highlighted=None):
    if highlighted is None:
        highlighted = []
    os.makedirs(output_dir, exist_ok=True)
    
    # Format prefix for title: "Full Data" or "Meta Filter"
    title_filter = "Full Data" if "Full Dataset" in prefix else "Meta Filter"
    title_header = f"{filename_stem} ({title_filter})"
    
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
    # Instead, group by the string and map the aspects back for plotting.
    aspect_stats = get_stats_local(data, ["LeaderAspectsStr"])
    aspect_stats = aspect_stats.sort_values("WinRate", ascending=False)
    
    # Mapping back aspects for coloring
    aspect_map = data.drop_duplicates("LeaderAspectsStr").set_index("LeaderAspectsStr")["LeaderAspects"].to_dict()
    
    fig, ax = plt.subplots(figsize=(12, 7))
    
    # Using plt.bar for more direct control over hatches and colors
    for i, (idx, row) in enumerate(aspect_stats.iterrows()):
        leader_aspects = aspect_map.get(row["LeaderAspectsStr"], [])
        color = get_aspect_color(leader_aspects)
        hatch = get_hatch_robust(row)
        edge_color = get_hatch_color_robust(row)
        
        bar_x = i
        ax.bar(bar_x, row["WinRate"], color=color, hatch=hatch, edgecolor=edge_color, 
                linewidth=1.5, label=row["LeaderAspectsStr"])
        
        # Add text labels on top of bars
        ax.annotate(f'{row["WinRate"]:.1f}%', 
                    (bar_x, row["WinRate"]), 
                    ha='center', va='bottom', 
                    xytext=(0, 5), 
                    textcoords='offset points',
                    color=SW_COLORS["Text"],
                    fontweight='bold',
                    fontsize=10)

    ax.set_xticks(range(len(aspect_stats)))
    ax.set_xticklabels(aspect_stats["LeaderAspectsStr"], rotation=45, ha='right')

    # Color the aspect labels on the X-axis
    for label in ax.get_xticklabels():
        text = label.get_text()
        # The series label color should align to the leader's aspect color (including blending)
        # Parse the aspects from the string "Aspect1, Aspect2, ..."
        aspect_list = [a.strip() for a in text.split(',')]
        label.set_color(get_aspect_color(aspect_list, use_text_colors=True))

    ax.set_title(f"Win Rate by Opponent Leader Aspects (%)\n{title_header}")
    ax.set_ylabel("Win Rate (%)")
    ax.set_ylim(0, 105)
    ax.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axhline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axhline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    
    # Update legend to have colored text for aspects
    leg = ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title="Aspects")
    for text_obj in leg.get_texts():
        text = text_obj.get_text()
        aspect_list = [a.strip() for a in text.split(',')]
        text_obj.set_color(get_aspect_color(aspect_list, use_text_colors=True))
                
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    # Add Star Wars property disclaimer at the bottom
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader_aspect.png"), dpi=300)
    plt.close()

    # 1c. Win Rate by Deck Aspect Combination (Leader + Base)
    # Grouping by Leader and Base Aspect
    deck_leader_stats_full = get_stats_local(data, ["LeaderNorm", "BaseAspect"])
    
    # Custom sorting for BaseAspect
    aspect_order = {"Cunning": 0, "Aggression": 1, "Command": 2, "Vigilance": 3}
    deck_leader_stats_full["AspectOrder"] = deck_leader_stats_full["BaseAspect"].map(lambda x: aspect_order.get(x, 4))
    
    # Identify unique leaders in the order they should appear
    # We first determine the top 10 most played leaders overall
    leader_popularity = deck_leader_stats_full.groupby("LeaderNorm")["TotalGames"].sum().reset_index()
    leader_popularity = leader_popularity.sort_values("TotalGames", ascending=False)
    top_10_names = leader_popularity.head(10)["LeaderNorm"].tolist()
    
    # Now we create a sorting dataframe for these 10 names
    leader_order_df = pd.DataFrame({"LeaderNorm": top_10_names})
    
    # Enrich with sorting keys for leaders
    leader_order_df[["SetRank", "AlignmentRank", "AspectRank"]] = leader_order_df["LeaderNorm"].apply(get_leader_sort_info)
    
    # Final sorting of the top 10: 1. Set Number 2. Alignment 3. Leader Aspect color 4. Alphabetical
    leader_order_df = leader_order_df.sort_values(["SetRank", "AlignmentRank", "AspectRank", "LeaderNorm"])
    
    top_10_leaders = leader_order_df["LeaderNorm"].tolist()
    
    # Create a ranking for leaders based on this order
    leader_rank = {name: rank for rank, name in enumerate(top_10_leaders)}
    
    # Filter the stats to only these top 10 leaders
    deck_leader_stats = deck_leader_stats_full[deck_leader_stats_full["LeaderNorm"].isin(top_10_leaders)].copy()
    
    # Apply leader rank for sorting
    deck_leader_stats["LeaderRank"] = deck_leader_stats["LeaderNorm"].map(leader_rank)
    
    # Group results by leader, then consistent order of Vigilance, Command, Aggression, Cunning
    deck_leader_stats = deck_leader_stats.sort_values(["LeaderRank", "AspectOrder"])
    
    # Mapping back leader aspects for coloring
    leader_aspect_map = data.drop_duplicates("LeaderNorm").set_index("LeaderNorm")["LeaderAspects"].to_dict()

    fig, ax = plt.subplots(figsize=(16, 12))
    y_positions = []
    current_y = 0
    last_leader = None
    tick_positions = []
    tick_labels = []
    
    for i, row in deck_leader_stats.iterrows():
        if last_leader is not None and row["LeaderNorm"] != last_leader:
            current_y += 1.5
        else:
            current_y += 1.0
        y_positions.append(current_y)
        last_leader = row["LeaderNorm"]

    for i, (idx, row) in enumerate(deck_leader_stats.iterrows()):
        leader_norm = row["LeaderNorm"]
        leader_aspects = leader_aspect_map.get(leader_norm, [])
        base_aspect = row["BaseAspect"]
        
        # Color based on BaseAspect if available, otherwise fallback to leader color
        color = SW_COLORS.get(base_aspect, get_aspect_color(leader_aspects))

        hatch = get_hatch_robust(row)
        edge_color = get_hatch_color_robust(row)
        
        bar_y = y_positions[i]
        ax.barh(bar_y, row["WinRate"], color=color, hatch=hatch, edgecolor=edge_color, linewidth=1.5)
        
        # Add horizontal text labels next to bars - updated to show number of games
        ax.text(row["WinRate"] + 1, bar_y, f'{int(row["TotalGames"])} games', 
                 ha='left', va='center', color=SW_COLORS["Text"], fontweight='bold', 
                 fontsize=10)
        
        tick_positions.append(bar_y)
        # Label is Leader name, subtitle (Base Aspect)
        tick_labels.append(f"{row['LeaderNorm']} ({row['BaseAspect']})")

    ax.set_title(f"Top 10 Leaders by Games Played\n{title_header}")
    ax.set_yticks(tick_positions)
    ax.set_yticklabels(tick_labels, fontsize=10)
    
    # Color the aspect descriptions in the Y-axis labels
    for label in ax.get_yticklabels():
        text = label.get_text()
        # The series label should align to the leader's aspect color
        # Label is "Leader Name | Subtitle (Base Aspect)"
        if '(' in text and ')' in text:
            leader_name_part = text.split('(')[0].strip()
            base_aspect_part = text.split('(')[1].split(')')[0].strip()
            
            l_aspects = leader_aspect_map.get(leader_name_part, [])
            # Priority: Base Aspect > Leader primary aspects > alignment
            # But the requirement says "series label should always align to the leaders aspect color"
            # and "This should prioritize vigilance, command, aggression, cunning over heroism / vigilance" (likely meant heroism/villainy)
            # and "blend the colors" for multiple non-villain/hero aspects.
            label.set_color(get_aspect_color(l_aspects, use_text_colors=True))

    ax.set_xlabel("Win Rate (%)")
    ax.set_xlim(0, 115)
    ax.axvline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axvline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axvline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    # Add Star Wars property disclaimer at the bottom
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "win_rate_by_deck_aspect.png"), dpi=300)
    plt.close()

    # 2. Top Opponent Leaders (prioritize most games played)
    leader_stats_full = get_stats_local(data, "LeaderNorm")
    
    # To handle split colors by base aspect, we need stats by (Leader, BaseAspect)
    leader_base_stats = get_stats_local(data, ["LeaderNorm", "BaseAspect"])
    
    # Identify truncation and add 'Other' category
    if len(leader_stats_full) > 20:
        # First, find the top 20 most played leaders
        top_20_popularity = leader_stats_full.sort_values("TotalGames", ascending=False).head(20)
        top_20_names = top_20_popularity["LeaderNorm"].tolist()
        
        # Use the top 19 leaders to make room for 'Other'
        top_19_names = top_20_names[:19]
        
        # Sort these 19 using the established hierarchy
        top_19_df = leader_stats_full[leader_stats_full["LeaderNorm"].isin(top_19_names)].copy()
        top_19_df[["SetRank", "AlignmentRank", "AspectRank"]] = top_19_df["LeaderNorm"].apply(get_leader_sort_info)
        leader_stats_sorted = top_19_df.sort_values(["SetRank", "AlignmentRank", "AspectRank", "LeaderNorm"])
        
        leader_stats = leader_stats_sorted.copy()
        top_leaders = leader_stats["LeaderNorm"].tolist()
        
        # Calculate 'Other'
        other_data = leader_stats_full[~leader_stats_full["LeaderNorm"].isin(top_19_names)]
        other_row = pd.DataFrame({
            "LeaderNorm": ["Other"],
            "Wins": [other_data["Wins"].sum()],
            "Losses": [other_data["Losses"].sum()],
            "Entries": [other_data["Entries"].sum()],
            "TotalGames": [other_data["TotalGames"].sum()],
            "SetRank": [100],
            "AspectRank": [5]
        })
        other_row["WinRate"] = (other_row["Wins"] / other_row["TotalGames"].apply(lambda x: max(x, 1))) * 100
        leader_stats = pd.concat([leader_stats, other_row], ignore_index=True)
    else:
        # Sort all available leaders using the established hierarchy
        leader_stats_full[["SetRank", "AlignmentRank", "AspectRank"]] = leader_stats_full["LeaderNorm"].apply(get_leader_sort_info)
        leader_stats = leader_stats_full.sort_values(["SetRank", "AlignmentRank", "AspectRank", "LeaderNorm"])
        top_leaders = leader_stats["LeaderNorm"].tolist()
    
    # Enrich labels
    def get_enriched_label(l):
        if l == "Other":
            count = leader_stats[leader_stats["LeaderNorm"] == "Other"]["Entries"].iloc[0]
            total_games = leader_stats[leader_stats["LeaderNorm"] == "Other"]["TotalGames"].iloc[0]
            return f"Other ({int(count)} leaders)\n({int(total_games)} games)"
        
        asp = leader_lookup.get(l, {}).get("aspects", [])
        asp_str = format_aspects(asp)
        total_games = leader_stats_full[leader_stats_full["LeaderNorm"] == l]["TotalGames"].iloc[0]
        return f"{l} ({int(total_games)} games)\n({asp_str})" if asp else f"{l} ({int(total_games)} games)"

    leader_stats["EnrichedLabel"] = leader_stats["LeaderNorm"].apply(get_enriched_label)
    
    fig, ax = plt.subplots(figsize=(14, 10))
    for i, (idx, row) in enumerate(leader_stats.iterrows()):
        leader_name = row["LeaderNorm"]
        bar_y = i
        
        if leader_name == "Other":
            ax.barh(bar_y, row["WinRate"], color=SW_COLORS["Neutral"], edgecolor=SW_COLORS["Text"], linewidth=1.5)
        else:
            # Split bar logic
            l_aspects = leader_lookup.get(leader_name.lower(), {}).get("aspects", [])
            hatch = get_hatch_robust(row)
            
            # Get base aspect breakdown for this leader
            bases = leader_base_stats[leader_base_stats["LeaderNorm"] == leader_name].copy()
            total_l_games = bases["TotalGames"].sum()
            
            if total_l_games > 0:
                left = 0
                # Sort bases to have consistent coloring order if possible
                bases = bases.sort_values("BaseAspect")
                for i_base, (_, b_row) in enumerate(bases.iterrows()):
                    b_aspect = b_row["BaseAspect"]
                    b_games = b_row["TotalGames"]
                    # Percentage of this leader's games that were with this base aspect
                    share = b_games / total_l_games
                    # Width of this segment in the bar (share of the win rate)
                    width = row["WinRate"] * share
                    
                    # Color selection: use BaseAspect color if valid, otherwise fallback to leader color
                    color = SW_COLORS.get(b_aspect, get_aspect_color(l_aspects))
                    
                    # Hatch color based on alignment
                    edge_color = get_hatch_color_robust(row)
                    line_width = 1.5
                    
                    ax.barh(bar_y, width, left=left, color=color, hatch=hatch, 
                             edgecolor=edge_color, linewidth=line_width)
                    left += width
            else:
                # Fallback if no game data (shouldn't happen)
                color = get_aspect_color(l_aspects)
                hatch = get_hatch_robust(row)
                edge_color = get_hatch_color_robust(row)
                line_width = 1.5
                ax.barh(bar_y, row["WinRate"], color=color, hatch=hatch, edgecolor=edge_color, linewidth=line_width)
        
        # Win rate text annotation - updated to show number of games
        ax.annotate(f'{int(row["TotalGames"])} games', 
                    (row["WinRate"] + 1.5, bar_y), 
                    ha='left', va='center', 
                    color=SW_COLORS["Text"],
                    fontweight='bold',
                    fontsize=10)

    ax.set_yticks(range(len(leader_stats)))
    ax.set_yticklabels(leader_stats["EnrichedLabel"])
    
    # Color the labels in the Y-axis
    for label in ax.get_yticklabels():
        text = label.get_text()
        # EnrichedLabel has aspects in parentheses at the end, e.g., "(Vigilance, Villainy)"
        if '(' in text and ')' in text:
            # Extract the content of the last parentheses which contains the aspects
            parts = text.split('(')
            aspects_part = parts[-1].replace(')', '')
            aspect_list = [a.strip() for a in aspects_part.split(',')]
            
            # Use the leader name part for robust lookup to ensure we get the right aspects
            leader_name_raw = parts[0].strip()
            l_data = get_leader_data(leader_name_raw)
            l_aspects = l_data.get("aspects", aspect_list)
            
            # Use the leader's primary aspect color for the entire label
            label.set_color(get_aspect_color(l_aspects, use_text_colors=True))
                
    ax.invert_yaxis() # Top leaders first

    ax.set_title(f"Top 20 Opponent Leaders (by Games Played)\n{title_header}")
    if highlighted:
        # plt.suptitle(f"(Highlighted: {', '.join(highlighted)})", fontsize=10, y=0.92, color=SW_COLORS["Text"])
        pass
    ax.set_xlabel("Win Rate (%)")
    ax.set_xlim(0, 115) # Extended X-axis
    ax.axvline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axvline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axvline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    ax.grid(axis='x', linestyle=':', alpha=0.3)
    # plt.legend removed as redundant with Y-axis labels
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    # Add Star Wars property disclaimer at the bottom
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader.png"), dpi=300)
    plt.close()

    # 3. Game Popularity Heatmap
    # We want to show what percentage of total games were played in each leader / base aspect pairing
    total_games_overall = data["Wins"].sum() + data["Losses"].sum()
    
    matchup = data.copy()
    matchup["PairingGames"] = matchup["Wins"] + matchup["Losses"]
    matchup = matchup.groupby(["LeaderAspectsStr", "BaseAspect"])["PairingGames"].sum().reset_index()
    
    # Filter out Unknowns to make it cleaner
    matchup = matchup[(matchup["LeaderAspectsStr"] != "Unknown") & (matchup["BaseAspect"].notna())]
    
    if not matchup.empty and total_games_overall > 0:
        matchup["GamePercentage"] = (matchup["PairingGames"] / total_games_overall) * 100
        pivot = matchup.pivot(index="LeaderAspectsStr", columns="BaseAspect", values="GamePercentage")
        
        # --- ORDERING FOR HEATMAP ---
        # 1. Row (Leader Aspects) sorting
        def get_aspect_str_sort_info(aspect_str):
            # Parse aspects
            aspect_list = [a.strip() for a in aspect_str.split(',')]
            
            # Since these are aspect strings (potentially multiple leaders),
            # we don't have a single Set Number. We will use:
            # 1. Alignment (Heroism --> Villainy)
            # 2. Leader Aspect color (Vigilance --> Command --> Aggression --> Cunning)
            # 3. Alphabetical (the aspect string itself)
            
            alignment_priority = {"Heroism": 0, "Villainy": 1}
            alignments = [alignment_priority[a] for a in aspect_list if a in alignment_priority]
            alignment_rank = min(alignments) if alignments else 2
            
            aspect_priority = {"Vigilance": 0, "Command": 1, "Aggression": 2, "Cunning": 3}
            primaries = [aspect_priority[a] for a in aspect_list if a in aspect_priority]
            aspect_rank = min(primaries) if primaries else 4
            
            return (alignment_rank, aspect_rank, aspect_str)

        row_order = sorted(pivot.index.tolist(), key=get_aspect_str_sort_info)
        
        # 2. Column (Base Aspect) sorting
        aspect_order_heatmap = {"Vigilance": 0, "Command": 1, "Aggression": 2, "Cunning": 3}
        col_order = sorted(pivot.columns.tolist(), key=lambda x: aspect_order_heatmap.get(x, 4))
        
        pivot = pivot.reindex(index=row_order, columns=col_order)
        
        fig, ax = plt.subplots(figsize=(12, 10))
        # Use a single-color sequential colormap for popularity (e.g., "Blues" or "YlGnBu")
        sns.heatmap(pivot, annot=True, cmap="YlGnBu", vmin=0, fmt=".1f", linewidths=.5,
                    cbar_kws={'label': 'Percentage of Total Games (%)'}, ax=ax)
        
        # --- COLORING AXIS LABELS ---
        # Color the aspect labels on the Y-axis
        for label in ax.get_yticklabels():
            text = label.get_text()
            aspect_list = [a.strip() for a in text.split(',')]
            label.set_color(get_aspect_color(aspect_list, use_text_colors=True))
            
        # Color the base aspect labels on the X-axis
        for label in ax.get_xticklabels():
            text = label.get_text()
            label.set_color(TEXT_ASPECT_COLORS.get(text, TEXT_ASPECT_COLORS["Neutral"]))

        ax.set_title(f"Game Sample Heatmap (%)\n{title_header}")
        plt.tight_layout(rect=[0, 0.03, 1, 1])
        # Add Star Wars property disclaimer at the bottom
        fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
        plt.savefig(os.path.join(output_dir, "aspect_matchup_heatmap.png"), dpi=300)
        plt.close()
    else:
        print(f"DEBUG: Skipping heatmap for {prefix} because data is empty or total games is 0.")

    # 4. Scatter Plot (Replacing Bubble Chart)
    fig, ax = plt.subplots(figsize=(16, 12)) # Slightly larger figure
    # For scatter plot, use the top 30 most played leaders to avoid overcrowding.
    # When truncating, add an 'Other' category
    other_info_text = ""
    if len(leader_stats_full) > 30:
        bubble_stats_sorted = leader_stats_full.sort_values("TotalGames", ascending=False)
        bubble_stats = bubble_stats_sorted.head(29).copy()
        other_data = bubble_stats_sorted.iloc[29:]
        
        other_wins = other_data["Wins"].sum()
        other_losses = other_data["Losses"].sum()
        other_games = other_data["TotalGames"].sum()
        other_wr = (other_wins / max(other_games, 1)) * 100
        
        other_info_text = (f"Other Leaders Combined:\n"
                           f"Count: {len(other_data)}\n"
                           f"Games: {other_games}\n"
                           f"Win Rate: {other_wr:.1f}%")
    else:
        bubble_stats = leader_stats_full.sort_values("TotalGames", ascending=False)
    
    # Prepare base-aspect composition per leader for pie slices
    base_counts = data.groupby(["LeaderNorm", "BaseAspect"]).size().unstack(fill_value=0)

    def draw_pie(ax_local, center_xy, ratios_dict, size_px=28):
        # Ensure consistent order for base aspects
        order = ["Cunning", "Aggression", "Command", "Vigilance"]
        vals = [float(ratios_dict.get(k, 0.0)) for k in order]
        total = sum(vals) if sum(vals) > 0 else 1.0
        fracs = [v / total for v in vals]
        colors = [SW_COLORS.get(k, SW_COLORS["Neutral"]) for k in order]

        da = DrawingArea(size_px, size_px, 0, 0)
        theta1 = 0.0
        cx = cy = size_px / 2.0
        r = size_px / 2.0
        for frac, col in zip(fracs, colors):
            theta2 = theta1 + 360.0 * frac
            if frac > 0:
                w = Wedge((cx, cy), r, theta1, theta2, facecolor=col, edgecolor=SW_COLORS["Text"], linewidth=0.6)
                da.add_artist(w)
            theta1 = theta2
        ab = AnnotationBbox(da, center_xy, frameon=False, box_alignment=(0.5, 0.5))
        ax_local.add_artist(ab)

    # Draw a pie at each data point representing base aspect mix; remove redundant winrate colormap
    for _, row in bubble_stats.iterrows():
        leader = row["LeaderNorm"]
        x, y = row["TotalGames"], row["WinRate"]
        ratios = {}
        if leader in base_counts.index:
            counts = base_counts.loc[leader]
            total = counts.sum()
            if total > 0:
                ratios = {k: counts.get(k, 0) for k in counts.index}
        # Fallback to neutral if no data
        if not ratios:
            ratios = {"Neutral": 1.0}
        draw_pie(ax, (x, y), ratios, size_px=28)
    
    ax.set_title(f"Leader Popularity vs Win Rate (%)\n{title_header}")
    ax.set_ylabel("Win Rate (%)")
    ax.set_xlabel("Total Games Played")
    
    for i, row in bubble_stats.iterrows():
        # Clean up labels to prevent them from being too lengthy
        label = row["LeaderNorm"]
        if len(label) > 25:
            label = label[:22] + "..."
            
        # Get leader aspects for coloring the annotation
        leader_name = row["LeaderNorm"]
        l_aspects = get_leader_data(leader_name).get("aspects", [])
        # The series label should align to the leader's aspect color (including blending)
        text_color = get_aspect_color(l_aspects, use_text_colors=True)

        # Offset labels to ensure they don't overlap with the point
        # Increased xytext offset to avoid overlap with the pie chart
        ax.annotate(label, (row["TotalGames"], row["WinRate"]), 
                     xytext=(14, 14), textcoords='offset points', fontsize=10, fontweight='bold',
                     rotation=10, color=text_color)
    
    if other_info_text:
        # Add 'Other' info box to the right
        ax.text(1.02, 0.5, other_info_text, transform=ax.transAxes, 
                fontsize=11, fontweight='bold', color=SW_COLORS["Text"],
                bbox=dict(facecolor=SW_COLORS["Background"], edgecolor=SW_COLORS["Grid"], alpha=0.8),
                verticalalignment='center')

    ax.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axhline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axhline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    ax.set_ylim(-10, 115) # Increased margins to prevent label clipping at 0% and 100%
    max_games = bubble_stats["TotalGames"].max()
    ax.set_xlim(-max_games * 0.05, max_games * 1.15) # Adjusted padding for labels now that Other is gone
    ax.grid(True, linestyle=':', alpha=0.3)
    # plt.legend removed as pies convey base-aspect composition and labels identify leaders
    plt.tight_layout(rect=[0, 0.03, 0.9, 1]) # Make room on the right for the text box
    # Add Star Wars property disclaimer at the bottom
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "popularity_vs_winrate.png"), dpi=300)
    plt.close()

# --- PERSISTENCE AND DATA PROCESSING HELPERS ---
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

def combine_deck_aspects(row):
    leader_aspects = row["LeaderAspects"]
    base_aspect = row["BaseAspect"]
    leader_norm = row["LeaderNorm"]
    
    if not isinstance(leader_aspects, list):
        found_data = get_leader_data(leader_norm)
        if found_data:
            leader_aspects = found_data.get("aspects")
    
    if not isinstance(leader_aspects, list):
        if base_aspect and pd.notna(base_aspect):
            return str(base_aspect)
        return "Unknown"
    
    combined = [str(a) for a in leader_aspects if a is not None]
    if base_aspect and pd.notna(base_aspect):
        combined.append(str(base_aspect))
    
    return format_aspects(combined)

def get_hatch_robust(row):
    aspects = None
    if "LeaderAspects" in row:
        aspects = row["LeaderAspects"]
    elif "LeaderAspectsStr" in row and isinstance(row["LeaderAspectsStr"], str):
        aspects = [a.strip() for a in row["LeaderAspectsStr"].split(",") if a.strip()]
        
    if isinstance(aspects, list) and aspects:
        return get_hatch(aspects)
    
    leader_norm = None
    if "LeaderNorm" in row:
        leader_norm = row["LeaderNorm"]
    
    if leader_norm:
        found_data = get_leader_data(leader_norm)
        if found_data and "aspects" in found_data:
            return get_hatch(found_data["aspects"])
            
    return ""

def get_hatch_color_robust(row):
    aspects = None
    if "LeaderAspects" in row:
        aspects = row["LeaderAspects"]
    elif "LeaderAspectsStr" in row and isinstance(row["LeaderAspectsStr"], str):
        aspects = [a.strip() for a in row["LeaderAspectsStr"].split(",") if a.strip()]
        
    if isinstance(aspects, list) and aspects:
        return get_alignment_color(aspects)
    
    leader_norm = None
    if "LeaderNorm" in row:
        leader_norm = row["LeaderNorm"]
    
    if leader_norm:
        found_data = get_leader_data(leader_norm)
        if found_data and "aspects" in found_data:
            return get_alignment_color(found_data["aspects"])
            
    return SW_COLORS["Neutral"]

def load_card_data():
    cards_file = "all_cards.json" if os.path.exists("all_cards.json") else "cards.json"
    if not os.path.exists(cards_file):
        print(f"Error: Card metadata file {cards_file} not found.")
        return []

    with open(cards_file, "r", encoding="utf-8") as f:
        return json.load(f)

def build_lookups(cards):
    l_lookup = {}
    b_lookup = {}
    for card in cards:
        name = card.get("Name")
        subtitle = card.get("Subtitle")
        card_type = card.get("Type")
        if card_type == "Leader":
            aspects = card.get("Aspects", [])
            entry = {
                "aspects": sorted(aspects), 
                "subtitle": subtitle,
                "set": card.get("Set"),
                "number": card.get("Number")
            }
            if subtitle:
                l_lookup[f"{name} | {subtitle}".lower()] = entry
                l_lookup[strip_accents(f"{name} | {subtitle}".lower())] = entry
            if name.lower() not in l_lookup:
                l_lookup[name.lower()] = entry
            name_stripped = strip_accents(name.lower())
            if name_stripped not in l_lookup:
                l_lookup[name_stripped] = entry
        elif card_type == "Base":
            aspects = card.get("Aspects", [])
            b_lookup[name] = {
                "aspect": aspects[0] if aspects else None,
                "hp": card.get("HP", None)
            }
    return l_lookup, b_lookup

def process_match_data(selected_files, l_lookup, b_lookup):
    dataframes = []
    for file_path in selected_files:
        try:
            temp_df = pd.read_csv(file_path)
            dataframes.append(temp_df)
        except Exception as e:
            print(f"Error loading {file_path}: {e}")

    if not dataframes:
        return None

    df = pd.concat(dataframes, ignore_index=True)
    df["LeaderNorm"] = df["OpponentLeader"].apply(normalize_leader)
    df["BaseNorm"] = df["OpponentBase"].apply(normalize_base)

    df["LeaderAspects"] = df["LeaderNorm"].apply(lambda x: get_leader_data(x).get("aspects"))
    df["LeaderSubtitle"] = df["LeaderNorm"].apply(lambda x: get_leader_data(x).get("subtitle"))
    df["LeaderSet"] = df["LeaderNorm"].apply(lambda x: get_leader_data(x).get("set"))
    df["LeaderNumber"] = df["LeaderNorm"].apply(lambda x: get_leader_data(x).get("number"))

    def get_base_aspect_local(name):
        aspect = b_lookup.get(name, {}).get("aspect")
        if aspect: return aspect
        name_lower = name.lower()
        if "aggression" in name_lower: return "Aggression"
        if "cunning" in name_lower: return "Cunning"
        if "vigilance" in name_lower: return "Vigilance"
        if "command" in name_lower: return "Command"
        return None

    df["BaseAspect"] = df["BaseNorm"].apply(get_base_aspect_local)
    df["BaseHP"] = df["BaseNorm"].apply(lambda x: b_lookup.get(x, {}).get("hp"))
    df["Wins"] = pd.to_numeric(df["Wins"], errors='coerce').fillna(0)
    df["Losses"] = pd.to_numeric(df["Losses"], errors='coerce').fillna(0)
    
    df["LeaderAspectsStr"] = df["LeaderAspects"].apply(lambda x: format_aspects(x) if isinstance(x, list) else "Unknown")
    df["DeckAspects"] = df.apply(combine_deck_aspects, axis=1)
    
    return df

# --- MAIN EXECUTION ---
def main():
    global leader_lookup, base_lookup
    # --- DATA PROCESSING ---
    selected_files = filedialog.askopenfilenames(
        title="Select one or more CSV files you want to process and aggregate",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        initialdir="."
    )
    root.attributes("-topmost", False)

    if not selected_files:
        print("No files selected. Exiting.")
        return

    # Load card metadata
    cards = load_card_data()
    if not cards:
        return

    leader_lookup, base_lookup = build_lookups(cards)

    # Determine output directory name based on selected files
    if len(selected_files) == 1:
        filename_stem = Path(selected_files[0]).stem
    else:
        filename_stem = f"Aggregated_{len(selected_files)}_files_{Path(selected_files[0]).stem}"

    # Load and process match results
    df = process_match_data(selected_files, leader_lookup, base_lookup)
    if df is None:
        print("No valid data found in selected files. Exiting.")
        return

    full_plots_dir = os.path.join("plots", filename_stem, "full_report")
    meta_plots_dir = os.path.join("plots", filename_stem, "meta_report")

    meta_list = load_meta_leaders()

    # 1. Process Full Dataset
    print("Generating full report...")
    generate_plots(df, full_plots_dir, prefix="Full Dataset: ", filename_stem=filename_stem)
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
        # Update the persistent meta list
        new_meta_set = set(meta_list)
        available_set = set(available_leaders)
        for sl in selected_leaders:
            new_meta_set.add(sl)
        for al in available_leaders:
            if al not in selected_leaders and al in new_meta_set:
                new_meta_set.remove(al)
                
        save_meta_leaders(sorted(list(new_meta_set)))
        
        print(f"Generating meta report for {len(selected_leaders)} leaders: {selected_leaders}")
        meta_df = df[df["LeaderNorm"].isin(selected_leaders)]
        generate_plots(meta_df, meta_plots_dir, prefix="Meta: ", filename_stem=filename_stem, highlighted=selected_leaders)
        print(f"Meta report generated in {meta_plots_dir}")
        messagebox.showinfo("Success", f"Reports generated!\n\nFull report: {full_plots_dir}\nMeta report: {meta_plots_dir}")
    else:
        print("No leaders selected for meta. Skipping meta analysis.")
        messagebox.showinfo("Success", f"Full report generated in:\n{full_plots_dir}")

    print("Processing complete.")
    root.destroy()

if __name__ == "__main__":
    main()

# --- PLOTTING FUNCTIONS ---
