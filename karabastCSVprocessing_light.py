"""
SWU Deck Stats (Light Version)
A portable, dependency-light version of the match data analysis tool.
Uses native CSV and JSON processing (no pandas required).
"""
import sys
import csv
import json
import matplotlib
matplotlib.use('Agg')
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

def get_resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                  if unicodedata.category(c) != 'Mn')

def safe_float(v):
    if not v: return 0.0
    try: return float(str(v).strip())
    except: return 0.0

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

SET_ORDER = {"SOR": 0, "SHD": 1, "TWI": 2, "JTL": 3, "LOF": 4, "SEC": 5, "LAW": 6, "PRM": 7}

SW_COLORS = {
    "Vigilance": "blue", "Command": "green", "Aggression": "red", "Cunning": "yellow",
    "Heroism": "white", "Villainy": "black", "Neutral": "#808080", "Highlight": "#ff00ff",
    "Background": "#1a1a1a", "Text": "#FFD700", "Grid": "#404040",
    "WinRateHigh": "#90EE90", "WinRateLow": "#FFB6C1"
}

def format_aspects(aspect_list):
    if not aspect_list or not isinstance(aspect_list, list):
        return ""
    priority = {"Vigilance": 0, "Command": 1, "Aggression": 2, "Cunning": 3, "Heroism": 4, "Villainy": 5, "Neutral": 6}
    clean_list = list(set([str(a).strip().title() for a in aspect_list if a is not None]))
    sorted_list = sorted(clean_list, key=lambda x: (priority.get(x, 99), x))
    return ", ".join(sorted_list)

def get_leader_sort_info(name):
    data = get_leader_data(name)
    l_set = data.get("set", "Unknown")
    l_aspects = data.get("aspects", [])
    set_rank = SET_ORDER.get(l_set, 99)
    alignment_priority = {"Heroism": 0, "Villainy": 1}
    alignments = [alignment_priority[a] for a in l_aspects if a in alignment_priority]
    alignment_rank = min(alignments) if alignments else 2
    aspect_priority = {"Cunning": 0, "Aggression": 1, "Command": 2, "Vigilance": 3}
    primaries = [aspect_priority[a] for a in l_aspects if a in aspect_priority]
    aspect_rank = min(primaries) if primaries else 4
    return (set_rank, alignment_rank, aspect_rank)

DISCLAIMER_TEXT = ("Output is in no way affiliated with Disney or Fantasy Flight Games. "
                   "Star Wars characters, cards, logos, and art are property of Disney and/or Fantasy Flight Games.")

TEXT_ASPECT_COLORS = {
    "Vigilance": "#4169E1", "Command": "#32CD32", "Aggression": "#FF4500", "Cunning": "#FFD700",
    "Heroism": "#FFFFFF", "Villainy": "#D3D3D3", "Neutral": "#A9A9A9"
}

ASPECT_HATCH = {"Heroism": "\\", "Villainy": "/", "Neutral": "|"}

def get_hatch(aspects):
    if not aspects or not isinstance(aspects, list): return ""
    aspect_list = [str(a).strip().title() for a in aspects]
    if "Heroism" in aspect_list: return ASPECT_HATCH["Heroism"]
    elif "Villainy" in aspect_list: return ASPECT_HATCH["Villainy"]
    else: return ASPECT_HATCH["Neutral"]

def get_alignment_color(aspects):
    if not aspects or not isinstance(aspects, list): return SW_COLORS["Neutral"]
    aspect_list = [str(a).strip().title() for a in aspects]
    if "Heroism" in aspect_list: return SW_COLORS["Heroism"]
    elif "Villainy" in aspect_list: return SW_COLORS["Villainy"]
    else: return SW_COLORS["Neutral"]

# Set aesthetic parameters
plt.rcParams.update({
    'figure.facecolor': SW_COLORS["Background"], 'axes.facecolor': SW_COLORS["Background"],
    'axes.edgecolor': SW_COLORS["Text"], 'axes.labelcolor': SW_COLORS["Text"],
    'axes.titlecolor': SW_COLORS["Text"], 'xtick.color': SW_COLORS["Text"],
    'ytick.color': SW_COLORS["Text"], 'grid.color': SW_COLORS["Grid"],
    'axes.titlesize': 18, 'axes.labelsize': 14, 'xtick.labelsize': 12, 'ytick.labelsize': 12,
    'legend.fontsize': 12, 'font.family': 'sans-serif', 'legend.facecolor': SW_COLORS["Background"],
    'legend.edgecolor': SW_COLORS["Text"], 'legend.title_fontsize': 13, 'text.color': SW_COLORS["Text"]
})

def get_aspect_color(aspects, use_text_colors=False):
    if not aspects or not isinstance(aspects, list):
        return TEXT_ASPECT_COLORS["Neutral"] if use_text_colors else SW_COLORS["Neutral"]
    priority = ["Vigilance", "Command", "Aggression", "Cunning"]
    aspect_list = [str(a).strip().title() for a in aspects]
    primaries = [a for a in priority if a in aspect_list]
    alignments = [a for a in ["Heroism", "Villainy"] if a in aspect_list]
    palette = TEXT_ASPECT_COLORS if use_text_colors else SW_COLORS
    if primaries: return palette.get(primaries[0], palette["Neutral"])
    if alignments: return palette.get(alignments[0], palette["Neutral"])
    return palette["Neutral"]

def normalize_leader(name): return name.strip()

base_map = {
    "30hp-cunning-base": "Cunning Base", "30hp-aggression-base": "Aggression Base",
    "30hp-vigilance-base": "Vigilance Base", "30hp-command-base": "Command Base",
    "28hp-cunning-force-base": "Cunning Base", "28hp-aggression-force-base": "Aggression Base",
    "28hp-vigilance-force-base": "Vigilance Base", "28hp-command-force-base": "Command Base",
}

def normalize_base(name):
    if not name or not isinstance(name, str): return name
    return base_map.get(name, name.strip())

def multi_select_listbox(title, prompt, options, initial_selections=None):
    if initial_selections is None: initial_selections = []
    print(f"DEBUG: Opening dialog '{title}' with {len(options)} options...")
    
    # Ensure any pending root updates are handled
    root.update_idletasks()
    
    win = tk.Toplevel(root)
    win.withdraw()  # Hide while building and centering
    win.title(title)
    win.configure(bg=SW_COLORS["Background"])
    
    # Label
    tk.Label(win, text=prompt, pady=10, bg=SW_COLORS["Background"], fg=SW_COLORS["Text"], font=("sans-serif", 10, "bold")).pack()
    
    # Listbox with scrollbar
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
        print(f"DEBUG: Selection confirmed: {len(selection_state['result'])} items.")
        selection_state["done"] = True
        win.destroy()
        
    def on_cancel():
        print("DEBUG: Selection cancelled.")
        selection_state["done"] = True
        win.destroy()
        
    win.protocol("WM_DELETE_WINDOW", on_cancel)
    
    btn_frame = tk.Frame(win, bg=SW_COLORS["Background"])
    btn_frame.pack(pady=10)
    
    tk.Button(btn_frame, text="Confirm Selection", command=on_confirm, width=18, pady=5, 
              bg=SW_COLORS["Cunning"], fg=SW_COLORS["Background"], font=("sans-serif", 10, "bold")).pack(side=tk.LEFT, padx=5)
              
    tk.Button(btn_frame, text="Cancel", command=on_cancel, width=10, pady=5, 
              bg=SW_COLORS["Grid"], fg=SW_COLORS["Text"], font=("sans-serif", 10)).pack(side=tk.LEFT, padx=5)
    
    # Center the window
    win.update_idletasks()
    width, height = 400, 600
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')
    
    # Show and bring to front
    win.deiconify()
    win.attributes("-topmost", True)
    win.lift()
    win.focus_force()
    win.grab_set()
    
    # Local event loop to wait for the dialog to close
    print("DEBUG: Entering wait_window...")
    win.wait_window(win)
    
    print("DEBUG: Dialog loop finished.")
    return selection_state["result"]

def get_stats_local(data_rows, group_cols):
    if isinstance(group_cols, str): group_cols = [group_cols]
    aggregated = {}
    for row in data_rows:
        key = tuple(row.get(col) for col in group_cols)
        if key not in aggregated:
            aggregated[key] = {"Wins": 0.0, "Losses": 0.0, "Entries": 0}
            for i, col in enumerate(group_cols): aggregated[key][col] = key[i]
        aggregated[key]["Wins"] += row.get("Wins", 0.0)
        aggregated[key]["Losses"] += row.get("Losses", 0.0)
        aggregated[key]["Entries"] += 1
    stats_list = list(aggregated.values())
    for s in stats_list:
        s["TotalGames"] = s["Wins"] + s["Losses"]
        s["WinRate"] = (s["Wins"] / max(s["TotalGames"], 1)) * 100
    return stats_list

def generate_plots(data, output_dir, prefix="", filename_stem="", highlighted=None):
    if highlighted is None: highlighted = []
    try:
        os.makedirs(output_dir, exist_ok=True)
    except PermissionError as e:
        # Fallback to current working directory if absolute path fails
        alt_dir = os.path.join("plots", os.path.basename(os.path.dirname(output_dir)), os.path.basename(output_dir))
        print(f"WARNING: Permission denied for {output_dir}. Trying fallback: {alt_dir}")
        try:
            os.makedirs(alt_dir, exist_ok=True)
            output_dir = alt_dir
        except:
            print(f"CRITICAL ERROR: Could not create output directory: {e}")
            messagebox.showerror("Error", f"Could not create output directory:\n{output_dir}\n\nPlease check folder permissions.")
            return

    title_filter = "Full Data" if "Full Dataset" in prefix else "Meta Filter"
    title_header = f"{filename_stem} ({title_filter})"

    # 1. Win Rate by Leader Aspect
    aspect_stats = get_stats_local(data, "LeaderAspectsStr")
    aspect_stats.sort(key=lambda x: x["WinRate"], reverse=True)
    aspect_map = {}
    for row in data:
        as_str = row.get("LeaderAspectsStr")
        if as_str and as_str not in aspect_map: aspect_map[as_str] = row["LeaderAspects"]
    fig, ax = plt.subplots(figsize=(12, 7))
    for i, row in enumerate(aspect_stats):
        leader_aspects = aspect_map.get(row["LeaderAspectsStr"], [])
        color = get_aspect_color(leader_aspects)
        hatch = get_hatch_robust(row)
        edge_color = get_hatch_color_robust(row)
        ax.bar(i, row["WinRate"], color=color, hatch=hatch, edgecolor=edge_color, linewidth=1.5, label=row["LeaderAspectsStr"])
        ax.annotate(f'{row["WinRate"]:.1f}%', (i, row["WinRate"]), ha='center', va='bottom', xytext=(0, 5), textcoords='offset points', color=SW_COLORS["Text"], fontweight='bold', fontsize=10)
    ax.set_xticks(range(len(aspect_stats)))
    ax.set_xticklabels([s["LeaderAspectsStr"] for s in aspect_stats], rotation=45, ha='right')
    for label in ax.get_xticklabels():
        aspect_list = [a.strip() for a in label.get_text().split(',')]
        label.set_color(get_aspect_color(aspect_list, use_text_colors=True))
    ax.set_title(f"Win Rate by Opponent Leader Aspects (%)\n{title_header}")
    ax.set_ylabel("Win Rate (%)")
    ax.set_ylim(0, 105)
    ax.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axhline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axhline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    leg = ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', title="Aspects")
    for text_obj in leg.get_texts():
        text_obj.set_color(SW_COLORS["Text"])
        text = text_obj.get_text()
        aspect_list = [a.strip() for a in text.split(',')]
        text_obj.set_color(get_aspect_color(aspect_list, use_text_colors=True))
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader_aspect.png"), dpi=300)
    plt.close()

    # 1c. Win Rate by Deck Aspect (Top 10 Leaders)
    deck_leader_stats_full = [s for s in get_stats_local(data, ["LeaderNorm", "BaseAspect"]) if s.get("BaseAspect")]
    aspect_order_map = {"Cunning": 0, "Aggression": 1, "Command": 2, "Vigilance": 3}
    for s in deck_leader_stats_full: s["AspectOrder"] = aspect_order_map.get(s["BaseAspect"], 4)
    leader_pop = {}
    for s in deck_leader_stats_full: leader_pop[s["LeaderNorm"]] = leader_pop.get(s["LeaderNorm"], 0) + s["TotalGames"]
    top_10_names = sorted(leader_pop.keys(), key=lambda x: leader_pop[x], reverse=True)[:10]
    leader_order_list = sorted([{"n": n, "r": get_leader_sort_info(n)} for n in top_10_names], key=lambda x: (x["r"][0], x["r"][1], x["r"][2], x["n"]))
    top_10_leaders = [x["n"] for x in leader_order_list]
    leader_rank = {name: r for r, name in enumerate(top_10_leaders)}
    deck_leader_stats = [s for s in deck_leader_stats_full if s["LeaderNorm"] in top_10_leaders]
    deck_leader_stats.sort(key=lambda x: (leader_rank[x["LeaderNorm"]], x["AspectOrder"]))
    leader_aspect_map = {}
    for row in data:
        if row["LeaderNorm"] not in leader_aspect_map: leader_aspect_map[row["LeaderNorm"]] = row["LeaderAspects"]

    fig, ax = plt.subplots(figsize=(16, 12))
    y_pos = []
    curr_y = 0
    last_l = None
    for row in deck_leader_stats:
        if last_l and row["LeaderNorm"] != last_l: curr_y += 1.5
        else: curr_y += 1.0
        y_pos.append(curr_y)
        last_l = row["LeaderNorm"]
    for i, row in enumerate(deck_leader_stats):
        color = SW_COLORS.get(row["BaseAspect"], get_aspect_color(leader_aspect_map.get(row["LeaderNorm"], [])))
        ax.barh(y_pos[i], row["WinRate"], color=color, hatch=get_hatch_robust(row), edgecolor=get_hatch_color_robust(row), linewidth=1.5)
        ax.text(row["WinRate"] + 1, y_pos[i], f'{int(row["TotalGames"])} games', ha='left', va='center', color=SW_COLORS["Text"], fontweight='bold', fontsize=10)
    ax.set_yticks(y_pos)
    def get_deck_label(row):
        prefix_l = "★ " if highlighted and row["LeaderNorm"] in highlighted else ""
        return f"{prefix_l}{row['LeaderNorm']} ({row['BaseAspect']})"
    ax.set_yticklabels([get_deck_label(row) for row in deck_leader_stats], fontsize=10)
    for label in ax.get_yticklabels():
        l_name = label.get_text().split('(')[0].strip()
        label.set_color(get_aspect_color(leader_aspect_map.get(l_name, []), use_text_colors=True))
    ax.set_title(f"Top 10 Leaders by Games Played\n{title_header}")
    ax.set_xlim(0, 115)
    ax.axvline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axvline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axvline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "win_rate_by_deck_aspect.png"), dpi=300)
    plt.close()

    # 2. Top Opponent Leaders (Top 20)
    leader_stats_full = get_stats_local(data, "LeaderNorm")
    leader_base_stats = get_stats_local(data, ["LeaderNorm", "BaseAspect"])
    leader_stats_full.sort(key=lambda x: x["TotalGames"], reverse=True)
    if len(leader_stats_full) > 20:
        top_19 = leader_stats_full[:19]
        others = leader_stats_full[19:]
        top_19 = sorted([{"row": r, "rank": get_leader_sort_info(r["LeaderNorm"])} for r in top_19], key=lambda x: (x["rank"][0], x["rank"][1], x["rank"][2], x["row"]["LeaderNorm"]))
        leader_stats = [x["row"] for x in top_19]
        o_wins = sum(r["Wins"] for r in others)
        o_games = sum(r["TotalGames"] for r in others)
        leader_stats.append({"LeaderNorm": "Other", "Wins": o_wins, "TotalGames": o_games, "WinRate": (o_wins/max(o_games,1))*100, "Entries": len(others)})
    else:
        leader_stats = sorted([{"row": r, "rank": get_leader_sort_info(r["LeaderNorm"])} for r in leader_stats_full], key=lambda x: (x["rank"][0], x["rank"][1], x["rank"][2], x["row"]["LeaderNorm"]))
        leader_stats = [x["row"] for x in leader_stats]

    fig, ax = plt.subplots(figsize=(14, 10))
    for i, row in enumerate(leader_stats):
        l_name = row["LeaderNorm"]
        if l_name == "Other":
            ax.barh(i, row["WinRate"], color=SW_COLORS["Neutral"], edgecolor=SW_COLORS["Text"], linewidth=1.5)
        else:
            l_asp = leader_lookup.get(l_name.lower(), {}).get("aspects", [])
            h = get_hatch_robust(row)
            ec = get_hatch_color_robust(row)
            l_bases = [b for b in leader_base_stats if b["LeaderNorm"] == l_name]
            l_bases.sort(key=lambda x: x["BaseAspect"] or "")
            total_l_g = sum(b["TotalGames"] for b in l_bases)
            left = 0
            for b_row in l_bases:
                w = row["WinRate"] * (b_row["TotalGames"] / max(total_l_g, 1))
                ax.barh(i, w, left=left, color=SW_COLORS.get(b_row["BaseAspect"], get_aspect_color(l_asp)), hatch=h, edgecolor=ec, linewidth=1.5)
                left += w
        ax.annotate(f'{int(row["TotalGames"])} games', (row["WinRate"] + 1.5, i), ha='left', va='center', color=SW_COLORS["Text"], fontweight='bold', fontsize=10)
    
    def get_enriched(l_name, stats):
        if l_name == "Other":
            row = [r for r in stats if r["LeaderNorm"] == "Other"][0]
            return f"Other ({int(row['Entries'])} leaders)\n({int(row['TotalGames'])} games)"
        asp = leader_lookup.get(l_name.lower(), {}).get("aspects", [])
        tg = [r for r in stats if r["LeaderNorm"] == l_name][0]["TotalGames"]
        prefix_l = "★ " if highlighted and l_name in highlighted else ""
        return f"{prefix_l}{l_name} ({int(tg)} games)\n({format_aspects(asp)})" if asp else f"{prefix_l}{l_name} ({int(tg)} games)"

    ax.set_yticks(range(len(leader_stats)))
    ax.set_yticklabels([get_enriched(r["LeaderNorm"], leader_stats) for r in leader_stats])
    for label in ax.get_yticklabels():
        l_raw = label.get_text().split('(')[0].strip()
        label.set_color(get_aspect_color(leader_lookup.get(l_raw.lower(), {}).get("aspects", []), use_text_colors=True))
    ax.invert_yaxis()
    ax.set_title(f"Top 20 Opponent Leaders (by Games Played)\n{title_header}")
    ax.set_xlim(0, 115)
    ax.axvline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axvline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axvline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "win_rate_by_leader.png"), dpi=300)
    plt.close()

    # 3. Game Popularity Heatmap
    total_games_overall = sum(r["Wins"] + r["Losses"] for r in data)
    if total_games_overall > 0:
        matchup_data = {}
        all_l_aspects = set()
        all_b_aspects = set()
        for r in data:
            las = r.get("LeaderAspectsStr")
            ba = r.get("BaseAspect")
            if las and las != "Unknown" and ba:
                key = (las, ba)
                matchup_data[key] = matchup_data.get(key, 0) + (r["Wins"] + r["Losses"])
                all_l_aspects.add(las)
                all_b_aspects.add(ba)
        
        if matchup_data:
            def get_aspect_str_sort_info(aspect_str):
                aspect_list = [a.strip() for a in aspect_str.split(',')]
                align_p = {"Heroism": 0, "Villainy": 1}
                align_rank = min([align_p[a] for a in aspect_list if a in align_p]) if any(a in align_p for a in aspect_list) else 2
                asp_p = {"Vigilance": 0, "Command": 1, "Aggression": 2, "Cunning": 3}
                asp_rank = min([asp_p[a] for a in aspect_list if a in asp_p]) if any(a in asp_p for a in aspect_list) else 4
                return (align_rank, asp_rank, aspect_str)

            row_order = sorted(list(all_l_aspects), key=get_aspect_str_sort_info)
            col_order = sorted(list(all_b_aspects), key=lambda x: {"Vigilance": 0, "Command": 1, "Aggression": 2, "Cunning": 3}.get(x, 4))
            matrix = [[(matchup_data.get((r_asp, c_asp), 0) / total_games_overall) * 100 for c_asp in col_order] for r_asp in row_order]
            
            fig, ax = plt.subplots(figsize=(12, 10))
            im = ax.imshow(matrix, cmap="YlGnBu", vmin=0)
            fig.colorbar(im, ax=ax, label='Percentage of Total Games (%)')
            for i in range(len(row_order)):
                for j in range(len(col_order)):
                    val = matrix[i][j]
                    ax.text(j, i, f"{val:.1f}", ha="center", va="center", color="white" if val > 5 else "black")
            ax.set_xticks(range(len(col_order)))
            ax.set_xticklabels(col_order)
            ax.set_yticks(range(len(row_order)))
            ax.set_yticklabels(row_order)
            ax.set_xlabel("Base Aspect")
            ax.set_ylabel("Leader Aspects")
            for label in ax.get_yticklabels():
                label.set_color(get_aspect_color([a.strip() for a in label.get_text().split(',')], use_text_colors=True))
            for label in ax.get_xticklabels():
                label.set_color(TEXT_ASPECT_COLORS.get(label.get_text(), TEXT_ASPECT_COLORS["Neutral"]))
            ax.set_title(f"Game Sample Heatmap (%)\n{title_header}")
            plt.tight_layout(rect=[0, 0.03, 1, 1])
            fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
            plt.savefig(os.path.join(output_dir, "aspect_matchup_heatmap.png"), dpi=300)
            plt.close()

    # 4. Scatter Plot (Popularity vs Win Rate)
    fig, ax = plt.subplots(figsize=(16, 12))
    leader_stats_full.sort(key=lambda x: x["TotalGames"], reverse=True)
    if len(leader_stats_full) > 30:
        bubble_stats = leader_stats_full[:29]
        others_for_box = leader_stats_full[29:]
    else:
        bubble_stats = leader_stats_full
        others_for_box = []
    
    o_text = ""
    if others_for_box:
        o_w = sum(r["Wins"] for r in others_for_box)
        o_g = sum(r["TotalGames"] for r in others_for_box)
        o_text = f"Other Leaders Combined:\nCount: {len(others_for_box)}\nGames: {o_g}\nWin Rate: {(o_w/max(o_g,1))*100:.1f}%"
    
    total_w = sum(r["Wins"] for r in leader_stats_full)
    total_g = sum(r["TotalGames"] for r in leader_stats_full)
    ov_text = f"Overall Deck Statistics:\nGames: {total_g}\nWins: {int(total_w)}  Losses: {int(total_g-total_w)}\nWin Rate: {(total_w/max(total_g,1))*100:.1f}%"
    
    leader_base_counts = {}
    for row in data:
        k = (row["LeaderNorm"], row["BaseAspect"])
        leader_base_counts[k] = leader_base_counts.get(k, 0) + 1

    def draw_pie_at(ax_l, x, y, l_name, counts_dict):
        order = ["Cunning", "Aggression", "Command", "Vigilance"]
        ratios = [counts_dict.get((l_name, k), 0) for k in order]
        tot = sum(ratios) if sum(ratios) > 0 else 1.0
        fracs = [v/tot for v in ratios]
        da = DrawingArea(28, 28, 0, 0)
        t1 = 0
        for f, col in zip(fracs, [SW_COLORS[k] for k in order]):
            t2 = t1 + 360.0 * f
            if f > 0: da.add_artist(Wedge((14,14), 14, t1, t2, facecolor=col, edgecolor=SW_COLORS["Text"], linewidth=0.6))
            t1 = t2
        ax_l.add_artist(AnnotationBbox(da, (x, y), frameon=False))

    for r in bubble_stats:
        draw_pie_at(ax, r["TotalGames"], r["WinRate"], r["LeaderNorm"], leader_base_counts)
        prefix_l = "★ " if highlighted and r["LeaderNorm"] in highlighted else ""
        ax.annotate(f"{prefix_l}{r['LeaderNorm']}"[:25], (r["TotalGames"], r["WinRate"]), xytext=(14,14), textcoords='offset points', fontsize=10, fontweight='bold', rotation=10, color=get_aspect_color(get_leader_data(r["LeaderNorm"]).get("aspects", []), use_text_colors=True))
    
    ax.text(1.02, 0.5, ov_text + ("\n\n" + o_text if o_text else ""), transform=ax.transAxes, fontsize=11, fontweight='bold', color=SW_COLORS["Text"], bbox=dict(facecolor=SW_COLORS["Background"], edgecolor=SW_COLORS["Grid"], alpha=0.8), verticalalignment='center')
    ax.set_title(f"Leader Popularity vs Win Rate (%)\n{title_header}")
    ax.set_ylabel("Win Rate (%)")
    ax.set_xlabel("Total Games Played")
    ax.set_ylim(-10, 115)
    max_g = max([r["TotalGames"] for r in leader_stats_full]) if leader_stats_full else 100
    ax.set_xlim(-max_g*0.05, max_g*1.25)
    ax.grid(True, linestyle=':', alpha=0.3)
    ax.axhline(50, color=SW_COLORS["Grid"], linestyle='--', alpha=0.5)
    ax.axhline(60, color=SW_COLORS["WinRateHigh"], linestyle='--', alpha=0.5)
    ax.axhline(40, color=SW_COLORS["WinRateLow"], linestyle='--', alpha=0.5)
    plt.tight_layout(rect=[0, 0.03, 0.9, 1])
    fig.text(0.5, 0.01, DISCLAIMER_TEXT, ha='center', fontsize=8, color=SW_COLORS["Text"], alpha=0.7)
    plt.savefig(os.path.join(output_dir, "popularity_vs_winrate.png"), dpi=300)
    plt.close()

# --- UTILS ---
META_FILE = "meta_leaders.json"
def load_meta_leaders():
    if os.path.exists(META_FILE):
        try:
            with open(META_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except: pass
    return ["Darth Vader", "Luke Skywalker", "Boba Fett", "Sabine Wren", "Iden Versio", "Han Solo", "Leia Organa", "Grand Admiral Thrawn"]

def save_meta_leaders(leaders):
    try:
        with open(META_FILE, "w", encoding="utf-8") as f: json.dump(leaders, f, indent=2)
    except: pass

def combine_deck_aspects(row):
    asp = row.get("LeaderAspects")
    if not isinstance(asp, list): asp = get_leader_data(row.get("LeaderNorm", "")).get("aspects", [])
    combined = [str(a) for a in asp if a]
    if row.get("BaseAspect"): combined.append(str(row["BaseAspect"]))
    return format_aspects(combined)

def get_hatch_robust(row):
    asp = row.get("LeaderAspects")
    if not asp and "LeaderAspectsStr" in row: asp = [a.strip() for a in row["LeaderAspectsStr"].split(",") if a.strip()]
    if not asp: asp = get_leader_data(row.get("LeaderNorm", "")).get("aspects", [])
    return get_hatch(asp)

def get_hatch_color_robust(row):
    asp = row.get("LeaderAspects")
    if not asp and "LeaderAspectsStr" in row: asp = [a.strip() for a in row["LeaderAspectsStr"].split(",") if a.strip()]
    if not asp: asp = get_leader_data(row.get("LeaderNorm", "")).get("aspects", [])
    return get_alignment_color(asp)

def load_card_data():
    for f in ["all_cards_min.json", "all_cards.json"]:
        p = get_resource_path(f)
        if os.path.exists(p) and not os.path.isdir(p):
            with open(p, "r", encoding="utf-8") as f_in: return json.load(f_in)
    return []

def build_lookups(cards):
    l_l, b_l = {}, {}
    for c in cards:
        n, s, t = c.get("Name"), c.get("Subtitle"), c.get("Type")
        if t == "Leader":
            entry = {"aspects": sorted(c.get("Aspects", [])), "subtitle": s, "set": c.get("Set"), "number": c.get("Number")}
            if s: l_l[f"{n} | {s}".lower()] = l_l[strip_accents(f"{n} | {s}".lower())] = entry
            l_l[n.lower()] = l_l[strip_accents(n.lower())] = entry
        elif t == "Base":
            aspects = c.get("Aspects", [])
            b_l[n] = {"aspect": aspects[0] if aspects else None, "hp": c.get("HP")}
    return l_l, b_l

def process_match_data(files, l_l, b_l):
    rows = []
    for f_p in files:
        try:
            with open(f_p, 'r', encoding='utf-8-sig') as f:
                for r in csv.DictReader(f):
                    ln, bn = normalize_leader(r.get("OpponentLeader", "")), normalize_base(r.get("OpponentBase", ""))
                    ld = get_leader_data(ln)
                    ba = b_l.get(bn, {}).get("aspect")
                    if not ba:
                        for a in ["Aggression", "Cunning", "Vigilance", "Command"]:
                            if a.lower() in bn.lower(): ba = a; break
                    p_r = {"LeaderNorm": ln, "BaseNorm": bn, "LeaderAspects": ld.get("aspects"), "BaseAspect": ba, "Wins": safe_float(r.get("Wins")), "Losses": safe_float(r.get("Losses")), "LeaderAspectsStr": format_aspects(ld.get("aspects"))}
                    p_r["DeckAspects"] = combine_deck_aspects(p_r)
                    rows.append(p_r)
        except: pass
    return rows if rows else None

def main():
    global leader_lookup, base_lookup
    print("DEBUG: Starting main application...")
    files = filedialog.askopenfilenames(title="Select CSV(s)", filetypes=[("CSV", "*.csv"), ("All", "*.*")])
    if not files: return
    cards = load_card_data()
    if not cards: return
    leader_lookup, base_lookup = build_lookups(cards)
    stem = Path(files[0]).stem if len(files)==1 else f"Aggregated_{len(files)}_files_{Path(files[0]).stem}"
    print(f"DEBUG: Processing {len(files)} files into stem '{stem}'...")
    df = process_match_data(files, leader_lookup, base_lookup)
    if not df: 
        print("DEBUG: No data processed. Exiting.")
        return
    
    all_leaders = sorted(list(set(r["LeaderNorm"] for r in df)))
    print(f"DEBUG: Found {len(all_leaders)} unique leaders in dataset.")
    
    # Use absolute paths for directories to avoid PermissionError in some environments
    base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    plots_root = os.path.join(base_dir, "plots")
    
    full_dir = os.path.join(plots_root, stem, "full_report")
    meta_dir = os.path.join(plots_root, stem, "meta_report")
    meta_list = load_meta_leaders()
    print(f"DEBUG: Loaded {len(meta_list)} meta leaders from persistence.")
    
    print("DEBUG: Generating Full Dataset plots...")
    generate_plots(df, full_dir, "Full Dataset: ", stem, highlighted=meta_list)
    print(f"DEBUG: Full Dataset plots generated in {full_dir}")
    
    # 2. Meta Selection
    available_leaders = sorted(list(set(r["LeaderNorm"] for r in df)))
    selected_leaders = multi_select_listbox("Meta Analysis", "Select leaders for Meta Report:", available_leaders, initial_selections=meta_list)
    
    print(f"DEBUG: Selected leaders for Meta Report: {selected_leaders}")
    if selected_leaders:
        # Update persistent meta list: add selected, remove available-but-unselected
        new_meta = set(meta_list)
        for sl in selected_leaders: new_meta.add(sl)
        for al in available_leaders:
            if al not in selected_leaders and al in new_meta: new_meta.remove(al)
        
        save_meta_leaders(sorted(list(new_meta)))
        m_df = [r for r in df if r["LeaderNorm"] in selected_leaders]
        generate_plots(m_df, meta_dir, "Meta: ", stem, highlighted=selected_leaders)
        messagebox.showinfo("Success", f"Reports generated!\n\nFull: {full_dir}\nMeta: {meta_dir}")
    else:
        messagebox.showinfo("Success", f"Full report generated in:\n{full_dir}")
    print("DEBUG: Finalizing and exiting...")
    root.destroy()
    print("DEBUG: Application finished.")

if __name__ == "__main__":
    if "--test" in sys.argv:
        print("CLI TEST MODE: Loading data...")
        cards = load_card_data()
        if not cards:
            print("ERROR: Could not load card data!")
            sys.exit(1)
        leader_lookup, base_lookup = build_lookups(cards)
        print("SUCCESS: Data loaded. Exiting.")
        sys.exit(0)
    root = tk.Tk(); root.withdraw()
    main()
