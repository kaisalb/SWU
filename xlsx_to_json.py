import pandas as pd
import json
import os
import re

def normalize_card_number(num):
    if not num:
        return "000"
    num_str = str(num).split('.')[0]
    return num_str.zfill(3)

def extract_subtitle(name, row=None):
    # Matches "Name | Subtitle" or "Name - Subtitle"
    if ' | ' in name:
        parts = name.split(' | ', 1)
        return parts[0].strip(), parts[1].strip()
    elif ' - ' in name:
        parts = name.split(' - ', 1)
        return parts[0].strip(), parts[1].strip()
    
    # Check "Subtitle" column if it exists in the row
    if row is not None:
        subtitle_val = str(row.get("Subtitle", "")).strip()
        if subtitle_val and subtitle_val.lower() != 'nan':
            return name.strip(), subtitle_val
            
    return name.strip(), None

def infer_alignment(traits, card_name=""):
    # Common traits for Heroism and Villainy
    hero_traits = ["REBEL", "REPUBLIC", "RESISTANCE", "JEDI", "NABOO", "NEW REPUBLIC"]
    villain_traits = ["IMPERIAL", "SEPARATIST", "FIRST ORDER", "SITH", "BOUNTY HUNTER", "UNDERWORLD", "TROOPER", "INQUISITOR"]
    
    # Check name for explicit hints if traits are ambiguous or missing
    # DEACTIVATED for debugging or if Excel is supposed to have everything
    # if "Saw Gerrera" in card_name: return "Heroism" 
    # if "Tobias Beckett" in card_name: return "Villainy"
    # if "Sebulba" in card_name: return "Villainy"
    
    for t in traits:
        if any(h in t.upper() for h in hero_traits):
            return "Heroism"
        if any(v in t.upper() for v in villain_traits):
            return "Villainy"
    return None

def process_set(xl, sheet_name, set_code, subtitle_lookup=None):
    print(f"Processing set: {sheet_name} ({set_code})")
    df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
    
    # Find header row
    header_row_idx = None
    for i, row in df.iterrows():
        if any(str(cell).strip() == "Card #" for cell in row):
            header_row_idx = i
            break
    
    if header_row_idx is None:
        print(f"Warning: Could not find header row in {sheet_name}")
        return []
    
    # We use columns from the original df to find header info
    header_row_info = [str(h) for h in df.iloc[header_row_idx].tolist()]
    aspect_idx = -1
    card_name_idx = -1
    for i, h in enumerate(header_row_info):
        if "Aspect(s)" in h:
            aspect_idx = i
        if "Card Name" in h:
            card_name_idx = i
            
    # Handle the case where Card Name is actually multiple columns due to merging
    # In some sets, there are empty columns between Aspect(s) and Card Name
    # We want to stop at the FIRST column that contains "Card Name"
    
    df.columns = header_row_info
    df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
    
    processed_cards = []
    for _, row in df.iterrows():
        card_num = normalize_card_number(row.get("Card #"))
        if not card_num or card_num == "000":
            continue
            
        full_name = str(row.get("Card Name", "")).strip()
        if isinstance(full_name, pd.Series): # Should not happen with current logic but being safe
            full_name = str(full_name.iloc[0]).strip()
            
        if not full_name or full_name == "nan":
            continue
            
        name, subtitle = extract_subtitle(full_name, row)
        
        # Determine Type
        type_str = str(row.get("Type(s)", "")).strip()
        final_type = "Unit"
        if "Leader" in type_str:
            final_type = "Leader"
        elif "Base" in type_str:
            final_type = "Base"
        elif "Upgrade" in type_str:
            final_type = "Upgrade"
        elif "Event" in type_str:
            final_type = "Event"

        # Override subtitle from lookup if available (using Set and Card Number)
        if final_type == "Leader" and subtitle_lookup:
            key = f"{set_code}_{card_num}"
            if key in subtitle_lookup:
                subtitle = subtitle_lookup[key]
            
        # Traits
        traits_val = str(row.get("Trait(s)", "")).strip()
        traits = [t.strip().upper() for t in re.split(r'[,/]', traits_val) if t.strip()]
        
        # Primary aspect is at aspect_idx, alignment is usually at aspect_idx + 1 or + 2
        # BUT we must be careful not to overshoot into Card Name or other columns
        raw_aspects = []
        if aspect_idx != -1:
            # Check columns between aspect_idx and card_name_idx (exclusive)
            # In some sets, card_name_idx might be aspect_idx + 3 (as in LAW)
            # In others, it might be aspect_idx + 2 (as in SHD)
            limit = card_name_idx if card_name_idx != -1 else aspect_idx + 3
            for i in range(aspect_idx, limit):
                if i >= len(row): break
                val = str(row.iloc[i]).strip()
                if val and val.lower() != 'nan':
                    # Split by comma in case multiple aspects are in one cell
                    # but ALSO skip if the value is the same as the card name (just in case)
                    parts = [a.strip().title() for a in val.split(',')]
                    for p in parts:
                        if p and p.lower() != name.lower():
                            raw_aspects.append(p)
            
        aspects = sorted(list(set([a for a in raw_aspects if a])))

        # Infer Heroism/Villainy for Leaders/Bases if missing
        # BUT only if it wasn't already in the Excel file
        if final_type in ["Leader", "Base"]:
            has_alignment = any(a in ["Heroism", "Villainy"] for a in aspects)
            if not has_alignment:
                alignment = infer_alignment(traits, name)
                if alignment:
                    aspects.append(alignment)
        
        # Arenas
        arena_val = str(row.get("Arena", "")).strip()
        arenas = [arena_val.title()] if arena_val and arena_val != "nan" else []
        
        # Stats
        cost = str(row.get("Cost", "")).split('.')[0]
        power = str(row.get("Power", "")).split('.')[0]
        hp = str(row.get("HP", "")).split('.')[0]
        
        card_data = {
            "Set": set_code,
            "Number": card_num,
            "Name": name,
            "Subtitle": subtitle,
            "Type": final_type,
            "Aspects": sorted(aspects),
            "Traits": traits,
            "Arenas": arenas,
            "Cost": cost if cost != "nan" else None,
            "Power": power if power != "nan" else "-",
            "HP": hp if hp != "nan" else "-",
            "FrontText": str(row.get("Ability", "-")),
            "Rarity": str(row.get("Rarity", None)),
            "Unique": True if final_type in ["Leader", "Base"] or "Unique" in str(row.get("Unique", "")) else False,
            "VariantType": "Normal"
        }
        processed_cards.append(card_data)
        
    return processed_cards

def main():
    excel_file = "all_sets.xlsx"
    if not os.path.exists(excel_file):
        print(f"Error: {excel_file} not found.")
        return
        
    xl = pd.ExcelFile(excel_file)
    sheet_map = {
        "1  - Spark of Rebellion": "SOR",
        "2 - Shadows of the Galaxy": "SHD",
        "3 - Twilight of the Republic": "TWI",
        "4 - Jump to Lightspeed": "JTL",
        "5 - Legends of the Force": "LOF",
        "6 - Secrets of Power": "SEC",
        "7 - A Lawless Time": "LAW",
        "Promos": "PRM"
    }
    
    # Load leader subtitles if mapping file exists
    subtitle_lookup = {}
    if os.path.exists("leader_subtitles.json"):
        with open("leader_subtitles.json", "r", encoding="utf-8") as f:
            subtitle_lookup = json.load(f)
            
    all_cards = []
    for sheet_name, set_code in sheet_map.items():
        if sheet_name in xl.sheet_names:
            all_cards.extend(process_set(xl, sheet_name, set_code, subtitle_lookup))
        else:
            print(f"Warning: Sheet {sheet_name} not found in Excel file.")
            
    with open("all_cards.json", "w", encoding="utf-8") as f:
        json.dump(all_cards, f, indent=2)
        
    print(f"Successfully processed {len(all_cards)} cards and updated all_cards.json")

if __name__ == "__main__":
    main()
