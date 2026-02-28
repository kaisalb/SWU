import pandas as pd
import json
import os
import re

def normalize_card_number(num):
    if not num:
        return "000"
    num_str = str(num).split('.')[0]
    return num_str.zfill(3)

def extract_subtitle(name):
    # Matches "Name | Subtitle" or "Name - Subtitle"
    if ' | ' in name:
        parts = name.split(' | ', 1)
        return parts[0].strip(), parts[1].strip()
    elif ' - ' in name:
        parts = name.split(' - ', 1)
        return parts[0].strip(), parts[1].strip()
    return name.strip(), None

def infer_alignment(traits):
    # Common traits for Heroism and Villainy
    hero_traits = ["REBEL", "REPUBLIC", "RESISTANCE", "JEDI"]
    villain_traits = ["IMPERIAL", "SEPARATIST", "FIRST ORDER", "SITH", "BOUNTY HUNTER", "UNDERWORLD", "TROOPER"]
    
    for t in traits:
        if any(h in t.upper() for h in hero_traits):
            return "Heroism"
        if any(v in t.upper() for v in villain_traits):
            return "Villainy"
    return None

def process_set(xl, sheet_name, set_code):
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
    
    df.columns = df.iloc[header_row_idx]
    df = df.iloc[header_row_idx + 1:].reset_index(drop=True)
    
    processed_cards = []
    for _, row in df.iterrows():
        card_num = normalize_card_number(row.get("Card #"))
        if not card_num or card_num == "000":
            continue
            
        full_name = str(row.get("Card Name", "")).strip()
        if not full_name or full_name == "nan":
            continue
            
        name, subtitle = extract_subtitle(full_name)
        
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
            
        # Aspects
        aspects_val = str(row.get("Aspect(s)", "")).strip()
        aspects = [a.strip().title() for a in aspects_val.split(',') if a.strip()]
        
        # Traits
        traits_val = str(row.get("Trait(s)", "")).strip()
        traits = [t.strip().upper() for t in re.split(r'[,/]', traits_val) if t.strip()]
        
        # Infer Heroism/Villainy for Leaders/Bases if missing
        if final_type in ["Leader", "Base"]:
            has_alignment = any(a in ["Heroism", "Villainy"] for a in aspects)
            if not has_alignment:
                alignment = infer_alignment(traits)
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
        "5 - Legends of the Force": "LOTF",
        "6 - Secrets of Power": "SOP",
        "7 - A Lawless Time": "LAW",
        "Promos": "PRM"
    }
    
    all_cards = []
    for sheet_name, set_code in sheet_map.items():
        if sheet_name in xl.sheet_names:
            all_cards.extend(process_set(xl, sheet_name, set_code))
        else:
            print(f"Warning: Sheet {sheet_name} not found in Excel file.")
            
    with open("all_cards.json", "w", encoding="utf-8") as f:
        json.dump(all_cards, f, indent=2)
        
    print(f"Successfully processed {len(all_cards)} cards and updated all_cards.json")

if __name__ == "__main__":
    main()
