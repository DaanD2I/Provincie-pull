import pgeocode
import pandas as pd
import openpyxl

nomi = pgeocode.Nominatim("nl")
df = nomi._data.copy()

# Alleen PC4 gebruiken
df["postcode4"] = df["postal_code"].astype(str).str.strip()

province_map = {
    "North Holland": "Noord-Holland",
    "South Holland": "Zuid-Holland",
    "North Brabant": "Noord-Brabant",
    "Zeeland": "Zeeland",
    "Utrecht": "Utrecht",
    "Guelders": "Gelderland",
    "Overijssel": "Overijssel",
    "Flevoland": "Flevoland",
    "Friesland": "Friesland",
    "Groningen": "Groningen",
    "Drenthe": "Drenthe",
    "Limburg": "Limburg",
}

# Provincie vertalen
df["provincie"] = df["state_name"].map(province_map)

def fix_municipality(x):
    if pd.isna(x):
        return None
    s = str(x).strip()

    if "Municipality" in s:
        s = s.replace("Municipality", "").strip(" ,-/")
        return f"Gemeente {s}" if s else "Gemeente"

    return s

# Gemeente opschonen
df["gemeente"] = df["county_name"].apply(fix_municipality)

# Tabel bouwen
postcode_table = df[["postcode4", "provincie", "gemeente"]].copy()

# Sorteer zodat je een consistente "eerste" kiest per postcode4
postcode_table = postcode_table.sort_values(["postcode4", "provincie", "gemeente"])

# Duplicates eruit: 1 regel per postcode4
postcode_table = postcode_table.drop_duplicates(subset=["postcode4"], keep="first")


# Naar Excel (.xlsx)
output_file = "nl_postcode4_provincie_gemeente.xlsx"
postcode_table.to_excel(output_file, index=False, sheet_name="Postcode4")

print(f"Done")