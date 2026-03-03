import pgeocode
import pandas as pd

nomi = pgeocode.Nominatim("be")
df = nomi._data.copy()

# PC4
df["postcode4"] = df["postal_code"].astype(str).str.strip()

# Regio (state_name) -> NL
region_map = {
    "Bruxelles-Capitale": "Brussels Hoofdstedelijk Gewest",
    "Vlaanderen": "Vlaanderen",
    "Wallonie": "Wallonië",
}
df["regio"] = df["state_name"].map(region_map).fillna(df["state_name"])

# Provincie (county_name) -> NL
province_map_be = {
    "Anvers": "Antwerpen",
    "Brabant Flamand": "Vlaams-Brabant",
    "Limbourg": "Limburg",
    "Flandre-Orientale": "Oost-Vlaanderen",
    "Flandre-Occidentale": "West-Vlaanderen",
    "Brabant Wallon": "Waals-Brabant",
    "Hainaut": "Henegouwen",
    "Liège": "Luik",
    "Luxembourg": "Luxemburg",
    "Namur": "Namen",
    # speciale Brussel-waarde
    "Bruxelles (19 communes)": "Brussels Hoofdstedelijk Gewest",
}
df["provincie"] = df["county_name"].map(province_map_be).fillna(df["county_name"])

# Plaatsnaam bewaren (optioneel maar vaak handig)
df["plaats"] = df["place_name"]

# Tabel maken
postcode_table = df[["postcode4", "regio", "provincie", "plaats"]].copy()

# Consistente dedupe
postcode_table = postcode_table.sort_values(["postcode4", "provincie", "plaats"])

# ✅ 1 regel per postcode4
postcode_table = postcode_table.drop_duplicates(subset=["postcode4"], keep="first")

# Export naar Excel
output_file = "be_postcode4_regio_provincie_plaats.xlsx"
postcode_table.to_excel(output_file, index=False, sheet_name="Postcode4")

print(f"Klaar 🎉 Excel opgeslagen als: {output_file}")