import pandas as pd

# Charger le fichier Excel
file_path = "DATA_ATO.xlsx"
xls = pd.ExcelFile(file_path)

# Charger les feuilles Demand et BOM
demand_df = xls.parse('Demand')
bom_df = xls.parse('BOM')

# Afficher les colonnes de Demand
print("Colonnes de Demand :", demand_df.columns.tolist())

# Adapte le nom de la colonne identifiant les produits finis, par exemple :
if 'End-item\\period' in demand_df.columns:
    demand_df.set_index('End-item\\period', inplace=True)
elif 'End-item' in demand_df.columns:
    demand_df.set_index('End-item', inplace=True)
else:
    print("La colonne identifiant les produits finis n'a pas été trouvée.")

print("Produits finis (index) :", demand_df.index.tolist())

# Afficher les colonnes de BOM
print("Colonnes de BOM :", bom_df.columns.tolist())
