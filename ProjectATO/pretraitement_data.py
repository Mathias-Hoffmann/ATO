import pandas as pd
import matplotlib.pyplot as plt

# Chemin vers le fichier Excel (assure-toi que DATA_ATO.xlsx est dans le même dossier que ce script)
file_path = "DATA_ATO.xlsx"

# Charger le fichier Excel
xls = pd.ExcelFile(file_path)
print("Feuilles disponibles :", xls.sheet_names)

# -------------------------------
# 1. Lecture et affichage des données
# -------------------------------

# Holding costs
holding_costs_df = xls.parse('Holding costs')
print("\nHolding costs:")
print(holding_costs_df.head())

# Lead Times Distirbution
lead_times_df = xls.parse('Lead Times Distirbution')
# Forcer les noms de colonnes à être des chaînes et enlever les espaces superflus
lead_times_df.columns = [str(col).strip() for col in lead_times_df.columns]
print("\nLead Times Distirbution:")
print(lead_times_df.head())

# Bill of Materials (BOM)
bom_df = xls.parse('BOM')
print("\nBill of Materials (BOM):")
print(bom_df.head())

# Demand
demand_df = xls.parse('Demand')
print("\nDemand:")
print(demand_df.head())

# -------------------------------
# 2. Vérification des distributions des délais de livraison
# -------------------------------
print("\nVérification des distributions des délais de livraison:")
# Récupérer dynamiquement les colonnes de probabilités (toutes sauf la première)
prob_cols = lead_times_df.columns[1:]
print("Colonnes de probabilités utilisées :", list(prob_cols))

for index, row in lead_times_df.iterrows():
    prob_sum = row[prob_cols].sum()
    comp = row['Component\\Lead time']
    print(f"Composant {comp}: Somme des probabilités = {prob_sum}")

# -------------------------------
# 3. Visualisation de la demande pour le produit A
# -------------------------------
# On suppose que la colonne 'End-item\\period' contient les identifiants des produits finis
if 'End-item\\period' in demand_df.columns:
    demand_df.set_index('End-item\\period', inplace=True)
    print("\nProduits disponibles dans la demande :", list(demand_df.index))
    
    if 'A' in demand_df.index:
        demand_A = demand_df.loc['A']
        plt.figure(figsize=(8, 4))
        demand_A.plot(marker='o')
        plt.title("Demande pour le produit A")
        plt.xlabel("Semaine")
        plt.ylabel("Nombre d'unités")
        plt.grid(True)
        plt.show()
    else:
        print("Le produit 'A' n'a pas été trouvé dans la feuille 'Demand'.")
else:
    print("La colonne 'End-item\\period' n'existe pas dans la feuille 'Demand'.")
