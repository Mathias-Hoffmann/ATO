import pandas as pd
import pulp

# Chargement des données depuis DATA_ATO.xlsx
file_path = "DATA_ATO.xlsx"
xls = pd.ExcelFile(file_path)

# Définir l'horizon de planification (en semaines)
T = 10

# Lecture des différentes feuilles
holding_costs_df = xls.parse('Holding costs')
lead_times_df = xls.parse('Lead Times Distirbution')
bom_df = xls.parse('BOM')
demand_df = xls.parse('Demand')

# Prétraitement de la feuille Demand :
demand_df.set_index('End-item\\period', inplace=True)
end_items = [str(item) for item in demand_df.index.tolist()]
print("Produits finis (index) :", end_items)

# Prétraitement du BOM :
# On conserve uniquement les composants listés dans Holding costs
components = holding_costs_df['Component'].astype(str).tolist()
bom_df.set_index('Component\\End-item', inplace=True)
print("Composants (d'après Holding costs) :", components)

# Extraire les composants de la feuille Holding costs
components_holding = set(holding_costs_df['Component'].astype(str).tolist())
# Extraire les composants de la feuille Lead Times Distirbution
components_leadtimes = set(lead_times_df['Component\\Lead time'].astype(str).tolist())

# On prend l'intersection pour ne garder que les composants pour lesquels nous avons toutes les données
components = list(components_holding.intersection(components_leadtimes))
print("Composants considérés :", components)


# Calcul de la demande en composants pour chaque composant et chaque semaine
DemandComp = {}
for c in components:
    DemandComp[c] = {}
    for t in range(1, T+1):
        week_key = t  # Les colonnes dans demand_df sont des entiers (1, 2, ..., 15)
        total_demand = 0
        for product in end_items:
            # Essayer d'accéder à la quantité dans le BOM pour le couple (composant, produit)
            try:
                qty = bom_df.loc[c, product]
            except KeyError:
                print(f"Avertissement : Clé '{product}' non trouvée dans BOM pour le composant {c}. On considère qty = 0.")
                qty = 0
            # Récupération de la demande pour le produit en semaine t
            try:
                demand_val = demand_df.loc[product, week_key]
            except KeyError:
                print(f"Avertissement : Pour le produit {product} ou la semaine {week_key}, la demande n'est pas trouvée. On considère demand_val = 0.")
                demand_val = 0
            total_demand += qty * demand_val
        DemandComp[c][t] = total_demand

# Calcul du délai fractile L_c^(90) pour chaque composant
L90 = {}
# Forcer les noms de colonnes de lead_times_df en chaînes et sans espaces superflus
lead_times_df.columns = [str(col).strip() for col in lead_times_df.columns]
for idx, row in lead_times_df.iterrows():
    comp = str(row['Component\\Lead time']).strip()
    cum_prob = 0.0
    L_val = None
    for l in [1, 2, 3, 4, 5]:
        key = str(l)
        value = row[key] if key in row.index else 0  # On vérifie l'existence de la clé
        cum_prob += value
        if cum_prob >= 0.90:
            L_val = l
            break
    if L_val is None:
        L_val = 5  # Valeur par défaut si 90% n'est pas atteint
    L90[comp] = L_val

# Extraction des coûts de stockage pour chaque composant
holding_cost = {}
for idx, row in holding_costs_df.iterrows():
    comp = str(row['Component']).strip()
    holding_cost[comp] = row['holding costs']

# Formulation du modèle d'optimisation avec PuLP
prob = pulp.LpProblem("ATO_Optimization", pulp.LpMinimize)

# Déclaration des variables de décision :
# Q_vars[c][t] : Quantité à commander pour le composant c en semaine t
# I_vars[c][t] : Niveau de stock du composant c à la fin de la semaine t
Q_vars = {}
I_vars = {}
for c in components:
    Q_vars[c] = {}
    I_vars[c] = {}
    for t in range(1, T+1):
        Q_vars[c][t] = pulp.LpVariable(f"Q_{c}_{t}", lowBound=0, cat="Continuous")
        I_vars[c][t] = pulp.LpVariable(f"I_{c}_{t}", lowBound=0, cat="Continuous")

# Fonction objectif : minimiser le coût total de stockage
prob += pulp.lpSum(holding_cost[c] * I_vars[c][t] for c in components for t in range(1, T+1)), "Total_Holding_Cost"

# Contraintes d'évolution du stock pour chaque composant et chaque semaine
for c in components:
    for t in range(1, T+1):
        if t - L90[c] >= 1:
            order_sum = pulp.lpSum(Q_vars[c][u] for u in range(1, t - L90[c] + 1))
        else:
            order_sum = 0
        prev_inventory = 0 if t == 1 else I_vars[c][t-1]
        prob += I_vars[c][t] == prev_inventory + order_sum - DemandComp[c][t], f"Inventory_Balance_{c}_{t}"

# Résolution du modèle
prob.solve()

# Affichage des résultats
print("Status :", pulp.LpStatus[prob.status])
for c in components:
    for t in range(1, T+1):
        print(f"Composant {c}, Semaine {t} : Commande Q = {Q_vars[c][t].varValue:.2f}, Stock I = {I_vars[c][t].varValue:.2f}")
