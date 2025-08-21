import pandas as pd

def create_excel_database(filename="database.xlsx"):
    """
    Creates an Excel file with predefined sheets and data for the cable manager project.
    """
    try:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            # --- Parametres Sheet ---
            params_data = {'Description': ['Taux de réserve'], 'Valeur': [0.20]}
            params_df = pd.DataFrame(params_data)
            params_df.to_excel(writer, sheet_name='Parametres', index=False)

            # --- Liste_Cables Sheet ---
            cables_data = {
                'ID_Cable': ['W001', 'W002', 'W003', 'W004'],
                'Type_Cable': ['U-1000 R2V 3G1.5', 'U-1000 R2V 3G2.5', 'H07V-K 1x16', 'H07V-K 1x25'],
                'Diametre_mm': [7.8, 8.8, 7.5, 9.2]
            }
            cables_df = pd.DataFrame(cables_data)
            cables_df.to_excel(writer, sheet_name='Liste_Cables', index=False)

            # --- Chemins_de_Cables Sheet ---
            trays_data = {
                'ID_Chemin': ['CDG-01-A', 'CDG-01-B', 'CDT-01-A', 'CDT-01-B'],
                'Largeur_mm': [150, 150, 300, 200],
                'Hauteur_mm': [50, 50, 100, 80]
            }
            trays_df = pd.DataFrame(trays_data)
            trays_df.to_excel(writer, sheet_name='Chemins_de_Cables', index=False)

            # --- Assignation Sheet ---
            assignments_data = {
                'ID_Cable': ['W001', 'W002', 'W003', 'W004'],
                'Chemins_de_cable_assignes': [
                    'CDG-01-A/CDG-01-B/CDT-01-A',
                    'CDG-01-A/CDG-01-B',
                    'CDT-01-A/CDT-01-B',
                    'CDT-01-A/CDT-01-B'
                ]
            }
            assignments_df = pd.DataFrame(assignments_data)
            assignments_df.to_excel(writer, sheet_name='Assignation', index=False)

        print(f"Fichier '{filename}' créé avec succès avec 4 feuilles.")

    except Exception as e:
        print(f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    create_excel_database()
