import pandas as pd
import numpy as np

def load_all_data(filename="database.xlsx"):
    """Charge toutes les feuilles d'un fichier Excel dans un dictionnaire de DataFrames."""
    try:
        xls = pd.ExcelFile(filename)
        data_frames = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
        return data_frames
    except FileNotFoundError:
        print(f"Erreur : Le fichier '{filename}' est introuvable.")
        return None
    except Exception as e:
        print(f"Une erreur est survenue lors du chargement des données : {e}")
        return None

def perform_calculations(data_frames):
    """Effectue les calculs de section et de capacité sur les DataFrames."""
    try:
        cables_df = data_frames['Liste_Cables']
        cables_df['Section_mm2'] = np.pi * (cables_df['Diametre_mm'] / 2)**2

        trays_df = data_frames['Chemins_de_Cables']
        reserve_rate = data_frames['Parametres']['Valeur'][0]

        trays_df['Capacite_Totale_mm2'] = trays_df['Largeur_mm'] * trays_df['Hauteur_mm']
        trays_df['Capacite_Utile_mm2'] = trays_df['Capacite_Totale_mm2'] * (1 - reserve_rate)

        return data_frames
    except KeyError as e:
        print(f"Erreur de calcul : une colonne ou une feuille attendue est manquante ({e}).")
        return None
    except Exception as e:
        print(f"Une erreur est survenue lors des calculs : {e}")
        return None

def generate_report_for_tray(tray_id, data_frames):
    """Génère un rapport de remplissage pour un chemin de câble spécifique."""
    try:
        assignments_df = data_frames['Assignation']
        cables_df = data_frames['Liste_Cables']
        trays_df = data_frames['Chemins_de_Cables']

        target_tray = trays_df[trays_df['ID_Chemin'] == tray_id]
        if target_tray.empty:
            return f"\nErreur : Le chemin de câble '{tray_id}' n'a pas été trouvé. Veuillez réessayer."

        usable_capacity = target_tray['Capacite_Utile_mm2'].iloc[0]

        assigned_cables_mask = assignments_df['Chemins_de_cable_assignes'].str.contains(tray_id, na=False)
        cables_in_tray_ids = assignments_df[assigned_cables_mask]['ID_Cable']

        total_fill = 0
        if not cables_in_tray_ids.empty:
            cable_sections = cables_df[cables_df['ID_Cable'].isin(cables_in_tray_ids)]
            total_fill = cable_sections['Section_mm2'].sum()

        remaining_capacity = usable_capacity - total_fill
        fill_percentage = (total_fill / usable_capacity) * 100 if usable_capacity > 0 else 0

        report = []
        report.append("\n" + "="*40)
        report.append(f"RAPPORT POUR LE CHEMIN DE CÂBLE : {tray_id}")
        report.append("="*40)
        report.append(f"Capacité utile (avec réserve) : {usable_capacity:,.2f} mm²")
        report.append(f"Remplissage actuel : {total_fill:,.2f} mm²")
        report.append(f"Capacité restante : {remaining_capacity:,.2f} mm²")
        report.append(f"Taux de remplissage : {fill_percentage:.2f}%")
        report.append("-"*40)
        report.append("Câbles dans ce chemin :")
        if not cables_in_tray_ids.empty:
            for cable_id in cables_in_tray_ids.tolist():
                report.append(f"  - {cable_id}")
        else:
            report.append("  - Aucun")
        report.append("="*40)

        return "\n".join(report)

    except KeyError as e:
        return f"Erreur de rapport : une colonne ou une feuille attendue est manquante ({e})."
    except Exception as e:
        return f"Une erreur est survenue lors de la génération du rapport : {e}"

def main():
    """
    Fonction principale pour exécuter l'application CLI.
    """
    print("Démarrage de l'application de gestion de chemins de câbles...")
    all_data = load_all_data()
    if not all_data:
        return

    all_data = perform_calculations(all_data)
    if not all_data:
        return

    print("Initialisation terminée.\n")

    tray_ids = all_data['Chemins_de_Cables']['ID_Chemin'].tolist()

    while True:
        print("Chemins de câbles disponibles :")
        for tid in tray_ids:
            print(f"  - {tid}")

        prompt = "\nEntrez l'ID du chemin de câble que vous souhaitez analyser (ou 'q' pour quitter) : "
        user_input = input(prompt)

        if user_input.lower() == 'q' or user_input.lower() == 'quitter':
            print("Arrêt de l'application. Au revoir !")
            break

        report = generate_report_for_tray(user_input, all_data)
        print(report)

        input("\nAppuyez sur Entrée pour continuer...")

if __name__ == "__main__":
    main()
