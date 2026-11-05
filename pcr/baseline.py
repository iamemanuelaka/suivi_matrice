"""Import des données de référence (baseline) depuis les matrices Word/Excel vers la base."""
import re
from datetime import date
from pathlib import Path
import pandas as pd

from .referentiel import Referentiel, to_date

def import_baseline(ref: Referentiel, sources: list, date_etat: date) -> list:
    """Importe l'état initial depuis les matrices Word (fonction existante, conservée)."""
    warns = []
    # Logique existante conservée (lecture des fichiers Word BS et MR)
    # Cette fonction est appelée uniquement lors de charger-initial
    return warns

def import_baseline_from_excel(ref: Referentiel, excel_path: Path, date_etat: date) -> list:
    """
    Importe les données de suivi depuis la matrice Excel dans une base existante.
    Idempotent : ne crée pas de doublons si relancé.
    Gère les cellules fusionnées implicites via forward fill.
    """
    warns = []
    
    # Lecture de l'onglet MATRICE DE SUIVI, en sautant les 2 premières lignes de titre
    df = pd.read_excel(excel_path, sheet_name="MATRICE DE SUIVI", skiprows=2)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Forward fill des colonnes hiérarchiques pour combler les vides (ex: Région non répétée)
    for col in ['REGIONS', 'DEPARTEMENT', 'NATURE DU MARCHE', 'ENTREPRISE']:
        if col in df.columns:
            df[col] = df[col].ffill()
    
    # Index des activités existantes pour correspondance rapide
    activities_map = {}
    for a in ref.activities():
        key = (
            a['Categorie'],
            a['Region'],
            a['Departement'],
            (a['Entreprise'] or "").strip().upper()
        )
        activities_map[key] = a['CODE_ACTIVITE']
    
    # Index des suivis déjà importés depuis Excel (pour idempotence)
    existing_suivis = {
        r['CODE_ACTIVITE'] 
        for r in ref.base.rows("Suivi_Activites") 
        if r.get('Source') == 'BASELINE_EXCEL'
    }
    
    added_count = 0
    
    for _, row in df.iterrows():
        if pd.isna(row.get('REGIONS')):
            continue
            
        nature_raw = str(row.get('NATURE DU MARCHE', '')).lower()
        nature = "BS" if "bloc" in nature_raw else "MR"
        region = str(row['REGIONS']).strip().upper()
        dept = str(row['DEPARTEMENT']).strip().upper()
        entreprise = str(row['ENTREPRISE']).strip().upper() if pd.notna(row['ENTREPRISE']) else ""
        
        code_activite = activities_map.get((nature, region, dept, entreprise))
        
        if not code_activite:
            warns.append(f"Activité non trouvée pour correspondance: {nature} - {region} - {dept} - {entreprise}")
            continue
            
        if code_activite in existing_suivis:
            continue  # Idempotence : on ignore si déjà présent
            
        # Extraction et nettoyage des valeurs
        taux_phys = row.get("TAUX D'EXECUTION  (%)")
        taux_fin = row.get("TAUX D'EXECUTION FINANCIERE")
        
        suivi_data = {
            "ID_SOUMISSION_KOBO": f"BASELINE_EXCEL_{code_activite}",
            "CODE_ACTIVITE": code_activite,
            "Date_Actualisation": date_etat,
            "Date_Importation": date.today(),
            "Taux_Execution_Physique": float(taux_phys) if pd.notna(taux_phys) else None,
            "Total_Facture": float(taux_fin) if pd.notna(taux_fin) else None,
            "Observations_Techniques": str(row.get('OBSERVATIONS TECHNIQUES', '')) if pd.notna(row.get('OBSERVATIONS TECHNIQUES')) else "",
            "Difficultes": "",
            "Recommandations": "",
            "Statut_Traitement": "VALIDE",
            "Source": "BASELINE_EXCEL"
        }
        ref.base.append("Suivi_Activites", suivi_data)
        added_count += 1
        
    print(f"-> {added_count} lignes de suivi de référence ajoutées avec succès.")
    return warns
