
"""Import des données de référence depuis les matrices Markdown.

Les matrices sources sont issues des matrices Word converties en texte.
Le module :
- nettoie le balisage Markdown ;
- lit les lignes des matrices BS et MR ;
- récupère les observations techniques, les taux et les commentaires ;
- importe les suivis sans créer de doublons ;
- conserve les anomalies dans une liste d'avertissements.
"""

import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

from .referentiel import Referentiel, to_date
from .seed import parse_date_loose


# ---------------------------------------------------------------------
# Nettoyage et conversions
# ---------------------------------------------------------------------

def clean(value) -> str:
    """Nettoie le balisage Markdown et les espaces sans altérer le contenu."""
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\u00a0", " ").replace("\u202f", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"\*+", "", text)
    text = re.sub(r"(?<!\w)_(.*?)_(?!\w)", r"\1", text)
    text = text.replace("**", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _empty(value) -> bool:
    return value is None or not str(value).strip()


def _number(value):
    """Convertit un montant ou un taux en nombre ; renvoie None si absent."""
    if _empty(value):
        return None

    text = clean(value)
    text = text.replace("%", "").replace("FCFA", "").replace("HT", "")
    text = text.replace(",", ".")
    text = re.sub(r"\s+", "", text)

    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None

    number = float(match.group())
    return int(number) if number.is_integer() else number


def _date_from_text(text):
    """Extrait la première date valide d'un texte."""
    if not text:
        return None

    result, _ = parse_date_loose(str(text))
    return result


def _split_markdown_row(line):
    """Découpe une ligne de tableau Markdown en cellules."""
    if not line.lstrip().startswith("|"):
        return []

    return [clean(cell) for cell in line.strip().strip("|").split("|")]


def _extract_region(text):
    """Extrait le nom de région placé au début de la cellule de désignation."""
    text = clean(text)
    match = re.match(
        r"^(BAFING|BAGOUE|BAGOUÉ|BERE|BÉRÉ|BOUNKANI|FOLON|"
        r"GONTOUGO|HAMBOL|KABADOUGOU|PORO|TCHOLOGO|WORODOUGOU)\b",
        text,
        flags=re.IGNORECASE,
    )
    return match.group(1).upper() if match else None


def _extract_fields(designation):
    """Extrait les champs fixes présents dans la première cellule."""
    text = clean(designation)

    labels = {
        "Intitulé": r"Intitulé\s*:",
        "Montant": r"Montant\s+marché\s*:",
        "Entreprise": r"Titulaire\s*:",
        "OS": r"O\.?\s*S\.?\s*:",
        "Date_Debut": r"Date\s+de\s+démarrage\s*:",
        "Delai": r"Délai\s*:",
        "Fin": r"Fin\s+prévue\s*:",
        "Bureau": r"Bureau\s+de\s+contrôle\s*:",
    }

    pattern = "(" + "|".join(labels.values()) + ")"
    parts = re.split(pattern, text, flags=re.IGNORECASE)

    result = {}
    reverse = {v.lower(): k for k, v in labels.items()}

    for i in range(1, len(parts) - 1, 2):
        label = parts[i].lower()
        key = reverse.get(label)
        if key:
            result[key] = parts[i + 1].strip()

    return result


def _extract_rate(text):
    """Extrait le taux d'avancement physique."""
    match = re.search(
        r"Taux\s+d['’]avancement\s*:\s*([\d\s]+(?:[,.]\d+)?)\s*%?",
        clean(text),
        flags=re.IGNORECASE,
    )
    return _number(match.group(1)) if match else None


def _extract_invoice(text):
    """Extrait le taux de facturation indiqué dans la matrice."""
    match = re.search(
        r"Total\s+factur[ée]\s*:?\s*([\d\s]+(?:[,.]\d+)?)\s*%?",
        clean(text),
        flags=re.IGNORECASE,
    )
    return _number(match.group(1)) if match else None


def _extract_visit_date(text):
    """Extrait la date de visite lorsqu'elle est mentionnée."""
    match = re.search(
        r"Visite\s+effectu[ée]e?\s+(?:le\s+)?(\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{4})",
        clean(text),
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    return _date_from_text(match.group(1).replace(" ", ""))


def _extract_technical_observations(text):
    """Retire les indicateurs de taux/facturation et conserve les observations."""
    text = clean(text)

    # Retirer les informations calculées qui ne sont pas des observations.
    text = re.sub(
        r"Taux\s+d['’]avancement\s*:.*?(?=Délai\s+consommé|Total\s+factur[ée]|$)",
        " ",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"Délai\s+consommé\s*:.*?(?=Total\s+factur[ée]|$)",
        " ",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"Total\s+factur[ée]\s*:.*$",
        " ",
        text,
        flags=re.IGNORECASE,
    )

    return clean(text)


def _extract_comments(text):
    """Sépare les commentaires généraux, difficultés et recommandations."""
    text = clean(text)

    # Certaines matrices répètent deux fois les blocs Difficultés /
    # Recommandations. On conserve une seule occurrence de chaque bloc.
    personnel_match = re.search(
        r"(Personnel\s+de\s+la\s+mission\s+de\s+contrôle\s+mobilisé\.?)",
        text,
        flags=re.IGNORECASE,
    )
    personnel = personnel_match.group(1).strip() if personnel_match else ""

    difficulties = re.findall(
        r"Difficultés\s*:?\s*(.*?)(?=Recommandations\s*:|$)",
        text,
        flags=re.IGNORECASE,
    )
    recommendations = re.findall(
        r"Recommandations\s*:?\s*(.*?)(?=Difficultés\s*:|$)",
        text,
        flags=re.IGNORECASE,
    )

    def unique_join(values):
        unique = []
        for value in values:
            value = clean(value).strip(" ;.")
            if value and value not in unique:
                unique.append(value)
        return " / ".join(unique)

    return {
        "Commentaires": personnel or None,
        "Difficultes": unique_join(difficulties) or None,
        "Recommandations": unique_join(recommendations) or None,
    }


# ---------------------------------------------------------------------
# Lecture des matrices
# ---------------------------------------------------------------------

def parse_suivis(source, categorie=None) -> list[dict]:
    """Lit une matrice Markdown et renvoie les suivis structurés.

    `source` peut être un chemin ou un couple (catégorie, chemin).
    Les lignes complémentaires des marchés ruraux sont rattachées
    à l'activité précédente.
    """
    if isinstance(source, tuple) and len(source) == 2:
        categorie, source = source

    path = Path(source)
    if not path.exists():
        raise FileNotFoundError(f"Matrice source introuvable : {path}")

    parsed = []
    current = None

    for line_number, line in enumerate(
        path.read_text(encoding="utf-8").splitlines(), start=1
    ):
        cells = _split_markdown_row(line)

        if len(cells) < 3:
            continue

        # Ignorer les lignes d'en-tête et de séparation.
        if "Désignation" in cells[0] or set(cells[0]) <= {"-", ":"}:
            continue

        if cells[0]:
            # Une nouvelle activité commence.
            fields = _extract_fields(cells[0])
            region = _extract_region(cells[0])

            technical_text = cells[1]
            comments_text = cells[2]

            current = {
                "Categorie": categorie,
                "Region_Texte": region,
                "Designation": fields.get("Intitulé") or None,
                "Montant_Texte": fields.get("Montant"),
                "Entreprise": clean(fields.get("Entreprise")) or None,
                "Date_OS": _date_from_text(fields.get("OS", "")),
                "Date_Debut": _date_from_text(fields.get("Date_Debut", "")),
                "Delai_Mois": _number(fields.get("Delai", "")),
                "Bureau_Controle": clean(fields.get("Bureau")) or None,
                "Date_Actualisation": _extract_visit_date(technical_text),
                "Observations_Techniques": _extract_technical_observations(
                    technical_text
                ),
                "Taux_Execution_Physique": _extract_rate(technical_text),
                "Total_Facture": _extract_invoice(technical_text),
                **_extract_comments(comments_text),
                "_ligne_source": line_number,
            }

            amount = _number(fields.get("Montant", ""))
            current["Montant_Marche"] = amount
            parsed.append(current)

        elif current is not None:
            # Ligne de continuation, notamment pour les matrices MR :
            # le taux et le délai consommé sont souvent sur des lignes séparées.
            continuation = " ".join(cells[1:])
            rate = _extract_rate(continuation)
            invoice = _extract_invoice(continuation)

            if rate is not None:
                current["Taux_Execution_Physique"] = rate
            if invoice is not None:
                current["Total_Facture"] = invoice

    return parsed


# ---------------------------------------------------------------------
# Import dans la base
# ---------------------------------------------------------------------

def _activity_key(activity):
    """Clé de rapprochement entre une activité et sa ligne source."""
    return (
        activity.get("Categorie"),
        activity.get("Region"),
        clean(activity.get("Departement")).upper(),
        clean(activity.get("Entreprise")).upper(),
        activity.get("Montant_Marche"),
    )


def _source_region_code(ref, region_text):
    if not region_text:
        return None
    try:
        return ref.cfg.region_code(region_text)
    except ValueError:
        return None


def _find_activity(ref, item, used_codes):
    """Rapproche une ligne source avec une activité du référentiel."""
    region_code = _source_region_code(ref, item.get("Region_Texte"))

    candidates = [
        activity
        for activity in ref.activities()
        if activity["Categorie"] == item["Categorie"]
        and activity["Region"] == region_code
        and activity["CODE_ACTIVITE"] not in used_codes
    ]

    def same(a, b):
        return clean(a).casefold() == clean(b).casefold()

    # Premier critère : entreprise + département + montant.
    for activity in candidates:
        if (
            same(activity.get("Entreprise"), item.get("Entreprise"))
            and same(activity.get("Departement"), item.get("Departement"))
            and activity.get("Montant_Marche") == item.get("Montant_Marche")
        ):
            return activity

    # Deuxième critère : désignation + entreprise.
    for activity in candidates:
        if (
            same(activity.get("Entreprise"), item.get("Entreprise"))
            and same(activity.get("Designation"), item.get("Designation"))
        ):
            return activity

    # Dernier recours : ordre des activités dans le référentiel et dans
    # la matrice, uniquement si une seule candidate reste.
    if len(candidates) == 1:
        return candidates[0]

    return None


def import_baseline(
    ref: Referentiel,
    sources: list,
    date_etat: date,
) -> list[str]:
    """Importe les suivis de référence des matrices BS et MR.

    L'import est idempotent : une activité ayant déjà un suivi
    BASELINE_WORD n'est pas réimportée.
    """
    warns = []
    used_codes = set()

    if isinstance(date_etat, datetime):
        date_etat = date_etat.date()

    for categorie, source_path in sources:
        parsed_rows = parse_suivis(source_path, categorie)

        for item in parsed_rows:
            activity = _find_activity(ref, item, used_codes)

            if activity is None:
                warns.append(
                    "Activité non rapprochée dans la matrice "
                    f"{categorie}, ligne {item['_ligne_source']} : "
                    f"{item.get('Designation') or 'désignation absente'}"
                )
                continue

            code = activity["CODE_ACTIVITE"]
            used_codes.add(code)

            # Un seul suivi initial par activité.
            existing = [
                row
                for row in ref.base.rows("Suivi_Activites")
                if row.get("CODE_ACTIVITE") == code
                and str(row.get("Source") or "").startswith("BASELINE")
            ]
            if existing:
                continue

            visit_date = item.get("Date_Actualisation") or date_etat

            observations = item.get("Observations_Techniques")
            comments = item.get("Commentaires")
            difficulties = item.get("Difficultes")
            recommendations = item.get("Recommandations")

            # Détecter les répétitions de blocs dans le texte d'origine.
            if difficulties and recommendations:
                if difficulties == recommendations:
                    warns.append(
                        f"{code} : bloc Difficultés / Recommandations répété "
                        "dans la matrice ; une seule occurrence a été conservée."
                    )

            ref.base.append(
                "Suivi_Activites",
                {
                    "ID_SOUMISSION_KOBO": f"BASELINE_{code}",
                    "CODE_ACTIVITE": code,
                    "Date_Actualisation": visit_date,
                    "Observations_Techniques": observations or None,
                    "Taux_Execution_Physique": item.get(
                        "Taux_Execution_Physique"
                    ),
                    "Commentaires": comments,
                    "Difficultes": difficulties,
                    "Recommandations": recommendations,
                    "Date_Importation": date.today(),
                    "Statut_Traitement": "VALIDE",
                    "Motif_Rejet": None,
                    "Nom_Agent": None,
                    "Total_Facture": (
                        item.get("Total_Facture")
                        if item.get("Total_Facture") is not None
                        else 0
                    ),
                    "Source": "BASELINE_WORD",
                },
            )

    return warns


# ---------------------------------------------------------------------
# Import depuis une matrice Excel
# ---------------------------------------------------------------------

def import_baseline_from_excel(
    ref: Referentiel,
    excel_path: Path,
    date_etat: date,
) -> list[str]:
    """Importe les données de suivi depuis la matrice Excel.

    Cette fonction conserve l'approche initiale du projet.
    Elle n'écrase jamais un suivi déjà importé.
    """
    import pandas as pd

    warns = []
    excel_path = Path(excel_path)

    if not excel_path.exists():
        raise FileNotFoundError(f"Fichier Excel introuvable : {excel_path}")

    df = pd.read_excel(
        excel_path,
        sheet_name="MATRICE DE SUIVI",
        skiprows=2,
    )
    df.columns = [str(column).strip() for column in df.columns]

    for column in (
        "REGIONS",
        "DEPARTEMENT",
        "NATURE DU MARCHE",
        "ENTREPRISE",
    ):
        if column in df.columns:
            df[column] = df[column].ffill()

    activities = ref.activities()
    existing_codes = {
        row.get("CODE_ACTIVITE")
        for row in ref.base.rows("Suivi_Activites")
        if row.get("Source") == "BASELINE_EXCEL"
    }

    added = 0

    for _, row in df.iterrows():
        region_value = row.get("REGIONS")
        if pd.isna(region_value):
            continue

        nature_raw = str(row.get("NATURE DU MARCHE", "")).lower()
        categorie = "BS" if "bloc" in nature_raw else "MR"

        region_raw = str(region_value).strip()
        dept_raw = (
            str(row.get("DEPARTEMENT")).strip()
            if pd.notna(row.get("DEPARTEMENT"))
            else ""
        )
        entreprise_raw = (
            str(row.get("ENTREPRISE")).strip()
            if pd.notna(row.get("ENTREPRISE"))
            else ""
        )

        matches = [
            activity
            for activity in activities
            if activity["Categorie"] == categorie
            and clean(activity["Region"]).upper()
            == clean(region_raw).upper()
            and clean(activity["Departement"]).upper()
            == clean(dept_raw).upper()
            and clean(activity["Entreprise"]).upper()
            == clean(entreprise_raw).upper()
        ]

        if len(matches) != 1:
            warns.append(
                "Correspondance Excel ambiguë ou absente : "
                f"{categorie} / {region_raw} / {dept_raw} / {entreprise_raw}"
            )
            continue

        code = matches[0]["CODE_ACTIVITE"]

        if code in existing_codes:
            continue

        taux = row.get("TAUX D'EXECUTION  (%)")
        taux_financier = row.get("TAUX D'EXECUTION FINANCIERE")
        observations = row.get("OBSERVATIONS TECHNIQUES")

        ref.base.append(
            "Suivi_Activites",
            {
                "ID_SOUMISSION_KOBO": f"BASELINE_EXCEL_{code}",
                "CODE_ACTIVITE": code,
                "Date_Actualisation": date_etat,
                "Observations_Techniques": (
                    str(observations).strip()
                    if pd.notna(observations)
                    else None
                ),
                "Taux_Execution_Physique": (
                    _number(taux) if pd.notna(taux) else None
                ),
                "Commentaires": None,
                "Difficultes": None,
                "Recommandations": None,
                "Date_Importation": date.today(),
                "Statut_Traitement": "VALIDE",
                "Total_Facture": (
                    _number(taux_financier)
                    if pd.notna(taux_financier)
                    else None
                ),
                "Source": "BASELINE_EXCEL",
            },
        )

        existing_codes.add(code)
        added += 1

    print(
        f"-> {added} lignes de suivi de référence Excel ajoutées."
    )
    return warns
