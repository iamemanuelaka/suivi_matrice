"""Import des suivis de référence depuis les matrices Markdown."""
import re
from datetime import date, datetime
from pathlib import Path

from .referentiel import Referentiel
from .seed import parse_date_loose


def clean(value) -> str:
    """Nettoie le Markdown et les espaces sans supprimer le contenu."""
    if value is None:
        return ""
    text = str(value)
    text = text.replace("\u00a0", " ").replace("\u202f", " ")
    text = text.replace("\u200b", "")
    text = re.sub(r"\*+", "", text)
    text = re.sub(r"(?<!\w)_(.*?)_(?!\w)", r"\1", text)
    return re.sub(r"\s+", " ", text).strip()


def _empty(value):
    return value is None or not str(value).strip()


def _number(value):
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
    if not text:
        return None
    result, _ = parse_date_loose(str(text))
    return result


def _split_row(line):
    if not line.lstrip().startswith("|"):
        return []
    return [clean(x) for x in line.strip().strip("|").split("|")]


def _extract_region(text):
    text = clean(text)
    match = re.match(
        r"^(BAFING|BAGOUE|BAGOUÉ|BERE|BÉRÉ|BOUNKANI|FOLON|"
        r"GONTOUGO|HAMBOL|KABADOUGOU|PORO|TCHOLOGO|WORODOUGOU)\b",
        text,
        re.I,
    )
    return match.group(1).upper() if match else None


def _extract_fields(text):
    """Extrait les champs entre deux libellés, sans confondre regex et texte."""
    text = clean(text)
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
    pattern = re.compile(
        r"(?P<label>" + "|".join(labels.values()) + r")",
        re.I,
    )
    matches = list(pattern.finditer(text))
    result = {}

    reverse = {
        "intitul": "Intitulé",
        "montant": "Montant",
        "titulaire": "Entreprise",
        "o.s.": "OS",
        "os": "OS",
        "date de démarrage": "Date_Debut",
        "délai": "Delai",
        "fin prévue": "Fin",
        "bureau de contrôle": "Bureau",
    }

    for i, match in enumerate(matches):
        label_text = re.sub(r"\s*:\s*$", "", match.group())
        key = reverse.get(label_text.lower())
        if not key:
            continue
        start = match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        result[key] = text[start:end].strip()

    return result


def _extract_rate(text):
    match = re.search(
        r"Taux\s+d['’]avancement\s*:\s*([\d\s]*(?:[,.]\d+)?)\s*%?",
        clean(text),
        re.I,
    )
    return _number(match.group(1)) if match else None


def _extract_invoice(text):
    match = re.search(
        r"Total\s+factur[ée]\s*:?\s*([\d\s]*(?:[,.]\d+)?)\s*%?",
        clean(text),
        re.I,
    )
    return _number(match.group(1)) if match else None


def _extract_visit_date(text):
    match = re.search(
        r"Visite\s+effectu[ée]e?\s+(?:le\s+)?"
        r"(\d{1,2}\s*/\s*\d{1,2}\s*/\s*\d{4})",
        clean(text),
        re.I,
    )
    return _date_from_text(match.group(1).replace(" ", "")) if match else None


def _extract_technical_observations(text):
    text = clean(text)
    patterns = [
        r"Taux\s+d['’]avancement\s*:.*?(?=Délai\s+consommé|Total\s+factur[ée]|$)",
        r"Délai\s+consommé\s*:.*?(?=Total\s+factur[ée]|$)",
        r"Total\s+factur[ée]\s*:.*$",
    ]
    for pattern in patterns:
        text = re.sub(pattern, " ", text, flags=re.I)
    return clean(text)


def _extract_comments(text):
    text = clean(text)

    personnel_match = re.search(
        r"Personnel\s+de\s+la\s+mission\s+de\s+contrôle\s+mobilisé\.?",
        text,
        re.I,
    )
    personnel = personnel_match.group(0).strip() if personnel_match else ""

    difficulties = re.findall(
        r"Difficultés\s*:?\s*(.*?)(?=Recommandations\s*:|$)",
        text,
        re.I,
    )
    recommendations = re.findall(
        r"Recommandations\s*:?\s*(.*?)(?=Difficultés\s*:|$)",
        text,
        re.I,
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


def parse_suivis(source, categorie=None):
    """Lit les suivis et rattache les lignes complémentaires MR."""
    if isinstance(source, tuple) and len(source) == 2:
        categorie, source = source

    path = Path(source)
    if not path.exists():
        raise FileNotFoundError(path)

    parsed = []
    current = None

    for line_number, line in enumerate(
        path.read_text(encoding="utf-8").splitlines(), 1
    ):
        cells = _split_row(line)
        if len(cells) < 3:
            continue

        if "Désignation" in cells[0] or set(cells[0]) <= {"-", ":"}:
            continue

        if cells[0]:
            fields = _extract_fields(cells[0])
            technical = cells[1]
            comments = cells[2]

            amount = _number(fields.get("Montant"))
            current = {
                "Categorie": categorie,
                "Region_Texte": _extract_region(cells[0]),
                "Designation": fields.get("Intitulé") or None,
                "Montant_Marche": amount,
                "Entreprise": clean(fields.get("Entreprise")) or None,
                "Date_OS": _date_from_text(fields.get("OS", "")),
                "Date_Debut": _date_from_text(fields.get("Date_Debut", "")),
                "Delai_Mois": _number(fields.get("Delai", "")),
                "Bureau_Controle": clean(fields.get("Bureau")) or None,
                "Date_Actualisation": _extract_visit_date(technical),
                "Observations_Techniques": _extract_technical_observations(
                    technical
                ),
                "Taux_Execution_Physique": _extract_rate(technical),
                "Total_Facture": _extract_invoice(technical),
                "_comments_source": comments,
                "_ligne_source": line_number,
                **_extract_comments(comments),
            }
            parsed.append(current)

        elif current is not None:
            continuation = " ".join(cells[1:])
            rate = _extract_rate(continuation)
            invoice = _extract_invoice(continuation)

            if rate is not None:
                current["Taux_Execution_Physique"] = rate
            if invoice is not None:
                current["Total_Facture"] = invoice

    return parsed


def _source_region_code(ref, region_text):
    if not region_text:
        return None
    try:
        return ref.cfg.region_code(region_text)
    except ValueError:
        return None


def _find_activity(ref, item, used_codes):
    """Rapproche par catégorie, région et ordre de création dans le référentiel."""
    region = _source_region_code(ref, item.get("Region_Texte"))
    candidates = [
        a for a in ref.activities()
        if a["Categorie"] == item["Categorie"]
        and a["Region"] == region
        and a["CODE_ACTIVITE"] not in used_codes
    ]
    candidates.sort(key=lambda a: a["CODE_ACTIVITE"])
    return candidates[0] if candidates else None


def import_baseline(ref: Referentiel, sources: list, date_etat: date):
    warns = []
    used_codes = set()

    if isinstance(date_etat, datetime):
        date_etat = date_etat.date()

    for categorie, source_path in sources:
        for item in parse_suivis(source_path, categorie):
            activity = _find_activity(ref, item, used_codes)

            if activity is None:
                warns.append(
                    f"Activité non rapprochée : {categorie}, "
                    f"ligne {item['_ligne_source']}"
                )
                continue

            code = activity["CODE_ACTIVITE"]
            used_codes.add(code)

            existing = [
                row for row in ref.base.rows("Suivi_Activites")
                if row.get("CODE_ACTIVITE") == code
                and str(row.get("Source") or "").startswith("BASELINE")
            ]
            if existing:
                continue

            comments_source = item.get("_comments_source") or ""
            for label in ("Difficultés", "Recommandations"):
                count = len(re.findall(
                    rf"{label}\s*:", comments_source, re.I
                ))
                if count > 1:
                    warns.append(
                        f"{code} : le bloc {label} se répète dans la matrice."
                    )

            ref.base.append(
                "Suivi_Activites",
                {
                    "ID_SOUMISSION_KOBO": f"BASELINE_{code}",
                    "CODE_ACTIVITE": code,
                    "Date_Actualisation": (
                        item.get("Date_Actualisation") or date_etat
                    ),
                    "Observations_Techniques": (
                        item.get("Observations_Techniques") or None
                    ),
                    "Taux_Execution_Physique": (
                        item.get("Taux_Execution_Physique")
                    ),
                    "Commentaires": item.get("Commentaires"),
                    "Difficultes": item.get("Difficultes"),
                    "Recommandations": item.get("Recommandations"),
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
