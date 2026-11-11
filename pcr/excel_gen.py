"""Phase 6 : alimentation du tableau de bord Excel EXISTANT à partir de la base.

Le tableau de bord lit la feuille technique masquée « MATRICE DE SUIVI » par formules. Sa capacité (500 activités par défaut, lignes 4 à 503)
est lue dans le modèle (`capacite_modele`) : elle n'est plus codée en dur. On n'écrit que les cellules d'ENTRÉE de cette feuille (identité de
l'activité, région sur CHAQUE ligne, début, délai, taux, observations, statut) et la feuille cachée META (date d'actualisation, historique du
taux moyen). Aucune formule, aucun graphique, aucune image du classeur n'est modifié.

Le classeur produit est diagnostiqué avant d'être publié : s'il présente un défaut qu'Excel signalerait à l'ouverture, la génération est refusée
et aucun fichier n'est écrit.

Non collectés par le nouveau questionnaire (spécification) : montant décaissé, taux financier, date de remise de site, PGES, ONG, réception
provisoire, photo, répondant. Ces colonnes restent TELLES QUELLES dans le modèle.
"""
import os
import unicodedata
from datetime import datetime
from pathlib import Path

from . import xlsx_repair as X
from .referentiel import Referentiel
from .word_gen import latest_valid, ordered_activities
from .xlsx_capacity import capacite_modele
from .xlsx_patch import Package

SHEET, META = "MATRICE DE SUIVI", "META"
FIRST_ROW = 4
COL = {"lot": 2, "region": 3, "dep": 4, "ent": 5, "montant": 6, "nature": 9, "loc": 10, "debut": 11, "delai": 12, "fin": 13, "taux": 17, "statut": 18, "obs": 19}
META_MAX_ROWS = 450


def upper_noaccent(s) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").upper().strip()


def statut_chantier(taux) -> str:
    """Statut dérivé du seul taux physique de la base (le nouveau questionnaire ne collecte pas de statut saisi)."""
    if taux is None or taux <= 0:
        return "Non démarré"
    return "Achevé" if taux >= 100 else "En cours"


def update_meta(pkg: Package, now: datetime, mean: float) -> None:
    if not pkg.has_sheet(META):
        return
    m = pkg.sheet(META)
    m.set(1, 2, now)
    r, last = 4, None
    while m.get(r, 1) is not None or m.get(r, 2) is not None:
        last = m.get(r, 2)
        r += 1
    if r > META_MAX_ROWS:               # le tableau de bord lit META!B4:B500 : on garde les 300 dernières mesures
        keep = [(m.get(i, 1), m.get(i, 2)) for i in range(r - 300, r)]
        for i in range(4, r):
            m.set(i, 1, None), m.set(i, 2, None)
        for i, (a, b) in enumerate(keep):
            m.set(4 + i, 1, a, style_from=(4, 1)), m.set(4 + i, 2, b, style_from=(4, 2))
        r = 4 + len(keep)
    if last is None or abs(float(last) - mean) >= 0.05:
        m.set(r, 1, now, style_from=(4, 1))
        m.set(r, 2, mean, style_from=(4, 2))


def generate_dashboard(ref: Referentiel, template, out_path, now: datetime | None = None) -> dict:
    now = (now or datetime.now()).replace(microsecond=0)
    out_path = Path(out_path)
    acts = [a for cat in ref.cfg.categories for a in ordered_activities(ref, cat)]
    capacite = capacite_modele(X.lire_parts(template))
    last_row = FIRST_ROW + capacite - 1
    if len(acts) > capacite:
        raise ValueError(f"{len(acts)} activités : le tableau de bord n'a de la place que pour {capacite} lignes ({FIRST_ROW} à {last_row}). "
                         "Étendre d'abord le modèle (pcr.xlsx_capacity.etendre_modele).")
    pkg = Package(template)
    sh = pkg.sheet(SHEET)
    last = latest_valid(ref)
    skipped, taux_values = [], []

    for n, a in enumerate(acts, 1):
        r = FIRST_ROW - 1 + n
        s = last.get(a["CODE_ACTIVITE"])
        taux = s["Taux_Execution_Physique"] if s else None

        def put(key, value):
            if not sh.set(r, COL[key], value):
                skipped.append(f"{key}{r}")

        put("lot", n)
        put("region", upper_noaccent(ref.cfg.regions[a["Region"]]))      # région sur CHAQUE ligne (plus de cellules fusionnées)
        put("dep", upper_noaccent(a["Departement"]) or None)
        put("ent", a["Entreprise"])
        put("montant", a["Montant_Marche"])
        put("nature", ref.cfg.categories[a["Categorie"]])
        put("loc", "\n".join(x.strip() for x in a["Localite"].split(",")) if a["Localite"] else None)
        put("debut", a["Date_Debut"])
        put("delai", a["Delai_Mois"])
        if not sh.has_formula(r, COL["fin"]):                           # fin selon la convention de la catégorie, sauf si la formule EDATE existe
            put("fin", a["Date_Fin_Previsionnelle"])
        put("taux", taux)
        put("statut", statut_chantier(taux))
        put("obs", "\n".join(l.strip() for l in (s["Observations_Techniques"] or "").splitlines() if l.strip()) if s else None)
        if isinstance(taux, (int, float)):
            taux_values.append(taux)

    for r in range(FIRST_ROW + len(acts), last_row + 1):               # lignes en trop (activité retirée) : vidées
        for key in ("lot", "region", "dep", "ent", "montant", "nature", "loc", "debut", "delai", "taux", "statut", "obs"):
            sh.set(r, COL[key], None)
        if not sh.has_formula(r, COL["fin"]):
            sh.set(r, COL["fin"], None)

    mean = round(sum(taux_values) / len(taux_values), 2) if taux_values else 0
    update_meta(pkg, now, mean)
    pkg.force_full_calc()

    # Écriture dans un fichier temporaire, diagnostic, puis publication : jamais de classeur défectueux en sortie.
    tmp = out_path.with_name(out_path.name + ".tmp")
    try:
        pkg.save(tmp)
        problemes = X.diagnostiquer(X.lire_parts(tmp))
        if problemes:
            raise ValueError("Classeur refusé, défauts détectés (Excel afficherait un message de réparation) : " + " ; ".join(problemes))
        os.replace(tmp, out_path)
    finally:
        if tmp.exists():
            tmp.unlink()
    return {"lignes": len(acts), "taux_moyen": mean, "formules_ignorees": skipped, "fichier": str(out_path),
            "corrections_modele": list(pkg.corrections)}
