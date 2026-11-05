"""Phase 6 : alimentation du tableau de bord Excel EXISTANT à partir de la base.

Le tableau de bord lit la feuille « MATRICE DE SUIVI » par formules (capacité : 45 lignes, lignes 4 à 48). On n'écrit que les cellules d'ENTRÉE de
cette feuille (identité de l'activité, début, délai, taux, observations, statut) et la feuille cachée META (date d'actualisation, historique du
taux moyen). Aucune formule, aucun graphique, aucune image, aucun style du classeur n'est modifié.

Non collectés par le nouveau questionnaire (spécification) : montant décaissé, taux financier, date de remise de site, PGES, ONG, réception
provisoire, photo, répondant. Ces colonnes restent TELLES QUELLES dans le modèle pour les blocs sanitaires et vides pour les marchés ruraux.
"""
import re
import unicodedata
from datetime import date, datetime
from pathlib import Path

from .referentiel import Referentiel
from .word_gen import latest_valid, ordered_activities
from .xlsx_patch import Package

SHEET, META = "MATRICE DE SUIVI", "META"
FIRST_ROW, LAST_ROW = 4, 48
COL = {"lot": 2, "region": 3, "dep": 4, "ent": 5, "montant": 6, "nature": 9, "loc": 10, "debut": 11, "delai": 12, "fin": 13, "taux": 17, "statut": 18, "obs": 19}
STYLE_ROW = 29                      # dernière ligne « décorée » du modèle : modèle de style des lignes ajoutées
META_MAX_ROWS = 450


def upper_noaccent(s) -> str:
    s = unicodedata.normalize("NFD", str(s or ""))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").upper().strip()


def statut_chantier(taux) -> str:
    """Statut dérivé du seul taux physique de la base (le nouveau questionnaire ne collecte pas de statut saisi)."""
    if taux is None or taux <= 0:
        return "Non démarré"
    return "Achevé" if taux >= 100 else "En cours"


def _region_by_row(sh, first, last) -> dict:
    """Région de chaque ligne du modèle (les cellules fusionnées ne portent la valeur que sur leur première ligne)."""
    out, cur = {}, ""
    for r in range(first, last + 1):
        cur = upper_noaccent(sh.get(r, COL["region"])) or cur
        out[r] = cur
    return out


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
    acts = [a for cat in ref.cfg.categories for a in ordered_activities(ref, cat)]
    if len(acts) > LAST_ROW - FIRST_ROW + 1:
        raise ValueError(f"{len(acts)} activités : le tableau de bord n'a de la place que pour {LAST_ROW - FIRST_ROW + 1} lignes (4 à 48). "
                         "Étendre d'abord les plages du classeur.")
    pkg = Package(template)
    sh = pkg.sheet(SHEET)
    template_regions = _region_by_row(sh, FIRST_ROW, STYLE_ROW)
    last = latest_valid(ref)
    skipped, taux_values = [], []

    for n, a in enumerate(acts, 1):
        r = 3 + n
        region = upper_noaccent(ref.cfg.regions[a["Region"]])
        decorated = r <= STYLE_ROW
        if decorated and template_regions[r] != region:
            raise ValueError(
                f"Ligne {r} du modèle : région « {template_regions[r]} » (cellules fusionnées) mais l'activité {a['CODE_ACTIVITE']} est en « {region} ». "
                "Le nombre d'activités par région a changé : réorganiser les cellules fusionnées de la colonne REGIONS dans le modèle.")
        s = last.get(a["CODE_ACTIVITE"])
        taux = s["Taux_Execution_Physique"] if s else None

        def put(key, value):
            if not sh.set(r, COL[key], value):
                skipped.append(f"{key}{r}")

        put("lot", n)
        if not decorated:
            put("region", ref.cfg.regions[a["Region"]].upper().replace("É", "E"))
        put("dep", upper_noaccent(a["Departement"]) or None)
        put("ent", a["Entreprise"])
        put("montant", a["Montant_Marche"])
        put("nature", ref.cfg.categories[a["Categorie"]])
        put("loc", "\n".join(x.strip() for x in a["Localite"].split(",")) if a["Localite"] else None)
        put("debut", a["Date_Debut"])
        put("delai", a["Delai_Mois"])
        if not decorated:
            put("fin", a["Date_Fin_Previsionnelle"])            # lignes ajoutées : fin selon la convention de la catégorie (la formule EDATE n'y existe pas)
        put("taux", taux)
        put("statut", statut_chantier(taux))
        put("obs", "\n".join(l.strip() for l in (s["Observations_Techniques"] or "").splitlines() if l.strip()) if s else None)
        if not decorated:
            for c in range(2, 25):                              # lignes ajoutées : même mise en forme que les lignes existantes
                src = sh.style(STYLE_ROW, c)
                if src is not None:
                    sh._cell(r, c, create=True).set("s", src)
            sh.dirty = True
            sh.set_row_height(r, 60)
        if isinstance(taux, (int, float)):
            taux_values.append(taux)

    for r in range(3 + len(acts) + 1, LAST_ROW + 1):            # lignes en trop (activité retirée) : vidées
        for key in ("lot", "dep", "ent", "montant", "nature", "loc", "debut", "delai", "taux", "statut", "obs"):
            sh.set(r, COL[key], None)
    mean = round(sum(taux_values) / len(taux_values), 2) if taux_values else 0
    update_meta(pkg, now, mean)
    pkg.force_full_calc()
    pkg.save(out_path)
    return {"lignes": len(acts), "taux_moyen": mean, "formules_ignorees": skipped, "fichier": str(out_path)}
