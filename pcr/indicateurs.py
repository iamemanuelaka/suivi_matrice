"""Indicateurs du tableau de bord : source unique pour l'application et le classeur Excel généré.

Règles (reprises de l'ancien classeur, vérifiées contre lui) :
  - population : activités actives, dans l'ordre du classeur (N° 1, 2, … = « lot ») ;
  - 4 filtres liés : nature (catégorie), région, entreprise, statut du chantier ; un filtre sans activité compatible est ignoré ;
  - statut : celui saisi par l'administrateur, sinon déduit du taux (0 % non démarré, 100 % achevé, sinon en cours) ;
  - délai consommé = temps écoulé depuis le début des travaux ÷ durée prévue (fin prévue − début) × 100, non plafonné ;
    non calculable pour une activité achevée, sans début, sans fin prévue, ou aux dates incohérentes ;
  - alerte : exécution < seuil (30 %) ET délai consommé ≥ seuil (50 %), classées par écart (délai − exécution) décroissant ;
  - graphique par région (par entreprise quand une région est choisie) : Σ temps écoulé ÷ Σ durées prévues,
    exécution pondérée par les montants (moyenne simple si montants absents), 24 groupes maximum, tri par délai consommé.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .referentiel import GENERIQUES_ONG, Referentiel, cle_ong, ong_canonique, statut_deduit
from .word_gen import latest_valid, ordered_activities

TOUS = ""                                     # valeur « pas de filtre »
STATUTS = ("Non démarré", "En cours", "Achevé", "Suspendu")
MAX_GROUPES = 24
PREFIXE = "P-"


@dataclass
class Ligne:
    lot: int
    code: str
    nature: str
    region: str
    departement: str
    entreprise: str
    montant: int
    decaisse: int
    taux: float | None
    statut: str
    pges: str
    ong: str
    debut: date | None
    fin: date | None
    observations: str
    calculabilite: str = ""                   # vide = délai calculable, sinon la raison
    ecoule: int | None = None
    duree: int | None = None
    conso: float | None = None
    alerte: bool = False

    @property
    def evaluable(self) -> bool:
        return self.taux is not None and self.conso is not None

    @property
    def numero(self) -> str:
        return f"{PREFIXE}{self.lot:03d}"


@dataclass
class Filtres:
    nature: str = TOUS
    region: str = TOUS
    entreprise: str = TOUS
    statut: str = TOUS


@dataclass
class TableauDeBord:
    total: int
    selection: list
    effectifs: Filtres
    ignores: list
    options: dict
    taux_moyen: float | None
    pges_valides: int
    ong_nommees: int
    ong_sans_nom: int
    montant_total: int
    decaisse_total: int
    taux_financier: float | None
    alertes: list
    groupes: list
    par: str                                  # "région" ou "entreprise"
    exclusions: dict = field(default_factory=dict)
    seuils: tuple = (30, 50)


def alias_ong(cfg) -> dict:
    """Variantes d'écriture des ONG (config.yaml, rubrique ong_alias) : {clé variante: clé retenue}."""
    return {cle_ong(str(k)): cle_ong(str(v)) for k, v in (cfg.data.get("ong_alias") or {}).items() if k and v}


def ong_cle(nom: str, alias: dict) -> tuple[str, bool]:
    """(clé de l'ONG nommée ou "", ONG mobilisée sans nom). Même lecture que l'ancien classeur."""
    if not nom:
        return "", False
    canon = ong_canonique(nom)
    if canon == GENERIQUES_ONG[0]:
        return "", False
    if canon == GENERIQUES_ONG[1] or len(nom) < 3:
        return "", True
    k = cle_ong(nom)
    return alias.get(k, k), False


def lignes(ref: Referentiel, au: date) -> list[Ligne]:
    """Une ligne par activité active, comme la feuille MATRICE DE SUIVI du classeur."""
    last = latest_valid(ref)
    acts = [a for cat in ref.cfg.categories for a in ordered_activities(ref, cat)]
    c = ref.cfg.data.get("controles", {})
    s_taux, s_conso = c.get("seuil_alerte_taux", 30), c.get("seuil_alerte_delai", 50)
    out = []
    for n, a in enumerate(acts, 1):
        s = last.get(a["CODE_ACTIVITE"])
        taux = s["Taux_Execution_Physique"] if s else None
        if isinstance(taux, (int, float)) and 0 < taux < 1:                 # même conversion que le classeur (0,35 -> 35 %)
            taux = taux * 100
        statut = a["Statut_Chantier"] or statut_deduit(taux)
        l = Ligne(lot=n, code=a["CODE_ACTIVITE"], nature=ref.cfg.categories[a["Categorie"]],
                  region=ref.cfg.regions[a["Region"]], departement=a["Departement"] or "", entreprise=(a["Entreprise"] or "").strip(),
                  montant=max(a["Montant_Marche"] or 0, 0), decaisse=a["Montant_Decaisse"] or 0, taux=taux, statut=statut,
                  pges=a["PGES"] or "Non renseigné", ong=(a["ONG_Mobilisees"] or "").strip(), debut=a["Date_Debut"],
                  fin=a["Date_Fin_Previsionnelle"], observations=(s["Observations_Techniques"] or "") if s else "")
        if statut == "Achevé":
            l.calculabilite = "activité achevée (fin réelle non disponible)"
        elif not l.debut:
            l.calculabilite = "début des travaux manquant"
        elif not l.fin:
            l.calculabilite = "fin prévue ou délai manquant"
        elif (l.fin - l.debut).days <= 0:
            l.calculabilite = "dates incohérentes ou durée nulle"
        else:
            l.duree = (l.fin - l.debut).days
            l.ecoule = max(0, (au - l.debut).days)
            l.conso = l.ecoule / l.duree * 100
        l.alerte = l.evaluable and l.taux < s_taux and l.conso >= s_conso
        out.append(l)
    return out


def _ok(l: Ligne, f: Filtres, sauf: str | None = None) -> bool:
    return all((k == sauf or not getattr(f, k) or getattr(l, k) == getattr(f, k)) for k in ("nature", "region", "entreprise", "statut"))


def calculer(ref: Referentiel, filtres: Filtres | None = None, au: date | None = None) -> TableauDeBord:
    au, f = au or date.today(), filtres or Filtres()
    tout = lignes(ref, au)
    c = ref.cfg.data.get("controles", {})
    seuils = (c.get("seuil_alerte_taux", 30), c.get("seuil_alerte_delai", 50))

    # Filtres liés (cascade nature → région → entreprise → statut, comme CALCULS!CE5:CE8) : un choix sans activité est ignoré.
    eff, ignores = Filtres(), []
    for k, lib in (("nature", "Nature"), ("region", "Région"), ("entreprise", "Entreprise"), ("statut", "Statut")):
        v = getattr(f, k)
        if v and any(getattr(l, k) == v and _ok(l, eff) for l in tout):
            setattr(eff, k, v)
        elif v:
            ignores.append(lib)
    sel = [l for l in tout if _ok(l, eff)]

    # Listes proposées dans chaque filtre : valeurs compatibles avec les AUTRES filtres (listes liées).
    options = {}
    for k in ("nature", "region", "entreprise"):
        options[k] = sorted({getattr(l, k) for l in tout if getattr(l, k) and _ok(l, eff, sauf=k)}, key=str.casefold)
    options["statut"] = [s for s in STATUTS if any(l.statut == s and _ok(l, eff, sauf="statut") for l in tout)]

    taux = [l.taux for l in sel if l.taux is not None]
    ong_noms, sans_nom, alias = set(), 0, alias_ong(ref.cfg)
    for l in sel:
        cle, anonyme = ong_cle(l.ong, alias)
        sans_nom += anonyme
        if cle:
            ong_noms.add(cle)
    montant, decaisse = sum(l.montant for l in sel), sum(l.decaisse for l in sel)

    alertes = sorted((l for l in sel if l.alerte), key=lambda l: (-(l.conso - l.taux), l.lot))

    # Graphique : par région, ou par entreprise quand une région est choisie (CALCULS!DX:EP).
    par = "entreprise" if eff.region else "région"
    groupes: dict = {}
    for l in sel:
        cle = l.entreprise if eff.region else l.region
        if not cle:                                                         # comme l'Excel : groupe sans libellé non affiché
            continue
        g = groupes.setdefault(cle, {"libelle": cle, "n": 0, "ecoule": 0, "duree": 0, "montant": 0, "mt": 0.0, "taux": 0.0})
        if l.evaluable:
            g["n"] += 1
            g["ecoule"] += l.ecoule
            g["duree"] += l.duree
            g["montant"] += l.montant
            g["mt"] += l.montant * l.taux
            g["taux"] += l.taux
    series = []
    for g in groupes.values():
        conso = g["ecoule"] / g["duree"] * 100 if g["duree"] > 0 else None
        execu = g["mt"] / g["montant"] if g["montant"] > 0 else (g["taux"] / g["n"] if g["n"] else None)
        series.append({"libelle": g["libelle"], "n": g["n"], "conso": conso, "execution": execu,
                       "methode": "non calculable" if not g["n"] else "pondéré par les montants" if g["montant"] > 0 else "moyenne simple"})
    series.sort(key=lambda g: -(g["conso"] if g["conso"] is not None else -1))

    exclusions = {"achevées": 0, "sans date de début": 0, "sans fin prévue/délai": 0, "aux dates incohérentes": 0, "sans taux d'exécution": 0}
    for l in sel:
        r = l.calculabilite
        if r.startswith("activité achevée"):
            exclusions["achevées"] += 1
        elif r.startswith("début"):
            exclusions["sans date de début"] += 1
        elif r.startswith("fin prévue"):
            exclusions["sans fin prévue/délai"] += 1
        elif r.startswith("dates"):
            exclusions["aux dates incohérentes"] += 1
        elif l.taux is None:
            exclusions["sans taux d'exécution"] += 1

    return TableauDeBord(
        total=len(tout), selection=sel, effectifs=eff, ignores=ignores, options=options,
        taux_moyen=sum(taux) / len(taux) if taux else None,
        pges_valides=sum(1 for l in sel if l.pges == "Validé"), ong_nommees=len(ong_noms), ong_sans_nom=sans_nom,
        montant_total=montant, decaisse_total=decaisse, taux_financier=decaisse / montant * 100 if montant else None,
        alertes=alertes, groupes=series[:MAX_GROUPES], par=par, exclusions=exclusions, seuils=seuils)
