"""Reprise UNIQUE des valeurs saisies à la main dans le modèle de tableau de bord (feuille « MATRICE DE SUIVI », lignes des blocs sanitaires).

Avant, ces colonnes (montant décaissé, remise de site, statut, PGES, ONG, réception provisoire) étaient conservées « par position » dans le modèle.
Elles sont désormais des informations de l'administrateur stockées dans la base, par code d'activité. Cette reprise les y copie :

  * rapprochement de chaque ligne du modèle avec une activité de la base par MONTANT du marché (puis nom d'entreprise le plus proche) ;
  * seules les valeurs ABSENTES de la base sont renseignées : une saisie de l'administrateur n'est jamais écrasée ;
  * statut : seul « Suspendu » est repris (les autres statuts se déduisent du taux et le statut du modèle est périmé) ;
  * PGES ramené à l'une des 4 valeurs ; les dates en texte libre (« 08 et 09/09/2026 ») ne sont pas reprises (signalées) ;
  * chaque reprise est historisée (utilisateur « Reprise du modèle »).
"""
from __future__ import annotations

import difflib
from datetime import date, datetime

from openpyxl import load_workbook

from .excel_gen import SHEET, upper_noaccent
from .referentiel import PGES_VALEURS, STATUTS_CHANTIER, Referentiel, to_int
from .word_gen import ordered_activities

USER = "Reprise du modèle"
# (mot-clé, catégorie) dans l'ordre des rangs du modèle : le plus grand rang qui correspond l'emporte
_PGES = [("valid", "Validé"), ("transmis", "En cours de validation"), ("en cours", "En cours de validation"),
         ("non transmis", "Non transmis"), ("npn transmis", "Non transmis"), ("pas de pges", "Sans PGES"), ("sans pges", "Sans PGES")]
_STATUT = [("effectif", "En cours"), ("en cours", "En cours"), ("non d", "Non démarré"), ("non encore", "Non démarré"),
           ("pas encore", "Non démarré"), ("non effectif", "Non démarré"), ("achev", "Achevé"), ("termin", "Achevé"), ("suspend", "Suspendu")]


def _categorie(texte, table):
    t = str(texte or "").lower()
    trouve = [cat for kw, cat in table if kw in t]
    return trouve[-1] if trouve else None


def _montant(v):
    m = _montants(v)
    return m[0] if m else None


def _montants(v):
    """Tous les montants lisibles d'une cellule (ex. « Lot 2 : 101 033 723 » ou « 185 115 390 F CFA »)."""
    import re
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return [int(v)]
    nombres = [re.sub(r"\D", "", x) for x in re.findall(r"\d[\d\s\u00a0]{4,}\d", str(v or ""))]
    return [int(n) for n in nombres if n]


def _date(v):
    return v.date() if isinstance(v, datetime) else v if isinstance(v, date) else None


def migrer_valeurs_modele(ref: Referentiel, template, user: str = USER, dry_run: bool = False) -> dict:
    ws = load_workbook(template, data_only=True)[SHEET]
    candidats = [a for cat in ref.cfg.categories for a in ordered_activities(ref, cat)]
    pris, rapport = set(), {"reprises": [], "non_rapprochees": [], "ignorees": []}
    for r in range(4, ws.max_row + 1):
        ent, montants = ws.cell(r, 5).value, _montants(ws.cell(r, 6).value)
        if ent is None and not montants:
            continue
        libres = [a for a in candidats if a["CODE_ACTIVITE"] not in pris]
        sim = lambda x: difflib.SequenceMatcher(None, upper_noaccent(ent), upper_noaccent(x["Entreprise"])).ratio()   # noqa: E731
        pool = [a for a in libres if a["Montant_Marche"] in montants]
        a = (max(pool, key=sim) if len(pool) > 1 else pool[0]) if pool else None
        if a is None:                                           # repli : une seule entreprise très proche parmi les activités libres
            proches = [x for x in libres if sim(x) >= 0.8]
            if len(proches) == 1:
                a = proches[0]
                rapport["ignorees"].append(f"{a['CODE_ACTIVITE']} : rapprochée par le NOM de l'entreprise (montant du modèle différent) : à vérifier")
        if a is None:
            rapport["non_rapprochees"].append(f"ligne {r} ({ent!s}, {ws.cell(r, 6).value!s})".replace("\n", " "))
            continue
        pris.add(a["CODE_ACTIVITE"])
        propose = {}
        dec = _montant(ws.cell(r, 7).value)
        if dec is not None:
            propose["Montant_Decaisse"] = dec
        remise = ws.cell(r, 14).value
        if _date(remise):
            propose["Date_Remise_Site"] = _date(remise)
        elif remise not in (None, ""):
            rapport["ignorees"].append(f"{a['CODE_ACTIVITE']} : remise de site « {remise} » (texte libre, non reprise)")
        if _categorie(ws.cell(r, 18).value, _STATUT) == "Suspendu":
            propose["Statut_Chantier"] = "Suspendu"
        pges = _categorie(ws.cell(r, 20).value, _PGES)
        if pges:
            propose["PGES"] = pges
        ong = str(ws.cell(r, 21).value or "").strip()
        if ong:
            propose["ONG_Mobilisees"] = " ".join(ong.split())
        if _date(ws.cell(r, 22).value):
            propose["Date_Reception_Provisoire"] = _date(ws.cell(r, 22).value)
        changes = {k: v for k, v in propose.items() if a[k] in (None, "")}          # jamais d'écrasement
        if changes and not dry_run:
            ref.update_activity(a["CODE_ACTIVITE"], user, changes)
        if changes:
            rapport["reprises"].append((a["CODE_ACTIVITE"], changes))
    return rapport


def reprise_deja_faite(ref: Referentiel) -> bool:
    return any(h["Utilisateur"] == USER for h in ref.history())


def reprise_automatique(ref: Referentiel, template) -> list[str]:
    """Au premier traitement après la mise en service des champs administrateur : reprend UNE FOIS les valeurs du modèle.
    Le marqueur est l'historique (utilisateur « Reprise du modèle ») : une fois la reprise faite, elle ne se rejoue jamais
    (un champ effacé volontairement par l'administrateur n'est donc pas re-rempli). Retourne des lignes d'information."""
    if reprise_deja_faite(ref):
        return []
    rap = migrer_valeurs_modele(ref, template)
    if not (rap["reprises"] or rap["non_rapprochees"] or rap["ignorees"]):
        return []
    infos = [f"Reprise unique des valeurs du modèle (montant décaissé, PGES, ONG…) : {len(rap['reprises'])} activité(s) renseignée(s) dans la base."]
    infos += [f"  à vérifier : {x}" for x in rap["ignorees"]] + [f"  non rapprochée : {x}" for x in rap["non_rapprochees"]]
    return infos
