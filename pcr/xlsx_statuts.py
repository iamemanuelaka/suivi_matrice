"""Limite les statuts du tableau de bord aux 4 statuts retenus : Non démarré, En cours, Achevé, Suspendu.

Le modèle d'origine contenait trois statuts supplémentaires dans ses réglages (feuille PARAMETRES, table « Lecture du statut des travaux ») :
« Démarrage autorisé », « OS à reprendre » et « Abandonné » ; la feuille CALCULS testait aussi « Abandonné » dans la colonne « actif ».
La fonction retire ces réglages sans toucher au reste (rangs et ordre de priorité des 4 statuts conservés, filtre « Statut » inchangé).
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path

from .xlsx_capacity import Q, _fichier_feuille
from .xlsx_patch import Sheet
from .xlsx_repair import lire_parts

STATUTS = ("Non démarré", "En cours", "Achevé", "Suspendu")
# (mot-clé contenu dans le texte du statut, catégorie) ; l'ordre = le rang (le plus grand rang qui correspond l'emporte)
MOTS_CLES = [("effectif", "En cours"), ("en cours", "En cours"),
             ("non d", "Non démarré"), ("non encore", "Non démarré"), ("pas encore", "Non démarré"), ("non effectif", "Non démarré"),
             ("achev", "Achevé"), ("termin", "Achevé"), ("suspend", "Suspendu")]
PREMIERE, DERNIERE = 13, 41                      # plage kwStatut / catStatut / rkStatut de PARAMETRES


def limiter_statuts(src, dst) -> dict:
    parts = lire_parts(src)
    f_par = _fichier_feuille(parts, "PARAMETRES")
    sh = Sheet(parts[f_par])
    avant = [(sh.get(r, 1), sh.get(r, 2)) for r in range(PREMIERE, DERNIERE + 1) if sh.get(r, 2)]
    retires = sorted({cat for _, cat in avant if cat not in STATUTS})
    for i, r in enumerate(range(PREMIERE, DERNIERE + 1)):
        mot, cat = MOTS_CLES[i] if i < len(MOTS_CLES) else (None, None)
        sh.set(r, 1, mot)
        sh.set(r, 2, cat)
    parts[f_par] = sh.tobytes()
    f_calc = _fichier_feuille(parts, "CALCULS")
    calc = Sheet(parts[f_calc])
    n = 0
    for f in calc.root.iter(Q("f")):
        if f.text and "Abandonné" in f.text:
            nouveau = re.sub(r'OR\((\$N\d+)="Suspendu",\1="Abandonné"\)', r'\1="Suspendu"', f.text)
            if nouveau != f.text:
                f.text, n = nouveau, n + 1
    calc.dirty = True
    parts[f_calc] = calc.tobytes()
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i.filename, parts[i.filename])
    return {"statuts_retires": retires, "formules_actif_modifiees": n}
