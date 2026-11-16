"""Ajustements ciblés du modèle de tableau de bord (formules et libellés), sans le reconstruire.

* Carte « Activités » : « sélection / total ». Sans filtre : 42 / 42 ; avec le filtre statut « En cours » : 27 / 42 (par exemple).
  Avant : « actives dans la sélection / activités de la sélection » (41 / 42 sans filtre, car « Suspendu » n'était pas compté actif).
"""
from __future__ import annotations

import zipfile

from lxml import etree

from .xlsx_capacity import Q, _fichier_feuille
from .xlsx_repair import lire_parts

ANCIEN_LIBELLE = "Activités en cours"
NOUVEAU_LIBELLE = "Activités sélectionnées / total"
NOUVELLE_FORMULE = 'COUNTIFS(cFiltre,1)&" / "&COUNT(cLot)'


def ajuster_kpi_activites(src, dst) -> dict:
    parts = lire_parts(src)
    f1 = _fichier_feuille(parts, "TABLEAU DE BORD")
    root = etree.fromstring(parts[f1])
    cell = next(c for c in root.iter(Q("c")) if c.get("r") == "G9")
    f = cell.find(Q("f"))
    avant = f.text
    f.text = NOUVELLE_FORMULE
    for v in cell.findall(Q("v")):
        cell.remove(v)
    n = 0
    lab = next(c for c in root.iter(Q("c")) if c.get("r") == "G8")
    for t in lab.iter(Q("t")):                       # le modèle stocke ses textes en « inline strings » dans la feuille
        if t.text == ANCIEN_LIBELLE:
            t.text, n = NOUVEAU_LIBELLE, n + 1
    parts[f1] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(src) as z:
        infos = z.infolist()
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for i in infos:
            z.writestr(i.filename, parts[i.filename])
    return {"formule_avant": avant, "formule_apres": NOUVELLE_FORMULE, "libelles_modifies": n}
