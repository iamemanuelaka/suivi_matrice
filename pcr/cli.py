"""Commandes de l'administrateur :  python -m pcr <commande>"""
import argparse
import sys
from pathlib import Path

from .admin_file import apply_admin_file, export_admin_file
from .config import load_config
from .referentiel import Referentiel
from .seed import seed, write_report
from .store import Base


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="pcr", description="Référentiel des activités PCR-CI")
    ap.add_argument("--config", help="fichier de configuration (défaut : config/config.yaml)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("charger-initial", help="chargement initial depuis les matrices Word (une seule fois)")
    s.add_argument("--bs", required=True), s.add_argument("--mr", required=True)
    s.add_argument("--excel", help="ancienne matrice Excel (lecture seule, pour les localités des BS)")
    s.add_argument("--rapport", default="docs/rapport_chargement_initial.md")
    sub.add_parser("exporter", help="(ré)écrit le fichier d'administration depuis la base")
    sub.add_parser("appliquer", help="applique le fichier d'administration à la base (sauvegarde avant écriture)")
    h = sub.add_parser("historique", help="historique des modifications d'une activité"), 
    sub.choices["historique"].add_argument("code")
    sub.add_parser("verifier", help="contrôles d'intégrité (unicité des codes, séquences, compteurs)")
    sub.add_parser("lister", help="liste des activités")
    a = ap.parse_args(argv)

    cfg = load_config(a.config)
    base = Base(cfg.path("base"))
    ref = Referentiel(base, cfg)
    backups = cfg.path("sauvegardes")

    if a.cmd == "charger-initial":
        rows = seed(ref, [("BS", Path(a.bs)), ("MR", Path(a.mr))], a.excel)
        problems = ref.verify()
        base.save(backups)
        export_admin_file(ref, cfg.path("fichier_admin"))
        write_report(rows, a.rapport, problems)
        print(f"{len(rows)} activités créées. Rapport : {a.rapport}. Fichier d'administration : {cfg.path('fichier_admin')}")
    elif a.cmd == "exporter":
        print(export_admin_file(ref, cfg.path("fichier_admin")))
    elif a.cmd == "appliquer":
        rep = apply_admin_file(ref, cfg.path("fichier_admin"))
        print("\n".join(rep.lines()))
        problems = ref.verify()
        if problems:
            print("Contrôle d'intégrité en échec, rien n'est enregistré :", *problems, sep="\n  ")
            return 1
        if rep.created or rep.updated:
            base.save(backups)
            export_admin_file(ref, cfg.path("fichier_admin"))   # pour voir les codes générés
    elif a.cmd == "historique":
        for r in ref.history(a.code.upper()):
            print(f"{r['Date_Modification']:%d/%m/%Y %H:%M}  {r['Utilisateur']}  {r['Champ_Modifie']} : {r['Ancienne_Valeur']!r} -> {r['Nouvelle_Valeur']!r}")
    elif a.cmd == "verifier":
        problems = ref.verify()
        print("Référentiel conforme." if not problems else "\n".join(problems))
        return 1 if problems else 0
    elif a.cmd == "lister":
        for x in ref.activities():
            print(f"{x['CODE_ACTIVITE']}  {x['Statut']:<8} {x['Departement'] or '—':<22} {x['Entreprise'] or '—'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
