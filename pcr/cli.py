"""Commandes de l'administrateur :  python -m pcr <commande>"""
import argparse
import os
import sys
from datetime import date
from pathlib import Path

from .admin_file import apply_admin_file, export_admin_file
from .config import load_config
from .referentiel import Referentiel
from .kobo_form import KoboClient, KoboError, activities_csv_rows, build_xlsform, sync_to_kobo, write_activities_csv
from .baseline import import_baseline, import_baseline_from_excel
from .ingest import ingest, set_statut
from .seed import seed, write_report
from .excel_gen import generate_dashboard
from .pipeline import load_json, run as run_pipeline
from .word_gen import generate_all
from .store import Base

def _load_env(path: Path) -> None:
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            if line.strip() and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="pcr", description="Référentiel des activités PCR-CI")
    ap.add_argument("--config", help="fichier de configuration (défaut : config/config.yaml)")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("charger-initial", help="chargement initial depuis les matrices Word (une seule fois)")
    s.add_argument("--bs", required=True), s.add_argument("--mr", required=True)
    s.add_argument("--excel", help="ancienne matrice Excel (lecture seule, pour les localités des BS)")
    s.add_argument("--rapport", default="docs/rapport_chargement_initial.md")
    cr = sub.add_parser("charger-reference", help="importe l'état de référence (baseline) depuis Excel dans une base existante (idempotent, sans doublon)")
    cr.add_argument("--excel", required=True, help="chemin vers la matrice Excel de référence (ex: MATRICE_DE_SUIVI_DES_ACTIVITES_du_28-10-26.xlsx)")
    sub.add_parser("exporter", help="(ré)écrit le fichier d'administration depuis la base")
    sub.add_parser("appliquer", help="applique le fichier d'administration à la base (sauvegarde avant écriture)")
    h = sub.add_parser("historique", help="historique des modifications d'une activité")
    h.add_argument("code")
    sub.add_parser("verifier", help="contrôles d'intégrité (unicité des codes, séquences, compteurs)")
    sub.add_parser("lister", help="liste des activités")
    r = sub.add_parser("recuperer", help="récupère les soumissions Kobo et les intègre (doublons ignorés, invalides conservées)")
    r.add_argument("--from-json", help="fichier JSON de soumissions (test hors ligne / démonstration)")
    r.add_argument("--env", default=".env")
    sv = sub.add_parser("suivis", help="liste des suivis reçus")
    sv.add_argument("--statut", choices=["VALIDE", "A_VERIFIER", "REJETE"])
    for name, hp in (("valider", "valide une soumission « à vérifier »"), ("rejeter", "rejette une soumission")):
        c = sub.add_parser(name, help=hp)
        c.add_argument("id_kobo")
        if name == "rejeter":
            c.add_argument("motif")
    w = sub.add_parser("generer-word", help="génère les matrices Word (une par catégorie) depuis la base")
    w.add_argument("--au", help="date de situation JJ/MM/AAAA (défaut : aujourd'hui)")
    vw = sub.add_parser("verifier-word", help="génère les Word et les compare aux matrices d'origine (rapport de fidélité)")
    vw.add_argument("--bs", required=True), vw.add_argument("--mr", required=True)
    vw.add_argument("--au", required=True, help="date de situation JJ/MM/AAAA")
    vw.add_argument("--rapport", default="docs/fidelite_word.md")
    d = sub.add_parser("generer-dashboard", help="alimente le tableau de bord Excel existant depuis la base")
    d.add_argument("--au", help="date pour le nom du fichier JJ/MM/AAAA (défaut : aujourd'hui)")
    l = sub.add_parser("lancer", help="cycle complet : Kobo -> base -> Word + Excel -> e-mail -> journal")
    l.add_argument("--from-json"), l.add_argument("--au"), l.add_argument("--mail", action="store_true", help="envoyer les documents par e-mail")
    l.add_argument("--declencheur", default="manuel"), l.add_argument("--env", default=".env")
    dm = sub.add_parser("demo", help="démonstration sur une COPIE de la base (la base réelle n'est pas modifiée)")
    dm.add_argument("--json", default="demo/soumissions_demo.json"), dm.add_argument("--sortie", default="demo_sortie")
    sub.add_parser("kobo-formulaire", help="génère le questionnaire XLSForm à importer dans Kobo")
    sub.add_parser("kobo-csv", help="génère la liste des activités (activites.csv) sans l'envoyer")
    k = sub.add_parser("kobo-sync", help="envoie la liste des activités à Kobo et redéploie le formulaire")
    k.add_argument("--sans-redeploiement", action="store_true")
    k.add_argument("--env", default=".env")
    a = ap.parse_args(argv)

    cfg = load_config(a.config)
    base = Base(cfg.path("base"))
    ref = Referentiel(base, cfg)
    backups = cfg.path("sauvegardes")

    if a.cmd == "charger-initial":
        rows = seed(ref, [("BS", Path(a.bs)), ("MR", Path(a.mr))], a.excel)
        problems = ref.verify()
        warns = import_baseline(ref, [("BS", Path(a.bs)), ("MR", Path(a.mr))], cfg.date_etat_initial)
        base.save(backups)
        export_admin_file(ref, cfg.path("fichier_admin"))
        write_report(rows, a.rapport, problems)
        for w in warns:
            print("  ! " + w)
        print(f"{len(rows)} activités créées. Rapport : {a.rapport}. Fichier d'administration : {cfg.path('fichier_admin')}")
    elif a.cmd == "charger-reference":
        excel_path = Path(a.excel)
        if not excel_path.exists():
            print(f"Erreur : le fichier {excel_path} n'existe pas.")
            return 1
        warns = import_baseline_from_excel(ref, excel_path, cfg.date_etat_initial)
        base.save(backups)
        for w in warns:
            print("  ! " + w)
        print("Référence chargée avec succès. Les suivis manquants ont été ajoutés sans créer de doublons.")
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
            export_admin_file(ref, cfg.path("fichier_admin"))
    elif a.cmd == "historique":
        for r in ref.history(a.code.upper()):
            print(f"{r['Date_Modification']:%d/%m/%Y %H:%M}  {r['Utilisateur']}  {r['Champ_Modifie']} : {r['Ancienne_Valeur']!r} -> {r['Nouvelle_Valeur']!r}")
    elif a.cmd == "verifier":
        problems = ref.verify()
        print("Référentiel conforme." if not problems else "\n".join(problems))
        return 1 if problems else 0
    elif a.cmd == "recuperer":
        if a.from_json:
            import json
            raw = json.loads(Path(a.from_json).read_text(encoding="utf-8"))
            subs = raw.get("results", raw) if isinstance(raw, dict) else raw
        else:
            _load_env(Path(a.env))
            for var in ("KOBO_ASSET_UID", "KOBO_TOKEN"):
                if not os.environ.get(var):
                    print(f"Variable manquante : {var} (voir .env.example)")
                    return 2
            try:
                subs = KoboClient(os.environ.get("KOBO_URL", "https://kf.kobotoolbox.org"), os.environ["KOBO_ASSET_UID"], os.environ["KOBO_TOKEN"]).submissions()
            except KoboError as e:
                print(f"ÉCHEC : {e}")
                return 1
        rep = ingest(ref, subs, declencheur="manuel")
        base.save(backups)
        print("\n".join(rep.lines()))
    elif a.cmd == "suivis":
        for r in ref.base.rows("Suivi_Activites"):
            if not a.statut or r["Statut_Traitement"] == a.statut:
                print(f"{str(r['ID_SOUMISSION_KOBO']):<18} {r['CODE_ACTIVITE'] or '?':<9} {r['Date_Actualisation'] or '—'}  {r['Taux_Execution_Physique']!s:>5} %  "
                      f"{r['Statut_Traitement']:<10} {r['Motif_Rejet'] or ''}")
    elif a.cmd in ("valider", "rejeter"):
        set_statut(ref, a.id_kobo, "VALIDE" if a.cmd == "valider" else "REJETE", getattr(a, "motif", None))
        base.save(backups)
        print("Enregistré.")
    elif a.cmd == "generer-word":
        from .referentiel import to_date
        for f in generate_all(ref, to_date(a.au) if a.au else date.today(), cfg.path("sorties")):
            print(f)
    elif a.cmd == "verifier-word":
        from .referentiel import to_date
        from .word_check import compare
        files = generate_all(ref, to_date(a.au), cfg.path("sorties"))
        gen = {cat: f for cat, f in zip(cfg.categories, files)}
        report = compare(ref, {"BS": Path(a.bs), "MR": Path(a.mr)}, gen)
        Path(a.rapport).parent.mkdir(parents=True, exist_ok=True)
        Path(a.rapport).write_text(report, encoding="utf-8")
        print(report.split("## Différences")[0].split("## Blocs")[1] if "## Blocs" in report else "", f"\nRapport : {a.rapport}")
    elif a.cmd == "generer-dashboard":
        from .referentiel import to_date
        jour = to_date(a.au) if a.au else date.today()
        res = generate_dashboard(ref, cfg.path("dashboard_modele"), cfg.path("sorties") / f"MATRICE_DE_SUIVI_DES_ACTIVITES_du_{jour:%d-%m-%y}.xlsx")
        print(f"{res['fichier']}  ({res['lignes']} lignes, taux moyen {res['taux_moyen']} %)")
        if res["formules_ignorees"]:
            print("Cellules à formule non modifiées :", ", ".join(res["formules_ignorees"]))
    elif a.cmd == "lancer":
        from .referentiel import to_date
        _load_env(Path(a.env))
        client = None
        if not a.from_json:
            if not (os.environ.get("KOBO_ASSET_UID") and os.environ.get("KOBO_TOKEN")):
                print("Variables manquantes : KOBO_ASSET_UID et KOBO_TOKEN (voir .env.example)")
                return 2
            client = KoboClient(os.environ.get("KOBO_URL", "https://kf.kobotoolbox.org"), os.environ["KOBO_ASSET_UID"], os.environ["KOBO_TOKEN"])
        res = run_pipeline(ref, load_json(a.from_json) if a.from_json else None, to_date(a.au) if a.au else date.today(), a.mail, a.declencheur, client)
        print("\n".join(res.lines()))
        return 0 if res.ok else 1
    elif a.cmd == "demo":
        import shutil, tempfile
        tmp = Path(tempfile.mkdtemp())
        for d in ("config", "modeles"):
            shutil.copytree(cfg.root / d, tmp / d)
        (tmp / "data").mkdir()
        shutil.copy(cfg.path("base"), tmp / "data" / Path(cfg.data["chemins"]["base"]).name)
        dcfg = load_config(tmp / "config" / "config.yaml")
        dref = Referentiel(Base(dcfg.path("base")), dcfg)
        res = run_pipeline(dref, load_json(a.json), date.today(), False, "demo")
        print("\n".join(res.lines()))
        dest = Path(a.sortie)
        dest.mkdir(parents=True, exist_ok=True)
        for f in res.documents:
            shutil.copy(f, dest / Path(f).name)
        print(f"\nDocuments de démonstration : {dest.resolve()}")
        return 0 if res.ok else 1
    elif a.cmd == "kobo-formulaire":
        print(build_xlsform(cfg.path("kobo_formulaire")))
    elif a.cmd == "kobo-csv":
        print(f"{write_activities_csv(ref, cfg.path('kobo_csv'))} activités -> {cfg.path('kobo_csv')}")
    elif a.cmd == "kobo-sync":
        _load_env(Path(a.env))
        for var in ("KOBO_ASSET_UID", "KOBO_TOKEN"):
            if not os.environ.get(var):
                print(f"Variable manquante : {var} (voir .env.example)")
                return 2
        write_activities_csv(ref, cfg.path("kobo_csv"))
        try:
            client = KoboClient(os.environ.get("KOBO_URL", "https://kf.kobotoolbox.org"), os.environ["KOBO_ASSET_UID"], os.environ["KOBO_TOKEN"])
            res = sync_to_kobo(client, activities_csv_rows(ref), redeploy=not a.sans_redeploiement)
        except KoboError as e:
            print(f"ÉCHEC : {e}")
            return 1
        print(f"Liste envoyée : {res['activites']} activités ({res['fichier']}), redéploiement : {'oui' if res['redeploye'] else 'non'}.")
    elif a.cmd == "lister":
        for x in ref.activities():
            print(f"{x['CODE_ACTIVITE']}  {x['Statut']:<8} {x['Departement'] or '—':<22} {x['Entreprise'] or '—'}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
