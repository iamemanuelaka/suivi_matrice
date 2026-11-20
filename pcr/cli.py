"""Commandes de l'administrateur :  python -m pcr <commande>"""
import argparse
import os
import sys
from datetime import date
from pathlib import Path

from .admin_file import apply_admin_file, export_admin_file
from .config import load_config
from .referentiel import Referentiel
from .kobo_form import KoboClient, KoboError, StructuresError, activities_csv_rows, build_xlsform, load_structures, sync_to_kobo, types_csv_rows, write_activities_csv
from .baseline import import_baseline, import_baseline_from_excel
from .ingest import ingest, set_statut
from .seed import seed, write_report
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
    pm = sub.add_parser("pges-en-masse", help="met le même PGES à toutes les activités d'un statut (ex. En cours → Validé), avec simulation")
    pm.add_argument("--statut", default="En cours", help="statut du chantier concerné (défaut : En cours)")
    pm.add_argument("--valeur", default="Validé", help="valeur du PGES à appliquer (défaut : Validé)")
    pm.add_argument("--seulement-vides", action="store_true", help="ne touche pas aux PGES déjà renseignés")
    pm.add_argument("--a-blanc", action="store_true", help="affiche ce qui serait modifié sans rien écrire")
    sf = sub.add_parser("statuts-figer", help="prend le statut suggéré par le taux comme statut de l'administrateur, pour les activités qui n'en ont pas encore")
    sf.add_argument("--a-blanc", action="store_true", help="affiche ce qui serait enregistré sans rien écrire")
    sub.add_parser("ong-liste", help="liste les ONG saisies (nombre d'activités) et signale les doublons d'écriture")
    ro = sub.add_parser("renommer-ong", help="remplace une ONG (et ses variantes d'écriture) par le nom retenu dans toutes les activités")
    ro.add_argument("ancien"), ro.add_argument("nouveau")
    for nom in ("tableau-de-bord", "generer-dashboard"):
        d = sub.add_parser(nom, help="génère le tableau de bord Excel depuis la base (filtres, indicateurs, graphique, alertes)"
                           + ("" if nom == "tableau-de-bord" else " — ancien nom de « tableau-de-bord »"))
        d.add_argument("--au", help="date de situation JJ/MM/AAAA (défaut : aujourd'hui)")
        d.add_argument("--mail", action="store_true", help="envoyer le tableau de bord par e-mail")
        d.add_argument("--declencheur", default="manuel"), d.add_argument("--env", default=".env")
    l = sub.add_parser("lancer", help="cycle complet : Kobo -> base -> matrices Word -> e-mail -> journal")
    l.add_argument("--from-json"), l.add_argument("--au"), l.add_argument("--mail", action="store_true", help="envoyer les documents par e-mail")
    l.add_argument("--declencheur", default="manuel"), l.add_argument("--env", default=".env")
    al = sub.add_parser("alerter", help="envoie un e-mail d'alerte aux destinataires (MAIL_TO) : utilisé par le workflow en cas d'échec")
    al.add_argument("--sujet", required=True), al.add_argument("--message", required=True), al.add_argument("--env", default=".env")
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
    elif a.cmd == "pges-en-masse":
        from .actions_masse import pges_en_masse
        rap = pges_en_masse(ref, a.statut, a.valeur, a.seulement_vides, dry_run=a.a_blanc)
        print(("[À BLANC] " if a.a_blanc else "") + f"PGES « {a.valeur} » pour les chantiers « {a.statut} » : "
              f"{len(rap['modifiees'])} modifiée(s), {len(rap['deja'])} déjà à cette valeur, {len(rap['gardees'])} conservée(s).")
        for code, ancien in rap["modifiees"]:
            print(f"  ~ {code} : PGES {ancien!r} -> {a.valeur!r}")
        for code, ancien in rap["gardees"]:
            print(f"  = {code} : conservé ({ancien!r})")
        if rap["modifiees"] and not a.a_blanc:
            problems = ref.verify()
            if problems:
                print("Contrôle d'intégrité en échec, rien n'est enregistré :", *problems, sep="\n  ")
                return 1
            base.save(backups)
            export_admin_file(ref, cfg.path("fichier_admin"))
    elif a.cmd == "statuts-figer":
        from .actions_masse import figer_statuts
        faits = figer_statuts(ref, dry_run=a.a_blanc)
        print(("[À BLANC] " if a.a_blanc else "") + f"{len(faits)} activité(s) reçoivent leur statut d'administrateur.")
        for code, st in faits:
            print(f"  ~ {code} : Statut_Chantier -> {st!r}")
        if faits and not a.a_blanc:
            problems = ref.verify()
            if problems:
                print("Contrôle d'intégrité en échec, rien n'est enregistré :", *problems, sep="\n  ")
                return 1
            base.save(backups)
            export_admin_file(ref, cfg.path("fichier_admin"))
    elif a.cmd == "ong-liste":
        liste = ref.ong_liste()
        for g in liste:
            noms = " / ".join(f"{n} ({k})" for n, k in g["noms"].items())
            print(("⚠ " if g["proches"] else "  ") + f"{noms} — activités : {', '.join(g['codes'])}"
                  + (f"   ← proche de : {', '.join(g['proches'])}" if g["proches"] else ""))
        nb = sum(1 for g in liste if g["proches"])
        print(f"{len(liste)} ONG distincte(s) ; {nb} à vérifier (noms proches comptés comme des ONG différentes : homogénéiser avec renommer-ong).")
    elif a.cmd == "renommer-ong":
        codes = ref.renommer_ong(a.ancien, a.nouveau)
        print(f"{len(codes)} activité(s) mise(s) à jour : {', '.join(codes) or '—'}")
        if codes:
            problems = ref.verify()
            if problems:
                print("Contrôle d'intégrité en échec, rien n'est enregistré :", *problems, sep="\n  ")
                return 1
            base.save(backups)
            export_admin_file(ref, cfg.path("fichier_admin"))
    elif a.cmd in ("tableau-de-bord", "generer-dashboard"):
        from .pipeline import run_tableau_de_bord
        from .referentiel import to_date
        if a.mail:
            _load_env(Path(a.env))
        res = run_tableau_de_bord(ref, to_date(a.au) if a.au else date.today(), a.mail, a.declencheur)
        print(f"Tableau de bord : {', '.join(str(d) for d in res.documents) or 'aucun'}")
        print(f"Activités en alerte : {len(res.alertes)}")
        print(f"E-mail : {res.email}")
        if res.erreur:
            print(f"ÉCHEC : {res.erreur}")
        return 0 if res.ok else 1
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
    elif a.cmd == "alerter":
        from .mailer import MailError, build_message, recipients, send
        _load_env(Path(a.env))
        to = recipients(cfg)
        if not to:
            print("Aucun destinataire (MAIL_TO ou email.destinataires).")
            return 2
        try:
            send(build_message(a.sujet, a.message, [], os.environ.get("MAIL_FROM") or os.environ.get("SMTP_USER", ""), to))
        except MailError as e:
            print(f"Alerte non envoyée : {e}")
            return 1
        print("Alerte envoyée à " + ", ".join(to))
        return 0
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
        try:
            n = write_activities_csv(ref, cfg.path('kobo_csv'))
        except StructuresError as e:
            print(f"ÉCHEC : {e}")
            return 1
        print(f"{n} activités -> {cfg.path('kobo_csv')} ; structures.csv et types_activite.csv -> {cfg.path('kobo_csv').parent}")
    elif a.cmd == "kobo-sync":
        _load_env(Path(a.env))
        for var in ("KOBO_ASSET_UID", "KOBO_TOKEN"):
            if not os.environ.get(var):
                print(f"Variable manquante : {var} (voir .env.example)")
                return 2
        try:
            write_activities_csv(ref, cfg.path("kobo_csv"))
            structures = load_structures(cfg)
            client = KoboClient(os.environ.get("KOBO_URL", "https://kf.kobotoolbox.org"), os.environ["KOBO_ASSET_UID"], os.environ["KOBO_TOKEN"])
            rows = activities_csv_rows(ref, structures)
            res = sync_to_kobo(client, rows, redeploy=not a.sans_redeploiement, structures=structures,
                               types=types_csv_rows(ref, rows, structures))
        except (KoboError, StructuresError) as e:
            print(f"ÉCHEC : {e}")
            return 1
        print(f"Listes envoyées : {res['structures']} structures, {res['types']} types, {res['activites']} activités ({res['fichier']}), "
              f"redéploiement : {'oui' if res['redeploye'] else 'non'}.")
    elif a.cmd == "lister":
        for x in ref.activities():
            print(f"{x['CODE_ACTIVITE']}  {x['Statut']:<8} {x['Departement'] or '—':<22} {x['Entreprise'] or '—'}")
    return 0

if __name__ == "__main__":
    sys.exit(main())
