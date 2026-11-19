"""Opérations du logiciel : une fonction par bouton de l'interface.

Toutes les opérations rechargent la base depuis le disque, et celles qui la modifient :
  1. récupèrent d'abord la dernière version sur GitHub (si configuré) ;
  2. sauvegardent la base (copie horodatée dans data/sauvegardes) puis l'enregistrent ;
  3. renvoient la base sur GitHub.
Un verrou empêche deux opérations simultanées (double clic, etc.).
"""
from __future__ import annotations

import os
import subprocess
import sys
import threading
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

from pcr.calculs import indicateurs_delai
from pcr.config import load_config
from pcr.ingest import set_statut
from pcr.referentiel import EDITABLE, PGES_VALEURS, STATUTS, STATUTS_CHANTIER, Referentiel
from pcr.store import Base
from pcr.word_gen import latest_valid

from .github_sync import GitHubSync, SyncConflict, SyncError, SyncReport
from .workspace import Settings, prepare_workspace

_LOCK = threading.RLock()

LIBELLES = {
    "CODE_ACTIVITE": "Code", "Categorie": "Catégorie", "Region": "Région", "Departement": "Département", "Localite": "Localité(s)",
    "Designation": "Intitulé du marché", "Montant_Marche": "Montant du marché (FCFA)", "Entreprise": "Titulaire (entreprise)",
    "Numero_OS": "N° de l'O.S.", "Date_OS": "Date de l'O.S.", "Date_Debut": "Date de démarrage", "Delai_Mois": "Délai (mois)",
    "Bureau_Controle": "Bureau de contrôle", "Date_Fin_Previsionnelle": "Fin prévue", "Statut": "Activité (base)",
    "Montant_Decaisse": "Montant décaissé (FCFA)", "Date_Remise_Site": "Date de remise du site", "Statut_Chantier": "Statut du chantier",
    "PGES": "PGES", "ONG_Mobilisees": "ONG mobilisée(s)", "Date_Reception_Provisoire": "Date de réception provisoire",
}
CHOIX = {"Statut": list(STATUTS), "Statut_Chantier": list(STATUTS_CHANTIER), "PGES": list(PGES_VALEURS)}


class AppError(Exception):
    """Erreur à afficher telle quelle à l'utilisateur."""


@dataclass
class Resultat:
    ok: bool
    message: str
    details: list | None = None


class Services:
    def __init__(self, settings: Settings | None = None, session=None):
        self.settings = settings or Settings.load()
        self._session = session                             # tests : faux serveur GitHub
        self.activate()

    # --- espace de travail -------------------------------------------------------
    def activate(self) -> None:
        self.ws = prepare_workspace(self.settings.ws)
        os.chdir(self.ws)                                   # le moteur cherche data/photos/ depuis le dossier courant
        self.settings.apply_env()

    @property
    def cfg(self):
        cfg = load_config(self.ws / "config" / "config.yaml")
        if self.settings.administrateur:
            cfg.data["administrateur"] = self.settings.administrateur
        return cfg

    def base_existe(self) -> bool:
        return self.cfg.path("base").is_file()

    def _ref(self) -> Referentiel:
        cfg = self.cfg
        if not cfg.path("base").is_file():
            raise AppError("La base n'est pas encore présente sur cet ordinateur. Allez dans Paramètres → GitHub "
                           "puis cliquez sur « Synchroniser maintenant » pour la récupérer.")
        return Referentiel(Base(cfg.path("base")), cfg)

    def _save(self, ref: Referentiel) -> None:
        problems = ref.verify()
        if problems:
            raise AppError("Contrôle d'intégrité en échec, rien n'a été enregistré :\n" + "\n".join(problems))
        ref.base.save(self.cfg.path("sauvegardes"))

    # --- GitHub --------------------------------------------------------------------
    def _github(self) -> GitHubSync | None:
        s = self.settings
        return GitHubSync(s.github_depot, s.github_branche, s.github_jeton, self.ws, session=self._session) if s.github_ok else None

    def synchroniser(self, push: bool = True, prefer: str | None = None, message: str | None = None) -> SyncReport | None:
        gh = self._github()
        if gh is None:
            return None
        with _LOCK:
            try:
                return gh.sync(push=push and bool(self.settings.github_jeton), prefer=prefer,
                               message=message or "Mise à jour depuis l'application PCR-CI")
            except SyncConflict as e:
                raise AppError(str(e)) from e
            except SyncError as e:
                raise AppError(f"Synchronisation GitHub impossible : {e}") from e

    def derniere_synchro(self) -> str | None:
        gh = self._github()
        return gh.last_sync() if gh else None

    def tester_github(self) -> str:
        gh = self._github()
        if gh is None:
            raise AppError("Indiquez le dépôt sous la forme propriétaire/dépôt.")
        try:
            return gh.check()
        except SyncError as e:
            raise AppError(str(e)) from e

    def _modifier(self, action, message: str):
        """Cadre commun des opérations qui modifient la base."""
        with _LOCK:
            avertissement = None
            try:
                self.synchroniser(push=False)
            except AppError as e:
                if "CONFLIT" in str(e).upper() or "à la fois" in str(e):
                    raise
                avertissement = str(e)                       # hors connexion : on travaille en local, envoi plus tard
            ref = self._ref()
            out = action(ref)
            self._save(ref)
            try:
                self.synchroniser(push=True, message=message)
            except AppError as e:
                avertissement = f"Enregistré sur cet ordinateur, mais pas encore sur GitHub : {e}"
            return out, avertissement

    # --- tableau de bord -------------------------------------------------------------
    def tableau_de_bord(self) -> dict:
        from pcr.pipeline import alertes
        with _LOCK:
            ref = self._ref()
            acts = [a for a in ref.activities() if a["Statut"] == "Actif"]
            suivis = ref.base.rows("Suivi_Activites")
            last = latest_valid(ref)
            journal = ref.base.rows("Journal_Executions")
            par_cat = {ref.cfg.categories[c]: sum(1 for a in acts if a["Categorie"] == c) for c in ref.cfg.categories}
            statuts = {}
            for a in acts:
                st = a.get("Statut_Chantier") or "Non renseigné"
                statuts[st] = statuts.get(st, 0) + 1
            taux = [last[a["CODE_ACTIVITE"]]["Taux_Execution_Physique"] for a in acts
                    if a["CODE_ACTIVITE"] in last and last[a["CODE_ACTIVITE"]]["Taux_Execution_Physique"] is not None]
            sans_suivi = [a["CODE_ACTIVITE"] for a in acts if a["CODE_ACTIVITE"] not in last]
            return {
                "actives": len(acts), "par_categorie": par_cat, "statuts": statuts,
                "taux_moyen": round(sum(taux) / len(taux), 1) if taux else None,
                "a_verifier": sum(1 for r in suivis if r["Statut_Traitement"] == "A_VERIFIER"),
                "suivis_total": len(suivis), "sans_suivi": len(sans_suivi),
                "alertes": alertes(ref, date.today()),
                "journal": sorted(journal, key=lambda r: r["Date_Execution"] or datetime.min, reverse=True)[:8],
                "derniere_synchro": self.derniere_synchro(),
            }

    def tableau_de_bord_excel(self, nature: str = "", region: str = "", entreprise: str = "", statut: str = ""):
        """Indicateurs, graphique et alertes calculés comme dans le tableau de bord Excel, pour les filtres choisis."""
        from .tableau_de_bord import Filtres, calculer
        with _LOCK:
            ref = self._ref()
            tb = calculer(ref, Filtres(nature or "", region or "", entreprise or "", statut or ""))
            journal = sorted(ref.base.rows("Journal_Executions"), key=lambda r: r["Date_Execution"] or datetime.min, reverse=True)[:5]
            a_verifier = sum(1 for r in ref.base.rows("Suivi_Activites") if r["Statut_Traitement"] == "A_VERIFIER")
        return tb, journal, a_verifier, self.derniere_synchro()

    # --- activités ---------------------------------------------------------------------
    def activites(self) -> list[dict]:
        with _LOCK:
            ref = self._ref()
            last = latest_valid(ref)
            out = []
            for a in ref.activities():
                s = last.get(a["CODE_ACTIVITE"])
                i = indicateurs_delai(a["Date_Debut"], a["Delai_Mois"], date.today(), ref.cfg.conventions[a["Categorie"]])
                out.append({**{k: v for k, v in a.items() if k != "_r"},
                            "Categorie_nom": ref.cfg.categories.get(a["Categorie"], a["Categorie"]),
                            "Region_nom": ref.cfg.regions.get(a["Region"], a["Region"]),
                            "Taux": s["Taux_Execution_Physique"] if s else None,
                            "Derniere_visite": s["Date_Actualisation"] if s else None,
                            "Delai_consomme": i["consommation_pct"] if i else None})
            return out

    def historique(self, code: str) -> list[dict]:
        with _LOCK:
            return list(reversed(self._ref().history(code)))

    def creer_activite(self, categorie: str, region: str, champs: dict) -> tuple[str, str | None]:
        champs = {k: v for k, v in champs.items() if k in EDITABLE and v not in (None, "")}
        try:
            return self._modifier(lambda ref: ref.create_activity(categorie, region, **champs),
                                  f"Nouvelle activité ({categorie} {region}) depuis l'application")
        except ValueError as e:
            raise AppError(str(e)) from e

    def modifier_activite(self, code: str, champs: dict, effacer=()) -> tuple[list, str | None]:
        try:
            return self._modifier(lambda ref: ref.update_activity(code, changes=champs, clear=tuple(effacer)),
                                  f"Modification de {code} depuis l'application")
        except (ValueError, KeyError) as e:
            raise AppError(str(e).strip("'")) from e

    # --- suivis Kobo ---------------------------------------------------------------------
    def suivis(self, statut: str | None = None) -> list[dict]:
        with _LOCK:
            ref = self._ref()
            acts = {a["CODE_ACTIVITE"]: a for a in ref.activities()}
            out = []
            for r in ref.base.rows("Suivi_Activites"):
                if statut and r["Statut_Traitement"] != statut:
                    continue
                a = acts.get(r["CODE_ACTIVITE"]) or {}
                out.append({**{k: v for k, v in r.items() if k != "_r"},
                            "Departement": a.get("Departement"), "Entreprise": a.get("Entreprise"),
                            "photo_chemin": str(self.photo(r.get("Photo_du_site"))) if self.photo(r.get("Photo_du_site")) else None})
            return sorted(out, key=lambda r: (str(r["Date_Importation"] or ""), str(r["ID_SOUMISSION_KOBO"])), reverse=True)

    def photo(self, nom: str | None) -> Path | None:
        if not nom:
            return None
        nom = Path(str(nom)).name
        for c in (self.ws / "data" / "photos" / (Path(nom).stem + ".jpg"), self.ws / "data" / "photos" / nom):
            if c.is_file():
                return c
        return None

    def decider_suivi(self, id_kobo, valider: bool, motif: str | None = None) -> str | None:
        if not valider and not (motif or "").strip():
            raise AppError("Indiquez le motif du rejet.")
        try:
            _, avert = self._modifier(lambda ref: set_statut(ref, str(id_kobo), "VALIDE" if valider else "REJETE", motif),
                                      f"Suivi {id_kobo} {'validé' if valider else 'rejeté'} depuis l'application")
        except KeyError as e:
            raise AppError(str(e).strip("'")) from e
        return avert

    # --- documents / cycle complet ----------------------------------------------------------
    def lancer(self, au: date | None = None, envoyer_mail: bool = False, avec_kobo: bool = True) -> Resultat:
        from pcr.kobo_form import KoboClient
        from pcr.pipeline import run as run_pipeline
        s = self.settings
        if avec_kobo and not s.kobo_ok:
            raise AppError("Renseignez l'UID du formulaire et le jeton KoboToolbox dans Paramètres → KoboToolbox.")
        if envoyer_mail and not s.mail_ok:
            raise AppError("Renseignez le serveur d'envoi et les destinataires dans Paramètres → E-mail.")
        s.apply_env()
        client = KoboClient(s.kobo_url, s.kobo_asset_uid, s.kobo_jeton) if avec_kobo else None

        def action(ref):
            return run_pipeline(ref, None if avec_kobo else [], au or date.today(), envoyer_mail, "application", client)

        with _LOCK:
            res, avert = self._modifier_pipeline(action, f"Suivi lancé depuis l'application ({(au or date.today()):%d/%m/%Y})")
        lignes = res.lines() + ([f"⚠ {avert}"] if avert else [])
        if not res.ok:
            return Resultat(False, f"Échec : {res.erreur}", lignes)
        return Resultat(True, f"{len(res.documents)} document(s) généré(s)" + (f", {res.email}" if envoyer_mail else ""), lignes)

    def generer_tableau_de_bord(self, au: date | None = None, envoyer_mail: bool = False) -> Resultat:
        """Tableau de bord Excel généré sur cet ordinateur (depuis la base, sans Kobo)."""
        from pcr.pipeline import run_tableau_de_bord
        if envoyer_mail and not self.settings.mail_ok:
            raise AppError("Renseignez le serveur d'envoi et les destinataires dans Paramètres → E-mail.")
        self.settings.apply_env()
        with _LOCK:
            res, avert = self._modifier_pipeline(lambda ref: run_tableau_de_bord(ref, au or date.today(), envoyer_mail, "application"),
                                                 "Tableau de bord généré depuis l'application")
        lignes = [f"Tableau de bord : {', '.join(Path(d).name for d in res.documents) or 'aucun'}",
                  f"Activités en alerte : {len(res.alertes)}", f"E-mail : {res.email}"] + ([f"⚠ {avert}"] if avert else [])
        if not res.ok:
            return Resultat(False, f"Échec : {res.erreur}", lignes + [f"ÉCHEC : {res.erreur}"])
        return Resultat(True, "Tableau de bord généré" + (f", {res.email}" if envoyer_mail else ""), lignes)

    def _modifier_pipeline(self, action, message):
        """Le pipeline enregistre lui-même la base (avec sauvegarde) : on encadre seulement par la synchronisation."""
        avert = None
        try:
            self.synchroniser(push=False)
        except AppError as e:
            if "à la fois" in str(e):
                raise
            avert = str(e)
        res = action(self._ref())
        if res.ok:
            try:
                self.synchroniser(push=True, message=message)
            except AppError as e:
                avert = f"Base enregistrée sur cet ordinateur, mais pas encore sur GitHub : {e}"
        return res, avert

    # --- traitement sur GitHub (le même que celui du lundi) ------------------------------
    WORKFLOWS = {"matrices": ("suivi.yml", "matrices Word"), "tableau": ("tableau-de-bord.yml", "tableau de bord Excel")}

    @property
    def _etat_execution(self) -> Path:
        return self.ws / ".execution_github.json"

    def execution_en_cours(self) -> dict | None:
        """Exécution lancée depuis l'application et pas encore récupérée (survit à la fermeture de l'application)."""
        import json
        try:
            return json.loads(self._etat_execution.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None

    def _ecrire_etat(self, etat: dict | None) -> None:
        import json
        if etat is None:
            self._etat_execution.unlink(missing_ok=True)
        else:
            self._etat_execution.write_text(json.dumps(etat, ensure_ascii=False), encoding="utf-8")

    def lancer_sur_github(self, envoyer_mail: bool = False, quoi: str = "matrices") -> dict:
        """Envoie d'abord les modifications faites ici, puis lance un traitement GitHub Actions :
        « matrices » (Kobo, base, matrices Word, e-mail) ou « tableau » (tableau de bord Excel, e-mail)."""
        workflow = self.WORKFLOWS[quoi][0]
        gh = self._github()
        if gh is None or not self.settings.github_jeton:
            raise AppError("Renseignez le dépôt et le jeton GitHub dans Paramètres → GitHub.")
        if self.execution_en_cours():
            raise AppError("Un traitement lancé depuis l'application est déjà en cours.")
        self.synchroniser(push=True, message="Mise à jour depuis l'application avant le traitement GitHub")
        with _LOCK:
            try:
                avant_sorties = gh.remote_tree(("sorties/",))
                avant_runs = gh.run_ids(workflow)
                gh.dispatch(workflow, {"send_mail": "true" if envoyer_mail else "false"})
            except SyncError as e:
                raise AppError(str(e)) from e
        etat = {"lance_le": datetime.now().isoformat(timespec="seconds"), "mail": envoyer_mail, "run_id": None, "quoi": quoi,
                "workflow": workflow,
                "runs_avant": sorted(avant_runs), "sorties_avant": avant_sorties}
        self._ecrire_etat(etat)
        return etat

    def suivre_github(self) -> dict | None:
        """État de l'exécution en cours : {status, conclusion, url, etape} ; None si rien n'est en cours."""
        etat = self.execution_en_cours()
        gh = self._github()
        if not etat or gh is None:
            return None
        try:
            if not etat.get("run_id"):
                run = gh.new_run(etat.get("workflow", "suivi.yml"), set(etat["runs_avant"]))
                if run is None:
                    return {"status": "queued", "conclusion": None, "url": None, "etape": "Démarrage sur GitHub…"}
                etat["run_id"] = run["id"]
                self._ecrire_etat(etat)
            return gh.run_status(etat["run_id"])
        except SyncError as e:
            raise AppError(str(e)) from e

    def recuperer_github(self) -> dict:
        """Après l'exécution : récupère la base mise à jour et les documents produits dans sorties/."""
        etat = self.execution_en_cours() or {}
        gh = self._github()
        with _LOCK:
            rep = self.synchroniser(push=False)
            try:
                docs = gh.download_new("sorties/", etat.get("sorties_avant") or {}) if gh else []
            except SyncError as e:
                raise AppError(str(e)) from e
        self._ecrire_etat(None)
        return {"documents": [Path(p).name for p in docs], "synchro": rep.text() if rep else ""}

    def abandonner_suivi_github(self) -> None:
        self._ecrire_etat(None)

    def envoyer_liste_kobo(self) -> str:
        from pcr.kobo_form import KoboClient, KoboError, activities_csv_rows, sync_to_kobo, write_activities_csv
        s = self.settings
        if not s.kobo_ok:
            raise AppError("Renseignez l'UID du formulaire et le jeton KoboToolbox dans Paramètres → KoboToolbox.")
        with _LOCK:
            ref = self._ref()
            write_activities_csv(ref, self.cfg.path("kobo_csv"))
            try:
                r = sync_to_kobo(KoboClient(s.kobo_url, s.kobo_asset_uid, s.kobo_jeton), activities_csv_rows(ref), redeploy=True)
            except KoboError as e:
                raise AppError(str(e)) from e
        return f"Liste envoyée à Kobo : {r['activites']} activités, formulaire redéployé : {'oui' if r['redeploye'] else 'non'}."

    def tester_kobo(self) -> str:
        from pcr.kobo_form import KoboClient, KoboError
        s = self.settings
        if not s.kobo_ok:
            raise AppError("Renseignez l'UID du formulaire et le jeton.")
        try:
            n = len(KoboClient(s.kobo_url, s.kobo_asset_uid, s.kobo_jeton).submissions())
        except KoboError as e:
            raise AppError(str(e)) from e
        return f"Connexion réussie : {n} soumission(s) sur le formulaire."

    def tester_mail(self) -> str:
        from pcr.mailer import MailError, build_message, recipients, send
        s = self.settings
        if not s.mail_ok:
            raise AppError("Renseignez le serveur, l'utilisateur, le mot de passe et les destinataires.")
        s.apply_env()
        to = recipients(self.cfg)
        try:
            send(build_message("Test – application PCR-CI", "Ceci est un message de test envoyé depuis l'application PCR-CI.",
                               [], s.mail_expediteur or s.smtp_utilisateur, to))
        except MailError as e:
            raise AppError(str(e)) from e
        return "E-mail de test envoyé à " + ", ".join(to)

    def documents(self, n: int = 30) -> list[Path]:
        d = self.ws / "sorties"
        files = [f for f in d.glob("*") if f.suffix.lower() in (".docx", ".xlsx")] if d.is_dir() else []
        return sorted(files, key=lambda f: f.stat().st_mtime, reverse=True)[:n]

    @staticmethod
    def ouvrir(path: Path) -> None:
        """Ouvre un fichier ou un dossier avec l'application par défaut de Windows (Word, Excel, Explorateur)."""
        path = Path(path)
        if sys.platform.startswith("win"):
            os.startfile(str(path))                        # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def enregistrer_parametres(self, valeurs: dict) -> None:
        ancien_ws = self.settings.ws
        for k, v in valeurs.items():
            if hasattr(self.settings, k):
                setattr(self.settings, k, (v or "").strip() if isinstance(v, str) else v)
        self.settings.save()
        if self.settings.ws != ancien_ws:
            self.activate()
        else:
            self.settings.apply_env()
