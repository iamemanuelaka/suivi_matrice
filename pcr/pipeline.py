"""Phase 7 : le cycle complet.  Kobo -> base -> (Word + Excel) -> sauvegarde -> e-mail -> journal.

Règles : sauvegarde de la base avant toute écriture ; aucun document n'est envoyé si une étape échoue (un e-mail d'erreur est alors envoyé
si le SMTP est configuré) ; relancer le cycle est sans danger (soumissions déjà traitées ignorées, idempotent).
"""
import json
import os
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .calculs import indicateurs_delai
from .excel_gen import generate_dashboard
from .ingest import Report, ingest, log_execution
from .kobo_form import KoboClient, KoboError
from .mailer import MailError, build_message, recipients, send
from .referentiel import Referentiel
from .word_gen import generate_all, latest_valid


@dataclass
class Result:
    ok: bool = True
    report: Report = field(default_factory=Report)
    documents: list = field(default_factory=list)
    email: str = "non envoyé"
    alertes: list = field(default_factory=list)
    erreur: str | None = None

    def lines(self) -> list[str]:
        out = self.report.lines() + [f"Documents : {', '.join(Path(d).name for d in self.documents) or 'aucun'}", f"E-mail : {self.email}"]
        out += [f"Activités en alerte : {len(self.alertes)}"] + [f"  - {a}" for a in self.alertes]
        return out + ([f"ÉCHEC : {self.erreur}"] if self.erreur else [])


def alertes(ref: Referentiel, au: date) -> list[str]:
    """Même règle que le tableau de bord : taux < seuil ET délai consommé >= seuil (jamais calculé si donnée manquante)."""
    c = ref.cfg.data.get("controles", {})
    s_taux, s_delai = c.get("seuil_alerte_taux", 30), c.get("seuil_alerte_delai", 50)
    last, out = latest_valid(ref), []
    for a in ref.activities():
        s = last.get(a["CODE_ACTIVITE"])
        if a["Statut"] != "Actif" or not s or s["Taux_Execution_Physique"] is None or s["Taux_Execution_Physique"] >= 100:
            continue
        i = indicateurs_delai(a["Date_Debut"], a["Delai_Mois"], au, ref.cfg.conventions[a["Categorie"]])
        if i and s["Taux_Execution_Physique"] < s_taux and i["consommation_pct"] >= s_delai:
            out.append((i["consommation_pct"] - s["Taux_Execution_Physique"],
                        f"{a['CODE_ACTIVITE']} – {a['Departement'] or a['Region']} ({a['Entreprise'] or '—'}) : exécution {s['Taux_Execution_Physique']} % "
                        f"pour {i['consommation_pct']:g} % du délai consommé"))
    return [t for _, t in sorted(out, reverse=True)]


def mail_body(res: Result, au: date, a_verifier: list[str]) -> str:
    L = ["Bonjour,", "", f"Veuillez trouver ci-joint les matrices de suivi (blocs sanitaires, marchés ruraux) et le tableau de bord, mis à jour au {au:%d/%m/%Y}.", "",
         f"Nouvelles actualisations reçues : {len(res.report.nouvelles)} ({res.report.count('VALIDE')} valide(s), {res.report.count('A_VERIFIER')} à vérifier, "
         f"{res.report.count('REJETE')} rejetée(s))."]
    if a_verifier:
        L += ["", "Soumissions à vérifier par l'administrateur :"] + [f"  - {x}" for x in a_verifier]
    L += ["", f"Activités en alerte (exécution < seuil et délai consommé ≥ seuil) : {len(res.alertes)}"] + [f"  - {x}" for x in res.alertes[:15]]
    return "\n".join(L + ["", "Message généré automatiquement à partir des données KoboToolbox.", "", "Cordialement"])


def run(ref: Referentiel, submissions=None, au: date | None = None, send_mail: bool = False, declencheur: str = "manuel", client: KoboClient | None = None) -> Result:
    cfg, au, res = ref.cfg, au or date.today(), Result()
    now = datetime.now().replace(microsecond=0)
    out_dir = cfg.path("sorties")
    try:
        ref.base.save(cfg.path("sauvegardes"))                                   # sauvegarde AVANT toute modification
        if submissions is None:
            if client is None:
                raise KoboError("Aucune source de soumissions (client Kobo ou fichier JSON requis).")
            submissions = client.submissions()
        res.report = ingest(ref, submissions, au, declencheur, log=False)
        res.alertes = alertes(ref, au)
        res.documents = [*generate_all(ref, au, out_dir),
                         Path(generate_dashboard(ref, cfg.path("dashboard_modele"), out_dir / f"MATRICE_DE_SUIVI_DES_ACTIVITES_du_{au:%d-%m-%y}.xlsx", now)["fichier"])]
        problems = ref.verify()
        if problems:
            raise RuntimeError("Contrôle d'intégrité du référentiel en échec : " + "; ".join(problems))
        if send_mail:
            to = recipients(cfg)
            if not to:
                raise MailError("Aucun destinataire (variable MAIL_TO ou email.destinataires dans la configuration).")
            a_verifier = [f"{i} {c} : {m}" for i, c, s, m in res.report.nouvelles if s == "A_VERIFIER"]
            send(build_message(cfg.data["email"]["objet"].format(date=f"{au:%d/%m/%Y}"), mail_body(res, au, a_verifier), res.documents,
                               os.environ.get("MAIL_FROM") or os.environ.get("SMTP_USER", ""), to))
            res.email = "envoyé à " + ", ".join(to)
        log_execution(ref, res.report, declencheur, now, documents=", ".join(Path(d).name for d in res.documents), email=res.email)
        ref.base.save()
    except Exception as e:                                                       # noqa: BLE001 : tout échec est journalisé et signalé
        res.ok, res.erreur = False, f"{type(e).__name__}: {e}"
        try:
            log_execution(ref, res.report, declencheur, now, documents=None, email="non envoyé", resultat="ÉCHEC", detail=res.erreur[:500])
            ref.base.save()
            to = recipients(cfg)
            if send_mail and to and not isinstance(e, MailError):
                send(build_message(f"[ÉCHEC] Suivi des activités PCR-CI – {au:%d/%m/%Y}", f"Le traitement automatique a échoué.\n\n{res.erreur}", [],
                                   os.environ.get("MAIL_FROM") or os.environ.get("SMTP_USER", ""), to))
        except Exception:                                                        # noqa: BLE001
            pass
    return res


def load_json(path) -> list[dict]:
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    return raw.get("results", raw) if isinstance(raw, dict) else raw
