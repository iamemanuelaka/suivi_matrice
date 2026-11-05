"""Phase 4 : intégration des soumissions Kobo dans Suivi_Activites.

Garanties : une soumission (identifiant Kobo) n'est jamais intégrée deux fois ; toute soumission reçue est CONSERVÉE
(valide, à vérifier ou rejetée avec son motif) ; une nouvelle soumission ajoute une ligne, n'écrase jamais l'existant.
Statuts : VALIDE (utilisée dans les documents), A_VERIFIER (conservée, non utilisée tant que l'administrateur ne la valide pas),
REJETE (conservée avec motif, jamais utilisée).
"""
import re
from dataclasses import dataclass, field
from datetime import date, datetime

from .referentiel import Referentiel, to_date

DATE_MIN = date(2020, 1, 1)


def flatten(sub: dict) -> dict:
    """Retire les préfixes de groupes Kobo (« avancement/taux_execution » -> « taux_execution »)."""
    flat = {k.split("/")[-1]: v for k, v in sub.items() if not k.startswith("_")}
    for k in ("_id", "_submission_time", "_validation_status"):
        flat[k] = sub.get(k)
    return flat


def _txt(v):
    s = re.sub(r"[ \t]+", " ", str(v).replace("\r", "")).strip() if v not in (None, "") else ""
    return s or None


def is_rejected_in_kobo(sub: dict) -> bool:
    vs = sub.get("_validation_status") or {}
    return (vs.get("uid") if isinstance(vs, dict) else vs) == "validation_status_not_approved"


@dataclass
class Report:
    recuperees: int = 0
    nouvelles: list = field(default_factory=list)      # [(id, code, statut, motif)]
    ignorees: int = 0                                  # déjà intégrées auparavant

    def count(self, statut): return sum(1 for _, _, s, _ in self.nouvelles if s == statut)

    def lines(self) -> list[str]:
        out = [f"{self.recuperees} soumission(s) récupérée(s) : {len(self.nouvelles)} nouvelle(s) "
               f"({self.count('VALIDE')} valide(s), {self.count('A_VERIFIER')} à vérifier, {self.count('REJETE')} rejetée(s)), "
               f"{self.ignorees} déjà traitée(s) et ignorée(s)."]
        out += [f"  {s:<10} {i} {c or '?'} : {m}" for i, c, s, m in self.nouvelles if s != "VALIDE"]
        return out


def evaluate(sub: dict, acts: dict, valid_rows: list[dict], today: date) -> tuple[dict, str, str | None]:
    """Retourne (champs, statut, motif). Les champs sont toujours conservés, même rejetés."""
    f = flatten(sub)
    code = (_txt(f.get("code_activite")) or "").upper() or None
    rec = {"CODE_ACTIVITE": code, "Observations_Techniques": _txt(f.get("observations_techniques")), "Commentaires": _txt(f.get("commentaires")),
           "Difficultes": _txt(f.get("difficultes")), "Recommandations": _txt(f.get("recommandations")), "Nom_Agent": _txt(f.get("nom_agent")),
           "Date_Actualisation": None, "Taux_Execution_Physique": None}
    problems = []
    if is_rejected_in_kobo(sub):
        problems.append("soumission refusée dans Kobo")
    if not code or code not in acts:
        problems.append(f"code d'activité inconnu : {code!r}")
    try:
        rec["Date_Actualisation"] = to_date(f.get("date_actualisation"))
        if rec["Date_Actualisation"] > today:
            problems.append("date d'actualisation dans le futur")
        elif rec["Date_Actualisation"] < DATE_MIN:
            problems.append("date d'actualisation avant 2020")
    except (ValueError, TypeError):
        problems.append(f"date d'actualisation invalide : {f.get('date_actualisation')!r}")
    try:
        t = float(str(f.get("taux_execution")).replace(",", "."))
        if not 0 <= t <= 100:
            raise ValueError
        rec["Taux_Execution_Physique"] = int(t) if t.is_integer() else t
    except (ValueError, TypeError):
        problems.append(f"taux d'exécution invalide : {f.get('taux_execution')!r} (attendu 0 à 100)")
    if not rec["Observations_Techniques"]:
        problems.append("observations techniques vides")
    if problems:
        return rec, "REJETE", " ; ".join(problems)
    doubts = []
    if acts[code]["Statut"] != "Actif":
        doubts.append("activité archivée")
    prev = [r for r in valid_rows if r["CODE_ACTIVITE"] == code and r["Taux_Execution_Physique"] is not None
            and (r["Date_Actualisation"] or date.min) <= rec["Date_Actualisation"]]
    if prev:
        last = max(prev, key=lambda r: (r["Date_Actualisation"] or date.min, str(r["ID_SOUMISSION_KOBO"])))
        if rec["Taux_Execution_Physique"] < last["Taux_Execution_Physique"]:
            doubts.append(f"taux en baisse : {rec['Taux_Execution_Physique']} % après {last['Taux_Execution_Physique']} % "
                          f"({last['Date_Actualisation']:%d/%m/%Y})")
    return rec, ("A_VERIFIER" if doubts else "VALIDE"), (" ; ".join(doubts) or None)


def ingest(ref: Referentiel, submissions: list[dict], today: date | None = None, declencheur: str = "manuel", log: bool = True) -> Report:
    """Intègre les soumissions (en mémoire : l'appelant sauvegarde la base). Idempotent."""
    today, now = today or date.today(), datetime.now().replace(microsecond=0)
    acts = {a["CODE_ACTIVITE"]: a for a in ref.activities()}
    known = {str(r["ID_SOUMISSION_KOBO"]) for r in ref.base.rows("Suivi_Activites")}
    rep = Report(recuperees=len(submissions))
    for sub in sorted(submissions, key=lambda s: (str(s.get("_submission_time") or ""), int(s.get("_id") or 0))):
        sid = sub.get("_id")
        if sid is None:
            rep.nouvelles.append(("?", None, "REJETE", "soumission sans identifiant Kobo (_id)"))
            continue
        if str(sid) in known:
            rep.ignorees += 1
            continue
        valid = [r for r in ref.base.rows("Suivi_Activites") if r["Statut_Traitement"] == "VALIDE"]
        rec, statut, motif = evaluate(sub, acts, valid, today)
        last_tf = next((r["Total_Facture"] for r in sorted(valid, key=lambda r: (r["Date_Actualisation"] or date.min), reverse=True)
                        if r["CODE_ACTIVITE"] == rec["CODE_ACTIVITE"] and r["Total_Facture"] is not None), None)
        ref.base.append("Suivi_Activites", {"ID_SOUMISSION_KOBO": sid, **rec, "Date_Importation": now, "Statut_Traitement": statut,
                                            "Motif_Rejet": motif, "Total_Facture": last_tf, "Source": "KOBO"})
        known.add(str(sid))
        rep.nouvelles.append((sid, rec["CODE_ACTIVITE"], statut, motif))
    if log:
        log_execution(ref, rep, declencheur, now)
    return rep


def log_execution(ref: Referentiel, rep: Report, declencheur: str, now: datetime, documents: str | None = None, email: str | None = None,
                  resultat: str = "OK", detail: str | None = None) -> None:
    rows = ref.base.rows("Journal_Executions")
    ref.base.append("Journal_Executions", {
        "ID_Execution": max([r["ID_Execution"] or 0 for r in rows] + [0]) + 1, "Date_Execution": now, "Declencheur": declencheur,
        "Recuperees": rep.recuperees, "Traitees": len(rep.nouvelles), "Ignorees": rep.ignorees, "Echouees": rep.count("REJETE"),
        "Documents": documents, "Email": email, "Resultat": resultat, "Detail": detail or "; ".join(rep.lines()[1:])[:500] or None})


def set_statut(ref: Referentiel, sid: str, statut: str, motif: str | None = None) -> None:
    """Décision de l'administrateur sur une soumission (valider une soumission « à vérifier », ou rejeter)."""
    for r in ref.base.rows("Suivi_Activites"):
        if str(r["ID_SOUMISSION_KOBO"]) == str(sid):
            ref.base.update("Suivi_Activites", r["_r"], {"Statut_Traitement": statut, "Motif_Rejet": motif if statut == "REJETE" else r["Motif_Rejet"]})
            return
    raise KeyError(f"Soumission introuvable : {sid}")
