"""Interface du logiciel PCR-CI (NiceGUI) : une fenêtre, un menu à gauche, un écran par tâche.

Chaque bouton appelle une fonction de services.py dans un fil séparé (run.io_bound) : la fenêtre ne se fige jamais.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

from nicegui import run, ui

from .services import CHOIX, LIBELLES, AppError, Services

S: Services | None = None            # un seul utilisateur, un seul poste : services partagés par les écrans

BLEU = "#1F4E78"
MENU = [("/", "Tableau de bord", "dashboard"), ("/suivi", "Lancer le suivi", "play_circle"),
        ("/suivis", "Suivis à valider", "fact_check"), ("/activites", "Activités", "construction"),
        ("/parametres", "Paramètres", "settings")]
FIXES = ["Departement", "Localite", "Designation", "Montant_Marche", "Entreprise", "Numero_OS", "Date_OS", "Date_Debut",
         "Delai_Mois", "Bureau_Controle"]
CHANTIER = ["Statut_Chantier", "PGES", "ONG_Mobilisees", "Montant_Decaisse", "Date_Remise_Site", "Date_Reception_Provisoire"]
DATES = {"Date_OS", "Date_Debut", "Date_Remise_Site", "Date_Reception_Provisoire"}
MONTANTS = {"Montant_Marche", "Montant_Decaisse"}


# --- mise en forme ---------------------------------------------------------------------------
def d(v) -> str:
    if isinstance(v, datetime):
        return f"{v:%d/%m/%Y %H:%M}"
    if isinstance(v, date):
        return f"{v:%d/%m/%Y}"
    return "" if v is None else str(v)


def money(v) -> str:
    return "" if v in (None, "") else f"{int(v):,}".replace(",", " ")


def pct(v) -> str:
    return "" if v is None else f"{v:g} %"


def erreur(e: Exception) -> None:
    ui.notify(str(e), type="negative", multi_line=True, close_button="OK", timeout=0)


def avertir(msg: str | None) -> None:
    if msg:
        ui.notify(msg, type="warning", multi_line=True, close_button="OK", timeout=0)


async def appel(fn, *args, **kw):
    """Exécute une opération du moteur sans figer la fenêtre ; les erreurs prévues sont affichées, pas levées."""
    try:
        return True, await run.io_bound(fn, *args, **kw)
    except AppError as e:
        erreur(e)
    except Exception as e:                                                     # noqa: BLE001
        erreur(Exception(f"Erreur inattendue : {type(e).__name__} : {e}"))
    return False, None


@contextmanager
def cadre(chemin: str, titre: str):
    ui.colors(primary=BLEU)
    ui.page_title(f"PCR-CI – {titre}")
    with ui.header().classes("items-center px-4 py-2").style(f"background:{BLEU}"):
        ui.icon("account_tree", size="md")
        ui.label("Suivi des activités PCR-CI").classes("text-lg font-semibold")
        ui.space()
        ui.label(titre).classes("text-base opacity-90")
    with ui.left_drawer(value=True, fixed=True).classes("bg-slate-50 p-0").props("width=230 bordered"):
        with ui.list().classes("w-full pt-2"):
            for url, nom, icone in MENU:
                actif = url == chemin
                with ui.item(on_click=lambda u=url: ui.navigate.to(u)).classes(
                        "rounded-r-full mr-3 " + ("bg-blue-100 text-blue-900 font-semibold" if actif else "")):
                    with ui.item_section().props("avatar"):
                        ui.icon(icone, color="primary" if actif else "grey-8")
                    with ui.item_section():
                        ui.label(nom)
    with ui.column().classes("w-full max-w-7xl mx-auto p-6 gap-5"):
        yield


def carte(titre: str, valeur, icone: str, couleur: str = "primary", sous: str = ""):
    with ui.card().classes("min-w-48 flex-1"):
        with ui.row().classes("items-center gap-3 no-wrap"):
            ui.icon(icone, size="lg", color=couleur)
            with ui.column().classes("gap-0"):
                ui.label(str(valeur)).classes("text-2xl font-bold")
                ui.label(titre).classes("text-sm text-gray-600")
                if sous:
                    ui.label(sous).classes("text-xs text-gray-500")


def base_absente():
    with ui.card().classes("w-full bg-amber-50"):
        ui.label("La base n'est pas encore sur cet ordinateur").classes("text-lg font-semibold")
        ui.label("Allez dans Paramètres, vérifiez le dépôt GitHub, puis cliquez sur « Synchroniser maintenant ».")
        ui.button("Ouvrir les paramètres", icon="settings", on_click=lambda: ui.navigate.to("/parametres"))


# --- tableau de bord (mêmes filtres, indicateurs et graphique que le classeur Excel) --------------------
ORANGE, BLEU_SERIE, GRIS = "#F2A33A", "#1F5FA8", "#6B7A90"
TOUTES = {"nature": "Toutes les natures", "region": "Toutes les régions", "entreprise": "Toutes les entreprises",
          "statut": "Tous les statuts"}
LIB_FILTRE = {"nature": "Nature du marché", "region": "Région", "entreprise": "Entreprise", "statut": "Statut du chantier"}


def fr(v, dec=1) -> str:
    return f"{v:,.{dec}f}".replace(",", " ").replace(".", ",")


def tuile(titre: str, clic=None):
    c = ui.card().classes("p-4 gap-1 min-h-32" + (" cursor-pointer hover:shadow-lg" if clic else ""))
    if clic:
        c.on("click", clic)
    with c:
        ui.label(titre).classes("text-xs uppercase tracking-wide text-gray-500 font-semibold")
    return c


def options_graphique(tb) -> dict:
    noms = [(g['libelle'][:23] + '…' if len(g['libelle']) > 24 else g['libelle']) + ("  – n.d." if g["conso"] is None else "")
            for g in tb.groupes]
    val = lambda x: round(x, 1) if x is not None else None                 # noqa: E731
    etiquette = {"show": True, "position": "top", "fontSize": 10, "color": "#1B2B41",
                 ":formatter": "(p) => p.value == null ? '' : String(p.value).replace('.', ',') + ' %'"}
    return {
        "color": [ORANGE, BLEU_SERIE, GRIS],
        "tooltip": {"trigger": "axis", "axisPointer": {"type": "shadow"}, ":valueFormatter": "(v) => v == null ? 'n.d.' : String(v).replace('.', ',') + ' %'"},
        "legend": {"bottom": 0, "textStyle": {"color": "#1B2B41"}},
        "grid": {"left": 48, "right": 16, "top": 24, "bottom": 96},
        "xAxis": {"type": "category", "data": noms, "axisLabel": {"rotate": 30, "fontSize": 11, "color": "#1B2B41", "interval": 0},
                  "axisTick": {"show": False}, "axisLine": {"lineStyle": {"color": "#C9D3E0"}}},
        "yAxis": {"type": "value", "min": 0, "axisLabel": {"formatter": "{value}%", "color": "#1B2B41"},
                  "splitLine": {"lineStyle": {"color": "#E3E8EF"}}},
        "series": [
            {"name": "Consommation du délai (%)", "type": "bar", "barGap": "10%", "label": etiquette,
             "data": [val(g["conso"]) for g in tb.groupes]},
            {"name": "Exécution physique (%)", "type": "bar", "label": etiquette, "data": [val(g["execution"]) for g in tb.groupes]},
            {"name": "Délai prévu atteint (100 %)", "type": "line", "symbol": "none", "lineStyle": {"type": "dashed", "width": 2},
             "data": [100] * len(tb.groupes)},
        ],
    }


def liste_alertes(alertes, limite=None):
    cols = [{"name": "n", "label": "N° activité", "field": "n", "align": "left"},
            {"name": "code", "label": "Code", "field": "code", "align": "left"},
            {"name": "reg", "label": "Région", "field": "reg", "align": "left"},
            {"name": "dep", "label": "Département", "field": "dep", "align": "left"},
            {"name": "ent", "label": "Entreprise", "field": "ent", "align": "left"},
            {"name": "nat", "label": "Nature", "field": "nat", "align": "left"},
            {"name": "debut", "label": "Début des travaux", "field": "debut", "align": "left"},
            {"name": "taux", "label": "Taux", "field": "taux"},
            {"name": "conso", "label": "Délai consommé", "field": "conso"}]
    rows = [{"n": a.numero, "code": a.code, "reg": a.region, "dep": a.departement.title(), "ent": a.entreprise or "—", "nat": a.nature,
             "debut": d(a.debut) or "—", "taux": f"{a.taux:.0f} %", "conso": f"{a.conso:.0f} %"} for a in alertes[:limite]]
    return cols, rows


@ui.page("/", response_timeout=30)
async def accueil():
    with cadre("/", "Tableau de bord"):
        if not S.base_existe():
            base_absente()
            return
        etat = {k: "" for k in TOUTES}
        with ui.row().classes("w-full gap-5 no-wrap items-start"):
            with ui.card().classes("w-64 shrink-0 p-4 gap-3 sticky top-20"):
                ui.label("▼  FILTRES").classes("font-bold text-primary")
                selects = {}
                for k in TOUTES:
                    selects[k] = ui.select({"": TOUTES[k]}, value="", label=LIB_FILTRE[k]).classes("w-full").mark(f"filtre-{k}")
                avert = ui.label().classes("text-xs text-warning")
                ui.button("Réinitialiser", icon="filter_alt_off", on_click=lambda: reinitialiser()).props("flat dense")
            zone = ui.column().classes("flex-1 min-w-0 gap-5")

        async def charger(sync: bool = False):
            if sync:
                ok, rep = await appel(S.synchroniser)
                if ok:
                    ui.notify(rep.text() if rep else "GitHub n'est pas configuré.", type="positive" if rep else "warning")
            ok, res = await appel(S.tableau_de_bord_excel, **etat)
            if not ok:
                return
            tb, journal, a_verifier, synchro = res
            for k, sel in selects.items():                                   # listes liées, comme dans l'Excel
                opts = {"": TOUTES[k], **{v: v for v in tb.options[k]}}
                if etat[k] and etat[k] not in opts:
                    opts[etat[k]] = etat[k]
                sel.set_options(opts, value=etat[k])
            avert.text = (f"⚠ Filtre ignoré (aucune activité avec les autres filtres) : {', '.join(tb.ignores)}. "
                          "Choisir « Toutes… » pour réinitialiser.") if tb.ignores else ""
            n = len(tb.selection)
            zone.clear()
            with zone:
                with ui.grid(columns=3).classes("w-full gap-4"):
                    with tuile("Activités sélectionnées / total"):
                        ui.label(f"{n} / {tb.total}").classes("text-3xl font-bold").mark("kpi-selection")
                        if a_verifier:
                            ui.link(f"{a_verifier} fiche(s) Kobo à valider", "/suivis").classes("text-sm")
                    with tuile("Taux d'exécution physique moyen"):
                        ui.label(f"{fr(tb.taux_moyen)} %" if tb.taux_moyen is not None else "—").classes("text-3xl font-bold").mark("kpi-taux")
                    with tuile("Situation des PGES"):
                        ui.label(f"{tb.pges_valides} Validés").classes("text-3xl font-bold")
                        ui.label(f"ONG mobilisées : {tb.ong_nommees}").classes("text-sm")
                        if tb.ong_sans_nom:
                            ui.label(f"{tb.ong_sans_nom} activité(s) avec ONG non nommée").classes("text-xs text-gray-600")
                    with tuile("Montant total des marchés"):
                        ui.label(f"{fr(tb.montant_total / 1e9, 2)} Mds FCFA").classes("text-3xl font-bold")
                        ui.label(f"Montant décaissé : {fr(tb.decaisse_total / 1e6)} M FCFA").classes("text-sm")
                        if tb.taux_financier is not None:
                            ui.label(f"Taux d'exécution financière : {fr(tb.taux_financier)} %").classes("text-sm")
                    with tuile("Activités en alerte  ▸ détail", clic=lambda: detail_alertes(tb)):
                        ui.label(f"{len(tb.alertes)} / {n}").classes("text-3xl font-bold " + ("text-negative" if tb.alertes else "text-positive"))
                        ui.label(f"Exécution < {tb.seuils[0]} % ET délai consommé ≥ {tb.seuils[1]} %").classes("text-xs text-gray-600")
                    with tuile("Aperçu des alertes  ▸ voir tout", clic=lambda: detail_alertes(tb)):
                        if not tb.alertes:
                            ui.label("Aucune activité en alerte dans le périmètre.").classes("text-sm text-positive")
                        for a in tb.alertes[:5]:
                            with ui.row().classes("w-full gap-2 no-wrap text-xs"):
                                ui.label(a.numero).classes("font-semibold w-12")
                                ui.label(a.region).classes("w-20 truncate")
                                ui.label(a.entreprise or "—").classes("flex-1 truncate")
                                ui.label(f"{a.taux:.0f}%").classes("w-8 text-right text-negative")
                        if len(tb.alertes) > 5:
                            ui.label(f"… et {len(tb.alertes) - 5} autre(s) : voir le détail ▸").classes("text-xs text-primary")
                with ui.card().classes("w-full"):
                    ui.label(f"Consommation des délais par {tb.par}   (délai consommé vs exécution physique)") \
                        .classes("text-lg font-semibold").mark("titre-graphique")
                    if tb.groupes:
                        ui.echart(options_graphique(tb)).classes("w-full h-96").mark("graphique")
                    else:
                        ui.label("Aucune activité évaluable dans le périmètre.").classes("text-gray-600")
                    evaluees = sum(1 for l in tb.selection if l.evaluable)
                    excl = ", ".join(f"{v} {k}" for k, v in tb.exclusions.items() if v) or "aucune"
                    ui.label(f"Méthode : consommation du délai = Σ temps écoulé ÷ Σ durées prévues (depuis la date de début des travaux) ; "
                             f"exécution physique pondérée par les montants des marchés (moyenne simple si montants absents). "
                             f"Activités évaluées : {evaluees} sur {n}. Exclues : {excl}.").classes("text-xs text-gray-500")
                with ui.row().classes("w-full items-center gap-4 text-sm text-gray-600"):
                    if journal:
                        j = journal[0]
                        ui.icon("check_circle" if j["Resultat"] == "OK" else "cancel", color="positive" if j["Resultat"] == "OK" else "negative")
                        ui.label(f"Dernier traitement : {d(j['Date_Execution'])} ({j['Declencheur'] or ''})")
                    ui.space()
                    ui.icon("cloud_sync")
                    ui.label(f"GitHub : {d(datetime.fromisoformat(synchro)) if synchro else 'jamais synchronisé'}")
                    ui.button("Synchroniser", icon="sync", on_click=lambda: charger(True)).props("flat dense")

        def detail_alertes(tb):
            with ui.dialog() as dlg, ui.card().classes("w-[1100px] max-w-full"):
                with ui.row().classes("w-full items-center"):
                    ui.label(f"Alertes et points d'attention — {len(tb.alertes)} activité(s)").classes("text-lg font-semibold")
                    ui.space()
                    ui.button(icon="close", on_click=dlg.close).props("flat round dense")
                perim = [f"{LIB_FILTRE[k]} : {getattr(tb.effectifs, k)}" for k in TOUTES if getattr(tb.effectifs, k)]
                ui.label("Périmètre : " + ("; ".join(perim) if perim else "toutes les activités")).classes("text-sm text-gray-600")
                cols, rows = liste_alertes(tb.alertes)
                ui.table(columns=cols, rows=rows, row_key="n", pagination=15).props("dense wrap-cells").classes("w-full")
                ui.label(f"Retard potentiel d'exécution : exécution < {tb.seuils[0]} % ET délai consommé ≥ {tb.seuils[1]} %. "
                         "Classement par écart (délai consommé − exécution) décroissant.").classes("text-xs text-gray-500")
            dlg.open()

        async def changer(k, e):
            if etat[k] != (e.value or ""):
                etat[k] = e.value or ""
                await charger()

        async def reinitialiser():
            for k in etat:
                etat[k] = ""
            await charger()

        for k, sel in selects.items():
            sel.on_value_change(lambda e, k=k: changer(k, e))
        await charger()


# --- lancer le suivi / documents ---------------------------------------------------------------------
@ui.page("/suivi", response_timeout=30)
async def suivi():
    with cadre("/suivi", "Lancer le suivi"):
        if not S.base_existe():
            base_absente()
            return
        # --- 1. sur GitHub : exactement le traitement du lundi, déclenché à la demande -----------------
        with ui.card().classes("w-full border-l-4 border-primary"):
            with ui.row().classes("items-center gap-2"):
                ui.icon("cloud_sync", color="primary", size="sm")
                ui.label("Lancer le traitement sur GitHub").classes("text-lg font-semibold")
                ui.badge("recommandé", color="primary").props("outline")
            ui.label("Les mêmes traitements que ceux du lundi, avec les accès Kobo et e-mail enregistrés sur GitHub. "
                     "Les documents produits arrivent ensuite dans votre dossier « sorties ». Durée : quelques minutes.") \
                .classes("text-sm text-gray-600")
            with ui.row().classes("items-center gap-6"):
                gh_quoi = ui.toggle({"matrices": "Matrices Word (fiches Kobo → base → Word)",
                                     "tableau": "Tableau de bord Excel"}, value="matrices").mark("gh-quoi")
                gh_mail = ui.checkbox("Envoyer par e-mail", value=False).mark("gh-mail")
                gh_bouton = ui.button("Lancer sur GitHub", icon="rocket_launch").mark("lancer-github")
            gh_etat = ui.row().classes("w-full items-center gap-3")
            gh_etat.visible = False
            with gh_etat:
                gh_spin = ui.spinner(size="md")
                gh_texte = ui.label().classes("text-sm").mark("etat-github")
                ui.space()
                gh_lien = ui.link("Voir sur GitHub", "#", new_tab=True).classes("text-sm")
                gh_lien.visible = False

        suivi_actif = {"on": False}

        def afficher_etat(texte: str, en_cours: bool, url: str | None = None):
            gh_etat.visible = True
            gh_texte.text = texte
            gh_spin.visible = en_cours
            gh_bouton.set_enabled(not en_cours)
            if url:
                gh_lien._props["href"] = url
                gh_lien.update()
                gh_lien.visible = True

        async def lancer_github():
            ok, _ = await appel(S.lancer_sur_github, gh_mail.value, gh_quoi.value)
            if ok:
                afficher_etat("Traitement demandé à GitHub…", True)
                suivi_actif["on"] = True

        async def verifier_github():
            if not suivi_actif["on"]:
                return
            ok, st = await appel(S.suivre_github)
            if not ok:
                suivi_actif["on"] = False
                afficher_etat("Suivi interrompu : voir le message d'erreur.", False)
                return
            if st is None:
                suivi_actif["on"] = False
                return
            if st["status"] != "completed":
                libelle = {"queued": "En file d'attente sur GitHub", "in_progress": "En cours sur GitHub"}.get(st["status"], st["status"])
                afficher_etat(f"{libelle}" + (f" : {st['etape']}" if st.get("etape") else "") + "…", True, st.get("url"))
                return
            suivi_actif["on"] = False
            if st["conclusion"] != "success":
                S.abandonner_suivi_github()
                afficher_etat(f"Le traitement GitHub a échoué ({st['conclusion']}). Ouvrez « Voir sur GitHub » pour le détail ; "
                              "la base n'a pas été modifiée.", False, st.get("url"))
                ui.notify("Le traitement GitHub a échoué.", type="negative")
                return
            afficher_etat("Traitement terminé : récupération de la base et des documents…", True, st.get("url"))
            ok, res = await appel(S.recuperer_github)
            if ok:
                n = len(res["documents"])
                afficher_etat(f"Terminé : {n} document(s) récupéré(s) dans « sorties »" + (f" ({', '.join(res['documents'])})" if n else "")
                              + ".", False, st.get("url"))
                ui.notify(f"Traitement GitHub terminé : {n} document(s) récupéré(s).", type="positive", multi_line=True)
                docs.refresh()

        gh_bouton.on_click(lancer_github)
        ui.timer(5.0, verifier_github)
        if S.execution_en_cours():                                    # reprise après fermeture de l'application
            suivi_actif["on"] = True
            afficher_etat("Traitement GitHub en cours (lancé précédemment)…", True)

        # --- 2. sur cet ordinateur -------------------------------------------------------------------
        with ui.card().classes("w-full"):
            ui.label("Ou générer sur cet ordinateur").classes("text-lg font-semibold")
            ui.label("Matrices : récupère les nouvelles fiches Kobo (et leurs photos), met la base à jour et génère les matrices Word. "
                     "Tableau de bord : génère le classeur Excel depuis la base. Accès Kobo et e-mail à renseigner dans Paramètres.") \
                .classes("text-sm text-gray-600")
            with ui.row().classes("items-end gap-6"):
                jour = ui.date_input("Situation au", value=date.today().isoformat()).classes("w-48")
                avec_kobo = ui.checkbox("Récupérer les fiches Kobo", value=S.settings.kobo_ok)
                mail = ui.checkbox("Envoyer par e-mail", value=False)
            with ui.row().classes("gap-3"):
                bouton = ui.button("Générer les matrices ici", icon="description").props("outline").mark("lancer")
                bouton_tb = ui.button("Générer le tableau de bord ici", icon="insert_chart").props("outline").mark("lancer-tableau")
                ui.button("Envoyer la liste des activités à Kobo", icon="upload",
                          on_click=lambda: envoyer_liste()).props("outline")
            barre = ui.linear_progress(show_value=False).props("indeterminate").classes("w-full")
            barre.visible = False
            journal = ui.log(max_lines=60).classes("w-full h-40")
            journal.visible = False

        async def lancer():
            bouton.disable()
            barre.visible = True
            try:
                ok, r = await appel(S.lancer, au=date.fromisoformat(jour.value) if jour.value else None,
                                    envoyer_mail=mail.value, avec_kobo=avec_kobo.value)
            finally:
                bouton.enable()
                barre.visible = False
            if ok:
                journal.visible = True
                journal.clear()
                for line in r.details or []:
                    journal.push(line)
                ui.notify(r.message, type="positive" if r.ok else "negative", multi_line=True)
                docs.refresh()

        async def lancer_tableau():
            bouton_tb.disable()
            barre.visible = True
            try:
                ok, r = await appel(S.generer_tableau_de_bord, au=date.fromisoformat(jour.value) if jour.value else None,
                                    envoyer_mail=mail.value)
            finally:
                bouton_tb.enable()
                barre.visible = False
            if ok:
                journal.visible = True
                journal.clear()
                for line in r.details or []:
                    journal.push(line)
                ui.notify(r.message, type="positive" if r.ok else "negative", multi_line=True)
                docs.refresh()

        bouton_tb.on_click(lancer_tableau)

        async def envoyer_liste():
            ok, msg = await appel(S.envoyer_liste_kobo)
            if ok:
                ui.notify(msg, type="positive")

        bouton.on_click(lancer)

        @ui.refreshable
        def docs():
            with ui.card().classes("w-full"):
                with ui.row().classes("w-full items-center"):
                    ui.label("Documents générés").classes("text-lg font-semibold")
                    ui.space()
                    ui.button("Ouvrir le dossier", icon="folder_open", on_click=lambda: S.ouvrir(S.ws / "sorties")).props("flat")
                fichiers = S.documents()
                if not fichiers:
                    ui.label("Aucun document pour l'instant.").classes("text-gray-600")
                for f in fichiers:
                    with ui.row().classes("w-full items-center gap-3 no-wrap hover:bg-slate-50 px-2 rounded"):
                        ui.icon("description" if f.suffix == ".docx" else "table_chart",
                                color="blue-8" if f.suffix == ".docx" else "green-8")
                        ui.label(f.name).classes("flex-1 text-sm")
                        ui.label(d(datetime.fromtimestamp(f.stat().st_mtime))).classes("text-xs text-gray-500")
                        ui.button("Ouvrir", on_click=lambda p=f: S.ouvrir(p)).props("flat dense")
        docs()


# --- suivis à valider ---------------------------------------------------------------------------------
STATUTS_SUIVI = {"A_VERIFIER": "À vérifier", "VALIDE": "Validés", "REJETE": "Rejetés", "": "Tous"}


@ui.page("/suivis", response_timeout=30)
async def suivis():
    with cadre("/suivis", "Suivis à valider"):
        if not S.base_existe():
            base_absente()
            return
        ui.label("Fiches envoyées par les agents depuis Kobo. Celles « à vérifier » n'apparaissent dans les matrices "
                 "qu'une fois validées. Cliquez sur une ligne pour voir le détail.").classes("text-sm text-gray-600")
        filtre = ui.toggle(STATUTS_SUIVI, value="A_VERIFIER")
        cols = [{"name": "date", "label": "Visite", "field": "date", "sortable": True, "align": "left"},
                {"name": "code", "label": "Activité", "field": "code", "sortable": True, "align": "left"},
                {"name": "dep", "label": "Département", "field": "dep", "align": "left"},
                {"name": "agent", "label": "Agent", "field": "agent", "align": "left"},
                {"name": "taux", "label": "Taux", "field": "taux", "sortable": True},
                {"name": "photo", "label": "Photo", "field": "photo"},
                {"name": "statut", "label": "Statut", "field": "statut", "align": "left"},
                {"name": "motif", "label": "Motif", "field": "motif", "align": "left"}]
        table = ui.table(columns=cols, rows=[], row_key="id", pagination=15).props("wrap-cells").classes("w-full")
        donnees: dict = {}

        async def charger():
            ok, rows = await appel(S.suivis, filtre.value or None)
            if not ok:
                return
            donnees.clear()
            donnees.update({str(r["ID_SOUMISSION_KOBO"]): r for r in rows})
            table.rows = [{"id": str(r["ID_SOUMISSION_KOBO"]), "date": d(r["Date_Actualisation"]), "code": r["CODE_ACTIVITE"] or "?",
                           "dep": r["Departement"] or "", "agent": r["Nom_Agent"] or "", "taux": pct(r["Taux_Execution_Physique"]),
                           "photo": "📷" if r["photo_chemin"] else "", "statut": STATUTS_SUIVI.get(r["Statut_Traitement"], r["Statut_Traitement"]),
                           "motif": r["Motif_Rejet"] or ""} for r in rows]

        def detail(e):
            r = donnees.get(str(e.args[1]["id"]))
            if not r:
                return
            with ui.dialog() as dlg, ui.card().classes("w-[860px] max-w-full"):
                with ui.row().classes("w-full items-center"):
                    ui.label(f"{r['CODE_ACTIVITE'] or '?'} – {r['Departement'] or ''}").classes("text-lg font-semibold")
                    ui.space()
                    ui.badge(STATUTS_SUIVI.get(r["Statut_Traitement"], ""), color="warning" if r["Statut_Traitement"] == "A_VERIFIER" else
                             "positive" if r["Statut_Traitement"] == "VALIDE" else "negative")
                with ui.row().classes("w-full gap-6 no-wrap items-start"):
                    with ui.column().classes("flex-1 gap-2"):
                        for lib, val in (("Date de la visite", d(r["Date_Actualisation"])), ("Agent", r["Nom_Agent"]),
                                         ("Taux d'exécution physique", pct(r["Taux_Execution_Physique"])),
                                         ("Fiche Kobo n°", r["ID_SOUMISSION_KOBO"]), ("Reçue le", d(r["Date_Importation"]))):
                            with ui.row().classes("gap-2"):
                                ui.label(f"{lib} :").classes("font-semibold")
                                ui.label(str(val or "—"))
                        for lib, val in (("Observations techniques", r["Observations_Techniques"]), ("Commentaires", r["Commentaires"]),
                                         ("Difficultés", r["Difficultes"]), ("Recommandations", r["Recommandations"])):
                            ui.label(lib).classes("font-semibold mt-2")
                            ui.label(val or "—").classes("whitespace-pre-line text-sm")
                        if r["Motif_Rejet"]:
                            ui.label("Motif du rejet" if r["Statut_Traitement"] == "REJETE" else "Pourquoi cette fiche est à vérifier") \
                                .classes("font-semibold mt-2 text-negative")
                            ui.label(r["Motif_Rejet"]).classes("text-sm")
                    if r["photo_chemin"]:
                        ui.image(Path(r["photo_chemin"])).classes("w-72 rounded shadow")
                motif = ui.input("Motif du rejet (obligatoire pour rejeter)").classes("w-full")
                with ui.row().classes("w-full justify-end gap-2"):
                    ui.button("Fermer", on_click=dlg.close).props("flat")

                    async def decider(valider: bool):
                        ok, avert = await appel(S.decider_suivi, r["ID_SOUMISSION_KOBO"], valider, motif.value)
                        if ok:
                            ui.notify("Suivi validé." if valider else "Suivi rejeté.", type="positive")
                            avertir(avert)
                            dlg.close()
                            await charger()
                    if r["Statut_Traitement"] != "REJETE":
                        ui.button("Rejeter", icon="block", color="negative", on_click=lambda: decider(False)).mark("rejeter")
                    if r["Statut_Traitement"] != "VALIDE":
                        ui.button("Valider", icon="check", color="positive", on_click=lambda: decider(True)).mark("valider")
            dlg.open()

        table.on("rowClick", detail)
        table.classes("cursor-pointer")
        filtre.on_value_change(charger)
        await charger()


# --- activités ----------------------------------------------------------------------------------------
@ui.page("/activites", response_timeout=30)
async def activites():
    with cadre("/activites", "Activités"):
        if not S.base_existe():
            base_absente()
            return
        cfg = S.cfg
        with ui.row().classes("w-full items-center gap-3"):
            recherche = ui.input(placeholder="Rechercher (code, département, titulaire…)").props("clearable outlined dense") \
                .classes("w-96")
            cat = ui.select({"": "Toutes catégories", **cfg.categories}, value="").props("outlined dense").classes("w-48")
            archivees = ui.checkbox("Afficher les archivées")
            ui.space()
            ui.button("Nouvelle activité", icon="add", on_click=lambda: formulaire(None)).mark("nouvelle")
        cols = [{"name": "code", "label": "Code", "field": "code", "sortable": True, "align": "left"},
                {"name": "reg", "label": "Région", "field": "reg", "sortable": True, "align": "left"},
                {"name": "dep", "label": "Département", "field": "dep", "sortable": True, "align": "left"},
                {"name": "ent", "label": "Titulaire", "field": "ent", "sortable": True, "align": "left"},
                {"name": "montant", "label": "Montant (FCFA)", "field": "montant"},
                {"name": "chantier", "label": "Chantier", "field": "chantier", "sortable": True, "align": "left"},
                {"name": "taux", "label": "Taux phys.", "field": "taux", "sortable": True},
                {"name": "delai", "label": "Délai consommé", "field": "delai", "sortable": True},
                {"name": "visite", "label": "Dernière visite", "field": "visite", "align": "left"}]
        table = ui.table(columns=cols, rows=[], row_key="code", pagination=20).props("dense wrap-cells").classes("w-full cursor-pointer text-sm")
        table.bind_filter_from(recherche, "value")
        tout: list = []

        def afficher():
            table.rows = [{"code": a["CODE_ACTIVITE"], "reg": a["Region_nom"], "dep": a["Departement"] or "", "ent": a["Entreprise"] or "—",
                           "montant": money(a["Montant_Marche"]), "chantier": a["Statut_Chantier"] or "",
                           "taux": pct(a["Taux"]), "delai": pct(a["Delai_consomme"]), "visite": d(a["Derniere_visite"])}
                          for a in tout if (not cat.value or a["Categorie"] == cat.value) and (archivees.value or a["Statut"] == "Actif")]

        async def charger():
            ok, rows = await appel(S.activites)
            if ok:
                tout[:] = rows
                afficher()

        cat.on_value_change(afficher)
        archivees.on_value_change(afficher)

        def champ(f, valeur):
            lib = LIBELLES.get(f, f)
            if f in CHOIX:
                return ui.select([""] + CHOIX[f], value=valeur or "", label=lib).classes("w-full")
            if f in DATES:
                return ui.date_input(lib, value=valeur.isoformat() if valeur else "").classes("w-full")
            if f == "Designation":
                return ui.textarea(lib, value=valeur or "").props("autogrow").classes("w-full")
            if f in MONTANTS:
                return ui.input(lib, value=money(valeur)).classes("w-full")
            if f == "ONG_Mobilisees":                                       # suggestions : évite les doublons d'écriture
                noms = sorted({a["ONG_Mobilisees"] for a in tout if a.get("ONG_Mobilisees")} | {"Non mobilisée", "Mobilisée"},
                              key=str.casefold)
                return ui.input(lib, value=valeur or "", autocomplete=noms).props("hint='Commencez à taper : les ONG déjà saisies sont proposées'") \
                    .classes("w-full")
            return ui.input(lib, value="" if valeur is None else str(valeur)).classes("w-full")

        def formulaire(code: str | None):
            act = next((a for a in tout if a["CODE_ACTIVITE"] == code), None) if code else None
            entrees: dict = {}
            with ui.dialog().props("maximized=false") as dlg, ui.card().classes("w-[900px] max-w-full"):
                with ui.row().classes("w-full items-center"):
                    ui.label(f"{code} – {act['Departement'] or ''}" if act else "Nouvelle activité").classes("text-lg font-semibold")
                    ui.space()
                    ui.button(icon="close", on_click=dlg.close).props("flat round dense")
                with ui.tabs().classes("w-full") as onglets:
                    t1, t2 = ui.tab("Marché"), ui.tab("Chantier")
                    t3 = ui.tab("Historique") if act else None
                with ui.tab_panels(onglets, value=t1).classes("w-full"):
                    with ui.tab_panel(t1):
                        with ui.grid(columns=2).classes("w-full gap-x-4"):
                            if act:
                                ui.input(LIBELLES["Categorie"], value=act["Categorie_nom"]).props("readonly")
                                ui.input(LIBELLES["Region"], value=act["Region_nom"]).props("readonly")
                            else:
                                entrees["_cat"] = ui.select(cfg.categories, label=LIBELLES["Categorie"]).mark("f-cat")
                                entrees["_reg"] = ui.select(cfg.regions, label=LIBELLES["Region"]).mark("f-reg")
                            for f in FIXES:
                                entrees[f] = champ(f, act.get(f) if act else None).mark(f"f-{f}")
                            if act:
                                ui.input(LIBELLES["Date_Fin_Previsionnelle"], value=d(act["Date_Fin_Previsionnelle"])).props("readonly")
                    with ui.tab_panel(t2):
                        with ui.grid(columns=2).classes("w-full gap-x-4"):
                            for f in CHANTIER:
                                entrees[f] = champ(f, act.get(f) if act else None).mark(f"f-{f}")
                            entrees["Statut"] = ui.select(CHOIX["Statut"], value=(act or {}).get("Statut") or "Actif",
                                                          label=LIBELLES["Statut"])
                        taux_fin = ui.label().classes("text-sm text-gray-600 mt-2").mark("taux-financier")

                        def maj_taux_fin(_=None):
                            try:
                                m = int(str(entrees["Montant_Marche"].value or "").replace(" ", "").replace("\u00a0", "") or 0)
                                dcs = int(str(entrees["Montant_Decaisse"].value or "").replace(" ", "").replace("\u00a0", "") or 0)
                                taux_fin.text = (f"Taux d'exécution financière : {dcs / m * 100:.1f} %".replace(".", ",") if m and dcs
                                                 else "Taux d'exécution financière : renseignez le montant du marché et le montant décaissé")
                            except ValueError:
                                taux_fin.text = "Montants : chiffres uniquement (ex. 12 500 000)"
                        maj_taux_fin()
                        entrees["Montant_Decaisse"].on_value_change(maj_taux_fin)
                        entrees["Montant_Marche"].on_value_change(maj_taux_fin)
                        ui.label("Ces informations sont saisies ici par l'administrateur (les agents ne les renseignent pas dans Kobo) : "
                                 "elles alimentent le tableau de bord (statut, PGES, ONG, finances).").classes("text-xs text-gray-500")
                    if t3:
                        with ui.tab_panel(t3):
                            hist = ui.column().classes("w-full gap-1")

                            async def charger_hist():
                                ok, h = await appel(S.historique, code)
                                if ok:
                                    with hist:
                                        if not h:
                                            ui.label("Aucune modification enregistrée.")
                                        for x in h:
                                            ui.label(f"{d(x['Date_Modification'])} · {x['Utilisateur'] or ''} · "
                                                     f"{LIBELLES.get(x['Champ_Modifie'], x['Champ_Modifie'])} : "
                                                     f"{x['Ancienne_Valeur'] or '∅'} → {x['Nouvelle_Valeur'] or '∅'}").classes("text-sm")
                            onglets.on_value_change(lambda e: charger_hist() if e.value == "Historique" and not hist.default_slot.children else None)
                ui.label("Laisser un champ vide = ne rien changer. Pour effacer une information, videz le champ : "
                         "il vous sera demandé de confirmer.").classes("text-xs text-gray-500")

                async def enregistrer():
                    valeurs = {f: (w.value or "") for f, w in entrees.items() if not f.startswith("_")}
                    if act is None:
                        if not entrees["_cat"].value or not entrees["_reg"].value:
                            ui.notify("Choisissez la catégorie et la région.", type="warning")
                            return
                        ok, res = await appel(S.creer_activite, entrees["_cat"].value, entrees["_reg"].value, valeurs)
                        if ok:
                            nouveau, avert = res
                            ui.notify(f"Activité créée : {nouveau}. Pensez à envoyer la liste à Kobo (écran « Lancer le suivi »).",
                                      type="positive", multi_line=True)
                            avertir(avert)
                            dlg.close()
                            await charger()
                        return
                    effacer = [f for f, v in valeurs.items() if f != "Statut" and not str(v).strip() and act.get(f) not in (None, "")]
                    if effacer and not await confirmer(f"Effacer : {', '.join(LIBELLES.get(f, f) for f in effacer)} ?"):
                        return
                    ok, res = await appel(S.modifier_activite, code, valeurs, effacer)
                    if ok:
                        changes, avert = res
                        ui.notify(f"{len(changes)} modification(s) enregistrée(s)." if changes else "Aucune modification.",
                                  type="positive" if changes else "info")
                        avertir(avert)
                        dlg.close()
                        await charger()

                with ui.row().classes("w-full justify-end gap-2"):
                    ui.button("Annuler", on_click=dlg.close).props("flat")
                    ui.button("Enregistrer", icon="save", on_click=enregistrer).mark("enregistrer")
            dlg.open()

        table.on("rowClick", lambda e: formulaire(e.args[1]["code"]))
        await charger()


async def confirmer(question: str) -> bool:
    with ui.dialog() as dlg, ui.card():
        ui.label(question)
        with ui.row().classes("w-full justify-end"):
            ui.button("Non", on_click=lambda: dlg.submit(False)).props("flat")
            ui.button("Oui", on_click=lambda: dlg.submit(True), color="negative")
    return bool(await dlg)


# --- paramètres ---------------------------------------------------------------------------------------
@ui.page("/parametres", response_timeout=30)
async def parametres():
    with cadre("/parametres", "Paramètres"):
        s = S.settings
        w: dict = {}

        def section(titre, icone, texte):
            c = ui.card().classes("w-full")
            with c:
                with ui.row().classes("items-center gap-2"):
                    ui.icon(icone, color="primary")
                    ui.label(titre).classes("text-lg font-semibold")
                ui.label(texte).classes("text-sm text-gray-600")
            return c

        with section("Général", "person", "Nom écrit dans l'historique des modifications, et dossier de travail "
                     "(documents générés, base, photos)."):
            with ui.grid(columns=2).classes("w-full gap-x-4"):
                w["administrateur"] = ui.input("Votre nom", value=s.administrateur)
                w["workspace"] = ui.input("Dossier de travail", value=s.workspace)

        with section("GitHub", "cloud_sync", "La base est conservée sur GitHub : le logiciel récupère la dernière version "
                     "avant chaque modification et renvoie la base après. Le traitement automatique du lundi continue de "
                     "fonctionner, même ordinateur éteint."):
            with ui.grid(columns=3).classes("w-full gap-x-4"):
                w["github_depot"] = ui.input("Dépôt (propriétaire/nom)", value=s.github_depot)
                w["github_branche"] = ui.input("Branche", value=s.github_branche)
                w["github_jeton"] = ui.input("Jeton d'accès", value=s.github_jeton, password=True, password_toggle_button=True)
            ui.label("Jeton : github.com → Settings → Developer settings → Fine-grained tokens, accès au dépôt, permissions "
                     "« Contents : Read and write » et « Actions : Read and write » (pour lancer le traitement).").classes("text-xs text-gray-500")
            with ui.row().classes("gap-2"):
                ui.button("Tester", icon="wifi_tethering", on_click=lambda: tester(S.tester_github)).props("outline")
                ui.button("Synchroniser maintenant", icon="sync", on_click=lambda: synchro(None)).props("outline")
                with ui.button("Conflit : garder…", icon="call_split").props("outline color=warning"):
                    with ui.menu():
                        ui.menu_item("la version GitHub (la version locale est sauvegardée)", on_click=lambda: synchro("github"))
                        ui.menu_item("la version de cet ordinateur (remplace celle de GitHub)", on_click=lambda: synchro("local"))

        with section("KoboToolbox", "assignment", "Formulaire de suivi rempli par les agents (fiches et photos)."):
            with ui.grid(columns=3).classes("w-full gap-x-4"):
                w["kobo_url"] = ui.input("Adresse du serveur", value=s.kobo_url)
                w["kobo_asset_uid"] = ui.input("UID du formulaire", value=s.kobo_asset_uid)
                w["kobo_jeton"] = ui.input("Jeton API", value=s.kobo_jeton, password=True, password_toggle_button=True)
            ui.button("Tester", icon="wifi_tethering", on_click=lambda: tester(S.tester_kobo)).props("outline")

        with section("E-mail", "mail", "Envoi des matrices aux destinataires."):
            with ui.grid(columns=3).classes("w-full gap-x-4"):
                w["smtp_hote"] = ui.input("Serveur SMTP (ex. smtp.gmail.com)", value=s.smtp_hote)
                w["smtp_port"] = ui.input("Port", value=s.smtp_port)
                w["smtp_utilisateur"] = ui.input("Utilisateur", value=s.smtp_utilisateur)
                w["smtp_mot_de_passe"] = ui.input("Mot de passe (d'application)", value=s.smtp_mot_de_passe,
                                                  password=True, password_toggle_button=True)
                w["mail_expediteur"] = ui.input("Adresse de l'expéditeur", value=s.mail_expediteur)
                w["mail_destinataires"] = ui.input("Destinataires (séparés par des virgules)", value=s.mail_destinataires)
            ui.button("Envoyer un e-mail de test", icon="send", on_click=lambda: tester(S.tester_mail)).props("outline")

        async def enregistrer(silencieux=False):
            ok, _ = await appel(S.enregistrer_parametres, {k: v.value for k, v in w.items()})
            if ok and not silencieux:
                ui.notify("Paramètres enregistrés.", type="positive")
            return ok

        async def tester(fn):
            if await enregistrer(silencieux=True):
                ok, msg = await appel(fn)
                if ok:
                    ui.notify(msg, type="positive", multi_line=True)

        async def synchro(prefer):
            if await enregistrer(silencieux=True):
                ok, rep = await appel(S.synchroniser, True, prefer)
                if ok:
                    ui.notify(rep.text() if rep else "Indiquez d'abord le dépôt GitHub.", type="positive" if rep else "warning")

        with ui.row().classes("w-full justify-end"):
            ui.button("Enregistrer", icon="save", on_click=enregistrer).mark("enregistrer-parametres")
        ui.label(f"Paramètres conservés dans {s.path()} (votre profil Windows), jamais dans le dépôt.").classes("text-xs text-gray-500")


def demarrer(services: Services | None = None) -> None:
    global S
    S = services or Services()
