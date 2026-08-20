import importlib

import streamlit as st

import app_config as _app_config
import audit_webfleet_rda as _audit_webfleet_rda
import ltr_checks as _ltr_checks
import merge_files as _merge_files
import nexus_batch_runner as _nexus_batch_runner
import planning_download as _planning_download
import ui_common as _ui_common
import webfleet as _webfleet

_app_config = importlib.reload(_app_config)
_ui_common = importlib.reload(_ui_common)
_webfleet = importlib.reload(_webfleet)
_merge_files = importlib.reload(_merge_files)
_nexus_batch_runner = importlib.reload(_nexus_batch_runner)
_planning_download = importlib.reload(_planning_download)
_ltr_checks = importlib.reload(_ltr_checks)
_audit_webfleet_rda = importlib.reload(_audit_webfleet_rda)

TASKS = _app_config.TASKS
render_audit_task = _audit_webfleet_rda.render_audit_task
render_ltr_task = _ltr_checks.render_ltr_task
render_merge_task = _merge_files.render_merge_task
render_nexus_batch_runner_task = _nexus_batch_runner.render_nexus_batch_runner_task
render_planning_download_task = _planning_download.render_planning_download_task
render_webfleet_task = _webfleet.render_webfleet_task


def render_home_task() -> None:
    st.title("Accueil")
    st.caption("Outils internes pour télécharger, fusionner, préparer et auditer les données Webfleet, RDA et LTR.")
    st.markdown(
        """
        ### Comment commencer une tache
        - Choisissez une section dans le menu de droite.
        - Ajoutez les fichiers demandés, puis lancez le traitement.
        - Les boutons de téléchargement principaux apparaissent à côté du bouton de lancement.
        - Les fichiers générés doivent être téléchargés depuis l'application pendant la session.
        """
    )

    st.warning(
        "Pendant un traitement en cours, ne changez pas de section et n'utilisez pas d'autres parties de l'application. "
        "Attendez que le traitement soit terminé."
    )

    _ui_common.render_template_downloads()

    st.divider()

    st.subheader("Comprendre les sections")

    with st.expander("1. Téléchargement Webfleet"):
        st.markdown(
            """
            Cette section sert à télécharger les données brutes des trajets Webfleet sur une période donnée.
            Pour récupérer ces données, il faut un accès API Webfleet actif: identifiants Webfleet, clé API et activation
            API sur l'utilisateur concerné.

            Pour demander la clé API, envoyez un email à `support.de@webfleet.com`. La clé n'est pas créée dans
            l'application.

            Le message peut reprendre le modèle ci-dessous. Chaque ligne explique directement quelle information mettre :

            ```text
            Bonjour,

            Je souhaite demander une clé API Webfleet. Voici les informations nécessaires :

            Nom de l'application : indiquez le nom de l'outil ou de l'application qui utilisera l'API.
            Nom de l'intégrateur : indiquez le nom de votre entreprise ou de l'organisation responsable de l'intégration.
            Site web : indiquez le site web officiel de l'entreprise.
            Personne de contact : indiquez le nom de la personne que Webfleet peut contacter pour cette demande.
            Adresse : indiquez l'adresse postale de l'entreprise.
            Numéro de téléphone : indiquez le numéro de téléphone de la personne de contact.
            Email : indiquez l'adresse email de la personne de contact.

            Description de l'application :
            Nous souhaitons automatiser des rapports internes et des contrôles de conformité pour nos opérations. L'application récupère les trajets et les données d'utilisation des véhicules depuis Webfleet avec l'API .connect, traite les données dans notre outil interne, puis permet de télécharger les résultats pour intégration dans nos processus existants. Il ne s'agit pas d'une intégration de Webfleet dans une plateforme tierce établie, mais d'une utilisation interne pour obtenir un flux de travail plus flexible et automatisé, adapté à nos besoins.

            Merci de me dire si d'autres informations sont nécessaires pour traiter la demande.

            Cordialement,
            indiquez votre nom
            ```

            Après la réponse de Webfleet avec la clé API, le compte maître Webfleet doit se connecter à Webfleet, ouvrir
            les paramètres de l'utilisateur concerné, puis activer l'utilisation de l'API pour cet utilisateur. Sans cette
            activation côté utilisateur, le téléchargement Webfleet ne fonctionnera pas même si la clé API existe.
            """
        )

    with st.expander("2. Téléchargement planning"):
        st.markdown(
            """
            Cette section sert à récupérer les fichiers planning CSV depuis le serveur SFTP.
            L'utilisateur choisit une plage de dates planning, puis l'application télécharge les fichiers correspondants.

            Point important: la date dans le nom du fichier est toujours le lendemain de la date planning. Par exemple,
            pour le planning du 2026-06-30, l'application cherche un fichier avec `20260701` dans le nom, comme
            `HAS_HAM_20260701_2330_PEPS-Visits.csv`.

            L'application utilise les quatre champs SFTP: serveur, port, utilisateur et mot de passe.
            Si un fichier demandé n'est pas disponible sur le serveur, il est listé dans le résumé.

            Après la recherche, trois sorties peuvent être téléchargées: un ZIP avec tous les fichiers trouvés, un CSV
            fusionné avec tous les jours sélectionnés, et chaque fichier CSV individuellement.
            """
        )

    with st.expander("3. Fusionner des fichiers"):
        st.markdown(
            """
            Cette section sert à regrouper plusieurs fichiers qui ont la même structure. L'application lit les fichiers
            CSV, XLSX ou XLS, vérifie les en-têtes, empile les lignes, puis génère un fichier unique.

            C'est l'étape à utiliser avant les contrôles LTR ou l'audit quand les données RDA ou Webfleet arrivent en
            plusieurs exports. La fusion ne réinterprète pas les règles métier: elle prépare seulement une base propre
            et unique pour le traitement suivant.
            """
        )

    with st.expander("4. Contrôles LTR"):
        st.markdown(
            """
            Les contrôles LTR transforment d'abord les lignes RDA en **services réels** par collaborateur. Un service
            commence au premier intervalle de travail et continue tant qu'il n'y a pas de repos d'au moins **8 h**. Les
            lignes après minuit restent attachées au service commencé la veille quand elles appartiennent au même service.

            Les pauses sont soustraites du temps de travail. Les prestations transport sont comptées comme travail pour
            les contrôles 50 h, 7 jours, pauses et construction des services.

            Conseil: incluez les données du mois précédent et complétez la dernière semaine du mois avec les premiers
            jours du mois suivant, au moins jusqu'au lundi suivant si nécessaire.
            """
        )
        st.markdown(
            """
            **OVER_50H_WEEK** additionne le temps net par semaine lundi-dimanche. Les services qui passent minuit sont
            coupés en créneaux calendrier pour mettre les heures sur la vraie date.

            **STREAK_7DAYS** cherche les séries de 7 jours de service consécutifs ou plus, en utilisant la date de début
            du service.

            **SPAN_OVER_14H** signale un service dont l'amplitude complète dépasse 14 h.

            **REST_UNDER_11H** compare le début d'un service avec la fin du précédent. Moins de 8 h est une infraction.
            Entre 8 h et 11 h, la réduction est acceptée seulement si c'est la première de la semaine et si la moyenne
            des 14 jours précédents est au moins 11 h; sinon elle est signalée ou mise en revue si l'historique manque.

            **PAUSE_INSUFF** exige 15 min d'interruption dès 5 h 30 de travail, 30 min dès 7 h, et 60 min dès 9 h.
            """
        )

    with st.expander("5. Audit Webfleet-RDA, PDFs et RDA cutting"):
        st.markdown(
            """
            **Audit principal**  
            L'audit croise quatre sources: RDA, Webfleet, mapping collaborateurs et planning. Les dates/heures sont
            normalisées en heure suisse, les identifiants collaborateur sont alignés via le mapping, puis les trajets
            Webfleet sont comparés aux plages RDA et planning du même collaborateur.

            Les flags principaux repèrent les trajets privés pendant un service, les trajets professionnels ou domicile-
            travail après le buffer de fin, avant le buffer de début, sur jour sans RDA, ou dans une grande coupure interne.
            Les trajets avec une vitesse supérieure à 160 km/h sont aussi signalés. Les kilomètres suspects additionnent
            les trajets de mode 2 ou 3 qui tombent hors du cadre attendu.

            **PDFs Gantt**  
            Pendant l'audit complet, les PDFs Gantt sont générés directement dans le même traitement. Le package final
            contient un PDF par collaborateur, avec une page par jour: Webfleet, RDA et planning sont affichés sur trois
            lignes. Les trajets suspects sont colorés en rouge, les prestations RDA gardent leurs couleurs par code, et
            la page affiche aussi les kilomètres mensuels privé, suspect et total privé + suspect.

            **RDA cutting**  
            Le cutting ne modifie que les bords d'une journée RDA: la première et/ou la dernière entrée du jour, jamais
            les prestations au milieu. Si un trajet Webfleet chevauche le début de la première prestation, le nouveau
            début RDA est placé 1 minute après la fin de la chaine de trajets Webfleet. Si un trajet chevauche la fin de
            la dernière prestation, la nouvelle fin RDA est placée 1 minute avant le début de la chaine Webfleet.

            Si toute une ligne RDA est couverte par la coupe, elle est marquée comme supprimée ou mise à zéro. Si une
            prestation 61010 reste en bord de journée après une coupe, elle peut aussi être retirée. Le package final
            contient aussi le RDA ajusté, le résumé des coupes, les trajets Webfleet utilisés et les PDFs des jours modifiés.

            **Manual review**  
            Le mode manual review applique la même détection, mais ne coupe pas directement le RDA. Il génère une liste
            de changements à vérifier et des PDFs où les lignes concernées sont marquées en rouge vif.
            """
        )

def inject_app_css() -> None:
    colors = {
        "teal": "#31b6a7",
        "teal_dark": "#188f83",
        "orange": "#f59b45",
        "yellow": "#ffc400",
        "danger": "#e45f4f",
    }
    st.markdown(
        f"""
        <style>
        :root {{
            --ha-bg: var(--background-color);
            --ha-surface: var(--secondary-background-color);
            --ha-surface-2: var(--background-color);
            --ha-text: var(--text-color);
            --ha-muted: color-mix(in srgb, var(--text-color) 64%, transparent);
            --ha-border: color-mix(in srgb, var(--text-color) 16%, transparent);
            --ha-teal: {colors["teal"]};
            --ha-teal-dark: {colors["teal_dark"]};
            --ha-orange: {colors["orange"]};
            --ha-yellow: {colors["yellow"]};
            --ha-danger: {colors["danger"]};
        }}
        .stApp {{
            background: var(--ha-bg);
            color: var(--ha-text);
        }}
        [data-testid="stHeader"] {{
            background: color-mix(in srgb, var(--ha-bg) 86%, transparent);
        }}
        .block-container {{
            padding-top: 1.35rem;
        }}
        h1, h2, h3, h4, h5, h6, p, label, span {{
            color: var(--ha-text);
        }}
        h1 {{
            letter-spacing: 0;
            font-weight: 800;
        }}
        h2, h3 {{
            font-weight: 750;
        }}
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stCaptionContainer"] {{
            color: var(--ha-muted);
        }}
        div[data-testid="stMetric"],
        div[data-testid="stExpander"],
        [data-testid="stDataFrame"],
        [data-testid="stTabs"] {{
            border-color: var(--ha-border);
        }}
        div[data-testid="stMetric"] {{
            background: var(--ha-surface);
            border: 1px solid var(--ha-border);
            border-radius: 8px;
            padding: 0.75rem 0.9rem;
        }}
        div[data-testid="stMetric"] label {{
            color: var(--ha-muted);
            font-weight: 700;
        }}
        div[data-testid="stMetricValue"] {{
            color: var(--ha-text);
        }}
        div[data-testid="stExpander"] {{
            background: var(--ha-surface);
            border-radius: 8px;
        }}
        .stButton > button,
        .stDownloadButton > button {{
            background: var(--ha-teal);
            color: white;
            border: 1px solid var(--ha-teal);
            border-radius: 6px;
            font-weight: 800;
            letter-spacing: 0;
            min-height: 2.4rem;
            box-shadow: 0 8px 18px rgba(49, 182, 167, 0.18);
        }}
        .stButton > button:hover,
        .stDownloadButton > button:hover {{
            background: var(--ha-teal-dark);
            border-color: var(--ha-teal-dark);
            color: white;
        }}
        .stButton > button[kind="primary"] {{
            background: var(--ha-orange);
            border-color: var(--ha-orange);
            box-shadow: 0 8px 18px rgba(245, 155, 69, 0.20);
        }}
        .stTextInput input,
        .stNumberInput input,
        .stDateInput input,
        .stTextArea textarea,
        div[data-baseweb="select"] > div,
        div[data-baseweb="base-input"] {{
            background: var(--ha-surface);
            color: var(--ha-text);
            border-color: var(--ha-border);
            border-radius: 6px;
        }}
        .stRadio,
        .stCheckbox,
        .stToggle {{
            color: var(--ha-text);
        }}
        div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {{
            border-color: var(--ha-teal);
        }}
        .stTabs [data-baseweb="tab-list"] {{
            gap: 0.85rem;
            padding: 0 0 0.35rem 0;
            margin-bottom: 1rem;
            background: transparent;
            border-bottom: 1px solid var(--ha-border);
            overflow-x: auto;
        }}
        .stTabs [data-baseweb="tab"] {{
            min-height: 2.35rem;
            padding: 0.35rem 0.1rem 0.55rem;
            color: var(--ha-muted);
            background: transparent;
            border: 1px solid transparent;
            border-bottom: 3px solid transparent;
            border-radius: 0;
            font-weight: 750;
            letter-spacing: 0;
            white-space: nowrap;
        }}
        .stTabs [data-baseweb="tab"] p {{
            color: inherit;
            font-weight: inherit;
        }}
        .stTabs [data-baseweb="tab"]:hover {{
            color: var(--ha-text);
            background: transparent;
            border-bottom-color: color-mix(in srgb, var(--ha-teal) 45%, transparent);
        }}
        .stTabs [aria-selected="true"] {{
            color: var(--ha-teal-dark);
            background: transparent;
            border-bottom-color: var(--ha-orange);
            box-shadow: none;
        }}
        .stTabs [data-baseweb="tab-highlight"] {{
            display: none;
        }}
        .side-brand {{
            display: inline-flex;
            align-items: center;
            gap: 0.65rem;
            font-weight: 900;
            color: var(--ha-text);
            margin: 0.15rem 0 0.85rem;
        }}
        .side-brand::before {{
            content: "";
            display: inline-block;
            width: 0.5rem;
            height: 1.5rem;
            background: var(--ha-teal);
            border-radius: 2px;
            box-shadow: 0.65rem 0 0 var(--ha-orange);
            margin-right: 0.65rem;
        }}
        .element-container:has(.stAlert) {{
            color: var(--ha-text);
        }}
        [data-testid="stDataFrame"] {{
            background: var(--ha-surface);
            border-radius: 8px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def main() -> None:
    st.set_page_config(page_title="Application Audit RDA et Trajets", layout="wide")
    inject_app_css()

    main_col, task_col = st.columns([5, 1.35], gap="large")

    with task_col:
        st.markdown('<div class="side-brand">HOME ASSISTANCE</div>', unsafe_allow_html=True)
        task_labels = list(TASKS.values())
        current_label = st.radio(
            "Sélectionner une tâche",
            task_labels,
            label_visibility="collapsed",
            key="selected_task_label",
        )

    selected_key = next(key for key, label in TASKS.items() if label == current_label)

    with main_col:
        if selected_key == "home":
            render_home_task()
        elif selected_key == "webfleet":
            render_webfleet_task()
        elif selected_key == "planning":
            render_planning_download_task()
        elif selected_key == "merge":
            render_merge_task()
        elif selected_key == "ltr":
            render_ltr_task()
        elif selected_key == "audit":
            render_audit_task()


if __name__ == "__main__":
    main()
