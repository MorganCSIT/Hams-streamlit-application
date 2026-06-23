import importlib

import streamlit as st

import app_config as _app_config
import audit_webfleet_rda as _audit_webfleet_rda
import ltr_checks as _ltr_checks
import merge_files as _merge_files
import nexus_batch_runner as _nexus_batch_runner
import rda_transfers as _rda_transfers
import ui_common as _ui_common
import webfleet as _webfleet

_app_config = importlib.reload(_app_config)
_ui_common = importlib.reload(_ui_common)
_webfleet = importlib.reload(_webfleet)
_merge_files = importlib.reload(_merge_files)
_nexus_batch_runner = importlib.reload(_nexus_batch_runner)
_rda_transfers = importlib.reload(_rda_transfers)
_ltr_checks = importlib.reload(_ltr_checks)
_audit_webfleet_rda = importlib.reload(_audit_webfleet_rda)

TASKS = _app_config.TASKS
render_audit_task = _audit_webfleet_rda.render_audit_task
render_ltr_task = _ltr_checks.render_ltr_task
render_merge_task = _merge_files.render_merge_task
render_nexus_batch_runner_task = _nexus_batch_runner.render_nexus_batch_runner_task
render_rda_task = _rda_transfers.render_rda_task
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

    _ui_common.render_template_downloads()

    st.subheader("Utiliser les batchs Nexus pour les transferts RDA")
    st.markdown(
        """
        Pour exécuter les batchs, téléchargez d'abord le dossier complet généré par la section **Transferts RDA**.
        Décompressez le fichier `.zip`, puis placez le dossier `nx-spi-client` dans le dossier RDA extrait, au même niveau que les dossiers `01_...`, `02_...`, `03_...` et les fichiers `HAS_map...csv`.
        
        [Télécharger le dossier nx-spi-client depuis le serveur](https://cloud.nexus-schweiz.ch/index.php/s/xokmebwanSH7DWH).
        
        Le dossier nx-spi-client ne peut pas être partagé entre plusieurs machines. Il doit être téléchargé directement depuis ce lien du serveur afin que les transferts RDA fonctionnent correctement. Veuillez demander le mot de passe à un administrateur.

        Exemple attendu :
        ```text
        RDA_xxxxx/
          nx-spi-client/
            Asebis.Client.StarterCommand.exe
          HAS_map_main.csv
          01_Standard_Transfer/
          01_All_Collabs_One_CSV/
          02_Collabs_With_61010_One_CSV/
          03_Per_Collab_Separate/
        ```

        Si le dossier `nx-spi-client` n'est pas téléchargé depuis le lien du serveur puis placé à cet endroit, les fichiers `.bat` ne trouveront pas `Asebis.Client.StarterCommand.exe` et le transfert Nexus échouera.
        """
    )

    st.subheader("Fichiers et dossiers RDA")
    rda_files = [
        {
            "Fichier / dossier": "Dossier complet .zip",
            "Rôle": "Archive à télécharger puis décompresser. Elle contient les CSV Nexus, batchs, mappings et contrôles.",
        },
        {
            "Fichier / dossier": "nx-spi-client/",
            "Rôle": "Client Nexus à ajouter manuellement dans le dossier RDA extrait pour pouvoir lancer les `.bat`.",
        },
        {
            "Fichier / dossier": "*.bat",
            "Rôle": "Fichier à lancer pour importer le CSV associé dans Nexus avec le bon OE et le bon mapping.",
        },
        {
            "Fichier / dossier": "*.csv dans les dossiers 01/02/03",
            "Rôle": "Données d'import Nexus. Le batch du même dossier utilise ce CSV.",
        },
        {
            "Fichier / dossier": "HAS_map.csv / HAS_map_main.csv",
            "Rôle": "Mapping des codes de prestations utilisé par les batchs pendant l'import Nexus.",
        },
        {
            "Fichier / dossier": "01_Standard_Transfer/",
            "Rôle": "Transfert standard UO vers UO, avec un CSV et un batch principal.",
        },
        {
            "Fichier / dossier": "01_All_Collabs_One_CSV/",
            "Rôle": "Import groupé de tous les collaborateurs dans un seul CSV.",
        },
        {
            "Fichier / dossier": "02_Collabs_With_61010_One_CSV/",
            "Rôle": "Import groupé limité aux collaborateurs qui avaient des prestations 61010.",
        },
        {
            "Fichier / dossier": "03_Per_Collab_Separate/",
            "Rôle": "Imports séparés par collaborateur, chacun avec son CSV et son batch.",
        },
        {
            "Fichier / dossier": "RDA_*_Source.xlsx / *_adjusted_*.xlsx",
            "Rôle": "Fichiers de référence ou de contrôle avant/après transformation RDA.",
        },
        {
            "Fichier / dossier": "RDA_duree_check.csv, QA, audit, overlaps",
            "Rôle": "Contrôles de cohérence pour vérifier les durées, lignes générées et éventuels chevauchements.",
        },
    ]
    st.dataframe(rda_files, use_container_width=True, hide_index=True)

    st.subheader("Obtenir et activer l'accès API Webfleet")
    st.markdown(
        """
        Pour utiliser la section **Téléchargement Webfleet**, l'utilisateur doit avoir un compte Webfleet et une clé API. Cette clé API n'est pas créée dans l'application : elle doit être demandée à Webfleet Support par email.

        Envoyez la demande à `support.de@webfleet.com`. Le message peut reprendre le modèle ci-dessous. Chaque ligne explique directement quelle information mettre :

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

        Après la réponse de Webfleet avec la clé API, le compte maître Webfleet doit se connecter à Webfleet, ouvrir les paramètres de l'utilisateur concerné, puis activer l'utilisation de l'API pour cet utilisateur. Sans cette activation côté utilisateur, le téléchargement Webfleet ne fonctionnera pas même si la clé API existe.
        """
    )

    st.divider()

    st.subheader("Sections disponibles")
    sections = [
        {
            "Section": "Téléchargement Webfleet",
            "Entrées": "Identifiants Webfleet, clé API, période",
            "Sorties": "CSV et Excel des trajets Webfleet",
        },
        {
            "Section": "Fusionner des fichiers",
            "Entrées": "Au moins deux fichiers CSV/XLSX/XLS avec les mêmes en-têtes",
            "Sorties": "Fichier fusionné CSV ou XLSX",
        },
        {
            "Section": "Transferts RDA",
            "Entrées": "RDA brut ou Nexus préparé, mapping UO, chemin nx-spi-client et identifiants Nexus",
            "Sorties": "Fichiers Nexus, contrôles QA, exécution du batch et log de transfert",
        },
        {
            "Section": "Contrôles LTR",
            "Entrées": "Classeur collaborateurs matchés et fichier RDA fusionné",
            "Sorties": "Classeur Excel LTR multi-feuilles",
        },
        {
            "Section": "Audit Webfleet-RDA",
            "Entrées": "Fichiers RDA, Webfleet, Mapping et Planning",
            "Sorties": "Rapport Excel d'audit et PDFs Gantt générés en arrière-plan",
        },
    ]
    st.dataframe(sections, use_container_width=True, hide_index=True)

    st.warning(
        "Pendant un traitement en cours, ne changez pas de section et n'utilisez pas d'autres parties de l'application. "
        "Attendez que le traitement soit terminé."
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
    st.set_page_config(page_title="Outils Webfleet", layout="wide")
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
        elif selected_key == "merge":
            render_merge_task()
        elif selected_key == "rda":
            render_rda_task(render_nexus_batch_runner_task)
        elif selected_key == "ltr":
            render_ltr_task()
        elif selected_key == "audit":
            render_audit_task()


if __name__ == "__main__":
    main()
