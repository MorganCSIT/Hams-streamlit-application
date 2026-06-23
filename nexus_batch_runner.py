import hashlib

from app_config import *
from rda_transfers import (
    RDA_ALLOWED_EXTENSIONS,
    _detect_columns,
    _duration_from_min,
    _nexus_df,
    _read_uploaded_df,
    _to_min,
    _write_batch_file,
)
from ui_common import render_blocking_run_warning, render_download_for_path, safe_folder_name


NEXUS_EXE_NAME = "Asebis.Client.StarterCommand.exe"
NEXUS_IMPORT_TYPE = "ImportLeistungen_CSV"
NEXUS_COLUMNS = [
    "Datum",
    "Von",
    "Bis",
    "Leistungscode",
    "Dauer_verrechnet",
    "OE",
    "KD-Nr",
    "Klient",
    "Einsatzgrund",
    "Mitarbeiter-ID",
]


@dataclass
class PreparedNexusTransfer:
    fingerprint: str
    output_dir: Path
    csv_path: Path
    map_path: Path
    batch_path: Path
    log_path: Path
    nexus_df: pd.DataFrame
    map_df: pd.DataFrame
    exe_path: Path
    oe: str


def _client_executable(folder_text: str) -> Path:
    folder = Path(folder_text.strip().strip('"')).expanduser()
    return folder / NEXUS_EXE_NAME


def _validation_rows(raw_df: pd.DataFrame, cols, nexus_df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    rows: list[dict] = []

    def add(level: str, check: str, value, detail: str) -> None:
        rows.append({"Statut": level, "Contrôle": check, "Résultat": value, "Détail": detail})

    add("OK" if len(raw_df) else "Erreur", "Lignes RDA", len(raw_df), "Le fichier ne doit pas être vide.")

    invalid_dates = int(nexus_df["Datum"].eq("").sum())
    add("OK" if invalid_dates == 0 else "Erreur", "Dates invalides", invalid_dates, "Toutes les dates doivent être lisibles.")

    start_minutes = raw_df[cols.start].apply(_to_min)
    end_minutes = raw_df[cols.end].apply(_to_min)
    invalid_times = int((start_minutes.isna() | end_minutes.isna()).sum())
    add("OK" if invalid_times == 0 else "Erreur", "Heures invalides", invalid_times, "Début et fin sont obligatoires.")

    durations = pd.to_numeric(raw_df[cols.duration], errors="coerce")
    invalid_durations = int(durations.isna().sum())
    negative_durations = int((durations.fillna(0) < 0).sum())
    add("OK" if invalid_durations == 0 else "Erreur", "Durées non numériques", invalid_durations, "La durée doit être numérique.")
    add("OK" if negative_durations == 0 else "Erreur", "Durées négatives", negative_durations, "La durée ne peut pas être négative.")

    calculated = pd.Series(
        [_duration_from_min(start, end) for start, end in zip(start_minutes, end_minutes)],
        index=raw_df.index,
        dtype="float64",
    )
    duration_mismatches = int(
        (
            durations.notna()
            & calculated.notna()
            & (durations.round().astype("Int64") != calculated.round().astype("Int64"))
        ).sum()
    )
    add(
        "OK" if duration_mismatches == 0 else "Erreur",
        "Durée différente de Début/Fin",
        duration_mismatches,
        "Corrigez les lignes avant le transfert.",
    )

    empty_codes = int(nexus_df["Leistungscode"].astype(str).str.strip().eq("").sum())
    add("OK" if empty_codes == 0 else "Erreur", "Prestations vides", empty_codes, "Chaque ligne doit avoir un code prestation.")

    raw_clients = raw_df[cols.client]
    missing_clients = raw_clients.isna() | raw_clients.astype(str).str.strip().eq("")
    numeric_clients = pd.to_numeric(raw_clients, errors="coerce")
    invalid_clients = int((~missing_clients & numeric_clients.isna()).sum())
    expected_kd = numeric_clients.fillna(0).astype(int)
    expected_reason = expected_kd.map(lambda value: 0 if value == 0 else 2)
    client_values_are_correct = bool(
        nexus_df["KD-Nr"].eq(expected_kd).all()
        and nexus_df["Klient"].eq(0).all()
        and nexus_df["Einsatzgrund"].eq(expected_reason).all()
    )
    add(
        "OK" if invalid_clients == 0 else "Erreur",
        "Clients invalides",
        invalid_clients,
        "Une valeur vide est autorisée; une valeur renseignée doit être numérique.",
    )
    add(
        "OK" if client_values_are_correct else "Erreur",
        "Valeurs client Nexus",
        int(missing_clients.sum()),
        "Client absent: 0/0/0. Client présent: KD-Nr/0/2.",
    )

    collaborators = pd.to_numeric(raw_df[cols.collab], errors="coerce")
    invalid_collaborators = int(collaborators.isna().sum())
    add(
        "OK" if invalid_collaborators == 0 else "Erreur",
        "Collaborateurs invalides",
        invalid_collaborators,
        "Les identifiants collaborateur doivent être numériques.",
    )

    duplicates = int(nexus_df.duplicated().sum())
    add("OK" if duplicates == 0 else "Attention", "Lignes dupliquées", duplicates, "Vérifiez les doublons avant le transfert.")
    add("Info", "Minutes totales", int(durations.fillna(0).sum()), "Somme de Durée dans le RDA brut.")

    checks = pd.DataFrame(rows)
    return checks, bool((checks["Statut"] == "Erreur").any())


def _normalize_oe(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _validate_prepared_nexus(nexus_df: pd.DataFrame) -> tuple[pd.DataFrame, bool, str]:
    rows: list[dict] = []

    def add(level: str, check: str, value, detail: str) -> None:
        rows.append({"Statut": level, "Contrôle": check, "Résultat": value, "Détail": detail})

    add("OK" if len(nexus_df) else "Erreur", "Lignes Nexus", len(nexus_df), "Le fichier ne doit pas être vide.")

    dates = pd.to_datetime(nexus_df["Datum"], dayfirst=True, errors="coerce")
    invalid_dates = int(dates.isna().sum())
    add("OK" if invalid_dates == 0 else "Erreur", "Dates invalides", invalid_dates, "Datum doit contenir une date lisible.")

    starts = nexus_df["Von"].apply(_to_min)
    ends = nexus_df["Bis"].apply(_to_min)
    invalid_times = int((starts.isna() | ends.isna()).sum())
    add("OK" if invalid_times == 0 else "Erreur", "Heures invalides", invalid_times, "Von et Bis sont obligatoires.")

    durations = pd.to_numeric(nexus_df["Dauer_verrechnet"], errors="coerce")
    invalid_durations = int(durations.isna().sum())
    negative_durations = int(durations.fillna(0).lt(0).sum())
    calculated = pd.Series(
        [_duration_from_min(start, end) for start, end in zip(starts, ends)],
        index=nexus_df.index,
        dtype="float64",
    )
    duration_mismatches = int(
        (
            durations.notna()
            & calculated.notna()
            & (durations.round().astype("Int64") != calculated.round().astype("Int64"))
        ).sum()
    )
    add("OK" if invalid_durations == 0 else "Erreur", "Durées non numériques", invalid_durations, "Dauer_verrechnet doit être numérique.")
    add("OK" if negative_durations == 0 else "Erreur", "Durées négatives", negative_durations, "La durée ne peut pas être négative.")
    add("OK" if duration_mismatches == 0 else "Erreur", "Durée différente de Von/Bis", duration_mismatches, "Corrigez les lignes avant le transfert.")

    empty_codes = int(nexus_df["Leistungscode"].fillna("").astype(str).str.strip().eq("").sum())
    add("OK" if empty_codes == 0 else "Erreur", "Prestations vides", empty_codes, "Leistungscode est obligatoire.")

    oe_values = sorted({_normalize_oe(value) for value in nexus_df["OE"] if _normalize_oe(value)})
    oe = oe_values[0] if len(oe_values) == 1 else ""
    add(
        "OK" if len(oe_values) == 1 else "Erreur",
        "OE unique",
        oe if oe else len(oe_values),
        "Le fichier préparé doit contenir exactement une OE.",
    )

    kd = pd.to_numeric(nexus_df["KD-Nr"], errors="coerce")
    klient = pd.to_numeric(nexus_df["Klient"], errors="coerce")
    reason = pd.to_numeric(nexus_df["Einsatzgrund"], errors="coerce")
    invalid_client_values = int((kd.isna() | klient.isna() | reason.isna()).sum())
    expected_reason = kd.fillna(0).map(lambda value: 0 if value == 0 else 2)
    inconsistent_clients = int((klient.fillna(-1).ne(0) | reason.fillna(-1).ne(expected_reason)).sum())
    add("OK" if invalid_client_values == 0 else "Erreur", "Valeurs client numériques", invalid_client_values, "KD-Nr, Klient et Einsatzgrund doivent être numériques.")
    add("OK" if inconsistent_clients == 0 else "Erreur", "Règle client", inconsistent_clients, "Sans client: 0/0/0. Avec client: KD-Nr/0/2.")

    collaborators = pd.to_numeric(nexus_df["Mitarbeiter-ID"], errors="coerce")
    invalid_collaborators = int(collaborators.isna().sum())
    add("OK" if invalid_collaborators == 0 else "Erreur", "Collaborateurs invalides", invalid_collaborators, "Mitarbeiter-ID doit être numérique.")

    duplicates = int(nexus_df.duplicated().sum())
    add("OK" if duplicates == 0 else "Attention", "Lignes dupliquées", duplicates, "Vérifiez les doublons avant le transfert.")
    add("Info", "Minutes totales", int(durations.fillna(0).sum()), "Somme de Dauer_verrechnet.")

    checks = pd.DataFrame(rows)
    return checks, bool((checks["Statut"] == "Erreur").any()), oe


def _fingerprint(uploaded_file, oe: str, exe_path: Path) -> str:
    digest = hashlib.sha256()
    digest.update(uploaded_file.getvalue())
    digest.update(oe.encode("utf-8"))
    digest.update(str(exe_path.resolve()).encode("utf-8", errors="replace"))
    return digest.hexdigest()


def _prepare_transfer(
    fingerprint: str,
    source_name: str,
    nexus_df: pd.DataFrame,
    map_df: pd.DataFrame,
    exe_path: Path,
    oe: str,
) -> PreparedNexusTransfer:
    root = get_session_output_root("NexusTransfers")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = root / f"{safe_folder_name(Path(source_name).stem)}_{timestamp}_{fingerprint[:8]}"
    output_dir.mkdir(parents=True, exist_ok=True)

    csv_path = output_dir / "RDA_Nexus.csv"
    map_path = output_dir / "HAS_map.csv"
    batch_path = output_dir / "RDA_Nexus_batch.bat"
    log_path = output_dir / "RDA_Nexus_transfer.log"

    nexus_df.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")
    map_df.to_csv(map_path, index=False, sep=";", encoding="utf-8-sig")
    _write_batch_file(batch_path, csv_path.name, oe, map_path.name, str(exe_path))

    return PreparedNexusTransfer(
        fingerprint=fingerprint,
        output_dir=output_dir,
        csv_path=csv_path,
        map_path=map_path,
        batch_path=batch_path,
        log_path=log_path,
        nexus_df=nexus_df,
        map_df=map_df,
        exe_path=exe_path,
        oe=oe,
    )


def _run_transfer(prepared: PreparedNexusTransfer, username: str, password: str) -> dict:
    args = [
        str(prepared.exe_path),
        f"/u={username}",
        f"/p={password}",
        f"/t={NEXUS_IMPORT_TYPE}",
        f"/o={prepared.oe}",
        f"/f={prepared.csv_path.name}",
        f"/map={prepared.map_path.name}",
        "/v",
    ]
    started = datetime.now()
    try:
        result = subprocess.run(
            args,
            cwd=str(prepared.output_dir),
            capture_output=True,
            text=True,
            errors="replace",
            timeout=1800,
            shell=False,
        )
        ended = datetime.now()
        status = "Réussi" if result.returncode == 0 else "Échec"
        return_code = result.returncode
        stdout = result.stdout or ""
        stderr = result.stderr or ""
    except Exception as exc:
        ended = datetime.now()
        status = "Erreur"
        return_code = ""
        stdout = ""
        stderr = str(exc)

    safe_command = " ".join([str(prepared.exe_path), f"/u={username}", "/p=********", *args[3:]])
    log_text = (
        f"Statut: {status}\n"
        f"Début: {started.isoformat(timespec='seconds')}\n"
        f"Fin: {ended.isoformat(timespec='seconds')}\n"
        f"Durée secondes: {(ended - started).total_seconds():.1f}\n"
        f"Code retour: {return_code}\n"
        f"Commande: {safe_command}\n\n"
        f"STDOUT\n------\n{stdout}\n\nSTDERR\n------\n{stderr}\n"
    )
    prepared.log_path.write_text(log_text, encoding="utf-8")
    return {
        "Statut": status,
        "Code retour": return_code,
        "Début": started.strftime("%Y-%m-%d %H:%M:%S"),
        "Fin": ended.strftime("%Y-%m-%d %H:%M:%S"),
        "Durée sec": round((ended - started).total_seconds(), 1),
        "Log": log_text,
    }


def _render_prepared_downloads(prepared: PreparedNexusTransfer) -> None:
    cols = st.columns(3)
    with cols[0]:
        render_download_for_path(prepared.map_path, "Télécharger HAS_map.csv", key="nexus_raw_map_download", width="stretch")
    with cols[1]:
        render_download_for_path(prepared.csv_path, "Télécharger le CSV Nexus", key="nexus_raw_csv_download", width="stretch")
    with cols[2]:
        render_download_for_path(prepared.batch_path, "Télécharger le batch", key="nexus_raw_batch_download", width="stretch")


def render_nexus_batch_runner_task(embedded: bool = False) -> None:
    if embedded:
        st.subheader("Transfert RDA vers Nexus")
    else:
        st.title("Transfert RDA vers Nexus")
    st.caption("Chargez un RDA brut ou un fichier Nexus déjà préparé, contrôlez les données, puis lancez le transfert.")
    st.info(
        "Le chemin nx-spi-client est lu sur le PC qui exécute Streamlit. Avec un lien partagé, indiquez un chemin "
        "présent sur le serveur Streamlit, pas sur le PC du navigateur."
    )

    input_mode = st.radio(
        "Type de fichier d'entrée",
        ["RDA brut", "Fichier Nexus déjà préparé"],
        horizontal=True,
        key="nexus_input_mode",
    )
    upload_col, path_col = st.columns(2)
    if input_mode == "RDA brut":
        input_file = upload_col.file_uploader("Fichier RDA brut", type=RDA_ALLOWED_EXTENSIONS, key="nexus_raw_rda")
    else:
        input_file = upload_col.file_uploader(
            "Fichier Nexus préparé",
            type=RDA_ALLOWED_EXTENSIONS,
            key="nexus_prepared_file",
            help="Colonnes attendues: Datum, Von, Bis, Leistungscode, Dauer_verrechnet, OE, KD-Nr, Klient, Einsatzgrund, Mitarbeiter-ID.",
        )
    nx_folder_text = path_col.text_input(
        "Chemin local du dossier nx-spi-client",
        placeholder=r"C:\Nexus\nx-spi-client",
        key="nexus_raw_client_folder",
    )
    oe = ""
    if input_mode == "RDA brut":
        uo_label = st.selectbox("UO cible", list(RDA_OE_MAP.keys()), key="nexus_raw_uo")
        oe = RDA_OE_MAP[uo_label]

    exe_path = _client_executable(nx_folder_text) if nx_folder_text.strip() else None
    if nx_folder_text.strip():
        if exe_path and exe_path.is_file():
            st.success(f"Client Nexus trouvé: {exe_path}")
        else:
            st.error(f"Client Nexus introuvable: {exe_path}")

    if input_file is None:
        expected = "un fichier RDA brut" if input_mode == "RDA brut" else "un fichier Nexus déjà préparé"
        st.warning(f"Ajoutez {expected} pour afficher les contrôles.")
        return

    try:
        source_df = _read_uploaded_df(input_file)
        if input_mode == "RDA brut":
            if not any(column in source_df.columns for column in RDA_CLIENT_COLS):
                source_df["N° du client"] = 0
            cols = _detect_columns(source_df)
            nexus_df = _nexus_df(source_df, cols, oe)
            checks, has_blocking_errors = _validation_rows(source_df, cols, nexus_df)
            source_tab_label = "RDA brut"
        else:
            source_df.columns = [str(column).strip() for column in source_df.columns]
            missing_columns = [column for column in NEXUS_COLUMNS if column not in source_df.columns]
            if missing_columns:
                raise ValueError(f"Colonnes Nexus manquantes: {', '.join(missing_columns)}")
            nexus_df = source_df[NEXUS_COLUMNS].copy()
            nexus_df["OE"] = nexus_df["OE"].map(_normalize_oe)
            checks, has_blocking_errors, oe = _validate_prepared_nexus(nexus_df)
            source_tab_label = "Fichier préparé chargé"
    except Exception as exc:
        st.error(f"Impossible de préparer le transfert: {exc}")
        return

    map_codes = sorted(code for code in nexus_df["Leistungscode"].dropna().astype(str).str.strip().unique() if code)
    map_df = pd.DataFrame({"Code_ext": map_codes, "Leistungstarif_nummer": map_codes})

    error_count = int((checks["Statut"] == "Erreur").sum())
    warning_count = int((checks["Statut"] == "Attention").sum())
    metric_cols = st.columns(4)
    metric_cols[0].metric("Lignes", f"{len(nexus_df):,}")
    metric_cols[1].metric("Prestations", f"{len(map_df):,}")
    metric_cols[2].metric("Erreurs", error_count)
    metric_cols[3].metric("Avertissements", warning_count)

    st.subheader("Contrôles avant transfert")
    st.dataframe(checks, width="stretch", hide_index=True)
    preview_tab, nexus_tab, map_tab = st.tabs([source_tab_label, "CSV Nexus utilisé", "HAS_map"])
    with preview_tab:
        st.dataframe(source_df, width="stretch", hide_index=True)
    with nexus_tab:
        st.dataframe(nexus_df, width="stretch", hide_index=True)
    with map_tab:
        st.dataframe(map_df, width="stretch", hide_index=True)

    if has_blocking_errors:
        st.error("Le transfert est bloqué. Corrigez les erreurs indiquées dans le fichier.")
        return
    if exe_path is None or not exe_path.is_file():
        st.warning("Indiquez un dossier nx-spi-client valide pour générer et exécuter le batch.")
        return

    fingerprint = _fingerprint(input_file, f"{input_mode}:{oe}", exe_path)
    prepared = st.session_state.get("nexus_raw_prepared")
    if not isinstance(prepared, PreparedNexusTransfer) or prepared.fingerprint != fingerprint:
        prepared = _prepare_transfer(fingerprint, input_file.name, nexus_df, map_df, exe_path, oe)
        st.session_state["nexus_raw_prepared"] = prepared
        st.session_state.pop("nexus_raw_run_result", None)

    st.success("HAS_map.csv, le CSV Nexus et le batch ont été générés.")
    _render_prepared_downloads(prepared)

    st.subheader("Lancer le transfert")
    credential_cols = st.columns(2)
    username = credential_cols[0].text_input("Utilisateur Nexus", key="nexus_raw_username")
    password = credential_cols[1].text_input("Mot de passe Nexus", type="password", key="nexus_raw_password")
    confirm = st.checkbox(
        "J'ai vérifié les données affichées et je confirme le transfert vers Nexus.",
        key="nexus_raw_confirm",
    )
    run_disabled = not username.strip() or not password or not confirm
    if st.button("Lancer le batch Nexus", type="primary", disabled=run_disabled, width="stretch", key="nexus_raw_run"):
        render_blocking_run_warning()
        with st.spinner("Transfert Nexus en cours..."):
            result = _run_transfer(prepared, username.strip(), password)
        st.session_state["nexus_raw_run_result"] = result

    result = st.session_state.get("nexus_raw_run_result")
    if result:
        if result["Statut"] == "Réussi":
            st.success(f"Transfert terminé avec le code retour {result['Code retour']}.")
        else:
            st.error(f"Transfert terminé avec le statut {result['Statut']} et le code retour {result['Code retour']}.")
        st.dataframe(pd.DataFrame([{key: value for key, value in result.items() if key != "Log"}]), width="stretch", hide_index=True)
        st.subheader("Log du transfert")
        st.code(result["Log"], language="text")
        render_download_for_path(prepared.log_path, "Télécharger le log", key="nexus_raw_log_download")
