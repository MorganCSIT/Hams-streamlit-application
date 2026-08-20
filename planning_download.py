from app_config import *
from ui_common import read_csv_flex, render_blocking_run_warning, render_download_for_path

import posixpath

try:
    import paramiko
except ImportError:
    paramiko = None


DEFAULT_SFTP_HOST = "sftp.nx-schweiz.ch"
DEFAULT_SFTP_PORT = 22
DEFAULT_SFTP_USERNAME = "spi-has"
REMOTE_ROOT = "/"
PLANNING_DOWNLOAD_FOLDER_NAME = "Webfleet Planning Downloads"
DEFAULT_PLANNING_PREFIXES = ("HAS_HAM", "HAS_MCT")
PLANNING_FILENAME_RE = re.compile(
    r"^(?P<prefix>HAS_[A-Z]+)_(?P<snapshot_date>\d{8})_(?P<time>\d{4})_PEPS-Visits\.csv$"
)


@dataclass(frozen=True)
class PlanningFile:
    prefix: str
    snapshot_date: str
    time: str
    name: str
    path: Path | None = None

    @property
    def planning_date(self) -> date:
        return snapshot_date_to_planning_date(self.snapshot_date)


def get_planning_archive_folder() -> Path:
    folder = Path.home() / "Downloads" / PLANNING_DOWNLOAD_FOLDER_NAME
    folder.mkdir(parents=True, exist_ok=True)
    return folder


def get_planning_archive_folders() -> tuple[Path, ...]:
    return (get_planning_archive_folder(),)


def get_planning_source_options(local_files: dict[tuple[str, str], PlanningFile]) -> list[str]:
    prefixes = {key[0] for key in local_files}
    prefixes.update(DEFAULT_PLANNING_PREFIXES)
    return sorted(prefixes)


def parse_planning_filename(filename: str) -> PlanningFile | None:
    match = PLANNING_FILENAME_RE.match(Path(filename).name)
    if not match:
        return None
    return PlanningFile(
        prefix=match.group("prefix"),
        snapshot_date=match.group("snapshot_date"),
        time=match.group("time"),
        name=Path(filename).name,
    )


def snapshot_date_to_planning_date(snapshot_date: str) -> date:
    return datetime.strptime(snapshot_date, "%Y%m%d").date() - timedelta(days=1)


def planning_range_to_snapshot_dates(start: date, end: date) -> list[str]:
    if start > end:
        raise ValueError("La date de début doit être antérieure ou égale à la date de fin.")

    days = []
    current = start
    while current <= end:
        days.append((current + timedelta(days=1)).strftime("%Y%m%d"))
        current += timedelta(days=1)
    return days


def choose_latest_files(files: list[PlanningFile]) -> dict[tuple[str, str], PlanningFile]:
    latest = {}
    for item in files:
        key = (item.prefix, item.snapshot_date)
        current = latest.get(key)
        if current is None or item.time > current.time:
            latest[key] = item
    return latest


def scan_local_planning_files(archive_folder: Path, snapshot_dates: set[str] | None = None) -> dict[tuple[str, str], PlanningFile]:
    archive_folder.mkdir(parents=True, exist_ok=True)

    files = []
    for path in archive_folder.glob("HAS_*_*_PEPS-Visits.csv"):
        parsed = parse_planning_filename(path.name)
        if parsed is None:
            continue
        if snapshot_dates is not None and parsed.snapshot_date not in snapshot_dates:
            continue
        files.append(
            PlanningFile(
                prefix=parsed.prefix,
                snapshot_date=parsed.snapshot_date,
                time=parsed.time,
                name=parsed.name,
                path=path,
            )
        )
    return choose_latest_files(files)


def scan_all_local_planning_files(snapshot_dates: set[str] | None = None) -> dict[tuple[str, str], PlanningFile]:
    files = []
    for folder in get_planning_archive_folders():
        files.extend(scan_local_planning_files(folder, snapshot_dates).values())
    return choose_latest_files(files)


def find_remote_planning_files(filenames: list[str], snapshot_dates: set[str]) -> dict[tuple[str, str], PlanningFile]:
    files = []
    for filename in filenames:
        parsed = parse_planning_filename(filename)
        if parsed is None or parsed.snapshot_date not in snapshot_dates:
            continue
        files.append(parsed)
    return choose_latest_files(files)


def list_remote_planning_files(filenames: list[str]) -> list[PlanningFile]:
    files = []
    for filename in filenames:
        parsed = parse_planning_filename(filename)
        if parsed is not None:
            files.append(parsed)
    return sorted(files, key=lambda item: (item.snapshot_date, item.prefix, item.time, item.name), reverse=True)


def planning_files_to_rows(files: list[PlanningFile]) -> list[dict[str, str]]:
    return [
        {
            "prefix": item.prefix,
            "planning_date": item.planning_date.isoformat(),
            "snapshot_date": item.snapshot_date,
            "time": item.time,
            "file": item.name,
        }
        for item in files
    ]


def is_placeholder_or_empty(value: str, placeholders: set[str]) -> bool:
    normalized = (value or "").strip()
    return not normalized or normalized in placeholders


def sftp_fields_ready(host: str, port: int, username: str, password: str) -> bool:
    return not any(
        [
            is_placeholder_or_empty(host, {"sftp.example.com"}),
            int(port) <= 0,
            is_placeholder_or_empty(username, {"USERNAME_PLACEHOLDER"}),
            is_placeholder_or_empty(password, {"PASSWORD_PLACEHOLDER"}),
        ]
    )


def list_sftp_names(host: str, username: str, password: str, remote_folder: str, port: int = 22) -> list[str]:
    if paramiko is None:
        raise RuntimeError("Paramiko n'est pas installé. Ajoutez-le aux dépendances puis relancez l'application.")

    transport = paramiko.Transport((host, port))
    try:
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        try:
            return sftp.listdir(remote_folder)
        finally:
            sftp.close()
    finally:
        transport.close()


def download_sftp_file(
    host: str,
    username: str,
    password: str,
    remote_folder: str,
    remote_name: str,
    local_folder: Path,
    port: int = 22,
) -> Path:
    if paramiko is None:
        raise RuntimeError("Paramiko n'est pas installé. Ajoutez-le aux dépendances puis relancez l'application.")

    local_folder.mkdir(parents=True, exist_ok=True)
    local_path = local_folder / remote_name
    remote_path = posixpath.join(remote_folder, remote_name)

    transport = paramiko.Transport((host, port))
    try:
        transport.connect(username=username, password=password)
        sftp = paramiko.SFTPClient.from_transport(transport)
        try:
            sftp.get(remote_path, str(local_path))
        finally:
            sftp.close()
    finally:
        transport.close()
    return local_path


def create_zip_bytes(paths: list[Path]) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            archive.write(path, arcname=path.name)
    buffer.seek(0)
    return buffer.getvalue()


def create_merged_planning_csv_bytes(paths: list[Path]) -> bytes:
    frames = []
    for path in paths:
        df = read_csv_flex(path)
        parsed = parse_planning_filename(path.name)
        if parsed is not None:
            df.insert(0, "source_file", path.name)
            df.insert(1, "planning_date", parsed.planning_date.isoformat())
        frames.append(df)

    if not frames:
        return b""

    merged = pd.concat(frames, ignore_index=True, sort=False)
    return merged.to_csv(index=False, encoding="utf-8-sig", sep=";").encode("utf-8-sig")


def render_planning_download_task() -> None:
    st.title("Téléchargement planning")
    st.caption("Télécharge les fichiers planning CSV depuis le serveur SFTP.")
    st.info("Note: la date dans le nom du fichier correspond à la date planning + 1 jour.")

    today = date.today()
    date_cols = st.columns(2)
    from_planning_date = date_cols[0].date_input("From planning date", value=today)
    to_planning_date = date_cols[1].date_input("To planning date", value=today)

    with st.expander("Serveur SFTP", expanded=True):
        server_cols = st.columns(4)
        host = server_cols[0].text_input("SFTP host", value=DEFAULT_SFTP_HOST)
        port = server_cols[1].number_input("SFTP port", min_value=1, max_value=65535, value=DEFAULT_SFTP_PORT)
        username = server_cols[2].text_input("SFTP username", value=DEFAULT_SFTP_USERNAME)
        password = server_cols[3].text_input("SFTP password", type="password")
        st.caption(f"Dossier distant utilisé automatiquement: {REMOTE_ROOT}.")

    if from_planning_date > to_planning_date:
        st.error("La date de début doit être antérieure ou égale à la date de fin.")
        return

    snapshot_dates = planning_range_to_snapshot_dates(from_planning_date, to_planning_date)
    st.caption("Dates recherchées dans les noms de fichiers: " + ", ".join(snapshot_dates))

    selected_prefixes = st.multiselect(
        "Sources",
        list(DEFAULT_PLANNING_PREFIXES),
        default=list(DEFAULT_PLANNING_PREFIXES),
        help="Filtre optionnel selon les préfixes à chercher sur le serveur.",
    )

    if st.button("Find / update files", type="primary", width="stretch"):
        render_blocking_run_warning()
        st.session_state.pop("planning_all_server_zip", None)
        download_folder = get_planning_archive_folder()
        progress = st.progress(0, text="Préparation de la recherche planning...")
        status_box = st.empty()

        wanted_dates = set(snapshot_dates)
        wanted_prefixes = set(selected_prefixes)
        downloaded = []
        errors = []
        remote_latest = {}
        server_total_count = 0
        server_planning_files = []
        progress.progress(10, text="Préparation de la connexion serveur...")
        st.caption(f"Dossier de téléchargement créé automatiquement: {download_folder}")

        ready_for_sftp = sftp_fields_ready(host, int(port), username, password)
        if ready_for_sftp:
            if paramiko is None:
                errors.append(
                    {
                        "file": "SFTP",
                        "error": "Paramiko n'est pas installé dans l'application. Le téléchargement SFTP ne peut pas démarrer.",
                    }
                )
            else:
                try:
                    status_box.info("Connexion au serveur SFTP et lecture de la liste des fichiers...")
                    progress.progress(25, text="Connexion au serveur SFTP...")
                    remote_names = list_sftp_names(host.strip(), username.strip(), password, REMOTE_ROOT, int(port))
                    server_total_count = len(remote_names)
                    server_planning_files = list_remote_planning_files(remote_names)
                    progress.progress(
                        45,
                        text=(
                            f"{server_total_count} fichier(s) lu(s) sur le serveur, "
                            f"{len(server_planning_files)} fichier(s) planning reconnu(s)."
                        ),
                    )
                    remote_latest = find_remote_planning_files(remote_names, wanted_dates)
                    if wanted_prefixes:
                        remote_latest = {key: value for key, value in remote_latest.items() if key[0] in wanted_prefixes}

                    files_to_download = [
                        (key, remote_file)
                        for key, remote_file in sorted(remote_latest.items(), key=lambda item: (item[0][1], item[0][0]))
                    ]
                    if files_to_download:
                        status_box.info(f"Téléchargement SFTP en cours: 0/{len(files_to_download)} fichier(s)...")
                    else:
                        status_box.info("Aucun fichier serveur disponible pour cette sélection.")
                        progress.progress(80, text="Aucun fichier serveur disponible pour cette sélection.")

                    for index, (key, remote_file) in enumerate(files_to_download, start=1):
                        try:
                            status_box.info(
                                f"Téléchargement SFTP en cours: {index}/{len(files_to_download)} - {remote_file.name}"
                            )
                            if files_to_download:
                                percent = 50 + int((index - 1) / len(files_to_download) * 40)
                                progress.progress(percent, text=f"Téléchargement: {remote_file.name}")
                            local_path = download_sftp_file(
                                host.strip(),
                                username.strip(),
                                password,
                                REMOTE_ROOT,
                                remote_file.name,
                                download_folder,
                                int(port),
                            )
                            downloaded.append(
                                PlanningFile(
                                    prefix=remote_file.prefix,
                                    snapshot_date=remote_file.snapshot_date,
                                    time=remote_file.time,
                                    name=remote_file.name,
                                    path=local_path,
                                )
                            )
                            progress.progress(
                                50 + int(index / len(files_to_download) * 40),
                                text=f"Téléchargé: {remote_file.name}",
                            )
                        except Exception as exc:
                            errors.append({"file": remote_file.name, "error": str(exc)})
                except Exception as exc:
                    errors.append({"file": "SFTP", "error": str(exc)})
        else:
            st.warning("Les 4 champs SFTP doivent être renseignés avec de vraies valeurs pour chercher sur le serveur.")

        available = sorted(downloaded, key=lambda item: (item.snapshot_date, item.prefix))
        available_keys = set(remote_latest)
        candidate_prefixes = wanted_prefixes or set(DEFAULT_PLANNING_PREFIXES)
        missing = []
        for snapshot_date in snapshot_dates:
            for prefix in sorted(candidate_prefixes):
                if (prefix, snapshot_date) not in available_keys:
                    missing.append(
                        {
                            "prefix": prefix,
                            "planning_date": snapshot_date_to_planning_date(snapshot_date).isoformat(),
                            "snapshot_date": snapshot_date,
                        }
                    )

        st.session_state["planning_download_result"] = {
            "downloaded": [item.name for item in downloaded],
            "missing": missing,
            "errors": errors,
            "available": [str(item.path) for item in available if item.path is not None],
            "server_total_count": server_total_count,
            "server_planning_count": len(server_planning_files),
            "requested_server_files": planning_files_to_rows(
                sorted(remote_latest.values(), key=lambda item: (item.snapshot_date, item.prefix, item.time, item.name))
            ),
            "server_file_names": [item.name for item in server_planning_files],
            "from_planning_date": from_planning_date.isoformat(),
            "to_planning_date": to_planning_date.isoformat(),
        }
        progress.progress(100, text="Recherche planning terminée.")
        if errors:
            status_box.error("Recherche terminée avec erreur(s). Voir le résumé ci-dessous.")
        else:
            status_box.success("Recherche terminée. Les fichiers demandés sont prêts au téléchargement.")

    result = st.session_state.get("planning_download_result")
    if not result:
        return

    st.subheader("Résumé")
    metric_cols = st.columns(4)
    metric_cols[0].metric("Fichiers serveur", result.get("server_total_count", 0))
    metric_cols[1].metric("Planning serveur", result.get("server_planning_count", 0))
    metric_cols[2].metric("Téléchargés serveur", len(result["downloaded"]))
    metric_cols[3].metric("Manquants", len(result["missing"]))

    if result.get("requested_server_files"):
        st.subheader("Fichiers demandés trouvés sur le serveur")
        st.dataframe(pd.DataFrame(result["requested_server_files"]), width="stretch", hide_index=True)
    elif not result["errors"]:
        st.info("Aucun fichier planning correspondant aux dates sélectionnées n'a été trouvé sur le serveur.")

    if result["missing"]:
        st.warning("Fichiers demandés non disponibles sur le serveur.")
        st.dataframe(pd.DataFrame(result["missing"]), width="stretch", hide_index=True)

    if result["errors"]:
        st.error("Certaines recherches ou téléchargements ont échoué.")
        st.dataframe(pd.DataFrame(result["errors"]), width="stretch", hide_index=True)

    server_file_names = result.get("server_file_names", [])
    if server_file_names:
        st.subheader("Archive serveur complète")
        all_server_zip = st.session_state.get("planning_all_server_zip")
        if st.button("Préparer tout fishier dans le server", width="stretch"):
            if not sftp_fields_ready(host, int(port), username, password):
                st.warning("Les 4 champs SFTP doivent être renseignés pour préparer le ZIP complet du serveur.")
            else:
                all_server_folder = get_planning_archive_folder()
                all_server_paths = []
                all_server_errors = []
                full_progress = st.progress(0, text="Préparation du ZIP complet du serveur...")
                for index, remote_name in enumerate(server_file_names, start=1):
                    try:
                        local_path = download_sftp_file(
                            host.strip(),
                            username.strip(),
                            password,
                            REMOTE_ROOT,
                            remote_name,
                            all_server_folder,
                            int(port),
                        )
                        all_server_paths.append(local_path)
                        full_progress.progress(
                            int(index / len(server_file_names) * 100),
                            text=f"Ajout au ZIP serveur: {index}/{len(server_file_names)}",
                        )
                    except Exception as exc:
                        all_server_errors.append({"file": remote_name, "error": str(exc)})

                if all_server_paths:
                    all_server_zip = create_zip_bytes(all_server_paths)
                    st.session_state["planning_all_server_zip"] = all_server_zip
                if all_server_errors:
                    st.error("Certains fichiers du serveur n'ont pas pu être ajoutés au ZIP.")
                    st.dataframe(pd.DataFrame(all_server_errors), width="stretch", hide_index=True)

        if all_server_zip:
            st.download_button(
                "Télécharger tout fishier dans le server",
                all_server_zip,
                file_name="tout fishier dans le server.zip",
                mime="application/zip",
                key="planning_all_server_zip_download",
            )

    available_paths = [Path(path) for path in result["available"] if Path(path).is_file()]
    if available_paths:
        st.subheader("Téléchargements")
        if len(available_paths) > 1:
            st.download_button(
                "Télécharger tous les fichiers ZIP",
                create_zip_bytes(available_paths),
                file_name="planning_files.zip",
                mime="application/zip",
                key="planning_download_zip",
            )

        st.download_button(
            "Télécharger le planning fusionné CSV",
            create_merged_planning_csv_bytes(available_paths),
            file_name=f"planning_merged_{result['from_planning_date']}_to_{result['to_planning_date']}.csv",
            mime="text/csv",
            key="planning_download_merged_csv",
        )

        for path in available_paths:
            render_download_for_path(path, f"Télécharger {path.name}", key=f"planning_download_{path.name}")
