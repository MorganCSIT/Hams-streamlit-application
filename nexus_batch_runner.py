from app_config import *
from ui_common import render_blocking_run_warning, safe_folder_name


BATCH_GROUPS = [
    "01_Standard_Transfer",
    "01_All_Collabs_One_CSV",
    "02_Collabs_With_61010_One_CSV",
    "02_Whitelisted_Ready_For_101",
    "03_Per_Collab_Separate",
]


@dataclass
class BatchCandidate:
    batch_path: Path
    group: str
    exe_text: str
    exe_path: Path | None
    csv_text: str
    csv_path: Path | None
    map_text: str
    map_path: Path | None
    oe: str
    import_type: str
    args: list[str]
    status: str
    reason: str

    @property
    def runnable(self) -> bool:
        return self.status == "Runnable"


def _latest_rda_folder() -> str:
    root = APP_ROOT / RDA_OUTPUT_FOLDER
    if not root.exists():
        return ""
    folders = [path for path in root.iterdir() if path.is_dir() and path.name != "BatchRunLogs"]
    if not folders:
        return ""
    return str(max(folders, key=lambda path: path.stat().st_mtime))


def _resolve_from(base: Path, value: str) -> Path | None:
    if not value:
        return None
    raw = Path(value)
    if raw.is_absolute():
        return raw
    return (base / raw).resolve()


def _strip_quotes(value: str) -> str:
    text = value.strip()
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        return text[1:-1]
    return text


def _extract_params(args_text: str) -> tuple[dict[str, str], list[str]]:
    params: dict[str, str] = {}
    args: list[str] = []
    pattern = re.compile(r"/(?P<key>u|p|t|o|f|map)=(?P<value>\"[^\"]*\"|\S+)|(?P<verbose>/v)\b", re.IGNORECASE)
    for match in pattern.finditer(args_text):
        if match.group("verbose"):
            args.append("/v")
            continue
        key = match.group("key").lower()
        value = _strip_quotes(match.group("value"))
        params[key] = value
        args.append(f"/{key}={value}")
    return params, args


def _parse_batch(batch_path: Path) -> BatchCandidate:
    try:
        text = batch_path.read_text(encoding="utf-8", errors="replace")
    except Exception as exc:
        return BatchCandidate(batch_path, _batch_group(batch_path), "", None, "", None, "", None, "", "", [], "Blocked", f"Lecture impossible: {exc}")

    command_line = next((line.strip() for line in text.splitlines() if "Asebis.Client.StarterCommand.exe" in line), "")
    if not command_line:
        return BatchCandidate(batch_path, _batch_group(batch_path), "", None, "", None, "", None, "", "", [], "Blocked", "Commande Nexus introuvable")

    quoted = re.match(r'^"(?P<exe>[^"]*Asebis\.Client\.StarterCommand\.exe)"\s*(?P<args>.*)$', command_line, re.IGNORECASE)
    unquoted = re.match(r"^(?P<exe>\S*Asebis\.Client\.StarterCommand\.exe)\s*(?P<args>.*)$", command_line, re.IGNORECASE)
    match = quoted or unquoted
    if not match:
        return BatchCandidate(batch_path, _batch_group(batch_path), "", None, "", None, "", None, "", "", [], "Blocked", "Format de commande non reconnu")

    exe_text = match.group("exe")
    params, nexus_args = _extract_params(match.group("args"))
    exe_path = _resolve_from(batch_path.parent, exe_text)
    csv_path = _resolve_from(batch_path.parent, params.get("f", ""))
    map_path = _resolve_from(batch_path.parent, params.get("map", ""))
    args = [str(exe_path)] + nexus_args if exe_path is not None else []

    missing = []
    if exe_path is None or not exe_path.is_file():
        missing.append("nx-spi-client/Asebis.Client.StarterCommand.exe")
    if csv_path is None or not csv_path.is_file():
        missing.append(params.get("f", "CSV"))
    if map_path is None or not map_path.is_file():
        missing.append(params.get("map", "map"))
    if not params.get("o"):
        missing.append("OE")
    if not params.get("t"):
        missing.append("type import")

    status = "Blocked" if missing else "Runnable"
    reason = f"Manquant: {', '.join(missing)}" if missing else ""
    return BatchCandidate(
        batch_path=batch_path,
        group=_batch_group(batch_path),
        exe_text=exe_text,
        exe_path=exe_path,
        csv_text=params.get("f", ""),
        csv_path=csv_path,
        map_text=params.get("map", ""),
        map_path=map_path,
        oe=params.get("o", ""),
        import_type=params.get("t", ""),
        args=args,
        status=status,
        reason=reason,
    )


def _batch_group(batch_path: Path) -> str:
    for part in batch_path.parts:
        if part in BATCH_GROUPS:
            return part
    return "Autre"


def _candidate_rows(candidates: list[BatchCandidate], package_root: Path) -> list[dict]:
    rows = []
    for idx, candidate in enumerate(candidates):
        rows.append(
            {
                "ID": idx,
                "Statut": candidate.status,
                "Groupe": candidate.group,
                "Batch": str(candidate.batch_path.relative_to(package_root)) if candidate.batch_path.is_relative_to(package_root) else str(candidate.batch_path),
                "CSV": candidate.csv_text,
                "Map": candidate.map_text,
                "OE": candidate.oe,
                "Type": candidate.import_type,
                "Raison": candidate.reason,
            }
        )
    return rows


def _scan_batches(package_root: Path) -> list[BatchCandidate]:
    if not package_root.exists() or not package_root.is_dir():
        return []
    return [_parse_batch(path) for path in sorted(package_root.rglob("*.bat"))]


def _safe_extract_zip(uploaded_file) -> Path:
    extract_root = APP_ROOT / RDA_OUTPUT_FOLDER / "UploadedBatchRuns"
    extract_root.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = extract_root / f"{safe_folder_name(Path(uploaded_file.name).stem)}_{timestamp}"
    target.mkdir(parents=True, exist_ok=False)
    target_resolved = target.resolve()

    with zipfile.ZipFile(BytesIO(uploaded_file.getvalue())) as archive:
        for member in archive.infolist():
            destination = (target / member.filename).resolve()
            if not str(destination).lower().startswith(str(target_resolved).lower()):
                raise ValueError(f"Chemin dangereux dans le zip: {member.filename}")
            if member.is_dir():
                destination.mkdir(parents=True, exist_ok=True)
                continue
            destination.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(member) as src, destination.open("wb") as dst:
                shutil.copyfileobj(src, dst)
    return target


def _mask_arg(arg: str) -> str:
    return "/p=********" if arg.lower().startswith("/p=") else arg


def _log_text(candidate: BatchCandidate, result: subprocess.CompletedProcess, started: datetime, ended: datetime) -> str:
    command = " ".join(_mask_arg(arg) for arg in candidate.args)
    return (
        f"Batch: {candidate.batch_path}\n"
        f"Started: {started.isoformat(timespec='seconds')}\n"
        f"Ended: {ended.isoformat(timespec='seconds')}\n"
        f"Duration seconds: {(ended - started).total_seconds():.1f}\n"
        f"Exit code: {result.returncode}\n"
        f"Command: {command}\n\n"
        "STDOUT\n"
        "------\n"
        f"{result.stdout or ''}\n\n"
        "STDERR\n"
        "------\n"
        f"{result.stderr or ''}\n"
    )


def _run_batches(candidates: list[BatchCandidate], package_root: Path) -> tuple[pd.DataFrame, Path]:
    log_root = APP_ROOT / RDA_OUTPUT_FOLDER / "BatchRunLogs"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_log_dir = log_root / f"{safe_folder_name(package_root.name)}_{timestamp}"
    run_log_dir.mkdir(parents=True, exist_ok=True)

    rows = []
    progress = st.progress(0)
    status_box = st.empty()
    for index, candidate in enumerate(candidates, start=1):
        status_box.info(f"Exécution {index}/{len(candidates)}: {candidate.batch_path.name}")
        started = datetime.now()
        try:
            result = subprocess.run(
                candidate.args,
                cwd=str(candidate.batch_path.parent),
                capture_output=True,
                text=True,
                timeout=None,
            )
            ended = datetime.now()
            outcome = "Réussi" if result.returncode == 0 else "Échec"
            log_content = _log_text(candidate, result, started, ended)
        except Exception as exc:
            ended = datetime.now()
            outcome = "Erreur"
            log_content = (
                f"Batch: {candidate.batch_path}\n"
                f"Started: {started.isoformat(timespec='seconds')}\n"
                f"Ended: {ended.isoformat(timespec='seconds')}\n"
                f"Error: {exc}\n"
            )
            result = None

        log_path = run_log_dir / f"{index:03d}_{safe_folder_name(candidate.batch_path.stem)}.log"
        log_path.write_text(log_content, encoding="utf-8")
        rows.append(
            {
                "Batch": str(candidate.batch_path.relative_to(package_root)) if candidate.batch_path.is_relative_to(package_root) else str(candidate.batch_path),
                "Groupe": candidate.group,
                "Statut": outcome,
                "Code retour": "" if result is None else result.returncode,
                "Début": started.strftime("%Y-%m-%d %H:%M:%S"),
                "Fin": ended.strftime("%Y-%m-%d %H:%M:%S"),
                "Durée sec": round((ended - started).total_seconds(), 1),
                "Log": str(log_path),
            }
        )
        progress.progress(index / len(candidates))

    summary = pd.DataFrame(rows)
    summary.to_csv(run_log_dir / "summary.csv", index=False, encoding="utf-8-sig", sep=";")
    status_box.success(f"Exécution terminée. Logs: {run_log_dir}")
    return summary, run_log_dir


def _render_run_results(summary: pd.DataFrame, log_dir: Path) -> None:
    st.subheader("Résultats")
    st.caption(f"Dossier de logs: {log_dir}")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    for _, row in summary.iterrows():
        log_path = Path(row["Log"])
        with st.expander(f"{row['Statut']} - {row['Batch']}"):
            if log_path.is_file():
                st.code(log_path.read_text(encoding="utf-8", errors="replace"), language="text")
            else:
                st.warning("Log introuvable.")


def render_nexus_batch_runner_task() -> None:
    st.title("Exécution batch Nexus")
    st.caption("Détecte les batchs Nexus générés, vérifie les CSV/maps et exécute seulement les batchs sélectionnés.")

    input_tab, results_tab = st.tabs(["Détection et exécution", "Derniers résultats"])
    with input_tab:
        default_path = st.session_state.get("nexus_batch_package_path", _latest_rda_folder())
        path_text = st.text_input("Dossier RDA local", value=default_path, key="nexus_batch_path")
        if path_text != default_path:
            st.session_state["nexus_batch_package_path"] = path_text
        uploaded_zip = st.file_uploader("Ou déposer un dossier complet en .zip", type=["zip"], key="nexus_batch_zip")

        if uploaded_zip is not None and st.button("Extraire le zip", type="primary", key="nexus_batch_extract"):
            try:
                extracted = _safe_extract_zip(uploaded_zip)
                st.session_state["nexus_batch_package_path"] = str(extracted)
                path_text = str(extracted)
                st.success(f"Zip extrait dans: {extracted}")
            except Exception as exc:
                st.exception(exc)
                return

        if not path_text.strip():
            st.warning("Indiquez un dossier RDA local ou extrayez un zip.")
            return

        package_root = Path(path_text.strip().strip('"')).expanduser()
        candidates = _scan_batches(package_root)
        runnable = [candidate for candidate in candidates if candidate.runnable]
        blocked = [candidate for candidate in candidates if not candidate.runnable]

        metric_cols = st.columns(4)
        metric_cols[0].metric("Batchs détectés", f"{len(candidates):,}")
        metric_cols[1].metric("Exécutables", f"{len(runnable):,}")
        metric_cols[2].metric("Bloqués", f"{len(blocked):,}")
        metric_cols[3].metric("Dossier", "OK" if package_root.is_dir() else "Introuvable")

        if not package_root.is_dir():
            st.warning("Indiquez un dossier RDA local valide ou extrayez un zip.")
            return
        if not candidates:
            st.info("Aucun fichier .bat détecté dans ce dossier.")
            return

        rows = _candidate_rows(candidates, package_root)
        st.subheader("Aperçu")
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        options = [
            f"{row['ID']} | {row['Groupe']} | {row['Batch']}"
            for row in rows
            if candidates[row["ID"]].runnable
        ]
        selected_options = st.multiselect("Batchs à exécuter", options=options, key="nexus_batch_selected")
        selected_ids = [int(option.split(" | ", 1)[0]) for option in selected_options]
        selected_candidates = [candidates[index] for index in selected_ids]

        confirm = st.checkbox(
            "Je confirme que ces imports Nexus doivent être exécutés depuis cette machine.",
            key="nexus_batch_confirm",
        )
        run_disabled = not selected_candidates or not confirm
        if st.button("Exécuter les batchs sélectionnés", type="primary", disabled=run_disabled, key="nexus_batch_run", width="stretch"):
            render_blocking_run_warning()
            with st.spinner("Exécution des batchs Nexus..."):
                summary, log_dir = _run_batches(selected_candidates, package_root)
            st.session_state["nexus_batch_last_summary"] = summary
            st.session_state["nexus_batch_last_log_dir"] = str(log_dir)
            _render_run_results(summary, log_dir)

    with results_tab:
        summary = st.session_state.get("nexus_batch_last_summary")
        log_dir = st.session_state.get("nexus_batch_last_log_dir")
        if summary is None or not log_dir:
            st.info("Aucune exécution dans cette session.")
        else:
            _render_run_results(summary, Path(log_dir))
