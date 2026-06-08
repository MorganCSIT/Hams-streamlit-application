from app_config import *
from ui_common import render_blocking_run_warning, render_download_or_placeholder


@dataclass
class DownloadConfig:
    start_date: date
    end_date: date
    account: str
    username: str
    password: str
    api_key: str
    output_root: Path
    max_chunk_days: int
    request_interval_secs: int
    max_retries: int
    timeout_secs: int

    @property
    def start_str(self) -> str:
        return date_str(self.start_date)

    @property
    def end_str(self) -> str:
        return date_str(self.end_date)

    @property
    def run_folder(self) -> Path:
        return self.output_root / f"webfleet_year_download_{self.start_str}_to_{self.end_str}"

    @property
    def cache_folder(self) -> Path:
        return self.run_folder / "checkpoint_chunks"

    @property
    def output_csv(self) -> Path:
        return self.run_folder / f"webfleet_ALL_TRIPS_{self.start_str}_to_{self.end_str}.csv"

    @property
    def output_xlsx(self) -> Path:
        return self.run_folder / f"webfleet_ALL_TRIPS_{self.start_str}_to_{self.end_str}.xlsx"

    @property
    def manifest_csv(self) -> Path:
        return self.run_folder / "download_manifest.csv"


class WebfleetDownloader:
    def __init__(self, config: DownloadConfig, status_box, progress_bar):
        self.config = config
        self.status_box = status_box
        self.progress_bar = progress_bar
        self.last_request_time = 0.0
        self.session = requests.Session()
        self.session.auth = (config.username, config.password)
        self.session.headers.update(
            {
                "User-Agent": "Mozilla/5.0",
                "Accept": "text/csv, text/plain, */*",
            }
        )

    def chunk_cache_path(self, start_d: date, end_d: date) -> Path:
        return self.config.cache_folder / f"trips_{date_str(start_d)}_to_{date_str(end_d)}.csv"

    def chunk_empty_marker_path(self, start_d: date, end_d: date) -> Path:
        return self.config.cache_folder / f"EMPTY_{date_str(start_d)}_to_{date_str(end_d)}.txt"

    def build_params(self, start_d: date, end_d: date) -> dict:
        return {
            "account": self.config.account,
            "apikey": self.config.api_key,
            "action": API_ACTION,
            "outputformat": OUTPUTFORMAT,
            "lang": LANG,
            "useISO8601": "true",
            "rangefrom_string": f"{date_str(start_d)}T00:00:00",
            "rangeto_string": f"{date_str(end_d)}T23:59:59",
        }

    def wait_for_rate_limit(self) -> None:
        elapsed = time.monotonic() - self.last_request_time
        if elapsed < self.config.request_interval_secs:
            wait_s = self.config.request_interval_secs - elapsed
            self.status_box.info(f"Pause limite API : attente de {wait_s:.1f} secondes")
            time.sleep(wait_s)

    def fetch_period_once(self, start_d: date, end_d: date) -> pd.DataFrame:
        label = f"{date_str(start_d)} to {date_str(end_d)}"
        self.wait_for_rate_limit()
        self.status_box.info(f"Requête en cours : {label}")

        resp = self.session.get(
            BASE_URL,
            params=self.build_params(start_d, end_d),
            timeout=self.config.timeout_secs,
        )
        self.last_request_time = time.monotonic()

        if resp.status_code != 200:
            raise RuntimeError(f"HTTP {resp.status_code}: {resp.text[:500]}")

        text = resp.text
        if looks_like_webfleet_error(text):
            first_line = text.strip().splitlines()[0]
            if "document is empty" in first_line.lower():
                return pd.DataFrame()
            raise RuntimeError(f"Webfleet error for {label}: {first_line}")

        if not text.strip():
            return pd.DataFrame()

        return parse_csv_robust(text, self.config.cache_folder, label.replace(" ", "_"))

    def fetch_period_with_retries(self, start_d: date, end_d: date) -> pd.DataFrame:
        label = f"{date_str(start_d)} to {date_str(end_d)}"
        last_exc = None

        for attempt in range(1, self.config.max_retries + 1):
            try:
                return self.fetch_period_once(start_d, end_d)
            except Exception as exc:
                last_exc = exc
                self.status_box.warning(
                    f"{label} : tentative {attempt}/{self.config.max_retries} échouée : {exc}"
                )
                if attempt < self.config.max_retries:
                    time.sleep(min(10 * attempt, 60))

        raise RuntimeError(f"{label}: failed after {self.config.max_retries} attempts. Last error: {last_exc}")

    def save_chunk_atomic(self, df: pd.DataFrame, start_d: date, end_d: date) -> Path:
        path = self.chunk_cache_path(start_d, end_d)
        tmp_path = path.with_suffix(path.suffix + ".tmp")

        out = df.copy()
        out["__download_range_from"] = date_str(start_d)
        out["__download_range_to"] = date_str(end_d)
        out.to_csv(tmp_path, index=False, encoding="utf-8-sig")
        os.replace(tmp_path, path)
        return path

    def save_empty_marker(self, start_d: date, end_d: date) -> Path:
        path = self.chunk_empty_marker_path(start_d, end_d)
        path.write_text("empty\n", encoding="utf-8")
        return path

    def log_manifest(self, row: dict) -> None:
        row = dict(row)
        row["logged_at"] = datetime.now().isoformat(timespec="seconds")

        exists = self.config.manifest_csv.exists()
        pd.DataFrame([row]).to_csv(
            self.config.manifest_csv,
            mode="a",
            header=not exists,
            index=False,
            encoding="utf-8-sig",
        )

    def fetch_period_safe(self, start_d: date, end_d: date) -> list[Path]:
        cached_csv = self.chunk_cache_path(start_d, end_d)
        cached_empty = self.chunk_empty_marker_path(start_d, end_d)

        if cached_csv.exists():
            self.status_box.info(f"Déjà en cache : {cached_csv.name}")
            return [cached_csv]

        if cached_empty.exists():
            self.status_box.info(f"Déjà en cache comme vide : {cached_empty.name}")
            return []

        days_count = (end_d - start_d).days + 1
        label = f"{date_str(start_d)} to {date_str(end_d)}"

        try:
            df = self.fetch_period_with_retries(start_d, end_d)
            if df.empty:
                self.save_empty_marker(start_d, end_d)
                self.log_manifest(
                    {"from": date_str(start_d), "to": date_str(end_d), "status": "empty", "rows": 0, "file": "", "error": ""}
                )
                return []

            path = self.save_chunk_atomic(df, start_d, end_d)
            self.log_manifest(
                {
                    "from": date_str(start_d),
                    "to": date_str(end_d),
                    "status": "ok",
                    "rows": len(df),
                    "file": str(path),
                    "error": "",
                }
            )
            self.status_box.success(f"{path.name} enregistré avec {len(df):,} lignes")
            return [path]
        except Exception as exc:
            self.status_box.error(f"{label} : échec sur ce bloc : {exc}")
            if days_count <= 1:
                self.log_manifest(
                    {
                        "from": date_str(start_d),
                        "to": date_str(end_d),
                        "status": "failed",
                        "rows": 0,
                        "file": "",
                        "error": str(exc),
                    }
                )
                raise

            mid_d = start_d + timedelta(days=(days_count // 2) - 1)
            files = []
            files.extend(self.fetch_period_safe(start_d, mid_d))
            files.extend(self.fetch_period_safe(mid_d + timedelta(days=1), end_d))
            return files

    def run(self) -> dict:
        self.config.run_folder.mkdir(parents=True, exist_ok=True)
        self.config.cache_folder.mkdir(parents=True, exist_ok=True)

        period_list = list(make_periods(self.config.start_date, self.config.end_date, self.config.max_chunk_days))
        all_chunk_files = []
        failed = []

        for i, (start_d, end_d) in enumerate(period_list, start=1):
            self.progress_bar.progress((i - 1) / len(period_list), text=f"Téléchargement du bloc {i} sur {len(period_list)}")
            try:
                all_chunk_files.extend(self.fetch_period_safe(start_d, end_d))
            except Exception as exc:
                failed.append((date_str(start_d), date_str(end_d), str(exc)))

        self.progress_bar.progress(1.0, text="Assemblage des fichiers checkpoint")
        unique_chunk_files = dedupe_paths(all_chunk_files)
        total_rows, final_cols = combine_checkpoint_csvs(unique_chunk_files, self.config.output_csv)

        audit = audit_download(self.config)
        return {
            "failed": failed,
            "total_rows": total_rows,
            "columns": final_cols,
            "audit": audit,
            "chunk_count": len(unique_chunk_files),
        }


def date_str(value: date) -> str:
    return value.strftime("%Y-%m-%d")


def make_periods(start: date, end: date, max_days: int):
    cur = start
    while cur <= end:
        chunk_end = min(cur + timedelta(days=max_days - 1), end)
        yield cur, chunk_end
        cur = chunk_end + timedelta(days=1)


def looks_like_webfleet_error(text: str) -> bool:
    if not text or not text.strip():
        return False
    first_line = text.strip().splitlines()[0].strip()
    return bool(re.match(r"^9\d{3}\s*[,;]", first_line))


def parse_csv_robust(text: str, cache_folder: Path, label: str) -> pd.DataFrame:
    attempts = [
        dict(sep=None, engine="python"),
        dict(sep=";", engine="python", quoting=csv.QUOTE_MINIMAL),
        dict(sep=",", engine="python"),
        dict(sep=None, engine="python", on_bad_lines="skip"),
    ]

    last_exc = None
    for kwargs in attempts:
        try:
            return pd.read_csv(StringIO(text), **kwargs)
        except Exception as exc:
            last_exc = exc

    raw_path = cache_folder / f"RAW_PARSE_FAILED_{label}.csv"
    raw_path.write_text(text, encoding="utf-8")
    raise RuntimeError(f"Could not parse CSV for {label}. Raw file saved to {raw_path}. Last error: {last_exc}")


def collect_union_columns(csv_files: list[Path]) -> list[str]:
    cols = []
    for path in csv_files:
        header = pd.read_csv(path, nrows=0).columns.tolist()
        for col in header:
            if col not in cols:
                cols.append(col)
    return cols


def combine_checkpoint_csvs(csv_files: list[Path], output_csv: Path) -> tuple[int, list[str]]:
    if not csv_files:
        return 0, []

    csv_files = sorted(csv_files)
    all_cols = collect_union_columns(csv_files)
    seen_tripids = set()
    use_tripid_dedupe = "tripid" in all_cols
    first_write = True
    total_rows = 0
    tmp_output = output_csv.with_suffix(output_csv.suffix + ".tmp")

    if tmp_output.exists():
        tmp_output.unlink()

    for path in csv_files:
        for chunk in pd.read_csv(path, chunksize=50_000):
            chunk = chunk.reindex(columns=all_cols)
            if use_tripid_dedupe and "tripid" in chunk.columns:
                tripids = chunk["tripid"].astype(str)
                keep_mask = ~tripids.isin(seen_tripids)
                seen_tripids.update(tripids[keep_mask].tolist())
                chunk = chunk.loc[keep_mask]

            if chunk.empty:
                continue

            chunk.to_csv(
                tmp_output,
                mode="w" if first_write else "a",
                header=first_write,
                index=False,
                encoding="utf-8-sig",
            )
            first_write = False
            total_rows += len(chunk)

    if tmp_output.exists():
        os.replace(tmp_output, output_csv)
    return total_rows, all_cols


def write_xlsx_from_csv(input_csv: Path, output_xlsx: Path, total_rows: int) -> bool:
    if total_rows == 0 or total_rows > EXCEL_MAX_ROWS - 1:
        return False

    from openpyxl import Workbook

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("All Trips")
    first = True

    for chunk in pd.read_csv(input_csv, chunksize=50_000):
        if first:
            ws.append(list(chunk.columns))
            first = False
        for row in chunk.itertuples(index=False, name=None):
            ws.append(row)

    wb.save(output_xlsx)
    return True


def dedupe_paths(paths: list[Path]) -> list[Path]:
    seen = set()
    out = []
    for path in paths:
        if path not in seen:
            out.append(path)
            seen.add(path)
    return out


def count_csv_rows(path: Path) -> int:
    if not path.exists():
        return 0
    with path.open("r", encoding="utf-8-sig", errors="replace") as handle:
        return max(sum(1 for _ in handle) - 1, 0)


def parse_period_from_filename(path: Path):
    name = path.name
    match = re.match(r"trips_(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})\.csv$", name)
    if match:
        return date.fromisoformat(match.group(1)), date.fromisoformat(match.group(2)), "data"

    match = re.match(r"EMPTY_(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})\.txt$", name)
    if match:
        return date.fromisoformat(match.group(1)), date.fromisoformat(match.group(2)), "empty"

    return None


def days_in_period(start_d: date, end_d: date):
    cur = start_d
    while cur <= end_d:
        yield cur
        cur += timedelta(days=1)


def audit_download(config: DownloadConfig) -> dict:
    requested_days = set(days_in_period(config.start_date, config.end_date))
    covered_days = set()
    empty_days = set()

    chunk_files = sorted(config.cache_folder.glob("trips_*.csv"))
    empty_markers = sorted(config.cache_folder.glob("EMPTY_*.txt"))

    for path in chunk_files + empty_markers:
        parsed = parse_period_from_filename(path)
        if parsed is None:
            continue
        start_d, end_d, kind = parsed
        for day in days_in_period(start_d, end_d):
            covered_days.add(day)
            if kind == "empty":
                empty_days.add(day)

    manifest_failed_count = 0
    if config.manifest_csv.exists():
        manifest = pd.read_csv(config.manifest_csv)
        if "status" in manifest.columns:
            manifest_failed_count = int(manifest["status"].astype(str).str.lower().eq("failed").sum())

    raw_checkpoint_rows = sum(count_csv_rows(path) for path in chunk_files)
    final_rows = count_csv_rows(config.output_csv)

    tripid_summary = {}
    raw_tripids = []
    tripid_exists = False
    for path in chunk_files:
        try:
            header = pd.read_csv(path, nrows=0).columns.tolist()
        except Exception:
            continue
        if "tripid" in header:
            tripid_exists = True
            for chunk in pd.read_csv(path, usecols=["tripid"], chunksize=100_000):
                raw_tripids.extend(chunk["tripid"].dropna().astype(str).tolist())

    missing_tripids_count = 0
    duplicate_tripids_count = 0
    if tripid_exists:
        raw_counter = Counter(raw_tripids)
        duplicate_tripids_count = sum(1 for count in raw_counter.values() if count > 1)
        final_tripids = set()
        if config.output_csv.exists() and "tripid" in pd.read_csv(config.output_csv, nrows=0).columns.tolist():
            for chunk in pd.read_csv(config.output_csv, usecols=["tripid"], chunksize=100_000):
                final_tripids.update(chunk["tripid"].dropna().astype(str).tolist())
        missing_tripids_count = len(set(raw_counter.keys()) - final_tripids)
        tripid_summary = {
            "raw_tripids": len(raw_tripids),
            "unique_raw_tripids": len(raw_counter),
            "duplicate_tripids": duplicate_tripids_count,
            "final_tripids": len(final_tripids),
            "missing_tripids": missing_tripids_count,
        }

    missing_coverage = sorted(requested_days - covered_days)
    hard_fail = bool(missing_coverage or manifest_failed_count or missing_tripids_count)
    if not tripid_exists and raw_checkpoint_rows != final_rows:
        hard_fail = True

    return {
        "passed": not hard_fail,
        "checkpoint_files": len(chunk_files),
        "empty_markers": len(empty_markers),
        "missing_days": [date_str(day) for day in missing_coverage],
        "empty_days": [date_str(day) for day in sorted(empty_days & requested_days)],
        "manifest_failed_count": manifest_failed_count,
        "raw_checkpoint_rows": raw_checkpoint_rows,
        "final_rows": final_rows,
        "tripid_exists": tripid_exists,
        "tripid_summary": tripid_summary,
    }


def render_export_buttons(config: DownloadConfig) -> None:
    if not config.output_csv.exists():
        return

    st.subheader("Export")
    total_rows = count_csv_rows(config.output_csv)
    export_cols = st.columns(2)

    with export_cols[0]:
        with config.output_csv.open("rb") as handle:
            st.download_button(
                "Télécharger le CSV",
                handle,
                file_name=config.output_csv.name,
                mime="text/csv",
                width="stretch",
            )

    can_make_excel = 0 < total_rows <= EXCEL_MAX_ROWS - 1
    xlsx_is_current = (
        config.output_xlsx.exists()
        and config.output_xlsx.stat().st_mtime >= config.output_csv.stat().st_mtime
    )
    if can_make_excel and not xlsx_is_current:
        with st.spinner("Préparation du fichier Excel..."):
            write_xlsx_from_csv(config.output_csv, config.output_xlsx, total_rows)

    with export_cols[1]:
        if can_make_excel and config.output_xlsx.exists():
            with config.output_xlsx.open("rb") as handle:
                st.download_button(
                    "Télécharger l'Excel",
                    handle,
                    file_name=config.output_xlsx.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                )
        else:
            st.button("Télécharger l'Excel", disabled=True, width="stretch")
            if total_rows == 0:
                st.caption("Excel indisponible : aucune ligne à exporter.")
            elif total_rows > EXCEL_MAX_ROWS - 1:
                st.caption("Excel indisponible : trop de lignes pour une seule feuille.")


def ensure_webfleet_xlsx(config: DownloadConfig) -> None:
    if not config.output_csv.exists():
        return
    total_rows = count_csv_rows(config.output_csv)
    can_make_excel = 0 < total_rows <= EXCEL_MAX_ROWS - 1
    xlsx_is_current = (
        config.output_xlsx.exists()
        and config.output_xlsx.stat().st_mtime >= config.output_csv.stat().st_mtime
    )
    if can_make_excel and not xlsx_is_current:
        write_xlsx_from_csv(config.output_csv, config.output_xlsx, total_rows)


@st.cache_data(show_spinner=False)
def load_dashboard_data(path: str, modified_time: float) -> pd.DataFrame:
    _ = modified_time
    return pd.read_csv(path)


def find_column(columns: list[str], target: str) -> str | None:
    normalized_target = target.lower().replace("_", "")
    for column in columns:
        normalized_column = column.lower().replace("_", "")
        if normalized_column == normalized_target:
            return column
    return None


def coerce_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def coerce_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def format_datetime_series(series: pd.Series) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    formatted = parsed.dt.strftime("%Y-%m-%d %H:%M:%S")
    return formatted.fillna(series.astype(str))


def format_duration_hh_mm_ss(series: pd.Series) -> pd.Series:
    seconds = coerce_numeric(series)

    def format_one(value):
        if pd.isna(value):
            return ""
        total_seconds = int(round(value))
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        remaining_seconds = total_seconds % 60
        return f"{hours}:{minutes:02d}:{remaining_seconds:02d}"

    return seconds.apply(format_one)


def missing_value_mask(series: pd.Series) -> pd.Series:
    return series.isna() | series.astype(str).str.strip().eq("")


def filter_options_with_missing(series: pd.Series) -> list[str]:
    present_values = sorted(series[~missing_value_mask(series)].astype(str).unique().tolist())
    if missing_value_mask(series).any():
        return [MISSING_FILTER_LABEL] + present_values
    return present_values


def apply_categorical_filter(df: pd.DataFrame, column: str, selected: list[str]) -> pd.DataFrame:
    if not selected:
        return df

    mask = pd.Series(False, index=df.index)
    if MISSING_FILTER_LABEL in selected:
        mask = mask | missing_value_mask(df[column])

    selected_present = [value for value in selected if value != MISSING_FILTER_LABEL]
    if selected_present:
        mask = mask | df[column].astype(str).isin(selected_present)

    return df[mask]


def render_dashboard(csv_path: Path) -> None:
    if not csv_path.exists():
        st.info("Téléchargez d'abord les données pour ouvrir le tableau de bord.")
        return

    try:
        df = load_dashboard_data(str(csv_path), csv_path.stat().st_mtime)
    except Exception as exc:
        st.error(f"Impossible de charger le CSV du tableau de bord : {exc}")
        return

    column_map = {name: find_column(df.columns.tolist(), name) for name in DASHBOARD_COLUMNS}
    available_dashboard_cols = [column_map[name] for name in DASHBOARD_COLUMNS if column_map[name]]
    missing_cols = [name for name in DASHBOARD_COLUMNS if not column_map[name]]

    if not available_dashboard_cols:
        st.warning("Aucune des colonnes attendues pour le tableau de bord n'a été trouvée dans le CSV de sortie.")
        st.dataframe(pd.DataFrame({"available_columns": df.columns.tolist()}), width="stretch")
        return

    st.subheader("Tableau de bord")
    st.caption(f"{len(df):,} trajets chargés depuis {csv_path.name}. Les indicateurs et le tableau se mettent à jour avec les filtres.")

    with st.expander("Colonnes", expanded=True):
        selected_display_cols = st.multiselect(
            "Colonnes à afficher",
            options=df.columns.tolist(),
            default=available_dashboard_cols,
            key=f"dashboard_columns_{csv_path}",
        )

    filters = st.container()
    filtered = df.copy()

    with filters:
        filter_cols = st.columns(4)

        if column_map["tripmode"]:
            values = filter_options_with_missing(filtered[column_map["tripmode"]])
            selected = filter_cols[0].multiselect("Mode de trajet", values, default=values)
            filtered = apply_categorical_filter(filtered, column_map["tripmode"], selected)

        if column_map["drivername"]:
            values = filter_options_with_missing(filtered[column_map["drivername"]])
            selected = filter_cols[1].multiselect("Conducteur", values)
            filtered = apply_categorical_filter(filtered, column_map["drivername"], selected)

        if column_map["driverno"]:
            values = filter_options_with_missing(filtered[column_map["driverno"]])
            selected = filter_cols[2].multiselect("N° conducteur", values)
            filtered = apply_categorical_filter(filtered, column_map["driverno"], selected)

        if column_map["objectname"]:
            values = filter_options_with_missing(filtered[column_map["objectname"]])
            selected = filter_cols[3].multiselect("Objet", values)
            filtered = apply_categorical_filter(filtered, column_map["objectname"], selected)

        date_filter_cols = st.columns(3)
        start_col = column_map["start_time"]
        end_col = column_map["end_time"]
        if start_col:
            start_times = coerce_datetime(filtered[start_col])
            valid_start_times = start_times.dropna()
            if not valid_start_times.empty:
                min_day = valid_start_times.min().date()
                max_day = valid_start_times.max().date()
                selected_days = date_filter_cols[0].date_input("Plage de dates de début", value=(min_day, max_day))
                if isinstance(selected_days, tuple) and len(selected_days) == 2:
                    start_day, end_day = selected_days
                    mask = start_times.dt.date.between(start_day, end_day)
                    filtered = filtered[mask.fillna(False)]

        if column_map["distance"]:
            distance_km = coerce_numeric(filtered[column_map["distance"]]) / 1000
            valid_distance = distance_km.dropna()
            if not valid_distance.empty:
                min_distance = float(valid_distance.min())
                max_distance = float(valid_distance.max())
                if min_distance < max_distance:
                    selected_distance = date_filter_cols[1].slider(
                        "Plage de distance (km)",
                        min_value=min_distance,
                        max_value=max_distance,
                        value=(min_distance, max_distance),
                    )
                    filtered = filtered[distance_km.between(*selected_distance).fillna(False)]
                else:
                    date_filter_cols[1].metric("Distance km", f"{min_distance:,.2f}")

        if column_map["duration"]:
            duration_hours = coerce_numeric(filtered[column_map["duration"]]) / 3600
            valid_duration = duration_hours.dropna()
            if not valid_duration.empty:
                min_duration = float(valid_duration.min())
                max_duration = float(valid_duration.max())
                if min_duration < max_duration:
                    selected_duration = date_filter_cols[2].slider(
                        "Plage de durée (heures)",
                        min_value=min_duration,
                        max_value=max_duration,
                        value=(min_duration, max_duration),
                    )
                    filtered = filtered[duration_hours.between(*selected_duration).fillna(False)]
                else:
                    date_filter_cols[2].metric("Durée en heures", f"{min_duration:,.2f}")

        search = st.text_input("Recherche")
        if search:
            search_cols = [col for col in [column_map["drivername"], column_map["driverno"], column_map["objectname"]] if col]
            if search_cols:
                mask = pd.Series(False, index=filtered.index)
                for col in search_cols:
                    mask = mask | filtered[col].astype(str).str.contains(search, case=False, na=False)
                filtered = filtered[mask]

    if missing_cols:
        st.caption(f"Colonnes manquantes dans ce CSV : {', '.join(missing_cols)}")

    st.subheader("Résumé filtré")
    filtered_metric_cols = st.columns(5)
    filtered_metric_cols[0].metric("Trajets", f"{len(filtered):,}")
    if column_map["distance"]:
        filtered_distance_km = coerce_numeric(filtered[column_map["distance"]]).sum(skipna=True) / 1000
        filtered_metric_cols[1].metric("Distance", f"{filtered_distance_km:,.1f} km")
    else:
        filtered_metric_cols[1].metric("Distance", "n/a")
    if column_map["duration"]:
        filtered_duration_seconds = coerce_numeric(filtered[column_map["duration"]]).sum(skipna=True)
        filtered_metric_cols[2].metric("Secondes", f"{filtered_duration_seconds:,.0f}")
        filtered_metric_cols[3].metric("Minutes", f"{filtered_duration_seconds / 60:,.2f}")
        filtered_metric_cols[4].metric("Heures", f"{filtered_duration_seconds / 3600:,.2f}")
        st.caption("La colonne durée du CSV est en secondes. Les minutes et heures sont calculées depuis ce total brut.")
    else:
        filtered_metric_cols[2].metric("Secondes", "n/a")
        filtered_metric_cols[3].metric("Minutes", "n/a")
        filtered_metric_cols[4].metric("Heures", "n/a")
    display_df = filtered.copy()
    for category_name in ["tripmode", "drivername", "driverno", "objectname"]:
        actual_col = column_map[category_name]
        if actual_col and actual_col in display_df.columns:
            values = display_df[actual_col]
            display_df[actual_col] = values.astype("string")
            display_df.loc[missing_value_mask(values), actual_col] = MISSING_FILTER_LABEL

    for time_col_name in ["start_time", "end_time"]:
        actual_col = column_map[time_col_name]
        if actual_col and actual_col in display_df.columns:
            display_df[actual_col] = format_datetime_series(display_df[actual_col])

    if column_map["duration"] and column_map["duration"] in display_df.columns:
        duration_col = column_map["duration"]
        display_df["duration_hh_mm_ss"] = format_duration_hh_mm_ss(display_df[duration_col])
        if duration_col in selected_display_cols:
            duration_index = selected_display_cols.index(duration_col)
            selected_display_cols = selected_display_cols.copy()
            selected_display_cols[duration_index] = "duration_hh_mm_ss"

    if column_map["distance"] and column_map["distance"] in display_df.columns:
        distance_col = column_map["distance"]
        display_df["distance_km"] = coerce_numeric(display_df[distance_col]) / 1000
        if distance_col in selected_display_cols:
            distance_index = selected_display_cols.index(distance_col)
            selected_display_cols = selected_display_cols.copy()
            selected_display_cols[distance_index] = "distance_km"
    if selected_display_cols:
        st.dataframe(display_df[selected_display_cols], width="stretch", hide_index=True)
    else:
        st.warning("Activez au moins une colonne pour afficher le tableau.")

    export_cols = selected_display_cols or available_dashboard_cols
    export_df = display_df
    st.download_button(
        "Télécharger le CSV filtré",
        export_df[export_cols].to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"filtered_{csv_path.name}",
        mime="text/csv",
    )


def render_webfleet_task() -> None:
    st.title("Téléchargement des journaux Webfleet")

    default_output = get_session_output_root("WebfleetReports")

    with st.expander("API Webfleet", expanded=True):
        api_cols = st.columns(4)
        account = api_cols[0].text_input("Compte")
        username = api_cols[1].text_input("Utilisateur")
        password = api_cols[2].text_input("Mot de passe", type="password")
        api_key = api_cols[3].text_input("Clé API", type="password")
        today = date.today()
        date_cols = st.columns(2)
        with date_cols[0]:
            start_date = st.date_input("Date de début", value=date(today.year, 1, 1))
        with date_cols[1]:
            end_date = st.date_input("Date de fin", value=today)

    output_root = default_output
    with st.expander("Paramètres avancés", expanded=False):
        advanced_cols = st.columns(4)
        max_chunk_days = advanced_cols[0].number_input("Jours max par requête", min_value=1, max_value=31, value=7)
        request_interval_secs = advanced_cols[1].number_input("Secondes entre requêtes", min_value=0, max_value=600, value=61)
        max_retries = advanced_cols[2].number_input("Nombre max de tentatives", min_value=1, max_value=20, value=5)
        timeout_secs = advanced_cols[3].number_input("Timeout requête en secondes", min_value=10, max_value=1200, value=300)

    if start_date > end_date:
        st.error("La date de début doit être antérieure ou égale à la date de fin.")
        return

    config = DownloadConfig(
        start_date=start_date,
        end_date=end_date,
        account=account.strip(),
        username=username.strip(),
        password=password,
        api_key=api_key,
        output_root=output_root,
        max_chunk_days=int(max_chunk_days),
        request_interval_secs=int(request_interval_secs),
        max_retries=int(max_retries),
        timeout_secs=int(timeout_secs),
    )

    st.caption("Les fichiers générés sont disponibles avec les boutons de téléchargement de cette session.")

    ready = all([config.account, config.username, config.password, config.api_key])
    if not ready:
        st.info("Saisissez vos identifiants Webfleet pour commencer.")

    download_tab, dashboard_tab = st.tabs(["Téléchargement", "Tableau de bord"])

    with download_tab:
        if config.output_csv.exists():
            ensure_webfleet_xlsx(config)

        action_cols = st.columns([2, 1, 1])
        with action_cols[0]:
            start_download = st.button("Commencer le téléchargement", type="primary", disabled=not ready, width="stretch")

        if start_download:
            render_blocking_run_warning()
            status_box = st.empty()
            progress_bar = st.progress(0.0, text="Démarrage")

            try:
                result = WebfleetDownloader(config, status_box, progress_bar).run()
                st.session_state["latest_output_csv"] = str(config.output_csv)
                ensure_webfleet_xlsx(config)
            except Exception as exc:
                st.exception(exc)
                return

        with action_cols[1]:
            render_download_or_placeholder(config.output_csv, "Télécharger le CSV", key="webfleet_main_csv")
        with action_cols[2]:
            render_download_or_placeholder(config.output_xlsx, "Télécharger l'Excel", key="webfleet_main_xlsx")

        if start_download:
            st.subheader("Résultat")
            metric_cols = st.columns(4)
            metric_cols[0].metric("Lignes", f"{result['total_rows']:,}")
            metric_cols[1].metric("Fichiers checkpoint", f"{result['chunk_count']:,}")
            metric_cols[2].metric("Colonnes", f"{len(result['columns']):,}")
            metric_cols[3].metric("Blocs échoués", f"{len(result['failed']):,}")

            if result["total_rows"] > 0:
                st.success("Trajets téléchargés. Le CSV et l'Excel sont disponibles au téléchargement, et le tableau de bord peut utiliser le CSV généré.")
            else:
                st.warning("Aucune ligne de trajet trouvée. Des marqueurs checkpoint vides peuvent quand même avoir été créés.")

            if result["failed"]:
                st.error("Certaines périodes ont échoué définitivement. Relancez avec les mêmes paramètres pour reprendre les blocs en cache.")
                st.dataframe(pd.DataFrame(result["failed"], columns=["from", "to", "error"]), width="stretch")

            audit = result["audit"]
            st.subheader("Audit")
            audit_cols = st.columns(4)
            audit_cols[0].metric("Audit", "Réussi" if audit["passed"] else "Échec")
            audit_cols[1].metric("Lignes checkpoint brutes", f"{audit['raw_checkpoint_rows']:,}")
            audit_cols[2].metric("Lignes finales", f"{audit['final_rows']:,}")
            audit_cols[3].metric("Marqueurs vides", f"{audit['empty_markers']:,}")

            if audit["missing_days"]:
                st.error("Couverture checkpoint manquante pour ces jours :")
                st.write(audit["missing_days"])

            if audit["manifest_failed_count"]:
                st.error(f"Le manifeste contient {audit['manifest_failed_count']} périodes en échec.")

            if audit["tripid_exists"]:
                st.write("Résumé des Trip ID")
                st.dataframe(pd.DataFrame([audit["tripid_summary"]]), width="stretch")

            if audit["empty_days"]:
                with st.expander("Jours marqués vides"):
                    st.write(audit["empty_days"])

    with dashboard_tab:
        dashboard_path = Path(st.session_state.get("latest_output_csv", config.output_csv))
        render_dashboard(dashboard_path)

