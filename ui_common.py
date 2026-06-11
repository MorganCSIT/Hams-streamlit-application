from app_config import *


TEMPLATE_FOLDER = APP_ROOT / "Templates"


def template_display_name(path: Path) -> str:
    name = re.sub(r"[_-]+", " ", path.stem).strip()
    return name.title() if name else path.stem


def render_template_downloads(folder: Path = TEMPLATE_FOLDER) -> None:
    st.subheader("Modèles de fichiers")
    st.caption("Téléchargez le dossier d'exemples pour comparer les en-têtes et les formats attendus.")

    if not folder.exists():
        st.info("Aucun modèle disponible pour le moment. Ajoutez un fichier ZIP dans le dossier Templates.")
        return

    template_paths = sorted(folder.glob("*.zip"), key=lambda path: path.name.lower())
    if not template_paths:
        st.info("Aucun modèle disponible pour le moment. Ajoutez un fichier ZIP dans le dossier Templates.")
        return

    for path in template_paths:
        render_download_for_path(
            path,
            f"Télécharger {template_display_name(path)}",
            key=f"template_zip_{path.name}",
        )


def read_csv_flex(source) -> pd.DataFrame:
    try:
        return pd.read_csv(source, sep=None, engine="python")
    except Exception as first_exc:
        for sep in [",", ";", "\t", "|"]:
            try:
                if hasattr(source, "seek"):
                    source.seek(0)
                return pd.read_csv(source, sep=sep, engine="python")
            except Exception:
                pass
        for enc in ["utf-8", "utf-8-sig", "latin-1"]:
            try:
                if hasattr(source, "seek"):
                    source.seek(0)
                return pd.read_csv(source, sep=None, engine="python", encoding=enc)
            except Exception:
                pass
        raise RuntimeError(f"Could not read CSV. Last error: {first_exc}")


def read_excel_flex(source) -> pd.DataFrame:
    try:
        xls = pd.ExcelFile(source)
    except Exception as exc:
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            xls = pd.ExcelFile(source, engine="openpyxl")
        except Exception as exc2:
            raise RuntimeError(f"Could not open Excel file. Errors: {exc} / {exc2}")

    chosen_sheet = None
    for sheet in xls.sheet_names:
        try:
            df_head = pd.read_excel(xls, sheet_name=sheet, nrows=5)
            if df_head.shape[1] > 0 and df_head.shape[0] > 0:
                chosen_sheet = sheet
                break
        except Exception:
            continue

    if chosen_sheet is None:
        chosen_sheet = xls.sheet_names[0]

    try:
        return pd.read_excel(xls, sheet_name=chosen_sheet)
    except Exception:
        return pd.read_excel(source, sheet_name=chosen_sheet, engine="openpyxl")


def read_any_flex(source, filename: str) -> pd.DataFrame:
    ext = Path(filename.lower()).suffix
    if ext == ".csv":
        return read_csv_flex(source)
    if ext in [".xlsx", ".xls"]:
        return read_excel_flex(source)
    raise ValueError(f"Unsupported file type for '{filename}'")


def clean_illegal_excel_chars(df: pd.DataFrame) -> pd.DataFrame:
    illegal_char_re = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")
    out = df.copy()

    def clean_value(value):
        try:
            if pd.isna(value):
                return value
        except Exception:
            pass
        return illegal_char_re.sub("", str(value))

    for col in out.select_dtypes(include=["object"]).columns:
        out[col] = out[col].map(clean_value)
    return out


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


def merge_dataframes(named_frames: list[tuple[str, pd.DataFrame]], stop_on_sanity_mismatch: bool = False) -> dict:
    if len(named_frames) < 2:
        raise ValueError("Please provide at least 2 files to merge.")

    dfs = []
    row_counts = []
    schemas = []

    first_name, first_df = named_frames[0]
    first_df = first_df.loc[:, ~first_df.columns.astype(str).str.contains("^Unnamed")]
    if first_df.empty:
        raise AssertionError(f"First file appears empty: {first_name}")

    ref_cols = list(first_df.columns)
    dfs.append(first_df[ref_cols])
    row_counts.append(len(first_df))
    schemas.append({"file": first_name, "rows": len(first_df), "columns": len(ref_cols)})

    for filename, df in named_frames[1:]:
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
        if df.empty:
            raise AssertionError(f"File appears empty: {filename}")

        if set(df.columns) != set(ref_cols):
            missing = [col for col in ref_cols if col not in df.columns]
            extra = [col for col in df.columns if col not in ref_cols]
            raise AssertionError(
                f"Schema mismatch in {filename}.\n"
                f"Missing columns: {missing}\n"
                f"Extra columns: {extra}\n"
                "All files must have the same header names."
            )

        df = df.reindex(columns=ref_cols)
        dfs.append(df)
        row_counts.append(len(df))
        schemas.append({"file": filename, "rows": len(df), "columns": len(df.columns)})

    merged = pd.concat(dfs, ignore_index=True)
    sum_input_rows = sum(row_counts)
    merged_rows = len(merged)
    sanity_reason = ""

    if merged_rows != sum_input_rows:
        sanity_reason = (
            f"Le contrôle des lignes a échoué : le fichier fusionné contient {merged_rows} ligne(s), "
            f"mais les fichiers d'entrée totalisent {sum_input_rows} ligne(s)."
        )
        if stop_on_sanity_mismatch:
            raise AssertionError(sanity_reason)

    return {
        "merged": clean_illegal_excel_chars(merged),
        "schemas": schemas,
        "row_counts": row_counts,
        "sum_input_rows": sum_input_rows,
        "merged_rows": merged_rows,
        "sanity_passed": merged_rows == sum_input_rows,
        "sanity_reason": sanity_reason,
    }


def write_merged_output(merged: pd.DataFrame, output_dir: Path, output_basename: str, output_format: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_basename = re.sub(r"[^A-Za-z0-9_.-]+", "_", output_basename.strip()) or "merged"

    if output_format.lower() == "xlsx":
        out_path = output_dir / f"{safe_basename}_{timestamp}.xlsx"
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            merged.to_excel(writer, index=False, sheet_name="merged")
    else:
        out_path = output_dir / f"{safe_basename}_{timestamp}.csv"
        merged.to_csv(out_path, index=False, encoding="utf-8-sig", sep=";")

    return out_path


def mime_for_path(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    if suffix == ".zip":
        return "application/zip"
    if suffix == ".csv":
        return "text/csv"
    return "application/octet-stream"


def render_download_for_path(path: Path, label: str, key: str | None = None, width: str | None = None) -> None:
    if not path.exists():
        return
    with path.open("rb") as handle:
        kwargs = {"width": width} if width else {}
        st.download_button(
            label,
            handle,
            file_name=path.name,
            mime=mime_for_path(path),
            key=key or f"download_{label}_{path.resolve()}",
            **kwargs,
        )


def render_download_or_placeholder(path, label: str, key: str, width: str = "stretch") -> None:
    path = Path(path) if path else None
    if path is not None and path.is_file():
        render_download_for_path(path, label, key=key, width=width)
    else:
        st.button(label, disabled=True, key=f"{key}_placeholder", width=width)


def render_blocking_run_warning() -> None:
    st.warning(
        "Veuillez ne pas changer de section et ne pas utiliser d'autres parties de l'application pendant l'exécution. "
        "Attendez la fin du traitement."
    )


def safe_folder_name(value) -> str:
    text = str(value).strip()
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", text)[:80] or "unknown"

