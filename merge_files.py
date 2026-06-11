from app_config import *
from ui_common import *

def uploaded_merge_frames(uploaded_files) -> list[tuple[str, pd.DataFrame]]:
    frames = []
    for uploaded in uploaded_files:
        data = BytesIO(uploaded.getvalue())
        frames.append((uploaded.name, read_any_flex(data, uploaded.name)))
    return frames


def local_merge_frames(search_dir: Path, patterns: list[str], include_subfolders: bool) -> list[tuple[str, pd.DataFrame]]:
    files = []
    for pattern in patterns:
        iterator = search_dir.rglob(pattern) if include_subfolders else search_dir.glob(pattern)
        files.extend(path for path in iterator if path.is_file() and not path.name.startswith("."))
    paths = sorted(set(files))
    return [(str(path), read_any_flex(path, path.name)) for path in paths]


def render_merge_dashboard(df: pd.DataFrame, key_prefix: str) -> None:
    st.subheader("Tableau de bord")
    with st.container():
        metric_cols = st.columns(4)
        metric_cols[0].metric("Lignes", f"{len(df):,}")
        metric_cols[1].metric("Colonnes", f"{len(df.columns):,}")
        metric_cols[2].metric("Cellules vides", f"{int(df.isna().sum().sum()):,}")
        metric_cols[3].metric("Doublons", f"{int(df.duplicated().sum()):,}")

        selected_cols = st.multiselect(
            "Colonnes à afficher",
            options=df.columns.tolist(),
            default=df.columns.tolist()[: min(12, len(df.columns))],
            key=f"{key_prefix}_display_cols",
        )

        filtered = df.copy()
        search = st.text_input("Recherche", key=f"{key_prefix}_search")
        if search:
            search_columns = selected_cols or df.columns.tolist()
            mask = pd.Series(False, index=filtered.index)
            for col in search_columns:
                mask = mask | filtered[col].astype(str).str.contains(search, case=False, na=False)
            filtered = filtered[mask]

        low_cardinality_cols = [
            col
            for col in df.columns
            if df[col].nunique(dropna=True) <= 100
        ]
        category_filter_cols = st.multiselect(
            "Filtres par valeur",
            options=df.columns.tolist(),
            default=low_cardinality_cols[: min(4, len(low_cardinality_cols))],
            key=f"{key_prefix}_category_cols",
        )
        if category_filter_cols:
            value_filter_cols = st.columns(min(4, len(category_filter_cols)))
            for index, col in enumerate(category_filter_cols):
                values = filter_options_with_missing(filtered[col])
                selected = value_filter_cols[index % len(value_filter_cols)].multiselect(
                    col,
                    values,
                    key=f"{key_prefix}_filter_{col}",
                )
                filtered = apply_categorical_filter(filtered, col, selected)

        numeric_cols = [
            col
            for col in df.columns
            if not pd.api.types.is_bool_dtype(df[col])
            and pd.to_numeric(df[col], errors="coerce").notna().sum() > 0
        ]
        filter_cols = st.columns(2)
        if numeric_cols:
            numeric_col = filter_cols[0].selectbox(
                "Filtre numérique",
                ["Aucun"] + numeric_cols,
                key=f"{key_prefix}_numeric_col",
            )
            if numeric_col != "Aucun":
                numeric_values = pd.to_numeric(filtered[numeric_col], errors="coerce")
                valid_numeric = numeric_values.dropna()
                if not valid_numeric.empty:
                    min_value = float(valid_numeric.min())
                    max_value = float(valid_numeric.max())
                    if min_value < max_value:
                        selected_range = filter_cols[0].slider(
                            "Plage",
                            min_value=min_value,
                            max_value=max_value,
                            value=(min_value, max_value),
                            key=f"{key_prefix}_numeric_range",
                        )
                        filtered = filtered[numeric_values.between(*selected_range).fillna(False)]

        date_cols = []
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]) or pd.api.types.is_bool_dtype(df[col]):
                continue
            parsed = pd.to_datetime(df[col], errors="coerce")
            if parsed.notna().sum() > 0 and parsed.notna().mean() >= 0.5:
                date_cols.append(col)
        if date_cols:
            date_col = filter_cols[1].selectbox(
                "Filtre date",
                ["Aucun"] + date_cols,
                key=f"{key_prefix}_date_col",
            )
            if date_col != "Aucun":
                date_values = pd.to_datetime(filtered[date_col], errors="coerce")
                valid_dates = date_values.dropna()
                if not valid_dates.empty:
                    min_day = valid_dates.min().date()
                    max_day = valid_dates.max().date()
                    selected_days = filter_cols[1].date_input(
                        "Plage de dates",
                        value=(min_day, max_day),
                        key=f"{key_prefix}_date_range",
                    )
                    if isinstance(selected_days, tuple) and len(selected_days) == 2:
                        start_day, end_day = selected_days
                        filtered = filtered[date_values.dt.date.between(start_day, end_day).fillna(False)]

        st.caption(f"{len(filtered):,} ligne(s) affichée(s) après filtres.")
        display_cols = selected_cols or df.columns.tolist()
        st.dataframe(filtered[display_cols], width="stretch", hide_index=True)
        st.download_button(
            "Télécharger le CSV filtré",
            filtered[display_cols].to_csv(index=False, encoding="utf-8-sig"),
            file_name="filtered_merged.csv",
            mime="text/csv",
            key=f"{key_prefix}_download",
        )


def render_merge_task() -> None:
    st.title("Fusion de fichiers")
    st.caption("Fusionne les fichiers CSV/XLSX/XLS avec les mêmes noms d'en-têtes. L'ordre des colonnes suit le premier fichier.")
    st.info(
        "Conseil: les fichiers à fusionner doivent avoir les mêmes en-têtes de colonnes. "
        "L'ordre peut être différent, mais les noms doivent correspondre exactement."
    )

    default_output = get_session_output_root(MERGE_OUTPUT_FOLDER)

    with st.expander("Paramètres de fusion", expanded=True):
        settings_cols = st.columns(2)
        output_format = settings_cols[0].selectbox("Format de sortie", ["xlsx", "csv"], index=0)
        output_basename = settings_cols[1].text_input("Nom de base de sortie", value="")
        output_dir = default_output
        st.caption("Le fichier fusionné sera disponible au téléchargement après la fusion.")

    uploaded_files = st.file_uploader(
        "Fichiers à fusionner",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
    )

    latest_output = Path(st.session_state.get("latest_merge_output", ""))
    latest_result = st.session_state.get("latest_merge_result")
    merge_tab, dashboard_tab = st.tabs(["Fusion", "Tableau de bord"])

    with merge_tab:
        action_cols = st.columns([2, 1])
        with action_cols[0]:
            run_merge = st.button("Fusionner les fichiers", type="primary", disabled=len(uploaded_files) < 2, width="stretch")

        if run_merge:
            try:
                render_blocking_run_warning()
                named_frames = uploaded_merge_frames(uploaded_files)

                result = merge_dataframes(named_frames)
                out_path = write_merged_output(result["merged"], output_dir, output_basename, output_format)
                st.session_state["latest_merge_output"] = str(out_path)
                st.session_state["latest_merge_result"] = {
                    key: value
                    for key, value in result.items()
                    if key != "merged"
                }
                latest_output = out_path
                latest_result = st.session_state["latest_merge_result"]
            except Exception as exc:
                st.exception(exc)
                return

        with action_cols[1]:
            render_download_or_placeholder(latest_output if latest_output.is_file() else None, "Télécharger le fichier fusionné", key="merge_main_output")

        if latest_output.is_file():
            st.success("Fichier fusionné prêt au téléchargement.")

        if latest_result:
            st.subheader("Résumé de fusion")
            metric_cols = st.columns(4)
            metric_cols[0].metric("Fichiers fusionnés", f"{len(latest_result['schemas']):,}")
            metric_cols[1].metric("Lignes en entrée", f"{latest_result['sum_input_rows']:,}")
            metric_cols[2].metric("Lignes fusionnées", f"{latest_result['merged_rows']:,}")
            metric_cols[3].metric("Contrôle", "Réussi" if latest_result["sanity_passed"] else "Échec")
            if not latest_result["sanity_passed"]:
                st.warning(latest_result.get("sanity_reason") or "Le contrôle de fusion a échoué.")
            st.dataframe(pd.DataFrame(latest_result["schemas"]), width="stretch", hide_index=True)

    with dashboard_tab:
        if latest_output.is_file():
            dashboard_df = read_any_flex(latest_output, latest_output.name)
            render_merge_dashboard(dashboard_df, f"merge_dashboard_{latest_output.name}")
        else:
            st.info("Fusionnez d'abord des fichiers pour ouvrir le tableau de bord.")

