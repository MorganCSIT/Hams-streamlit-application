from app_config import *
from ui_common import read_any_flex, render_blocking_run_warning, render_download_or_placeholder, safe_folder_name

def ltr_unique_output_root(output_name: str) -> Path:
    safe_name = safe_folder_name(output_name) if output_name.strip() else f"LTR_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    root = get_session_output_root(LTR_OUTPUT_FOLDER) / safe_name
    if root.exists():
        root = root.with_name(f"{root.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    root.mkdir(parents=True, exist_ok=True)
    return root


def ltr_save_upload(uploaded_file, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    path = target_dir / uploaded_file.name
    path.write_bytes(uploaded_file.getvalue())
    return path


class LtrPytzCompat:
    @staticmethod
    def timezone(name: str):
        return ZoneInfo(name)


@st.cache_resource(show_spinner=False)
def ltr_load_notebook_functions(notebook_mtime: float) -> dict:
    _ = notebook_mtime
    notebook_path = LTR_NOTEBOOK_PATH
    if not notebook_path.exists():
        raise FileNotFoundError(f"LTR notebook not found: {notebook_path}")

    from openpyxl import load_workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    env = {
        "__builtins__": __builtins__,
        "os": os,
        "re": re,
        "np": np,
        "pd": pd,
        "pytz": LtrPytzCompat,
        "Path": Path,
        "datetime": datetime,
        "timedelta": timedelta,
        "dtime": dtime,
        "load_workbook": load_workbook,
        "Alignment": Alignment,
        "get_column_letter": get_column_letter,
        "SWISS_TZ": ZoneInfo("Europe/Zurich"),
        "MAX_SINGLE_INTERVAL_HOURS": 24,
        "SERVICE_RESET_GAP_HOURS": 8.0,
        "WEEKLY_LIMIT_HOURS": 50.0,
        "SPAN_LIMIT_HOURS": 14.0,
        "REST_NORMAL_HOURS": 11.0,
        "REST_ABSOLUTE_MIN_HOURS": 8.0,
        "REQUIRE_AVG_FOR_REDUCED_REST": True,
        "RUN_OVER50H": True,
        "RUN_STREAK": True,
        "RUN_SPAN": True,
        "RUN_REST11": True,
        "RUN_BREAKS": True,
        "COUNT_TRANSPORT_AS_WORK_FOR_50H": True,
        "COUNT_TRANSPORT_AS_WORK_FOR_STREAK": True,
        "COUNT_TRANSPORT_AS_WORK_FOR_BREAKS": True,
        "COUNT_TRANSPORT_FOR_SERVICE_BOUNDARY": True,
        "PAUSE_CODES": {"16009", "95900"},
        "TRANSPORT_CODES": {"61800", "61010"},
        "EXCLUDE_PRESTATIONS": {"196", "60041"},
        "PSEUDO_NON_WORK_CODES": {"195"},
    }

    notebook = json.loads(notebook_path.read_text(encoding="utf-8"))
    for cell_index in range(2, 9):
        source = "".join(notebook["cells"][cell_index]["source"])
        exec(compile(source, f"{notebook_path.name}:cell{cell_index + 1}", "exec"), env)
    return env


def ltr_compute_summary(env: dict, df: pd.DataFrame, services_df: pd.DataFrame, calendar_slices_df: pd.DataFrame, all_infractions: pd.DataFrame, data_quality: pd.DataFrame) -> pd.DataFrame:
    raw_months = set(df["Mois"].dropna().astype(str).tolist()) if "Mois" in df.columns else set()
    service_months = set(services_df["service_month"].dropna().astype(str).tolist()) if not services_df.empty else set()
    calendar_months = set(calendar_slices_df["target_month"].dropna().astype(str).tolist()) if not calendar_slices_df.empty else set()
    inf_months = set(all_infractions["TARGET_MONTH"].dropna().astype(str).tolist()) if not all_infractions.empty else set()
    months = sorted(raw_months | service_months | calendar_months | inf_months)

    rows = []
    for month in months:
        row = {"TARGET_MONTH": month}
        for rule in [
            f"OVER_{int(env['WEEKLY_LIMIT_HOURS'])}H_WEEK",
            "STREAK_7DAYS",
            f"SPAN_OVER_{int(env['SPAN_LIMIT_HOURS'])}H",
            "REST_UNDER_11H",
            "PAUSE_INSUFF",
        ]:
            if all_infractions.empty:
                row[rule] = 0
            else:
                row[rule] = int(
                    (
                        all_infractions.get("TARGET_MONTH", pd.Series(dtype=str)).astype(str).eq(month)
                        & all_infractions.get("RULE", pd.Series(dtype=str)).astype(str).eq(rule)
                    ).sum()
                )
        row["TOTAL_INFRACTIONS"] = int(all_infractions["TARGET_MONTH"].astype(str).eq(month).sum()) if not all_infractions.empty else 0
        row["SERVICES_STARTED"] = int(services_df["service_month"].astype(str).eq(month).sum()) if not services_df.empty else 0
        row["CALENDAR_HOURS"] = float(calendar_slices_df.loc[calendar_slices_df["target_month"].astype(str).eq(month), "hours"].sum()) if not calendar_slices_df.empty else 0.0
        row["ORPHAN_PAUSE_ROWS"] = int(
            (
                data_quality.get("TARGET_MONTH", pd.Series(dtype=str)).astype(str).eq(month)
                & data_quality.get("QUALITY_TYPE", pd.Series(dtype=str)).astype(str).eq("ORPHAN_PAUSE_NO_SERVICE")
            ).sum()
        ) if not data_quality.empty else 0
        row["INVALID_INTERVAL_ROWS"] = int(
            (
                data_quality.get("TARGET_MONTH", pd.Series(dtype=str)).astype(str).eq(month)
                & data_quality.get("QUALITY_TYPE", pd.Series(dtype=str)).astype(str).eq("INVALID_INTERVAL")
            ).sum()
        ) if not data_quality.empty else 0
        rows.append(row)
    return pd.DataFrame(rows)


def ltr_write_workbook(env: dict, workbook_path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    orders = {
        "ALL_INFRACTIONS": ["TARGET_MONTH", "EVENT_DATE", "RULE", "SEVERITY", "Collaborateur", "No_collaborateur_codes", "Match_status", "DETAIL", "service_id", "service_date", "service_start", "service_end"],
        "SERVICES_AUDIT": ["service_month", "service_date", "service_start", "service_end", "continues_after_midnight", "continuation_row_count", "creates_worked_day", "Collaborateur", "No_collaborateur_codes", "Match_status", "amplitude_hours", "net_50h_minutes", "net_breaks_minutes", "pause_minutes_inside_service", "attached_calendar_dates", "service_id"],
        "CALENDAR_HOUR_SLICES": ["target_month", "calendar_date", "week_monday", "slice_start", "slice_end", "minutes", "hours", "continuation_from_previous_service", "service_date", "Collaborateur", "No_collaborateur_codes", "service_id", "note"],
        "DATA_QUALITY": ["TARGET_MONTH", "EVENT_DATE", "QUALITY_TYPE", "DETAIL", "Collaborateur", "No collaborateur", "No prestation", "Prestation", "start_dt_local", "end_dt_local", "interval_status", "service_id", "service_date", "continuation_from_previous_service"],
    }
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(workbook_path, engine="openpyxl", mode="w") as writer:
        for sheet_name, df_sheet in sheets.items():
            clean = env["excel_safe_no_tz"](env["prep_for_export"](df_sheet, drop_collaborateur_id=True))
            clean = env["reorder_columns"](clean, orders.get(sheet_name, []))
            clean = env["ensure_unique_columns"](clean)
            clean.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    env["apply_swiss_formats_xlsx"](str(workbook_path))


def ltr_process(matched_upload, rda_upload, output_name: str) -> dict:
    output_root = ltr_unique_output_root(output_name)
    input_dir = output_root / "inputs"
    matched_path = ltr_save_upload(matched_upload, input_dir)
    rda_path = ltr_save_upload(rda_upload, input_dir)
    workbook_path = output_root / "multiple" / "LTR_CHECKS_MULTIPLE_ALL_MONTHS_HYBRID.xlsx"

    notebook_path = LTR_NOTEBOOK_PATH
    env = ltr_load_notebook_functions(notebook_path.stat().st_mtime)

    df_raw = env["load_and_normalize"](str(rda_path))
    matched = env["load_matched_collabs"](str(matched_path))
    df_raw = env["attach_collab_master"](df_raw, matched)
    df = env["add_interval_columns"](df_raw)
    services_df, df_tagged, orphan_pauses_df = env["build_services_and_tagged_rows"](df)
    calendar_slices_df = env["build_calendar_slices"](services_df)

    over50_detail, over50_all = env["check_over_50h"](calendar_slices_df)
    streak_detail, streak_all = env["check_streak_7days"](services_df)
    span_detail, span_all = env["check_span_over_14h"](services_df)
    rest_detail, rest_all, rest_review = env["check_rest_under_11h"](services_df)
    breaks_detail, breaks_all, breaks_audit = env["check_breaks"](services_df)
    data_quality = env["build_data_quality"](df_tagged, orphan_pauses_df)

    infraction_parts = [part for part in [over50_all, streak_all, span_all, rest_all, breaks_all] if part is not None and not part.empty]
    if infraction_parts:
        all_infractions = pd.concat(infraction_parts, ignore_index=True, sort=False)
        all_infractions = env["ensure_unique_columns"](all_infractions)
        sort_cols = [col for col in ["TARGET_MONTH", "EVENT_DATE", "Collaborateur", "RULE"] if col in all_infractions.columns]
        if sort_cols:
            all_infractions = all_infractions.sort_values(sort_cols, kind="stable").reset_index(drop=True)
    else:
        all_infractions = pd.DataFrame(columns=["TARGET_MONTH", "EVENT_DATE", "RULE", "SEVERITY", "DETAIL", "collab_uid", "Collaborateur"])

    summary_by_month = ltr_compute_summary(env, df, services_df, calendar_slices_df, all_infractions, data_quality)
    services_audit = services_df.drop(columns=[col for col in ["_net_50_intervals", "_net_breaks_intervals"] if col in services_df.columns]).copy()
    unmatched_mask = df["collab_key"].astype(str).str.startswith("UNMATCHED")
    ambig_mask = df["collab_match_status"].astype(str).eq("AMBIG_MATCH")
    unmatched_df = df[unmatched_mask | ambig_mask].copy()
    if unmatched_df.empty:
        unrecognized_summary = pd.DataFrame(columns=["Collaborateur", "No collaborateur", "collab_match_status", "collab_key", "Row Count"])
        unrecognized_rows = pd.DataFrame()
    else:
        unrecognized_rows = unmatched_df.copy()
        if "start_dt_local" in unrecognized_rows.columns:
            unrecognized_rows["TARGET_MONTH"] = pd.to_datetime(unrecognized_rows["start_dt_local"], errors="coerce").dt.strftime("%Y-%m")
        unrecognized_summary = (
            unmatched_df.groupby(["Collaborateur", "No collaborateur", "collab_match_status", "collab_key"], dropna=False)
            .size()
            .reset_index(name="Row Count")
            .sort_values("Row Count", ascending=False)
            .reset_index(drop=True)
        )

    sheets = {
        "SUMMARY_BY_MONTH": summary_by_month,
        "ALL_INFRACTIONS": all_infractions,
        "OVER_50H_WEEK": over50_detail,
        "STREAK_7DAYS": streak_detail,
        "SPAN_OVER_14H": span_detail,
        "REST_UNDER_11H": rest_detail,
        "REST_REVIEW_ALLOWED": rest_review,
        "PAUSE_INSUFF": breaks_detail,
        "PAUSE_AUDIT_SERVICES": breaks_audit,
        "SERVICES_AUDIT": services_audit,
        "CALENDAR_HOUR_SLICES": calendar_slices_df,
        "DATA_QUALITY": data_quality,
    }
    ltr_write_workbook(env, workbook_path, sheets)

    return {
        "output_root": output_root,
        "workbook_path": workbook_path,
        "sheets": sheets,
        "summary_by_month": summary_by_month,
        "all_infractions": all_infractions,
        "rest_review": rest_review,
        "data_quality": data_quality,
        "services_audit": services_audit,
        "calendar_slices": calendar_slices_df,
        "breaks_audit": breaks_audit,
        "unrecognized_summary": unrecognized_summary,
        "unrecognized_rows": unrecognized_rows,
        "metrics": {
            "raw_rows": len(df),
            "services": len(services_df),
            "calendar_slices": len(calendar_slices_df),
            "infractions": len(all_infractions),
            "affected_collaborators": all_infractions["collab_uid"].nunique() if "collab_uid" in all_infractions.columns and not all_infractions.empty else 0,
            "rest_review_rows": len(rest_review),
            "data_quality_rows": len(data_quality),
            "unrecognized_rows": int(unrecognized_summary["Row Count"].sum()) if not unrecognized_summary.empty else 0,
        },
    }


def ltr_filtered_df(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        return out
    if filters.get("months") and "TARGET_MONTH" in out.columns:
        out = out[out["TARGET_MONTH"].astype(str).isin(filters["months"])]
    if filters.get("rules") and "RULE" in out.columns:
        out = out[out["RULE"].astype(str).isin(filters["rules"])]
    if filters.get("severities") and "SEVERITY" in out.columns:
        out = out[out["SEVERITY"].astype(str).isin(filters["severities"])]
    if filters.get("collaborators") and "Collaborateur" in out.columns:
        out = out[out["Collaborateur"].astype(str).isin(filters["collaborators"])]
    return out


def ltr_filtered_support_df(df: pd.DataFrame, filters: dict, month_cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        return out
    if filters.get("months"):
        for col in month_cols:
            if col in out.columns:
                out = out[out[col].astype(str).isin(filters["months"])]
                break
    if filters.get("collaborators") and "Collaborateur" in out.columns:
        out = out[out["Collaborateur"].astype(str).isin(filters["collaborators"])]
    return out


def ltr_summary_from_filtered_infractions(filtered: pd.DataFrame) -> pd.DataFrame:
    if filtered.empty or "TARGET_MONTH" not in filtered.columns:
        return pd.DataFrame(columns=["TARGET_MONTH", "TOTAL_INFRACTIONS"])

    summary = filtered.groupby("TARGET_MONTH", dropna=False).size().reset_index(name="TOTAL_INFRACTIONS")
    for rule in sorted(filtered["RULE"].dropna().astype(str).unique().tolist()) if "RULE" in filtered.columns else []:
        counts = (
            filtered[filtered["RULE"].astype(str).eq(rule)]
            .groupby("TARGET_MONTH", dropna=False)
            .size()
            .rename(rule)
        )
        summary = summary.merge(counts, left_on="TARGET_MONTH", right_index=True, how="left")
    return summary.fillna(0)


def ltr_unrecognized_summary(unrecognized_rows: pd.DataFrame) -> pd.DataFrame:
    columns = ["Collaborateur", "No collaborateur", "collab_match_status", "collab_key", "Row Count"]
    if unrecognized_rows.empty:
        return pd.DataFrame(columns=columns)
    return (
        unrecognized_rows.groupby(["Collaborateur", "No collaborateur", "collab_match_status", "collab_key"], dropna=False)
        .size()
        .reset_index(name="Row Count")
        .sort_values("Row Count", ascending=False)
        .reset_index(drop=True)
    )


def render_ltr_chart(title: str, df: pd.DataFrame, x_col: str, y_col: str) -> None:
    st.subheader(title)
    if df.empty or x_col not in df.columns or y_col not in df.columns:
        st.info("Aucune donnée de graphique.")
        return
    chart_df = df[[x_col, y_col]].copy()
    chart_df = chart_df.rename(columns={x_col: "Catégorie", y_col: "Nombre"})
    chart_df["Catégorie"] = chart_df["Catégorie"].astype(str)

    try:
        import altair as alt

        chart = (
            alt.Chart(chart_df)
            .mark_bar()
            .encode(
                x=alt.X("Catégorie:N", title=None, sort=None),
                y=alt.Y("Nombre:Q", title=None),
                tooltip=[
                    alt.Tooltip("Catégorie:N", title="Catégorie"),
                    alt.Tooltip("Nombre:Q", title="Nombre"),
                ],
            )
            .properties(title=None)
        )
        st.altair_chart(chart, use_container_width=True)
    except Exception:
        st.bar_chart(chart_df.set_index("Catégorie"))


def render_ltr_dashboard(result: dict) -> None:
    all_infractions = result["all_infractions"]
    filter_cols = st.columns(3)
    months = sorted(all_infractions["TARGET_MONTH"].dropna().astype(str).unique().tolist()) if "TARGET_MONTH" in all_infractions.columns and not all_infractions.empty else []
    rules = sorted(all_infractions["RULE"].dropna().astype(str).unique().tolist()) if "RULE" in all_infractions.columns and not all_infractions.empty else []
    collaborators = sorted(all_infractions["Collaborateur"].dropna().astype(str).unique().tolist()) if "Collaborateur" in all_infractions.columns and not all_infractions.empty else []
    filters = {
        "months": filter_cols[0].multiselect("Mois", months),
        "rules": filter_cols[1].multiselect("Règle", rules),
        "collaborators": filter_cols[2].multiselect("Collaborateur", collaborators),
    }
    filtered = ltr_filtered_df(all_infractions, filters)
    filtered_services = ltr_filtered_support_df(result["services_audit"], filters, ["service_month", "TARGET_MONTH"])
    filtered_calendar = ltr_filtered_support_df(result["calendar_slices"], filters, ["target_month", "TARGET_MONTH"])
    filtered_rest_review = ltr_filtered_support_df(result["rest_review"], filters, ["TARGET_MONTH"])
    unrecognized_rows = result.get("unrecognized_rows", pd.DataFrame())
    if unrecognized_rows.empty and not result["unrecognized_summary"].empty:
        filtered_unrecognized_summary = ltr_filtered_support_df(result["unrecognized_summary"], filters, [])
        filtered_unrecognized_count = int(filtered_unrecognized_summary["Row Count"].sum()) if "Row Count" in filtered_unrecognized_summary.columns else 0
    else:
        filtered_unrecognized_rows = ltr_filtered_support_df(unrecognized_rows, filters, ["TARGET_MONTH"])
        filtered_unrecognized_summary = ltr_unrecognized_summary(filtered_unrecognized_rows)
        filtered_unrecognized_count = len(filtered_unrecognized_rows)
    filtered_summary = ltr_summary_from_filtered_infractions(filtered)

    metric_cols = st.columns(5)
    metric_cols[0].metric("Infractions", f"{len(filtered):,}")
    affected = filtered["collab_uid"].nunique() if "collab_uid" in filtered.columns and not filtered.empty else 0
    metric_cols[1].metric("Collaborateurs concernés", f"{affected:,}")
    metric_cols[2].metric("Services", f"{len(filtered_services):,}")
    metric_cols[3].metric("Créneaux calendrier", f"{len(filtered_calendar):,}")
    metric_cols[4].metric("Lignes non reconnues", f"{filtered_unrecognized_count:,}")

    chart_cols = st.columns(2)
    with chart_cols[0]:
        render_ltr_chart("Infractions par mois", filtered_summary, "TARGET_MONTH", "TOTAL_INFRACTIONS")
    with chart_cols[1]:
        if filtered.empty or "RULE" not in filtered.columns:
            st.subheader("Infractions par règle")
            st.info("Aucune donnée de graphique.")
        else:
            by_rule = filtered.groupby("RULE").size().reset_index(name="Count")
            render_ltr_chart("Infractions par règle", by_rule, "RULE", "Count")

    if filtered.empty or "Collaborateur" not in filtered.columns:
        st.subheader("Principaux collaborateurs")
        st.info("Aucune donnée de graphique.")
    else:
        top_collabs = filtered.groupby("Collaborateur").size().sort_values(ascending=False).head(15).reset_index(name="Count")
        render_ltr_chart("Principaux collaborateurs", top_collabs, "Collaborateur", "Count")

    st.subheader("Infractions filtrées")
    st.dataframe(filtered, width="stretch", hide_index=True)
    if not filtered.empty:
        st.download_button(
            "Télécharger le CSV des infractions filtrées",
            filtered.to_csv(index=False, encoding="utf-8-sig"),
            file_name="ltr_filtered_infractions.csv",
            mime="text/csv",
        )

    tab_names = ["Résumé", "Revue des repos", "Services", "Audit des pauses", "Créneaux calendrier", "Non reconnus"]
    tabs = st.tabs(tab_names)
    with tabs[0]:
        st.dataframe(filtered_summary, width="stretch", hide_index=True)
    with tabs[1]:
        st.dataframe(filtered_rest_review, width="stretch", hide_index=True)
    with tabs[2]:
        st.dataframe(filtered_services, width="stretch", hide_index=True)
    with tabs[3]:
        st.dataframe(ltr_filtered_support_df(result["breaks_audit"], filters, ["TARGET_MONTH"]), width="stretch", hide_index=True)
    with tabs[4]:
        st.dataframe(filtered_calendar, width="stretch", hide_index=True)
    with tabs[5]:
        st.dataframe(filtered_unrecognized_summary, width="stretch", hide_index=True)


def render_ltr_task() -> None:
    st.title("Contrôles LTR")
    st.caption("Exécute les contrôles LTR hybrides et crée le classeur Excel multi-feuilles avec un tableau de bord d'audit.")
    st.info(
        "Conseil: pour contrôler un mois, incluez le RDA depuis 14 jours avant le début du mois jusqu'à 7 jours "
        "après la fin du mois. Exemple pour mars: du 15 février au 7 avril. Cela permet aux contrôles de repos, "
        "séries de jours, services qui passent minuit et semaines de 50h de voir le contexte autour du mois."
    )

    cols = st.columns(3)
    matched_file = cols[0].file_uploader("Classeur collaborateurs matchés", type=["xlsx", "xls"], key="ltr_matched")
    rda_file = cols[1].file_uploader("Fichier RDA fusionné", type=["xlsx", "xls", "csv"], key="ltr_rda")
    output_name = cols[2].text_input("Nom du dossier de sortie", value="")

    action_cols = st.columns([2, 1])
    with action_cols[0]:
        run_ltr = st.button("Lancer les contrôles LTR", type="primary", disabled=matched_file is None or rda_file is None, width="stretch")

    if run_ltr:
        render_blocking_run_warning()
        progress = st.progress(0.0, text="Démarrage des contrôles LTR")
        try:
            progress.progress(0.1, text="Chargement des fichiers et de la logique notebook")
            result = ltr_process(matched_file, rda_file, output_name)
            progress.progress(1.0, text="Contrôles LTR terminés")
            st.session_state["latest_ltr_result"] = result
        except Exception as exc:
            progress.empty()
            st.exception(exc)
            return

    result = st.session_state.get("latest_ltr_result")
    workbook_path = result["workbook_path"] if result else None
    with action_cols[1]:
        render_download_or_placeholder(workbook_path, "Télécharger le classeur LTR", key="ltr_main_workbook")

    if result:
        st.success("Classeur LTR créé et disponible au téléchargement.")
        render_ltr_dashboard(result)


# ============================================================
# Audit Webfleet-RDA — low-level helpers
# ============================================================
