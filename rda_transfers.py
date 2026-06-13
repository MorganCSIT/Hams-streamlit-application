from app_config import *
from ui_common import read_any_flex, render_blocking_run_warning, render_download_or_placeholder, safe_folder_name


RDA_ALLOWED_EXTENSIONS = ["xlsx", "xls", "csv"]
RDA_DEFAULT_ALLOWED_CODES = "11000,11100,11200,14000,14100,14200"
RDA_DEFAULT_WHITELIST_CODES = "16011,909,16009,195"
RDA_MAIN_FOLDERS = [
    "01_Standard_Transfer",
    "01_All_Collabs_One_CSV",
    "02_Collabs_With_61010_One_CSV",
    "03_Per_Collab_Separate",
    "02_Whitelisted_Ready_For_101",
]


@dataclass
class RdaColumns:
    date: str
    start: str
    end: str
    code: str
    duration: str
    client: str
    collab: str
    collab_name: str | None = None


@dataclass
class RdaRunResult:
    output_dir: Path
    zip_path: Path
    adjusted_df: pd.DataFrame
    folder_summary: pd.DataFrame
    csv_summary: pd.DataFrame
    overlaps: pd.DataFrame
    audit_summary: pd.DataFrame
    generated_files: list[Path]
    minutes_added_by_client: pd.DataFrame | None = None


def _pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    return next((c for c in candidates if c in df.columns), None)


def _norm_code(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (int, np.integer)):
        return str(int(value))
    if isinstance(value, (float, np.floating)):
        if np.isfinite(value) and abs(value - round(value)) < 1e-9:
            return str(int(round(value)))
        text = str(value).strip()
    else:
        text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    if text.isdigit():
        return text
    runs = re.findall(r"\d+", text)
    return sorted(runs, key=len, reverse=True)[0] if runs else text


def _parse_code_set(raw: str, fallback: str = "") -> set[str]:
    text = raw if str(raw).strip() else fallback
    return {_norm_code(part) for part in str(text).split(",") if _norm_code(part)}


def _to_min(value):
    if pd.isna(value):
        return np.nan
    if isinstance(value, pd.Timestamp):
        return int(value.hour) * 60 + int(value.minute)
    if isinstance(value, dtime):
        return int(value.hour) * 60 + int(value.minute)
    if isinstance(value, (float, int, np.integer, np.floating)) and not isinstance(value, bool):
        x = float(value)
        if 0 <= x < 1.0:
            return int(round(x * 24 * 60)) % (24 * 60)
        if 0 <= x < 24 * 60 and abs(x - round(x)) < 1e-9:
            return int(round(x))
        if 0 <= x < 24:
            return int(round(x * 60))
    text = str(value).strip()
    parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
    if not pd.isna(parsed):
        return int(parsed.hour) * 60 + int(parsed.minute)
    token = text.split()[-1]
    parts = token.split(":")
    if len(parts) >= 2 and parts[0].lstrip("+-").isdigit() and parts[1].isdigit():
        return int(parts[0]) * 60 + int(parts[1])
    return np.nan


def _duration_from_min(start_min, end_min):
    if pd.isna(start_min) or pd.isna(end_min):
        return np.nan
    diff = int(end_min) - int(start_min)
    return diff if diff >= 0 else diff + 1440


def _end_abs_min(end_min, start_min):
    if pd.isna(start_min) or pd.isna(end_min):
        return np.nan
    end_i = int(end_min)
    start_i = int(start_min)
    return end_i if end_i >= start_i else end_i + 1440


def _hhmmss_from_abs(min_abs) -> str:
    if pd.isna(min_abs):
        return ""
    minute = int(min_abs) % 1440
    return f"{minute // 60:02d}:{minute % 60:02d}:00"


def _format_date(value) -> str:
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    return "" if pd.isna(parsed) else parsed.strftime("%d.%m.%Y")


def _format_time(value) -> str:
    if pd.isna(value):
        return ""
    minute = _to_min(value)
    if pd.isna(minute):
        return str(value).strip()[:5]
    return f"{int(minute) // 60:02d}:{int(minute) % 60:02d}"


def _detect_columns(df: pd.DataFrame) -> RdaColumns:
    date_col = _pick_col(df, RDA_DATE_COLS)
    start_col = _pick_col(df, RDA_START_COLS)
    end_col = _pick_col(df, RDA_END_COLS)
    code_col = _pick_col(df, RDA_CODE_COLS)
    duration_col = _pick_col(df, RDA_DUREE_COLS)
    client_col = _pick_col(df, RDA_CLIENT_COLS)
    collab_col = _pick_col(df, RDA_COLLAB_COLS)
    missing = [
        name
        for name, col in [
            ("date", date_col),
            ("début", start_col),
            ("fin", end_col),
            ("prestation", code_col),
            ("durée", duration_col),
            ("client", client_col),
            ("collaborateur", collab_col),
        ]
        if col is None
    ]
    if missing:
        raise ValueError(f"Colonnes RDA manquantes: {', '.join(missing)}")
    collab_name_col = next((c for c in ["Collaborateur", "Nom Collaborateur"] if c in df.columns and c != collab_col), None)
    return RdaColumns(date_col, start_col, end_col, code_col, duration_col, client_col, collab_col, collab_name_col)


def _read_uploaded_df(uploaded_file) -> pd.DataFrame:
    return read_any_flex(BytesIO(uploaded_file.getvalue()), uploaded_file.name)


def _make_output_dir(label: str) -> Path:
    root = get_session_output_root(RDA_OUTPUT_FOLDER)
    root.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return root / f"{safe_folder_name(label)}_{timestamp}"


def _normalize_duration_from_times(df: pd.DataFrame, cols: RdaColumns) -> pd.DataFrame:
    out = df.copy()
    start_min = out[cols.start].apply(_to_min)
    end_min = out[cols.end].apply(_to_min)
    calc = [
        _duration_from_min(start, end)
        for start, end in zip(start_min, end_min)
    ]
    if f"{cols.duration}_original" not in out.columns:
        out[f"{cols.duration}_original"] = out[cols.duration]
    out[cols.duration] = pd.Series(calc, index=out.index).where(
        pd.Series(calc, index=out.index).notna(),
        pd.to_numeric(out[cols.duration], errors="coerce"),
    ).astype("Int64")
    return out


def _adjust_61010(df: pd.DataFrame, cols: RdaColumns, allowed_codes: set[str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    df_work = df.copy()
    original_cols = df_work.columns.tolist()
    df_work[cols.code] = df_work[cols.code].apply(_norm_code)
    df_work["Temps retiré"] = 0
    df_work["Temps ajouté"] = 0
    df_work["_start_min"] = df_work[cols.start].apply(_to_min)
    df_work["_end_min"] = df_work[cols.end].apply(_to_min)
    df_work["_start_abs0"] = df_work["_start_min"].astype("Float64")
    df_work["_end_abs0"] = df_work.apply(lambda r: _end_abs_min(r["_end_min"], r["_start_min"]), axis=1).astype("Float64")
    df_work["_dur0"] = df_work.apply(lambda r: _duration_from_min(r["_start_min"], r["_end_min"]), axis=1).astype("Int64")
    df_work["_jour"] = pd.to_datetime(df_work[cols.date], dayfirst=True, errors="coerce").dt.date

    code_61010 = _norm_code(RDA_CODE_61010)
    allowed_norm = {_norm_code(c) for c in allowed_codes}

    def process_block(block: pd.DataFrame) -> pd.DataFrame:
        b = block.sort_values(by=["_start_abs0", "_end_abs0"]).copy().reset_index(drop=False)
        index_col = "index"
        valid = b["_start_abs0"].notna() & b["_end_abs0"].notna() & b["_dur0"].notna()
        valid_idx = b.index[valid].tolist()
        if not valid_idx:
            return b.set_index(index_col).sort_index()

        dur0 = pd.to_numeric(b["_dur0"], errors="coerce").fillna(0).astype(int)
        day_start = int(np.nanmin(b.loc[valid_idx, "_start_abs0"].astype(float)))
        day_end = int(np.nanmax(b.loc[valid_idx, "_end_abs0"].astype(float)))
        day_span = day_end - day_start
        order = b.loc[valid_idx].sort_values(by=["_start_abs0", "_end_abs0"]).index.tolist()
        gaps0 = [
            max(0, int(b.loc[order[k + 1], "_start_abs0"]) - int(b.loc[order[k], "_end_abs0"]))
            for k in range(len(order) - 1)
        ]
        total_gap0 = sum(gaps0)
        dur_new = b["_dur0"].copy()
        target_idx = [i for i in valid_idx if _norm_code(b.loc[i, cols.code]) in allowed_norm]

        def nearest_later_target(i):
            if not target_idx:
                return None
            start_i = float(b.loc[i, "_start_abs0"])
            best = None
            best_dist = float("inf")
            for j in target_idx:
                start_j = float(b.loc[j, "_start_abs0"])
                if start_j >= start_i and start_j - start_i < best_dist:
                    best = j
                    best_dist = start_j - start_i
            return best

        for i in valid_idx:
            if _norm_code(b.loc[i, cols.code]) != code_61010:
                continue
            duration = dur_new.loc[i]
            if pd.isna(duration) or int(duration) <= 15:
                continue
            target = nearest_later_target(i)
            if target is None:
                continue
            surplus = int(duration) - 15
            dur_new.loc[i] = 15
            dur_new.loc[target] = int(dur_new.loc[target]) + surplus

        dur_new_numeric = pd.to_numeric(dur_new, errors="coerce").fillna(0).astype(int)
        delta = dur_new_numeric - dur0
        b["Temps ajouté"] = np.where(delta > 0, delta, 0).astype(int)
        b["Temps retiré"] = np.where(delta < 0, -delta, 0).astype(int)

        work_new = int(dur_new_numeric.loc[valid_idx].sum())
        slack = day_span - work_new
        b["_start_abs"] = pd.NA
        b["_end_abs"] = pd.NA

        if len(order) == 1:
            b.loc[order[0], "_start_abs"] = b.loc[order[0], "_start_abs0"]
            b.loc[order[0], "_end_abs"] = b.loc[order[0], "_end_abs0"]
        elif slack >= 0:
            scaled = [g * slack / total_gap0 for g in gaps0] if total_gap0 > 0 else [slack / (len(order) - 1)] * (len(order) - 1)
            gaps_int = [int(np.floor(x)) for x in scaled]
            remainder = int(slack - sum(gaps_int))
            for k in range(remainder):
                gaps_int[k % (len(order) - 1)] += 1
            cur = day_start
            for pos, i in enumerate(order):
                duration_i = int(dur_new_numeric.loc[i])
                b.loc[i, "_start_abs"] = cur
                b.loc[i, "_end_abs"] = cur + duration_i
                if pos < len(order) - 1:
                    cur = cur + duration_i + gaps_int[pos]
            last_i = order[-1]
            drift = int(b.loc[last_i, "_end_abs"]) - day_end
            if drift != 0:
                b.loc[last_i, "_end_abs"] = int(b.loc[last_i, "_end_abs"]) - drift
        else:
            previous_end = None
            for i in order:
                start_i = int(b.loc[i, "_start_abs0"])
                end_i = start_i + int(dur_new_numeric.loc[i])
                if previous_end is not None and start_i < previous_end:
                    shift = previous_end - start_i
                    start_i += shift
                    end_i += shift
                b.loc[i, "_start_abs"] = start_i
                b.loc[i, "_end_abs"] = end_i
                previous_end = end_i

        for i in valid_idx:
            b.loc[i, cols.start] = _hhmmss_from_abs(b.loc[i, "_start_abs"])
            b.loc[i, cols.end] = _hhmmss_from_abs(b.loc[i, "_end_abs"])
        b[cols.duration] = pd.to_numeric(dur_new, errors="coerce").astype("Int64")
        return b.set_index(index_col).sort_index()

    processed = [process_block(sub) for _, sub in df_work.groupby(["_jour", cols.collab], dropna=False, sort=False)]
    if processed:
        df_full = pd.concat(processed).sort_index()
    else:
        df_full = df_work.copy()

    dur_in = pd.to_numeric(df_work["_dur0"], errors="coerce").fillna(0).astype(int)
    dur_out = pd.to_numeric(df_full[cols.duration], errors="coerce").fillna(0).astype(int)
    delta = dur_out - dur_in
    codes = df_full[cols.code].apply(_norm_code)
    bad_recv = df_full[(delta > 0) & (~codes.isin(allowed_norm))].copy()
    bad_give = df_full[(delta < 0) & (codes != code_61010)].copy()
    if int(delta.sum()) != 0:
        raise ValueError(f"STOP: Sum delta must be 0, got {int(delta.sum())}.")
    if not bad_recv.empty:
        raise ValueError("STOP: minutes were added to codes outside the authorized prestations.")
    if not bad_give.empty:
        raise ValueError("STOP: minutes were removed from codes other than 61010.")

    recap = (
        df_full.assign(_code=df_full[cols.code].apply(_norm_code))
        .groupby("_code")[["Temps ajouté", "Temps retiré"]]
        .sum()
        .query("`Temps ajouté` > 0 or `Temps retiré` > 0")
        .sort_values(["Temps ajouté", "Temps retiré"], ascending=False)
        .reset_index()
    )
    for helper_col in ["_start_min", "_end_min", "_start_abs0", "_end_abs0", "_dur0", "_jour", "_start_abs", "_end_abs"]:
        if helper_col in df_full.columns:
            df_full.drop(columns=helper_col, inplace=True)
    final_cols = original_cols.copy()
    for extra in ["Temps retiré", "Temps ajouté"]:
        if extra not in final_cols:
            final_cols.append(extra)
    out = df_full.reindex(columns=final_cols)
    _assert_duration_matches(out, cols)
    return out, recap


def _assert_duration_matches(df: pd.DataFrame, cols: RdaColumns) -> None:
    check = df.copy()
    start_min = check[cols.start].apply(_to_min)
    end_min = check[cols.end].apply(_to_min)
    from_times = pd.Series([_duration_from_min(s, e) for s, e in zip(start_min, end_min)], index=check.index)
    duration_col = pd.to_numeric(check[cols.duration], errors="coerce")
    bad = check[duration_col.fillna(-999999).astype(int) != from_times.fillna(-999999).astype(int)]
    if not bad.empty:
        raise ValueError(f"STOP: {len(bad)} row(s) have Durée different from Début/Fin.")


def _nexus_df(df: pd.DataFrame, cols: RdaColumns, oe_value: str, client_override: pd.Series | None = None, collab_override: pd.Series | None = None) -> pd.DataFrame:
    clients = client_override if client_override is not None else df[cols.client]
    collabs = collab_override if collab_override is not None else df[cols.collab]
    out = pd.DataFrame(
        {
            "Datum": df[cols.date].apply(_format_date),
            "Von": df[cols.start].apply(_format_time),
            "Bis": df[cols.end].apply(_format_time),
            "Leistungscode": df[cols.code].apply(_norm_code),
            "Dauer_verrechnet": pd.to_numeric(df[cols.duration], errors="coerce").fillna(0).astype(int),
            "OE": oe_value,
            "KD-Nr": pd.to_numeric(clients, errors="coerce").fillna(0).astype(int),
            "Klient": 0,
            "Einsatzgrund": pd.to_numeric(clients, errors="coerce").fillna(0).apply(lambda x: 0 if int(x) == 0 else 2).astype(int),
            "Mitarbeiter-ID": pd.to_numeric(collabs, errors="coerce").fillna(0).astype(int),
        }
    )
    return out


def _batch_text(
    csv_name: str,
    oe_value: str,
    map_rel: str,
    nx_client_path: str = "",
    default_exe_path: str = r"..\nx-spi-client\Asebis.Client.StarterCommand.exe",
) -> str:
    exe = nx_client_path.strip() if nx_client_path.strip() else default_exe_path
    return (
        "@echo off\n"
        "chcp 65001\n"
        "set /p NEXUS_USER=Nexus user: \n"
        "set /p NEXUS_PASSWORD=Nexus password: \n"
        f"\"{exe}\" \"/u=%NEXUS_USER%\" \"/p=%NEXUS_PASSWORD%\" /t=ImportLeistungen_CSV /o={oe_value} "
        f"/f=\"{csv_name}\" /map=\"{map_rel}\" /v\n"
        "Pause\n"
    )


def _write_batch_file(
    path: Path,
    csv_name: str,
    oe_value: str,
    map_rel: str,
    nx_client_path: str = "",
    default_exe_path: str = r"..\nx-spi-client\Asebis.Client.StarterCommand.exe",
) -> None:
    path.write_text(_batch_text(csv_name, oe_value, map_rel, nx_client_path, default_exe_path), encoding="utf-8")


def _write_has_map(df: pd.DataFrame, cols: RdaColumns, path: Path) -> None:
    codes = sorted(c for c in df[cols.code].apply(_norm_code).dropna().astype(str).unique().tolist() if c)
    pd.DataFrame({"Code_ext": codes, "Leistungstarif_nummer": codes}).to_csv(path, index=False, sep=";")


def _write_main_exports(df: pd.DataFrame, cols: RdaColumns, output_dir: Path, oe_value: str, nx_client_path: str = "") -> list[Path]:
    generated: list[Path] = []
    _write_has_map(df, cols, output_dir / "HAS_map_main.csv")
    generated.append(output_dir / "HAS_map_main.csv")

    summary = df.groupby([cols.collab] + ([cols.collab_name] if cols.collab_name else []))[cols.duration].sum().reset_index()
    summary_path = output_dir / "RDA_duree_check.csv"
    summary.to_csv(summary_path, index=False, sep=";")
    generated.append(summary_path)

    all_folder = output_dir / "01_All_Collabs_One_CSV"
    all_folder.mkdir(parents=True, exist_ok=True)
    all_total = int(pd.to_numeric(df[cols.duration], errors="coerce").fillna(0).sum())
    all_csv = all_folder / f"RDA_AllCollabs+{all_total}.csv"
    _nexus_df(df, cols, oe_value).to_csv(all_csv, index=False, sep=";")
    all_bat = all_folder / "RDA_AllCollabs_batch.bat"
    _write_batch_file(all_bat, all_csv.name, oe_value, r"..\HAS_map_main.csv", nx_client_path)
    generated.extend([all_csv, all_bat])

    folder_61010 = output_dir / "02_Collabs_With_61010_One_CSV"
    folder_61010.mkdir(parents=True, exist_ok=True)
    collabs_61010 = df.loc[df[cols.code].apply(_norm_code) == _norm_code(RDA_CODE_61010), cols.collab].unique().tolist()
    if collabs_61010:
        df_61010 = df[df[cols.collab].isin(collabs_61010)].copy()
        total_61010 = int(pd.to_numeric(df_61010[cols.duration], errors="coerce").fillna(0).sum())
        csv_61010 = folder_61010 / f"RDA_CollabsWith61010+{total_61010}.csv"
        _nexus_df(df_61010, cols, oe_value).to_csv(csv_61010, index=False, sep=";")
        bat_61010 = folder_61010 / "RDA_CollabsWith61010_batch.bat"
        _write_batch_file(bat_61010, csv_61010.name, oe_value, r"..\HAS_map_main.csv", nx_client_path)
        generated.extend([csv_61010, bat_61010])

    per_folder = output_dir / "03_Per_Collab_Separate"
    per_folder.mkdir(parents=True, exist_ok=True)
    for collab_id, group in df.groupby(cols.collab):
        total = int(pd.to_numeric(group[cols.duration], errors="coerce").fillna(0).sum())
        identifier = safe_folder_name(collab_id)
        if cols.collab_name and not group[cols.collab_name].dropna().empty:
            identifier = f"{identifier}-{safe_folder_name(group[cols.collab_name].dropna().iloc[0])}"
        collab_folder = per_folder / f"RDA-{identifier}+{total}"
        collab_folder.mkdir(parents=True, exist_ok=True)
        csv_path = collab_folder / f"{identifier}+{total}.csv"
        _nexus_df(group, cols, oe_value).to_csv(csv_path, index=False, sep=";")
        bat_path = collab_folder / f"{identifier}_batch.bat"
        _write_batch_file(
            bat_path,
            csv_path.name,
            oe_value,
            r"..\..\..\HAS_map_main.csv",
            nx_client_path,
            r"..\..\nx-spi-client\Asebis.Client.StarterCommand.exe",
        )
        generated.extend([csv_path, bat_path])
    return generated


def _norm_mapping_col(value) -> str:
    import unicodedata

    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    text = text.strip().lower().replace("#", "no")
    text = re.sub(r"n\s*o", "no", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return re.sub(r"_+", "_", text).strip("_")


def _mapping_sheet(sheets: dict[str, pd.DataFrame], role: str) -> pd.DataFrame:
    tokens = ["client", "klient", "kd"] if role == "client" else ["collab", "collaborateur", "mitarbeiter", "employee"]
    for name, df in sheets.items():
        normalized = _norm_mapping_col(name)
        if any(token in normalized for token in tokens):
            return df
    return next(iter(sheets.values()))


def _mapping_column(df: pd.DataFrame, role: str, uo_label: str) -> str:
    uo_tokens = [_norm_mapping_col(part) for part in re.split(r"\s+", uo_label) if part]
    role_tokens = ["client", "klient", "kd"] if role == "client" else ["collab", "collaborateur", "mitarbeiter", "employee"]
    scored = []
    for col in df.columns:
        key = _norm_mapping_col(col)
        score = sum(token in key for token in uo_tokens) * 3 + sum(token in key for token in role_tokens)
        if score:
            scored.append((score, col))
    if not scored:
        numeric_cols = [c for c in df.columns if pd.to_numeric(df[c].astype(str).str.extract(r"(\d+)")[0], errors="coerce").notna().sum() > 0]
        if len(numeric_cols) >= 2:
            return numeric_cols[0 if role == "client" else min(1, len(numeric_cols) - 1)]
        raise ValueError(f"Could not auto-detect mapping column for {role} / {uo_label}.")
    scored.sort(reverse=True, key=lambda item: item[0])
    return scored[0][1]


def _numeric_ids(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.extract(r"(\d+)")[0], errors="coerce").astype("Int64")


def _load_mapping(uploaded_mapping, source_label: str, target_label: str) -> tuple[dict[int, int], dict[int, int], pd.DataFrame]:
    if uploaded_mapping is None:
        if source_label == target_label:
            summary = pd.DataFrame(
                [["Mapping mode", "Identity mapping: source and target UO are the same."]],
                columns=["Item", "Value"],
            )
            return {}, {}, summary
        raise ValueError("A mapping file is required when source UO and target UO are different.")
    sheets = pd.read_excel(BytesIO(uploaded_mapping.getvalue()), sheet_name=None)
    clients_df = _mapping_sheet(sheets, "client")
    collabs_df = _mapping_sheet(sheets, "collab")
    src_client_col = _mapping_column(clients_df, "client", source_label)
    tgt_client_col = _mapping_column(clients_df, "client", target_label)
    src_collab_col = _mapping_column(collabs_df, "collab", source_label)
    tgt_collab_col = _mapping_column(collabs_df, "collab", target_label)
    client_map = {
        int(src): int(tgt)
        for src, tgt in zip(_numeric_ids(clients_df[src_client_col]), _numeric_ids(clients_df[tgt_client_col]))
        if not pd.isna(src) and not pd.isna(tgt)
    }
    collab_map = {
        int(src): int(tgt)
        for src, tgt in zip(_numeric_ids(collabs_df[src_collab_col]), _numeric_ids(collabs_df[tgt_collab_col]))
        if not pd.isna(src) and not pd.isna(tgt)
    }
    summary = pd.DataFrame(
        [
            ["Client source", src_client_col],
            ["Client target", tgt_client_col],
            ["Collaborator source", src_collab_col],
            ["Collaborator target", tgt_collab_col],
            ["Client mappings", len(client_map)],
            ["Collaborator mappings", len(collab_map)],
        ],
        columns=["Item", "Value"],
    )
    return client_map, collab_map, summary


def _write_transfer_export(
    df: pd.DataFrame,
    cols: RdaColumns,
    output_dir: Path,
    source_label: str,
    target_label: str,
    mapping_file,
    nx_client_path: str = "",
    folder_name: str = "02_Whitelisted_Ready_For_101",
    export_stem: str = "RDA_Whitelisted_Ready_For_101",
) -> tuple[list[Path], pd.DataFrame, pd.DataFrame]:
    generated: list[Path] = []
    folder = output_dir / folder_name
    folder.mkdir(parents=True, exist_ok=True)
    if df.empty:
        return generated, pd.DataFrame(), pd.DataFrame()

    client_map, collab_map, mapping_summary = _load_mapping(mapping_file, source_label, target_label)
    source_clients = _numeric_ids(df[cols.client])
    source_collabs = _numeric_ids(df[cols.collab])
    if source_label == target_label and not client_map and not collab_map:
        mapped_clients = source_clients.fillna(0).astype(int)
        mapped_collabs = source_collabs.fillna(0).astype(int)
    else:
        mapped_clients = source_clients.map(client_map).fillna(0).astype(int)
        mapped_collabs = source_collabs.map(collab_map).fillna(0).astype(int)

    target_oe = RDA_OE_MAP[target_label]
    csv_df = _nexus_df(df, cols, target_oe, mapped_clients, mapped_collabs)
    total = int(csv_df["Dauer_verrechnet"].fillna(0).sum())
    csv_path = folder / f"{export_stem}+{total}.csv"
    csv_df.to_csv(csv_path, index=False, sep=";")
    map_path = output_dir / "HAS_map.csv"
    _write_has_map(df, cols, map_path)
    batch_path = folder / f"{export_stem}_batch.bat"
    _write_batch_file(batch_path, csv_path.name, target_oe, r"..\HAS_map.csv", nx_client_path)
    qa_path = folder / f"{export_stem}_QA.xlsx"

    unmapped_clients = df[source_clients.notna() & (source_clients != 0) & (mapped_clients == 0)].copy()
    unmapped_collabs = df[source_collabs.notna() & (source_collabs != 0) & (mapped_collabs == 0)].copy()
    with pd.ExcelWriter(qa_path, engine="openpyxl") as writer:
        mapping_summary.to_excel(writer, index=False, sheet_name="Summary")
        unmapped_clients.to_excel(writer, index=False, sheet_name="Unmapped clients")
        unmapped_collabs.to_excel(writer, index=False, sheet_name="Unmapped collabs")
    generated.extend([csv_path, batch_path, map_path, qa_path])
    qa_summary = pd.DataFrame(
        [
            [f"{export_stem} rows", len(csv_df)],
            [f"{export_stem} minutes", total],
            ["Unmapped clients", len(unmapped_clients)],
            ["Unmapped collaborators", len(unmapped_collabs)],
        ],
        columns=["Check", "Value"],
    )
    return generated, qa_summary, csv_df


def _write_adjusted_outputs(raw_df: pd.DataFrame, adjusted_df: pd.DataFrame, cols: RdaColumns, output_dir: Path, base_name: str) -> list[Path]:
    generated: list[Path] = []
    adjusted_path = output_dir / f"{safe_folder_name(base_name)}_61010_adjusted_no_overlap.xlsx"
    adjusted_df.to_excel(adjusted_path, index=False)
    generated.append(adjusted_path)

    raw_compare = raw_df[[c for c in [cols.date, cols.collab, cols.code, cols.start, cols.end, cols.duration] if c in raw_df.columns]].copy()
    adj_compare = adjusted_df[[c for c in [cols.date, cols.collab, cols.code, cols.start, cols.end, cols.duration, "Temps retiré", "Temps ajouté"] if c in adjusted_df.columns]].copy()
    diff_path = output_dir / "RDA_StartEnd_Differences_RAW_vs_ADJ.xlsx"
    with pd.ExcelWriter(diff_path, engine="openpyxl") as writer:
        raw_compare.to_excel(writer, index=False, sheet_name="RAW")
        adj_compare.to_excel(writer, index=False, sheet_name="ADJUSTED")
    generated.append(diff_path)

    all_collabs = adjusted_df[[cols.collab] + ([cols.collab_name] if cols.collab_name else [])].drop_duplicates()
    collabs_61010 = adjusted_df[adjusted_df[cols.code].apply(_norm_code) == _norm_code(RDA_CODE_61010)][all_collabs.columns].drop_duplicates()
    all_clients = pd.DataFrame(adjusted_df[cols.client].drop_duplicates(), columns=[cols.client])
    summary_path = output_dir / "Collaborator_Client_Summary.xlsx"
    with pd.ExcelWriter(summary_path, engine="openpyxl") as writer:
        all_collabs.to_excel(writer, sheet_name="All Collaborators", index=False)
        collabs_61010.to_excel(writer, sheet_name="61010 Collaborators", index=False)
        all_clients.to_excel(writer, sheet_name="All Clients", index=False)
    generated.append(summary_path)
    return generated


def _find_overlaps(df: pd.DataFrame, cols: RdaColumns) -> pd.DataFrame:
    check = df.copy()
    check["_jour_chk"] = pd.to_datetime(check[cols.date], dayfirst=True, errors="coerce").dt.date
    check["_start_chk"] = check[cols.start].apply(_to_min)
    check["_end_chk"] = [
        _end_abs_min(end, start)
        for start, end in zip(check["_start_chk"], check[cols.end].apply(_to_min))
    ]
    overlaps = []
    for (day, collab), sub in check.groupby(["_jour_chk", cols.collab], dropna=False, sort=False):
        block = sub.sort_values(["_start_chk", "_end_chk"])
        rows = list(block.index)
        for pos in range(len(rows) - 1):
            left = block.loc[rows[pos]]
            right = block.loc[rows[pos + 1]]
            if pd.notna(left["_end_chk"]) and pd.notna(right["_start_chk"]) and int(left["_end_chk"]) > int(right["_start_chk"]):
                overlaps.append(
                    {
                        "Jour": day,
                        "Collaborateur": collab,
                        "Row1": rows[pos],
                        "Row1_Start": left[cols.start],
                        "Row1_End": left[cols.end],
                        "Row1_Code": left[cols.code],
                        "Row2": rows[pos + 1],
                        "Row2_Start": right[cols.start],
                        "Row2_End": right[cols.end],
                        "Row2_Code": right[cols.code],
                    }
                )
    return pd.DataFrame(overlaps)


def _audit_generated_csvs(output_dir: Path, adjusted_df: pd.DataFrame, cols: RdaColumns) -> tuple[pd.DataFrame, pd.DataFrame]:
    folder_rows = []
    csv_rows = []
    for folder_name in RDA_MAIN_FOLDERS:
        folder = output_dir / folder_name
        csv_files = sorted(folder.rglob("*.csv")) if folder.exists() else []
        row_count = 0
        duration_sum = 0
        time_sum = 0
        mismatch_rows = 0
        audited_files = 0
        for csv_path in csv_files:
            try:
                csv_df = pd.read_csv(csv_path, sep=";")
            except Exception:
                continue
            if not {"Von", "Bis", "Dauer_verrechnet"}.issubset(csv_df.columns):
                continue
            audited_files += 1
            durations = pd.to_numeric(csv_df["Dauer_verrechnet"], errors="coerce").fillna(0).astype(int)
            starts = csv_df["Von"].apply(_to_min)
            ends = csv_df["Bis"].apply(_to_min)
            calc = pd.Series([_duration_from_min(s, e) for s, e in zip(starts, ends)]).fillna(0).astype(int)
            mismatches = int((durations.reset_index(drop=True) != calc.reset_index(drop=True)).sum())
            row_count += len(csv_df)
            duration_sum += int(durations.sum())
            time_sum += int(calc.sum())
            mismatch_rows += mismatches
            csv_rows.append(
                {
                    "Folder": folder_name,
                    "CSV file": str(csv_path.relative_to(output_dir)),
                    "Rows": len(csv_df),
                    "Sum Dauer_verrechnet": int(durations.sum()),
                    "Sum calculated Von/Bis": int(calc.sum()),
                    "Difference": int(durations.sum() - calc.sum()),
                    "Mismatch rows": mismatches,
                }
            )
        folder_rows.append(
            {
                "Folder": folder_name,
                "Folder exists": folder.exists(),
                "CSV files": audited_files,
                "Rows": row_count,
                "Sum Dauer_verrechnet": duration_sum,
                "Sum calculated Von/Bis": time_sum,
                "Difference": duration_sum - time_sum,
                "Mismatch rows": mismatch_rows,
            }
        )
    return pd.DataFrame(folder_rows), pd.DataFrame(csv_rows)


def _write_zip(output_dir: Path) -> Path:
    zip_path = output_dir.with_suffix(".zip")
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in output_dir.rglob("*"):
            if path.is_file():
                zf.write(path, Path(output_dir.name) / path.relative_to(output_dir))
    return zip_path


def _collect_files(output_dir: Path, zip_path: Path) -> list[Path]:
    files = sorted(path for path in output_dir.rglob("*") if path.is_file())
    if zip_path.exists():
        files.append(zip_path)
    return files


def _standard_transfer_run(df: pd.DataFrame, mapping_file, source_label: str, target_label: str, output_name: str, nx_client_path: str = "") -> RdaRunResult:
    output_dir = _make_output_dir(output_name)
    output_dir.mkdir(parents=True, exist_ok=True)
    cols = _detect_columns(df)
    normalized = _normalize_duration_from_times(df, cols)
    generated, transfer_qa, _ = _write_transfer_export(
        normalized,
        cols,
        output_dir,
        source_label,
        target_label,
        mapping_file,
        nx_client_path,
        folder_name="01_Standard_Transfer",
        export_stem="RDA_Standard_Transfer",
    )
    # Also create a main HAS map so the standard package has the same import skeleton.
    _write_has_map(normalized, cols, output_dir / "HAS_map_main.csv")
    generated.append(output_dir / "HAS_map_main.csv")
    rda_path = output_dir / "RDA_Standard_Transfer_Source.xlsx"
    normalized.to_excel(rda_path, index=False)
    generated.append(rda_path)
    overlaps = _find_overlaps(normalized, cols)
    folder_summary, csv_summary = _audit_generated_csvs(output_dir, normalized, cols)
    audit_summary = pd.concat([transfer_qa], ignore_index=True) if not transfer_qa.empty else pd.DataFrame()
    zip_path = _write_zip(output_dir)
    return RdaRunResult(output_dir, zip_path, normalized, folder_summary, csv_summary, overlaps, audit_summary, _collect_files(output_dir, zip_path))


def _transfer_excluding_codes_run(
    df: pd.DataFrame,
    mapping_file,
    source_label: str,
    target_label: str,
    output_name: str,
    excluded_codes: set[str],
    nx_client_path: str = "",
) -> RdaRunResult:
    output_dir = _make_output_dir(output_name)
    output_dir.mkdir(parents=True, exist_ok=True)
    cols = _detect_columns(df)
    normalized = _normalize_duration_from_times(df, cols)
    excluded_norm = {_norm_code(code) for code in excluded_codes if _norm_code(code)}
    if excluded_norm:
        keep_mask = ~normalized[cols.code].apply(_norm_code).isin(excluded_norm)
    else:
        keep_mask = pd.Series(True, index=normalized.index)
    transfer_df = normalized.loc[keep_mask].copy()
    excluded_df = normalized.loc[~keep_mask].copy()

    generated, transfer_qa, _ = _write_transfer_export(
        transfer_df,
        cols,
        output_dir,
        source_label,
        target_label,
        mapping_file,
        nx_client_path,
        folder_name="01_Standard_Transfer",
        export_stem="RDA_UO_Transfer",
    )
    _write_has_map(transfer_df, cols, output_dir / "HAS_map_main.csv")
    generated.append(output_dir / "HAS_map_main.csv")
    source_path = output_dir / "RDA_UO_Transfer_Source.xlsx"
    normalized.to_excel(source_path, index=False)
    generated.append(source_path)
    if not excluded_df.empty:
        excluded_path = output_dir / "RDA_UO_Transfer_Excluded.xlsx"
        excluded_df.to_excel(excluded_path, index=False)
        generated.append(excluded_path)

    overlaps = _find_overlaps(transfer_df, cols)
    folder_summary, csv_summary = _audit_generated_csvs(output_dir, transfer_df, cols)
    audit_summary = pd.DataFrame(
        [
            ["Raw rows", len(normalized)],
            ["Excluded rows", len(excluded_df)],
            ["Transferred rows", len(transfer_df)],
            ["Raw minutes", int(pd.to_numeric(normalized[cols.duration], errors="coerce").fillna(0).sum())],
            ["Excluded minutes", int(pd.to_numeric(excluded_df[cols.duration], errors="coerce").fillna(0).sum())],
            ["Transferred minutes", int(pd.to_numeric(transfer_df[cols.duration], errors="coerce").fillna(0).sum())],
        ],
        columns=["Check", "Value"],
    )
    if not transfer_qa.empty:
        audit_summary = pd.concat([audit_summary, transfer_qa], ignore_index=True)
    zip_path = _write_zip(output_dir)
    return RdaRunResult(output_dir, zip_path, transfer_df, folder_summary, csv_summary, overlaps, audit_summary, _collect_files(output_dir, zip_path))


def _adjustment_run(
    df: pd.DataFrame,
    mapping_file,
    source_label: str,
    target_label: str,
    output_name: str,
    allowed_codes: set[str],
    include_whitelist_transfer: bool,
    whitelist_codes: set[str],
    nx_client_path: str = "",
) -> RdaRunResult:
    output_dir = _make_output_dir(output_name)
    output_dir.mkdir(parents=True, exist_ok=True)
    cols = _detect_columns(df)
    normalized = _normalize_duration_from_times(df, cols)
    whitelist_mask = normalized[cols.code].apply(_norm_code).isin(whitelist_codes) if whitelist_codes else pd.Series(False, index=normalized.index)
    whitelist_df = normalized.loc[whitelist_mask].copy()
    main_df = normalized.loc[~whitelist_mask].copy()
    adjusted_df, allocation_recap = _adjust_61010(main_df, cols, allowed_codes)

    generated = _write_adjusted_outputs(main_df, adjusted_df, cols, output_dir, output_name)
    generated.extend(_write_main_exports(adjusted_df, cols, output_dir, RDA_OE_MAP[source_label], nx_client_path))
    transfer_qa = pd.DataFrame()
    if include_whitelist_transfer:
        transfer_generated, transfer_qa, _ = _write_transfer_export(whitelist_df, cols, output_dir, source_label, target_label, mapping_file, nx_client_path)
        generated.extend(transfer_generated)

    overlaps = _find_overlaps(adjusted_df, cols)
    folder_summary, csv_summary = _audit_generated_csvs(output_dir, adjusted_df, cols)
    total_raw = int(pd.to_numeric(normalized[cols.duration], errors="coerce").fillna(0).sum())
    total_adjusted = int(pd.to_numeric(adjusted_df[cols.duration], errors="coerce").fillna(0).sum())
    code_61010 = _norm_code(RDA_CODE_61010)
    raw_61010 = normalized[normalized[cols.code].apply(_norm_code) == code_61010]
    adjusted_61010 = adjusted_df[adjusted_df[cols.code].apply(_norm_code) == code_61010]
    total_61010_before = int(pd.to_numeric(raw_61010[cols.duration], errors="coerce").fillna(0).sum())
    total_61010_after = int(pd.to_numeric(adjusted_61010[cols.duration], errors="coerce").fillna(0).sum())
    collabs_with_61010 = int(raw_61010[cols.collab].nunique(dropna=True))
    minutes_added_by_client = pd.DataFrame()
    if "Temps ajouté" in adjusted_df.columns:
        added_rows = adjusted_df[pd.to_numeric(adjusted_df["Temps ajouté"], errors="coerce").fillna(0) > 0].copy()
        if not added_rows.empty:
            minutes_added_by_client = (
                added_rows.assign(
                    Client=added_rows[cols.client].apply(_norm_code),
                    **{"Minutes ajoutées": pd.to_numeric(added_rows["Temps ajouté"], errors="coerce").fillna(0).astype(int)},
                )
                .groupby("Client", dropna=False, as_index=False)["Minutes ajoutées"]
                .sum()
                .sort_values("Minutes ajoutées", ascending=False)
                .reset_index(drop=True)
            )
    audit_summary = pd.DataFrame(
        [
            ["Total 61010 minutes before", total_61010_before],
            ["Total 61010 minutes after", total_61010_after],
            ["Collaborators with 61010", collabs_with_61010],
            ["Raw rows", len(normalized)],
            ["Main adjusted rows", len(adjusted_df)],
            ["Whitelisted transfer rows", len(whitelist_df)],
            ["Raw minutes", total_raw],
            ["Adjusted main minutes", total_adjusted],
            ["Temps retiré", int(adjusted_df.get("Temps retiré", pd.Series(dtype=int)).sum())],
            ["Temps ajouté", int(adjusted_df.get("Temps ajouté", pd.Series(dtype=int)).sum())],
            ["Overlap pairs", len(overlaps)],
        ],
        columns=["Check", "Value"],
    )
    if not allocation_recap.empty:
        recap_path = output_dir / "RDA_61010_allocation_recap.xlsx"
        allocation_recap.to_excel(recap_path, index=False)
        generated.append(recap_path)
    if not transfer_qa.empty:
        audit_summary = pd.concat([audit_summary, transfer_qa], ignore_index=True)
    zip_path = _write_zip(output_dir)
    return RdaRunResult(
        output_dir,
        zip_path,
        adjusted_df,
        folder_summary,
        csv_summary,
        overlaps,
        audit_summary,
        _collect_files(output_dir, zip_path),
        minutes_added_by_client,
    )


def _result_downloads(result: RdaRunResult, key_prefix: str) -> None:
    st.subheader("Téléchargements")
    with result.zip_path.open("rb") as handle:
        st.download_button(
            "Télécharger le dossier complet (.zip)",
            handle,
            file_name=result.zip_path.name,
            mime="application/zip",
            key=f"{key_prefix}_zip",
        )
    rows = []
    for path in result.generated_files:
        rows.append({"Fichier": str(path.relative_to(result.output_dir.parent)), "Taille KB": round(path.stat().st_size / 1024, 1)})
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True)
        selected = st.multiselect(
            "Fichiers individuels",
            options=[str(path.relative_to(result.output_dir)) for path in result.generated_files if path.is_file() and path != result.zip_path],
            key=f"{key_prefix}_selected_files",
        )
        for rel in selected:
            path = result.output_dir / rel
            mime = "text/csv" if path.suffix.lower() == ".csv" else "application/octet-stream"
            with path.open("rb") as handle:
                st.download_button(
                    f"Télécharger {path.name}",
                    handle,
                    file_name=path.name,
                    mime=mime,
                    key=f"{key_prefix}_download_{rel}",
                )


def _audit_value(result: RdaRunResult, check_name: str):
    if result.audit_summary.empty or not {"Check", "Value"}.issubset(result.audit_summary.columns):
        return None
    matches = result.audit_summary.loc[result.audit_summary["Check"] == check_name, "Value"]
    return None if matches.empty else matches.iloc[0]


def _render_result(result: RdaRunResult, key_prefix: str) -> None:
    st.success("Dossier RDA généré. Téléchargez le ZIP complet pour récupérer les fichiers Nexus.")
    total_61010_before = _audit_value(result, "Total 61010 minutes before")
    total_61010_after = _audit_value(result, "Total 61010 minutes after")
    collabs_with_61010 = _audit_value(result, "Collaborators with 61010")
    if all(value is not None for value in [total_61010_before, total_61010_after, collabs_with_61010]):
        metric_cols = st.columns(3)
        metric_cols[0].metric("Total 61010 minutes before", total_61010_before)
        metric_cols[1].metric("Total 61010 minutes after", total_61010_after)
        metric_cols[2].metric("Number of collabs with at least one 61010", collabs_with_61010)
    minutes_added_by_client = getattr(result, "minutes_added_by_client", None)
    if minutes_added_by_client is not None and not minutes_added_by_client.empty:
        st.subheader("Minutes ajoutées par client")
        st.dataframe(minutes_added_by_client, use_container_width=True, hide_index=True)
    tabs = st.tabs(["Résumé"])
    with tabs[0]:
        st.subheader("QA")
        if not result.audit_summary.empty:
            st.dataframe(result.audit_summary, use_container_width=True)
        else:
            st.info("Aucun résumé QA disponible.")
        st.subheader("CSV audit")
        if result.csv_summary.empty:
            st.info("Aucun CSV Nexus auditable trouvé.")
        else:
            st.dataframe(result.csv_summary, use_container_width=True)
        st.subheader("Overlaps")
        if result.overlaps.empty:
            st.success("Aucun overlap détecté.")
        else:
            st.warning(f"{len(result.overlaps)} overlap(s) détecté(s).")
            st.dataframe(result.overlaps, use_container_width=True)


def _uo_options() -> tuple[list[str], int]:
    options = list(RDA_OE_MAP.keys())
    default_idx = options.index("SA 101") if "SA 101" in options else 0
    return options, default_idx


def _adjustment_inputs(key_prefix: str):
    rda_file = st.file_uploader("Fichier RDA à ajuster", type=RDA_ALLOWED_EXTENSIONS, key=f"{key_prefix}_rda")
    config_cols = st.columns(2)
    source_label = config_cols[0].selectbox("UO du fichier RDA", list(RDA_OE_MAP.keys()), key=f"{key_prefix}_source")
    output_name = config_cols[1].text_input("Nom du dossier", value=f"RDA {datetime.now().strftime('%m%Y')}", key=f"{key_prefix}_output")
    nx_path = ""
    ready = rda_file is not None
    if not ready:
        st.info("Ajoutez le fichier RDA pour lancer l'ajustement.")
    return rda_file, source_label, output_name, nx_path, ready


def _transfer_inputs(key_prefix: str):
    upload_cols = st.columns(2)
    rda_file = upload_cols[0].file_uploader("Fichier RDA à transférer", type=RDA_ALLOWED_EXTENSIONS, key=f"{key_prefix}_rda")
    mapping_file = upload_cols[1].file_uploader("Fichier Mapping", type=["xlsx", "xls"], key=f"{key_prefix}_mapping")
    target_options, default_target_idx = _uo_options()
    config_cols = st.columns(3)
    source_label = config_cols[0].selectbox("UO source", list(RDA_OE_MAP.keys()), key=f"{key_prefix}_source")
    target_label = config_cols[1].selectbox("UO cible", target_options, index=default_target_idx, key=f"{key_prefix}_target")
    output_name = config_cols[2].text_input("Nom du dossier", value=f"RDA {datetime.now().strftime('%m%Y')}", key=f"{key_prefix}_output")
    nx_path = ""
    needs_mapping = source_label != target_label
    ready = rda_file is not None and (mapping_file is not None or not needs_mapping)
    if not ready:
        if rda_file is None:
            st.info("Ajoutez le fichier RDA pour préparer le transfert.")
        elif needs_mapping:
            st.info("Ajoutez le fichier Mapping pour un transfert entre deux UO différentes.")
    elif source_label == target_label:
        st.caption("Mapping optionnel: la source et la cible sont identiques, les IDs client/collab sont conservés.")
    return rda_file, mapping_file, source_label, target_label, output_name, nx_path, ready


def _run_button(label: str, key: str) -> bool:
    return st.button(label, type="primary", key=key)


def _store_result(key: str, result: RdaRunResult) -> None:
    st.session_state[key] = result


def _show_stored_result(key: str, key_prefix: str) -> None:
    result = st.session_state.get(key)
    if result is not None:
        _render_result(result, key_prefix)


def _render_adjustment_section() -> None:
    st.subheader("Ajustement 15 minutes")
    st.caption("Reconstruit le RDA complet: les prestations 61010 de plus de 15 minutes sont réduites, et le surplus est ajouté aux prestations autorisées.")
    rda_file, source_label, output_name, nx_path, ready = _adjustment_inputs("rda_adjustment")
    allowed_codes = _parse_code_set(
        st.text_input(
            "Prestations autorisées à recevoir le surplus",
            RDA_DEFAULT_ALLOWED_CODES,
            key="rda_adjustment_allowed",
        ),
        RDA_DEFAULT_ALLOWED_CODES,
    )
    result = st.session_state.get("rda_adjustment_result")
    action_cols = st.columns([2, 1])
    with action_cols[0]:
        run_adjustment = st.button("Ajuster le RDA et générer les fichiers Nexus", type="primary", key="rda_adjustment_run", disabled=not ready, width="stretch")

    if run_adjustment:
        render_blocking_run_warning()
        with st.spinner("Ajustement 61010 en cours..."):
            df = _read_uploaded_df(rda_file)
            result = _adjustment_run(df, None, source_label, source_label, output_name, allowed_codes, False, set(), nx_path)
            _store_result("rda_adjustment_result", result)
    result = st.session_state.get("rda_adjustment_result")
    with action_cols[1]:
        render_download_or_placeholder(result.zip_path if result else None, "Télécharger le dossier complet (.zip)", key="rda_adjustment_zip")
    _show_stored_result("rda_adjustment_result", "rda_adjustment")


def _render_uo_transfer_section() -> None:
    st.subheader("Transfert UO vers UO")
    st.caption("Prépare un transfert Nexus vers l'UO cible.")
    rda_file, mapping_file, source_label, target_label, output_name, nx_path, ready = _transfer_inputs("rda_transfer")
    result = st.session_state.get("rda_transfer_result")
    action_cols = st.columns([2, 1])
    with action_cols[0]:
        run_transfer = st.button("Préparer le transfert UO vers UO", type="primary", key="rda_transfer_run", disabled=not ready, width="stretch")

    if run_transfer:
        render_blocking_run_warning()
        with st.spinner("Préparation du transfert UO vers UO..."):
            df = _read_uploaded_df(rda_file)
            result = _transfer_excluding_codes_run(df, mapping_file, source_label, target_label, output_name, set(), nx_path)
            _store_result("rda_transfer_result", result)
    result = st.session_state.get("rda_transfer_result")
    with action_cols[1]:
        render_download_or_placeholder(result.zip_path if result else None, "Télécharger le dossier complet (.zip)", key="rda_transfer_zip")
    _show_stored_result("rda_transfer_result", "rda_transfer")


def render_rda_task() -> None:
    st.title("Transferts RDA")
    st.caption("Choisissez si vous voulez reconstruire le RDA en 15 minutes ou préparer un transfert entre deux UO. Les batchs Nexus sont générés, mais pas exécutés depuis l'application.")
    tabs = st.tabs(["Ajustement 15 minutes", "Transfert UO vers UO"])
    with tabs[0]:
        _render_adjustment_section()
    with tabs[1]:
        _render_uo_transfer_section()
