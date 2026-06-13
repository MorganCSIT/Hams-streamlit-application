from app_config import *
import concurrent.futures

from ui_common import read_csv_flex, render_blocking_run_warning, render_download_or_placeholder


def rda_pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    return next((c for c in candidates if c in df.columns), None)


def audit_to_int_str(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"\s+", "", s)
    try:
        return str(int(float(s)))
    except Exception:
        return s if s else None


def audit_extract_code_any(x):
    if pd.isna(x):
        return np.nan
    m = re.search(r"(\d{4,6})", str(x))
    return m.group(1) if m else np.nan


def audit_clean_legend_text(x):
    if pd.isna(x):
        return ""
    s = re.sub(r"\s+", " ", str(x).strip())
    return "" if s.lower() in {"nan", "nat", "none", "<na>"} else s


def audit_shorten_text(text, max_len):
    text = audit_clean_legend_text(text)
    if len(text) <= max_len:
        return text
    return text[:max(0, max_len - 1)].rstrip() + "…"


def audit_swiss_date(series_like):
    return pd.to_datetime(series_like, errors="coerce", dayfirst=True).dt.date


def audit_swiss_dt(series_like, tz_name=AUDIT_TZ_NAME):
    s = pd.to_datetime(series_like, errors="coerce", dayfirst=True)
    try:
        if s.dt.tz is None:
            return s.dt.tz_localize(tz_name, ambiguous="infer", nonexistent="shift_forward")
        return s.dt.tz_convert(tz_name)
    except Exception:
        def one(x):
            if pd.isna(x):
                return pd.NaT
            x = pd.Timestamp(x)
            if x.tzinfo is None:
                return x.tz_localize(tz_name, ambiguous="infer", nonexistent="shift_forward")
            return x.tz_convert(tz_name)
        return pd.Series(series_like).apply(one)


def audit_ensure_tz(series_like, tz_name=AUDIT_TZ_NAME):
    raw = pd.Series(series_like).copy()
    s = raw.astype(str).str.strip()
    has_tz = s.str.contains(r"(?:Z|[+-]\d{2}:?\d{2})$", regex=True, na=False)
    out = pd.Series(pd.NaT, index=raw.index, dtype=f"datetime64[ns, {tz_name}]")
    if has_tz.any():
        parsed = pd.to_datetime(raw.loc[has_tz], errors="coerce", utc=True)
        out.loc[has_tz] = parsed.dt.tz_convert(tz_name)
    if (~has_tz).any():
        parsed = pd.to_datetime(raw.loc[~has_tz], errors="coerce", dayfirst=True)
        out.loc[~has_tz] = parsed.dt.tz_localize(tz_name, ambiguous="infer", nonexistent="shift_forward")
    return out


def audit_align_to_zurich(s, tz_name=AUDIT_TZ_NAME):
    s2 = pd.to_datetime(s, errors="coerce")
    try:
        if s2.dt.tz is None:
            return s2.dt.tz_localize(tz_name, ambiguous="infer", nonexistent="shift_forward")
        return s2.dt.tz_convert(tz_name)
    except Exception:
        def fix_one(x):
            if pd.isna(x):
                return pd.NaT
            x = pd.Timestamp(x)
            try:
                if x.tzinfo is None:
                    return x.tz_localize(tz_name, ambiguous="infer", nonexistent="shift_forward")
                return x.tz_convert(tz_name)
            except Exception:
                return x
        return s.apply(fix_one)


def audit_to_local_naive(ts, tz_name=AUDIT_TZ_NAME):
    if pd.isna(ts):
        return pd.NaT
    ts = pd.Timestamp(ts)
    try:
        if ts.tzinfo is not None:
            return ts.tz_convert(tz_name).tz_localize(None)
    except Exception:
        try:
            return ts.tz_localize(None)
        except Exception:
            pass
    return ts


def audit_series_to_local_naive(s, tz_name=AUDIT_TZ_NAME):
    return pd.to_datetime(s.apply(lambda x: audit_to_local_naive(x, tz_name)), errors="coerce")


def audit_merge_blocks(grp_df, gap_min):
    df = grp_df[["start", "end"]].dropna().sort_values("start")
    if df.empty:
        return []
    merged = []
    cur_s, cur_e = df.iloc[0]["start"], df.iloc[0]["end"]
    for _, r in df.iloc[1:].iterrows():
        s, e = r["start"], r["end"]
        gap = (s - cur_e).total_seconds() / 60.0
        if gap <= gap_min:
            cur_e = max(cur_e, e)
        else:
            merged.append((cur_s, cur_e))
            cur_s, cur_e = s, e
    merged.append((cur_s, cur_e))
    return merged


def audit_trip_km_from_cols(df):
    cand_km = ["distance", "Distance", "distance_km", "mileage [km]", "km", "Mileage (km)", "Trip distance [km]"]
    cand_m = ["distance [m]", "Distance [m]", "Trip distance [m]", "distance_m", "mileage [m]", "Distance (m)", "Meters"]
    dcol = rda_pick_col(df, cand_km + cand_m)
    if dcol:
        vals = pd.to_numeric(df[dcol], errors="coerce")
        header = str(dcol).lower()
        looks_meter = ("[m]" in header) or ("(m)" in header) or ("meter" in header) or (dcol in cand_m)
        q95 = np.nanquantile(vals, 0.95) if np.isfinite(vals).any() else np.nan
        if looks_meter or (pd.notna(q95) and q95 > 800):
            return (vals / 1000.0).astype(float)
        return vals.astype(float)
    return pd.Series(np.nan, index=df.index, dtype=float)


def audit_series_speed_kmh(km, start, end):
    dur_h = (pd.to_datetime(end) - pd.to_datetime(start)).dt.total_seconds() / 3600.0
    with np.errstate(divide="ignore", invalid="ignore"):
        spd = km / dur_h
    spd = spd.replace([np.inf, -np.inf], np.nan)
    return spd


def audit_pick_best_sheet(file_bytes, required_groups, prefer_code=None):
    xf = pd.ExcelFile(BytesIO(file_bytes))
    best = None
    best_score = -1
    for name in xf.sheet_names:
        df = xf.parse(name)
        cols = set(df.columns.astype(str))
        score = sum(1 for grp in required_groups if any(c in cols for c in grp))
        if prefer_code is not None:
            try:
                hits = df.astype(str).apply(lambda c: c.str.contains(prefer_code, na=False)).sum().sum()
                if hits > 0:
                    score += 2
            except Exception:
                pass
        if score > best_score:
            best_score = score
            best = (name, df)
    return best


def audit_drop_tz_excel_safe(df, tz_name=AUDIT_TZ_NAME):
    from pandas.api import types as pdt
    out = df.copy()
    for c in out.columns:
        if pdt.is_datetime64tz_dtype(out[c]):
            s = out[c]
            try:
                s = s.dt.tz_convert(tz_name)
            except Exception:
                pass
            out[c] = s.dt.tz_localize(None)
    return out


def audit_fmt_hhmm(ts, tz_name=AUDIT_TZ_NAME):
    ts = audit_to_local_naive(ts, tz_name)
    if pd.isna(ts):
        return "-"
    return pd.Timestamp(ts).strftime("%H:%M")


def audit_fmt_hours(delta):
    if delta is None or pd.isna(delta):
        return "-"
    total_min = int(round(delta.total_seconds() / 60.0))
    if total_min < 0:
        return "-"
    hh = total_min // 60
    mm = total_min % 60
    return f"{hh}h{mm:02d}" if mm else f"{hh}h"


def audit_fmt_span(s, e):
    s = audit_to_local_naive(s)
    e = audit_to_local_naive(e)
    if pd.isna(s) or pd.isna(e) or e <= s:
        return "-"
    return f"{audit_fmt_hhmm(s)} → {audit_fmt_hhmm(e)} ({audit_fmt_hours(e - s)})"


def audit_uniq_flags(series):
    vals = series.fillna("").astype(str)
    s = set()
    for v in vals:
        if not v or v == "nan":
            continue
        for f in v.split(","):
            f = f.strip()
            if f:
                s.add(f)
    return ",".join(sorted(s))


def audit_safe_filename(x):
    s = ("" if pd.isna(x) else str(x)).strip()
    s = re.sub(r"[^\w\-. ]+", "_", s)
    s = re.sub(r"\s+", "_", s)
    return s[:180] if s else "collaborateur"


def audit_duration_mins(a, b):
    if pd.isna(a) or pd.isna(b):
        return np.nan
    return (pd.Timestamp(b) - pd.Timestamp(a)).total_seconds() / 60.0


# ============================================================
# Audit — planning date resolution helpers
# ============================================================

def _audit_as_local_timestamp(value, tz_name=AUDIT_TZ_NAME):
    if pd.isna(value):
        return pd.NaT
    try:
        ts = pd.Timestamp(value)
    except Exception:
        return pd.NaT
    if pd.isna(ts):
        return pd.NaT
    try:
        if ts.tzinfo is not None:
            return ts.tz_convert(tz_name).tz_localize(None)
    except Exception:
        try:
            return ts.tz_localize(None)
        except Exception:
            pass
    return ts


def _audit_valid_planning_date(ts):
    ts = _audit_as_local_timestamp(ts)
    if pd.isna(ts):
        return None
    if AUDIT_PLANNING_DATE_MIN_YEAR <= ts.year <= AUDIT_PLANNING_DATE_MAX_YEAR:
        return ts.date()
    return None


def _audit_excel_or_compact_date(value):
    try:
        v = float(value)
    except Exception:
        return None
    if not np.isfinite(v):
        return None
    if 20000 <= v <= 80000:
        return _audit_valid_planning_date(pd.Timestamp("1899-12-30") + pd.to_timedelta(v, unit="D"))
    if float(v).is_integer():
        text = str(int(v))
        if re.fullmatch(r"20\d{6}", text):
            return _audit_valid_planning_date(pd.to_datetime(text, format="%Y%m%d", errors="coerce"))
    return None


def _audit_try_timestamp_ymd(year, month, day):
    try:
        return pd.Timestamp(year=year, month=month, day=day)
    except ValueError:
        return pd.NaT


def _audit_planning_scalar_date_candidates(value):
    from datetime import date as _dt, datetime as _dtt, time as _tt
    candidates = []

    def add(val, src, pri):
        dt = _audit_valid_planning_date(val)
        if dt is not None:
            candidates.append((dt, src, pri))

    if pd.isna(value):
        return candidates

    if isinstance(value, (pd.Timestamp, _dtt)):
        add(value, "native_datetime", 500)
        return _audit_unique_date_candidates(candidates)

    if isinstance(value, _dt):
        add(pd.Timestamp(value), "native_date", 500)
        return _audit_unique_date_candidates(candidates)

    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        dt = _audit_excel_or_compact_date(value)
        if dt is not None:
            candidates.append((dt, "numeric_date", 480))
        return _audit_unique_date_candidates(candidates)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none", "null"}:
        return candidates

    dt = _audit_excel_or_compact_date(text)
    if dt is not None:
        candidates.append((dt, "numeric_text_date", 470))
        return _audit_unique_date_candidates(candidates)

    if re.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:\s|$)", text):
        add(pd.to_datetime(text, errors="coerce", yearfirst=True), "string_iso", 460)
        return _audit_unique_date_candidates(candidates)

    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?:\s.*)?$", text)
    if m:
        a = int(m.group(1))
        b = int(m.group(2))
        year = int(m.group(3))
        if year < 100:
            year += 2000
        if 1 <= a <= 31 and 1 <= b <= 31:
            if a > 12 and b <= 12:
                add(_audit_try_timestamp_ymd(year, b, a), "string_dayfirst_unambiguous", 430)
            elif b > 12 and a <= 12:
                add(_audit_try_timestamp_ymd(year, a, b), "string_monthfirst_unambiguous", 430)
            else:
                add(_audit_try_timestamp_ymd(year, b, a), "string_dayfirst_ambiguous", 300)
                add(_audit_try_timestamp_ymd(year, a, b), "string_monthfirst_ambiguous", 300)

    add(pd.to_datetime(text, errors="coerce", dayfirst=True), "string_dayfirst_fallback", 120)
    add(pd.to_datetime(text, errors="coerce", dayfirst=False), "string_monthfirst_fallback", 120)
    return _audit_unique_date_candidates(candidates)


def _audit_unique_date_candidates(candidates):
    best = {}
    for dt, source, priority in candidates:
        old = best.get(dt)
        if old is None or priority > old[1]:
            best[dt] = (source, priority)
    return [(dt, source, priority) for dt, (source, priority) in best.items()]


def _audit_planning_date_match_score(cid, dt, ref_by_collab, ref_all):
    if dt is None:
        return 0
    score = 0
    cid_text = None if pd.isna(cid) else str(cid)
    if cid_text and dt in ref_by_collab.get(cid_text, set()):
        score += 1000
    if dt in ref_all:
        score += 100
    return score


def _audit_candidate_for_order(value, order, ref_by_collab, ref_all, date_order_resolved):
    target = f"string_{order}"
    matches = [c for c in _audit_planning_scalar_date_candidates(value) if c[1].startswith(target)]
    if not matches:
        return None
    return sorted(matches, key=lambda x: x[2], reverse=True)[0][0]


def audit_build_reference_dates(rda_df, wf_df):
    ref_by_collab = {}
    ref_all = set()

    def add_ref(cid, value):
        if pd.isna(cid):
            return
        cands = _audit_planning_scalar_date_candidates(value)
        if not cands:
            return
        dt = cands[0][0]
        cid = str(cid)
        ref_by_collab.setdefault(cid, set()).add(dt)
        ref_all.add(dt)

    if not rda_df.empty and "collab_id" in rda_df.columns:
        for rr in rda_df.dropna(subset=["collab_id"]).itertuples(index=False):
            jour = getattr(rr, "jour", pd.NaT)
            start = getattr(rr, "start", pd.NaT)
            add_ref(rr.collab_id, jour if pd.notna(jour) else start)

    if not wf_df.empty and "collab_id" in wf_df.columns:
        for rr in wf_df.dropna(subset=["collab_id"]).itertuples(index=False):
            d = getattr(rr, "date", pd.NaT)
            start = getattr(rr, "start", pd.NaT)
            add_ref(rr.collab_id, d if pd.notna(d) else start)

    return ref_by_collab, ref_all


def audit_resolve_planning_date_order(raw_dates, collab_ids, ref_by_collab, ref_all):
    scores = {"dayfirst": 0, "monthfirst": 0}
    for value, cid in zip(raw_dates, collab_ids):
        if pd.isna(value):
            continue
        text = str(value).strip()
        if not re.match(r"^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}(?:\s.*)?$", text):
            continue
        for order in ["dayfirst", "monthfirst"]:
            dt = _audit_candidate_for_order(value, order, ref_by_collab, ref_all, order)
            scores[order] += _audit_planning_date_match_score(cid, dt, ref_by_collab, ref_all)
    if scores["dayfirst"] > scores["monthfirst"]:
        return "dayfirst", scores
    if scores["monthfirst"] > scores["dayfirst"]:
        return "monthfirst", scores
    return AUDIT_PLANNING_DATE_TIE_BREAKER, scores


def audit_select_planning_date(value, cid, date_order_resolved, ref_by_collab, ref_all):
    candidates = _audit_planning_scalar_date_candidates(value)
    if not candidates:
        return pd.NaT, "unparsed"
    ranked = []
    for dt, source, priority in candidates:
        score = priority + _audit_planning_date_match_score(cid, dt, ref_by_collab, ref_all)
        if source.startswith(f"string_{date_order_resolved}"):
            score += 20
        ranked.append((score, dt, source))
    ranked.sort(key=lambda x: x[0], reverse=True)
    _, dt, source = ranked[0]
    return dt, source


def audit_planning_time_parts(value):
    from datetime import datetime as _dtt, time as _tt
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, _dtt)):
        ts = _audit_as_local_timestamp(value)
        if pd.notna(ts):
            return int(ts.hour), int(ts.minute), int(ts.second)
    if isinstance(value, _tt):
        return int(value.hour), int(value.minute), int(value.second)
    if isinstance(value, (int, float, np.integer, np.floating)) and not isinstance(value, bool):
        try:
            v = float(value)
        except Exception:
            return None
        if not np.isfinite(v):
            return None
        if 0 <= v < 1:
            total_seconds = int(round(v * 86400)) % 86400
        elif 1 <= v < 24:
            total_seconds = int(round(v * 3600)) % 86400
        elif 20000 <= v <= 80000:
            total_seconds = int(round((v % 1) * 86400)) % 86400
        else:
            return None
        return total_seconds // 3600, (total_seconds % 3600) // 60, total_seconds % 60
    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none", "null"}:
        return None
    m = re.search(r"(\d{1,2})[:hH](\d{2})(?::(\d{2}))?", text)
    if m:
        hh, mm, ss = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        if 0 <= hh < 24 and 0 <= mm < 60 and 0 <= ss < 60:
            return hh, mm, ss
    ts = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.notna(ts):
        ts = _audit_as_local_timestamp(ts)
        return int(ts.hour), int(ts.minute), int(ts.second)
    return None


def audit_combine_plan_date_time(parsed_dates, time_ser, tz_name=AUDIT_TZ_NAME):
    from datetime import time as _tt
    values = []
    for d, time_value in zip(parsed_dates, time_ser):
        parts = audit_planning_time_parts(time_value)
        if pd.isna(d) or parts is None:
            values.append(pd.NaT)
            continue
        hh, mm, ss = parts
        values.append(pd.Timestamp.combine(d, _tt(hh, mm, ss)))
    out = pd.to_datetime(pd.Series(values, index=time_ser.index), errors="coerce")
    out = out.dt.tz_localize(tz_name, ambiguous="infer", nonexistent="shift_forward")
    return out


# ============================================================
# Audit — main processing function
# ============================================================

def audit_process(rda_file, wf_file, mapping_file, planning_file, progress_cb=None):
    def _prog(pct, msg):
        if progress_cb:
            progress_cb(pct, msg)

    _prog(0.05, "Chargement des fichiers...")

    RDA_REQUIRED = [
        ["Jour", "Date", "date"],
        ["Début", "Debut", "Heure Début", "Heure Debut", "Start", "start"],
        ["Fin", "Heure fin", "End", "end", "fin"],
        ["Durée", "Duree", "duration", "Minutes"],
        ["No collaborateur", "No Collaborateur", "Employee No", "no_collaborateur"],
    ]
    WF_REQUIRED = [
        ["tripid", "Trip ID", "TripId"],
        ["tripmode", "Trip Mode", "trip_mode"],
        ["start_time", "Start Time", "Start time"],
        ["end_time", "End Time", "End time"],
        ["driverno", "Driver No", "driver_no"],
    ]
    PLANNING_REQUIRED = [
        ["emp_nr"], ["date"], ["start"], ["end"], ["event_color"], ["client_absent"],
    ]

    rda_bytes = rda_file.read()
    wf_bytes = wf_file.read()
    map_bytes = mapping_file.read()
    plan_bytes = planning_file.read()

    rda_name = rda_file.name
    wf_name = wf_file.name

    def _strip_bom(df):
        df.columns = [str(c).lstrip("﻿").strip() for c in df.columns]
        return df

    if rda_name.lower().endswith(".csv"):
        RDA = _strip_bom(read_csv_flex(BytesIO(rda_bytes)))
    else:
        _, RDA = audit_pick_best_sheet(
            rda_bytes, RDA_REQUIRED,
            prefer_code=AUDIT_PRESTATION_61010_CODE if AUDIT_ENABLE_61010_FEATURE else None
        )
        RDA = _strip_bom(RDA)

    if wf_name.lower().endswith(".csv"):
        WF = _strip_bom(read_csv_flex(BytesIO(wf_bytes)))
    else:
        _, WF = audit_pick_best_sheet(wf_bytes, WF_REQUIRED)
        WF = _strip_bom(WF)

    _, PLANNING = audit_pick_best_sheet(plan_bytes, PLANNING_REQUIRED)
    PLANNING = _strip_bom(PLANNING)

    map_sheets = pd.read_excel(BytesIO(map_bytes), sheet_name=None)
    MAP = _strip_bom(map_sheets.get("Matched Collaborateurs", next(iter(map_sheets.values()))))

    _prog(0.12, "Normalisation RDA...")

    rda_cols = {
        "jour": rda_pick_col(RDA, ["Jour", "Date", "date"]),
        "debut": rda_pick_col(RDA, ["Début", "Debut", "Heure Début", "Heure Debut", "Start", "start"]),
        "fin": rda_pick_col(RDA, ["Fin", "Heure fin", "End", "end", "fin"]),
        "duree": rda_pick_col(RDA, ["Durée", "Duree", "duration", "Minutes"]),
        "collab_name": rda_pick_col(RDA, ["Collaborateur", "collaborateur", "Employee"]),
        "collab_no": rda_pick_col(RDA, ["No collaborateur", "No Collaborateur", "Employee No", "no_collaborateur"]),
        "client_nr": rda_pick_col(RDA, ["N° du client", "No client", "ID client", "Client No", "client_nr", "KD-Nr", "KD_Nr"]),
        "client_name": rda_pick_col(RDA, ["Client"]),
        "prestation_name": rda_pick_col(RDA, ["Prestation", "prestation"]),
    }
    missing = [k for k, v in rda_cols.items() if v is None and k in ["jour", "debut", "fin", "duree", "collab_no"]]
    if missing:
        raise KeyError(f"RDA colonnes manquantes : {missing}. Colonnes trouvées : {list(RDA.columns)}")

    rda = pd.DataFrame({
        "jour": audit_swiss_date(RDA[rda_cols["jour"]]),
        "start": audit_swiss_dt(RDA[rda_cols["debut"]]),
        "end": audit_swiss_dt(RDA[rda_cols["fin"]]),
        "duree_min": pd.to_numeric(RDA[rda_cols["duree"]], errors="coerce"),
        "collab_name": (RDA[rda_cols["collab_name"]].astype(str) if rda_cols["collab_name"] else ""),
        "collab_no_sarl": RDA[rda_cols["collab_no"]].apply(audit_to_int_str),
        "client_nr": (RDA[rda_cols["client_nr"]].apply(audit_to_int_str) if rda_cols["client_nr"] else None),
        "client_name": (RDA[rda_cols["client_name"]].apply(audit_clean_legend_text) if rda_cols["client_name"] else ""),
        "prestation_text": (
            RDA[rda_cols["prestation_name"]].apply(audit_clean_legend_text)
            if rda_cols["prestation_name"] else ""
        ),
    })
    mask = rda["duree_min"].isna() & rda["start"].notna() & rda["end"].notna()
    rda.loc[mask, "duree_min"] = (rda.loc[mask, "end"] - rda.loc[mask, "start"]).dt.total_seconds() / 60.0

    best_col, best_score = None, -1
    for col in RDA.columns:
        ser = RDA[col].apply(audit_extract_code_any)
        hits_61010 = int((ser.astype(str) == AUDIT_PRESTATION_61010_CODE).sum())
        hits_any = int(ser.notna().sum())
        score = hits_61010 * 100000 + hits_any
        if score > best_score:
            best_score = score
            best_col = col
    rda["prestation_code"] = np.nan
    if best_col is not None:
        rda["prestation_code"] = RDA[best_col].apply(audit_extract_code_any)
        if not rda_cols["prestation_name"]:
            rda["prestation_text"] = RDA[best_col].apply(audit_clean_legend_text)
    rda["rda_row_id"] = np.arange(len(rda), dtype=int)

    _prog(0.20, "Normalisation Webfleet...")

    wf_cols = {
        "tripid": rda_pick_col(WF, ["tripid", "Trip ID", "TripId"]),
        "tripmode": rda_pick_col(WF, ["tripmode", "Trip Mode", "trip_mode"]),
        "start": rda_pick_col(WF, ["start_time", "Start Time", "Start time"]),
        "end": rda_pick_col(WF, ["end_time", "End Time", "End time"]),
        "driverno": rda_pick_col(WF, ["driverno", "Driver No", "driver_no"]),
        "drivername": rda_pick_col(WF, ["drivername", "Driver Name", "driver_name"]),
    }
    missing_wf = [k for k, v in wf_cols.items() if v is None and k in ["tripid", "tripmode", "start", "end", "driverno"]]
    if missing_wf:
        raise KeyError(f"Webfleet colonnes manquantes : {missing_wf}. Colonnes trouvées : {list(WF.columns)}")

    wf = pd.DataFrame({
        "tripid": WF[wf_cols["tripid"]].astype(str),
        "tripmode": pd.to_numeric(WF[wf_cols["tripmode"]], errors="coerce").astype("Int64"),
        "start": audit_ensure_tz(WF[wf_cols["start"]]),
        "end": audit_ensure_tz(WF[wf_cols["end"]]),
        "driverno": WF[wf_cols["driverno"]].apply(audit_to_int_str),
        "drivername": (WF[wf_cols["drivername"]].astype(str) if wf_cols["drivername"] else ""),
    })
    wf["km"] = pd.to_numeric(audit_trip_km_from_cols(WF), errors="coerce").round(3)
    wf["duration_min"] = (pd.to_datetime(wf["end"]) - pd.to_datetime(wf["start"])).dt.total_seconds() / 60.0
    wf["speed_kmh"] = audit_series_speed_kmh(wf["km"], wf["start"], wf["end"])
    wf["date"] = wf["start"].dt.date

    _prog(0.28, "Chargement mapping collaborateurs...")

    collab_id_col = rda_pick_col(MAP, ["collaborateur-id", "collab_id", "Collaborateur_ID", "collaborateur_id"])
    if not collab_id_col:
        raise KeyError("Mapping : colonne collab_id introuvable (collaborateur-id/collab_id/Collaborateur_ID).")

    map_no_sarl_col = rda_pick_col(MAP, ["no-collaborateur-sarl-102", "collab_no_sarl", "No collaborateur", "No Collaborateur"])
    map_name_sarl_col = rda_pick_col(MAP, ["name-collaborateur", "collaborateur-sarl-102", "collab_sarl", "Collaborateur"])
    map_name_wf_col = rda_pick_col(MAP, ["collaborateur-webfleet", "collab_webfleet", "Driver Name", "name-collaborateur"])
    map_drv_main_col = rda_pick_col(MAP, ["no-collaborateur-wf", "no-collaborateur-webfleet", "driverno", "Driver No"])

    map_df = pd.DataFrame({
        "collab_id": MAP[collab_id_col].astype(str),
        "collab_no_sarl": (MAP[map_no_sarl_col].apply(audit_to_int_str) if map_no_sarl_col else None),
        "collab_name_sarl": (MAP[map_name_sarl_col].astype(str) if map_name_sarl_col else ""),
        "collab_name_wf": (MAP[map_name_wf_col].astype(str) if map_name_wf_col else ""),
        "driverno": (MAP[map_drv_main_col].apply(audit_to_int_str) if map_drv_main_col else None),
    }).dropna(subset=["collab_id"]).drop_duplicates()

    RDA_ID_CANDIDATES = [
        "no-collaborateur-sarl-102", "no-collaborateur-sa-101", "no-collaborateur-ne-103",
        "No collaborateur", "No Collaborateur", "collab_no_sarl",
    ]
    sarlno_to_id = {}
    for col in RDA_ID_CANDIDATES:
        if col in MAP.columns:
            tmp = MAP[[collab_id_col, col]].dropna()
            for _, rr in tmp.iterrows():
                rno = audit_to_int_str(rr[col])
                cid = rr[collab_id_col]
                if rno and pd.notna(cid):
                    sarlno_to_id.setdefault(str(rno), str(cid))
    rda["collab_id"] = rda["collab_no_sarl"].map(sarlno_to_id)

    UO_ID_CANDIDATES = [
        "no-collaborateur-wf", "no-collaborateur-webfleet", "no-collaborateur-sa-101",
        "no-collaborateur-sarl-102", "no-collaborateur-ne-103",
        "collab_no_webfleet", "UO ID", "UO ID 2", "UO ID 3", "driverno", "Driver No",
    ]
    driverno_to_id = {}
    for col in UO_ID_CANDIDATES:
        if col in MAP.columns:
            tmp = MAP[[collab_id_col, col]].dropna()
            for _, rr in tmp.iterrows():
                dno = audit_to_int_str(rr[col])
                cid = rr[collab_id_col]
                if dno and pd.notna(cid):
                    driverno_to_id.setdefault(str(dno), str(cid))
    wf["collab_id"] = wf["driverno"].map(driverno_to_id)

    _prog(0.35, "Normalisation planning...")

    plan_cols = {
        "emp_nr": rda_pick_col(PLANNING, ["emp_nr"]),
        "date": rda_pick_col(PLANNING, ["date"]),
        "start": rda_pick_col(PLANNING, ["start"]),
        "end": rda_pick_col(PLANNING, ["end"]),
        "duration": rda_pick_col(PLANNING, ["duration"]),
        "event_color": rda_pick_col(PLANNING, ["event_color"]),
        "client_absent": rda_pick_col(PLANNING, ["client_absent"]),
        "type": rda_pick_col(PLANNING, ["type"]),
        "note": rda_pick_col(PLANNING, ["note"]),
        "client_nr": rda_pick_col(PLANNING, ["client_nr"]),
        "client_firstname": rda_pick_col(PLANNING, ["client_firstname"]),
        "client_lastname": rda_pick_col(PLANNING, ["client_lastname"]),
    }
    missing_plan = [k for k, v in plan_cols.items() if v is None and k in ["emp_nr", "date", "start", "end", "event_color"]]
    if missing_plan:
        raise KeyError(f"Planning colonnes manquantes : {missing_plan}. Colonnes trouvées : {list(PLANNING.columns)}")

    planning_emp_nr = PLANNING[plan_cols["emp_nr"]].apply(audit_to_int_str)

    PLAN_EMP_ID_CANDIDATES = list(dict.fromkeys(RDA_ID_CANDIDATES + UO_ID_CANDIDATES + [collab_id_col, "emp_nr", "employee_nr", "Employee No"]))
    plan_emp_to_id = dict(sarlno_to_id)

    cols_list = list(MAP.columns)
    id_positions = [i for i, c in enumerate(cols_list) if c == collab_id_col]
    for col in PLAN_EMP_ID_CANDIDATES:
        if col in MAP.columns and col != collab_id_col:
            value_positions = [i for i, c in enumerate(cols_list) if c == col]
            for id_pos in id_positions:
                for value_pos in value_positions:
                    if id_pos == value_pos:
                        continue
                    tmp = MAP.iloc[:, [id_pos, value_pos]].copy()
                    tmp.columns = ["_cid", "_emp"]
                    for _, rr in tmp.dropna(subset=["_cid", "_emp"]).iterrows():
                        emp = audit_to_int_str(rr["_emp"])
                        cid = rr["_cid"]
                        if emp and pd.notna(cid):
                            plan_emp_to_id.setdefault(str(emp), str(cid))

    for cid_pos in id_positions:
        for cid in MAP.iloc[:, cid_pos].dropna().astype(str):
            plan_emp_to_id.setdefault(cid, cid)
            cid_norm = audit_to_int_str(cid)
            if cid_norm:
                plan_emp_to_id.setdefault(cid_norm, cid)

    planning_collab_id = planning_emp_nr.map(plan_emp_to_id)

    ref_by_collab, ref_all = audit_build_reference_dates(rda, wf)
    date_order_resolved, _ = audit_resolve_planning_date_order(
        PLANNING[plan_cols["date"]], planning_collab_id, ref_by_collab, ref_all
    )

    selected_plan_dates = [
        audit_select_planning_date(v, cid, date_order_resolved, ref_by_collab, ref_all)
        for v, cid in zip(PLANNING[plan_cols["date"]], planning_collab_id)
    ]
    planning_date_values = pd.Series([x[0] for x in selected_plan_dates], index=PLANNING.index)
    planning_date_sources = pd.Series([x[1] for x in selected_plan_dates], index=PLANNING.index)

    planning = pd.DataFrame({
        "emp_nr": planning_emp_nr,
        "collab_id": planning_collab_id,
        "date": planning_date_values,
        "date_parse_source": planning_date_sources,
        "raw_date": PLANNING[plan_cols["date"]],
        "raw_start": PLANNING[plan_cols["start"]],
        "raw_end": PLANNING[plan_cols["end"]],
        "start": audit_combine_plan_date_time(planning_date_values, PLANNING[plan_cols["start"]]),
        "end": audit_combine_plan_date_time(planning_date_values, PLANNING[plan_cols["end"]]),
        "duration_min": (pd.to_numeric(PLANNING[plan_cols["duration"]], errors="coerce") if plan_cols["duration"] else np.nan),
        "event_color": PLANNING[plan_cols["event_color"]].fillna("").astype(str).str.strip(),
        "client_absent": (
            PLANNING[plan_cols["client_absent"]].fillna("N").astype(str).str.strip().str.upper()
            if plan_cols["client_absent"] else "N"
        ),
        "type": (pd.to_numeric(PLANNING[plan_cols["type"]], errors="coerce") if plan_cols["type"] else np.nan),
        "note": (PLANNING[plan_cols["note"]].fillna("").astype(str) if plan_cols["note"] else ""),
        "client_nr": (PLANNING[plan_cols["client_nr"]].apply(audit_to_int_str) if plan_cols["client_nr"] else None),
    })
    planning_firstname = (
        PLANNING[plan_cols["client_firstname"]].apply(audit_clean_legend_text)
        if plan_cols["client_firstname"] else pd.Series("", index=PLANNING.index)
    )
    planning_lastname = (
        PLANNING[plan_cols["client_lastname"]].apply(audit_clean_legend_text)
        if plan_cols["client_lastname"] else pd.Series("", index=PLANNING.index)
    )
    planning["client_name"] = (planning_lastname + " " + planning_firstname).str.strip()
    planning["client_label"] = planning["client_name"].fillna("").astype(str).str.strip()
    if "client_nr" in planning.columns:
        nr_label = planning["client_nr"].fillna("").astype(str).str.strip()
        planning["client_label"] = np.where(
            planning["client_label"].astype(str).str.len() > 0,
            planning["client_label"],
            np.where(nr_label.str.len() > 0, "Client " + nr_label, ""),
        )

    overnight = planning["start"].notna() & planning["end"].notna() & (planning["end"] < planning["start"])
    planning.loc[overnight, "end"] = planning.loc[overnight, "end"] + pd.Timedelta(days=1)
    dur_missing = planning["duration_min"].isna() & planning["start"].notna() & planning["end"].notna()
    planning.loc[dur_missing, "duration_min"] = (planning.loc[dur_missing, "end"] - planning.loc[dur_missing, "start"]).dt.total_seconds() / 60.0

    planning["date_only"] = planning["date"]
    planning["event_color_key"] = planning["event_color"].fillna("").astype(str).str.strip().str.lower()

    drop_reasons = []
    for rr in planning.itertuples(index=False):
        reasons = []
        if pd.isna(getattr(rr, "collab_id", pd.NA)):
            reasons.append("unmapped_emp_nr")
        if pd.isna(getattr(rr, "date", pd.NaT)):
            reasons.append("bad_date")
        if pd.isna(getattr(rr, "start", pd.NaT)):
            reasons.append("bad_start_time")
        if pd.isna(getattr(rr, "end", pd.NaT)):
            reasons.append("bad_end_time")
        if (pd.notna(getattr(rr, "start", pd.NaT)) and pd.notna(getattr(rr, "end", pd.NaT))
                and getattr(rr, "end") <= getattr(rr, "start")):
            reasons.append("non_positive_span")
        drop_reasons.append(",".join(reasons))
    planning["_drop"] = drop_reasons
    planning = planning[planning["_drop"].str.len() == 0].copy()
    planning.drop(columns=["_drop"], inplace=True)
    planning["collab_id"] = planning["collab_id"].astype(str)
    planning["plot_color"] = np.where(
        planning["client_absent"].eq("Y"),
        "#000000",
        planning["event_color_key"].map(AUDIT_PLAN_COLOR_MAP).fillna("#bdbdbd"),
    )

    _prog(0.45, "Alignement timezone et construction rda_daily...")

    for df_obj in [rda, wf]:
        for col in ["start", "end"]:
            if col in df_obj.columns:
                df_obj[col] = audit_align_to_zurich(df_obj[col])
    wf["date"] = wf["start"].dt.date
    wf["duration_min"] = (pd.to_datetime(wf["end"]) - pd.to_datetime(wf["start"])).dt.total_seconds() / 60.0
    wf["speed_kmh"] = audit_series_speed_kmh(wf["km"], wf["start"], wf["end"])

    rows_daily = []
    for (cid, day), grp in rda.dropna(subset=["collab_id", "jour", "start", "end"]).groupby(["collab_id", "jour"]):
        grp = grp.copy().sort_values(["start", "end"])
        starts = grp["start"].dropna()
        ends = grp["end"].dropna()
        if starts.empty or ends.empty:
            continue
        first_start = starts.min()
        last_end = ends.max()
        total_min = pd.to_numeric(grp["duree_min"], errors="coerce")
        fallback = (grp["end"] - grp["start"]).dt.total_seconds() / 60.0
        total_min = float(total_min.fillna(fallback).sum())
        blocks_tight = audit_merge_blocks(grp, gap_min=AUDIT_GAP_MERGE_MIN)
        block_cnt = len(blocks_tight)
        duty_span_min = float((pd.to_datetime(last_end) - pd.to_datetime(first_start)).total_seconds() / 60.0)
        time_full = total_min >= AUDIT_FULL_DAY_MINUTES
        span_full = block_cnt >= AUDIT_MIN_BLOCKS_FOR_SPAN and duty_span_min >= AUDIT_FULL_SPAN_MINUTES
        day_class = "Full" if (time_full or span_full) else "Half"
        main_blocks = audit_merge_blocks(grp, gap_min=AUDIT_INTERNAL_BLOCK_GAP_MIN)
        ibs = pd.NaT
        ibe = pd.NaT
        if len(main_blocks) >= 2:
            raw_ibs = main_blocks[0][1]
            raw_ibe = main_blocks[1][0]
            ibs = raw_ibs + pd.Timedelta(minutes=AUDIT_INTERNAL_BUFFER_MIN)
            ibe = raw_ibe - pd.Timedelta(minutes=AUDIT_INTERNAL_BUFFER_MIN)
            if pd.notna(ibs) and pd.notna(ibe) and ibe <= ibs:
                ibs = pd.NaT
                ibe = pd.NaT
        rows_daily.append({
            "collab_id": str(cid),
            "date": day,
            "rda_first_start": first_start,
            "rda_last_end": last_end,
            "rda_total_min": total_min,
            "rda_block_count": int(block_cnt),
            "duty_span_min": duty_span_min,
            "day_class": day_class,
            "buffer_start": first_start - pd.Timedelta(minutes=AUDIT_PRE_SHIFT_BUFFER_MIN) if AUDIT_CHECK_PRE_SHIFT else pd.NaT,
            "buffer_end": last_end + pd.Timedelta(minutes=AUDIT_WORK_END_BUFFER_MIN),
            "internal_buf_start": ibs,
            "internal_buf_end": ibe,
        })
    rda_daily = pd.DataFrame(rows_daily)
    for col in ["rda_first_start", "rda_last_end", "buffer_start", "buffer_end", "internal_buf_start", "internal_buf_end"]:
        if col in rda_daily.columns:
            rda_daily[col] = audit_align_to_zurich(rda_daily[col])

    _prog(0.55, "Annotation des trajets Webfleet...")

    wf = wf.copy()
    wf["collab_id"] = wf["collab_id"].astype(str)
    merge_cols = ["rda_first_start", "rda_last_end", "buffer_start", "buffer_end", "internal_buf_start", "internal_buf_end", "day_class"]
    wf = wf.drop(columns=[c for c in merge_cols if c in wf.columns], errors="ignore")
    if not rda_daily.empty:
        wf = wf.merge(rda_daily[["collab_id", "date"] + merge_cols], how="left", on=["collab_id", "date"])
    else:
        for c in merge_cols:
            wf[c] = np.nan

    wf["offday"] = wf["rda_first_start"].isna()
    wf["within_service"] = (~wf["offday"]) & (wf["start"] <= wf["rda_last_end"]) & (wf["end"] >= wf["rda_first_start"])
    wf["after_buffer"] = (~wf["buffer_end"].isna()) & (wf["start"] >= wf["buffer_end"])
    if AUDIT_CHECK_PRE_SHIFT:
        wf["before_shift"] = (~wf["buffer_start"].isna()) & (wf["end"] <= wf["buffer_start"])
    else:
        wf["before_shift"] = False

    def _interval_overlap(a_s, a_e, b_s, b_e):
        if pd.isna(a_s) or pd.isna(a_e) or pd.isna(b_s) or pd.isna(b_e):
            return False
        ov_s = max(a_s, b_s)
        ov_e = min(a_e, b_e)
        ov_min = (ov_e - ov_s).total_seconds() / 60.0
        if ov_min <= 0:
            return False
        if AUDIT_INTERNAL_BUF_MIN_OVERLAP_MIN > 0 and ov_min < AUDIT_INTERNAL_BUF_MIN_OVERLAP_MIN:
            return False
        if AUDIT_INTERNAL_BUF_MIN_OVERLAP_RATIO > 0:
            trip_min = (a_e - a_s).total_seconds() / 60.0
            if trip_min > 0 and (ov_min / trip_min) < AUDIT_INTERNAL_BUF_MIN_OVERLAP_RATIO:
                return False
        return True

    wf["in_internal_buffer"] = [
        _interval_overlap(s, e, ibs, ibe)
        for s, e, ibs, ibe in zip(wf["start"], wf["end"], wf["internal_buf_start"], wf["internal_buf_end"])
    ]

    def _build_flags(row):
        km = float(row["km"]) if pd.notna(row["km"]) else 0.0
        f = []
        if row["within_service"] and row["tripmode"] == 1 and km >= AUDIT_KM_MIN_FOR_FLAG_CONTEXT:
            f.append("MODE1_DURING_SERVICE")
        if row["after_buffer"] and row["tripmode"] in [2, 3] and km >= AUDIT_KM_MIN_FOR_FLAG_CONTEXT:
            f.append("MODE23_AFTER_BUFFER")
        if row["offday"] and row["tripmode"] in [2, 3] and km >= AUDIT_KM_MIN_FOR_FLAG_CONTEXT:
            f.append("MODE23_ON_OFFDAY")
            f.append("NO_RDA_BUT_MODE23")
        if row["before_shift"] and row["tripmode"] in [2, 3] and km >= AUDIT_KM_MIN_FOR_FLAG_CONTEXT:
            f.append("PRE_SHIFT_MODE2")
        if row["in_internal_buffer"] and row["tripmode"] in [2, 3] and km >= AUDIT_KM_MIN_FOR_FLAG_CONTEXT:
            f.append("MODE23_IN_INTERNAL_BUFFER")
        if pd.isna(row["collab_id"]) or row["collab_id"] in ("None", "nan"):
            f.append("UNMAPPED_DRIVER")
        if pd.notna(row.get("speed_kmh", np.nan)) and row["speed_kmh"] > AUDIT_MAX_REASONABLE_SPEED_KMH:
            f.append("SPEED_EXCEEDS_MAX")
        has_rda_today = not bool(row["offday"])
        outside_service = (not bool(row["within_service"])) and (bool(row["after_buffer"]) or bool(row["before_shift"]))
        if has_rda_today and outside_service and row["tripmode"] in [2, 3] and km >= AUDIT_KM_MIN_FOR_FLAG_CONTEXT:
            f.append("MODE23_OUTSIDE_SERVICE_ON_RDA_DAY")
        return ",".join(sorted(set(f)))

    wf["flags"] = wf.apply(_build_flags, axis=1)
    wf["suspect_private_km"] = np.where(
        wf["tripmode"].isin([2, 3]) & (wf["after_buffer"] | wf["offday"] | wf["before_shift"] | wf["in_internal_buffer"]),
        wf["km"].fillna(0.0),
        0.0,
    )

    _prog(0.63, "Labels d'usage journaliers...")

    rda_dates = rda[["collab_id", "jour"]].dropna().copy()
    rda_dates["collab_id"] = rda_dates["collab_id"].astype(str)
    rda_dates["date"] = rda_dates["jour"]
    rda_dates = rda_dates[["collab_id", "date"]]
    wf_dates = wf[["collab_id"]].copy()
    wf_dates["collab_id"] = wf_dates["collab_id"].astype(str)
    wf_dates["date"] = wf["start"].dt.date
    all_days = pd.concat([rda_dates, wf_dates], ignore_index=True).dropna(subset=["collab_id", "date"]).drop_duplicates()
    all_days["date"] = pd.to_datetime(all_days["date"]).dt.date

    wf_day_agg = wf.groupby(["collab_id", "date"]).agg(
        wf_trip_count=("tripid", "count"),
        wf_km_total=("km", "sum"),
        wf_any_within_service=("within_service", lambda s: bool(pd.Series(s).fillna(False).any())),
        wf_any_after_buffer=("after_buffer", lambda s: bool(pd.Series(s).fillna(False).any())),
        wf_any_before_shift=("before_shift", lambda s: bool(pd.Series(s).fillna(False).any())),
    ).reset_index()

    if not rda_daily.empty:
        ad = (all_days
              .merge(rda_daily[["collab_id", "date", "rda_first_start", "rda_last_end", "day_class"]], how="left", on=["collab_id", "date"])
              .merge(wf_day_agg, how="left", on=["collab_id", "date"]))
    else:
        ad = all_days.merge(wf_day_agg, how="left", on=["collab_id", "date"])
        ad["rda_first_start"] = pd.NaT
        ad["rda_last_end"] = pd.NaT
        ad["day_class"] = None

    ad["wf_trip_count"] = ad["wf_trip_count"].fillna(0).astype(int)
    ad["wf_km_total"] = ad["wf_km_total"].fillna(0.0).astype(float)
    for c in ["wf_any_within_service", "wf_any_after_buffer", "wf_any_before_shift"]:
        ad[c] = ad[c].fillna(False).astype(bool)
    ad["has_rda"] = ad["rda_first_start"].notna()
    ad["has_wf"] = ad["wf_trip_count"] > 0
    ad["car_used_within_service"] = ad["has_rda"] & ad["wf_any_within_service"]
    ad["car_used_outside_service"] = ad["has_rda"] & (ad["wf_any_after_buffer"] | ad["wf_any_before_shift"])

    def _usage_label(row):
        if row["has_rda"]:
            if not row["has_wf"]:
                return "WORK_NO_CAR"
            if row["car_used_within_service"]:
                return "WORK_CAR_IN_SERVICE"
            return "WORK_CAR_ONLY_OUTSIDE"
        return "NO_WORK_CAR_USED" if row["has_wf"] else "OFF_NO_CAR"

    ad["usage_label"] = ad.apply(_usage_label, axis=1)
    daily_usage = ad[["collab_id", "date", "has_rda", "has_wf", "car_used_within_service", "car_used_outside_service", "usage_label", "day_class"]].copy()

    _prog(0.70, "Vérifications 61010...")

    prestation61010_checks = pd.DataFrame()
    prestation61010_day_counts = pd.DataFrame()

    if AUDIT_ENABLE_61010_FEATURE:
        r61010 = rda[
            (rda["prestation_code"].astype(str) == str(AUDIT_PRESTATION_61010_CODE)) &
            rda["collab_id"].notna() & rda["start"].notna() & rda["end"].notna()
        ].copy()
        r61010["collab_id"] = r61010["collab_id"].astype(str)
        r61010["start"] = audit_series_to_local_naive(r61010["start"])
        r61010["end"] = audit_series_to_local_naive(r61010["end"])
        r61010 = r61010[r61010["start"].notna() & r61010["end"].notna() & (r61010["end"] > r61010["start"])].copy()

        wf_ok = wf[wf["collab_id"].notna() & wf["start"].notna() & wf["end"].notna()].copy()
        wf_ok["collab_id"] = wf_ok["collab_id"].astype(str)
        wf_ok["start"] = audit_series_to_local_naive(wf_ok["start"])
        wf_ok["end"] = audit_series_to_local_naive(wf_ok["end"])
        wf_ok = wf_ok[wf_ok["start"].notna() & wf_ok["end"].notna() & (wf_ok["end"] > wf_ok["start"])].copy()
        wf_groups = {cid: g.sort_values("start").copy() for cid, g in wf_ok.groupby("collab_id")}

        def _gap_min(tr_s, tr_e, win_s, win_e):
            if pd.isna(tr_s) or pd.isna(tr_e) or pd.isna(win_s) or pd.isna(win_e):
                return np.nan
            if tr_e < win_s:
                return (win_s - tr_e).total_seconds() / 60.0
            if tr_s > win_e:
                return (tr_s - win_e).total_seconds() / 60.0
            return 0.0

        def _overlap_min(a_s, a_e, b_s, b_e):
            if pd.isna(a_s) or pd.isna(a_e) or pd.isna(b_s) or pd.isna(b_e):
                return np.nan
            ov_s = max(a_s, b_s)
            ov_e = min(a_e, b_e)
            return max(0.0, (ov_e - ov_s).total_seconds() / 60.0)

        rows_61010 = []
        for rr in r61010.itertuples(index=False):
            cid = str(rr.collab_id)
            st, en = rr.start, rr.end
            buf_s = st - pd.Timedelta(minutes=AUDIT_PRESTATION_61010_BUFFER_MIN)
            buf_e = en + pd.Timedelta(minutes=AUDIT_PRESTATION_61010_BUFFER_MIN)
            trips = wf_groups.get(cid, pd.DataFrame())
            has_inside = False
            nearest_gap = np.nan
            nearest_tripid = None
            nearest_tripmode = None
            nearest_trip_start = pd.NaT
            nearest_trip_end = pd.NaT
            nearest_overlap_buf_min = np.nan
            nearest_overlap_61010_min = np.nan
            start_delta = np.nan
            end_delta = np.nan
            rda_dur = max(0.0, (en - st).total_seconds() / 60.0)
            uncovered = np.nan
            total_dev = np.nan
            if not trips.empty:
                cand = trips[(trips["start"] <= buf_e + pd.Timedelta(hours=12)) & (trips["end"] >= buf_s - pd.Timedelta(hours=12))].copy()
                if not cand.empty:
                    inside = (cand["start"] >= buf_s) & (cand["end"] <= buf_e)
                    nearest_row = None
                    if inside.any():
                        has_inside = True
                        nearest_gap = 0.0
                        nearest_row = cand.loc[inside].sort_values("start").iloc[0]
                    else:
                        cand["gap_min"] = [_gap_min(ts, te, buf_s, buf_e) for ts, te in zip(cand["start"], cand["end"])]
                        cand = cand.dropna(subset=["gap_min"]).sort_values(["gap_min", "start"])
                        if not cand.empty:
                            nearest_row = cand.iloc[0]
                            nearest_gap = float(nearest_row.get("gap_min", np.nan))
                    if nearest_row is not None:
                        nearest_tripid = str(nearest_row.get("tripid", ""))
                        nearest_tripmode = nearest_row.get("tripmode", None)
                        nearest_trip_start = nearest_row.get("start", pd.NaT)
                        nearest_trip_end = nearest_row.get("end", pd.NaT)
                        nearest_overlap_buf_min = _overlap_min(nearest_trip_start, nearest_trip_end, buf_s, buf_e)
                        nearest_overlap_61010_min = _overlap_min(nearest_trip_start, nearest_trip_end, st, en)
                        if pd.notna(nearest_trip_start):
                            start_delta = (st - nearest_trip_start).total_seconds() / 60.0
                        if pd.notna(nearest_trip_end):
                            end_delta = (en - nearest_trip_end).total_seconds() / 60.0
                        ov_for_calc = nearest_overlap_61010_min if pd.notna(nearest_overlap_61010_min) else 0.0
                        uncovered = max(0.0, rda_dur - ov_for_calc)
                        if pd.notna(start_delta) and pd.notna(end_delta):
                            total_dev = abs(start_delta) + abs(end_delta)
            if pd.isna(uncovered):
                uncovered = rda_dur
            if pd.isna(total_dev):
                total_dev = rda_dur
            rows_61010.append({
                "rda_row_id": int(rr.rda_row_id), "collab_id": cid,
                "date": pd.to_datetime(st).date(), "start": st, "end": en,
                "buf_start": buf_s, "buf_end": buf_e,
                "has_wf_trip_in_buffer": bool(has_inside), "nearest_gap_min": nearest_gap,
                "nearest_tripid": nearest_tripid, "nearest_tripmode": nearest_tripmode,
                "nearest_trip_start": nearest_trip_start, "nearest_trip_end": nearest_trip_end,
                "rda_61010_duration_min": rda_dur,
                "nearest_overlap_buf_min": nearest_overlap_buf_min,
                "nearest_overlap_61010_min": nearest_overlap_61010_min,
                "start_delta_rda_vs_wf_min": start_delta, "end_delta_rda_vs_wf_min": end_delta,
                "uncovered_61010_min": uncovered, "total_deviation_min": total_dev,
            })

        prestation61010_checks = pd.DataFrame(rows_61010)
        if not prestation61010_checks.empty:
            prestation61010_day_counts = (
                prestation61010_checks.groupby(["collab_id", "date"]).agg(
                    prestation_61010_count=("rda_row_id", "count"),
                    prestation_61010_missing=("has_wf_trip_in_buffer", lambda s: int((~pd.Series(s).astype(bool)).sum())),
                ).reset_index()
            )
        merge_back_cols = ["has_wf_trip_in_buffer", "buf_start", "buf_end", "rda_61010_duration_min",
                           "nearest_overlap_buf_min", "nearest_overlap_61010_min",
                           "start_delta_rda_vs_wf_min", "end_delta_rda_vs_wf_min",
                           "uncovered_61010_min", "total_deviation_min"]
        rda = rda.drop(columns=[c for c in merge_back_cols if c in rda.columns], errors="ignore")
        if not prestation61010_checks.empty:
            rda = rda.merge(prestation61010_checks[["rda_row_id"] + merge_back_cols], how="left", on="rda_row_id")
        else:
            for c in merge_back_cols:
                rda[c] = np.nan

    _prog(0.80, "Agrégats et statistiques collaborateurs...")

    wf2 = wf.copy()
    wf2["collab_id"] = wf2["collab_id"].astype(str)
    wf2["km"] = pd.to_numeric(wf2["km"], errors="coerce").fillna(0.0)
    wf2["km_m1"] = np.where(wf2["tripmode"] == 1, wf2["km"], 0.0)
    wf2["km_m2"] = np.where(wf2["tripmode"] == 2, wf2["km"], 0.0)
    wf2["km_m3"] = np.where(wf2["tripmode"] == 3, wf2["km"], 0.0)
    wf2["trips_m1"] = (wf2["tripmode"] == 1).astype(int)
    wf2["trips_m2"] = (wf2["tripmode"] == 2).astype(int)
    wf2["trips_m3"] = (wf2["tripmode"] == 3).astype(int)

    agg_daily = wf2.groupby(["collab_id", "date"], dropna=False).agg(
        wf_trip_count=("tripid", "count"),
        trips_m1=("trips_m1", "sum"),
        trips_m2=("trips_m2", "sum"),
        trips_m3=("trips_m3", "sum"),
        km_mode1=("km_m1", "sum"),
        km_mode2=("km_m2", "sum"),
        km_mode3=("km_m3", "sum"),
        suspect_private_km=("suspect_private_km", "sum"),
        flags_concat=("flags", audit_uniq_flags),
    ).reset_index()
    if not rda_daily.empty:
        agg_daily = agg_daily.merge(rda_daily, how="left", on=["collab_id", "date"])
    agg_daily["private_km_total_for_day"] = (agg_daily.get("km_mode1", 0).fillna(0) + agg_daily["suspect_private_km"].fillna(0)).round(3)
    agg_daily = agg_daily.merge(daily_usage[["collab_id", "date", "usage_label", "has_rda", "has_wf", "car_used_within_service", "car_used_outside_service"]], how="left", on=["collab_id", "date"])

    name_cols = map_df[["collab_id", "collab_name_sarl", "collab_no_sarl", "collab_name_wf", "driverno"]].drop_duplicates()
    name_cols["collab_id"] = name_cols["collab_id"].astype(str)
    agg_daily = agg_daily.merge(name_cols, how="left", on="collab_id")

    collab_stats = wf2.groupby("collab_id", dropna=False).agg(
        trips_total=("tripid", "count"),
        trips_m1=("trips_m1", "sum"),
        trips_m2=("trips_m2", "sum"),
        trips_m3=("trips_m3", "sum"),
        km_total=("km", "sum"),
        km_m1=("km_m1", "sum"),
        km_m2=("km_m2", "sum"),
        km_m3=("km_m3", "sum"),
        suspect_private_km=("suspect_private_km", "sum"),
        flags_concat=("flags", audit_uniq_flags),
    ).reset_index().merge(name_cols, how="left", on="collab_id")

    if not rda_daily.empty:
        day_counts = rda_daily.pivot_table(index="collab_id", columns="day_class", values="date", aggfunc="count").fillna(0).astype(int)
        day_counts = day_counts.rename(columns={"Full": "days_full", "Half": "days_half"})
        for c in ["days_full", "days_half"]:
            if c not in day_counts.columns:
                day_counts[c] = 0
        collab_stats = collab_stats.merge(day_counts.reset_index(), how="left", on="collab_id")
    else:
        collab_stats["days_full"] = 0
        collab_stats["days_half"] = 0

    usage_counts_pivot = agg_daily.groupby("collab_id")["usage_label"].value_counts().unstack(fill_value=0)
    for col in ["WORK_NO_CAR", "NO_WORK_CAR_USED", "WORK_CAR_IN_SERVICE", "WORK_CAR_ONLY_OUTSIDE", "OFF_NO_CAR"]:
        if col not in usage_counts_pivot.columns:
            usage_counts_pivot[col] = 0
    collab_stats = collab_stats.merge(usage_counts_pivot.reset_index(), how="left", on="collab_id")
    for col in ["WORK_NO_CAR", "NO_WORK_CAR_USED", "WORK_CAR_IN_SERVICE", "WORK_CAR_ONLY_OUTSIDE", "OFF_NO_CAR"]:
        collab_stats[col] = collab_stats[col].fillna(0).astype(int)

    base = agg_daily.copy()
    base["car_used_within_service"] = base.get("car_used_within_service", pd.Series(False)).fillna(False).astype(bool)
    full_wcar = base[(base.get("day_class", "") == "Full") & base["car_used_within_service"]].groupby("collab_id")["date"].nunique()
    half_wcar = base[(base.get("day_class", "") == "Half") & base["car_used_within_service"]].groupby("collab_id")["date"].nunique()
    collab_stats["full_days_with_car"] = collab_stats["collab_id"].map(full_wcar).fillna(0).astype(int)
    collab_stats["half_days_with_car"] = collab_stats["collab_id"].map(half_wcar).fillna(0).astype(int)
    collab_stats["work_days_total"] = (collab_stats["days_full"].fillna(0).astype(int) + collab_stats["days_half"].fillna(0).astype(int))

    if AUDIT_ENABLE_61010_FEATURE and not prestation61010_checks.empty:
        tot = prestation61010_checks.groupby("collab_id").agg(
            prestation_61010_total=("rda_row_id", "count"),
            prestation_61010_missing_total=("has_wf_trip_in_buffer", lambda s: int((~pd.Series(s).astype(bool)).sum())),
        ).reset_index()
        collab_stats = collab_stats.merge(tot, how="left", on="collab_id")
    else:
        collab_stats["prestation_61010_total"] = 0
        collab_stats["prestation_61010_missing_total"] = 0
    collab_stats["prestation_61010_total"] = collab_stats["prestation_61010_total"].fillna(0).astype(int)
    collab_stats["prestation_61010_missing_total"] = collab_stats["prestation_61010_missing_total"].fillna(0).astype(int)

    flags_exp = wf[["collab_id", "flags", "km"]].copy()
    flags_exp["flag"] = flags_exp["flags"].fillna("").astype(str).str.split(",")
    flags_exp = flags_exp.explode("flag")
    flags_exp = flags_exp[flags_exp["flag"].astype(str).str.strip() != ""]
    flag_summary = (
        flags_exp.groupby("flag").agg(count=("collab_id", "count"), km_sum=("km", "sum"))
        .reset_index().sort_values(["count", "km_sum"], ascending=[False, False])
    )

    _prog(0.90, "Export Excel...")

    output_dir = get_session_output_root(AUDIT_OUTPUT_FOLDER)
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = output_dir / f"audit_{timestamp}.xlsx"

    excel_bytes = BytesIO()
    with pd.ExcelWriter(excel_bytes, engine="openpyxl") as xw:
        audit_drop_tz_excel_safe(collab_stats).to_excel(xw, index=False, sheet_name="Collaborator_Summary")
        audit_drop_tz_excel_safe(agg_daily).to_excel(xw, index=False, sheet_name="Daily_Aggregates")
        audit_drop_tz_excel_safe(wf2).to_excel(xw, index=False, sheet_name="All_Trips")
        audit_drop_tz_excel_safe(rda).to_excel(xw, index=False, sheet_name="All_RDA_Entries")
        if not rda_daily.empty:
            audit_drop_tz_excel_safe(rda_daily).to_excel(xw, index=False, sheet_name="RDA_Daily")
        map_df.to_excel(xw, index=False, sheet_name="Mapping")
        flag_summary.to_excel(xw, index=False, sheet_name="Flag_Summary")
        audit_drop_tz_excel_safe(planning).to_excel(xw, index=False, sheet_name="Planning")
        if AUDIT_ENABLE_61010_FEATURE and not prestation61010_checks.empty:
            audit_drop_tz_excel_safe(prestation61010_checks).to_excel(xw, index=False, sheet_name="Prestation_61010_Checks")
        if AUDIT_ENABLE_61010_FEATURE and not prestation61010_day_counts.empty:
            prestation61010_day_counts.to_excel(xw, index=False, sheet_name="Prestation_61010_DayCounts")
    excel_bytes.seek(0)
    with open(excel_path, "wb") as f:
        f.write(excel_bytes.getvalue())
    excel_bytes.seek(0)

    total_trips = len(wf)
    flagged = int((wf["flags"].fillna("").str.strip() != "").sum())
    flagged_pct = round(100.0 * flagged / total_trips, 1) if total_trips > 0 else 0.0
    suspect_km = float(wf["suspect_private_km"].fillna(0).sum())

    _prog(1.0, "Audit terminé")

    return {
        "rda": rda,
        "wf": wf,
        "wf2": wf2,
        "map_df": map_df,
        "planning": planning,
        "rda_daily": rda_daily,
        "agg_daily": agg_daily,
        "collab_stats": collab_stats,
        "daily_usage": daily_usage,
        "flag_summary": flag_summary,
        "prestation61010_checks": prestation61010_checks,
        "prestation61010_day_counts": prestation61010_day_counts,
        "excel_path": excel_path,
        "excel_bytes": excel_bytes,
        "metrics": {
            "total_trips": total_trips,
            "total_collabs": wf["collab_id"].nunique(),
            "flagged_trip_pct": flagged_pct,
            "suspect_km_total": suspect_km,
        },
    }


# ============================================================
# Audit — chart data builder (no matplotlib)
# ============================================================

def audit_build_chart_data(result: dict):
    """Build events DataFrame and lookup dicts needed for Gantt charts. No matplotlib dependency."""
    tz_name = AUDIT_TZ_NAME
    wf = result["wf"]
    rda = result["rda"]
    planning = result["planning"]
    map_df = result["map_df"]

    def _to_ln(ts):
        return audit_to_local_naive(ts, tz_name)

    def _safe_text(x):
        return "" if pd.isna(x) else str(x)

    def _cmp_date_str(value, fallback_ts=None):
        if value is not None and not pd.isna(value):
            d = pd.to_datetime(value, errors="coerce", dayfirst=True)
            if pd.notna(d):
                return pd.Timestamp(d).strftime("%Y-%m-%d")
        if fallback_ts is not None and not pd.isna(fallback_ts):
            fb = _to_ln(fallback_ts)
            if pd.notna(fb):
                return pd.Timestamp(fb).strftime("%Y-%m-%d")
        return None

    def _cmp_planning_color(row_dict):
        if str(row_dict.get("client_absent", "")).strip().upper() == "Y":
            return "#000000"
        key = str(row_dict.get("event_color_key", "")).strip().lower()
        return AUDIT_PLAN_COLOR_MAP.get(key, row_dict.get("plot_color", "#bdbdbd"))

    def _prestation_desc(code, text):
        code = str(code).strip() if pd.notna(code) else ""
        text = audit_clean_legend_text(text)
        if code:
            text = re.sub(rf"^\s*{re.escape(code)}\s*[-:–—]?\s*", "", text).strip()
        return text

    def _cmp_rda_orig_color(code):
        code = str(code).strip() if pd.notna(code) else ""
        if code in ["16009", "95900"]:
            return "#FFD700"
        if code == str(AUDIT_PRESTATION_61010_CODE):
            return "#800080"
        return "#2ca02c"

    LANE_Y = {"WF": 0.0, "RDA_ORIG": 1.0, "Planning": 2.0}

    map_local = map_df.copy()
    map_local["collab_id"] = map_local["collab_id"].astype(str)
    all_ids = set()
    for df_obj in [wf, rda, planning]:
        if not df_obj.empty and "collab_id" in df_obj.columns:
            all_ids.update(df_obj["collab_id"].dropna().astype(str).unique().tolist())
    collab_labels = {}
    for cid in sorted(all_ids):
        label = None
        if not map_local.empty:
            row = map_local[map_local["collab_id"] == cid]
            if not row.empty:
                rr = row.iloc[0]
                nm = _safe_text(rr.get("collab_name_sarl", "")).strip() or _safe_text(rr.get("collab_name_wf", "")).strip()
                sa = _safe_text(rr.get("collab_no_sarl", "")).strip() or "-"
                wf_no = _safe_text(rr.get("driverno", "")).strip() or "-"
                if nm:
                    label = f"{nm} (ID collab-{cid}, SA:{sa}, WF:{wf_no})"
        collab_labels[cid] = label or f"Collaborateur {cid}"

    rows_ev = []
    wf_local = wf.dropna(subset=["collab_id", "start", "end"]).copy()
    wf_local["collab_id"] = wf_local["collab_id"].astype(str)
    wf_stack = {}
    for rr in wf_local.sort_values(["collab_id", "start", "end"]).itertuples(index=False):
        s = _to_ln(rr.start)
        e = _to_ln(rr.end)
        if pd.isna(s) or pd.isna(e) or e <= s:
            continue
        cid = str(rr.collab_id)
        date_str = s.strftime("%Y-%m-%d")
        day_key = (cid, date_str)
        wf_idx = wf_stack.get(day_key, 0) + 1
        wf_stack[day_key] = wf_idx
        is_suspect = float(getattr(rr, "suspect_private_km", 0.0) or 0.0) > 0
        try:
            tm = int(rr.tripmode)
        except Exception:
            tm = -1
        fill_c = "#d62728" if is_suspect else {1: "#6c757d", 2: "#ff7f0e", 3: "#1f77b4"}.get(tm, "#999999")
        km_txt = f"{float(rr.km):.1f}km" if hasattr(rr, "km") and pd.notna(rr.km) else ""
        rows_ev.append({
            "collab_id": cid, "collab_label": collab_labels.get(cid, cid), "date_str": date_str,
            "kind": "WF", "y": LANE_Y["WF"], "left": s, "right": e, "mid": s + (e - s) / 2, "height": 0.34,
            "fill_color": fill_c, "line_color": "#7f0000" if is_suspect else "#202020",
            "label_text": f"{int(round(audit_duration_mins(s, e)))}m", "label_y": -0.20, "label_color": "#222222",
            "km_label": km_txt, "km_label_y": -0.38 if ((wf_idx - 1) % 2 == 0) else -0.48,
            "km_label_color": "#d62728" if is_suspect else "#111111", "wf_index": wf_idx,
        })

    rda_local = rda.dropna(subset=["collab_id", "start", "end"]).copy()
    rda_local["collab_id"] = rda_local["collab_id"].astype(str)
    for rr in rda_local.itertuples(index=False):
        s = _to_ln(rr.start)
        e = _to_ln(rr.end)
        if pd.isna(s) or pd.isna(e) or e <= s:
            continue
        cid = str(rr.collab_id)
        jour_value = getattr(rr, "jour", pd.NaT)
        date_str = _cmp_date_str(jour_value, fallback_ts=s)
        if not date_str:
            continue
        code = getattr(rr, "prestation_code", np.nan)
        prestation_text = getattr(rr, "prestation_text", "")
        code_key = str(code).strip() if pd.notna(code) else ""
        client_nr = audit_clean_legend_text(getattr(rr, "client_nr", ""))
        client_name = audit_clean_legend_text(getattr(rr, "client_name", ""))
        is_61010 = str(code) == str(AUDIT_PRESTATION_61010_CODE)
        rows_ev.append({
            "collab_id": cid, "collab_label": collab_labels.get(cid, cid), "date_str": date_str,
            "kind": "RDA_ORIG", "y": LANE_Y["RDA_ORIG"], "left": s, "right": e, "mid": s + (e - s) / 2, "height": 0.34,
            "fill_color": _cmp_rda_orig_color(code_key), "line_color": "#202020",
            "label_text": f"{int(round(audit_duration_mins(s, e)))}m", "label_y": 0.80,
            "label_color": "#b30000" if is_61010 else "#111111",
            "km_label": "", "km_label_y": np.nan, "km_label_color": "#111111", "wf_index": np.nan,
            "client_nr": client_nr, "client_label": client_name or (f"Client {client_nr}" if client_nr else ""),
            "event_color": "", "event_color_key": "",
            "prestation_code": code_key, "prestation_text": _prestation_desc(code_key, prestation_text),
        })

    plan_local = planning.dropna(subset=["collab_id", "start", "end"]).copy()
    plan_local["collab_id"] = plan_local["collab_id"].astype(str)
    for rr in plan_local.itertuples(index=False):
        s = _to_ln(rr.start)
        e = _to_ln(rr.end)
        if pd.isna(s) or pd.isna(e) or e <= s:
            continue
        cid = str(rr.collab_id)
        rr_dict = rr._asdict()
        client_nr = audit_clean_legend_text(rr_dict.get("client_nr", ""))
        client_label = audit_clean_legend_text(rr_dict.get("client_label", ""))
        event_color = audit_clean_legend_text(rr_dict.get("event_color", ""))
        event_color_key = audit_clean_legend_text(rr_dict.get("event_color_key", ""))
        plan_date_value = rr_dict.get("date_only", rr_dict.get("date", pd.NaT))
        date_str = _cmp_date_str(plan_date_value, fallback_ts=s)
        if not date_str:
            continue
        planning_color = _cmp_planning_color(rr_dict)
        rows_ev.append({
            "collab_id": cid, "collab_label": collab_labels.get(cid, cid), "date_str": date_str,
            "kind": "Planning", "y": LANE_Y["Planning"], "left": s, "right": e, "mid": s + (e - s) / 2, "height": 0.34,
            "fill_color": planning_color, "line_color": "#4A3B33",
            "label_text": f"{int(round(audit_duration_mins(s, e)))}m", "label_y": 1.80, "label_color": "#111111",
            "km_label": "", "km_label_y": np.nan, "km_label_color": "#111111", "wf_index": np.nan,
            "client_nr": client_nr, "client_label": client_label,
            "event_color": event_color, "event_color_key": event_color_key,
            "prestation_code": "", "prestation_text": "",
        })

    if not rows_ev:
        return None

    events_cmp = pd.DataFrame(rows_ev).sort_values(["collab_label", "date_str", "y", "left", "right"]).reset_index(drop=True)

    def _build_rda_rest_lookup():
        lookup = {}
        rd = rda.copy()
        rd = rd[rd["collab_id"].notna() & rd["start"].notna() & rd["end"].notna()].copy()
        if rd.empty:
            return lookup
        rd["collab_id"] = rd["collab_id"].astype(str)
        rd["start"] = rd["start"].apply(_to_ln)
        rd["end"] = rd["end"].apply(_to_ln)
        rd = rd[rd["start"].notna() & rd["end"].notna() & (rd["end"] > rd["start"])].copy()
        if "jour" in rd.columns:
            rd["date_str"] = rd.apply(lambda r: _cmp_date_str(r.get("jour", pd.NaT), fallback_ts=r["start"]), axis=1)
        else:
            rd["date_str"] = rd["start"].apply(lambda x: pd.Timestamp(x).strftime("%Y-%m-%d"))
        rd = rd[rd["date_str"].notna()].copy()
        daily = (rd.groupby(["collab_id", "date_str"]).agg(day_first_rda_start=("start", "min"), day_last_rda_end=("end", "max")).reset_index().sort_values(["collab_id", "date_str"]))
        for cid, grp in daily.groupby("collab_id"):
            grp = grp.sort_values("date_str").reset_index(drop=True)
            for i, row in grp.iterrows():
                cur_s = row["day_first_rda_start"]
                cur_e = row["day_last_rda_end"]
                prev_e = grp.loc[i - 1, "day_last_rda_end"] if i > 0 else pd.NaT
                next_s = grp.loc[i + 1, "day_first_rda_start"] if i < len(grp) - 1 else pd.NaT
                prev_rest = max(0.0, audit_duration_mins(prev_e, cur_s)) if pd.notna(prev_e) and pd.notna(cur_s) else np.nan
                next_rest = max(0.0, audit_duration_mins(cur_e, next_s)) if pd.notna(cur_e) and pd.notna(next_s) else np.nan
                lookup[(str(cid), str(row["date_str"]))] = {
                    "cur_start": cur_s, "cur_end": cur_e,
                    "prev_end": prev_e, "next_start": next_s,
                    "prev_rest_min": prev_rest, "next_rest_min": next_rest,
                }
        return lookup

    rda_rest_lookup = _build_rda_rest_lookup()
    return {"events_cmp": events_cmp, "collab_labels": collab_labels, "rda_rest_lookup": rda_rest_lookup}


# ============================================================
# Audit — single-day figure builder (requires matplotlib)
# ============================================================

def audit_build_day_fig(data_ctx: dict, result: dict, cid: str, date_str: str):
    """Render the Gantt chart for one collaborator/day. Returns matplotlib Figure or None."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.dates as mdates
        import matplotlib.patches as mpatches
    except ImportError:
        return None

    tz_name = AUDIT_TZ_NAME
    events_cmp = data_ctx["events_cmp"]
    rda_rest_lookup = data_ctx["rda_rest_lookup"]
    collab_labels = data_ctx["collab_labels"]
    agg_daily = result["agg_daily"]
    FIG_W, FIG_H = 19.0, 9.4

    def _to_ln(ts):
        return audit_to_local_naive(ts, tz_name)

    def _month_km_stats(c, d):
        out = {"month_label": "-", "km_private": 0.0, "km_professional": 0.0, "km_commute": 0.0,
               "km_total": 0.0, "km_suspect_private": 0.0, "km_private_plus_suspect": 0.0}
        if agg_daily.empty:
            return out
        base_day = pd.Timestamp(str(d))
        period = base_day.to_period("M")
        out["month_label"] = base_day.strftime("%B %Y")
        sub = agg_daily.copy()
        sub["collab_id"] = sub["collab_id"].astype(str)
        sub["date"] = pd.to_datetime(sub["date"], errors="coerce")
        sub = sub[(sub["collab_id"] == str(c)) & (sub["date"].dt.to_period("M") == period)]
        for key, col in [("km_private", "km_mode1"), ("km_professional", "km_mode2"), ("km_commute", "km_mode3"),
                         ("km_suspect_private", "suspect_private_km"), ("km_private_plus_suspect", "private_km_total_for_day")]:
            out[key] = float(sub[col].fillna(0).sum()) if col in sub.columns else 0.0
        out["km_total"] = out["km_private"] + out["km_professional"] + out["km_commute"]
        return out

    def _fmt_dt_short(ts):
        ts = _to_ln(ts)
        return "-" if pd.isna(ts) else pd.Timestamp(ts).strftime("%d/%m %H:%M")

    def _rest_texts(c, d):
        info = rda_rest_lookup.get((str(c), str(d)))
        if not info:
            return "REST BEFORE\n-", "REST AFTER\n-"
        pr = info["prev_rest_min"]
        nr = info["next_rest_min"]
        prev_text = ("REST BEFORE\n" + f"{_fmt_dt_short(info['prev_end'])} → {_fmt_dt_short(info['cur_start'])}\n{audit_fmt_hours(pd.Timedelta(minutes=float(pr)))}"
                     if pd.notna(pr) else "REST BEFORE\nAucun jour RDA précédent")
        next_text = ("REST AFTER\n" + f"{_fmt_dt_short(info['cur_end'])} → {_fmt_dt_short(info['next_start'])}\n{audit_fmt_hours(pd.Timedelta(minutes=float(nr)))}"
                     if pd.notna(nr) else "REST AFTER\nAucun jour RDA suivant")
        return prev_text, next_text

    def _wf_all_text(day_events):
        sub = day_events[day_events["kind"] == "WF"].sort_values(["left", "right"]).reset_index(drop=True)
        if sub.empty:
            return "ALL WF TRIPS\n-"
        lines = []
        for _, row in sub.iterrows():
            idx_txt = f"{int(row['wf_index']):02d}" if pd.notna(row.get("wf_index")) else "--"
            s = audit_fmt_hhmm(row["left"])
            e = audit_fmt_hhmm(row["right"])
            mins = audit_duration_mins(row["left"], row["right"])
            mins_txt = f"{int(round(mins))}m" if pd.notna(mins) else "-"
            km_txt = f" | {str(row.get('km_label', '')).strip()}" if str(row.get("km_label", "")).strip() else ""
            lines.append(f"{idx_txt}. {s} → {e} ({mins_txt}){km_txt}")
        return "ALL WF TRIPS\n" + "\n".join(lines)

    def _legend_items(day_events):
        client_items = []
        seen_clients = set()
        for _, row in day_events[day_events["kind"].isin(["RDA_ORIG", "Planning"])].sort_values(["kind", "left", "right"]).iterrows():
            client_nr = audit_clean_legend_text(row.get("client_nr", ""))
            label = audit_clean_legend_text(row.get("client_label", ""))
            if not client_nr:
                continue
            if not label:
                label = f"Client {client_nr}"
            if client_nr not in seen_clients:
                seen_clients.add(client_nr)
                client_items.append((client_nr, label))

        planning_color_items = []
        seen_plan_colors = set()
        for _, row in day_events[day_events["kind"] == "Planning"].sort_values(["event_color", "left", "right"]).iterrows():
            event_color = audit_clean_legend_text(row.get("event_color", ""))
            color = audit_clean_legend_text(row.get("fill_color", "#bdbdbd")) or "#bdbdbd"
            label = event_color or "event_color vide"
            key = (label.lower(), color)
            if key not in seen_plan_colors:
                seen_plan_colors.add(key)
                planning_color_items.append((label, color))

        prestation_items = []
        seen_prest = set()
        for _, row in day_events[day_events["kind"] == "RDA_ORIG"].sort_values(["left", "right"]).iterrows():
            code = audit_clean_legend_text(row.get("prestation_code", ""))
            if not code or code in seen_prest:
                continue
            seen_prest.add(code)
            prestation_items.append((
                code,
                audit_clean_legend_text(row.get("prestation_text", "")),
                audit_clean_legend_text(row.get("fill_color", "#2ca02c")),
            ))
        return client_items, planning_color_items, prestation_items

    def _wrap_client_label(label, max_len=18):
        label = audit_clean_legend_text(label)
        if len(label) <= max_len:
            return label
        parts = label.split()
        if len(parts) <= 1:
            return label
        lines = []
        cur = ""
        for part in parts:
            candidate = f"{cur} {part}".strip()
            if cur and len(candidate) > max_len:
                lines.append(cur)
                cur = part
            else:
                cur = candidate
        if cur:
            lines.append(cur)
        return "\n".join(lines[:2])

    def _draw_client_index_box(fig, rect, client_items):
        if not client_items:
            return
        rows = len(client_items)
        fs = 7.2 if rows <= 8 else 6.4 if rows <= 12 else 5.6
        ax_client = fig.add_axes(rect)
        ax_client.set_axis_off()
        ax_client.text(0.0, 1.0, "CLIENTS", transform=ax_client.transAxes,
                       ha="left", va="top", fontsize=fs + 0.5, fontweight="bold", color="#222222")
        top_y = 0.90
        row_step = 0.82 / max(1, rows)
        for idx, (client_nr, label) in enumerate(client_items):
            row = idx
            x = 0.0
            y = top_y - row * row_step
            max_name_len = 21
            client_txt = audit_shorten_text(client_nr, 9)
            if len(client_txt) > 4:
                client_prefix, client_suffix = client_txt[:-4], client_txt[-4:]
            else:
                client_prefix, client_suffix = "", client_txt
            prefix_pts = len(client_prefix) * fs * 0.55
            suffix_pts = len(client_suffix) * (fs + 2.0) * 0.58
            if client_prefix:
                ax_client.annotate(
                    client_prefix, xy=(x, y), xycoords=ax_client.transAxes,
                    xytext=(0, 0), textcoords="offset points",
                    ha="left", va="top", fontsize=max(5.2, fs - 1.3),
                    family="monospace", color="#222222", fontweight="normal",
                )
            ax_client.annotate(
                client_suffix, xy=(x, y), xycoords=ax_client.transAxes,
                xytext=(prefix_pts, 0), textcoords="offset points",
                ha="left", va="top", fontsize=fs + 2.0,
                family="monospace", color="#111111", fontweight="bold",
            )
            ax_client.annotate(
                f" = {audit_shorten_text(label, max_name_len)}", xy=(x, y), xycoords=ax_client.transAxes,
                xytext=(prefix_pts + suffix_pts + 4, 0), textcoords="offset points",
                ha="left", va="top", fontsize=fs,
                family="monospace", color="#222222", fontweight="bold",
            )

    def _draw_header_indexes(fig, plan_items, prest_items):
        _draw_lane_box(fig, [0.300, 0.785, 0.195, 0.095], "PLANNING COLORS", plan_items, fs=8.2)
        _draw_lane_box(fig, [0.505, 0.785, 0.210, 0.095], "PRESTATION INDEX", prest_items, fs=8.2)

    def _draw_lane_box(fig, rect, title, items, fs=6.5):
        if not items:
            return
        leg_ax = fig.add_axes(rect)
        leg_ax.set_axis_off()
        y = 0.98
        item_count = len(items)
        fs = min(fs, 8.0 if item_count <= 5 else 7.0 if item_count <= 8 else 6.0)
        line_h = min(0.150 if fs >= 7.4 else 0.125, 0.86 / max(1, item_count + 1))
        leg_ax.text(0.0, y, title, transform=leg_ax.transAxes, ha="left", va="top",
                    fontsize=fs + 0.6, fontweight="bold", color="#222222")
        y -= line_h
        for item in items:
            text, color = item if isinstance(item, tuple) else (str(item), None)
            x = 0.0
            if color:
                leg_ax.add_patch(mpatches.Rectangle((0.0, y - line_h * 0.62), 0.060, line_h * 0.54,
                                                    transform=leg_ax.transAxes, facecolor=color,
                                                    edgecolor="#333333", linewidth=0.4))
                x = 0.075
            leg_ax.text(x, y, text, transform=leg_ax.transAxes, ha="left", va="top",
                        fontsize=fs, color="#222222", linespacing=0.90)
            y -= line_h

    def _draw_right_index(fig, day_events, wf_all_text, wf_fs):
        client_items, planning_color_items, prestation_items = _legend_items(day_events)

        plan_items = [(audit_shorten_text(label, 28), color or "#bdbdbd") for label, color in planning_color_items]
        prest_items = []
        for code, desc, color in prestation_items:
            label = f"{code} {desc}".strip()
            prest_items.append((audit_shorten_text(label, 28), color or "#2ca02c"))

        _draw_header_indexes(fig, plan_items, prest_items)
        _draw_client_index_box(fig, [0.875, 0.365, 0.12, 0.365], client_items)
        fig.text(0.875, 0.285, wf_all_text, ha="left", va="top", fontsize=wf_fs,
                 family="monospace", color="#222222",
                 bbox=dict(facecolor="#ffffff", edgecolor="#cfcfcf", boxstyle="round,pad=0.38", alpha=0.96))

    def _day_bounds(day_events, day_str, min_hours=16):
        if day_events.empty:
            base = pd.Timestamp(f"{day_str} 00:00:00")
            return base, base + pd.Timedelta(hours=min_hours)
        left = _to_ln(day_events["left"].min())
        right = _to_ln(day_events["right"].max())
        if pd.isna(left) or pd.isna(right) or right <= left:
            base = pd.Timestamp(f"{day_str} 00:00:00")
            return base, base + pd.Timedelta(hours=min_hours)
        if (right - left) < pd.Timedelta(hours=min_hours):
            right = left + pd.Timedelta(hours=min_hours)
        return left, right

    def _lane_span(day_events, kind):
        sub = day_events[day_events["kind"] == kind]
        if sub.empty:
            return pd.NaT, pd.NaT
        s = _to_ln(sub["left"].min())
        e = _to_ln(sub["right"].max())
        return (s, e) if (pd.notna(s) and pd.notna(e) and e > s) else (pd.NaT, pd.NaT)

    def _draw_markers(ax, day_events):
        sub = day_events[day_events["kind"] == "RDA_ORIG"]
        if sub.empty:
            return
        rs = _to_ln(sub["left"].min())
        re_ = _to_ln(sub["right"].max())
        if pd.isna(rs) or pd.isna(re_) or re_ <= rs:
            return
        ymin, ymax = -0.36, 2.36
        for x, lab in [(rs, f"Start {rs.strftime('%H:%M')}"), (re_, f"End {re_.strftime('%H:%M')}")]:
            xx = mdates.date2num(x)
            ax.vlines(xx, ymin=ymin, ymax=ymax, colors="#555555", linestyles="--", linewidth=2.0, alpha=0.90, zorder=1)
            ax.text(xx, 2.52, lab, ha="center", va="bottom", fontsize=11.5, color="#555555", zorder=6, clip_on=False)
        span_hours = max(0.0, (re_ - rs).total_seconds() / 3600.0)
        for h in range(1, int(math.floor(span_hours)) + 2):
            x = rs + pd.Timedelta(hours=h)
            xx = mdates.date2num(x)
            ax.vlines(xx, ymin=ymin, ymax=ymax, colors="#9a9a9a", linestyles=":", linewidth=1.2, alpha=0.80, zorder=1)
            ax.text(xx, 2.45, f"{h}h", ha="center", va="bottom", fontsize=10, color="#6f6f6f", zorder=6, clip_on=False)

    def _draw_day(ax, day_events):
        def _draw_client_id_label(mid_num, y_pos, client_nr, kind):
            client_nr = audit_shorten_text(client_nr, 14)
            if len(client_nr) > 4:
                prefix, suffix = client_nr[:-4], client_nr[-4:]
            else:
                prefix, suffix = "", client_nr
            box_color = "#111111" if kind == "Planning" else "#202020"
            bbox = dict(facecolor=box_color, edgecolor="none", alpha=0.62, pad=0.12)
            if prefix:
                ax.annotate(
                    prefix, xy=(mid_num, y_pos), xytext=(0, -17), textcoords="offset points",
                    ha="center", va="center", rotation=90, fontsize=5.6, color="#ffffff",
                    fontweight="bold", zorder=7, clip_on=True,
                )
            ax.annotate(
                suffix, xy=(mid_num, y_pos), xytext=(0, 5), textcoords="offset points",
                ha="center", va="center", rotation=90, fontsize=9.4, color="#ffffff",
                fontweight="bold", zorder=7, clip_on=True, bbox=bbox,
            )

        for _, rr in day_events.iterrows():
            left = _to_ln(rr["left"])
            right = _to_ln(rr["right"])
            if pd.isna(left) or pd.isna(right) or right <= left:
                continue
            y = float(rr["y"])
            h = float(rr["height"])
            ax.broken_barh(
                [(mdates.date2num(left), mdates.date2num(right) - mdates.date2num(left))],
                (y - h / 2.0, h),
                facecolors=rr["fill_color"], edgecolors=rr["line_color"], linewidth=1.0, alpha=0.95, zorder=3,
            )
            mid = _to_ln(rr["mid"])
            lab = rr["label_text"]
            if pd.notna(mid) and str(lab).strip():
                ax.text(mdates.date2num(mid), rr["label_y"], lab, ha="center", va="center", fontsize=8.5,
                        color=rr["label_color"], zorder=6, clip_on=False, bbox=dict(facecolor="white", edgecolor="none", alpha=0.55, pad=0.18))
            if str(rr["kind"]) in ["RDA_ORIG", "Planning"]:
                client_nr = audit_clean_legend_text(rr.get("client_nr", ""))
                width_min = audit_duration_mins(left, right)
                if pd.notna(mid) and client_nr and pd.notna(width_min) and width_min >= 5:
                    _draw_client_id_label(mdates.date2num(mid), y, client_nr, str(rr["kind"]))
            if str(rr["kind"]) == "WF":
                wf_idx = rr.get("wf_index", np.nan)
                if pd.notna(mid) and pd.notna(wf_idx):
                    ax.text(mdates.date2num(mid), float(rr["y"]) + 0.26, f"{int(wf_idx)}", ha="center", va="center",
                            fontsize=8.3, color="#000000", fontweight="bold", zorder=7, clip_on=False,
                            bbox=dict(facecolor="white", edgecolor="#222222", alpha=0.88, boxstyle="round,pad=0.18"))
            km_lab = str(rr.get("km_label", "") or "").strip()
            km_y = rr.get("km_label_y", np.nan)
            if pd.notna(mid) and km_lab and pd.notna(km_y):
                ax.text(mdates.date2num(mid), km_y, km_lab, ha="center", va="center", fontsize=8.2,
                        color=rr["km_label_color"], zorder=6, clip_on=False, bbox=dict(facecolor="white", edgecolor="none", alpha=0.55, pad=0.15))

    collab_label = collab_labels.get(str(cid), f"Collaborateur {cid}")
    day_events = events_cmp[
        (events_cmp["collab_id"].astype(str) == str(cid)) &
        (events_cmp["date_str"].astype(str) == str(date_str))
    ].copy()
    if day_events.empty:
        return None
    day_events = day_events.sort_values(["y", "left", "right"]).reset_index(drop=True)
    left_bound, right_bound = _day_bounds(day_events, date_str)
    wf_s, wf_e = _lane_span(day_events, "WF")
    ro_s, ro_e = _lane_span(day_events, "RDA_ORIG")
    pl_s, pl_e = _lane_span(day_events, "Planning")
    km_stats = _month_km_stats(cid, date_str)
    km_line = (f"Month KM ({km_stats['month_label']})    |    Private: {km_stats['km_private']:.1f}km"
               f"    |    Suspect: {km_stats['km_suspect_private']:.1f}km"
               f"    |    Private+Suspect: {km_stats['km_private_plus_suspect']:.1f}km"
               f"    |    Dom-travail: {km_stats['km_commute']:.1f}km"
               f"    |    Professionnel: {km_stats['km_professional']:.1f}km"
               f"    |    Total: {km_stats['km_total']:.1f}km")
    wf_all_text = _wf_all_text(day_events)
    wf_count = int((day_events["kind"] == "WF").sum())
    wf_fs = 8.4 if wf_count <= 12 else 7.3 if wf_count <= 25 else 6.4 if wf_count <= 40 else 5.5 if wf_count <= 60 else 4.8
    prev_rest_text, next_rest_text = _rest_texts(cid, date_str)
    subtitle = (f"Frame: {audit_fmt_hhmm(left_bound)} → {audit_fmt_hhmm(right_bound)}"
                f"    |    WF span: {audit_fmt_span(wf_s, wf_e)}"
                f"    |    RDA original: {audit_fmt_span(ro_s, ro_e)}"
                f"    |    Planning: {audit_fmt_span(pl_s, pl_e)}")
    fig, ax = plt.subplots(figsize=(FIG_W, FIG_H), constrained_layout=False)
    _draw_markers(ax, day_events)
    _draw_day(ax, day_events)
    ax.set_xlim(mdates.date2num(left_bound), mdates.date2num(right_bound))
    ax.set_ylim(-0.52, 2.72)
    ax.margins(x=0, y=0)
    ax.set_yticks([0.0, 1.0, 2.0])
    ax.set_yticklabels(["WF Trips", "RDA Original", "Planning"], fontsize=11.5)
    span_hours = max(1.0, (right_bound - left_bound).total_seconds() / 3600.0)
    major_interval = 1 if span_hours <= 14 else 2 if span_hours <= 22 else 3
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=major_interval))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
    ax.tick_params(axis="x", labelsize=11)
    ax.grid(axis="x", linestyle=":", linewidth=0.8, alpha=0.45)
    ax.grid(axis="y", visible=False)
    for sp in ax.spines.values():
        sp.set_linewidth(0.8)
        sp.set_alpha(0.70)
    fig.suptitle(f"{collab_label} | {date_str}", fontsize=14, y=0.985, fontweight="bold")
    fig.text(0.5, 0.948, km_line, ha="center", va="center", fontsize=10.5, fontweight="bold",
             bbox=dict(facecolor="#f2f2f2", edgecolor="#d0d0d0", boxstyle="round,pad=0.30"))
    fig.text(0.5, 0.915, subtitle, ha="center", va="center", fontsize=9.2, color="#333333")
    fig.text(0.105, 0.875, prev_rest_text, ha="left", va="top", fontsize=8.7, family="monospace", color="#222222",
             bbox=dict(facecolor="#ffffff", edgecolor="#cfcfcf", boxstyle="round,pad=0.35", alpha=0.96))
    fig.text(0.855, 0.875, next_rest_text, ha="right", va="top", fontsize=8.7, family="monospace", color="#222222",
             bbox=dict(facecolor="#ffffff", edgecolor="#cfcfcf", boxstyle="round,pad=0.35", alpha=0.96))
    _draw_right_index(fig, day_events, wf_all_text, wf_fs)
    fig.subplots_adjust(left=0.10, right=0.855, top=0.72, bottom=0.12)
    return fig


# ============================================================
# Audit — PDF generation (optional, requires matplotlib)
# ============================================================

def audit_generate_pdfs(result: dict, progress_cb=None):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib.backends.backend_pdf import PdfPages
    except ImportError:
        raise RuntimeError("matplotlib n'est pas installé. Installez-le avec : pip install matplotlib")

    data_ctx = audit_build_chart_data(result)
    if data_ctx is None:
        return None

    events_cmp = data_ctx["events_cmp"]
    pairs = events_cmp[["collab_id", "collab_label", "date_str"]].drop_duplicates().copy()
    pairs["collab_id"] = pairs["collab_id"].astype(str)
    pairs["date_str"] = pairs["date_str"].astype(str)
    pairs = pairs.sort_values(["collab_label", "date_str"]).reset_index(drop=True)

    collab_groups = list(pairs.groupby("collab_label", sort=True))
    total_collabs = len(collab_groups)

    def _prog(i, label, pages=None):
        if not progress_cb:
            return
        pct = min(1.0, i / total_collabs) if total_collabs > 0 else 1.0
        pages_txt = f" — {pages} page(s)" if pages is not None else ""
        progress_cb(pct, f"PDF {i}/{total_collabs} : {label}{pages_txt}")

    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, (collab_label, grp) in enumerate(collab_groups):
            _prog(idx, collab_label)
            grp = grp.sort_values("date_str").reset_index(drop=True)
            cid_values = grp["collab_id"].dropna().astype(str).unique().tolist()
            cid = cid_values[0] if cid_values else "unknown"
            pdf_name = f"{audit_safe_filename(collab_label)}__collab_{cid}__wf_rda_planning.pdf"
            pdf_buf = BytesIO()
            pages_written = 0
            with PdfPages(pdf_buf) as pdf:
                for _, rr in grp.iterrows():
                    fig = audit_build_day_fig(data_ctx, result, str(rr["collab_id"]), str(rr["date_str"]))
                    if fig is None:
                        continue
                    pdf.savefig(fig, dpi=180, bbox_inches="tight")
                    plt.close(fig)
                    pages_written += 1
            _prog(idx + 1, collab_label, pages=pages_written)
            pdf_buf.seek(0)
            zf.writestr(pdf_name, pdf_buf.read())

    zip_buf.seek(0)
    return zip_buf


def audit_pdf_worker(result: dict, status: dict):
    def _progress(pct, text=None):
        status["progress"] = float(pct)
        if text:
            status["text"] = text

    status.update({"state": "running", "progress": 0.0, "text": "Initialisation des PDFs...", "error": None})
    try:
        zip_bytes = audit_generate_pdfs(result, progress_cb=_progress)
        if zip_bytes is None:
            status.update({"state": "empty", "progress": 1.0, "text": "Aucune donnée à exporter en PDF."})
        else:
            status.update({"state": "complete", "progress": 1.0, "text": "PDFs Gantt terminés."})
        return zip_bytes
    except Exception as exc:
        status.update({"state": "error", "error": str(exc), "text": f"Erreur PDFs : {exc}"})
        raise


def audit_start_pdf_job(result: dict, force: bool = False) -> None:
    future = st.session_state.get("latest_audit_pdf_future")
    if future is not None and not future.done() and not force:
        return

    if force:
        st.session_state.pop("latest_audit_pdf_zip", None)

    executor = st.session_state.get("audit_pdf_executor")
    if executor is None:
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=1, thread_name_prefix="audit-pdf")
        st.session_state["audit_pdf_executor"] = executor

    status = {"state": "queued", "progress": 0.0, "text": "PDFs Gantt en attente...", "error": None}
    st.session_state["latest_audit_pdf_status"] = status
    st.session_state["latest_audit_pdf_future"] = executor.submit(audit_pdf_worker, result, status)


def audit_collect_pdf_job() -> None:
    future = st.session_state.get("latest_audit_pdf_future")
    if future is None or not future.done():
        return
    status = st.session_state.get("latest_audit_pdf_status", {})
    if status.get("collected"):
        return
    try:
        zip_bytes = future.result()
        if zip_bytes:
            zip_bytes.seek(0)
            st.session_state["latest_audit_pdf_zip"] = zip_bytes
    except Exception:
        pass
    status["collected"] = True


def audit_render_pdf_controls(result: dict) -> None:
    audit_collect_pdf_job()
    status = st.session_state.get("latest_audit_pdf_status")
    future = st.session_state.get("latest_audit_pdf_future")

    p1, p2 = st.columns([2, 1])
    with p1:
        if status:
            state = status.get("state", "queued")
            text = status.get("text", "PDFs Gantt en cours...")
            pct = float(status.get("progress", 0.0) or 0.0)
            if state in ["queued", "running"] and future is not None and not future.done():
                st.progress(min(max(pct, 0.0), 1.0), text=text)
                st.caption("Vous pouvez utiliser le dashboard pendant la génération.")
            elif state == "complete":
                st.success(text)
            elif state == "empty":
                st.info(text)
            elif state == "error":
                st.error(text)
        else:
            st.caption("Les PDFs Gantt démarrent automatiquement après l'audit.")
    with p2:
        if st.button("Relancer les PDFs", key="audit_gen_pdf", use_container_width=True):
            audit_start_pdf_job(result, force=True)
            st.rerun()
        pdf_zip = st.session_state.get("latest_audit_pdf_zip")
        if pdf_zip:
            pdf_zip.seek(0)
            st.download_button(
                "Télécharger les PDFs (zip)", pdf_zip,
                file_name="audit_pdfs_gantt.zip", mime="application/zip",
                use_container_width=True,
            )


@st.fragment(run_every="2s")
def audit_render_pdf_controls_live(result: dict) -> None:
    audit_render_pdf_controls(result)
    future = st.session_state.get("latest_audit_pdf_future")
    if future is None or future.done():
        st.rerun()


# ============================================================
# Audit — dashboard render
# ============================================================

def render_audit_dashboard(result: dict) -> None:
    audit_collect_pdf_job()

    metrics = result["metrics"]
    collab_stats = result["collab_stats"]
    flag_summary = result["flag_summary"]
    daily_usage = result["daily_usage"]
    agg_daily = result["agg_daily"]
    wf_df = result["wf"]

    # --- Metrics row ---
    m_cols = st.columns(4)
    m_cols[0].metric("Trajets WF total", f"{metrics['total_trips']:,}")
    m_cols[1].metric("Collaborateurs", f"{metrics['total_collabs']:,}")
    m_cols[2].metric("% Trajets flaggés", f"{metrics['flagged_trip_pct']:.1f}%")
    m_cols[3].metric("KM suspects total", f"{metrics['suspect_km_total']:,.1f} km")

    st.divider()

    # --- PDF generation ---
    pdf_future = st.session_state.get("latest_audit_pdf_future")
    if pdf_future is not None and not pdf_future.done():
        audit_render_pdf_controls_live(result)
    else:
        audit_render_pdf_controls(result)

    # --- In-UI Gantt viewer ---
    chart_data = st.session_state.get("latest_audit_chart_data")
    pdf_future = st.session_state.get("latest_audit_pdf_future")
    pdf_running = pdf_future is not None and not pdf_future.done()
    if chart_data is None and not pdf_running:
        chart_data = audit_build_chart_data(result)
        if chart_data is not None:
            st.session_state["latest_audit_chart_data"] = chart_data

    if chart_data is not None:
        events_cmp = chart_data["events_cmp"]
        collab_options = (
            events_cmp[["collab_id", "collab_label"]]
            .drop_duplicates()
            .sort_values("collab_label")
            .reset_index(drop=True)
        )
        collab_display = [str(r["collab_label"]) for _, r in collab_options.iterrows()]
        collab_id_map = {str(r["collab_label"]): str(r["collab_id"]) for _, r in collab_options.iterrows()}

        with st.expander("Visualiser un graphique Gantt", expanded=False):
            vc1, vc2, vc3 = st.columns([3, 2, 1])
            with vc1:
                sel_collab_label = st.selectbox("Collaborateur", collab_display, key="audit_chart_collab")
            sel_cid = collab_id_map.get(sel_collab_label or "", "")
            dates_for_collab = sorted(
                events_cmp[events_cmp["collab_id"].astype(str) == sel_cid]["date_str"].unique().tolist()
            ) if sel_cid else []
            with vc2:
                sel_date = st.selectbox("Date", dates_for_collab, key="audit_chart_date")
            with vc3:
                st.write("")
                st.write("")
                show_chart_btn = st.button("Afficher", key="audit_show_chart", use_container_width=True)
            if show_chart_btn and sel_cid and sel_date:
                render_blocking_run_warning()
                with st.spinner("Génération du graphique..."):
                    try:
                        import matplotlib.pyplot as plt
                        fig = audit_build_day_fig(chart_data, result, sel_cid, sel_date)
                        if fig:
                            img_buf = BytesIO()
                            fig.savefig(img_buf, format="png", dpi=120, bbox_inches="tight")
                            plt.close(fig)
                            img_buf.seek(0)
                            st.session_state["audit_gantt_img"] = {
                                "key": (sel_cid, sel_date),
                                "data": img_buf.getvalue(),
                            }
                        else:
                            st.session_state["audit_gantt_img"] = None
                            st.info("Aucune donnée pour ce collaborateur et cette date.")
                    except Exception as exc:
                        st.exception(exc)

            gantt_img = st.session_state.get("audit_gantt_img")
            if gantt_img and gantt_img.get("key") == (sel_cid, sel_date):
                st.image(gantt_img["data"], use_container_width=True)
            elif gantt_img and gantt_img.get("key") != (sel_cid, sel_date):
                st.caption("Sélection modifiée — cliquez « Afficher » pour régénérer.")
    elif pdf_running:
        st.caption("Le visualiseur Gantt sera disponible quand la génération automatique des PDFs aura avancé ou terminé.")

    st.divider()

    # --- Charts section (dataframe-based to avoid sticky Vega-Lite tooltips) ---
    with st.expander("Graphiques de synthèse", expanded=True):
        chart_cols = st.columns(2)
        with chart_cols[0]:
            st.subheader("Répartition usage quotidien")
            if not daily_usage.empty and "usage_label" in daily_usage.columns:
                usage_counts = daily_usage["usage_label"].value_counts().reset_index()
                usage_counts.columns = ["Usage", "Jours"]
                usage_max = int(usage_counts["Jours"].max()) if not usage_counts.empty else 1
                st.dataframe(
                    usage_counts,
                    column_config={"Jours": st.column_config.ProgressColumn("Jours", min_value=0, max_value=usage_max, format="%d")},
                    hide_index=True, use_container_width=True,
                )
            else:
                st.info("Aucune donnée d'usage.")

        with chart_cols[1]:
            st.subheader("Répartition des flags")
            if not flag_summary.empty and "flag" in flag_summary.columns:
                flag_df = flag_summary[["flag", "count"]].copy()
                flag_df.columns = ["Flag", "Occurrences"]
                flag_max = int(flag_df["Occurrences"].max()) if not flag_df.empty else 1
                st.dataframe(
                    flag_df,
                    column_config={"Occurrences": st.column_config.ProgressColumn("Occurrences", min_value=0, max_value=flag_max, format="%d")},
                    hide_index=True, use_container_width=True,
                )
            else:
                st.info("Aucun flag détecté.")

        chart_cols2 = st.columns(2)
        with chart_cols2[0]:
            st.subheader("KM par mode (top 15 collaborateurs)")
            if not collab_stats.empty:
                top15 = collab_stats.nlargest(15, "km_total")[["collab_id", "km_m1", "km_m2", "km_m3", "km_total"]].copy()
                top15.columns = ["Collaborateur", "Privé (m1)", "Pro (m2)", "Dom-travail (m3)", "Total"]
                km_max = int(top15["Total"].max()) if not top15.empty else 1
                st.dataframe(
                    top15,
                    column_config={
                        "Privé (m1)": st.column_config.ProgressColumn("Privé (m1)", min_value=0, max_value=km_max, format="%.1f km"),
                        "Pro (m2)": st.column_config.ProgressColumn("Pro (m2)", min_value=0, max_value=km_max, format="%.1f km"),
                        "Dom-travail (m3)": st.column_config.ProgressColumn("Dom-travail (m3)", min_value=0, max_value=km_max, format="%.1f km"),
                        "Total": st.column_config.NumberColumn("Total km", format="%.1f km"),
                    },
                    hide_index=True, use_container_width=True,
                )
            else:
                st.info("Aucune donnée.")

        with chart_cols2[1]:
            st.subheader("Top 15 collaborateurs — KM suspects")
            if not collab_stats.empty and "suspect_private_km" in collab_stats.columns:
                top_susp = collab_stats.nlargest(15, "suspect_private_km")[["collab_id", "suspect_private_km"]].copy()
                top_susp.columns = ["Collaborateur", "KM suspects"]
                susp_max = float(top_susp["KM suspects"].max()) if not top_susp.empty else 1.0
                st.dataframe(
                    top_susp,
                    column_config={"KM suspects": st.column_config.ProgressColumn("KM suspects", min_value=0, max_value=susp_max, format="%.1f km")},
                    hide_index=True, use_container_width=True,
                )
            else:
                st.info("Aucune donnée.")

        if not agg_daily.empty and "date" in agg_daily.columns and "suspect_private_km" in agg_daily.columns:
            st.subheader("Évolution KM suspects par jour")
            trend = agg_daily.copy()
            trend["date"] = pd.to_datetime(trend["date"], errors="coerce")
            trend = (trend.dropna(subset=["date"])
                     .groupby("date")["suspect_private_km"].sum()
                     .reset_index().sort_values("date"))
            trend.columns = ["Date", "KM suspects"]
            trend_max = float(trend["KM suspects"].max()) if not trend.empty else 1.0
            st.dataframe(
                trend,
                column_config={
                    "Date": st.column_config.DateColumn("Date"),
                    "KM suspects": st.column_config.ProgressColumn("KM suspects", min_value=0, max_value=trend_max, format="%.1f km"),
                },
                hide_index=True, use_container_width=True,
            )

    # --- Data tabs ---
    tab_names = ["Résumé Collaborateurs", "Agrégats Journaliers", "Tous les Trajets", "Entrées RDA", "Planning"]
    tabs = st.tabs(tab_names)

    with tabs[0]:
        f0c1, f0c2 = st.columns([3, 1])
        with f0c1:
            search_collab = st.text_input("Rechercher", key="audit_collab_search", placeholder="ID ou nom...")
        with f0c2:
            only_suspect_cs = st.checkbox("Suspects uniquement", key="audit_collab_only_suspect")
        cs_display = collab_stats.copy()
        if search_collab.strip():
            mask = cs_display.apply(lambda r: search_collab.lower() in str(r).lower(), axis=1)
            cs_display = cs_display[mask]
        if only_suspect_cs and "suspect_private_km" in cs_display.columns:
            cs_display = cs_display[cs_display["suspect_private_km"] > 0]
        st.caption(f"{len(cs_display):,} collaborateurs")
        st.dataframe(cs_display, use_container_width=True, hide_index=True)

    with tabs[1]:
        agg = agg_daily.copy()
        f1c1, f1c2, f1c3 = st.columns(3)
        with f1c1:
            agg_collabs = sorted(agg["collab_id"].dropna().astype(str).unique().tolist()) if not agg.empty and "collab_id" in agg.columns else []
            sel_agg_collabs = st.multiselect("Collaborateur(s)", agg_collabs, key="audit_agg_collab_filter")
        with f1c2:
            date_range_agg = None
            if not agg.empty and "date" in agg.columns:
                agg_dates = pd.to_datetime(agg["date"], errors="coerce").dropna()
                if not agg_dates.empty:
                    min_d, max_d = agg_dates.min().date(), agg_dates.max().date()
                    date_range_agg = st.date_input("Période", value=(min_d, max_d), min_value=min_d, max_value=max_d, key="audit_agg_date_range")
        with f1c3:
            only_suspect_agg = st.checkbox("Suspects uniquement", key="audit_agg_only_suspect")
        if sel_agg_collabs:
            agg = agg[agg["collab_id"].astype(str).isin(sel_agg_collabs)]
        if date_range_agg and len(date_range_agg) == 2:
            agg_dc = pd.to_datetime(agg["date"], errors="coerce")
            agg = agg[(agg_dc.dt.date >= date_range_agg[0]) & (agg_dc.dt.date <= date_range_agg[1])]
        if only_suspect_agg and "suspect_private_km" in agg.columns:
            agg = agg[agg["suspect_private_km"] > 0]
        st.caption(f"{len(agg):,} entrées")
        st.dataframe(agg, use_container_width=True, hide_index=True)

    with tabs[2]:
        wf_filtered = wf_df.copy()
        f2c1, f2c2, f2c3 = st.columns(3)
        with f2c1:
            if "flags" in wf_df.columns:
                all_flags = sorted(set(
                    f.strip()
                    for flags_str in wf_df["flags"].fillna("").astype(str)
                    for f in flags_str.split(",")
                    if f.strip()
                ))
            else:
                all_flags = []
            sel_flags = st.multiselect("Flag(s)", all_flags, key="audit_trip_flag_filter")
        with f2c2:
            trip_collabs = sorted(wf_df["collab_id"].dropna().astype(str).unique().tolist()) if "collab_id" in wf_df.columns else []
            sel_trip_collabs = st.multiselect("Collaborateur(s)", trip_collabs, key="audit_trip_collab_filter")
        with f2c3:
            only_suspect_trips = st.checkbox("Suspects uniquement", key="audit_trip_only_suspect")
        if sel_flags:
            wf_filtered = wf_filtered[wf_filtered["flags"].fillna("").apply(lambda x: any(f in x.split(",") for f in sel_flags))]
        if sel_trip_collabs:
            wf_filtered = wf_filtered[wf_filtered["collab_id"].astype(str).isin(sel_trip_collabs)]
        if only_suspect_trips and "suspect_private_km" in wf_filtered.columns:
            wf_filtered = wf_filtered[wf_filtered["suspect_private_km"] > 0]
        st.caption(f"{len(wf_filtered):,} trajets")
        st.dataframe(audit_drop_tz_excel_safe(wf_filtered), use_container_width=True, hide_index=True)

    with tabs[3]:
        rda_f = result["rda"].copy()
        rda_collabs = sorted(rda_f["collab_id"].dropna().astype(str).unique().tolist()) if "collab_id" in rda_f.columns else []
        sel_rda_collabs = st.multiselect("Collaborateur(s)", rda_collabs, key="audit_rda_collab_filter")
        if sel_rda_collabs:
            rda_f = rda_f[rda_f["collab_id"].astype(str).isin(sel_rda_collabs)]
        st.caption(f"{len(rda_f):,} entrées")
        st.dataframe(audit_drop_tz_excel_safe(rda_f), use_container_width=True, hide_index=True)

    with tabs[4]:
        plan_f = result["planning"].copy()
        plan_collabs = sorted(plan_f["collab_id"].dropna().astype(str).unique().tolist()) if "collab_id" in plan_f.columns else []
        sel_plan_collabs = st.multiselect("Collaborateur(s)", plan_collabs, key="audit_plan_collab_filter")
        if sel_plan_collabs:
            plan_f = plan_f[plan_f["collab_id"].astype(str).isin(sel_plan_collabs)]
        st.caption(f"{len(plan_f):,} entrées")
        st.dataframe(audit_drop_tz_excel_safe(plan_f), use_container_width=True, hide_index=True)


# ============================================================
# Audit — task entry point
# ============================================================

def render_audit_task() -> None:
    st.title("Audit Webfleet-RDA")
    st.caption("Croise les données Webfleet, RDA et planning pour détecter les usages suspects du véhicule de fonction.")

    upload_cols = st.columns(4)
    rda_file = upload_cols[0].file_uploader("Fichier RDA", type=["xlsx", "xls", "csv"], key="audit_rda")
    wf_file = upload_cols[1].file_uploader("Fichier Webfleet", type=["xlsx", "xls", "csv"], key="audit_wf")
    mapping_file = upload_cols[2].file_uploader("Fichier Mapping", type=["xlsx", "xls"], key="audit_map")
    planning_file = upload_cols[3].file_uploader("Fichier Planning", type=["xlsx", "xls"], key="audit_plan")

    all_uploaded = all([rda_file, wf_file, mapping_file, planning_file])

    result = st.session_state.get("latest_audit_result")
    excel_path = result["excel_path"] if result else None

    action_cols = st.columns([2, 1])
    with action_cols[0]:
        run_audit = st.button("Lancer l'audit", type="primary", disabled=not all_uploaded, width="stretch")

    if run_audit:
        render_blocking_run_warning()
        progress = st.progress(0.0, text="Démarrage de l'audit...")
        try:
            result = audit_process(rda_file, wf_file, mapping_file, planning_file, progress_cb=progress.progress)
            progress.progress(1.0, text="Audit terminé")
            st.session_state["latest_audit_result"] = result
            st.session_state.pop("latest_audit_pdf_zip", None)
            st.session_state.pop("latest_audit_pdf_future", None)
            st.session_state.pop("latest_audit_pdf_status", None)
            st.session_state.pop("latest_audit_chart_data", None)
            st.session_state.pop("audit_gantt_img", None)
            audit_start_pdf_job(result)
        except Exception as exc:
            progress.empty()
            st.exception(exc)
            return

    result = st.session_state.get("latest_audit_result")
    excel_path = result["excel_path"] if result else None
    with action_cols[1]:
        render_download_or_placeholder(excel_path, "Télécharger le rapport Excel", key="audit_main_excel")

    if result:
        st.success("Rapport Excel créé et disponible au téléchargement.")
        render_audit_dashboard(result)


def render_placeholder_task(title: str) -> None:
    st.title(title)
    st.info("Cette tâche est préparée pour la prochaine passe d'intégration. Le notebook a été revu ; sa logique sera déplacée dans des fonctions appelables sans modifier les entrées ni les sorties.")
