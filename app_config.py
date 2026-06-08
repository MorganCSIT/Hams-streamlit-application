import csv
import json
import logging
import os
import re
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, time as dtime, timedelta
from io import BytesIO, StringIO
from pathlib import Path
from zoneinfo import ZoneInfo

import math

import numpy as np
import pandas as pd
import requests
import streamlit as st


APP_ROOT = Path(__file__).resolve().parent
BASE_URL = "https://csv.webfleet.com/extern"
API_ACTION = "showTripReportExtern"
LANG = "en"
OUTPUTFORMAT = "csv"
EXCEL_MAX_ROWS = 1_048_576
MISSING_FILTER_LABEL = "(Manquant)"
DASHBOARD_COLUMNS = [
    "tripmode",
    "start_time",
    "end_time",
    "duration",
    "distance",
    "drivername",
    "driverno",
    "objectname",
]
MERGE_OUTPUT_FOLDER = "MergedReports"
RDA_OUTPUT_FOLDER = "RDAReports"
LTR_OUTPUT_FOLDER = "LTRReports"
LTR_NOTEBOOK_PATH = APP_ROOT / "Scripts" / "Check LTR.ipynb"
AUDIT_OUTPUT_FOLDER = "AuditReports"
AUDIT_WORK_END_BUFFER_MIN = 30
AUDIT_PRE_SHIFT_BUFFER_MIN = 30
AUDIT_INTERNAL_BLOCK_GAP_MIN = 180
AUDIT_INTERNAL_BUFFER_MIN = 30
AUDIT_INTERNAL_BUF_MIN_OVERLAP_MIN = 0
AUDIT_INTERNAL_BUF_MIN_OVERLAP_RATIO = 0.0
AUDIT_MAX_REASONABLE_SPEED_KMH = 160
AUDIT_KM_MIN_FOR_FLAG_CONTEXT = 0.1
AUDIT_FULL_DAY_MINUTES = 420
AUDIT_FULL_SPAN_MINUTES = 480
AUDIT_MIN_BLOCKS_FOR_SPAN = 2
AUDIT_BLOCKS_FULL = 4
AUDIT_GAP_MERGE_MIN = 10
AUDIT_PRESTATION_61010_CODE = "61010"
AUDIT_PRESTATION_61010_BUFFER_MIN = 2
AUDIT_ENABLE_61010_FEATURE = True
AUDIT_TZ_NAME = "Europe/Zurich"
AUDIT_CHECK_PRE_SHIFT = True
AUDIT_PLAN_COLOR_MAP = {
    "as": "#FFD166",
    "inf": "#F15BB5",
    "adv": "#9C6644",
    "annule": "#D4A373",
    "annul?": "#D4A373",
    "demande d'horaire specifique": "#E9C46A",
    "demande d'horaire sp?cifique": "#E9C46A",
    "facturation (-24h)": "#7F5539",
    "avertir": "#FF99C8",
    "a ete averti - adv": "#B08968",
    "a ?t? averti - adv": "#B08968",
    "geplant": "#A68A64",
    "a avertir - inf": "#E6BE8A",
    "? avertir - inf": "#E6BE8A",
    "a avertir - adv": "#C9ADA7",
    "? avertir - adv": "#C9ADA7",
    "a ete averti - as": "#F6D6AD",
    "a ?t? averti - as": "#F6D6AD",
}
AUDIT_PLANNING_DATE_TIE_BREAKER = "monthfirst"
AUDIT_PLANNING_DATE_MIN_YEAR = 2000
AUDIT_PLANNING_DATE_MAX_YEAR = 2100
TASKS = {
    "home": "Accueil",
    "webfleet": "Téléchargement Webfleet",
    "merge": "Fusionner des fichiers",
    "rda": "Transferts RDA",
    "ltr": "Contrôles LTR",
    "audit": "Audit Webfleet-RDA",
}
RDA_OE_MAP = {
    "NE 301": "100000000000000301",
    "SARL 201": "100000000000000201",
    "SA 101": "100000000000000101",
}
RDA_TRANSFER_TARGET_OE = "100000000000000101"
RDA_ALLOWED_TARGET_CODES = {"11000", "11100", "11200", "14000", "14100", "14200"}
RDA_CODE_61010 = "61010"
RDA_WHITELIST_CODES = {"16011", "909", "16009", "195"}
RDA_DATE_COLS = ["Date Début", "Date", "Jour", "Date de prestation"]
RDA_START_COLS = ["Début", "Heure début", "Heure Début", "Von"]
RDA_END_COLS = ["Fin", "Heure fin", "Heure Fin", "Bis"]
RDA_CODE_COLS = ["N° Prestation", "No prestation", "Code prestation", "Prestation", "Code"]
RDA_DUREE_COLS = ["Durée", "Duree", "Durée (min)", "Durée (minutes)", "Dauer_verrechnet"]
RDA_CLIENT_COLS = ["N° du client", "No client", "ID client", "Client", "KD-Nr", "KD_Nr"]
RDA_COLLAB_COLS = ["No collaborateur", "Collaborateur", "ID collaborateur", "Mitarbeiter-ID"]


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


def get_session_output_root(folder_name: str) -> Path:
    session_id = st.session_state.get("_output_session_id")
    if not session_id:
        session_id = uuid.uuid4().hex
        st.session_state["_output_session_id"] = session_id

    output_root = Path(tempfile.gettempdir()) / "webfleet_tools_outputs" / session_id / folder_name
    output_root.mkdir(parents=True, exist_ok=True)
    return output_root


