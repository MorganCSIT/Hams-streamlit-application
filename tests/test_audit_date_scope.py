import unittest
from datetime import date
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import pandas as pd
import audit_webfleet_rda as audit
from ltr_checks import ltr_complete_months, ltr_history_warnings, ltr_scope_sheets, ltr_load_notebook_functions, ltr_notebook_mtime, ltr_process


def upload(frame, name, sheet="Sheet1"):
    stream = BytesIO()
    with pd.ExcelWriter(stream, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name=sheet)
    stream.seek(0)
    stream.name = name
    return stream


class AuditDateScopeTests(unittest.TestCase):
    def test_audit_period_applies_to_sources_summaries_and_cut_export(self):
        days = ["2026-07-31", "2026-08-01", "2026-08-02"]
        rda = pd.DataFrame([{"Jour": d, "Début": d+" 08:00", "Fin": d+" 09:00", "Durée": 60,
                             "No collaborateur": 1, "Collaborateur": "Test", "No prestation": 11200} for d in days])
        wf = pd.DataFrame([dict(tripid=str(i), tripmode=2, start_time=d+" 08:00", end_time=d+" 08:15", driverno=1, km=5) for i,d in enumerate(days)])
        planning = pd.DataFrame([dict(emp_nr=1, date=d, start="08:00", end="09:00", duration=60, event_color="INF", client_absent="N") for d in days])
        mapping = pd.DataFrame([{"collab_id":"a", "No collaborateur":1, "driverno":1}])
        with TemporaryDirectory() as tmp, patch.object(audit, "get_session_output_root", return_value=Path(tmp)):
            result = audit.audit_process(upload(rda,"rda.xlsx"), upload(wf,"wf.xlsx"), upload(mapping,"mapping.xlsx"), upload(planning,"planning.xlsx"), date_range=(date(2026,8,1),date(2026,8,1)))
            self.assertEqual(len(result["rda"]),1)
            self.assertEqual(len(result["planning_comparison"]),1)
            self.assertEqual(result["metrics"]["total_trips"],1)
            self.assertEqual(result["rda_daily"].date.nunique(),1)
            self.assertEqual(result["rda_source_df"].iloc[0]["Jour"],"2026-08-01")
            self.assertEqual(result["rda"].iloc[0].rda_row_id,0)
            cutting = audit.audit_run_rda_cutting(result)
            self.assertEqual(len(cutting["rda_cut"]),1)
            self.assertEqual(len(cutting["rda_input_export"]),1)

    def test_ltr_process_exports_only_selected_dates_with_full_lookback(self):
        days = pd.date_range("2026-07-15", "2026-09-02")
        rda = pd.DataFrame([{"Jour":d.strftime("%d.%m.%Y"), "Début":d.strftime("%d.%m.%Y")+" 08:00", "Fin":d.strftime("%d.%m.%Y")+" 17:00", "Durée":540,
                             "No collaborateur":1,"Collaborateur":"Test", "No prestation":11200,"Prestation":"Soins","UO":"Home Assistance SA"} for d in days])
        rda.loc[days == pd.Timestamp("2026-08-30"), "Fin"] = "30.08.2026 23:00"
        mapping = pd.DataFrame([{"collaborateur-id":"a", "name-collaborateur":"Test", "no-collaborateur-sa-101":1,"no-collaborateur-sarl-102":None,"no-collaborateur-ne-103":None}])
        with TemporaryDirectory() as tmp:
            result = ltr_process(upload(mapping,"map.xlsx","Matched Collaborateurs"), upload(rda,"rda.xlsx"), output_root=Path(tmp), date_range=(date(2026,8,31), date(2026,9,1)), exclude_incomplete_months=True)
            self.assertTrue(all(result["services_audit"].service_date.between(date(2026,8,31),date(2026,9,1))))
            self.assertFalse(result["all_infractions"].TARGET_MONTH.eq("2026-09").any())
            review = result["rest_review"]
            target = review[review.service_date.eq(date(2026,8,31))].iloc[0]
            self.assertEqual(target.rest_decision,"ALLOWED_REDUCED_REST")
            self.assertTrue(pd.notna(target.roll14_back_true_hours))
            self.assertTrue(result["warnings"])
            self.assertTrue(result["workbook_path"].exists())

    def test_complete_months_and_history_warning(self):
        dates = pd.date_range("2026-07-15", "2026-09-02")
        self.assertEqual(ltr_complete_months(dates, today=date(2026,10,2)), {"2026-08"})
        self.assertFalse(ltr_history_warnings(dates,date(2026,8,1),date(2026,8,10)))
        self.assertTrue(ltr_history_warnings(dates,date(2026,7,20),date(2026,8,10)))

    def test_ltr_scopes_results_after_checks_and_excludes_partial_month(self):
        infra = pd.DataFrame({"EVENT_DATE":pd.to_datetime(["2026-07-31","2026-08-01","2026-08-02","2026-09-01"]),"TARGET_MONTH":["2026-07","2026-08","2026-08","2026-09"]})
        tables = ltr_scope_sheets({"ALL_INFRACTIONS":infra,"SERVICES_AUDIT":infra}, (date(2026,8,2),date(2026,9,1)), {"2026-08"})
        self.assertEqual(len(tables["ALL_INFRACTIONS"]),1)
        self.assertEqual(len(tables["SERVICES_AUDIT"]),2)

    def test_rest_average_keeps_history_but_marks_partial_history_unknown(self):
        env = ltr_load_notebook_functions(ltr_notebook_mtime())
        days = pd.date_range("2026-07-15","2026-08-01")
        services = pd.DataFrame([dict(collab_key="a",collab_uid="a",Collaborateur="Test",No_collaborateur_codes="1",Match_status="MATCHED",service_id=str(i),service_date=d.date(),service_start=d+pd.Timedelta(hours=8), service_end=d+pd.Timedelta(hours=17)) for i,d in enumerate(days)])
        services.loc[services.index[-2],"service_end"] = pd.Timestamp("2026-07-31 23:00")
        _, _, review = env["check_rest_under_11h"](services)
        selected = review[review.service_date.eq(date(2026,8,1))].iloc[0]
        self.assertEqual(selected.rest_decision,"ALLOWED_REDUCED_REST")
        self.assertTrue(pd.notna(selected.roll14_back_true_hours))
        _, _, review = env["check_rest_under_11h"](services.tail(4))
        self.assertEqual(review.iloc[-1].rest_decision,"REVIEW_MISSING_14D_CONTEXT")
        self.assertTrue(pd.isna(review.iloc[-1].roll14_back_true_hours))
