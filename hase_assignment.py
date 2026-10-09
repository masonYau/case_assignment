import argparse
import json
import logging
import re
import posixpath
import sys
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile
from pathlib import Path
from typing import Optional, Sequence, Tuple
import numpy as np
import pandas as pd
from case_assignment import CDDTaskAllocator
from header import AnalystHeader as ah
from header import ClosedReviewHeader as ch
from header import MasterGroupHeader as mh
from header import OpenReviewHeader as oh
from header import StageType as st
from header import TaskHeader as th


class DefaultFields:
    class CaseList:
        case_id = "Review ID"
        customer = "CIN"
        cin = "CIN"
        segment = "Team code"
        assign_date = "Assign Date"
        designated_ocm = "Designated OCM"
        trigger_type = "Trigger type"
        due_date = "Planned T0 date"
        workload = "Workload Units"
        business_cm_id = "Staff ID"
        business_cm_name = "Case manager"
        business_team = "CM Team"
        mg_id = "Master Group ID"
        mg_name = "Master Group"
        imis_id = "IMIS Group No"
        imis_name = "IMIS Group"

    class CMList:
        cm_id = "Staff ID"
        cm_name = "CRT Name"
        team = "Team"
        supported_team_codes = "Support RM Segment"
        mass = "Mass"
        remark = "Remark"
        current_wip = "Current WIP"
        optimal_wip = "Optimal WIP"
        productivity = "Last 3 month Productivity"

    class HorisMG:
        cin = "CIN"
        mg_id = "Master_Group_ID"
        mg_name = "Master_Group_Name"
        customer_name = "Customer_Name"

    class IMIS:
        customer = "ITL_CUST_NUM"
        imis_id = "IMIS Group No"
        imis_name = "Group Name"

    class Closed:
        review_id = "Review ID"
        customer = "Customer Number"
        cm_id = "Latest DC Finalised by ID"
        first_claimed_id = "First Claimed ID"
        review_status = "Review Status"
        completed_status = "Approval Completed"
        cancelled_status = "Cancelled"
        date_candidates = ["Latest DC Finalised Date", "Approval/Cancel Date", "Initiated Date"]
        rm_num = "RM Num"
        segment = "Segment"

    class Open:
        review_id = "Review ID"
        customer = "Customer Number"
        stage = "Stage"
        task_status = "Task Status"
        assigned_to_user = "Assigned to User"
        user_id = "User ID"
        latest_dc_by = "Latest DC Finalised by ID"
        cm_id = "Staff ID"
        date_candidates = ["Initiated Date", "Latest DC Finalised Date", "Date of latest action", "Stage Claimed Date"]
        rm_num = "RM Num"
        segment = "Segment"

def read_json(json_file: str) -> dict:
    with open(json_file, "r", encoding="utf-8") as f:
        return json.load(f)


def clean_code(value, strip_zero: bool = False) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.lstrip("0") if strip_zero else text


def norm_text(value) -> str:
    return "" if pd.isna(value) else " ".join(str(value).upper().split())


def extract_staff_id(value) -> str:
    """Extract a bracketed staff ID without treating free text as an ID."""
    match = re.search(r"\[\s*([^\[\]]+?)\s*\]\s*$", clean_code(value))
    return clean_code(match.group(1)) if match else ""


def read_table(file_config, default_sheet=None, base_dir=None) -> pd.DataFrame:
    """Read one raw table, preserving identifiers and literal missing-value markers."""
    if isinstance(file_config, str):
        path = file_config
        sheet_name = default_sheet
        header = 0
    else:
        path = file_config["path"]
        sheet_name = file_config.get("sheet_name", default_sheet)
        header = file_config.get("header", 0)

    path = Path(path).expanduser()
    if not path.is_absolute():
        path = (Path(base_dir) if base_dir is not None else Path.cwd()) / path

    options = {"header": header, "dtype": str, "keep_default_na": False}
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(path, sheet_name=0 if sheet_name is None else sheet_name, **options)
    if suffix == ".csv":
        return pd.read_csv(path, encoding="utf-8-sig", **options)
    if suffix in {".txt", ".tsv"}:
        return pd.read_csv(path, sep="\t", encoding="utf-8-sig", **options)
    raise ValueError(f"Unsupported file type: {path}")


def read_many(file_config, fields_as_rows_markers=None, base_dir=None) -> pd.DataFrame:
    """Concatenate reports in config order, matching AMH orientation handling."""
    if file_config is None:
        files = []
    elif isinstance(file_config, (list, tuple)):
        files = file_config
    else:
        files = [file_config]

    frames = []
    for file in files:
        df = read_table(file, base_dir=base_dir)
        orientation = file.get("orientation", "records") if isinstance(file, dict) else "records"
        if orientation not in {"records", "fields_as_rows", "auto"}:
            raise ValueError(f"Unsupported table orientation: {orientation}")
        fields_as_rows = orientation == "fields_as_rows"
        if orientation == "auto" and fields_as_rows_markers and not df.empty and df.shape[1] > 1:
            first_column = set(df.iloc[:, 0].astype(str).str.strip())
            fields_as_rows = any(marker in first_column for marker in fields_as_rows_markers)
        if fields_as_rows:
            df = df.set_index(df.columns[0]).T
            df.columns = [str(column).strip() for column in df.columns]
            df = df.reset_index(drop=True)
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


class HaseAssignment:
    def __init__(self, config_file: str, current_date: str):
        self.config_file = str(Path(config_file).expanduser().resolve())
        self.config = read_json(self.config_file)
        self.as_of_date = pd.to_datetime(current_date)
        self.lookback_years = self.config.get("lookback_years", [2, 5])
        self.debug = self.config.get("debug", False)
        self.warnings = list()

        self.case_df = pd.DataFrame()
        self.cm_df = pd.DataFrame()
        self.closed_history_df = pd.DataFrame()
        self.open_history_df = pd.DataFrame()
        self.history_df = pd.DataFrame()
        self.proposed_df = pd.DataFrame()
        self.case_detail_df = pd.DataFrame()

    def read_input(self):
        """Load HASE raw inputs without cleaning, deduplication or allocation."""
        files = self.config["files"]
        base_dir = Path(self.config_file).expanduser().resolve().parent
        raw_case_df = read_table(files["business_output"], default_sheet="Case Assignment", base_dir=base_dir)
        raw_cm_df = read_table(files["cm_list"], default_sheet="OCM Segment Mapping", base_dir=base_dir)
        raw_closed_df = read_many(files["closed_report"], base_dir=base_dir)
        raw_open_df = read_many(
            files["open_report"],
            fields_as_rows_markers=["Customer Number", "Review ID", "Stage", "Assigned to User"],
            base_dir=base_dir,
        )
        self.raw_case_df = raw_case_df
        self.raw_cm_df = raw_cm_df
        self.raw_closed_df = raw_closed_df
        self.raw_open_df = raw_open_df

    def prepare_case_list(self):
        """Normalize HASE cases using the AMH case-list preparation structure."""
        fields = self.config.get("fields", {}).get("case_list", {})
        file_config = self.config["files"]["business_output"]
        orientation = file_config.get("orientation", "auto") if isinstance(file_config, dict) else "auto"
        if orientation not in {"records", "fields_as_rows", "auto"}:
            raise ValueError(f"Unsupported table orientation: {orientation}")
        df = self.raw_case_df.copy()
        df.columns = [str(column).strip() for column in df.columns]

        def column(key):
            return fields.get(key, getattr(DefaultFields.CaseList, key))

        case_id_col = column("case_id")
        if orientation == "fields_as_rows" or (
            orientation == "auto" and case_id_col not in df.columns
            and not df.empty and df.shape[1] > 1
            and case_id_col in df.iloc[:, 0].astype(str).str.strip().to_list()
        ):
            df = df.set_index(df.columns[0]).T
            df.columns = [str(value).strip() for value in df.columns]
            df = df.reset_index(drop=True)

        required = [column(key) for key in ("case_id", "customer", "segment", "business_cm_name")]
        missing = [name for name in required if name not in df.columns]
        if missing:
            raise ValueError(f"business_output.Case Assignment missing required columns: {missing}")

        def values(key):
            name = column(key)
            return df[name] if name in df.columns else pd.Series("", index=df.index, dtype=object)

        case_df = pd.DataFrame(index=df.index)
        case_df["case_row_id"] = range(1, len(df) + 1)
        case_df["case_id"] = values("case_id").apply(clean_code)
        # Match AMH/TaskClean join keys while keeping the display CIN's leading zeros.
        case_df["customer_id"] = values("customer").apply(lambda value: clean_code(value, strip_zero=True))
        case_df["cin"] = (values("cin") if column("cin") in df.columns else values("customer")).apply(clean_code)
        case_df["segment"] = values("segment").apply(clean_code)
        case_df["segment_norm"] = case_df["segment"].apply(norm_text)
        case_df["business_cm_name"] = values("business_cm_name").apply(clean_code)
        case_df["business_cm_id"] = values("business_cm_id").apply(clean_code)
        missing_staff = case_df["business_cm_id"].eq("")
        case_df.loc[missing_staff, "business_cm_id"] = case_df.loc[
            missing_staff, "business_cm_name"
        ].apply(extract_staff_id)
        case_df["business_team"] = values("business_team").apply(clean_code)
        case_df["assign_date"] = pd.to_datetime(values("assign_date"), errors="coerce", format="mixed")
        case_df["designated_ocm"] = values("designated_ocm").apply(clean_code)
        case_df["designated_ocm_id"] = case_df["designated_ocm"].apply(extract_staff_id)
        case_df["trigger_type"] = values("trigger_type").apply(clean_code)
        # HASE has no confirmed due-date rule; Assign Date is not a due date.
        case_df["due_date"] = pd.to_datetime(values("due_date"), errors="coerce", format="mixed")
        default_workload = self.config.get("default_workload", 1.0)
        case_df["workload"] = pd.to_numeric(values("workload"), errors="coerce").fillna(default_workload)

        duplicated_customer = case_df["customer_id"].duplicated().sum()
        if duplicated_customer:
            self.warnings.append(
                f"There are {duplicated_customer} duplicated customer_id rows. "
                "Proposed algorithm is keyed by customer_id, so duplicated rows may share the same proposed CM."
            )
        self.case_df = case_df

    def prepare_cm_list(self):
        """Normalize the HASE CM roster using configured WIP and capacity values."""
        fields = self.config.get("fields", {}).get("cm_list", {})
        df = self.raw_cm_df.copy()
        df.columns = [str(column).strip() for column in df.columns]

        def column(key):
            return fields.get(key, getattr(DefaultFields.CMList, key))

        required = [column(key) for key in ("cm_id", "cm_name", "team", "supported_team_codes")]
        missing = [name for name in required if name not in df.columns]
        if missing:
            raise ValueError(f"cm_list.OCM Segment Mapping missing required columns: {missing}")

        def values(key):
            name = column(key)
            return df[name] if name in df.columns else pd.Series("", index=df.index, dtype=object)

        def team_codes(value):
            # Split only the roster's comma-separated codes, retaining display spelling.
            return list(dict.fromkeys(code.strip() for code in clean_code(value).split(",") if code.strip()))

        cm_df = pd.DataFrame(index=df.index)
        cm_df["cm_id"] = values("cm_id").apply(clean_code)
        cm_df["cm_name"] = values("cm_name").apply(clean_code)
        cm_df["team"] = values("team").apply(clean_code)
        cm_df["supported_team_codes"] = values("supported_team_codes").apply(team_codes)
        cm_df["supported_team_codes_norm"] = cm_df["supported_team_codes"].apply(
            lambda codes: list(dict.fromkeys(norm_text(code) for code in codes))
        )
        cm_df["mass"] = values("mass").apply(norm_text)
        cm_df["remark"] = values("remark").apply(clean_code)

        defaults = self.config.get("cm_defaults", {})
        for key in ("current_wip", "optimal_wip", "productivity"):
            value = defaults.get(key)
            if isinstance(value, bool):
                raise ValueError(f"cm_defaults.{key} must be a finite non-negative number")
            try:
                value = float(value)
            except (TypeError, ValueError):
                raise ValueError(f"cm_defaults.{key} must be a finite non-negative number") from None
            if not np.isfinite(value) or value < 0:
                raise ValueError(f"cm_defaults.{key} must be a finite non-negative number")
            cm_df[key] = value

        cm_df = cm_df[cm_df["cm_id"] != ""].drop_duplicates("cm_id", keep="last").reset_index(drop=True)
        self.cm_df = cm_df

    def prepare_history(self) -> pd.DataFrame:
        """Combine HASE history using AMH's CM/date filters, without group keys."""
        closed_df = self.prepare_closed_history()
        open_df = self.prepare_open_history()
        history_df = pd.concat([closed_df, open_df], ignore_index=True)
        history_df = history_df[
            history_df["cm_id"].notna()
            & history_df["cm_id"].ne("")
            & history_df["history_date"].notna()
            & history_df["history_date"].le(self.as_of_date)
        ].copy()
        self.history_df = history_df.reset_index(drop=True)
        return self.history_df

    def prepare_closed_history(self) -> pd.DataFrame:
        """Normalize closed reviews without the AMH Master Group/IMIS mappings."""
        fields = self.config.get("fields", {}).get("closed_report", {})
        df = self.raw_closed_df.copy()
        df.columns = [str(column).strip() for column in df.columns]

        def column(key):
            return fields.get(key, getattr(DefaultFields.Closed, key))

        required = [column("customer"), column("cm_id")]
        if self.config.get("closed_completed_only", True):
            required.append(column("review_status"))
        missing = [name for name in required if name not in df.columns]
        if missing:
            raise ValueError(f"closed_report missing required columns: {missing}")

        if self.config.get("closed_completed_only", True):
            df = df[df[column("review_status")].eq(column("completed_status"))].copy()

        def values(key):
            name = column(key)
            return df[name] if name in df.columns else pd.Series("", index=df.index, dtype=object)

        closed = pd.DataFrame(index=df.index)
        closed["source"] = "closed"
        closed["review_id"] = values("review_id").apply(clean_code)
        closed["customer_id"] = values("customer").apply(lambda value: clean_code(value, strip_zero=True))
        closed["cin"] = values("customer").apply(clean_code)
        closed["cm_id"] = values("cm_id").apply(clean_code)
        cancelled = values("review_status").eq(column("cancelled_status"))
        fallback = cancelled & closed["cm_id"].eq("")
        closed.loc[fallback, "cm_id"] = values("first_claimed_id").loc[fallback].apply(clean_code)

        # Raw readers preserve empty strings; parse each candidate before combining
        # so blanks/invalid dates do not prevent fallback to the next date column.
        history_date = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
        for name in fields.get("date_candidates", DefaultFields.Closed.date_candidates):
            if name in df.columns:
                parsed = pd.to_datetime(df[name], errors="coerce", format="mixed")
                history_date = history_date.fillna(parsed)
        closed["history_date"] = history_date
        closed["rm_num"] = values("rm_num").apply(clean_code)
        # HASE report Segment values represent Team code.
        closed["segment"] = values("segment").apply(clean_code)
        closed["segment_norm"] = closed["segment"].apply(norm_text)
        self.closed_history_df = closed.reset_index(drop=True)
        return self.closed_history_df

    def prepare_open_history(self) -> pd.DataFrame:
        """Normalize WIP history, attributing QC/APP reviews to the DC CM."""
        fields = self.config.get("fields", {}).get("open_report", {})
        df = self.raw_open_df.copy()
        df.columns = [str(column).strip() for column in df.columns]

        def column(key):
            return fields.get(key, getattr(DefaultFields.Open, key))

        customer_col = column("customer")
        if customer_col not in df.columns:
            raise ValueError(f"open_report missing required columns: {[customer_col]}")

        def values(key):
            name = column(key)
            return df[name] if name in df.columns else pd.Series("", index=df.index, dtype=object)

        open_df = pd.DataFrame(index=df.index)
        open_df["source"] = "open"
        open_df["review_id"] = values("review_id").apply(clean_code)
        open_df["customer_id"] = values("customer").apply(lambda value: clean_code(value, strip_zero=True))
        open_df["cin"] = values("customer").apply(clean_code)
        open_df["stage"] = values("stage").apply(norm_text)
        open_df["task_status"] = values("task_status").apply(clean_code)
        open_df["assigned_to_user"] = values("assigned_to_user").apply(clean_code)
        open_df["user_id"] = values("user_id").apply(clean_code)
        open_df["latest_dc_by"] = values("latest_dc_by").apply(clean_code)
        open_df["cm_id"] = self.derive_open_cm_id(
            df,
            stage_col=column("stage"),
            assigned_to_user_col=column("assigned_to_user"),
            user_id_col=column("user_id"),
            latest_dc_by_col=column("latest_dc_by"),
            cm_col=column("cm_id"),
        )
        history_date = pd.Series(pd.NaT, index=df.index, dtype="datetime64[ns]")
        for name in fields.get("date_candidates", DefaultFields.Open.date_candidates):
            if name in df.columns:
                parsed = pd.to_datetime(df[name], errors="coerce", format="mixed")
                history_date = history_date.fillna(parsed)
        open_df["history_date"] = history_date
        open_df["rm_num"] = values("rm_num").apply(clean_code)
        open_df["segment"] = values("segment").apply(clean_code)
        open_df["segment_norm"] = open_df["segment"].apply(norm_text)
        self.open_history_df = open_df.reset_index(drop=True)
        return self.open_history_df

    def build_algorithm_analyst_df(self) -> pd.DataFrame:
        return pd.DataFrame({
            ah.analyst_id: self.cm_df["cm_id"],
            ah.current_wip: self.cm_df["current_wip"],
            ah.target_wip: self.cm_df["optimal_wip"],
            ah.daily_productivity: self.cm_df["productivity"],
            ah.team_head: self.cm_df["team"],
            ah.analyst_name: self.cm_df["cm_name"],
        })

    def build_algorithm_task_df(self) -> pd.DataFrame:
        # Planned T0 is the initiation date. CDDTask derives initiation as due - 120.
        planned_t0 = self.case_df["due_date"].fillna(self.as_of_date)
        return pd.DataFrame({
            th.task_id: self.case_df["customer_id"],
            th.due_date: planned_t0 + pd.Timedelta(days=120),
            th.effort: self.case_df["workload"],
            th.market: self.config.get("market", ""),
            th.legal_entity: self.config.get("legal_entity", ""),
        })

    def build_algorithm_closed_df(self) -> pd.DataFrame:
        closed = self.history_df[self.history_df["source"].eq("closed")]
        return pd.DataFrame({
            ch.ReviewID: closed["review_id"],
            ch.CustomerNumber: closed["customer_id"],
            ch.ReviewStatus: "Approval Completed",
            ch.LatestDcFinalisedById: closed["cm_id"],
            ch.RmNum: closed["rm_num"],
            ch.InitiatedDate: closed["history_date"],
            ch.Segment: "",
            # The shared allocator requires this column even when groups are unused.
            mh.MasterGroupCode: "",
        })

    def build_algorithm_open_df(self) -> pd.DataFrame:
        opened = self.history_df[self.history_df["source"].eq("open")]
        return pd.DataFrame({
            oh.ReviewID: opened["review_id"],
            oh.CustomerNumber: opened["customer_id"],
            oh.StaffID: opened["cm_id"],
            oh.RmNum: opened["rm_num"],
            oh.InitiatedDate: opened["history_date"],
            oh.Segment: "",
            mh.MasterGroupCode: "",
        })

    def run_proposed_assignment(self, output_result=True) -> pd.DataFrame:
        """Run the shared allocator with Team code experience from the CM roster only."""
        df_analysts = self.build_algorithm_analyst_df()
        df_tasks = self.build_algorithm_task_df()
        df_close = self.build_algorithm_closed_df()
        df_open = self.build_algorithm_open_df()
        allocator = CDDTaskAllocator(
            df_analysts=df_analysts,
            df_tasks=df_tasks,
            df_close=df_close,
            df_open=df_open,
            current_date=str(self.as_of_date.date()),
            config=self.config.get("algorithm_config", {}),
        )
        self.allocator = allocator
        allocator.read_data()
        allocator.add_history_customer_to_analyst()

        # Replace any history-derived segment state, including blank segment keys.
        for row in self.cm_df.itertuples(index=False):
            analyst = allocator.analysts[row.cm_id]
            analyst.history_segment = {norm_text(code) for code in row.supported_team_codes if norm_text(code)}
            analyst.history_segment_reviews.clear()
        for row in self.case_df.itertuples(index=False):
            allocator.tasks[row.customer_id].segment = norm_text(row.segment)

        result = allocator.run_top_n_analyst(n=self.config.get("top_n", 3))
        self.algorithm_result_df = result
        if result.empty:
            self.proposed_df = pd.DataFrame(columns=["customer_id", "proposed_cm_id", "proposed_score"])
            self.warnings.append("Proposed algorithm returns empty result.")
        else:
            self.proposed_df = pd.DataFrame({
                "customer_id": result["Task ID"].apply(lambda value: clean_code(value, strip_zero=True)),
                "proposed_cm_id": result["Assigned Analyst ID 1"].apply(clean_code),
                "proposed_score": result["Score Total 1"],
            })
        if output_result:
            output = Path(self.config.get("output_file", "hase_assignment_result.xlsx")).expanduser()
            output = output.with_name(f"{output.stem}_{self.as_of_date:%Y-%m-%d}{output.suffix}")
            if not output.is_absolute():
                output = Path(self.config_file).parent / output
            output.parent.mkdir(parents=True, exist_ok=True)
            result.to_excel(output, index=False)
            self.export_assignment_workbook()
        return result

    def export_assignment_workbook(self) -> Path:
        """Copy the complete business workbook and patch only assignment cells."""
        file_config = self.config["files"]["business_output"]
        if isinstance(file_config, str):
            file_config = {"path": file_config}
        if file_config.get("orientation", "records") not in {"records", "auto"}:
            raise ValueError("Assignment workbook export requires records orientation")
        base_dir = Path(self.config_file).parent
        source = Path(file_config["path"]).expanduser()
        if not source.is_absolute():
            source = base_dir / source
        if source.suffix.lower() not in {".xlsx", ".xlsm"}:
            raise ValueError("Assignment workbook export requires an .xlsx or .xlsm input")
        output = Path(self.config.get("filled_output_file", f"hase_business_output_assigned{source.suffix}")).expanduser()
        output = output.with_name(f"{output.stem}_{self.as_of_date:%Y-%m-%d}{output.suffix}")
        if not output.is_absolute():
            output = base_dir / output
        if output.suffix.lower() != source.suffix.lower():
            raise ValueError("Filled output must use the same Excel extension as business_output")
        if output.resolve() == source.resolve():
            raise ValueError("Filled output cannot overwrite business_output")

        fields = self.config.get("fields", {}).get("case_list", {})
        header = file_config.get("header", 0)
        if not isinstance(header, int) or header < 0:
            raise ValueError("Assignment workbook export requires a non-negative integer header")
        sheet_name = file_config.get("sheet_name", "Case Assignment")
        # Re-read source rows so matching stays keyed by CIN, never allocator sort order.
        source_rows = pd.read_excel(source, sheet_name=sheet_name, header=header, dtype=str, keep_default_na=False)
        source_rows.columns = [str(name).strip() for name in source_rows.columns]
        customer_column = fields.get("customer", DefaultFields.CaseList.customer)
        targets = [fields.get("business_team", DefaultFields.CaseList.business_team),
                   fields.get("business_cm_name", DefaultFields.CaseList.business_cm_name), "CM 1", "CM 2", "CM 3"]
        missing = [name for name in [customer_column, *targets] if name not in source_rows.columns]
        if missing:
            raise ValueError(f"business_output missing assignment output columns: {missing}")
        results = self.algorithm_result_df.copy()
        if not results.empty:
            results["customer_id"] = results["Task ID"].apply(lambda value: clean_code(value, strip_zero=True))
            if results["customer_id"].duplicated().any():
                raise ValueError("Assignment result contains duplicated customer IDs")
        by_customer = {} if results.empty else results.set_index("customer_id").to_dict("index")
        cms = self.cm_df.set_index("cm_id").to_dict("index")

        def label(cm_id):
            if not cm_id:
                return ""
            if cm_id not in cms:
                raise ValueError(f"Assigned CM is absent from cm_list: {cm_id}")
            name = cms[cm_id]["cm_name"]
            return name if extract_staff_id(name) == cm_id else f"{name} [{cm_id}]"

        def column_letter(index):
            value = ""
            while index:
                index, remainder = divmod(index - 1, 26)
                value = chr(65 + remainder) + value
            return value

        patches = {}
        for position, (_, row) in enumerate(source_rows.iterrows()):
            customer = clean_code(row[customer_column], strip_zero=True)
            if customer not in by_customer:
                continue
            result = by_customer[customer]
            ids = [clean_code(result.get(f"Assigned Analyst ID {rank}", "")) for rank in (1, 2, 3)]
            if not ids[0]:
                continue
            labels = [label(cm_id) for cm_id in ids]
            values = [cms[ids[0]]["team"], labels[0], *labels]
            excel_row = header + position + 2
            patches[excel_row] = {
                f"{column_letter(source_rows.columns.get_loc(name) + 1)}{excel_row}": value
                for name, value in zip(targets, values)
            }

        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        with ZipFile(source) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets = workbook.findall("s:sheets/s:sheet", ns)
            sheet = sheets[sheet_name] if isinstance(sheet_name, int) else next(s for s in sheets if s.attrib["name"] == sheet_name)
            relation = sheet.attrib[f"{{{rel_ns}}}id"]
            relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            target = next(r.attrib["Target"] for r in relationships if r.attrib["Id"] == relation)
            sheet_path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            xml = archive.read(sheet_path).decode("utf-8")
            prefix = re.search(r'<((?:[\w.-]+:)?)worksheet\b', xml).group(1)
            tag = re.escape(prefix)

            def patch_row(match):
                row_xml = match.group(0)
                row_number = int(re.search(r'\br="(\d+)"', row_xml).group(1))
                for address, value in patches.pop(row_number, {}).items():
                    cell_pattern = rf'<{tag}c\b[^>]*\br="{address}"[^>]*?(?:/>|>.*?</{tag}c>)'
                    old = re.search(cell_pattern, row_xml, flags=re.DOTALL)
                    attributes = re.match(rf'<{tag}c\b([^>]*?)(?:/?>)', old.group(0)).group(1) if old else f' r="{address}"'
                    attributes = re.sub(r'\s+t="[^"]*"', '', attributes)
                    new = f'<{prefix}c{attributes} t="inlineStr"><{prefix}is><{prefix}t xml:space="preserve">{escape(str(value))}</{prefix}t></{prefix}is></{prefix}c>'
                    if old:
                        row_xml = row_xml[:old.start()] + new + row_xml[old.end():]
                    else:
                        # Insert missing cells in column order, leaving all other XML intact.
                        following = next((c for c in re.finditer(rf'<{tag}c\b[^>]*\br="([A-Z]+)\d+"', row_xml)
                                          if (len(c.group(1)), c.group(1)) > (len(re.sub(r'\d', '', address)), re.sub(r'\d', '', address))), None)
                        offset = following.start() if following else row_xml.rfind(f"</{prefix}row>")
                        row_xml = row_xml[:offset] + new + row_xml[offset:]
                return row_xml

            xml = re.sub(rf'<{tag}row\b[^>]*>.*?</{tag}row>', patch_row, xml, flags=re.DOTALL)
            if patches:
                raise ValueError(f"Could not locate worksheet rows for assignment: {sorted(patches)}")
            output.parent.mkdir(parents=True, exist_ok=True)
            with ZipFile(output, "w") as destination:
                destination.comment = archive.comment
                for entry in archive.infolist():
                    destination.writestr(entry, xml.encode("utf-8") if entry.filename == sheet_path else archive.read(entry.filename))
        self.filled_output_path = output
        return output

    def derive_open_cm_id(
        self, df: pd.DataFrame, stage_col: str, assigned_to_user_col: str,
        user_id_col: str, latest_dc_by_col: str, cm_col: str,
    ) -> pd.Series:
        """Use AMH's CM priority, allowing optional source columns to be absent."""
        cm_id = pd.Series("", index=df.index, dtype=object)
        for index, row in df.iterrows():
            if norm_text(row.get(stage_col, "")) in (st.QC, st.APP):
                # QC/APP assignees are reviewers/approvers, not the handling DC CM.
                cm_id.loc[index] = clean_code(row.get(latest_dc_by_col, ""))
                continue
            assigned = clean_code(row.get(assigned_to_user_col, ""))
            user = clean_code(row.get(user_id_col, ""))
            cm_id.loc[index] = (
                assigned if assigned.isdigit() else user if user.isdigit()
                else clean_code(row.get(cm_col, ""))
            )
        return cm_id


# Retain compatibility with the original placeholder class name.
if __name__ == "__main__":
    from datetime import datetime
    current_date = datetime.now().strftime("%Y-%m-%d")
    self = HaseAssignment(config_file=str(Path(__file__).parent / "hase_data" / "config.json"), current_date=current_date)
    self.read_input()
    self.prepare_case_list()
    self.prepare_cm_list()
    self.prepare_history()
    self.run_proposed_assignment()
    self.export_assignment_workbook()
