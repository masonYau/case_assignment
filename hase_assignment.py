import argparse
import json
import logging
import re
import sys
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
        latest_dc_by = "DC Finalized by ID"
        cm_id = "Staff ID"
        date_candidates = ["Initiated Date", "DC Finalized Date", "Date of latest action", "Stage Start Date"]
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
        """Normalize the HASE CM roster without inferring capacity or availability."""
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

        # Unlike AMH, the HASE source need not contain capacity metrics. Leave unknown
        # values as NaN so later WIP/capacity preparation can supply actual values.
        for key in ("current_wip", "optimal_wip", "productivity"):
            cm_df[key] = pd.to_numeric(values(key), errors="coerce")

        cm_df = cm_df[cm_df["cm_id"] != ""].drop_duplicates("cm_id", keep="last").reset_index(drop=True)
        for key in ("current_wip", "optimal_wip", "productivity"):
            unknown = int(cm_df[key].isna().sum())
            if unknown:
                message = (
                    f"cm_list: {unknown} CM rows have missing or invalid {key}; "
                    "values remain unknown until WIP/capacity preparation."
                )
                if message not in self.warnings:
                    self.warnings.append(message)
        self.cm_df = cm_df


# Retain compatibility with the original placeholder class name.
if __name__ == "__main__":
    from datetime import datetime
    current_date = datetime.now().strftime("%Y-%m-%d")
    self = HaseAssignment(config_file=str(Path(__file__).parent / "hase_data" / "config.json"), current_date=current_date)
    self.read_input()
    self.prepare_case_list()
    self.prepare_cm_list()
