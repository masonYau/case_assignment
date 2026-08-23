# experiment/run_experiment.py

```python
import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Optional, Sequence, Tuple

import numpy as np
import pandas as pd

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from case_assignment import CDDTaskAllocator
from header import AnalystHeader as ah
from header import ClosedReviewHeader as ch
from header import MasterGroupHeader as mh
from header import OpenReviewHeader as oh
from header import StageType as st
from header import TaskHeader as th


class DefaultFields:
    class CaseList:
        case_id = "Customer_ID_S1"
        customer = "CIN"
        cin = "CIN"
        segment = "CMB_Segment_S1"
        due_date = "Planned T0 date"
        workload = "Workload Units"
        business_cm_id = "Case_Manager_Staff_ID_S1"
        business_cm_name = "CM Name"
        business_team = "Assign by TH"
        mg_id = "Master Group ID"
        mg_name = "Master Group"
        imis_id = "IMIS Group No"
        imis_name = "IMIS Group"

    class CMList:
        cm_id = "Staff ID"
        cm_name = "CM Name"
        team = "Team"
        current_wip = "Current WIP"
        optimal_wip = "Optimal WIP"
        productivity = "Last 3 month Productivity"

    class HorisMG:
        cin = "CIN"
        mg_id = "Master_Group_ID"
        mg_name = "Master_Group_Name"
        customer_name = "Customer_Name"

    class IMIS:
        customer = "Real customer id"
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


def resolve_working_path(path_value: str) -> Path:
    path = Path(path_value).expanduser()
    if path.is_absolute():
        return path
    return Path.cwd() / path


def get_nested(cfg: dict, path: Sequence[str], default=None):
    cur = cfg
    for key in path:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def as_list(value) -> list:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    return [value]


def clean_code(value, strip_zero: bool = False) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    if strip_zero:
        text = text.lstrip("0")
    return text


def clean_customer(value) -> str:
    return clean_code(value, strip_zero=True)


def clean_staff(value) -> str:
    return clean_code(value, strip_zero=False)


def clean_group(value) -> str:
    text = clean_code(value, strip_zero=True)
    if text in {"0", "0.0", "-", "NA", "N/A", "NONE", "NAN"}:
        return ""
    return text


def norm_text(value) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def coalesce_columns(df: pd.DataFrame, columns: Sequence[str]) -> pd.Series:
    result = pd.Series(pd.NA, index=df.index)
    for col in columns:
        if col in df.columns:
            result = result.fillna(df[col])
    return result


def get_col(fields: dict, key: str, default: str) -> str:
    return fields.get(key, default)


def normalize_fields_as_rows(df: pd.DataFrame) -> pd.DataFrame:
    field_col = df.columns[0]
    out = df.set_index(field_col).T
    out.columns = [str(c).strip() for c in out.columns]
    out = out.reset_index(drop=True)
    return out


def looks_like_fields_as_rows(df: pd.DataFrame, field_names: Sequence[str]) -> bool:
    if df.empty or df.shape[1] <= 1:
        return False
    first_col_values = set(df.iloc[:, 0].dropna().astype(str).str.strip())
    return any(field_name in first_col_values for field_name in field_names)


def read_table(file_config, default_sheet=None) -> pd.DataFrame:
    if isinstance(file_config, str):
        path = file_config
        sheet_name = default_sheet
        header = 0
    else:
        path = file_config["path"]
        sheet_name = file_config.get("sheet_name", default_sheet)
        header = file_config.get("header", 0)

    path = resolve_working_path(path)
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(path, sheet_name=sheet_name or 0, header=header)
    if suffix == ".csv":
        return pd.read_csv(path, header=header)
    if suffix in {".txt", ".tsv"}:
        return pd.read_csv(path, sep="\t", header=header)
    raise ValueError(f"Unsupported file type: {path}")


def read_many(file_config, fields_as_rows_markers: Optional[Sequence[str]] = None) -> pd.DataFrame:
    files = as_list(file_config)
    frames = []
    for f in files:
        df = read_table(f)
        orientation = f.get("orientation", "records") if isinstance(f, dict) else "records"
        if orientation == "fields_as_rows" or (
                orientation == "auto" and fields_as_rows_markers and looks_like_fields_as_rows(df, fields_as_rows_markers)
        ):
            df = normalize_fields_as_rows(df)
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


class AssignmentExperiment:
    def __init__(self, config_file: str):
        self.config_file = config_file
        self.config = read_json(config_file)
        self.as_of_date = pd.to_datetime(self.config["as_of_date"])
        self.lookback_years = self.config.get("lookback_years", [2, 5])
        self.debug = self.config.get("debug", False)
        self.warnings = list()

        self.case_df = pd.DataFrame()
        self.cm_df = pd.DataFrame()
        self.horis_mg_df = pd.DataFrame()
        self.imis_df = pd.DataFrame()
        self.closed_history_df = pd.DataFrame()
        self.open_history_df = pd.DataFrame()
        self.history_df = pd.DataFrame()
        self.proposed_df = pd.DataFrame()
        self.case_detail_df = pd.DataFrame()

    def run(self):
        logging.info("Start assignment comparison experiment")
        self.read_input()
        self.prepare_case_list()
        self.prepare_cm_list()
        self.prepare_group_mapping()
        self.prepare_history()
        self.run_proposed_assignment()
        self.build_case_detail()
        summary_df = self.build_summary()
        wip_by_cm_df, wip_by_team_cm_df, wip_by_team_df = self.build_wip_summary()
        mapping_check_df = self.build_mapping_check()
        self.output_result(
            summary_df=summary_df,
            wip_by_cm_df=wip_by_cm_df,
            wip_by_team_cm_df=wip_by_team_cm_df,
            wip_by_team_df=wip_by_team_df,
            mapping_check_df=mapping_check_df,
        )
        logging.info("Assignment comparison experiment finished")

    def read_input(self):
        files = self.config["files"]
        self.raw_case_df = read_table(files["business_output"], default_sheet="CMAssignmentFullList")
        self.raw_cm_df = read_table(files["cm_list"])
        self.raw_closed_df = read_many(files["closed_report"])
        self.raw_open_df = read_many(
            files["open_report"],
            fields_as_rows_markers=["Customer Number", "Review ID", "Stage", "Assigned to User"],
        )
        self.raw_horis_mg_df = read_table(files["horis_mg"])
        self.raw_imis_df = read_table(files["bb_rm_imis_group"])

    def prepare_case_list(self):
        fields = get_nested(self.config, ["fields", "case_list"], {})
        orientation = get_nested(self.config, ["files", "business_output", "orientation"], "auto")
        df = self.raw_case_df.copy()

        case_id_col = get_col(fields, "case_id", DefaultFields.CaseList.case_id)
        if orientation == "fields_as_rows" or (
                orientation == "auto" and case_id_col not in df.columns and df.shape[1] > 1
                and case_id_col in df.iloc[:, 0].astype(str).to_list()
        ):
            df = normalize_fields_as_rows(df)

        customer_col = get_col(fields, "customer", DefaultFields.CaseList.customer)
        cin_col = get_col(fields, "cin", DefaultFields.CaseList.cin)
        segment_col = get_col(fields, "segment", DefaultFields.CaseList.segment)
        due_date_col = get_col(fields, "due_date", DefaultFields.CaseList.due_date)
        workload_col = get_col(fields, "workload", DefaultFields.CaseList.workload)
        business_cm_id_col = get_col(fields, "business_cm_id", DefaultFields.CaseList.business_cm_id)
        business_cm_name_col = get_col(fields, "business_cm_name", DefaultFields.CaseList.business_cm_name)
        business_team_col = get_col(fields, "business_team", DefaultFields.CaseList.business_team)
        mg_id_col = get_col(fields, "mg_id", DefaultFields.CaseList.mg_id)
        mg_name_col = get_col(fields, "mg_name", DefaultFields.CaseList.mg_name)
        imis_id_col = get_col(fields, "imis_id", DefaultFields.CaseList.imis_id)
        imis_name_col = get_col(fields, "imis_name", DefaultFields.CaseList.imis_name)

        self.check_columns(
            df,
            [case_id_col, customer_col, business_cm_id_col],
            "business_output.CMAssignmentFullList",
        )

        case_df = pd.DataFrame(index=df.index)
        case_df["case_row_id"] = range(1, len(df) + 1)
        case_df["case_id"] = df[case_id_col].apply(clean_code)
        case_df["customer_id"] = df[customer_col].apply(clean_customer)
        case_df["cin"] = df[cin_col].apply(clean_customer) if cin_col in df.columns else case_df["customer_id"]
        case_df["segment"] = df[segment_col].fillna("").astype(str) if segment_col in df.columns else ""
        case_df["segment_norm"] = case_df["segment"].apply(norm_text)
        case_df["business_cm_id"] = df[business_cm_id_col].apply(clean_staff)
        case_df["business_cm_name"] = df[business_cm_name_col].fillna("").astype(str) if business_cm_name_col in df.columns else ""
        case_df["business_team"] = df[business_team_col].fillna("").astype(str) if business_team_col in df.columns else ""
        case_df["business_mg_id"] = df[mg_id_col].apply(clean_group) if mg_id_col in df.columns else ""
        case_df["business_mg_name"] = df[mg_name_col].fillna("").astype(str) if mg_name_col in df.columns else ""
        case_df["business_imis_id"] = df[imis_id_col].apply(clean_group) if imis_id_col in df.columns else ""
        case_df["business_imis_name"] = df[imis_name_col].fillna("").astype(str) if imis_name_col in df.columns else ""

        if due_date_col in df.columns:
            case_df["due_date"] = pd.to_datetime(df[due_date_col], errors="coerce", format="mixed")
        else:
            case_df["due_date"] = self.as_of_date + pd.Timedelta(days=120)
            self.warnings.append(f"Case due date column not found: {due_date_col}. Use as_of_date + 120 days.")

        default_workload = self.config.get("default_workload", 1.0)
        if workload_col in df.columns:
            case_df["workload"] = pd.to_numeric(df[workload_col], errors="coerce").fillna(default_workload)
        else:
            case_df["workload"] = default_workload

        duplicated_customer = case_df["customer_id"].duplicated().sum()
        if duplicated_customer:
            self.warnings.append(
                f"There are {duplicated_customer} duplicated customer_id rows. "
                "Proposed algorithm is keyed by customer_id, so duplicated rows may share the same proposed CM."
            )

        self.case_df = case_df

    def prepare_cm_list(self):
        fields = get_nested(self.config, ["fields", "cm_list"], {})
        df = self.raw_cm_df.copy()

        cm_id_col = get_col(fields, "cm_id", DefaultFields.CMList.cm_id)
        cm_name_col = get_col(fields, "cm_name", DefaultFields.CMList.cm_name)
        team_col = get_col(fields, "team", DefaultFields.CMList.team)
        current_wip_col = get_col(fields, "current_wip", DefaultFields.CMList.current_wip)
        optimal_wip_col = get_col(fields, "optimal_wip", DefaultFields.CMList.optimal_wip)
        productivity_col = get_col(fields, "productivity", DefaultFields.CMList.productivity)

        self.check_columns(df, [cm_id_col, current_wip_col, optimal_wip_col], "cm_list")

        cm_df = pd.DataFrame(index=df.index)
        cm_df["cm_id"] = df[cm_id_col].apply(clean_staff)
        cm_df["cm_name"] = df[cm_name_col].fillna("").astype(str) if cm_name_col in df.columns else ""
        cm_df["team"] = df[team_col].fillna("").astype(str) if team_col in df.columns else ""
        cm_df["current_wip"] = pd.to_numeric(df[current_wip_col], errors="coerce").fillna(0)
        cm_df["optimal_wip"] = pd.to_numeric(df[optimal_wip_col], errors="coerce").fillna(0)
        if productivity_col in df.columns:
            cm_df["productivity"] = pd.to_numeric(df[productivity_col], errors="coerce").fillna(1)
        else:
            cm_df["productivity"] = 1

        cm_df = cm_df[cm_df["cm_id"] != ""].drop_duplicates("cm_id", keep="last")
        self.cm_df = cm_df

    def prepare_group_mapping(self):
        mg_fields = get_nested(self.config, ["fields", "horis_mg"], {})
        imis_fields = get_nested(self.config, ["fields", "bb_rm_imis_group"], {})

        mg_df = self.raw_horis_mg_df.copy()
        cin_col = get_col(mg_fields, "cin", DefaultFields.HorisMG.cin)
        mg_id_col = get_col(mg_fields, "mg_id", DefaultFields.HorisMG.mg_id)
        mg_name_col = get_col(mg_fields, "mg_name", DefaultFields.HorisMG.mg_name)
        self.check_columns(mg_df, [cin_col, mg_id_col], "horis_mg")

        mg_map = pd.DataFrame({
            "cin": mg_df[cin_col].apply(clean_customer),
            "mapped_mg_id": mg_df[mg_id_col].apply(clean_group),
            "mapped_mg_name": mg_df[mg_name_col].fillna("").astype(str) if mg_name_col in mg_df.columns else "",
        })
        exclude_mg_ids = set(str(x) for x in self.config.get("exclude_mg_ids", [1514726659, 1514726667]))
        mg_map = mg_map[~mg_map["mapped_mg_id"].isin(exclude_mg_ids)]
        mg_map = mg_map[mg_map["cin"] != ""].drop_duplicates("cin", keep="last")
        self.horis_mg_df = mg_map

        imis_df = self.raw_imis_df.copy()
        imis_customer_col = get_col(imis_fields, "customer", DefaultFields.IMIS.customer)
        imis_id_col = get_col(imis_fields, "imis_id", DefaultFields.IMIS.imis_id)
        imis_name_col = get_col(imis_fields, "imis_name", DefaultFields.IMIS.imis_name)
        self.check_columns(imis_df, [imis_customer_col, imis_id_col], "bb_rm_imis_group")

        imis_map = pd.DataFrame({
            "imis_customer_key": imis_df[imis_customer_col].apply(clean_code),
            "mapped_imis_id": imis_df[imis_id_col].apply(clean_group),
            "mapped_imis_name": imis_df[imis_name_col].fillna("").astype(str) if imis_name_col in imis_df.columns else "",
        })
        imis_map = imis_map[imis_map["imis_customer_key"] != ""].drop_duplicates("imis_customer_key", keep="last")
        self.imis_df = imis_map

        case_df = self.case_df.merge(self.horis_mg_df, how="left", on="cin")
        imis_key_col = get_nested(
            self.config,
            ["fields", "case_list", "imis_mapping_key"],
            get_nested(self.config, ["fields", "case_list", "case_id"], DefaultFields.CaseList.case_id),
        )
        if imis_key_col in self.raw_case_df.columns:
            case_df["imis_customer_key"] = self.raw_case_df[imis_key_col].apply(clean_code).values
        else:
            case_df["imis_customer_key"] = case_df["case_id"]
        case_df = case_df.merge(self.imis_df, how="left", on="imis_customer_key")

        case_df["mg_id"] = case_df["mapped_mg_id"].fillna("")
        case_df.loc[case_df["mg_id"].eq(""), "mg_id"] = case_df.loc[case_df["mg_id"].eq(""), "business_mg_id"]
        case_df["mg_name"] = case_df["mapped_mg_name"].fillna("")
        case_df.loc[case_df["mg_name"].eq(""), "mg_name"] = case_df.loc[case_df["mg_name"].eq(""), "business_mg_name"]

        case_df["imis_id"] = case_df["mapped_imis_id"].fillna("")
        case_df.loc[case_df["imis_id"].eq(""), "imis_id"] = case_df.loc[case_df["imis_id"].eq(""), "business_imis_id"]
        case_df["imis_name"] = case_df["mapped_imis_name"].fillna("")
        case_df.loc[case_df["imis_name"].eq(""), "imis_name"] = case_df.loc[
            case_df["imis_name"].eq(""), "business_imis_name"
        ]

        self.case_df = self.add_group_key(case_df)

    def add_group_key(self, df: pd.DataFrame) -> pd.DataFrame:
        use_imis_for_bbrm = self.config.get("use_imis_as_group_for_bbrm", True)
        is_bbrm = df["segment_norm"].eq("BBRM")
        df["group_type"] = ""
        df["group_id"] = ""
        df.loc[df["mg_id"].ne(""), "group_type"] = "MG"
        df.loc[df["mg_id"].ne(""), "group_id"] = df.loc[df["mg_id"].ne(""), "mg_id"]
        if use_imis_for_bbrm:
            sel = is_bbrm & df["imis_id"].ne("")
            df.loc[sel, "group_type"] = "IMIS"
            df.loc[sel, "group_id"] = df.loc[sel, "imis_id"]
        return df

    def prepare_history(self):
        closed_df = self.prepare_closed_history()
        open_df = self.prepare_open_history()
        history_df = pd.concat([closed_df, open_df], ignore_index=True)
        history_df = history_df[history_df["cm_id"] != ""].copy()
        history_df = history_df[history_df["history_date"].notna()].copy()
        history_df = history_df[history_df["history_date"] <= self.as_of_date].copy()
        history_df = self.add_group_key(history_df)
        self.history_df = history_df

    def prepare_closed_history(self) -> pd.DataFrame:
        fields = get_nested(self.config, ["fields", "closed_report"], {})
        df = self.raw_closed_df.copy()

        review_id_col = get_col(fields, "review_id", DefaultFields.Closed.review_id)
        customer_col = get_col(fields, "customer", DefaultFields.Closed.customer)
        cm_col = get_col(fields, "cm_id", DefaultFields.Closed.cm_id)
        first_claimed_col = get_col(fields, "first_claimed_id", DefaultFields.Closed.first_claimed_id)
        status_col = get_col(fields, "review_status", DefaultFields.Closed.review_status)
        completed_status = get_col(fields, "completed_status", DefaultFields.Closed.completed_status)
        cancelled_status = get_col(fields, "cancelled_status", DefaultFields.Closed.cancelled_status)
        rm_col = get_col(fields, "rm_num", DefaultFields.Closed.rm_num)
        segment_col = get_col(fields, "segment", DefaultFields.Closed.segment)
        date_candidates = fields.get("date_candidates", DefaultFields.Closed.date_candidates)

        self.check_columns(df, [customer_col, cm_col], "closed_report")

        if self.config.get("closed_completed_only", True) and status_col in df.columns:
            df = df[df[status_col].eq(completed_status)].copy()

        closed = pd.DataFrame(index=df.index)
        closed["source"] = "closed"
        closed["review_id"] = df[review_id_col].apply(clean_code) if review_id_col in df.columns else ""
        closed["customer_id"] = df[customer_col].apply(clean_customer)
        closed["cin"] = closed["customer_id"]
        closed["cm_id"] = df[cm_col].apply(clean_staff)
        if first_claimed_col in df.columns and status_col in df.columns:
            cancelled = df[status_col].eq(cancelled_status)
            use_first_claimed = cancelled & closed["cm_id"].eq("")
            closed.loc[use_first_claimed, "cm_id"] = df.loc[use_first_claimed, first_claimed_col].apply(clean_staff)
        closed["history_date"] = pd.to_datetime(coalesce_columns(df, date_candidates), errors="coerce", format="mixed")
        closed["rm_num"] = df[rm_col].fillna("").astype(str) if rm_col in df.columns else ""
        closed["segment"] = df[segment_col].fillna("").astype(str) if segment_col in df.columns else ""
        closed["segment_norm"] = closed["segment"].apply(norm_text)
        closed = closed.merge(self.horis_mg_df, how="left", on="cin")
        closed["mg_id"] = closed["mapped_mg_id"].fillna("")
        closed["mg_name"] = closed["mapped_mg_name"].fillna("")
        closed = self.map_history_imis(closed)
        return closed

    def prepare_open_history(self) -> pd.DataFrame:
        fields = get_nested(self.config, ["fields", "open_report"], {})
        df = self.raw_open_df.copy()

        review_id_col = get_col(fields, "review_id", DefaultFields.Open.review_id)
        customer_col = get_col(fields, "customer", DefaultFields.Open.customer)
        stage_col = get_col(fields, "stage", DefaultFields.Open.stage)
        assigned_to_user_col = get_col(fields, "assigned_to_user", DefaultFields.Open.assigned_to_user)
        user_id_col = get_col(fields, "user_id", DefaultFields.Open.user_id)
        latest_dc_by_col = get_col(fields, "latest_dc_by", DefaultFields.Open.latest_dc_by)
        cm_col = get_col(fields, "cm_id", DefaultFields.Open.cm_id)
        rm_col = get_col(fields, "rm_num", DefaultFields.Open.rm_num)
        segment_col = get_col(fields, "segment", DefaultFields.Open.segment)
        date_candidates = fields.get("date_candidates", DefaultFields.Open.date_candidates)

        self.check_columns(df, [customer_col], "open_report")

        open_df = pd.DataFrame(index=df.index)
        open_df["source"] = "open"
        open_df["review_id"] = df[review_id_col].apply(clean_code) if review_id_col in df.columns else ""
        open_df["customer_id"] = df[customer_col].apply(clean_customer)
        open_df["cin"] = open_df["customer_id"]
        open_df["cm_id"] = self.derive_open_cm_id(
            df,
            stage_col=stage_col,
            assigned_to_user_col=assigned_to_user_col,
            user_id_col=user_id_col,
            latest_dc_by_col=latest_dc_by_col,
            cm_col=cm_col,
        )
        open_df["history_date"] = pd.to_datetime(coalesce_columns(df, date_candidates), errors="coerce", format="mixed")
        open_df["rm_num"] = df[rm_col].fillna("").astype(str) if rm_col in df.columns else ""
        open_df["segment"] = df[segment_col].fillna("").astype(str) if segment_col in df.columns else ""
        open_df["segment_norm"] = open_df["segment"].apply(norm_text)
        open_df = open_df.merge(self.horis_mg_df, how="left", on="cin")
        open_df["mg_id"] = open_df["mapped_mg_id"].fillna("")
        open_df["mg_name"] = open_df["mapped_mg_name"].fillna("")
        open_df = self.map_history_imis(open_df)
        return open_df

    def map_history_imis(self, history_df: pd.DataFrame) -> pd.DataFrame:
        history_df["imis_customer_key"] = history_df["customer_id"].apply(clean_code)
        history_df = history_df.merge(self.imis_df, how="left", on="imis_customer_key")
        history_df["imis_id"] = history_df["mapped_imis_id"].fillna("")
        history_df["imis_name"] = history_df["mapped_imis_name"].fillna("")
        return history_df

    def derive_open_cm_id(
            self,
            df: pd.DataFrame,
            stage_col: str,
            assigned_to_user_col: str,
            user_id_col: str,
            latest_dc_by_col: str,
            cm_col: str,
    ) -> pd.Series:
        if cm_col in df.columns:
            cm_id = df[cm_col].apply(clean_staff)
        else:
            cm_id = pd.Series("", index=df.index)

        has_required_cols = all(c in df.columns for c in [stage_col, assigned_to_user_col, user_id_col, latest_dc_by_col])
        if not has_required_cols:
            return cm_id

        def get_staff_id(row):
            if row[stage_col] in (st.QC, st.APP):
                return clean_staff(row[latest_dc_by_col])
            assigned_to_user = clean_staff(row[assigned_to_user_col])
            if assigned_to_user.isdigit():
                return assigned_to_user
            user_id = clean_staff(row[user_id_col])
            if user_id.isdigit():
                return user_id
            return clean_staff(row.get(cm_col, ""))

        return df.apply(get_staff_id, axis=1)

    def run_proposed_assignment(self):
        analyst_df = self.build_algorithm_analyst_df()
        task_df = self.build_algorithm_task_df()
        close_df = self.build_algorithm_closed_df()
        open_df = self.build_algorithm_open_df()

        allocator = CDDTaskAllocator(
            df_analysts=analyst_df,
            df_tasks=task_df,
            df_close=close_df,
            df_open=open_df,
            current_date=self.as_of_date,
            config=self.config.get("algorithm_config", {}),
        )
        allocator.read_data()
        allocator.add_history_customer_to_analyst()
        res_df = allocator.run_top_n_analyst(n=self.config.get("top_n", 3))

        if res_df.empty:
            self.proposed_df = pd.DataFrame(columns=["customer_id", "proposed_cm_id"])
            self.warnings.append("Proposed algorithm returns empty result.")
            return

        self.proposed_df = pd.DataFrame({
            "customer_id": res_df["Task ID"].apply(clean_customer),
            "proposed_cm_id": res_df["Assigned Analyst ID 1"].apply(clean_staff),
            "proposed_score": res_df.get("Score Total 1", np.nan),
        })

    def build_algorithm_analyst_df(self) -> pd.DataFrame:
        return pd.DataFrame({
            ah.analyst_id: self.cm_df["cm_id"],
            ah.current_wip: self.cm_df["current_wip"],
            ah.target_wip: self.cm_df["optimal_wip"],
            ah.daily_productivity: self.cm_df["productivity"],
        })

    def build_algorithm_task_df(self) -> pd.DataFrame:
        market = self.config.get("market", "")
        legal_entity = self.config.get("legal_entity", "")
        due_date = self.case_df["due_date"].fillna(self.as_of_date + pd.Timedelta(days=120))
        return pd.DataFrame({
            th.task_id: self.case_df["customer_id"],
            th.customer_id: self.case_df["customer_id"],
            th.due_date: due_date,
            th.effort: self.case_df["workload"],
            th.market: market,
            th.legal_entity: legal_entity,
            th.ppm_cdd_source: "",
            th.risk_rating: "",
            mh.MasterGroupCode: self.case_df["mg_id"],
        })

    def build_algorithm_closed_df(self) -> pd.DataFrame:
        closed = self.history_df[self.history_df["source"].eq("closed")].copy()
        return pd.DataFrame({
            ch.CustomerNumber: closed["customer_id"],
            ch.ReviewStatus: "Approval Completed",
            ch.LatestDcFinalisedById: closed["cm_id"],
            ch.RmNum: closed["rm_num"],
            ch.InitiatedDate: closed["history_date"],
            ch.Segment: closed["segment"],
            mh.MasterGroupCode: closed["mg_id"],
        })

    def build_algorithm_open_df(self) -> pd.DataFrame:
        open_df = self.history_df[self.history_df["source"].eq("open")].copy()
        return pd.DataFrame({
            oh.CustomerNumber: open_df["customer_id"],
            oh.StaffID: open_df["cm_id"],
            oh.RmNum: open_df["rm_num"],
            oh.InitiatedDate: open_df["history_date"],
            ch.Segment: open_df["segment"],
            mh.MasterGroupCode: open_df["mg_id"],
        })

    def build_case_detail(self):
        df = self.case_df.copy()
        df = df.merge(self.proposed_df, how="left", on="customer_id")
        df["proposed_cm_id"] = df["proposed_cm_id"].fillna("").apply(clean_staff)
        df = self.add_cm_info(df, cm_col="business_cm_id", prefix="business")
        df = self.add_cm_info(df, cm_col="proposed_cm_id", prefix="proposed")

        for method in ["business", "proposed"]:
            cm_col = f"{method}_cm_id"
            for year in self.lookback_years:
                df = self.add_hit_flags(df, method=method, cm_col=cm_col, year=year)

        self.case_detail_df = df

    def add_cm_info(self, df: pd.DataFrame, cm_col: str, prefix: str) -> pd.DataFrame:
        cm_info = self.cm_df[["cm_id", "cm_name", "team"]].copy()
        cm_info.columns = [cm_col, f"{prefix}_cm_name_from_list", f"{prefix}_team_from_list"]
        return df.merge(cm_info, how="left", on=cm_col)

    def add_hit_flags(self, df: pd.DataFrame, method: str, cm_col: str, year: int) -> pd.DataFrame:
        start_date = self.as_of_date - pd.DateOffset(years=year)
        history = self.history_df[
            (self.history_df["history_date"] >= start_date) &
            (self.history_df["history_date"] <= self.as_of_date)
        ].copy()

        customer_hit_set = set(zip(history["cm_id"], history["customer_id"]))
        mg_hit_set = set(zip(history.loc[history["mg_id"].ne(""), "cm_id"], history.loc[history["mg_id"].ne(""), "mg_id"]))
        group_hit_set = set(
            zip(history.loc[history["group_id"].ne(""), "cm_id"], history.loc[history["group_id"].ne(""), "group_id"])
        )
        customer_review_lookup = self.latest_review_lookup(
            history[history["customer_id"].ne("")],
            ["cm_id", "customer_id"],
        )
        mg_review_lookup = self.latest_review_lookup(
            history[history["mg_id"].ne("")],
            ["cm_id", "mg_id"],
        )

        df[f"{method}_customer_hit_{year}y"] = df.apply(
            lambda r: (r[cm_col], r["customer_id"]) in customer_hit_set if r[cm_col] else False,
            axis=1,
        )
        df[f"{method}_customer_hit_{year}y_last_review_id"] = df.apply(
            lambda r: customer_review_lookup.get((r[cm_col], r["customer_id"]), "")
            if r[f"{method}_customer_hit_{year}y"] else "",
            axis=1,
        )
        df[f"{method}_mg_hit_{year}y"] = df.apply(
            lambda r: (r[cm_col], r["mg_id"]) in mg_hit_set if r[cm_col] and r["mg_id"] else False,
            axis=1,
        )
        df[f"{method}_mg_hit_{year}y_last_review_id"] = df.apply(
            lambda r: mg_review_lookup.get((r[cm_col], r["mg_id"]), "")
            if r[f"{method}_mg_hit_{year}y"] else "",
            axis=1,
        )
        df[f"{method}_group_hit_{year}y"] = df.apply(
            lambda r: (r[cm_col], r["group_id"]) in group_hit_set if r[cm_col] and r["group_id"] else False,
            axis=1,
        )
        df[f"{method}_customer_or_group_hit_{year}y"] = (
                df[f"{method}_customer_hit_{year}y"] | df[f"{method}_group_hit_{year}y"]
        )
        return df

    @staticmethod
    def latest_review_lookup(history: pd.DataFrame, key_columns: Sequence[str]) -> dict:
        if history.empty:
            return {}
        lookup_df = history[list(key_columns) + ["history_date", "review_id"]].copy()
        lookup_df["review_id"] = lookup_df["review_id"].fillna("").astype(str)
        lookup_df = lookup_df.sort_values("history_date")
        latest_df = lookup_df.drop_duplicates(list(key_columns), keep="last")
        return latest_df.set_index(list(key_columns))["review_id"].to_dict()

    def build_summary(self) -> pd.DataFrame:
        rows = []
        for method in ["business", "proposed"]:
            cm_col = f"{method}_cm_id"
            assigned_df = self.case_detail_df[self.case_detail_df[cm_col].ne("")].copy()
            for year in self.lookback_years:
                rows.append({
                    "method": method,
                    "window_year": year,
                    "assigned_case_count": len(assigned_df),
                    "customer_hit_count": assigned_df[f"{method}_customer_hit_{year}y"].sum(),
                    "customer_hit_rate": self.safe_rate(
                        assigned_df[f"{method}_customer_hit_{year}y"].sum(),
                        assigned_df["customer_id"].ne("").sum(),
                    ),
                    "mg_hit_count": assigned_df[f"{method}_mg_hit_{year}y"].sum(),
                    "mg_hit_rate": self.safe_rate(
                        assigned_df[f"{method}_mg_hit_{year}y"].sum(),
                        assigned_df["mg_id"].ne("").sum(),
                    ),
                    "group_hit_count": assigned_df[f"{method}_group_hit_{year}y"].sum(),
                    "group_hit_rate": self.safe_rate(
                        assigned_df[f"{method}_group_hit_{year}y"].sum(),
                        assigned_df["group_id"].ne("").sum(),
                    ),
                    "customer_or_group_hit_count": assigned_df[f"{method}_customer_or_group_hit_{year}y"].sum(),
                    "customer_or_group_hit_rate": self.safe_rate(
                        assigned_df[f"{method}_customer_or_group_hit_{year}y"].sum(),
                        len(assigned_df),
                    ),
                })

        summary_df = pd.DataFrame(rows)
        wip_by_cm_df, _, _ = self.build_wip_summary()
        wip_summary = wip_by_cm_df.groupby("method", as_index=False).agg(
            cm_count=("cm_id", "count"),
            avg_wip_after=("wip_after_assignment", "mean"),
            max_wip_after=("wip_after_assignment", "max"),
            over_target_cm_count=("over_target", "sum"),
            total_over_target=("over_target_amount", "sum"),
        )
        summary_df = summary_df.merge(wip_summary, how="left", on="method")
        return summary_df

    def build_wip_summary(self) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        rows = []
        for method, cm_col, team_col in [
            ("business", "business_cm_id", "business_team"),
            ("proposed", "proposed_cm_id", "proposed_team_from_list"),
        ]:
            assigned = self.case_detail_df[self.case_detail_df[cm_col].ne("")].copy()
            if method == "proposed":
                assigned[team_col] = assigned[team_col].fillna("")
            assigned_group = assigned.groupby(cm_col, as_index=False).agg(
                assigned_case_count=("case_id", "count"),
                assigned_workload=("workload", "sum"),
                assigned_team=(team_col, lambda s: self.first_non_blank(s)),
            )
            assigned_group = assigned_group.rename(columns={cm_col: "cm_id"})
            missing_cm = sorted(set(assigned_group["cm_id"]) - set(self.cm_df["cm_id"]))
            if missing_cm:
                self.warnings.append(
                    f"{method} assignment has CM not found in CM_LIST: {', '.join(missing_cm[:20])}"
                )
            method_df = self.cm_df.merge(assigned_group, how="outer", on="cm_id")
            method_df["method"] = method
            method_df["cm_name"] = method_df["cm_name"].fillna("")
            method_df["team"] = method_df["team"].fillna("")
            method_df["current_wip"] = method_df["current_wip"].fillna(0)
            method_df["optimal_wip"] = method_df["optimal_wip"].fillna(0)
            method_df["assigned_case_count"] = method_df["assigned_case_count"].fillna(0)
            method_df["assigned_workload"] = method_df["assigned_workload"].fillna(0)
            method_df["assigned_team"] = method_df["assigned_team"].fillna("")
            method_df.loc[method_df["team"].eq(""), "team"] = method_df.loc[method_df["team"].eq(""), "assigned_team"]
            rows.append(method_df)

        wip_by_cm = pd.concat(rows, ignore_index=True)
        wip_by_cm["wip_after_assignment"] = wip_by_cm["current_wip"] + wip_by_cm["assigned_workload"]
        wip_by_cm["over_target_amount"] = (wip_by_cm["wip_after_assignment"] - wip_by_cm["optimal_wip"]).clip(lower=0)
        wip_by_cm["over_target"] = wip_by_cm["over_target_amount"] > 0
        wip_by_cm["utilization_after"] = np.where(
            wip_by_cm["optimal_wip"] > 0,
            wip_by_cm["wip_after_assignment"] / wip_by_cm["optimal_wip"],
            np.nan,
        )

        total_assigned = wip_by_cm.groupby("method")["assigned_case_count"].transform("sum")
        wip_by_cm["share_of_total_assigned"] = np.where(
            total_assigned > 0,
            wip_by_cm["assigned_case_count"] / total_assigned,
            np.nan,
        )

        wip_by_team_cm = wip_by_cm[
            [
                "method", "team", "cm_id", "cm_name", "current_wip", "optimal_wip",
                "assigned_case_count", "assigned_workload", "wip_after_assignment",
                "over_target", "over_target_amount", "utilization_after", "share_of_total_assigned",
            ]
        ].copy()

        wip_by_team = wip_by_cm.groupby(["method", "team"], as_index=False).agg(
            cm_count=("cm_id", "count"),
            initial_wip_sum=("current_wip", "sum"),
            optimal_wip_sum=("optimal_wip", "sum"),
            assigned_case_count=("assigned_case_count", "sum"),
            assigned_workload=("assigned_workload", "sum"),
            wip_after_assignment_sum=("wip_after_assignment", "sum"),
            over_target_cm_count=("over_target", "sum"),
            over_target_amount_sum=("over_target_amount", "sum"),
            avg_utilization_after=("utilization_after", "mean"),
            max_utilization_after=("utilization_after", "max"),
        )
        return wip_by_cm, wip_by_team_cm, wip_by_team

    def build_mapping_check(self) -> pd.DataFrame:
        df = self.case_df.copy()
        df["mg_id_diff_from_business_output"] = (
                df["business_mg_id"].ne("") &
                df["mapped_mg_id"].fillna("").ne("") &
                df["business_mg_id"].ne(df["mapped_mg_id"].fillna(""))
        )
        df["imis_id_diff_from_business_output"] = (
                df["business_imis_id"].ne("") &
                df["mapped_imis_id"].fillna("").ne("") &
                df["business_imis_id"].ne(df["mapped_imis_id"].fillna(""))
        )
        return df[
            [
                "case_row_id", "case_id", "customer_id", "cin", "segment",
                "business_mg_id", "mapped_mg_id", "mg_id", "business_mg_name", "mapped_mg_name",
                "business_imis_id", "mapped_imis_id", "imis_id", "business_imis_name", "mapped_imis_name",
                "group_type", "group_id", "mg_id_diff_from_business_output",
                "imis_id_diff_from_business_output",
            ]
        ]

    def build_metric_definitions(self) -> pd.DataFrame:
        windows = ", ".join(f"{year}y" for year in self.lookback_years)
        rows = [
            {
                "sheet": "summary",
                "field": "assigned_case_count",
                "definition": "该方法下已分配 CM 的 case 数。",
                "calculation_method": "统计 case_detail 中 <method>_cm_id 不为空的行数。",
                "denominator": "",
            },
            {
                "sheet": "summary",
                "field": "customer_hit_count",
                "definition": "分配到的 CM 在回看窗口内做过同一个 customer 的 case 数。",
                "calculation_method": "汇总 case_detail.<method>_customer_hit_<window>y。历史窗口为 history_date >= as_of_date - window_year 且 <= as_of_date。",
                "denominator": "",
            },
            {
                "sheet": "summary",
                "field": "customer_hit_rate",
                "definition": "已分配 case 中 customer 命中的比例。",
                "calculation_method": "customer_hit_count / 已分配且 customer_id 不为空的 case 数。",
                "denominator": "已分配且 customer_id 不为空的 case。",
            },
            {
                "sheet": "summary",
                "field": "mg_hit_count",
                "definition": "分配到的 CM 在回看窗口内做过同一个 Master Group 的 case 数。",
                "calculation_method": "汇总 case_detail.<method>_mg_hit_<window>y。",
                "denominator": "",
            },
            {
                "sheet": "summary",
                "field": "mg_hit_rate",
                "definition": "有 MG 的已分配 case 中 MG 命中的比例。",
                "calculation_method": "mg_hit_count / 已分配且 mg_id 不为空的 case 数。",
                "denominator": "已分配且 mg_id 不为空的 case。",
            },
            {
                "sheet": "summary",
                "field": "group_hit_count",
                "definition": "分配到的 CM 在回看窗口内做过同一个比较 group 的 case 数。",
                "calculation_method": "汇总 case_detail.<method>_group_hit_<window>y。CORP/CMB 使用 MG；use_imis_as_group_for_bbrm=true 时 BBRM 使用 IMIS。",
                "denominator": "",
            },
            {
                "sheet": "summary",
                "field": "group_hit_rate",
                "definition": "有比较 group 的已分配 case 中 group 命中的比例。",
                "calculation_method": "group_hit_count / 已分配且 group_id 不为空的 case 数。",
                "denominator": "已分配且 group_id 不为空的 case。",
            },
            {
                "sheet": "summary",
                "field": "customer_or_group_hit_count",
                "definition": "分配到的 CM 命中同 customer 或同比较 group 任一条件的 case 数。",
                "calculation_method": "汇总 (<method>_customer_hit_<window>y OR <method>_group_hit_<window>y)。",
                "denominator": "",
            },
            {
                "sheet": "summary",
                "field": "customer_or_group_hit_rate",
                "definition": "已分配 case 中命中同 customer 或同比较 group 任一条件的比例。",
                "calculation_method": "customer_or_group_hit_count / assigned_case_count。",
                "denominator": "已分配 case 数。",
            },
            {
                "sheet": "summary / wip_by_cm",
                "field": "wip_after_assignment",
                "definition": "应用本次分配结果后 CM 的在手工作量。",
                "calculation_method": "current_wip + assigned_workload。",
                "denominator": "",
            },
            {
                "sheet": "summary / wip_by_cm",
                "field": "over_target, over_target_amount",
                "definition": "分配后 CM 是否超过 optimal WIP，以及超过多少。",
                "calculation_method": "over_target_amount = max(wip_after_assignment - optimal_wip, 0)；over_target = over_target_amount > 0。",
                "denominator": "",
            },
            {
                "sheet": "case_detail",
                "field": "<method>_customer_hit_<window>y_last_review_id",
                "definition": "支撑 customer hit 标记的最近一条同 CM、同 customer 历史记录的 Review ID。",
                "calculation_method": f"对每个配置窗口（{windows}），先过滤历史窗口，再按 assigned CM + customer_id 取 history_date 最新的一条记录。",
                "denominator": "",
            },
            {
                "sheet": "case_detail",
                "field": "<method>_mg_hit_<window>y_last_review_id",
                "definition": "支撑 MG hit 标记的最近一条同 CM、同 MG 历史记录的 Review ID。",
                "calculation_method": f"对每个配置窗口（{windows}），先过滤历史窗口，再按 assigned CM + mg_id 取 history_date 最新的一条记录。",
                "denominator": "",
            },
            {
                "sheet": "all",
                "field": "as_of_date / lookback window",
                "definition": "实验只统计 as_of_date 当天及之前的历史记录。",
                "calculation_method": f"as_of_date = {self.as_of_date.date()}；配置的回看窗口 = {windows}。",
                "denominator": "",
            },
        ]
        return pd.DataFrame(rows)

    def output_result(
            self,
            summary_df: pd.DataFrame,
            wip_by_cm_df: pd.DataFrame,
            wip_by_team_cm_df: pd.DataFrame,
            wip_by_team_df: pd.DataFrame,
            mapping_check_df: pd.DataFrame,
    ):
        output_file = self.config.get("output_file", "assignment_experiment_result.xlsx")
        output_path = resolve_working_path(output_file)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(output_path) as writer:
            summary_df.to_excel(writer, sheet_name="summary", index=False)
            self.build_metric_definitions().to_excel(writer, sheet_name="metric_definitions", index=False)
            self.case_detail_df.to_excel(writer, sheet_name="case_detail", index=False)
            wip_by_cm_df.to_excel(writer, sheet_name="wip_by_cm", index=False)
            wip_by_team_cm_df.to_excel(writer, sheet_name="wip_by_team_cm", index=False)
            wip_by_team_df.to_excel(writer, sheet_name="wip_by_team", index=False)
            mapping_check_df.to_excel(writer, sheet_name="mapping_check", index=False)
            pd.DataFrame({"warning": self.warnings}).to_excel(writer, sheet_name="warnings", index=False)
            self.write_field_config(writer)
            if self.debug:
                self.history_df.head(5000).to_excel(writer, sheet_name="debug_history_sample", index=False)
                self.case_df.to_excel(writer, sheet_name="debug_case_normalized", index=False)

        print(f"Experiment result written to: {output_path}")

    def write_field_config(self, writer):
        rows = []
        for section, values in self.config.get("fields", {}).items():
            for key, value in values.items():
                rows.append({"section": section, "field_key": key, "column_name": value})
        pd.DataFrame(rows).to_excel(writer, sheet_name="field_config_used", index=False)

    def check_columns(self, df: pd.DataFrame, columns: Sequence[str], source_name: str):
        missing = [col for col in columns if col and col not in df.columns]
        if missing:
            raise ValueError(f"{source_name} missing required columns: {missing}")

    @staticmethod
    def safe_rate(num, den):
        if den == 0:
            return np.nan
        return num / den

    @staticmethod
    def first_non_blank(values):
        for value in values:
            if pd.notna(value) and str(value).strip() != "":
                return value
        return ""


def setup_log():
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )


def main():
    parser = argparse.ArgumentParser(description="Run simple assignment familiarity and WIP experiment.")
    parser.add_argument("--config", required=True, help="Path to experiment config json.")
    args = parser.parse_args()
    setup_log()
    experiment = AssignmentExperiment(args.config)
    experiment.run()


if __name__ == "__main__":
    main()
```
