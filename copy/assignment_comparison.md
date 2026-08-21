# assignment_comparison.py

```python
import argparse
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd

from case_assignment import CDDTaskAllocator
from header import (
    AnalystHeader as ah,
    ClosedReviewHeader as ch,
    MasterGroupHeader as mh,
    OpenReviewHeader as oh,
    TaskHeader as th,
)
from task_clean import TaskClean


DEFAULT_SCORE_CONFIG = {
    "cus_affinity_weight": 100.0,
    "mg_affinity_weight": 80.0,
    "rm_affinity_weight": 50.0,
    "segment_affinity_weight": 100.0,
    "capacity_weight": 5.0,
    "overload_penalty": 5.0,
    "simulate_decay": True,
}

METRIC_NAMES = [
    "assigned_cases",
    "assigned_effort",
    "same_customer_rate",
    "same_mg_rate",
    "same_rm_rate",
    "same_segment_rate",
    "mg_split_group_rate",
    "rm_split_group_rate",
    "analyst_load_cv",
    "analyst_load_gini",
    "target_share_abs_error",
    "over_target_analysts",
    "over_target_effort",
    "unique_assignees",
    "unassigned_cases",
]


@dataclass
class InputFrames:
    analysts: pd.DataFrame
    tasks: pd.DataFrame
    close: pd.DataFrame
    open: pd.DataFrame
    raw_task: pd.DataFrame
    ram: Optional[pd.DataFrame] = None
    mg: Optional[pd.DataFrame] = None


def read_table(path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    suffix = Path(path).suffix.lower()
    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(path, sheet_name=sheet_name or 0)
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix in {".txt", ".tsv"}:
        return pd.read_csv(path, sep="\t")
    raise ValueError(f"Unsupported input file type: {path}")


def read_many(paths: Sequence[str]) -> pd.DataFrame:
    frames = [read_table(path) for path in paths]
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def as_list(value) -> List:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, tuple):
        return list(value)
    return [value]


def clean_id(value, strip_zero: bool = True) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    if strip_zero:
        text = text.lstrip("0")
    return text


def norm_text(value) -> str:
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def first_existing_col(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    for col in candidates:
        if col in df.columns:
            return col
    return None


def safe_ratio(num: float, den: float) -> float:
    if den == 0 or pd.isna(den):
        return np.nan
    return float(num) / float(den)


def gini(values: Sequence[float]) -> float:
    arr = np.array([v for v in values if pd.notna(v)], dtype=float)
    if arr.size == 0:
        return np.nan
    if np.all(arr == 0):
        return 0.0
    arr = np.sort(np.abs(arr))
    n = arr.size
    return float((2 * np.arange(1, n + 1).dot(arr) / arr.sum() - (n + 1)) / n)


def load_input_frames(market_config: dict, current_date: str) -> InputFrames:
    oe_report_df = read_many(as_list(market_config.get("OE_REPORT_FILES", [])))
    close_df = read_many(as_list(market_config.get("CLOSE_REPORT_FILES", [])))
    open_files = as_list(market_config.get("OPEN_REPORT_FILES", []))
    if len(open_files) != 1:
        raise ValueError("OPEN_REPORT_FILES should point to one file for comparison.")
    open_df = read_table(open_files[0])

    analyst_file = market_config["CASE_MANAGER_LIST_FILES"]
    analyst_df = read_table(analyst_file)

    ram_df = None
    ram_files = as_list(market_config.get("RAM_FILES", []))
    if ram_files:
        ram_df = read_many(ram_files)

    mg_df = None
    mg_files = as_list(market_config.get("MG_FILES", []))
    if mg_files:
        mg_df = read_many(mg_files)

    cleaner = TaskClean(
        open_df,
        close_df,
        oe_report_df,
        ram_file=ram_df,
        df_mg=mg_df,
        market=market_config.get("MARKET", ""),
        legal_entity=as_list(market_config.get("LEGAL_ENTITY", [])),
        current_date=current_date,
        init_lead_month=market_config.get("INIT_LEAD_MONTH", 4),
    )
    cleaner.data_preprocess()
    cleaner.close_data_process()
    cleaner.clean_task()

    return InputFrames(
        analysts=analyst_df,
        tasks=cleaner.df_task,
        close=cleaner.df_close,
        open=cleaner.df_open,
        raw_task=oe_report_df,
        ram=ram_df,
        mg=mg_df,
    )


def build_history_maps(
    close_df: pd.DataFrame, open_df: pd.DataFrame
) -> Tuple[pd.DataFrame, Dict[str, set], Dict[str, set], Dict[str, set], Dict[str, str], Dict[str, str]]:
    close = close_df.copy()
    open_cases = open_df.copy()

    for df, customer_col in ((close, ch.CustomerNumber), (open_cases, oh.CustomerNumber)):
        if customer_col in df.columns:
            df[customer_col] = df[customer_col].apply(clean_id)

    if ch.InitiatedDate in close.columns:
        close[ch.InitiatedDate] = pd.to_datetime(close[ch.InitiatedDate], errors="coerce", format="mixed")
    if oh.InitiatedDate in open_cases.columns:
        open_cases[oh.InitiatedDate] = pd.to_datetime(open_cases[oh.InitiatedDate], errors="coerce", format="mixed")

    if ch.ReviewStatus in close.columns:
        close_completed = close[close[ch.ReviewStatus].eq("Approval Completed")].copy()
    else:
        close_completed = close.copy()

    staff_col = ch.StaffID if ch.StaffID in close_completed.columns else ch.LatestDcFinalisedById
    close_completed[staff_col] = close_completed[staff_col].apply(clean_id)
    close_completed = close_completed[close_completed[staff_col] != ""].copy()
    close_completed = close_completed.sort_values(ch.InitiatedDate) if ch.InitiatedDate in close_completed else close_completed

    latest_customer = pd.DataFrame(columns=[ch.CustomerNumber, "last_customer_analyst"])
    if ch.CustomerNumber in close_completed.columns and not close_completed.empty:
        latest_customer = close_completed.drop_duplicates(ch.CustomerNumber, keep="last")[
            [ch.CustomerNumber, staff_col]
        ].rename(columns={staff_col: "last_customer_analyst"})

    mg_owners: Dict[str, set] = {}
    if mh.MasterGroupCode in close_completed.columns:
        for mg, sub_df in close_completed.groupby(mh.MasterGroupCode, dropna=True):
            mg_key = clean_id(mg)
            if mg_key:
                mg_owners[mg_key] = set(sub_df[staff_col].apply(clean_id)) - {""}

    rm_owners: Dict[str, set] = {}
    if ch.RmNum in close_completed.columns:
        for rm, sub_df in close_completed.groupby(ch.RmNum, dropna=True):
            rm_key = clean_id(rm, strip_zero=False)
            if rm_key and rm_key != "-":
                rm_owners[rm_key] = set(sub_df[staff_col].apply(clean_id)) - {""}

    segment_owners: Dict[str, set] = {}
    if ch.Segment in close_completed.columns:
        for segment, sub_df in close_completed.groupby(ch.Segment, dropna=True):
            seg_key = norm_text(segment)
            if seg_key and seg_key != "-":
                segment_owners[seg_key] = set(sub_df[staff_col].apply(clean_id)) - {""}

    wip_customer_owner: Dict[str, str] = {}
    wip_group_owner: Dict[str, str] = {}
    if not open_cases.empty and oh.StaffID in open_cases.columns:
        open_cases[oh.StaffID] = open_cases[oh.StaffID].apply(clean_id)
        open_cases = open_cases[open_cases[oh.StaffID] != ""].copy()
        if oh.InitiatedDate in open_cases.columns:
            open_cases = open_cases.sort_values(oh.InitiatedDate)
        if oh.CustomerNumber in open_cases.columns:
            latest_open_customer = open_cases.drop_duplicates(oh.CustomerNumber, keep="last")
            wip_customer_owner = dict(zip(latest_open_customer[oh.CustomerNumber], latest_open_customer[oh.StaffID]))
        if mh.MasterGroupCode in open_cases.columns:
            latest_open_group = open_cases[open_cases[mh.MasterGroupCode].apply(clean_id) != ""]
            latest_open_group = latest_open_group.drop_duplicates(mh.MasterGroupCode, keep="last")
            wip_group_owner.update(
                {
                    clean_id(row[mh.MasterGroupCode]): clean_id(row[oh.StaffID])
                    for _, row in latest_open_group.iterrows()
                }
            )
        if oh.RmNum in open_cases.columns:
            latest_open_rm = open_cases[open_cases[oh.RmNum].apply(lambda x: clean_id(x, False) not in {"", "-"})]
            latest_open_rm = latest_open_rm.drop_duplicates(oh.RmNum, keep="last")
            wip_group_owner.update(
                {
                    "RM:" + clean_id(row[oh.RmNum], strip_zero=False): clean_id(row[oh.StaffID])
                    for _, row in latest_open_rm.iterrows()
                }
            )

    return latest_customer, mg_owners, rm_owners, segment_owners, wip_customer_owner, wip_group_owner


def build_task_features(tasks_df: pd.DataFrame, close_df: pd.DataFrame, open_df: pd.DataFrame) -> pd.DataFrame:
    tasks = tasks_df.copy()
    latest_customer, mg_owners, rm_owners, segment_owners, wip_customer_owner, wip_group_owner = build_history_maps(
        close_df, open_df
    )

    tasks[th.task_id] = tasks[th.task_id].apply(clean_id)
    tasks[th.customer_id] = tasks[th.customer_id].apply(clean_id)
    tasks[th.due_date] = pd.to_datetime(tasks[th.due_date], errors="coerce", format="mixed")
    if th.effort not in tasks.columns:
        tasks[th.effort] = 1.0
    tasks[th.effort] = pd.to_numeric(tasks[th.effort], errors="coerce").fillna(1.0)

    features = pd.DataFrame(index=tasks.index)
    features["task_id"] = tasks[th.task_id].apply(clean_id)
    features["customer_id"] = tasks[th.customer_id].apply(clean_id)
    features["due_date"] = tasks[th.due_date]
    features["effort"] = tasks[th.effort]
    for optional_col, canonical in [
        (th.market, "market"),
        (th.legal_entity, "legal_entity"),
        (th.risk_rating, "risk_rating"),
        (mh.MasterGroupCode, "mg_num"),
        (ch.RmNum, "rm_num"),
        (ch.Segment, "segment"),
    ]:
        if optional_col in tasks.columns:
            features[canonical] = tasks[optional_col]
    features["init_date"] = features["due_date"] - pd.to_timedelta(120, unit="D")
    if "mg_num" not in features.columns:
        features["mg_num"] = ""
    if "rm_num" not in features.columns:
        features["rm_num"] = ""
    if "segment" not in features.columns:
        features["segment"] = ""

    features["mg_num"] = features["mg_num"].apply(clean_id)
    features["rm_num"] = features["rm_num"].apply(lambda x: clean_id(x, strip_zero=False))
    features["segment"] = features["segment"].apply(norm_text)

    if not latest_customer.empty:
        features = features.merge(latest_customer.rename(columns={ch.CustomerNumber: "customer_id"}), how="left")
    else:
        features["last_customer_analyst"] = ""

    features["last_customer_analyst"] = features["last_customer_analyst"].fillna("").apply(clean_id)
    features["wip_customer_analyst"] = features["customer_id"].map(wip_customer_owner).fillna("")
    features["mg_history_analysts"] = features["mg_num"].map(lambda x: sorted(mg_owners.get(x, set())))
    features["rm_history_analysts"] = features["rm_num"].map(lambda x: sorted(rm_owners.get(x, set())))
    features["segment_history_analysts"] = features["segment"].map(lambda x: sorted(segment_owners.get(x, set())))
    features["wip_group_analyst"] = features["mg_num"].map(wip_group_owner).fillna("")
    rm_wip = features["rm_num"].map(lambda x: wip_group_owner.get("RM:" + x, ""))
    features.loc[features["wip_group_analyst"].eq(""), "wip_group_analyst"] = rm_wip
    features["group_key_mg"] = features["mg_num"].apply(lambda x: "MG:" + x if x else "")
    features["group_key_rm"] = features["rm_num"].apply(lambda x: "RM:" + x if x and x != "-" else "")
    features["legacy_group_key"] = np.where(
        features["group_key_mg"] != "",
        features["group_key_mg"],
        np.where(features["group_key_rm"] != "", features["group_key_rm"], ""),
    )
    return features


def analyst_capacity_frame(analyst_df: pd.DataFrame) -> pd.DataFrame:
    analysts = analyst_df.copy()
    analysts[ah.analyst_id] = analysts[ah.analyst_id].apply(clean_id)
    for col, default in [(ah.current_wip, 0.0), (ah.target_wip, 10.0), (ah.daily_productivity, 1.0)]:
        if col not in analysts.columns:
            analysts[col] = default
        analysts[col] = pd.to_numeric(analysts[col], errors="coerce").fillna(default)
    return analysts


def run_proposed(frames: InputFrames, current_date: str, config: dict, top_n: int = 3) -> pd.DataFrame:
    score_config = DEFAULT_SCORE_CONFIG.copy()
    score_config.update(config or {})
    allocator = CDDTaskAllocator(
        df_analysts=frames.analysts.copy(),
        df_tasks=frames.tasks.copy(),
        df_close=frames.close.copy(),
        df_open=frames.open.copy(),
        current_date=current_date,
        config=score_config,
    )
    allocator.read_data()
    allocator.add_history_customer_to_analyst()
    result = allocator.run_top_n_analyst(n=top_n)
    if result.empty:
        return pd.DataFrame(columns=["task_id", "assigned_analyst_id", "assigned_unit", "method"])

    out = pd.DataFrame(
        {
            "task_id": result["Task ID"].apply(clean_id),
            "assigned_analyst_id": result["Assigned Analyst ID 1"].apply(clean_id),
            "assigned_unit": result["Assigned Analyst ID 1"].apply(clean_id),
            "method": "proposed_score",
            "score_total": result.get("Score Total 1", np.nan),
            "score_customer_affinity": result.get("Score Customer Affinity 1", np.nan),
            "score_mg_affinity": result.get("Score MG Affinity 1", np.nan),
            "score_rm_affinity": result.get("Score RM Affinity 1", np.nan),
            "score_segment_affinity": result.get("Score Segment Affinity 1", np.nan),
            "score_capacity": result.get("Score Capacity 1", np.nan),
        }
    )
    return out


def choose_member_for_unit(
    unit_id: str,
    unit_members: Dict[str, List[str]],
    analyst_load: Dict[str, float],
) -> str:
    members = unit_members.get(unit_id, [unit_id])
    members = [m for m in members if m]
    if not members:
        return unit_id
    return min(members, key=lambda member: (analyst_load.get(member, 0.0), member))


def make_legacy_units(
    analyst_df: pd.DataFrame, legacy_config: dict
) -> Tuple[pd.DataFrame, Dict[str, List[str]], Dict[str, float], Dict[str, str]]:
    analysts = analyst_capacity_frame(analyst_df)
    requested_unit_col = legacy_config.get("assignment_unit_col")
    if requested_unit_col and requested_unit_col in analysts.columns:
        unit_col = requested_unit_col
    else:
        unit_col = ah.analyst_id

    analysts["unit_id"] = analysts[unit_col].apply(norm_text if unit_col != ah.analyst_id else clean_id)
    analysts = analysts[analysts["unit_id"] != ""].copy()

    unit_members = (
        analysts.groupby("unit_id")[ah.analyst_id]
        .apply(lambda s: [clean_id(v) for v in s if clean_id(v)])
        .to_dict()
    )
    analyst_to_unit = dict(zip(analysts[ah.analyst_id], analysts["unit_id"]))
    analyst_load = dict(zip(analysts[ah.analyst_id], analysts[ah.current_wip]))

    if unit_col == ah.analyst_id:
        grouped = analysts.rename(columns={ah.analyst_id: "unit_id_raw"})
        units = pd.DataFrame(
            {
                "unit_id": grouped["unit_id"],
                "target_capacity": grouped[ah.target_wip],
                "current_load": grouped[ah.current_wip],
                "unit_sort": grouped["unit_id"],
            }
        )
    else:
        units = analysts.groupby("unit_id", as_index=False).agg(
            target_capacity=(ah.target_wip, "sum"),
            current_load=(ah.current_wip, "sum"),
        )
        units["unit_sort"] = units["unit_id"]

    fte_col = legacy_config.get("fte_col")
    if fte_col and fte_col in analysts.columns:
        fte = analysts.groupby("unit_id")[fte_col].sum()
        units["share_basis"] = units["unit_id"].map(fte).fillna(0.0)
    else:
        units["share_basis"] = (units["target_capacity"] - units["current_load"]).clip(lower=0)
        if units["share_basis"].sum() <= 0:
            units["share_basis"] = units["target_capacity"].clip(lower=0)
    if units["share_basis"].sum() <= 0:
        units["share_basis"] = 1.0

    skip_units = {norm_text(x) for x in legacy_config.get("skip_units", [])}
    units["skipped"] = units["unit_id"].apply(lambda x: norm_text(x) in skip_units)
    return units, unit_members, analyst_load, analyst_to_unit


def assign_by_capacity(
    units: pd.DataFrame,
    required_cases: int,
    cursor: int,
    rng: np.random.Generator,
) -> Tuple[str, int]:
    active = units[~units["skipped"]].copy()
    if active.empty:
        active = units.copy()
    active = active.sort_values("assign_seq")
    rows = active.to_dict("records")
    n = len(rows)
    if n == 0:
        return "", cursor

    for offset in range(n):
        idx = (cursor + offset) % n
        row = rows[idx]
        if row.get("remaining_capacity", 0) >= required_cases:
            return row["unit_id"], (idx + 1) % n

    best = max(rows, key=lambda row: (row.get("remaining_capacity", 0), row["unit_id"]))
    best_pos = [i for i, row in enumerate(rows) if row["unit_id"] == best["unit_id"]][0]
    return best["unit_id"], (best_pos + 1) % n


def run_legacy_simulator(
    frames: InputFrames,
    features: pd.DataFrame,
    legacy_config: dict,
    seed: int = 42,
) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    units, unit_members, analyst_load, analyst_to_unit = make_legacy_units(frames.analysts, legacy_config)
    units = units.copy()
    units["assign_seq"] = rng.permutation(len(units)) if len(units) else []

    total_cases = len(features)
    active_mask = ~units["skipped"]
    total_share = units.loc[active_mask, "share_basis"].sum()
    if total_share <= 0:
        total_share = units["share_basis"].sum()
        active_mask = pd.Series(True, index=units.index)
    units["quota"] = 0
    units.loc[active_mask, "quota"] = np.ceil(units.loc[active_mask, "share_basis"] / total_share * total_cases)
    units["assigned_cases"] = 0
    units["remaining_capacity"] = units["quota"].astype(float)

    sequence = features.sample(frac=1.0, random_state=int(seed)).reset_index(drop=True)
    assignments = []
    assigned_task_ids = set()
    cursor = 0

    def assign_tasks(sel: pd.Series, unit_id: str, method: str, forced_member: Optional[str] = None):
        task_rows = sequence[sel & ~sequence["task_id"].isin(assigned_task_ids)].copy()
        if task_rows.empty:
            return
        member_id = clean_id(forced_member) if forced_member else choose_member_for_unit(unit_id, unit_members, analyst_load)
        for _, task in task_rows.iterrows():
            assignments.append(
                {
                    "task_id": task["task_id"],
                    "assigned_analyst_id": member_id,
                    "assigned_unit": unit_id,
                    "method": method,
                }
            )
            assigned_task_ids.add(task["task_id"])
            analyst_load[member_id] = analyst_load.get(member_id, 0.0) + float(task["effort"])
        unit_idx = units["unit_id"].eq(unit_id)
        units.loc[unit_idx, "assigned_cases"] += len(task_rows)
        units.loc[unit_idx, "remaining_capacity"] -= len(task_rows)

    available_analysts = set(analyst_capacity_frame(frames.analysts)[ah.analyst_id].apply(clean_id))

    if legacy_config.get("use_open_wip_owner", True):
        for owner_col, method in [
            ("wip_customer_analyst", "legacy_wip_customer"),
            ("wip_group_analyst", "legacy_wip_group"),
        ]:
            for owner in sequence[owner_col].dropna().apply(clean_id).unique():
                if owner and owner in available_analysts:
                    unit_id = analyst_to_unit.get(owner, owner)
                    sel = sequence[owner_col].apply(clean_id).eq(owner)
                    assign_tasks(sel, unit_id, method, forced_member=owner)

    if legacy_config.get("use_closed_customer_owner", False):
        for owner in sequence["last_customer_analyst"].dropna().apply(clean_id).unique():
            if owner and owner in available_analysts:
                sel = sequence["last_customer_analyst"].apply(clean_id).eq(owner)
                assign_tasks(sel, analyst_to_unit.get(owner, owner), "legacy_history_customer", forced_member=owner)

    group_keys = [key for key in sequence["legacy_group_key"].dropna().unique() if key]
    for group_key in group_keys:
        sel = sequence["legacy_group_key"].eq(group_key)
        sel = sel & ~sequence["task_id"].isin(assigned_task_ids)
        if not sel.any():
            continue
        n_cases = int(sel.sum())
        unit_id, cursor = assign_by_capacity(units, n_cases, cursor, rng)
        assign_tasks(sel, unit_id, "legacy_capacity_group")

    remaining = sequence[~sequence["task_id"].isin(assigned_task_ids)].copy()
    for _, task in remaining.iterrows():
        unit_id, cursor = assign_by_capacity(units, 1, cursor, rng)
        assign_tasks(sequence["task_id"].eq(task["task_id"]), unit_id, "legacy_capacity_single")

    result = pd.DataFrame(assignments)
    return result.sort_values("task_id").reset_index(drop=True)


def load_legacy_assignment_result(path: str, legacy_config: dict) -> pd.DataFrame:
    df = read_table(path, sheet_name=legacy_config.get("sheet_name"))
    task_col = first_existing_col(
        df,
        [
            legacy_config.get("task_id_col", ""),
            "task_id",
            "Task ID",
            "Customer Number",
            "Customer_ID_S1",
            "Customer_CIN_S1",
            "CIN",
        ],
    )
    analyst_col = first_existing_col(
        df,
        [
            legacy_config.get("analyst_id_col", ""),
            "assigned_analyst_id",
            "Assigned Analyst ID",
            "Case_Manager_Staff_ID_S1",
            "Case Manager Staff ID",
            "Staff ID",
        ],
    )
    unit_col = first_existing_col(
        df,
        [
            legacy_config.get("unit_col", ""),
            "assigned_unit",
            "Assign by TH",
            "CMTH",
            "Team",
            "Assigned Team",
        ],
    )
    method_col = first_existing_col(df, [legacy_config.get("method_col", ""), "Assigment Method", "Assignment Method"])
    if task_col is None:
        raise ValueError("Cannot locate a task/customer id column in legacy assignment file.")

    out = pd.DataFrame({"task_id": df[task_col].apply(clean_id)})
    if analyst_col:
        out["assigned_analyst_id"] = df[analyst_col].apply(clean_id)
    else:
        out["assigned_analyst_id"] = ""
    if unit_col:
        out["assigned_unit"] = df[unit_col].apply(norm_text)
    else:
        out["assigned_unit"] = out["assigned_analyst_id"]
    if method_col:
        out["method"] = df[method_col].fillna("").astype(str)
    else:
        out["method"] = "legacy_imported"
    return out


def history_hit(assigned: str, owners: Sequence[str]) -> bool:
    assigned = clean_id(assigned)
    return bool(assigned and assigned in {clean_id(x) for x in owners})


def add_metric_flags(assignments: pd.DataFrame, features: pd.DataFrame, method_name: str) -> pd.DataFrame:
    detail = features.merge(assignments, how="left", on="task_id")
    detail["assignment_method_name"] = method_name
    detail["assigned_analyst_id"] = detail["assigned_analyst_id"].fillna("").apply(clean_id)
    detail["assigned_unit"] = detail["assigned_unit"].fillna("")
    detail["same_customer"] = (
        detail["assigned_analyst_id"].eq(detail["last_customer_analyst"]) & detail["last_customer_analyst"].ne("")
    )
    detail["same_open_customer"] = (
        detail["assigned_analyst_id"].eq(detail["wip_customer_analyst"]) & detail["wip_customer_analyst"].ne("")
    )
    detail["same_mg"] = detail.apply(lambda row: history_hit(row["assigned_analyst_id"], row["mg_history_analysts"]), axis=1)
    detail["same_rm"] = detail.apply(lambda row: history_hit(row["assigned_analyst_id"], row["rm_history_analysts"]), axis=1)
    detail["same_segment"] = detail.apply(
        lambda row: history_hit(row["assigned_analyst_id"], row["segment_history_analysts"]), axis=1
    )
    return detail


def rate_for_eligible(df: pd.DataFrame, flag_col: str, eligible_col: str) -> float:
    eligible = df[eligible_col]
    return safe_ratio(df.loc[eligible, flag_col].sum(), eligible.sum())


def split_rate(df: pd.DataFrame, group_col: str) -> float:
    groups = df[(df[group_col] != "") & (df["assigned_analyst_id"] != "")]
    if groups.empty:
        return np.nan
    counts = groups.groupby(group_col).agg(cases=("task_id", "count"), assignees=("assigned_analyst_id", "nunique"))
    counts = counts[counts["cases"] > 1]
    if counts.empty:
        return 0.0
    return safe_ratio((counts["assignees"] > 1).sum(), len(counts))


def summarize_method(detail: pd.DataFrame, analyst_df: pd.DataFrame) -> Tuple[Dict[str, float], pd.DataFrame]:
    analysts = analyst_capacity_frame(analyst_df)
    assigned = detail[detail["assigned_analyst_id"] != ""].copy()
    assigned["effort"] = pd.to_numeric(assigned["effort"], errors="coerce").fillna(1.0)
    load = assigned.groupby("assigned_analyst_id", as_index=False).agg(
        assigned_cases=("task_id", "count"),
        assigned_effort=("effort", "sum"),
    )
    load = analysts[[ah.analyst_id, ah.target_wip, ah.current_wip]].merge(
        load, how="left", left_on=ah.analyst_id, right_on="assigned_analyst_id"
    )
    load["assigned_cases"] = load["assigned_cases"].fillna(0)
    load["assigned_effort"] = load["assigned_effort"].fillna(0.0)
    load["wip_after"] = load[ah.current_wip] + load["assigned_effort"]
    load["over_target_effort"] = (load["wip_after"] - load[ah.target_wip]).clip(lower=0)
    load["target_share"] = load[ah.target_wip] / load[ah.target_wip].sum()
    total_effort = load["assigned_effort"].sum()
    load["actual_share"] = np.where(total_effort > 0, load["assigned_effort"] / total_effort, 0)
    load["target_share_error"] = (load["actual_share"] - load["target_share"]).abs()

    mg_eligible = detail["mg_history_analysts"].apply(lambda x: len(x) > 0)
    rm_eligible = detail["rm_history_analysts"].apply(lambda x: len(x) > 0)
    seg_eligible = detail["segment_history_analysts"].apply(lambda x: len(x) > 0)
    customer_eligible = detail["last_customer_analyst"].ne("")

    load_ratio = np.where(
        load[ah.target_wip].replace(0, np.nan).notna(),
        load["wip_after"] / load[ah.target_wip].replace(0, np.nan),
        np.nan,
    )
    metrics = {
        "assigned_cases": float(len(assigned)),
        "assigned_effort": float(total_effort),
        "same_customer_rate": rate_for_eligible(detail.assign(_eligible=customer_eligible), "same_customer", "_eligible"),
        "same_mg_rate": rate_for_eligible(detail.assign(_eligible=mg_eligible), "same_mg", "_eligible"),
        "same_rm_rate": rate_for_eligible(detail.assign(_eligible=rm_eligible), "same_rm", "_eligible"),
        "same_segment_rate": rate_for_eligible(detail.assign(_eligible=seg_eligible), "same_segment", "_eligible"),
        "mg_split_group_rate": split_rate(detail, "group_key_mg"),
        "rm_split_group_rate": split_rate(detail, "group_key_rm"),
        "analyst_load_cv": safe_ratio(np.nanstd(load_ratio), np.nanmean(load_ratio)),
        "analyst_load_gini": gini(load["assigned_effort"]),
        "target_share_abs_error": float(load["target_share_error"].sum() / 2),
        "over_target_analysts": float((load["over_target_effort"] > 0).sum()),
        "over_target_effort": float(load["over_target_effort"].sum()),
        "unique_assignees": float(assigned["assigned_analyst_id"].nunique()),
        "unassigned_cases": float(detail["assigned_analyst_id"].eq("").sum()),
    }
    return metrics, load


def compare_methods(
    features: pd.DataFrame,
    analyst_df: pd.DataFrame,
    proposed_assignments: pd.DataFrame,
    legacy_assignments: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    proposed_detail = add_metric_flags(proposed_assignments, features, "proposed")
    legacy_detail = add_metric_flags(legacy_assignments, features, "legacy")
    proposed_metrics, proposed_load = summarize_method(proposed_detail, analyst_df)
    legacy_metrics, legacy_load = summarize_method(legacy_detail, analyst_df)

    rows = []
    for metric in METRIC_NAMES:
        proposed_value = proposed_metrics.get(metric, np.nan)
        legacy_value = legacy_metrics.get(metric, np.nan)
        rows.append(
            {
                "metric": metric,
                "proposed": proposed_value,
                "legacy": legacy_value,
                "delta_proposed_minus_legacy": proposed_value - legacy_value
                if pd.notna(proposed_value) and pd.notna(legacy_value)
                else np.nan,
            }
        )
    summary = pd.DataFrame(rows)

    detail = proposed_detail.merge(
        legacy_detail[
            [
                "task_id",
                "assigned_analyst_id",
                "assigned_unit",
                "method",
                "same_customer",
                "same_mg",
                "same_rm",
                "same_segment",
            ]
        ].rename(
            columns={
                "assigned_analyst_id": "legacy_assigned_analyst_id",
                "assigned_unit": "legacy_assigned_unit",
                "method": "legacy_method",
                "same_customer": "legacy_same_customer",
                "same_mg": "legacy_same_mg",
                "same_rm": "legacy_same_rm",
                "same_segment": "legacy_same_segment",
            }
        ),
        how="left",
        on="task_id",
    ).rename(
        columns={
            "assigned_analyst_id": "proposed_assigned_analyst_id",
            "assigned_unit": "proposed_assigned_unit",
            "method": "proposed_method",
            "same_customer": "proposed_same_customer",
            "same_mg": "proposed_same_mg",
            "same_rm": "proposed_same_rm",
            "same_segment": "proposed_same_segment",
        }
    )
    detail["same_assigned_analyst"] = detail["proposed_assigned_analyst_id"].eq(
        detail["legacy_assigned_analyst_id"].fillna("")
    )

    proposed_load = proposed_load.add_prefix("proposed_")
    legacy_load = legacy_load.add_prefix("legacy_")
    load = proposed_load.merge(
        legacy_load,
        how="outer",
        left_on="proposed_" + ah.analyst_id,
        right_on="legacy_" + ah.analyst_id,
    )

    group_rows = []
    for method_name, method_detail in [("proposed", proposed_detail), ("legacy", legacy_detail)]:
        for group_type, group_col in [("mg", "group_key_mg"), ("rm", "group_key_rm")]:
            groups = method_detail[(method_detail[group_col] != "") & (method_detail["assigned_analyst_id"] != "")]
            if groups.empty:
                continue
            group_stat = groups.groupby(group_col).agg(
                cases=("task_id", "count"),
                distinct_assignees=("assigned_analyst_id", "nunique"),
                assignees=("assigned_analyst_id", lambda s: ",".join(sorted(set(s)))),
            )
            group_stat = group_stat.reset_index().rename(columns={group_col: "group_key"})
            group_stat["group_type"] = group_type
            group_stat["method"] = method_name
            group_rows.append(group_stat)
    group_fragmentation = pd.concat(group_rows, ignore_index=True) if group_rows else pd.DataFrame()
    return summary, detail, load, group_fragmentation


def paired_bootstrap_deltas(
    detail: pd.DataFrame, analyst_df: pd.DataFrame, iterations: int, seed: int
) -> pd.DataFrame:
    if iterations <= 0 or detail.empty:
        return pd.DataFrame()
    rng = np.random.default_rng(seed)
    rows = []
    for i in range(iterations):
        sample_idx = rng.integers(0, len(detail), len(detail))
        sample = detail.iloc[sample_idx].copy()
        sample["task_id"] = [f"{task_id}__boot{i}_{j}" for j, task_id in enumerate(sample["task_id"])]

        feature_cols = [
            "task_id",
            "customer_id",
            "due_date",
            "effort",
            "init_date",
            "mg_num",
            "rm_num",
            "segment",
            "last_customer_analyst",
            "wip_customer_analyst",
            "mg_history_analysts",
            "rm_history_analysts",
            "segment_history_analysts",
            "wip_group_analyst",
            "group_key_mg",
            "group_key_rm",
            "legacy_group_key",
        ]
        features = sample[feature_cols].copy()
        proposed_assignments = sample[
            ["task_id", "proposed_assigned_analyst_id", "proposed_assigned_unit", "proposed_method"]
        ].rename(
            columns={
                "proposed_assigned_analyst_id": "assigned_analyst_id",
                "proposed_assigned_unit": "assigned_unit",
                "proposed_method": "method",
            }
        )
        legacy_assignments = sample[
            ["task_id", "legacy_assigned_analyst_id", "legacy_assigned_unit", "legacy_method"]
        ].rename(
            columns={
                "legacy_assigned_analyst_id": "assigned_analyst_id",
                "legacy_assigned_unit": "assigned_unit",
                "legacy_method": "method",
            }
        )
        summary, _, _, _ = compare_methods(features, analyst_df, proposed_assignments, legacy_assignments)
        for _, row in summary.iterrows():
            rows.append({"iteration": i, "metric": row["metric"], "delta": row["delta_proposed_minus_legacy"]})

    boot = pd.DataFrame(rows)
    if boot.empty:
        return boot
    return boot.groupby("metric")["delta"].agg(
        mean_delta="mean",
        std_delta="std",
        p05=lambda s: s.quantile(0.05),
        p50=lambda s: s.quantile(0.50),
        p95=lambda s: s.quantile(0.95),
    ).reset_index()


def legacy_seed_sensitivity(
    frames: InputFrames,
    features: pd.DataFrame,
    proposed_assignments: pd.DataFrame,
    legacy_config: dict,
    iterations: int,
    seed: int,
) -> pd.DataFrame:
    if iterations <= 0:
        return pd.DataFrame()
    rows = []
    for i in range(iterations):
        run_seed = seed + i
        legacy_assignments = run_legacy_simulator(frames, features, legacy_config, seed=run_seed)
        summary, _, _, _ = compare_methods(features, frames.analysts, proposed_assignments, legacy_assignments)
        summary["iteration"] = i
        summary["seed"] = run_seed
        rows.append(summary)
    if not rows:
        return pd.DataFrame()
    all_runs = pd.concat(rows, ignore_index=True)
    return all_runs.groupby("metric")["delta_proposed_minus_legacy"].agg(
        mean_delta="mean",
        std_delta="std",
        min_delta="min",
        p05=lambda s: s.quantile(0.05),
        p50=lambda s: s.quantile(0.50),
        p95=lambda s: s.quantile(0.95),
        max_delta="max",
    ).reset_index()


def write_report(
    output_path: str,
    summary: pd.DataFrame,
    detail: pd.DataFrame,
    load: pd.DataFrame,
    group_fragmentation: pd.DataFrame,
    bootstrap: pd.DataFrame,
    legacy_sensitivity: pd.DataFrame,
):
    out_path = Path(output_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path) as writer:
        summary.to_excel(writer, sheet_name="metric_summary", index=False)
        detail.to_excel(writer, sheet_name="case_detail", index=False)
        load.to_excel(writer, sheet_name="analyst_load", index=False)
        if not group_fragmentation.empty:
            group_fragmentation.to_excel(writer, sheet_name="group_fragmentation", index=False)
        if not bootstrap.empty:
            bootstrap.to_excel(writer, sheet_name="paired_bootstrap_delta", index=False)
        if not legacy_sensitivity.empty:
            legacy_sensitivity.to_excel(writer, sheet_name="legacy_seed_sensitivity", index=False)


def run_from_config(config_path: str, market_name: Optional[str] = None):
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    current_date = config.get("current_date")
    if not current_date:
        raise ValueError("current_date is required in the comparison config.")

    markets = config.get("markets")
    if markets:
        if market_name is None:
            if len(markets) != 1:
                raise ValueError("Please pass --market when the config contains more than one market.")
            market_name = next(iter(markets.keys()))
        market_config = markets[market_name]
    else:
        market_config = config
        market_name = market_name or config.get("MARKET", "market")

    frames = load_input_frames(market_config, current_date)
    features = build_task_features(frames.tasks, frames.close, frames.open)

    proposed_config = config.get("proposed_config", {})
    proposed_config.update({k: v for k, v in market_config.items() if k in DEFAULT_SCORE_CONFIG})
    proposed_assignments = run_proposed(frames, current_date, proposed_config, top_n=config.get("top_n", 3))

    legacy_config = config.get("legacy_simulation", {})
    legacy_result_file = market_config.get("LEGACY_ASSIGNMENT_RESULT_FILE") or legacy_config.get("assignment_result_file")
    if legacy_result_file:
        legacy_assignments = load_legacy_assignment_result(legacy_result_file, legacy_config)
    else:
        legacy_assignments = run_legacy_simulator(
            frames,
            features,
            legacy_config,
            seed=int(legacy_config.get("seed", config.get("seed", 42))),
        )

    summary, detail, load, group_fragmentation = compare_methods(
        features,
        frames.analysts,
        proposed_assignments,
        legacy_assignments,
    )

    bootstrap = paired_bootstrap_deltas(
        detail,
        frames.analysts,
        iterations=int(config.get("bootstrap_iterations", 0)),
        seed=int(config.get("seed", 42)),
    )
    legacy_sensitivity = pd.DataFrame()
    if not legacy_result_file:
        legacy_sensitivity = legacy_seed_sensitivity(
            frames,
            features,
            proposed_assignments,
            legacy_config,
            iterations=int(config.get("legacy_seed_iterations", 0)),
            seed=int(config.get("seed", 42)) + 1000,
        )

    output_dir = Path(config.get("output_dir", "comparison_outputs"))
    output_name = config.get("output_file", f"assignment_comparison_{market_name}_{current_date}.xlsx")
    output_path = output_dir / output_name
    write_report(
        str(output_path),
        summary,
        detail,
        load,
        group_fragmentation,
        bootstrap,
        legacy_sensitivity,
    )
    print(f"Comparison report written to: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Compare proposed CDD assignment with legacy assignment logic.")
    parser.add_argument("--config", required=True, help="Path to comparison_config.json.")
    parser.add_argument("--market", default=None, help="Market key if the config contains multiple markets.")
    args = parser.parse_args()
    run_from_config(args.config, args.market)


if __name__ == "__main__":
    main()
```
