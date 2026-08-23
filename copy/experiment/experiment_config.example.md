# experiment/experiment_config.example.json

```json
{
  "as_of_date": "2026-08-13",
  "market": "Hong Kong",
  "legal_entity": "HBAP",
  "output_file": "assignment_experiment_result.xlsx",
  "debug": true,
  "lookback_years": [2, 5],
  "default_workload": 1.0,
  "top_n": 3,
  "closed_completed_only": true,
  "use_imis_as_group_for_bbrm": true,
  "exclude_mg_ids": [1514726659, 1514726667],
  "algorithm_config": {
    "cus_affinity_weight": 100.0,
    "mg_affinity_weight": 80.0,
    "rm_affinity_weight": 50.0,
    "segment_affinity_weight": 100.0,
    "capacity_weight": 5.0,
    "overload_penalty": 5.0,
    "simulate_decay": true
  },
  "files": {
    "business_output": {
      "path": "CM Assignment output.xlsx",
      "sheet_name": "CMAssignmentFullList",
      "header": 0,
      "orientation": "auto"
    },
    "cm_list": {
      "path": "CM_LIST.xlsx",
      "sheet_name": "Sheet1",
      "header": 0
    },
    "closed_report": [
      "closed_report.csv"
    ],
    "open_report": [
      {
        "path": "OpenReviews_CMB_20260720.csv",
        "header": null,
        "orientation": "fields_as_rows"
      }
    ],
    "horis_mg": {
      "path": "horis_mg.xlsx",
      "sheet_name": "Sheet1",
      "header": 0
    },
    "bb_rm_imis_group": {
      "path": "BBRM_IMIS_GROUP.xlsx",
      "sheet_name": "Sheet1",
      "header": 0
    }
  },
  "fields": {
    "case_list": {
      "case_id": "Customer_ID_S1",
      "customer": "CIN",
      "cin": "CIN",
      "segment": "CMB_Segment_S1",
      "due_date": "Planned T0 date",
      "workload": "Workload Units",
      "business_cm_id": "Case_Manager_Staff_ID_S1",
      "business_cm_name": "CM Name",
      "business_team": "Assign by TH",
      "mg_id": "Master Group ID",
      "mg_name": "Master Group",
      "imis_id": "IMIS Group No",
      "imis_name": "IMIS Group",
      "imis_mapping_key": "Customer_ID_S1"
    },
    "cm_list": {
      "cm_id": "Staff ID",
      "cm_name": "CM Name",
      "team": "Team",
      "current_wip": "Current WIP",
      "optimal_wip": "Optimal WIP",
      "productivity": "Last 3 month Productivity"
    },
    "horis_mg": {
      "cin": "CIN",
      "mg_id": "Master_Group_ID",
      "mg_name": "Master_Group_Name"
    },
    "bb_rm_imis_group": {
      "customer": "Real customer id",
      "imis_id": "IMIS Group No",
      "imis_name": "Group Name"
    },
    "closed_report": {
      "review_id": "Review ID",
      "customer": "Customer Number",
      "cm_id": "Latest DC Finalised by ID",
      "first_claimed_id": "First Claimed ID",
      "review_status": "Review Status",
      "completed_status": "Approval Completed",
      "cancelled_status": "Cancelled",
      "date_candidates": [
        "Latest DC Finalised Date",
        "Approval/Cancel Date",
        "Initiated Date"
      ],
      "rm_num": "RM Num",
      "segment": "Segment"
    },
    "open_report": {
      "review_id": "Review ID",
      "customer": "Customer Number",
      "stage": "Stage",
      "assigned_to_user": "Assigned to User",
      "user_id": "User ID",
      "latest_dc_by": "DC Finalized by ID",
      "cm_id": "Staff ID",
      "date_candidates": [
        "Initiated Date",
        "DC Finalized Date",
        "Date of latest action",
        "Stage Start Date"
      ],
      "rm_num": "RM Num",
      "segment": "Segment"
    }
  }
}
```
