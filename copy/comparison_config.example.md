# comparison_config.example.json

```json
{
  "current_date": "2026-08-13",
  "output_dir": "comparison_outputs",
  "bootstrap_iterations": 500,
  "legacy_seed_iterations": 100,
  "seed": 42,
  "top_n": 3,
  "proposed_config": {
    "cus_affinity_weight": 100.0,
    "mg_affinity_weight": 80.0,
    "rm_affinity_weight": 50.0,
    "segment_affinity_weight": 100.0,
    "capacity_weight": 5.0,
    "overload_penalty": 5.0,
    "simulate_decay": true
  },
  "legacy_simulation": {
    "assignment_unit_col": "Staff ID",
    "use_open_wip_owner": true,
    "use_closed_customer_owner": false,
    "seed": 42,
    "skip_units": [],
    "task_id_col": "Customer_ID_S1",
    "analyst_id_col": "Case_Manager_Staff_ID_S1",
    "unit_col": "Assign by TH",
    "method_col": "Assigment Method",
    "sheet_name": "CMAssignmentFullList"
  },
  "markets": {
    "AMH": {
      "MARKET": "Hong Kong",
      "LEGAL_ENTITY": ["HBAP"],
      "OE_REPORT_FILES": ["path/to/oe_report.csv"],
      "CLOSE_REPORT_FILES": ["path/to/closed_review.csv"],
      "OPEN_REPORT_FILES": "path/to/open_review.csv",
      "RAM_FILES": ["path/to/ram.csv"],
      "MG_FILES": ["path/to/master_group.xlsx"],
      "CASE_MANAGER_LIST_FILES": "path/to/cm_list.xlsx",
      "LEGACY_ASSIGNMENT_RESULT_FILE": ""
    }
  }
}
```
