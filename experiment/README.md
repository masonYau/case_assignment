# Assignment Experiment

这个目录里的代码用于和业务团队 output 做一版简单、可解释的实验比较。

核心口径：

- case list 从业务团队 output 的 `CMAssignmentFullList` 读取。
- 初始 WIP 从你的 `CM_LIST` 读取。
- CORP/CMB 的 MG 用 `horis_mg` 按对方口径 map。
- BBRM 的 group 用 `BBRM_IMIS_GROUP` 按对方口径 map。
- open/closed history 的“做过”口径尽量沿用你现有代码。
- 你的方法会直接调用 `CDDTaskAllocator` 重新分配同一批 case。

## Run

复制 `experiment_config.example.json`，填入公司环境文件路径和字段名。

```powershell
python experiment/run_experiment.py --config experiment/experiment_config.local.json
```

如果本机 `python` 不在 PATH，可以用公司环境里的 Python 解释器直接运行同一个脚本。

## Important Config

`files.business_output.orientation` 支持：

- `auto`: 默认。字段正常作为 column header 时直接读取；如果检测到字段在第一列，会自动转置。
- `records`: 每一行是一个 case。
- `fields_as_rows`: 第一列是字段名，每一列是一个 case。

如果业务 output 是截图里那种字段在第一列的形式，建议配置：

```json
"business_output": {
  "path": "...xlsx",
  "sheet_name": "CMAssignmentFullList",
  "header": null,
  "orientation": "fields_as_rows"
}
```

## Output Sheets

输出文件默认是 `experiment/assignment_experiment_result.xlsx`。

| Sheet | 内容 |
|---|---|
| `summary` | business/proposed 在 2y/5y 的 customer/group 命中率，以及总体 WIP 统计 |
| `case_detail` | 每个 case 的 business/proposed CM、MG/IMIS mapping、2y/5y 命中标记 |
| `wip_by_cm` | CM 粒度 WIP |
| `wip_by_team_cm` | Team + CM 粒度 WIP |
| `wip_by_team` | Team 汇总 WIP |
| `mapping_check` | business output 自带 MG/IMIS 与重新 mapping 的差异 |
| `warnings` | 运行时发现的口径/字段提醒 |
| `field_config_used` | 本次实际使用的字段配置 |

`debug=true` 时额外输出：

- `debug_history_sample`
- `debug_case_normalized`

## Metric Notes

`customer_hit_2y/5y`:

```text
assigned CM 在 lookback window 内做过同 customer
```

`mg_hit_2y/5y`:

```text
assigned CM 在 lookback window 内做过同 Master Group
```

`group_hit_2y/5y`:

```text
CORP/CMB 使用 MG；BBRM 默认使用 IMIS Group
```

`customer_or_group_hit_2y/5y`:

```text
customer_hit OR group_hit
```

WIP:

```text
wip_after_assignment = current_wip + assigned_workload
over_target_amount = max(wip_after_assignment - optimal_wip, 0)
```
