import pandas as pd
from header import AnalystHeader as ah
from header import TaskHeader as th
from datetime import datetime, timedelta
from collections import defaultdict


class Analyst:
    def __init__(self, row: pd.Series):
        # 使用 getattr 获取 header 中配置的列名，再从 row 中取值
        self.analyst_id = str(row.get(ah.analyst_id))
        self.target_wip = float(row.get(ah.target_wip, 10))
        self.current_wip = float(row.get(ah.current_wip, 0))
        self._ori_wip = float(row.get(ah.current_wip, 0))
        self.weekly_throughput = float(row.get(ah.daily_productivity, 1.0)) * 5
        self.team_head = row.get(ah.team_head, '')
        self.analyst_name = row.get(ah.analyst_name, '')

        self.history_customers = set()
        self.history_rms = set()
        self.history_mgs = set()
        self.wip_mgs = set()
        self.batch_mgs = set()
        self.history_segment = set()

        self.history_customers_reviews = defaultdict(list)
        self.history_rms_reviews = defaultdict(list)
        self.history_mgs_reviews = defaultdict(list)
        self.wip_mgs_reviews = defaultdict(list)
        self.batch_mgs_reviews = defaultdict(list)
        self.history_segment_reviews = defaultdict(list)

    def reset_wip(self):
        self.current_wip = self._ori_wip
        self.batch_mgs = set()

    def __repr__(self):
        return f"<Analyst {self.analyst_id}>"


class CDDTask:
    def __init__(self, row: pd.Series):
        self.task_id = str(row.get(th.task_id))
        self.customer_id = str(row.get(th.customer_id))
        self.effort = float(row.get(th.effort, 1.0))
        self.rm_num = ""
        self.mg_num = ""
        self.segment = ""
        self.history_cm = set()

        # 处理日期格式
        raw_date = row.get(th.due_date)
        if isinstance(raw_date, pd.Timestamp):
            self.due_date = raw_date.to_pydatetime()
        elif isinstance(raw_date, str):
            try:
                self.due_date = datetime.strptime(raw_date, "%Y-%m-%d")
            except ValueError:
                self.due_date = datetime.strptime(raw_date, "%Y/%m/%d")  # 兼容其他格式
        else:
            self.due_date = raw_date

        self.init_date = self.due_date - timedelta(days=120)

    def __repr__(self):
        return f"<Task {self.task_id}>"