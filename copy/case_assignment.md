# case_assignment.py

```python
import logging
import os.path

import pandas as pd
import math
from datetime import datetime, timedelta
from header import AnalystHeader as ah
from header import TaskHeader as th
from header import ClosedReviewHeader as ch
from header import OpenReviewHeader as oh
from header import MasterGroupHeader as mh
from header import StageType as st
from header import ReviewStatusType as rst
from objects import Analyst, CDDTask
from typing import List, DefaultDict
from collections import defaultdict


class CDDTaskAllocator:
    def __init__(self,
                 df_analysts: pd.DataFrame,
                 df_tasks: pd.DataFrame,
                 df_close: pd.DataFrame,
                 df_open: pd.DataFrame,
                 current_date: str = None,
                 config: dict = None):
        """
        初始化分配器
        :param df_analysts: 分析师数据的 DataFrame
        :param df_tasks: 任务数据的 DataFrame
        :param analyst_header: 分析师列名配置对象
        :param task_header: 任务列名配置对象
        :param config: 算法参数配置
        """
        self.df_analysts = df_analysts
        self.df_tasks = df_tasks
        self.df_close = df_close
        self.df_open = df_open
        # 数据转换：DataFrame Row -> Object
        self.analysts: DefaultDict[str: Analyst] = defaultdict(Analyst)
        self.tasks: DefaultDict[str: CDDTask] = defaultdict(CDDTask)
        self.selected_analyst: DefaultDict[str: list] = defaultdict(list)

        # 默认算法参数
        self.config = {
            'cus_affinity_weight': 100.0,
            'mg_affinity_weight': 80.0,
            'rm_affinity_weight': 50.0,
            'segment_affinity_weight': 100.0,
            'capacity_weight': 5.0,
            'overload_penalty': 5.0,
            'simulate_decay': True
        }

        if config:
            self.config.update(config)

        if current_date is None:
            self.current_date = datetime.now()
        else:
            self.current_date = pd.to_datetime(current_date)

        self.res_df = pd.DataFrame()
        self.out_dir = None
        self.analyst_out_dir = None
        self.cm_statistics = dict()

    def read_data(self):
        self.analysts = {str(row[ah.analyst_id]): Analyst(row) for _, row in self.df_analysts.iterrows()}
        self.tasks = {str(row[th.task_id]): CDDTask(row) for _, row in self.df_tasks.iterrows()}

    def add_wip_segment_to_analyst(self):
        open_df = self.df_open.copy()
        open_df[oh.CustomerNumber] = open_df[
            oh.CustomerNumber].astype(str).apply(lambda x: x.lstrip('0'))
        open_df = open_df.sort_values(by=[ch.InitiatedDate])

        for analyst, sub_df in open_df.groupby(oh.StaffID):
            segments = sub_df[ch.Segment].dropna().iloc[-10:].to_list()
            if analyst in self.analysts:
                self.analysts[analyst].history_segment = self.analysts[analyst].history_segment.union(
                    set(segments))

    def add_history_customer_to_analyst(self):
        history_close_df = self.df_close.copy()
        history_close_df[ch.CustomerNumber] = history_close_df[
            ch.CustomerNumber].astype(str).apply(lambda x: x.lstrip('0'))
        history_close_df[ch.LatestDcFinalisedById] = history_close_df[ch.LatestDcFinalisedById].astype(str)
        history_close_df[ch.RmNum] = history_close_df[ch.RmNum].astype(str)
        completed_history_close_df = history_close_df[
            history_close_df[ch.ReviewStatus] == "Approval Completed"
        ]
        completed_history_close_df = completed_history_close_df.sort_values(by=[ch.InitiatedDate])
        for analyst, sub_df in completed_history_close_df.groupby(ch.LatestDcFinalisedById):
            customers = sub_df[ch.CustomerNumber].drop_duplicates().to_list()
            rms = sub_df[ch.RmNum]
            rms = rms[~rms.isin({'-', ''})].drop_duplicates().to_list()
            mgs = sub_df[mh.MasterGroupCode]
            mgs = mgs[~mgs.isin({'-', ''})].drop_duplicates().to_list()

            if analyst in self.analysts:
                self.analysts[analyst].history_customers = self.analysts[analyst].history_customers.union(
                    set(customers))
                self.analysts[analyst].history_rms = self.analysts[analyst].history_rms.union(
                    set(rms))
                self.analysts[analyst].history_mgs = self.analysts[analyst].history_mgs.union(
                    set(mgs))
            for customer in customers:
                if customer in self.tasks:
                    self.tasks[customer].history_cm.add(analyst)

        last_review = completed_history_close_df.drop_duplicates(subset=[ch.CustomerNumber], keep='last')
        last_review_with_rm = last_review[~last_review[ch.RmNum].isin({'-', ''})]
        for idx, row in last_review_with_rm.iterrows():
            customer, rm = row[ch.CustomerNumber], row[ch.RmNum]
            if customer in self.tasks:
                self.tasks[customer].rm_num = rm

        last_review_with_mg = last_review[~last_review[mh.MasterGroupCode].isin({'-', ''})]
        for idx, row in last_review_with_mg.iterrows():
            customer, mg = row[ch.CustomerNumber], row[mh.MasterGroupCode]
            if customer in self.tasks:
                self.tasks[customer].mg_num = mg

        last_review_with_segment = last_review[~last_review[ch.Segment].isin({'-', ''})]
        for idx, row in last_review_with_segment.iterrows():
            customer, segment = row[ch.CustomerNumber], row[ch.Segment]
            if customer in self.tasks:
                self.tasks[customer].segment = segment

        self.add_wip_segment_to_analyst()

    def _calculate_score(self, analyst, task):
        """计算匹配分数 (逻辑保持不变)"""
        reasons = {}

        # 1. 历史经验分 (Affinity)
        has_cust_history = 1 if task.customer_id in analyst.history_customers else 0
        affinity_score = has_cust_history * self.config['cus_affinity_weight']
        reasons['customer_affinity'] = affinity_score

        has_segment_history = 1 if task.segment in analyst.history_segment else 0
        segment_score = has_segment_history * self.config['segment_affinity_weight']
        reasons['segment_affinity'] = segment_score

        has_mg_history = 1 if (
                task.mg_num != "" and task.mg_num in analyst.history_mgs and not has_cust_history) else 0
        mg_affinity_score = has_mg_history * self.config['mg_affinity_weight']
        reasons['mg_affinity'] = mg_affinity_score

        has_rm_history = 1 if (
                task.rm_num != "" and task.rm_num in analyst.history_rms and not has_cust_history and not has_mg_history
        ) else 0
        rm_affinity_score = has_rm_history * self.config['rm_affinity_weight']
        reasons['rm_affinity'] = rm_affinity_score

        # 2. 容量均衡分 (Capacity)
        gap = analyst.target_wip - analyst.current_wip
        if gap < 0:
            capacity_score = gap * self.config['capacity_weight'] * self.config['overload_penalty']
        else:
            capacity_score = gap * self.config['capacity_weight']
            reasons['customer_affinity'] *= 10
            reasons['mg_affinity'] *= 10
            reasons['rm_affinity'] *= 10
            reasons['segment_affinity'] *= 10
        reasons['capacity'] = capacity_score

        score = sum(v for k, v in reasons.items())
        return score, reasons

    def _simulate_wip_release(self, current_week_start, last_week_start):
        """模拟时间推移释放产能"""
        if not self.config['simulate_decay'] or last_week_start is None:
            return

        weeks_passed = (current_week_start - last_week_start).days // 7
        if weeks_passed > 0:
            # print(f"--- Simulating {weeks_passed} weeks passing ---")
            for analyst_id, analyst in self.analysts.items():
                released = weeks_passed * analyst.weekly_throughput
                analyst.current_wip = max(0, analyst.current_wip - released)

    def run(self) -> pd.DataFrame:
        """运行分配算法并返回 DataFrame 格式的结果"""
        # 1. 按 Due Date 排序任务
        sorted_tasks = sorted(list(self.tasks.values()), key=lambda x: x.init_date)
        assignments = list()
        if not sorted_tasks:
            return pd.DataFrame()

        last_week_start = sorted_tasks[0].init_date - timedelta(days=sorted_tasks[0].init_date.weekday())

        for task in sorted_tasks:
            # 时间推进检查
            current_week_start = task.init_date - timedelta(days=task.init_date.weekday())
            if current_week_start > last_week_start:
                self._simulate_wip_release(current_week_start, last_week_start)
                last_week_start = current_week_start

            # 寻找最佳匹配
            best_analyst = None
            best_score = -math.inf
            best_reasons = {}
            task_reasons = {}

            for analyst_id, analyst in self.analysts.items():
                if analyst_id in self.selected_analyst.get(task.task_id, []):
                    continue
                score, reasons = self._calculate_score(analyst, task)
                task_reasons[analyst_id] = reasons
                # 贪心策略：分数高优先；分数相同选当前活最少的
                if score > best_score:
                    best_score = score
                    best_analyst = analyst
                    best_reasons = reasons
                elif score == best_score:
                    if analyst.current_wip < best_analyst.current_wip:
                        best_analyst = analyst
                        best_reasons = reasons

            # 记录结果
            if best_analyst:
                assignments.append({
                    "Task ID": task.task_id,
                    "Customer": task.customer_id,
                    "Init Date": task.init_date,
                    "Due Date": task.due_date,
                    "Assigned Analyst ID": best_analyst.analyst_id,
                    "Score Total": best_score,
                    "Score Customer Affinity": best_reasons.get('customer_affinity'),
                    "Score MG Affinity": best_reasons.get('mg_affinity'),
                    "Score RM Affinity": best_reasons.get('rm_affinity'),
                    "Score Capacity": best_reasons.get('capacity'),
                    "Analyst WIP After": best_analyst.current_wip + task.effort
                })
                # 更新状态
                best_analyst.current_wip += task.effort
                self.selected_analyst[task.task_id].append(best_analyst.analyst_id)
                # best_analyst.history_customers.add(task.customer_id)
                if task.rm_num != "":
                    best_analyst.history_rms.add(task.rm_num)

        res_df = pd.DataFrame(assignments)
        self.res_df = res_df
        return res_df

    def run_top_n_analyst(self, n=3) -> pd.DataFrame:
        """运行分配算法并返回 DataFrame 格式的结果"""
        # 1. 按 Due Date 排序任务
        sorted_tasks = sorted(list(self.tasks.values()), key=lambda x: x.init_date)
        assignments = list()
        if not sorted_tasks:
            return pd.DataFrame()

        last_week_start = sorted_tasks[0].init_date - timedelta(days=sorted_tasks[0].init_date.weekday())

        for task in sorted_tasks:
            # 时间推进检查
            current_week_start = task.init_date - timedelta(days=task.init_date.weekday())
            if current_week_start > last_week_start:
                self._simulate_wip_release(current_week_start, last_week_start)
                last_week_start = current_week_start

            # 寻找最佳匹配
            analyst_reason = {}
            analyst_score = {}

            for analyst_id, analyst in self.analysts.items():
                if analyst_id in self.selected_analyst.get(task.task_id, []):
                    continue
                score, reasons = self._calculate_score(analyst, task)
                analyst_reason[analyst_id] = reasons
                analyst_score[analyst_id] = score
            analyst_rank = sorted(
                analyst_score, key=lambda x: (analyst_score[x], -1 * self.analysts[x].current_wip), reverse=True
            )
            selected_analyst = analyst_rank[:n]
            task_assign_info = {
                "Task ID": task.task_id,
                "Customer": task.customer_id,
                "Init Date": task.init_date,
                "Due Date": task.due_date,
                "Segment": task.segment,
            }
            for i, analyst_id in enumerate(selected_analyst):
                task_assign_info.update(
                    {
                        f"Assigned Analyst ID {i+1}": analyst_id,
                        f"Score Total {i+1}": analyst_score[analyst_id],
                        f"Score Customer Affinity {i+1}": analyst_reason[analyst_id].get('customer_affinity'),
                        f"Score MG Affinity {i+1}": analyst_reason[analyst_id].get('mg_affinity'),
                        f"Score RM Affinity {i+1}": analyst_reason[analyst_id].get('rm_affinity'),
                        f"Score Segment Affinity {i + 1}": analyst_reason[analyst_id].get('segment_affinity'),
                        f"Score Capacity {i+1}": analyst_reason[analyst_id].get('capacity'),
                        f"Analyst WIP After {i+1}": self.analysts[analyst_id].current_wip + task.effort
                    }
                )
                if i == 0:
                    best_analyst = self.analysts[analyst_id]
                    best_analyst.current_wip += task.effort
                    self.selected_analyst[task.task_id].append(best_analyst.analyst_id)
                    # best_analyst.history_customers.add(task.customer_id)
                    if task.rm_num != "":
                        best_analyst.history_rms.add(task.rm_num)

            assignments.append(task_assign_info)
        res_df = pd.DataFrame(assignments)
        self.res_df = res_df

        return res_df

    def init_for_rerun(self):
        for analyst in self.analysts.values():
            analyst.reset_wip()

    def run_multiple_times(self, run_times=3):
        res_dict = dict()
        for i in range(run_times):
            if i > 0:
                self.init_for_rerun()
            res = self.run()
            res_dict[i] = res

        res_col = [
            'Assigned Analyst ID', 'Score Total', 'Score Customer Affinity', 'Score MG Affinity', 'Score RM Affinity',
            'Score Capacity', 'Analyst WIP After'
        ]
        all_df = list()
        for i, res_df in res_dict.items():
            if i == 0:
                res_df.columns = [f"{c} {i + 1}" if c in res_col else c for c in res_df.columns]
            else:
                res_df = res_df[res_col]
                res_df.columns = [f"{c} {i + 1}" for c in res_df.columns]
            all_df.append(res_df)

        multi_res_df = pd.concat(all_df, axis=1)
        self.res_df = multi_res_df

        return multi_res_df

    def output_result(self, output_file):
        self.calculated_same_cm_statistics()
        out_file = self.current_date.strftime(output_file)
        out_dir = "/".join(out_file.split("/")[:-1])
        if out_dir != "" and not os.path.isdir(out_dir):
            os.mkdir(out_dir)
            logging.info(f"new output dir {out_dir}")

        if out_file.split('.')[-1] == 'xlsx':
            self.res_df.to_excel(out_file, index=False)
        else:
            self.res_df.to_csv(out_file, index=False)
        self.out_dir = out_file


    def output_analyst_df(self, output_file):
        analyst_df = self.df_analysts
        out_file = self.current_date.strftime(output_file)
        out_dir = "/".join(out_file.split("/")[:-1])
        if out_dir != "" and not os.path.isdir(out_dir):
            os.mkdir(out_dir)
            logging.info(f"new output dir {out_dir}")

        if out_file.split('.')[-1] == 'xlsx':
            analyst_df.to_excel(out_file, index=False)
        else:
            analyst_df.to_csv(out_file, index=False)
        self.analyst_out_dir = out_file

    def calculated_same_cm_statistics(self):
        closed_df = self.df_close.copy()
        open_df = self.df_open.copy()

        open_df[oh.InitiatedDate] = pd.to_datetime(open_df[oh.InitiatedDate], format='mixed')
        closed_df[ch.InitiatedDate] = pd.to_datetime(closed_df[ch.InitiatedDate], format='mixed')
        closed_df = closed_df[closed_df[ch.ReviewStatus] == "Approval Completed"]
        combine_df = pd.concat([open_df, closed_df], axis=0)
        combine_df = combine_df[combine_df[ch.StaffID] != '']

        combine_df['Week'] = combine_df[ch.InitiatedDate].dt.to_period(freq='W')

        combine_df = combine_df.sort_values(by=ch.InitiatedDate)
        combine_df['Is Familiar Customer'] = combine_df.duplicated(
            subset=[ch.CustomerNumber, ch.StaffID], keep='first')

        combine_df['Is Familiar MG'] = combine_df.duplicated(
            subset=[mh.MasterGroupCode, ch.StaffID], keep='first')

        combine_df['Is Familiar RM'] = combine_df.duplicated(
            subset=[ch.RmNum, ch.StaffID], keep='first')

        combine_df['Is Familiar MG'] = (
                (combine_df['Is Familiar MG']) &
                (~combine_df['Is Familiar Customer']) &
                (~combine_df[mh.MasterGroupCode].eq(''))
        )

        combine_df['Is Familiar RM'] = (
                (combine_df['Is Familiar RM']) &
                (~combine_df['Is Familiar Customer']) &
                (~combine_df['Is Familiar MG']) &
                (~combine_df[ch.RmNum].eq('-'))
        )

        time_delta = {"6m": 180, "3m": 90, "1m": 30}
        stats = {"cm_handled_customer": {}, "cm_handled_mg": {}, "cm_handled_rm": {}, 'highest_same_cm': {}}
        hist_combine_df = combine_df[
            combine_df[ch.InitiatedDate] < self.current_date - pd.Timedelta(days=max(time_delta.values()))]
        same_customer_avl = hist_combine_df.groupby(
            ch.CustomerNumber).apply(lambda df: df[ch.StaffID].isin(self.analysts).sum() > 0)
        pr_df = combine_df[combine_df[ch.ReviewReason].isin({"Periodic"})]

        for m, td in time_delta.items():
            selected_hist_completed_df = pr_df[
                pr_df[ch.InitiatedDate].apply(
                    lambda x: self.current_date - pd.Timedelta(days=td) <= x <= self.current_date)
            ]
            selected_hist_completed_df = selected_hist_completed_df[
                ~selected_hist_completed_df[ch.ReviewStatus].eq('Cancelled')
            ]
            selected_hist_completed_df = selected_hist_completed_df[
                selected_hist_completed_df[ch.StaffID].isin(self.analysts)
            ]
            selected_hist_completed_df['Have Same CM'] = selected_hist_completed_df[ch.CustomerNumber].map(
                same_customer_avl).fillna(False)
            familiar_customer_num = selected_hist_completed_df['Is Familiar Customer'].sum()
            familiar_mg_num = selected_hist_completed_df["Is Familiar MG"].sum()
            familiar_rm_num = selected_hist_completed_df["Is Familiar RM"].sum()
            have_same_cm_num = selected_hist_completed_df['Have Same CM'].sum()

            cm_handled_customer = familiar_customer_num / len(selected_hist_completed_df)
            highest_same_cm = have_same_cm_num / len(selected_hist_completed_df)
            cm_handled_mg = familiar_mg_num / len(selected_hist_completed_df)
            cm_handled_rm = familiar_rm_num / len(selected_hist_completed_df)
            stats["cm_handled_customer"][m] = f"{round(cm_handled_customer * 100, 1)}%"
            stats["cm_handled_mg"][m] = f"{round(cm_handled_mg * 100, 1)}%"
            stats["cm_handled_rm"][m] = f"{round(cm_handled_rm * 100, 1)}%"
            stats['highest_same_cm'][m] = f"{round(highest_same_cm * 100, 1)}%"

        result = self.res_df.copy()
        result = result[result["Init Date"] <= self.current_date + pd.Timedelta(days=30)]
        familiar_customer_num = (result['Score Customer Affinity 1'] > 0).sum()
        familiar_mg_num = (result['Score MG Affinity 1'] > 0).sum()
        familiar_rm_num = (result['Score RM Affinity 1'] > 0).sum()
        cm_handled_customer = familiar_customer_num / len(result)
        cm_handled_mg = familiar_mg_num / len(result)
        cm_handled_rm = familiar_rm_num / len(result)
        stats["cm_handled_customer"]['current'] = f"{round(cm_handled_customer * 100, 1)}%"
        stats["cm_handled_mg"]['current'] = f"{round(cm_handled_mg * 100, 1)}%"
        stats["cm_handled_rm"]['current'] = f"{round(cm_handled_rm * 100, 1)}%"

        self.cm_statistics = stats
        return stats


if __name__ == "__main__":
    df_case_df = pd.read_excel("Hase Pending 2627 Cases.xlsx")
    df_case_df = df_case_df[
        (pd.to_datetime("2026-07-01") <= df_case_df[th.due_date])
    ]
    df_analyst_df = pd.read_excel("CM_List_HASE.xlsx")
    history_close_df = pd.read_excel("hase_closed_data.xlsx")
    self = CDDTaskAllocator(
        df_analysts=df_analyst_df,
        df_tasks=df_case_df,

    )
    self.read_data()
    self.add_history_customer_to_analyst(history_close_df)
    res = self.run_multiple_times()
    res["Week"] = res['Init_Date'].apply(lambda x: x - timedelta(days=x.weekday()))
    res.to_excel("Result_HASE.xlsx", index=False)

    # familiar_customer_num = (res['Score_Customer_Affinity_2'] > 0).sum()
    # familiar_rm_num = (res['Score_RM_Affinity_2'] > 0).sum()
    # other_num = len(res) - familiar_customer_num - familiar_rm_num
    # print(f"familiar_customer {familiar_customer_num}, familiar_rm {familiar_rm_num}, other {other_num}")
```
