import pandas as pd
from header import TaskHeader, OpenReviewHeader, RamHeader, ClosedReviewHeader, MasterGroupHeader
from header import ReviewStatusType as rst
from header import StageType as st
import os
import json
import datetime


class TaskClean:

    def __init__(
            self,
            df_open,
            df_close,
            df_task,
            ram_file=None,
            df_mg=None,
            market="",
            legal_entity="",
            current_date=None,
            init_lead_month=4
    ):
        self.df_open = df_open
        self.df_task = df_task
        self.df_close = df_close
        self.market = market
        self.legal_entity = legal_entity
        self.init_lead_month = init_lead_month
        if current_date is None:
            self.current_date = datetime.datetime.now()
        else:
            self.current_date = pd.to_datetime(current_date)

        self.df_ram = None
        if ram_file is not None:
            self.df_ram = ram_file

        self.df_mg = None
        if df_mg is not None:
            self.df_mg = df_mg

    def data_preprocess(self):
        df_open = self.df_open.copy()
        df_task = self.df_task.copy()
        oh = OpenReviewHeader
        th = TaskHeader
        mh = MasterGroupHeader

        df_open[oh.CustomerNumber] = df_open[oh.CustomerNumber].astype(str).apply(lambda x: x.lstrip('0'))
        df_open[oh.LegalEntity] = df_open[oh.LegalEntity].astype(str)
        df_open[oh.CNAndLE] = df_open.apply(lambda x: x[oh.CustomerNumber] + "&" + x[oh.LegalEntity], axis=1)
        df_open = df_open[
            (df_open[oh.Market] == self.market) & (df_open[oh.LegalEntity].isin(self.legal_entity))
        ]
        df_open[oh.LatestDcFinalisedDate] = pd.to_datetime(
            df_open[oh.LatestDcFinalisedDate], errors="coerce", format='mixed')
        df_open[oh.InitiatedDate] = pd.to_datetime(
            df_open[oh.InitiatedDate], errors="coerce", format='mixed')
        df_open[oh.StaffID] = df_open.apply(
            lambda r: r[oh.LatestDcFinalisedById] if r[oh.Stage] in (st.QC, st.APP) else
            r[oh.AssignedToUser] if r[oh.AssignedToUser].isdigit() else
            r[oh.UserID] if r[oh.UserID].isdigit() else
            "",
            axis=1
        )

        df_task[th.task_id] = df_task[th.task_id].astype(str).apply(lambda x: x.lstrip('0'))
        df_task[th.legal_entity] = df_task[th.legal_entity].astype(str)
        df_task[th.CNAndLE] = df_task.apply(lambda x: x[th.task_id] + "&" + x[th.legal_entity], axis=1)
        df_task = df_task[
            (df_task[th.market] == self.market) & (df_task[th.legal_entity].isin(self.legal_entity))
        ]
        df_task[th.due_date] = pd.to_datetime(df_task[th.due_date], errors="coerce")

        if self.df_ram is not None:
            rh = RamHeader
            df_ram = self.df_ram
            df_ram[rh.CustomerNumber] = df_ram[rh.CIN].astype(str).apply(lambda x: x.lstrip('0'))
            df_ram[rh.DueDate] = pd.to_datetime(df_ram[rh.NextCDDReviewDate], errors="coerce")
            df_ram[rh.LOB] = df_ram[rh.LineOfBusiness].fillna("")
            self.df_ram = df_ram

        if self.df_mg is not None:
            mh = MasterGroupHeader
            df_mg = self.df_mg
            df_mg[mh.CustomerNumber] = df_mg[mh.CustomerNumber].astype(str).apply(lambda x: x.lstrip('0'))
            df_mg[mh.MasterGroupCode] = df_mg[mh.MasterGroupCode].astype(str).apply(lambda x: x.lstrip('0'))
            task_mg = df_task[th.task_id].map(df_mg.set_index(df_mg[mh.CustomerNumber])[mh.MasterGroupCode])
            df_task[mh.MasterGroupCode] = task_mg.fillna('')

            open_mg = df_open[oh.CustomerNumber].map(df_mg.set_index(df_mg[mh.CustomerNumber])[mh.MasterGroupCode])
            df_open[mh.MasterGroupCode] = open_mg.fillna('')
        else:
            df_task[mh.MasterGroupCode] = pd.Series('', index=df_task.index)
            df_open[mh.MasterGroupCode] = pd.Series('', index=df_open.index)

        self.df_open = df_open
        self.df_task = df_task

        return

    def close_data_process(self):
        df_close = self.df_close.copy()
        ch = ClosedReviewHeader
        mh = MasterGroupHeader

        df_close[ch.CustomerNumber] = df_close[ch.CustomerNumber].astype(str).apply(lambda x: x.lstrip('0'))
        df_close[ch.ReviewID] = df_close[ch.ReviewID].astype(str)
        df_close[ch.LegalEntity] = df_close[ch.LegalEntity].astype(str)
        df_close[ch.LatestDcFinalisedById] = df_close[ch.LatestDcFinalisedById].astype(str)
        df_close[ch.LatestDcFinalisedByDate] = pd.to_datetime(
            df_close[ch.LatestDcFinalisedByDate], errors="coerce", format='mixed')
        df_close[ch.ApprovalCancelDate] = pd.to_datetime(
            df_close[ch.ApprovalCancelDate], errors="coerce", format='mixed')
        df_close[ch.InitiatedDate] = pd.to_datetime(
            df_close[ch.InitiatedDate], errors="coerce", format='mixed')
        df_close[ch.Segment] = df_close[ch.Segment].fillna('-')
        df_close[ch.StaffID] = df_close.apply(
            lambda r: r[ch.LatestDcFinalisedById] if r[ch.LatestDcFinalisedById].isdigit() else
            r[ch.FirstClaimedID] if r[ch.ReviewStatus] == rst.Cancelled and r[ch.FirstClaimedID].isdigit() else
            "",
            axis=1
        )
        df_close[ch.CNAndLE] = df_close.apply(lambda x: x[ch.CustomerNumber] + "&" + x[ch.LegalEntity], axis=1)
        df_close = df_close[
            (df_close[ch.Market] == self.market) & (df_close[ch.LegalEntity].isin(self.legal_entity))
        ]
        df_close = df_close.drop_duplicates(subset=[ch.ReviewID], keep="last")
        if self.df_mg is not None:
            df_mg = self.df_mg
            closed_mg = df_close[ch.CustomerNumber].map(df_mg.set_index(df_mg[mh.CustomerNumber])[mh.MasterGroupCode])
            df_close[mh.MasterGroupCode] = closed_mg.fillna('')

        else:
            df_close[mh.MasterGroupCode] = pd.Series('', index=df_close.index)
        self.df_close = df_close

        return

    def clean_task(self):
        df_open = self.df_open.copy()
        df_close = self.df_close.copy()
        df_task = self.df_task.copy()
        ch = ClosedReviewHeader
        oh = OpenReviewHeader
        th = TaskHeader
        current_month_start = self.current_date - pd.Timedelta(days=(self.current_date.day - 1))
        init_month = current_month_start + pd.Timedelta(days=31 * self.init_lead_month)


        df_task = df_task[~df_task[th.CNAndLE].isin(df_open[oh.CNAndLE].to_list())]
        df_task = df_task[~df_task[th.ppm_cdd_source].isin({'No CDD Required'})]
        df_task = df_task[
            df_task[th.due_date].apply(lambda x: (x.year, x.month)) >= (init_month.year, init_month.month)]

        if self.market == "Singapore":
            df_task = df_task[df_task[th.risk_rating].isin({"High", "SCC", "No Mapping"})]

        if self.df_ram is not None:
            rh = RamHeader
            df_ram = self.df_ram.copy()
            bow_df = self.df_task.copy()
            gsc_china_lob = {"BBPM", "GP", "ISB", "IVB"}
            # df_ram = df_ram[df_ram[rh.LOB] != "BBPM"]
            task_lob = df_task[th.task_id].map(df_ram.set_index(df_ram[rh.CustomerNumber])[rh.LOB])
            close_lob = df_close[ch.CustomerNumber].map(df_ram.set_index(df_ram[rh.CustomerNumber])[rh.LOB])
            open_lob = df_open[ch.CustomerNumber].map(df_ram.set_index(df_ram[rh.CustomerNumber])[rh.LOB])

            non_bbpm_task_df = df_task[~task_lob.isin(gsc_china_lob)]
            df_close = df_close[~close_lob.isin(gsc_china_lob)]
            df_open = df_open[~open_lob.isin(gsc_china_lob)]

            extra_task = df_ram[~df_ram[rh.CustomerNumber].isin(bow_df[th.task_id])]
            extra_task = extra_task[~extra_task[rh.LOB].isin(gsc_china_lob)]
            extra_task = extra_task[
                extra_task[rh.DueDate].apply(lambda x: (x.year, x.month)) >= (init_month.year, init_month.month)
            ]
            extra_task = extra_task[[rh.CustomerNumber, rh.DueDate]]
            df_task = pd.concat([non_bbpm_task_df, extra_task])

        self.df_task = df_task
        self.df_close = df_close
        self.df_open = df_open

        return


