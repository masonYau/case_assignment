# main.py

```python
import pandas as pd

from qlik_config import Settings
from qlik_connection import QlikConnection
from qlik_export_csv import QlikExporter
import logging
import json
from log import setup_log
import os
from task_clean import TaskClean
from case_assignment import CDDTaskAllocator
from email_sender import EmailSender
import shutil

def download_qlik_sense_report(config, current_date):
    settings = Settings(config)
    conn = QlikConnection(settings)
    exporter = QlikExporter(conn)

    if not os.path.isdir("history_data/"):
        os.mkdir("history_data/")
    try:
        # 步骤1: 建立连接
        conn.connect()

        # 步骤2: 打开文档 (关键修复：使用request方法)
        if not exporter.open_doc():
            raise Exception("Unable to Open Qlik Doc")

        # 步骤3: 逐个处理报表
        for name, cfg in settings.REPORTS.items():
            logging.info(f"\n► Handling Qlik Report: {name}")

            # 初始化报表环境（应用专属书签）
            if cfg.get("bookmark_id"):
                exporter.init_report(cfg["bookmark_id"])
            elif cfg.get("bookmark_config"):
                exporter.init_report_by_config(cfg["bookmark_config"])
            else:
                raise Exception(f"No Bookmark Configuration of {name}")
            # 导出数据
            success = exporter.export_object(cfg["obj_id"], cfg["output_file"])
            shutil.copy(cfg["output_file"], f'history_data/{cfg["output_file"].split(".")[0]}_{current_date}.csv')

            result = "Success" if success else "Failed"
            logging.info(f"  {result} → {cfg['output_file']}")

    except Exception as e:
        raise e
    finally:
        conn.close()


setup_log("./")

current_date = "2026-08-13"

if os.path.isfile('qlik_config.json'):
    with open('qlik_config.json', 'r') as f:
        qlik_config = json.load(f)
logging.info(f"Loaded Qlik Config: {qlik_config}")
download_qlik_sense_report(qlik_config, current_date)

"""
    AMH
"""
if os.path.isfile('config_amh.json'):
    with open('config_amh.json', 'r') as f:
        amh_config = json.load(f)

logging.info(f"Loaded AMH Config: {amh_config}")
oe_report_df = pd.concat([pd.read_csv(f) for f in amh_config["OE_REPORT_FILES"]])
close_df = pd.concat([pd.read_csv(f) for f in amh_config["CLOSE_REPORT_FILES"]])
open_df = pd.read_csv(amh_config["OPEN_REPORT_FILES"])
ram_df = pd.read_csv(amh_config["RAM_FILES"])
analyst_df = pd.read_excel(amh_config["CASE_MANAGER_LIST_FILES"])
mg_df = pd.read_excel(amh_config["MG_FILES"])

amh_data = TaskClean(
    open_df,
    close_df,
    oe_report_df,
    ram_file=ram_df,
    df_mg=mg_df,
    market=amh_config["MARKET"],
    legal_entity=amh_config["LEGAL_ENTITY"],
    current_date=current_date
)
amh_data.data_preprocess()
amh_data.close_data_process()
amh_data.clean_task()
logging.info(f"AMH complete data clean")


amh_task_assignment = CDDTaskAllocator(
    df_analysts=analyst_df,
    df_tasks=amh_data.df_task,
    df_close=amh_data.df_close,
    df_open=amh_data.df_open,
    current_date=current_date,
    config=amh_config
)
amh_task_assignment.read_data()
amh_task_assignment.add_history_customer_to_analyst()
amh_task_assignment.run_top_n_analyst()
amh_task_assignment.output_result(amh_config['OUT_RESULT_FILES'])
amh_task_assignment.output_analyst_df(amh_config['OUT_TARGET_WIP_FILES'])
logging.info(f"AMH complete case assignment")

"""
    HASE
"""
if os.path.isfile('config_hase.json'):
    with open('config_hase.json', 'r') as f:
        hase_config = json.load(f)

logging.info(f"Loaded HASE Config: {hase_config}")
oe_report_df = pd.concat([pd.read_csv(f) for f in hase_config["OE_REPORT_FILES"]])
close_df = pd.concat([pd.read_csv(f) for f in hase_config["CLOSE_REPORT_FILES"]])
open_df = pd.read_csv(hase_config["OPEN_REPORT_FILES"])
analyst_df = pd.read_excel(hase_config["CASE_MANAGER_LIST_FILES"])

hase_data = TaskClean(
    open_df,
    close_df,
    oe_report_df,
    ram_file=None,
    market=hase_config["MARKET"],
    legal_entity=hase_config["LEGAL_ENTITY"],
    current_date=current_date
)
hase_data.data_preprocess()
hase_data.close_data_process()
hase_data.clean_task()
logging.info(f"HASE complete data clean")

hase_task_assignment = CDDTaskAllocator(
    df_analysts=analyst_df,
    df_tasks=hase_data.df_task,
    df_close=hase_data.df_close,
    df_open=hase_data.df_open,
    current_date=current_date,
    config=hase_config
)
hase_task_assignment.read_data()
hase_task_assignment.add_history_customer_to_analyst()
hase_task_assignment.run_top_n_analyst()
hase_task_assignment.output_result(hase_config['OUT_RESULT_FILES'])
hase_task_assignment.output_analyst_df(hase_config['OUT_TARGET_WIP_FILES'])
logging.info(f"HASE complete case assignment")


"""
    CHN
"""
if os.path.isfile('config_chn.json'):
    with open('config_chn.json', 'r') as f:
        chn_config = json.load(f)

logging.info(f"Loaded CHN Config: {chn_config}")
oe_report_df = pd.concat([pd.read_csv(f) for f in chn_config["OE_REPORT_FILES"]])
close_df = pd.concat([pd.read_csv(f) for f in chn_config["CLOSE_REPORT_FILES"]])
open_df = pd.read_csv(chn_config["OPEN_REPORT_FILES"])
analyst_df = pd.read_excel(chn_config["CASE_MANAGER_LIST_FILES"])

chn_data = TaskClean(
    open_df,
    close_df,
    oe_report_df,
    ram_file=None,
    market=chn_config["MARKET"],
    legal_entity=chn_config["LEGAL_ENTITY"],
    current_date=current_date
)
chn_data.data_preprocess()
chn_data.close_data_process()
chn_data.clean_task()
logging.info(f"CHN complete data clean")

chn_task_assignment = CDDTaskAllocator(
    df_analysts=analyst_df,
    df_tasks=chn_data.df_task,
    df_close=chn_data.df_close,
    df_open=chn_data.df_open,
    current_date=current_date,
    config=chn_config
)
chn_task_assignment.read_data()
chn_task_assignment.add_history_customer_to_analyst()
chn_task_assignment.run_top_n_analyst()
chn_task_assignment.output_result(chn_config['OUT_RESULT_FILES'])
chn_task_assignment.output_analyst_df(chn_config['OUT_TARGET_WIP_FILES'])
logging.info(f"CHN complete case assignment")

"""
    SGP
"""
if os.path.isfile('config_sgp.json'):
    with open('config_sgp.json', 'r') as f:
        sgp_config = json.load(f)

logging.info(f"Loaded SGP Config: {sgp_config}")
oe_report_df = pd.concat([pd.read_csv(f) for f in sgp_config["OE_REPORT_FILES"]])
close_df = pd.concat([pd.read_csv(f) for f in sgp_config["CLOSE_REPORT_FILES"]])
open_df = pd.read_csv(sgp_config["OPEN_REPORT_FILES"])
analyst_df = pd.read_excel(sgp_config["CASE_MANAGER_LIST_FILES"])

sgp_data = TaskClean(
    open_df,
    close_df,
    oe_report_df,
    ram_file=None,
    market=sgp_config["MARKET"],
    legal_entity=sgp_config["LEGAL_ENTITY"],
    current_date=current_date
)
sgp_data.data_preprocess()
sgp_data.close_data_process()
sgp_data.clean_task()
logging.info(f"SGP complete data clean")

sgp_task_assignment = CDDTaskAllocator(
    df_analysts=analyst_df,
    df_tasks=sgp_data.df_task,
    df_close=sgp_data.df_close,
    df_open=sgp_data.df_open,
    current_date=current_date,
    config=sgp_config
)
sgp_task_assignment.read_data()
sgp_task_assignment.add_history_customer_to_analyst()
sgp_task_assignment.run_top_n_analyst()
sgp_task_assignment.output_result(sgp_config['OUT_RESULT_FILES'])
sgp_task_assignment.output_analyst_df(sgp_config['OUT_TARGET_WIP_FILES'])
logging.info(f"SGP complete case assignment")

amh_email_sender = EmailSender(
    config=amh_config,
    case_assignment=amh_task_assignment
)
amh_email_sender.send_case_assigment_emails()
logging.info(f"AMH complete send email")

hase_email_sender = EmailSender(
    config=hase_config,
    case_assignment=hase_task_assignment
)
hase_email_sender.send_case_assigment_emails()
logging.info(f"HASE complete send email")

chn_email_sender = EmailSender(
    config=chn_config,
    case_assignment=chn_task_assignment
)
chn_email_sender.send_case_assigment_emails()
logging.info(f"CHN complete send email")

sgp_email_sender = EmailSender(
    config=sgp_config,
    case_assignment=sgp_task_assignment
)
sgp_email_sender.send_case_assigment_emails()
logging.info(f"SGP complete send email")
```
