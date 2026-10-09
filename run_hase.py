from hase_assignment import HaseAssignment
import pandas as pd
from qlik_config import Settings
from qlik_connection import QlikConnection
from qlik_export_csv import QlikExporter
import logging
import os
import shutil
from log import setup_log
import json
import datetime


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

if os.path.isfile('config.json'):
    with open('config.json', 'r') as f:
        config = json.load(f)
if config['as_of_date'] is not None:
    current_date = config['as_of_date']
else:
    current_date = str(datetime.datetime.now().date())

logging.info(f"Current Date is: {current_date}")

if os.path.isfile('qlik_config.json'):
    with open('qlik_config.json', 'r') as f:
        qlik_config = json.load(f)
logging.info(f"Loaded Qlik Config: {qlik_config}")
download_qlik_sense_report(qlik_config, current_date)

self = HaseAssignment('config.json', current_date)
self.read_input()
self.prepare_case_list()
self.prepare_cm_list()
self.prepare_history()
self.run_proposed_assignment()
self.export_assignment_workbook()
