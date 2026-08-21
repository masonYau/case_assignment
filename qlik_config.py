import os
import requests
from requests_ntlm import HttpNtlmAuth  # 需要 pip install requests_ntlm
import logging


class Settings:

    def __init__(self, config: dict):
        self.config = config
        # Qlik Sense 服务器配置
        self.HOST = self.config.get("QLIK_HOST", "ws-qs-sense.systems.uk.hsbc")
        self.APP_ID = self.config.get("QLIK_APP_ID", "af1844ea-0c16-4805-bd6f-97759e648164")
        self.OBJ_ID = {
            "closed_review": self.config.get("CLOSED_REVIEW_OBJ_ID", "jZas"),
            "open_review": self.config.get("OPEN_REVIEW_OBJ", "KWnDsz")
        }
        self.qlik_base_url = self.config.get("QLIK_URL", "https://ws-qs-sense.systems.uk.hsbc")  # 注意是用 https
        self.username = self.config.get("QLIK_USER", "")
        self.password = self.config.get("QLIK_PSW", "")
        self.X_Qlik_User = f"UserDirectory=HBAP; UserId={self.username}"
        self.User_agent ="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

        # 会话配置
        # SESSION_VALUE = self.config.get("QLIK_SESSION", "d1e55531-39ac-4753-9056-fb55c0063c41")
        # 配置

        # 1. 创建一个 Session 对象，它会自动保存 Cookies
        session = requests.Session()
        #
        # # 2. 发起一个 HTTP 请求来“激活”登录
        # # 这里使用 NTLM 认证，或者如果不需要密码直接访问 hub 也行
        try:
            # 尝试访问 Hub 页面，触发认证流程
            target_url = f"{self.qlik_base_url}/sense/app/{self.APP_ID}"
            response = session.get(
                target_url,
                verify=False,
                auth=HttpNtlmAuth(self.username, self.password),
                headers={"X-Qlik-User": self.X_Qlik_User,"User-Agent": self.User_agent}
            )
            logging.info(f"HTTP 状态码: {response.status_code}")
            # 3. 提取 Cookie
            cookies = session.cookies.get_dict()
            logging.info("获取到的 Cookies:", cookies)
            # Qlik 的 Session ID 通常叫 'X-Qlik-Session' 或类似名字
            # 只要有了这个 cookies 字典，把它格式化成 WebSocket 需要的 header 字符串即可
            session_id = "; ".join([f"{k}={v}" for k, v in cookies.items()])
            logging.info("测试内容为:" + session_id)
            self.HEADERS = {
                "Cookie": session_id,
                "X-Qlik-User": self.X_Qlik_User,
                "User-Agent": self.User_agent
            }
        except Exception as e:
            logging.info(f"登录失败: {e}")

        # 导出配置和bookmark配置
        self.REPORTS = self.config.get("REPORTS", dict())

        # 变量配置
        self.CALC_DETAILS_VAR = "v_calc_Details"
        self.ATTRIBUTES_VAR = "v_Attributes"

        # WebSocket 配置
        # HEADERS = {
        #     "Cookie": f"X-Qlik-Session={SESSION_VALUE}",
        #     "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        # }

        self.WS_URL = f"wss://{self.HOST}/app/{self.APP_ID}"

        # 导出格式配置
        self.EXPORT_FORMAT = "CSV_C"
        self.EXPORT_PATH = "export.csv"
        self.EXPORT_MODE = "A"  # A=All, P=Possible

