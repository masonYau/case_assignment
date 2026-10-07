import time
import win32com.client as win32
import win32gui, win32con
import win32api
import re
import datetime
import uuid
import logging
from log import setup_log
import os

from case_assignment import CDDTaskAllocator


def get_outlook():
    try:
        return win32.GetActiveObject('Outlook.Application')
    except BaseException:
        return win32.Dispatch('Outlook.Application')


def get_inspector_hwnd(mail, ensure_display=True, wait=0.2):
    """
    尝试通过 Inspector.Hwnd 获取窗口句柄（优先）。
    如果失败，回退到根据邮件 Subject 在顶层窗口标题中查找匹配项。
    注意：Inspector.Hwnd 只有在 Inspector 已经显示后才可靠。
    """
    if ensure_display:
        try:
            # 非模态显示（避免阻塞）
            mail.Display(False)
            time.sleep(wait)
        except BaseException:
            pass

    # 1) 优先使用 COM Inspector.Hwnd
    try:
        insp = mail.GetInspector()   # 注意：必须加 ()，是方法
        hwnd = int(insp.Hwnd)       # 有时返回的是长整型
        if hwnd:
            return hwnd
    except BaseException:
        # 继续尝试枚举窗口
        pass

    # 2) 回退：根据邮件主题在顶层窗口标题中匹配
    subject = getattr(mail, "Subject", "") or ""
    if subject:
        found = []
        def enum_proc(h, l):
            if win32gui.IsWindowVisible(h):
                title = win32gui.GetWindowText(h)
                if title and subject in title:
                    found.append(h)
            return True
        win32gui.EnumWindows(enum_proc, None)
        if found:
            return found[0]

    # 最终失败返回 None
    return None


def bring_window_foreground(hwnd):
    """尝试把窗口恢复并置为前台。返回 True/False。"""
    if not hwnd:
        return False
    try:
        # 如果最小化/隐藏，先恢复
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.05)
        win32gui.SetForegroundWindow(hwnd)
        return True
    except BaseException:
        # 更复杂的跨线程/权限问题时，尝试 AttachThreadInput 技巧
        try:
            import ctypes
            user32 = ctypes.windll.user32
            kernel32 = ctypes.windll.kernel32

            fg = user32.GetForegroundWindow()
            if fg == hwnd:
                return True

            cur_tid = kernel32.GetCurrentThreadId()
            fg_tid = user32.GetWindowThreadProcessId(fg, 0)
            # AttachThreadInput( fg_tid, cur_tid, True )
            user32.AttachThreadInput(fg_tid, cur_tid, True)
            user32.SetForegroundWindow(hwnd)
            user32.AttachThreadInput(fg_tid, cur_tid, False)
            return True
        except BaseException:
            return False


def press_key_vk(vk, press_time=0.02):
    # 按下
    win32api.keybd_event(vk, 0, 0, 0)
    time.sleep(press_time)
    # 抬起
    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)


def press_alt_plus_char(ch, hold_alt=0.06):
    # Alt + 字母（例如发送 Alt+S）
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)  # 按下 Alt
    time.sleep(0.01)
    win32api.keybd_event(ord(ch.upper()), 0, 0, 0)   # 按下字母
    time.sleep(0.02)
    win32api.keybd_event(ord(ch.upper()), 0, win32con.KEYEVENTF_KEYUP, 0)  # 抬起字母
    time.sleep(0.01)
    win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)  # 抬起 Alt
    time.sleep(hold_alt)


def _subject_escape(s):
    return (s or "").replace("'", "''")


def is_mail_in_sent_items(mail_info, outlook=None, timeout=30, poll_interval=0.5,
                          body_snippet_len=40, lookback_count=50, verbose=False):
    """
    在 Outlook 已发送邮件中查找是否包含与 mail 匹配的项。
    返回 True 表示找到（即认为已发送），超时返回 False。

    参数:
      mail: COM MailItem（创建或显示过的邮件对象）
      outlook: 可选 Outlook.Application（None 则自动获取）
      timeout: 最长等待秒数
      poll_interval: 每次轮询间隔秒数
      body_snippet_len: 用于匹配的正文片段长度（若为0则不匹配正文）
      lookback_count: 回退检查时扫描已发送邮件的数量（按 SentOn 降序）
      verbose: 是否打印调试信息
    """
    if outlook is None:
        outlook = get_outlook()
    ns = outlook.GetNamespace("MAPI")
    try:
        sent_folder = ns.GetDefaultFolder(5)  # olFolderSentMail = 5
    except BaseException:
        sent_folder = None

    subj = mail_info["Subject"]
    to_field = mail_info["To"]
    email_code = mail_info["Code"]

    start_time = datetime.datetime.now()
    end_time = time.time() + timeout

    while time.time() < end_time:
        # 方案 A：如果有 subject，先用 Restrict 精确匹配 Subject（速度较快）
        candidates = []
        if sent_folder is not None and subj:
            try:
                restr = "[Subject] = '{}'".format(_subject_escape(subj))
                items = sent_folder.Items.Restrict(restr)
                # 注意 Items 不支持 len() 直接，迭代取到候选列表
                for it in items:
                    candidates.append(it)
            except BaseException as e:
                if verbose:
                    logging.info("Restrict 出错，回退到遍历已发送最近项：", e)
                candidates = []

        # 方案 B（或 A 无结果）：回退到取最近若干已发送邮件
        if not candidates and sent_folder is not None:
            try:
                items_all = sent_folder.Items
                # 按 SentOn 降序，确保最近的项先检查
                items_all.Sort("[SentOn]", True)  # True 表示降序
                # 取前 lookback_count 项
                # 注意：Items 对象并不是像列表那样切片，这里逐项迭代计数
                cnt = 0
                for it in items_all:
                    candidates.append(it)
                    cnt += 1
                    if cnt >= lookback_count:
                        break
            except BaseException as e:
                if verbose:
                    logging.info("遍历已发送文件夹出错:", e)
                candidates = []

        # 遍历候选项并判断匹配
        for it in candidates:
            try:
                it_subj = getattr(it, "Subject", "") or ""
                it_to = getattr(it, "To", "") or ""
                it_body = getattr(it, "Body", "") or ""
                it_sent_on = getattr(it, "SentOn", None)

                if email_code in it_body:
                    # 若通过所有检测项，则认为已发送
                    if verbose:
                        logging.info(
                            "在已发送文件夹中找到匹配项: Email Code {}".format(email_code)
                        )
                    return True
            except BaseException:
                # 单项检查失败则跳过
                continue

        # 未找到，等待下一轮
        time.sleep(poll_interval)

    if verbose:
        logging.info("超时未在已发送文件夹中找到匹配邮件（timeout={}s）".format(timeout))
    return False


def send_email(mail_info, select_classification=True, draft_only=True):
    outlook = get_outlook()
    mail = outlook.CreateItem(0)

    email_unique_code = str(uuid.uuid4())
    mail.To = mail_info["To"]
    mail.Subject = mail_info["Subject"]
    if "HTMLBody" in mail_info and mail_info["HTMLBody"]:
        mail.HTMLBody = mail_info["HTMLBody"]
    else:
        mail.Body = mail_info.get("Body", "")

    for f in mail_info.get("Attachments", []):
        try:
            mail.Attachments.Add(os.path.abspath(f))
        except BaseException as e:
            logging.error("Cannot add attachment: %s" % e)


    # 显示且确保 Inspector 已创建
    mail.Display(False)
    time.sleep(0.15)

    hwnd = get_inspector_hwnd(mail, ensure_display=False)
    if not hwnd:
        logging.info("无法通过 Inspector.Hwnd 或标题匹配到窗口，按键可能不会送达目标。")
    else:
        ok = bring_window_foreground(hwnd)
        if not ok:
            logging.info("尝试置前失败，按键可能不会送达目标。")

    # 示例：使用 Alt+S (Outlook 中通常是 Send 或 Save视语言/版本而定)
    # 示例：用 Alt+S 发送（或保存，视 Outlook 快捷键而定）
    if not draft_only:
        press_alt_plus_char('S')

        if select_classification:
            # 按两次 Tab（示例），每次完整的 down/up
            for _ in range(2):
                press_key_vk(win32con.VK_TAB, press_time=0.05)
                time.sleep(0.05)

            # 按空格两次
            for _ in range(2):
                press_key_vk(win32con.VK_SPACE, press_time=0.05)
                time.sleep(0.05)

        # press_key_vk(win32con.VK_RETURN, press_time=0.05)
        # time.sleep(0.05)
        #
        # for _ in range(2):
        #     press_key_vk(win32con.VK_DOWN, press_time=0.05)
        #     time.sleep(0.05)
        #
        # press_key_vk(win32con.VK_RETURN, press_time=0.05)
        # time.sleep(0.05)
        #
        # press_key_vk(win32con.VK_TAB, press_time=0.05)
        # time.sleep(0.05)
        #
        # press_key_vk(win32con.VK_RETURN, press_time=0.05)
        # time.sleep(0.05)

def build_case_assignment_stats_table_with_mg_html(stats: dict) -> str:
    """
    stats example:
    {
      "cm_handled_customer": {"6m": "12.3%", "3m": "10.1%", "1m": "9.8%", "current": "11.0%"},
      "cm_handled_rm":       {"6m": "45.0%", "3m": "41.2%", "1m": "39.9%", "current": "42.5%"},
    }
    """
    return f"""
            <table style="border-collapse:collapse; width: 100%; font-family: Calibri, Arial; font-size: 11pt;">
              <tr>
                <th style="border:1px solid #999; padding:8px; background:#f2f2f2; text-align:left;">Metric</th>
                <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Last 6 months</th>
                <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Last 3 months</th>
                <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Last 1 month</th>
                <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Current report</th>
              </tr>
            
              <!-- NEW ROW: cap/upper-limit percentage -->
              <tr>
                <td style="border:1px solid #999; padding:8px;">
                  Cap (%) of cases allowed to be assigned to a CM has previously handled the same customer
                </td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["6m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["3m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["1m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["current"]}</td>
              </tr>
            
              <tr>
                <td style="border:1px solid #999; padding:8px;">
                  % of cases where the assigned CM has previously handled the same customer
                </td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["6m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["3m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["1m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["current"]}</td>
              </tr>
            
              <tr>
                <td style="border:1px solid #999; padding:8px;">
                  % of cases where the assigned CM has previously handled the same master group
                </td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_mg"]["6m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_mg"]["3m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_mg"]["1m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_mg"]["current"]}</td>
              </tr>
            
              <tr>
                <td style="border:1px solid #999; padding:8px;">
                  % of cases where the assigned CM has previously handled customers managed by the same RM
                </td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["6m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["3m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["1m"]}</td>
                <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["current"]}</td>
              </tr>
            </table>
            """

def build_case_assignment_stats_table_html(stats: dict) -> str:
    """
    stats example:
    {
      "cm_handled_customer": {"6m": "12.3%", "3m": "10.1%", "1m": "9.8%", "current": "11.0%"},
      "cm_handled_rm":       {"6m": "45.0%", "3m": "41.2%", "1m": "39.9%", "current": "42.5%"},
    }
    """
    return f"""
    <table style="border-collapse:collapse; width: 100%; font-family: Calibri, Arial; font-size: 11pt;">
      <tr>
        <th style="border:1px solid #999; padding:8px; background:#f2f2f2; text-align:left;">Metric</th>
        <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Last 6 months</th>
        <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Last 3 months</th>
        <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Last 1 month</th>
        <th style="border:1px solid #999; padding:8px; background:#f2f2f2;">Current report</th>
      </tr>
      <tr>
        <td style="border:1px solid #999; padding:8px;">
          Cap (%) of cases allowed to be assigned to a CM has previously handled the same customer
        </td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["6m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["3m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["1m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["highest_same_cm"]["current"]}</td>
      </tr>
      <tr>
        <td style="border:1px solid #999; padding:8px;">
          % of cases where the assigned CM has previously handled the same customer
        </td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["6m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["3m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["1m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_customer"]["current"]}
      </td>
      <tr>
        <td style="border:1px solid #999; padding:8px;">
          % of cases where the assigned CM has previously handled customers managed by the same RM
        </td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["6m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["3m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["1m"]}</td>
        <td style="border:1px solid #999; padding:8px; text-align:center;">{stats["cm_handled_rm"]["current"]}</td>
      </tr>
    </table>
    """


class EmailSender:

    def __init__(self,
                 config: dict,
                 case_assignment: CDDTaskAllocator
                 ):
        self.config = config
        self.case_assignment = case_assignment
        self.fail_msg = []

    def send_case_assigment_emails(self, draft_only=True):
        email_to_list = self.config["EMAIL_TO_LIST"]
        market = self.config["MARKET"]
        legal_entity = ", ".join(self.config["LEGAL_ENTITY"])
        staff_email = ";".join(email_to_list)
        report_file = self.case_assignment.out_dir
        target_wip_file = self.case_assignment.analyst_out_dir
        stats = self.case_assignment.cm_statistics

        if report_file is None:
            logging.info(f"Case assignment report of {market} {market} is not generated， email will not sent")
            return

        logging.info(f"generating TL daily email for {market} {market} {staff_email}")

        email_unique_code = str(uuid.uuid4())
        if legal_entity == "HK-HSBC":
            table_html = build_case_assignment_stats_table_with_mg_html(stats)
        else:
            table_html = build_case_assignment_stats_table_html(stats)

        html_body = f"""
                <p>Attached is your Case Assignment Report ({market} {legal_entity}).</p>
                <p><b>Summary statistics</b></p>
                {table_html}
                <p><b>Statistical criteria/definition for the figures above:</b></p>
                <ol>
                    <li>Only PR cases are included.</li>
                    <li>Only cases with a status of <b>Completed</b> or <b>WIP (Work in Progress)</b> are included.</li>
                    <li>Only <b>GSC China BoW analysts</b> listed on the <b>HC report</b> are included.</li>
                </ol>
                <p style="color:#666; font-size:10pt;">email code: {email_unique_code}</p>
                """
        mail_info = {
            "To": staff_email,
            "Subject": f'[Internal] Case Assignment Report ({market} {legal_entity}) {datetime.datetime.now().date()}',
            "HTMLBody": html_body,
            "Attachments": [report_file, target_wip_file],
            "Code": email_unique_code
        }

        send_email(mail_info, select_classification=True, draft_only=draft_only)

        if not draft_only:
            time.sleep(0.15)
            sent = is_mail_in_sent_items(mail_info, verbose=True)
            if sent:
                logging.info(f"Email has been sent {mail_info['Subject']} {email_unique_code}")
            else:
                logging.error(f"Email sent fail {mail_info['Subject']} {email_unique_code}")
                self.fail_msg.append(f"Email sent fail {mail_info['Subject']} {email_unique_code}")

            logging.info(f"All email has been sent")

    def save_error_info(self):
        with open("email_error.txt", "w") as f:
            for msg in self.fail_msg:
                f.write(msg + "\n")

