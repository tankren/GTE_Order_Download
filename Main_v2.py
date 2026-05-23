# -*- coding: utf-8 -*-
"""
GTE Order Auto-Download Tool v0.2 (PyQt6)
Created on Thu Aug  9 21:05:37 2022
Refactored 2024
@author: REC3WX
"""

from datetime import datetime, timedelta
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton,
    QDateEdit, QCheckBox, QPlainTextEdit, QFrame,
    QGridLayout, QMessageBox, QStyle, QCalendarWidget,
    QSystemTrayIcon, QTimeEdit,
)
from PyQt6.QtGui import QFont
from PyQt6.QtCore import pyqtSlot, Qt, QThread, pyqtSignal, QEvent, QDate
from apscheduler.schedulers.background import BackgroundScheduler
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
from zipfile import ZipFile
from pathlib import Path
from lxml import etree
import re
import os
import tempfile
import sys
import requests
import smtplib

DEFAULT_CONFIG = {
    "user": "1509A01",
    "pwd": "1509-A01",
    "recipients": (
        "fixed-term.shiyu.wei@vhit-weifu.com;feng.he@vhit-weifu.com;"
        "wenzhuo.gu@vhit-weifu.com;wenjie.shen@vhit-weifu.com;"
        "CS01.VHCN@Fiege.com.cn"
    ),
    "mail_host": "smtp.163.com",
    "smtp_port": 465,
    "server_url": "http://192.168.10.33",
    "timezone": "Asia/Shanghai",
    "page_size": 100,
    "request_timeout": 5,
}

APP_TITLE = "GTE订单查询下载工具 v0.2   - Made by REC3WX"


def to_unicode(text):
    """将文本转换为unicode编码字符串"""
    return "".join(
        hex(ord(c)).lower().replace("0x", "\\\\u") for c in text
    )


def decode_filename(raw):
    """解码乱码文件名: cp437 -> gbk, fallback utf-8"""
    try:
        return raw.encode("cp437").decode("gbk")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return raw.encode("utf-8").decode("utf-8")


class EmailService:
    """邮件发送服务"""

    @staticmethod
    def html_content(body_text):
        return f"""\
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
  <title>GTE订单自动下载结果</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
</head>
<body style="margin: 0; padding: 0;">
  <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td style="padding: 20px 0 30px 0;">
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="600" style="border-collapse: collapse; border: 1px solid #cccccc;">
          <tr>
            <td align="center" bgcolor="#005691" style="padding: 20px 0 20px 0;">
                <h1 style="color: #ffffff; font-size: 24px; margin: 0; font-family: Microsoft YaHei;">GTE订单自动下载结果</h1>
            </td>
          </tr>
          <tr>
            <td bgcolor="#ffffff" style="padding: 30px 20px 30px 20px;">
              <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                <tr>
                  <td align="center" style="color: #000000; font-family: Microsoft YaHei;">
                    <p style="margin: 0;">{body_text}</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td bgcolor="#005691" style="padding: 10px 10px;">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse;">
                <tr>
                  <td align="center" style="color: #00000; font-family: Microsoft YaHei; font-size: 14px;">
                    <p style="color: #ffffff; margin: 0;">Powered by VHCN ICO</p>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>"""

    @staticmethod
    def send(folder, recipients, is_once):
        mail_host = os.environ.get("MAIL_HOST") or DEFAULT_CONFIG["mail_host"]
        mail_user = os.environ.get("MAIL_USER", "")
        mail_pass = os.environ.get("MAIL_PASS", "")

        email = MIMEMultipart()
        now_str = datetime.now().isoformat(" ", "seconds")
        email["Subject"] = f"no-reply: {now_str} GTE订单自动下载结果"
        email["From"] = formataddr(("Mail Bot@VHCN", mail_user))
        email["To"] = recipients

        file_list = [f.name for f in Path(folder).glob("*.zip")]
        for filename in file_list:
            file_path = os.path.join(folder, filename)
            with open(file_path, "rb") as f:
                part = MIMEApplication(f.read())
            part.add_header("Content-Disposition", "attachment", filename=filename)
            email.attach(part)

        if not file_list and is_once:
            return "单次执行无订单不发送邮件!"
        elif not file_list:
            email_content = EmailService.html_content("没有新的订单, 请知悉!")
        else:
            email_content = EmailService.html_content("附件为新的订单, 请查收!")

        email.attach(MIMEText(email_content, "html", "utf-8"))

        try:
            smtp = smtplib.SMTP_SSL(mail_host, DEFAULT_CONFIG["smtp_port"])
            smtp.login(mail_user, mail_pass)
            smtp.sendmail(mail_user, recipients.split(";"), email.as_string())
            smtp.quit()
            return "邮件发送成功..."
        except smtplib.SMTPException as e:
            return str(e)


class GteOrderService:
    """GTE 订单查询与下载服务"""

    def __init__(self, user, password, supplier_slice=5):
        self.user = user
        self.user2 = f"{user[0:4]}-A"
        self.password = password
        self.supplier = to_unicode(user[0:supplier_slice])
        self._stop_requested = False
        self.base_url = DEFAULT_CONFIG["server_url"]
        self.timeout = DEFAULT_CONFIG["request_timeout"]
        self.OrderNo = ""

    def stop(self):
        self._stop_requested = True

    @property
    def wcf_url(self):
        return f"{self.base_url}/WCFService/WcfService.svc"

    def _soap_headers(self, action):
        return {
            "content-type": "text/xml",
            "Referer": f"{self.base_url}/ClientBin/SilverlightUI.xap",
            "SOAPAction": f'"SysManager/WcfService/{action}"',
        }

    def mark_downloaded(self, order_no):
        """回传首次下载时间"""
        body = (
            '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">'
            "<s:Body>"
            '<UpeFirstDownLoadTimeIPO0710 xmlns="SysManager">'
            '<baseInfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities"'
            ' xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            "</baseInfo>"
            f'<prms>{{"OrderNo":"{to_unicode(order_no)}"}}</prms>'
            "</UpeFirstDownLoadTimeIPO0710>"
            "</s:Body>"
            "</s:Envelope>"
        )
        requests.post(
            self.wcf_url, data=body,
            headers=self._soap_headers("UpeFirstDownLoadTimeIPO0710"),
            timeout=self.timeout,
        )

    def query_orders(self, ordfrom="", ordtill="", include_downloaded=False, is_once=False):
        """查询订单列表，返回 (found, filenames_with_data) 的生成器"""
        chk = "1" if include_downloaded else "0"
        if include_downloaded and is_once:
            scope = (
                f'<parm>{{"SupplierNum":"{self.supplier}","Series":"-1",'
                f'"ActualOrderTimeS":"{ordfrom}","ActualOrderTimeE":"{ordtill}",'
                f'"Check":"{chk}"}}</parm>'
            )
        else:
            scope = (
                f'<parm>{{"SupplierNum":"{self.supplier}","Series":"-1","Check":"{chk}"}}</parm>'
            )

        body = (
            '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">'
            "<s:Body>"
            '<IPO0710GetInfo xmlns="SysManager">'
            '<baseInfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities"'
            ' xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            "</baseInfo>"
            '<pagerinfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities"'
            ' xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            f"<d4p1:pageIndex>1</d4p1:pageIndex>"
            f"<d4p1:pageSize>{DEFAULT_CONFIG['page_size']}</d4p1:pageSize>"
            "</pagerinfo>"
            f"{scope}"
            "</IPO0710GetInfo>"
            "</s:Body>"
            "</s:Envelope>"
        )

        resp = requests.post(
            self.wcf_url, data=body,
            headers=self._soap_headers("IPO0710GetInfo"),
            timeout=self.timeout,
        )
        xml = etree.fromstring(resp.content.decode("utf-8"))

        found = False
        for filenm in xml.iter("{*}FileNm"):
            if self._stop_requested:
                break
            found = True
            filename = filenm.text
            order_no = ""
            for ancestor in filenm.iterancestors("{*}IPO0710DTO"):
                order_no = ancestor.find("./{*}OrderNo").text
            yield filename, order_no
        if not found:
            yield None, None

    def download_file(self, filename):
        """下载单个订单文件"""
        date_part = filename[7:15]
        url = f"{self.base_url}/GTESGFile/Instruct/{self.user2}/{date_part}/{filename}"
        return requests.get(url=url, stream=True, timeout=self.timeout)


def rezip(filepath, temp_dir):
    """解压 -> 修复乱码文件名 -> 重新打包"""
    file_list = []
    with ZipFile(filepath, allowZip64=True) as unzip:
        for name in unzip.namelist():
            extracted = unzip.extract(name, temp_dir)
            file_list.append(extracted)

    for file in file_list:
        new_path = decode_filename(file)
        old_path = os.path.join(os.path.dirname(new_path), os.path.basename(file))
        if old_path != new_path:
            Path(old_path).rename(new_path)
    os.remove(filepath)
    with ZipFile(filepath, "w") as newzip:
        for dirpath, _, filenames in os.walk(temp_dir):
            for fname in filenames:
                newzip.write(os.path.join(dirpath, fname), fname)
    for fname in os.listdir(temp_dir):
        os.remove(os.path.join(temp_dir, fname))
    os.rmdir(temp_dir)


class DownloadWorker(QThread):
    log_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._stop_requested = False
        self.scheduler = None
        self.folder = os.path.join(tempfile.gettempdir(), "GET", "Order")
        self._ensure_folder()
        self.gte_service = None

    def _ensure_folder(self):
        if os.path.exists(self.folder):
            self.purge_folder()
        else:
            os.makedirs(self.folder)

    def purge_folder(self):
        for fname in os.listdir(self.folder):
            os.remove(os.path.join(self.folder, fname))

    def stop(self):
        self._stop_requested = True
        if self.gte_service is not None:
            self.gte_service.stop()
        try:
            if self.scheduler:
                self.scheduler.shutdown()
        except Exception:
            pass
        self.log_signal.emit("进程已终止...")

    def configure(self, ordfrom, ordtill, user, pwd, recipients,
                  include_downloaded, is_once, timer, workday_only):
        self.ordfrom = ordfrom
        self.ordtill = ordtill
        self.recipients = recipients
        self.include_downloaded = include_downloaded
        self.is_once = is_once
        self.timer = timer
        self.workday_only = workday_only
        self.gte_service = GteOrderService(user, pwd)

    def run(self):
        if self._stop_requested:
            return
        if self.is_once:
            try:
                self.chain()
            except Exception as e:
                self.log_signal.emit(str(e))
        else:
            self._start_scheduler()

    def _start_scheduler(self):
        try:
            self.scheduler = BackgroundScheduler(timezone=DEFAULT_CONFIG["timezone"])
            h, m = self.timer[0:2], self.timer[3:5]
            day_of_week = "*" if not self.workday_only else "mon-fri"
            self.scheduler.add_job(
                self.chain, trigger="cron",
                day_of_week=day_of_week, hour=h, minute=m,
            )
            self.scheduler.start()
            desc = f"每天{self.timer}" if day_of_week == "*" else f"每个工作日的{self.timer}"
            self.log_signal.emit(f"已设置定时任务为{desc} ")
            self._emit_next_run()
        except Exception as e:
            self.log_signal.emit(str(e))

    def _emit_next_run(self):
        now = datetime.now()
        now_hm = now.strftime("%H:%M")
        now_d = now.strftime("%Y-%m-%d")
        now_dhm = now.strftime("%Y-%m-%d %H:%M")
        delta_map = {4: 3, 5: 2}
        delta = delta_map.get(now.weekday(), 1)
        if now_hm < self.timer:
            sch_dhm = f"{now_d} {self.timer}"
        else:
            sch_date = datetime.strftime(now + timedelta(days=delta), "%Y-%m-%d")
            sch_dhm = f"{sch_date} {self.timer}"
        gap = datetime.strptime(sch_dhm, "%Y-%m-%d %H:%M") - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
        hours = int(gap.total_seconds() // 3600)
        minutes = int((gap.total_seconds() % 3600) // 60)
        self.log_signal.emit(f"下次任务时间为: {sch_dhm}, 距现在{hours}小时{minutes}分")

    def chain(self):
        try:
            prefix = "定时任务" if not self.is_once else "单次任务"
            self.log_signal.emit(f"{datetime.now()} - 开始运行{prefix}...")
            self.log_signal.emit("获取Zip文件清单...")
            self._process_orders()
            self.log_signal.emit("获取结束, 发送邮件...")
            result = EmailService.send(self.folder, self.recipients, self.is_once)
            self.log_signal.emit(result)
            self.purge_folder()
            if not self.is_once:
                self.log_signal.emit(f"{datetime.now()} - 定时任务已完成! ")
                self._emit_next_run()
            else:
                self.log_signal.emit(f"{datetime.now()} - 单次任务已完成! ")
        except Exception as e:
            self.log_signal.emit(str(e))

    def _process_orders(self):
        if self.gte_service is None:
            self.log_signal.emit("订单服务未初始化!")
            return
        for filename, order_no in self.gte_service.query_orders(
            self.ordfrom, self.ordtill,
            self.include_downloaded == "1",
            self.is_once,
        ):
            if self._stop_requested:
                break
            if filename is None or order_no is None:
                self.log_signal.emit("没有找到新的订单...")
                break
            try:
                self.log_signal.emit(f"命中: {filename}")
                self.gte_service.OrderNo = order_no
                self.log_signal.emit(f"开始下载Zip文件 {filename}")
                filepath = os.path.join(self.folder, filename)
                dld_zip = self.gte_service.download_file(filename)
                with open(filepath, "wb") as f:
                    f.write(dld_zip.content)
                temp_dir = filepath[:-4]
                rezip(filepath, temp_dir)
                self.log_signal.emit("下载完成! 回传下载时间")
                self.gte_service.mark_downloaded(order_no)
            except Exception as e:
                self.log_signal.emit(str(e))

    def stop_scheduler(self):
        self.log_signal.emit("正在终止定时任务...")
        try:
            if self.scheduler:
                self.scheduler.shutdown()
        except Exception as e:
            self.log_signal.emit(str(e))


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.thread = None
        self._setup_ui()
        self._setup_tray()

    def _setup_ui(self):
        self.setWindowTitle(APP_TITLE)
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_DialogOpenButton)
        self.setWindowIcon(icon)
        self.setFixedSize(800, 350)

        self._create_widgets()
        self._create_layout()

    def _create_widgets(self):
        self.lbl_user = QLabel("用户名:")
        self.edit_user = QLineEdit()
        self.edit_user.setClearButtonEnabled(True)
        self.edit_user.setText(DEFAULT_CONFIG["user"])

        self.lbl_pwd = QLabel("密码:")
        self.edit_pwd = QLineEdit()
        self.edit_pwd.setClearButtonEnabled(True)
        self.edit_pwd.setText(DEFAULT_CONFIG["pwd"])
        self.edit_pwd.setEchoMode(QLineEdit.EchoMode.PasswordEchoOnEdit)

        self.lbl_from = QLabel("开始日期:")
        self.date_from = QDateEdit(QDate.currentDate())
        self.date_from.setCursor(Qt.CursorShape.PointingHandCursor)
        self.date_from.setDisplayFormat("yyyy-MM-dd")
        self.date_from.setCalendarPopup(True)
        self.date_from.calendarWidget().setGridVisible(True)
        self.date_from.calendarWidget().setVerticalHeaderFormat(
            QCalendarWidget.VerticalHeaderFormat.ISOWeekNumbers
        )
        self.date_from.dateChanged.connect(self._on_date_changed)

        self.lbl_till = QLabel("结束日期:")
        self.date_till = QDateEdit(QDate.currentDate().addDays(5))
        self.date_till.setCursor(Qt.CursorShape.PointingHandCursor)
        self.date_till.setDisplayFormat("yyyy-MM-dd")
        self.date_till.setCalendarPopup(True)
        self.date_till.calendarWidget().setGridVisible(True)
        self.date_till.calendarWidget().setVerticalHeaderFormat(
            QCalendarWidget.VerticalHeaderFormat.ISOWeekNumbers
        )
        self.date_till.dateChanged.connect(self._on_date_changed)

        self.chk_downloaded = QCheckBox("包含已下载")
        self.chk_downloaded.setCursor(Qt.CursorShape.PointingHandCursor)
        self.chk_workday = QCheckBox("仅工作日")
        self.chk_workday.setCursor(Qt.CursorShape.PointingHandCursor)

        self.btn_once = QPushButton("运行单次任务")
        self.btn_once.clicked.connect(self.execute_once)

        self.lbl_time = QLabel("定时任务时间:")
        self.edit_time = QTimeEdit()
        self.edit_time.setDisplayFormat("HH:mm")
        self.edit_time.setCursor(Qt.CursorShape.PointingHandCursor)

        self.lbl_email = QLabel("收件人:")
        self.edit_email = QLineEdit()
        self.edit_email.setClearButtonEnabled(True)
        self.edit_email.setPlaceholderText("多个收件人之间用分号;分开")
        self.edit_email.setText(DEFAULT_CONFIG["recipients"])
        self.edit_email.setToolTip(self.edit_email.text())
        self.edit_email.editingFinished.connect(self._validate_email)

        self.lbl_log = QLabel("运行日志:")
        self.text_log = QPlainTextEdit()
        self.text_log.setReadOnly(True)

        self.separator = QFrame()
        self.separator.setFrameShape(QFrame.Shape.HLine)
        self.separator.setFrameShadow(QFrame.Shadow.Sunken)

        self.btn_schedule = QPushButton("设置定时任务")
        self.btn_schedule.clicked.connect(self.set_schedule)
        self.btn_cancel_sch = QPushButton("取消定时任务")
        self.btn_cancel_sch.clicked.connect(self.cancel_schedule)
        self.btn_cancel_sch.setEnabled(False)

        self.btn_reset = QPushButton("清空日志")
        self.btn_reset.clicked.connect(self.reset_log)

        self.btn_stop = QPushButton("终止任务进程")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_thread)

    def _create_layout(self):
        layout = QGridLayout()
        layout.addWidget(self.lbl_user, 0, 0)
        layout.addWidget(self.edit_user, 0, 1)
        layout.addWidget(self.lbl_pwd, 0, 2)
        layout.addWidget(self.edit_pwd, 0, 3)
        layout.addWidget(self.lbl_from, 0, 4)
        layout.addWidget(self.date_from, 0, 5)
        layout.addWidget(self.lbl_till, 1, 4)
        layout.addWidget(self.date_till, 1, 5)
        layout.addWidget(self.lbl_email, 1, 0)
        layout.addWidget(self.edit_email, 1, 1, 1, 3)
        layout.addWidget(self.lbl_time, 0, 6)
        layout.addWidget(self.edit_time, 0, 7)
        layout.addWidget(self.chk_downloaded, 1, 6)
        layout.addWidget(self.chk_workday, 1, 7)

        layout.addWidget(self.separator, 2, 0, 1, 8)
        layout.addWidget(self.lbl_log, 3, 0)
        layout.addWidget(self.text_log, 4, 0, 6, 6)
        layout.addWidget(self.btn_reset, 4, 6, 1, 2)
        layout.addWidget(self.btn_once, 6, 6, 1, 2)
        layout.addWidget(self.btn_schedule, 8, 6, 1, 2)
        layout.addWidget(self.btn_cancel_sch, 9, 6, 1, 2)

        self.setLayout(layout)

    def _setup_tray(self):
        icon = self.windowIcon()
        self.tray = QSystemTrayIcon()
        self.tray.setIcon(icon)
        self.tray.setToolTip(f"{APP_TITLE} 正在后台运行")
        self.tray.activated.connect(self._on_tray_activated)

    # ─── Slots ──────────────────────────────────────────────

    @pyqtSlot(str)
    def on_log(self, message):
        self.text_log.appendPlainText(message)

    @pyqtSlot(QSystemTrayIcon.ActivationReason)
    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            if self.isHidden():
                self.show()
                self.activateWindow()
                self.tray.hide()

    def changeEvent(self, event):
        if event.type() == QEvent.Type.WindowStateChange:
            if self.windowState() & Qt.WindowState.WindowMinimized:
                event.ignore()
                self.hide()
                self.tray.show()

    def closeEvent(self, event):
        result = QMessageBox.question(
            self, "警告", "是否确认退出? ",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if result == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()

    @pyqtSlot()
    def _validate_email(self):
        text = self.edit_email.text().strip()
        if not text:
            return
        for email in text.split(";"):
            if not re.match(
                r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", email
            ):
                self._msgbox("error", "邮箱格式错误, 请重新输入 !")
                self.edit_email.clear()
                self.edit_email.setFocus()
                return
        self.edit_email.setToolTip(text)

    @pyqtSlot()
    def _on_date_changed(self):
        ordfrom = self.date_from.date().toString("yyyy-MM-dd")
        ordtill = self.date_till.date().toString("yyyy-MM-dd")
        if ordfrom and ordtill:
            self.setWindowTitle(f"{APP_TITLE} - 订单日期范围: {ordfrom} 至 {ordtill}")

    @pyqtSlot()
    def set_schedule(self):
        self._on_date_changed()
        user = self.edit_user.text().strip()
        pwd = self.edit_pwd.text().strip()
        rec = self.edit_email.text().strip()
        timer = self.edit_time.text().strip()

        if not user or not pwd:
            self._msgbox("error", "请输入用户名和密码!! ")
            return
        if not rec:
            self._msgbox("error", "请输入邮箱地址!! ")
            return

        confirm = QMessageBox.question(
            self, "确认", f"是否设置定时任务: 每天{timer}? ",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        self.btn_cancel_sch.setEnabled(True)
        self.btn_schedule.setEnabled(False)
        self._on_date_changed()
        ordfrom = self.date_from.date().toString("yyyy-MM-dd")
        ordtill = self.date_till.date().toString("yyyy-MM-dd")
        chk_dld = "1" if self.chk_downloaded.isChecked() else "0"
        chk_workday = self.chk_workday.isChecked()

        self.thread = DownloadWorker()
        self.thread.log_signal.connect(self.on_log)
        self.thread.configure(
            ordfrom, ordtill, user, pwd, rec, chk_dld, False, timer, chk_workday,
        )
        self.thread.start()

    @pyqtSlot()
    def cancel_schedule(self):
        confirm = QMessageBox.question(
            self, "警告", "是否取消定时任务? ",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            if self.thread is not None:
                self.thread.stop_scheduler()
            self.text_log.appendPlainText("定时任务已取消")
            self.btn_cancel_sch.setEnabled(False)
            self.btn_schedule.setEnabled(True)

    @pyqtSlot()
    def execute_once(self):
        self._on_date_changed()
        user = self.edit_user.text().strip()
        pwd = self.edit_pwd.text().strip()
        rec = self.edit_email.text().strip()

        if not user or not pwd:
            self._msgbox("error", "请输入用户名和密码!! ")
            return
        if not rec:
            self._msgbox("error", "请输入邮箱地址!! ")
            return

        self.btn_stop.setEnabled(True)
        ordfrom = self.date_from.date().toString("yyyy-MM-dd")
        ordtill = self.date_till.date().toString("yyyy-MM-dd")
        chk_dld = "1" if self.chk_downloaded.isChecked() else "0"

        self.thread = DownloadWorker()
        self.thread.log_signal.connect(self.on_log)
        self.thread.configure(
            ordfrom, ordtill, user, pwd, rec, chk_dld, True, "", False,
        )
        self.thread.start()

    @pyqtSlot()
    def stop_thread(self):
        confirm = QMessageBox.question(
            self, "警告", "是否终止进程? ",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            if self.thread is not None:
                self.thread.stop()
            self.btn_stop.setEnabled(False)

    @pyqtSlot()
    def reset_log(self):
        confirm = QMessageBox.question(
            self, "警告", "是否清空日志? ",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if confirm == QMessageBox.StandardButton.Yes:
            self.text_log.clear()

    def _msgbox(self, title, text):
        tip = QMessageBox(self)
        if title == "error":
            tip.setIcon(QMessageBox.Icon.Critical)
        elif title == "DONE":
            tip.setIcon(QMessageBox.Icon.Warning)
        tip.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        font = QFont()
        font.setFamily("Microsoft YaHei")
        font.setPointSize(9)
        tip.setFont(font)
        tip.setText(text)
        tip.exec()


def main():
    app = QApplication.instance()
    if not app:
        app = QApplication(sys.argv)
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    app.setStyle("fusion")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
