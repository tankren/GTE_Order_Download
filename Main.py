# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 21:05:37 2022

@author: REC3WX

vK8FXz+w~UxHR2L

"""

from datetime import datetime, timedelta
from PySide6.QtWidgets import *
from PySide6.QtGui import QFont
from PySide6.QtCore import Slot, Qt, QThread, Signal, QEvent, QDate
import sys
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
import re
import os
from lxml import etree
import tempfile
from apscheduler.schedulers.background import BackgroundScheduler


class Worker(QThread):
    sinOut = Signal(str)

    def __init__(self, parent=None):
        super(Worker, self).__init__(parent)
        self.folder = f"{tempfile.gettempdir()}\GET\Order"
        if os.path.exists(self.folder):
            for zip in os.listdir(self.folder):
                os.remove(f"{self.folder}\\{zip}")
        else:
            os.makedirs(self.folder)

    def stop_self(self):
        self.terminate()
        message = f"进程已终止..."
        self.sinOut.emit(message)

    def getdata(
        self, ordfrom, ordtill, user, pwd, rec, chk_dld, once, timer, chk_workday
    ):
        self.ordfrom = ordfrom
        self.ordtill = ordtill
        self.user = user
        self.pwd = pwd
        self.rec = rec
        self.chk_dld = chk_dld
        self.once = once
        self.timer = timer
        self.chk_workday = chk_workday
        self.supplier = self.to_unicode(self.user[0:5])

    def send_mail(self):
        mail_host = "smtp.163.com"
        mail_user = "vhcn_lop@163.com"
        mail_pass = "YQXQMPJDDTSOJLOC"
        # 邮件发送方邮箱地址
        sender = "vhcn_lop@163.com"
        # 邮件接受方邮箱地址，注意需要[]包裹，这意味着你可以写多个邮件地址群发
        receivers = self.rec
        # 设置email信息
        # 邮件内容设置
        email = MIMEMultipart()
        # 邮件主题
        time = datetime.now().isoformat(" ", "seconds")
        email["Subject"] = f"no-reply: {time} GTE订单自动下载结果"
        # 发送方信息
        email["From"] = formataddr(["Mail Bot@VHCN", sender])
        # 接受方信息
        email["To"] = receivers

        # add attachments
        file_list = [f for f in os.listdir(self.folder) if f.endswith(".zip")]
        for file in file_list:
            part = MIMEApplication(open(f"{self.folder}\\{file}", "rb").read())
            part.add_header("Content-Disposition", "attachment", filename=file)
            email.attach(part)
        # dynamic email content
        if not file_list and self.once == "1":
            message = f"单次执行无订单不发送邮件!"
            self.sinOut.emit(message)
            return
        elif not file_list:
            email_content = f"Dear all,\n\tGTE {QDate.currentDate().toString('yyyy-MM-dd')} 没有新的订单, 请知悉! \n\nVHCN ICO"
        else:
            email_content = f"Dear all,\n\t附件为GTE {QDate.currentDate().toString('yyyy-MM-dd')} 订单, 请查收! \n\nVHCN ICO"
        # insert content
        email.attach(MIMEText(email_content, "plain", "utf-8"))

        try:
            smtpObj = smtplib.SMTP_SSL(mail_host, 465)
            # 登录到服务器
            smtpObj.login(mail_user, mail_pass)
            # 发送
            smtpObj.sendmail(sender, receivers.split(";"), email.as_string())
            # 退出
            smtpObj.quit()
            message = f"邮件发送成功..."
            self.sinOut.emit(message)
        except smtplib.SMTPException as e:
            message = f"{e}"
            self.sinOut.emit(message)

    def to_unicode(self, text):
        ret = ""
        for v in text:
            ret = ret + hex(ord(v)).lower().replace("0x", "\\\\u")
        return ret

    def first_download(self, filename):  # filename不含后缀.zip
        ufilename = self.to_unicode(self.OrderNo)
        url = "http://192.168.10.33/WCFService/WcfService.svc"
        headers = {
            "content-type": "text/xml",
            "Referer": "http://192.168.10.33/ClientBin/SilverlightUI.xap",
            "SOAPAction": '"SysManager/WcfService/UpeFirstDownLoadTimeIPO0710"',
        }
        prms = f'<prms>{{"OrderNo":"{ufilename}"}}</prms>'
        body = (
            '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">'
            "<s:Body>"
            '<UpeFirstDownLoadTimeIPO0710 xmlns="SysManager">'
            '<baseInfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            "</baseInfo>"
            f"{prms}"
            "</UpeFirstDownLoadTimeIPO0710>"
            "</s:Body>"
            "</s:Envelope>"
        )

        self.resp = requests.post(
            url, data=body, headers=headers, timeout=5
        ).content.decode("utf8")

    def post_download(self):
        url = "http://192.168.10.33/WCFService/WcfService.svc"
        headers = {
            "content-type": "text/xml",
            "Referer": "http://192.168.10.33/ClientBin/SilverlightUI.xap",
            "SOAPAction": '"SysManager/WcfService/IPO0710GetInfo"',
        }
        if self.chk_dld == "1" and self.once == "1":  # 考虑已下载且手动执行，才考虑日期范围，否则只查询未下载的所有订单
            scope = f'<parm>{{"SupplierNum":"{self.supplier}","Series":"-1","ActualOrderTimeS":"{self.ordfrom}","ActualOrderTimeE":"{self.ordtill}","Check":"{self.chk_dld}"}}</parm>'
        else:
            scope = f'<parm>{{"SupplierNum":"{self.supplier}","Series":"-1","Check":"{self.chk_dld}"}}</parm>'

        body = (
            '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">'
            "<s:Body>"
            '<IPO0710GetInfo xmlns="SysManager">'
            '<baseInfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            "</baseInfo>"
            '<pagerinfo xmlns:d4p1="http://schemas.datacontract.org/2004/07/Silverlight.BaseDTO.Entities" xmlns:i="http://www.w3.org/2001/XMLSchema-instance">'
            "<d4p1:pageIndex>1</d4p1:pageIndex>"
            "<d4p1:pageSize>100</d4p1:pageSize>"
            "</pagerinfo>"
            f"{scope}"
            "</IPO0710GetInfo>"
            "</s:Body>"
            "</s:Envelope>"
        )
        self.response = requests.post(url, data=body, headers=headers, timeout=5)
        self.resp_str = str(self.response.content.decode("utf8"))
        xml = etree.fromstring(self.resp_str)
        message = f"临时下载目录为: {self.folder}"
        self.sinOut.emit(message)
        for filenm in xml.iter("{*}FileNm"):
            try:
                message = f"命中: {filenm.text}"
                self.sinOut.emit(message)
                for ancestor in filenm.iterancestors("{*}IPO0710DTO"):
                    self.OrderNo = ancestor.find("./{*}OrderNo").text
                dld_path = f"http://192.168.10.33/GTESGFile/Instruct/3334-A/{filenm.text[7:15]}/{filenm.text}"
                # e.g.: http://192.168.10.33/GTESGFile/Instruct/3334-A/20220908/M23334A2022090805ZL.zip
                message = f"开始下载Zip文件 {filenm.text}"
                self.sinOut.emit(message)
                self.filepath = f"{self.folder}\\{filenm.text}"
                dld_zip = requests.get(url=dld_path, stream=True)
                with open(self.filepath, "wb") as zip:
                    zip.write(dld_zip.content)
                message = f"下载完成! 回传下载时间"
                self.sinOut.emit(message)
                self.first_download(filenm.text)  # 下载完成就回传下载时间

            except Exception as e:
                message = f"{e}"
                self.sinOut.emit(message)
        else:
            message = f"没有找到新的订单..."
            self.sinOut.emit(message)

    def chain(self):
        try:
            if self.once == "0":
                message = f"{datetime.now()} - 开始运行定时任务..."
                self.sinOut.emit(message)
            else:
                message = f"{datetime.now()} - 开始运行单次任务..."
                self.sinOut.emit(message)

            message = f"获取Zip文件清单..."
            self.sinOut.emit(message)
            self.post_download()
            message = f"获取结束, 发送邮件..."
            self.sinOut.emit(message)
            self.send_mail()
            if self.once == "0":
                message = f"{datetime.now()} - 定时任务已完成! "
                self.sinOut.emit(message)
                self.time_gap()
                message = f"下次任务时间为: {self.sch_dhm}, 距现在{self.gap_h}小时{self.gap_m}分"
                self.sinOut.emit(message)
            else:
                message = f"{datetime.now()} - 单次任务已完成! "
                self.sinOut.emit(message)

        except Exception as e:
            message = f"{e}"
            self.sinOut.emit(message)

    def stop_scheduler(self):
        message = f"正在终止定时任务..."
        self.sinOut.emit(message)
        try:
            self.scheduler.shutdown()
        except Exception as e:
            message = f"{e}"
            self.sinOut.emit(message)

    def time_gap(self):
        now = datetime.now()
        now_hm = now.strftime("%H:%M")
        now_d = now.strftime("%Y-%m-%d")
        now_dhm = now.strftime("%Y-%m-%d %H:%M")
        # sch_hm =input('输入时间(格式09:00): ')
        if now.weekday() == 4:
            delta = 3
        elif now.weekday() == 5:
            delta = 2
        else:
            delta = 1
        sch_hm = self.timer
        if self.chk_workday == "5":
            if now_hm < sch_hm:
                self.sch_dhm = f"{now_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                self.gap_h = gap.split(":")[0]
                self.gap_m = gap.split(":")[1]
            else:
                sch_d = datetime.strftime((now + timedelta(days=delta)), "%Y-%m-%d")
                self.sch_dhm = f"{sch_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                if "day" in str(gap):
                    self.gap_d = gap[0]
                    self.gap_h = int(gap.split(":")[0][-2:]) + 24 * int(self.gap_d)
                else:
                    self.gap_h = int(gap.split(":")[0])
                self.gap_m = gap.split(":")[1]
        else:
            if now_hm < sch_hm:
                self.sch_dhm = f"{now_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                self.gap_h = gap.split(":")[0]
                self.gap_m = gap.split(":")[1]
            else:
                sch_d = datetime.strftime((now + timedelta(days=delta)), "%Y-%m-%d")
                self.sch_dhm = f"{sch_d} {sch_hm}"
                gap = str(
                    datetime.strptime(self.sch_dhm, "%Y-%m-%d %H:%M")
                    - datetime.strptime(now_dhm, "%Y-%m-%d %H:%M")
                )
                self.gap_h = gap.split(":")[0]
                self.gap_m = gap.split(":")[1]

    def run(self):
        # 主逻辑
        if self.once == "1":
            try:
                self.chain()
            except Exception as e:
                message = f"{e}"
                self.sinOut.emit(message)
        else:
            try:
                self.scheduler = BackgroundScheduler(timezone="Asia/Shanghai")
                h = self.timer[0:2]
                m = self.timer[3:5]
                if self.chk_workday == "7":
                    day_of_week = "*"
                else:
                    day_of_week = "mon-fri"
                try:
                    self.scheduler.add_job(
                        self.chain,
                        trigger="cron",
                        day_of_week=day_of_week,
                        hour=h,
                        minute=m,
                    )

                    self.scheduler.start()
                    if day_of_week == "*":
                        message = f"已设置定时任务为每天{self.timer} "
                    else:
                        message = f"已设置定时任务为每个工作日的{self.timer} "
                    self.sinOut.emit(message)
                    self.time_gap()
                    message = f"下次任务时间为: {self.sch_dhm}, 距现在{self.gap_h}小时{self.gap_m}分"
                    self.sinOut.emit(message)
                except Exception as e:
                    message = f"{e}"
                    self.sinOut.emit(message)
            except Exception as e:
                message = f"{e}"
                self.sinOut.emit(message)


class MyWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = Worker()
        self.title = f"GTE订单查询下载工具 v0.1   - Made by REC3WX"
        self.setWindowTitle(self.title)
        pixmapi = QStyle.SP_FileDialogDetailedView
        icon = self.style().standardIcon(pixmapi)
        self.setWindowIcon(icon)
        self.setFixedSize(700, 300)

        self.tray = QSystemTrayIcon()
        self.tray.setIcon(icon)
        self.tray.setToolTip(f"{self.title} 正在后台运行")
        self.tray.activated.connect(self.on_systemTrayIcon_activated)

        self.fld_user = QLabel("用户名:")
        self.line_user = QLineEdit()
        self.line_user.setClearButtonEnabled(True)
        self.line_user.setText("3334A01")  # 测试
        self.fld_pwd = QLabel("密码:")
        self.line_pwd = QLineEdit()
        self.line_pwd.setClearButtonEnabled(True)
        self.line_pwd.setText("123456")  # 测试
        self.line_pwd.setEchoMode(QLineEdit.PasswordEchoOnEdit)

        self.fld_ordfrom = QLabel("开始日期:")
        self.cb_ordfrom = QDateEdit(QDate.currentDate())
        self.cb_ordfrom.setCursor(Qt.PointingHandCursor)
        self.cb_ordfrom.setDisplayFormat("yyyy-MM-dd")
        self.cb_ordfrom.setCalendarPopup(True)
        self.cb_ordfrom.calendarWidget().setGridVisible(True)
        self.cb_ordfrom.calendarWidget().setVerticalHeaderFormat(
            QCalendarWidget.ISOWeekNumbers
        )
        self.cb_ordfrom.dateChanged.connect(self.get_ordfrom_ordtill)

        self.fld_ordtill = QLabel("结束日期:")
        self.cb_ordtill = QDateEdit(QDate.currentDate().addDays(5))
        self.cb_ordtill.setCursor(Qt.PointingHandCursor)
        self.cb_ordtill.setDisplayFormat("yyyy-MM-dd")
        self.cb_ordtill.setCalendarPopup(True)
        self.cb_ordtill.calendarWidget().setGridVisible(True)
        self.cb_ordtill.calendarWidget().setVerticalHeaderFormat(
            QCalendarWidget.ISOWeekNumbers
        )
        self.cb_ordtill.dateChanged.connect(self.get_ordfrom_ordtill)

        self.chk_dld = QCheckBox("包含已下载", self)
        self.chk_dld.setCursor(Qt.PointingHandCursor)
        self.chk_workday = QCheckBox("仅工作日", self)
        self.chk_workday.setCursor(Qt.PointingHandCursor)

        self.btn_start = QPushButton("运行单次任务")
        self.btn_start.clicked.connect(self.execute_once)

        self.fld_sch = QLabel("定时任务时间:")
        self.time_sch = QTimeEdit()
        self.time_sch.setDisplayFormat("HH:mm")
        self.time_sch.setCursor(Qt.PointingHandCursor)

        self.fld_email = QLabel("收件人:")
        self.line_email = QLineEdit()
        self.line_email.setClearButtonEnabled(True)
        self.line_email.setPlaceholderText("多个收件人之间用分号;分开")
        self.line_email.setText("chenlong.ren@cn.bosch.com;feng.he@cn.bosch.com")  # 测试
        self.line_email.setToolTip(self.line_email.text())
        self.line_email.editingFinished.connect(self.check_email)

        self.fld_result = QLabel("运行日志:")
        self.text_result = QPlainTextEdit()
        self.text_result.setReadOnly(True)

        self.line = QFrame()
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)

        self.btn_schedule = QPushButton("设置定时任务")
        self.btn_schedule.clicked.connect(self.set_schedule)
        self.btn_cnl_sch = QPushButton("取消定时任务")
        self.btn_cnl_sch.clicked.connect(self.cancel_schedule)
        self.btn_cnl_sch.setEnabled(False)

        self.btn_reset = QPushButton("清空日志")
        self.btn_reset.clicked.connect(self.reset_log)
        self.btn_stop = QPushButton("终止任务进程")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_thread)

        self.layout = QGridLayout()
        self.layout.addWidget((self.fld_user), 0, 0)
        self.layout.addWidget((self.line_user), 0, 1)
        self.layout.addWidget((self.fld_pwd), 0, 2)
        self.layout.addWidget((self.line_pwd), 0, 3)
        self.layout.addWidget((self.fld_ordfrom), 0, 4)
        self.layout.addWidget((self.cb_ordfrom), 0, 5)
        self.layout.addWidget((self.fld_ordtill), 1, 4)
        self.layout.addWidget((self.cb_ordtill), 1, 5)
        self.layout.addWidget((self.fld_email), 1, 0)
        self.layout.addWidget((self.line_email), 1, 1, 1, 3)
        self.layout.addWidget((self.fld_sch), 0, 6)
        self.layout.addWidget((self.time_sch), 0, 7)
        self.layout.addWidget((self.chk_dld), 1, 6)
        self.layout.addWidget((self.chk_workday), 1, 7)

        self.layout.addWidget((self.line), 2, 0, 1, 8)
        self.layout.addWidget((self.fld_result), 3, 0)
        self.layout.addWidget((self.text_result), 4, 0, 6, 6)
        self.layout.addWidget((self.btn_reset), 4, 6, 1, 2)
        self.layout.addWidget((self.btn_start), 6, 6, 1, 2)
        # self.layout.addWidget((self.btn_stop), 7, 6, 1, 2)
        self.layout.addWidget((self.btn_schedule), 8, 6, 1, 2)
        self.layout.addWidget((self.btn_cnl_sch), 9, 6, 1, 2)

        self.setLayout(self.layout)

        self.thread.sinOut.connect(self.Addmsg)  # 解决重复emit

    @Slot()
    def on_systemTrayIcon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            if self.isHidden():
                self.show()
                self.activateWindow()
                self.tray.hide()

    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                event.ignore()
                self.hide()
                self.tray.show()

    def closeEvent(self, event):
        result = QMessageBox.question(
            self, "警告", "是否确认退出? ", QMessageBox.Yes | QMessageBox.No
        )
        if result == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def check_email(self):
        if not self.line_email.text() == "":
            email_list = self.line_email.text().split(";")
            for email in email_list:
                if not re.match(
                    r"(^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$)", email
                ):
                    self.msgbox("error", "邮箱格式错误, 请重新输入 !")
                    self.line_email.setText("")
                    self.line_email.setFocus()
        self.line_email.setToolTip(self.line_email.text())

    def Addmsg(self, message):
        self.text_result.appendPlainText(message)

    def stop_thread(self):
        confirm = QMessageBox.question(
            self, "警告", "是否终止进程? ", QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.thread.stop_self()
            self.btn_stop.setEnabled(False)

    def get_ordfrom_ordtill(self):
        self.ordfrom = self.cb_ordfrom.date().toString("yyyy-MM-dd")
        self.ordtill = self.cb_ordtill.date().toString("yyyy-MM-dd")
        if not self.ordfrom == "" and not self.ordtill == "":
            # self.text_result.appendPlainText(f'当前选择的发票年月为: {year}年{till}月')
            self.setWindowTitle(
                f"{self.title} - 订单日期范围: {self.ordfrom} 至 {self.ordtill}"
            )

    def set_schedule(self):
        year = str(self.cb_ordtill.currentText())
        till = str(self.cb_ordtill.currentText())
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        rec = str(self.line_email.text())
        timer = str(self.time_sch.text())
        if self.chk_dld.isChecked():
            chk_dld = "1"
        else:
            chk_dld = "0"

        if self.chk_workday.isChecked():
            chk_workday = "5"
        else:
            chk_workday = "7"

        if user == "" or pwd == "":
            self.msgbox("error", "请输入用户名和密码!! ")
        else:
            if not rec == "":
                confirm = QMessageBox.question(
                    self,
                    "确认",
                    f"是否设置定时任务: 每天{timer}? ",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if confirm == QMessageBox.Yes:
                    self.btn_cnl_sch.setEnabled(True)
                    self.btn_schedule.setEnabled(False)
                    # 订单开始日期，订单结束日期，用户名，密码，收件人，包含已下载，计划执行=0，定时时间，仅工作日）
                    self.thread.getdata(
                        year, till, user, pwd, rec, chk_dld, "0", timer, chk_workday
                    )
                    self.thread.start()
            else:
                self.msgbox("error", "请输入邮箱地址!! ")

    def cancel_schedule(self):
        confirm = QMessageBox.question(
            self, "警告", "是否取消定时任务? ", QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.thread.stop_scheduler()
            self.Addmsg(f"定时任务已取消")
            self.btn_cnl_sch.setEnabled(False)
            self.btn_schedule.setEnabled(True)

    def execute_once(self):
        self.get_ordfrom_ordtill()
        user = str(self.line_user.text())
        pwd = str(self.line_pwd.text())
        rec = str(self.line_email.text())
        if self.chk_dld.isChecked():
            chk_dld = "1"
        else:
            chk_dld = "0"
        if user == "" or pwd == "":
            self.msgbox("error", "请输入用户名和密码!! ")
        else:
            if not rec == "":
                self.btn_stop.setEnabled(True)
                # 订单开始日期，订单结束日期，用户名，密码，收件人，包含已下载，单次执行=1，定时时间，仅工作日）
                self.thread.getdata(
                    self.ordfrom, self.ordtill, user, pwd, rec, chk_dld, "1", "", ""
                )
                self.thread.start()
            else:
                self.msgbox("error", "请输入邮箱地址!! ")

    def msgbox(self, title, text):
        tip = QMessageBox(self)
        if title == "error":
            tip.setIcon(QMessageBox.Critical)
        elif title == "DONE":
            tip.setIcon(QMessageBox.Warning)
        tip.setWindowFlag(Qt.FramelessWindowHint)
        font = QFont()
        font.setFamily("Microsoft YaHei")
        font.setPointSize(9)
        tip.setFont(font)
        tip.setText(text)
        tip.exec()

    def reset_log(self):
        confirm = QMessageBox.question(
            self, "警告", "是否清空日志? ", QMessageBox.Yes | QMessageBox.No
        )
        if confirm == QMessageBox.Yes:
            self.text_result.clear()


def main():
    if not QApplication.instance():
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    font = QFont()
    font.setFamily("Microsoft YaHei")
    font.setPointSize(10)
    app.setFont(font)
    widget = MyWidget()
    # app.setStyleSheet(qdarktheme.load_stylesheet(border="rounded"))
    app.setStyle("fusion")

    widget.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
