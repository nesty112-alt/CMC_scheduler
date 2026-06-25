import sys
import os
import json
import uuid
import subprocess
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QFileDialog, QCheckBox, QTimeEdit, QDateEdit, QTableWidget,
                             QTableWidgetItem, QTextEdit, QLabel, QGroupBox, QHeaderView, QLineEdit,
                             QRadioButton, QComboBox, QStackedWidget, QSystemTrayIcon, QMenu, QAction, QStyle)
from PyQt5.QtCore import pyqtSignal, QObject, Qt, QDate, QSettings
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

import pythoncom
import win32com.client

CONFIG_FILE = "scheduler_config.json"
AUTOSTART_REG_KEY = "PythonMultiScheduler_v6"


class WorkerSignals(QObject):
    log_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str, str, str)


class MultiSchedulerApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("CMC 스케줄러")
        self.setGeometry(100, 100, 1150, 800)

        self.tasks = []

        self.scheduler = BackgroundScheduler()
        self.scheduler.start()
        self.signals = WorkerSignals()
        self.signals.log_signal.connect(self.append_log)
        self.signals.status_signal.connect(self.update_table_status)

        self.init_ui()
        self.init_tray_icon()  # 트레이 아이콘 초기화
        self.load_config()

    def init_ui(self) -> None:
        main_layout = QVBoxLayout()

        config_box = QGroupBox("새 작업 등록 및 시스템 설정")
        config_layout = QVBoxLayout()

        # [신규 추가] 시스템 설정 (자동 시작)
        sys_row = QHBoxLayout()
        self.autostart_cb = QCheckBox("PC 부팅 시 백그라운드로 자동 시작 (윈도우 시작프로그램 등록)")

        # 현재 레지스트리 상태 읽어와서 체크박스에 반영 (r 추가하여 이스케이프 경고 해결)
        settings = QSettings(r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run",
                             QSettings.NativeFormat)
        if settings.contains(AUTOSTART_REG_KEY):
            self.autostart_cb.setChecked(True)

        self.autostart_cb.stateChanged.connect(self.toggle_autostart)
        sys_row.addWidget(self.autostart_cb)
        sys_row.addStretch(1)
        config_layout.addLayout(sys_row)

        # 구분선 역할
        frame = QWidget()
        frame.setFixedHeight(1)
        frame.setStyleSheet("background-color: #ccc;")
        config_layout.addWidget(frame)

        # 1. 파일 선택
        file_row = QHBoxLayout()
        self.file_label = QLabel("선택된 파일 없음")
        btn_browse = QPushButton("파일 찾아보기")
        btn_browse.clicked.connect(self.browse_file)
        file_row.addWidget(btn_browse)
        file_row.addWidget(self.file_label, 1)
        config_layout.addLayout(file_row)

        # 2. 매크로 이름 입력
        macro_row = QHBoxLayout()
        macro_row.addWidget(QLabel("VBA 매크로명:"))
        self.macro_input = QLineEdit()
        self.macro_input.setPlaceholderText("엑셀(.xlsm, .xlsb) 파일 선택 시 필수 입력 (예: Module1.MyMacro)")
        self.macro_input.setEnabled(False)
        macro_row.addWidget(self.macro_input, 1)
        config_layout.addLayout(macro_row)

        # 3. 실행 기간 설정
        period_row = QHBoxLayout()
        period_row.addWidget(QLabel("실행 기간:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDate(QDate.currentDate())
        period_row.addWidget(self.start_date_edit)
        period_row.addWidget(QLabel(" ~ "))
        self.use_end_date_cb = QCheckBox("종료일 지정")
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDate(QDate.currentDate().addDays(30))
        self.end_date_edit.setEnabled(False)
        self.use_end_date_cb.toggled.connect(self.end_date_edit.setEnabled)
        period_row.addWidget(self.use_end_date_cb)
        period_row.addWidget(self.end_date_edit)
        period_row.addStretch(1)
        config_layout.addLayout(period_row)

        # 4. 실행 방식 (매주 / 매월) 선택 라디오 버튼
        type_row = QHBoxLayout()
        type_row.addWidget(QLabel("반복 방식:"))
        self.radio_weekly = QRadioButton("매주 (특정 요일)")
        self.radio_monthly = QRadioButton("매월 (특정 날짜)")
        self.radio_weekly.setChecked(True)
        type_row.addWidget(self.radio_weekly)
        type_row.addWidget(self.radio_monthly)
        type_row.addStretch(1)
        config_layout.addLayout(type_row)

        # 5. 실행 상세 조건 (스택 위젯)
        self.schedule_stack = QStackedWidget()

        self.weekly_widget = QWidget()
        day_layout = QHBoxLayout(self.weekly_widget)
        day_layout.setContentsMargins(0, 0, 0, 0)
        day_layout.addWidget(QLabel("실행 요일:"))
        self.day_mapping = {
            "월": "mon", "화": "tue", "수": "wed", "목": "thu", "금": "fri", "토": "sat", "일": "sun"
        }
        self.day_checkboxes = {}
        for kor, eng in self.day_mapping.items():
            cb = QCheckBox(kor)
            day_layout.addWidget(cb)
            self.day_checkboxes[eng] = cb
        day_layout.addStretch(1)
        self.schedule_stack.addWidget(self.weekly_widget)

        self.monthly_widget = QWidget()
        month_layout = QHBoxLayout(self.monthly_widget)
        month_layout.setContentsMargins(0, 0, 0, 0)
        month_layout.addWidget(QLabel("실행 날짜:"))
        self.month_day_combo = QComboBox()
        for i in range(1, 32):
            self.month_day_combo.addItem(f"매월 {i}일", str(i))
        self.month_day_combo.addItem("매월 말일(마지막 날)", "last")
        month_layout.addWidget(self.month_day_combo)
        month_layout.addStretch(1)
        self.schedule_stack.addWidget(self.monthly_widget)

        self.radio_weekly.toggled.connect(lambda: self.schedule_stack.setCurrentWidget(self.weekly_widget))
        self.radio_monthly.toggled.connect(lambda: self.schedule_stack.setCurrentWidget(self.monthly_widget))

        config_layout.addWidget(self.schedule_stack)

        # 6. 시간 선택 및 등록
        time_row = QHBoxLayout()
        time_row.addWidget(QLabel("실행 시간:"))
        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        btn_add = QPushButton("스케줄 추가")
        btn_add.clicked.connect(self.add_task)
        time_row.addWidget(self.time_edit)
        time_row.addWidget(btn_add, 1)
        config_layout.addLayout(time_row)

        config_box.setLayout(config_layout)
        main_layout.addWidget(config_box)

        # 테이블
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(["작업 대상 (파일/매크로)", "실행 기간", "실행 주기", "상태", "마지막 실행", "다음 실행", "관리"])
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        main_layout.addWidget(self.table)

        # 로그
        log_box = QGroupBox("실시간 실행 로그")
        log_layout = QVBoxLayout()
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        log_layout.addWidget(self.log_view)
        log_box.setLayout(log_layout)
        main_layout.addWidget(log_box, 1)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        self.current_selected_file = ""

    def init_tray_icon(self) -> None:
        # 시스템 기본 아이콘 사용
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

        # 우클릭 메뉴 설정
        tray_menu = QMenu()

        show_action = QAction("관리자 화면 열기", self)
        show_action.triggered.connect(self.showNormal)
        tray_menu.addAction(show_action)

        quit_action = QAction("완전히 종료", self)
        quit_action.triggered.connect(self.quit_application)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)

        # 더블클릭 이벤트 연결
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.showNormal()
            self.activateWindow()

    def toggle_autostart(self, state):
        # r 추가하여 이스케이프 경고 해결
        settings = QSettings(r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run",
                             QSettings.NativeFormat)
        if state == Qt.Checked:
            # exe로 패키징되었는지 파이썬 스크립트 상태인지 구분하여 경로 등록
            if getattr(sys, 'frozen', False):
                cmd_path = f'"{sys.executable}" --hidden'
            else:
                cmd_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}" --hidden'

            settings.setValue(AUTOSTART_REG_KEY, cmd_path)
            self.append_log("[시스템] PC 부팅 시 자동 시작이 설정되었습니다.")
        else:
            settings.remove(AUTOSTART_REG_KEY)
            self.append_log("[시스템] 부팅 시 자동 시작이 해제되었습니다.")

    def closeEvent(self, event) -> None:
        # 'X' 버튼을 눌렀을 때 완전히 종료되지 않고 트레이로 숨김
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "백그라운드 실행 중",
            "스케줄러가 시스템 트레이에서 계속 작동합니다.",
            QSystemTrayIcon.Information,
            2000
        )

    def quit_application(self) -> None:
        # 트레이에서 '완전히 종료'를 눌렀을 때만 실제 종료 프로세스 수행
        if self.scheduler.running:
            self.scheduler.shutdown()
        QApplication.quit()

    def browse_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self, "파일 선택", "",
            "지원하는 파일 (*.py *.exe *.xlsm *.xlsb);;Python Files (*.py *.exe);;Excel Macro Files (*.xlsm *.xlsb)"
        )
        if file_path:
            self.current_selected_file = file_path
            self.file_label.setText(os.path.basename(file_path))
            if file_path.lower().endswith(('.xlsm', '.xlsb')):
                self.macro_input.setEnabled(True)
                self.macro_input.setFocus()
            else:
                self.macro_input.setEnabled(False)
                self.macro_input.clear()

    def add_task(self) -> None:
        if not self.current_selected_file:
            self.append_log("[경고] 먼저 실행할 파일을 선택해주세요.")
            return

        macro_name = self.macro_input.text().strip()
        if self.current_selected_file.lower().endswith(('.xlsm', '.xlsb')) and not macro_name:
            self.append_log("[경고] 엑셀 파일을 스케줄링하려면 실행할 VBA 매크로 이름을 반드시 입력해야 합니다.")
            return

        schedule_type = "weekly" if self.radio_weekly.isChecked() else "monthly"
        cron_dow = "*"
        cron_day = "*"
        display_cycle = ""

        if schedule_type == "weekly":
            selected_days = [eng for eng, cb in self.day_checkboxes.items() if cb.isChecked()]
            if not selected_days:
                self.append_log("[경고] 매주 실행의 경우 최소 하나 이상의 요일을 선택해주세요.")
                return
            cron_dow = ",".join(selected_days)
            display_cycle = f"매주({cron_dow})"
        else:
            cron_day = self.month_day_combo.currentData()
            display_text = self.month_day_combo.currentText()
            display_cycle = f"매월({display_text})"

        time_obj = self.time_edit.time()
        hour = time_obj.hour()
        minute = time_obj.minute()

        start_date_str = self.start_date_edit.date().toString("yyyy-MM-dd")
        end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd") if self.use_end_date_cb.isChecked() else None

        self.register_task(
            self.current_selected_file, macro_name, start_date_str, end_date_str,
            schedule_type, cron_dow, cron_day, hour, minute, display_cycle, save=True
        )

    def register_task(self, file_path: str, macro_name: str, start_date: str, end_date: str,
                      schedule_type: str, cron_dow: str, cron_day: str, hour: int, minute: int,
                      display_cycle: str, save: bool = True) -> None:
        job_id = str(uuid.uuid4())

        try:
            if schedule_type == "weekly":
                trigger = CronTrigger(day_of_week=cron_dow, hour=hour, minute=minute, start_date=start_date,
                                      end_date=end_date)
            else:
                trigger = CronTrigger(day=cron_day, hour=hour, minute=minute, start_date=start_date, end_date=end_date)

            job = self.scheduler.add_job(self.execute_script, trigger=trigger, id=job_id,
                                         args=[file_path, macro_name, job_id])
        except ValueError as e:
            self.append_log(f"[설정 오류] 스케줄을 등록할 수 없습니다: {str(e)}")
            return

        row_index = self.table.rowCount()
        self.table.insertRow(row_index)

        display_name = os.path.basename(file_path)
        if macro_name: display_name += f" [{macro_name}]"

        file_item = QTableWidgetItem(display_name)
        file_item.setData(Qt.UserRole, job_id)

        # \n 을 사용하여 한 줄로 작성해 f-string 오류 해결
        file_item.setToolTip(f"{file_path}\n매크로: {macro_name}")

        self.table.setItem(row_index, 0, file_item)
        period_str = f"{start_date} ~ {end_date if end_date else '제한 없음'}"
        self.table.setItem(row_index, 1, QTableWidgetItem(period_str))
        self.table.setItem(row_index, 2, QTableWidgetItem(f"{display_cycle} {hour:02d}:{minute:02d}"))
        self.table.setItem(row_index, 3, QTableWidgetItem("대기 중"))
        self.table.setItem(row_index, 4, QTableWidgetItem("-"))

        next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job.next_run_time else "기간 만료"
        self.table.setItem(row_index, 5, QTableWidgetItem(next_run))

        btn_delete = QPushButton("삭제")
        btn_delete.setProperty("job_id", job_id)
        btn_delete.clicked.connect(self.delete_task)
        self.table.setCellWidget(row_index, 6, btn_delete)

        self.tasks.append({
            "file_path": file_path, "macro_name": macro_name, "start_date": start_date, "end_date": end_date,
            "schedule_type": schedule_type, "cron_dow": cron_dow, "cron_day": cron_day, "hour": hour, "minute": minute,
            "job_id": job_id
        })

        if save:
            self.save_config()
            self.append_log(f"[등록 완료] {display_name} 스케줄이 지정되었습니다.")

    def delete_task(self) -> None:
        button = self.sender()
        if not button: return
        job_id = button.property("job_id")
        try:
            if self.scheduler.get_job(job_id): self.scheduler.remove_job(job_id)
        except Exception:
            pass

        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.UserRole) == job_id:
                file_name = item.text()
                self.table.removeRow(row)
                self.append_log(f"[삭제 완료] {file_name} 스케줄이 삭제되었습니다.")
                break

        self.tasks = [t for t in self.tasks if t["job_id"] != job_id]
        self.save_config()

    def execute_script(self, file_path: str, macro_name: str, job_id: str) -> None:
        file_name = os.path.basename(file_path)
        start_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.signals.log_signal.emit(f"[{start_time}] [시작] {file_name} 작업을 실행합니다.")
        self.signals.status_signal.emit(job_id, "실행 중", start_time)

        try:
            if not os.path.exists(file_path): raise FileNotFoundError("파일이 존재하지 않습니다.")

            if file_path.lower().endswith(('.xlsm', '.xlsb')):
                pythoncom.CoInitialize()
                excel = None
                wb = None
                try:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    abs_path = os.path.abspath(file_path)
                    wb = excel.Workbooks.Open(abs_path)
                    excel.Application.Run(f"'{abs_path}'!{macro_name}")
                    wb.Save()
                    self.signals.log_signal.emit(f"[{file_name}] VBA 매크로({macro_name}) 정상 종료 및 저장 완료")
                finally:
                    if wb: wb.Close(SaveChanges=False)
                    if excel: excel.Quit()
                    pythoncom.CoUninitialize()
            else:
                if file_path.endswith('.py'):
                    # exe로 패키징된 상태인지 체크하여 Python 스크립트 실행 분기 처리
                    if getattr(sys, 'frozen', False):
                        result = subprocess.run(["python", file_path], capture_output=True, text=True, check=True)
                    else:
                        result = subprocess.run([sys.executable, file_path], capture_output=True, text=True, check=True)
                else:
                    result = subprocess.run([file_path], capture_output=True, text=True, check=True)

                if result.stdout:
                    self.signals.log_signal.emit(f"[{file_name} 출력]:\n{result.stdout.strip()}")

            end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.signals.log_signal.emit(f"[{end_time}] [성공] {file_name} 작업 완료")
            job = self.scheduler.get_job(job_id)
            next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job and job.next_run_time else "기간 만료"
            self.signals.status_signal.emit(job_id, "정상 종료", next_run)

        except Exception as e:
            end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.signals.log_signal.emit(f"[{end_time}] [에러 발생] {file_name} 실패: {str(e)}")
            job = self.scheduler.get_job(job_id)
            next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job and job.next_run_time else "기간 만료"
            self.signals.status_signal.emit(job_id, "오류 발생", next_run)

    def append_log(self, text: str) -> None:
        self.log_view.append(text)

    def update_table_status(self, job_id: str, status: str, time_info: str) -> None:
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.UserRole) == job_id:
                self.table.setItem(row, 3, QTableWidgetItem(status))
                self.table.setItem(row, 4 if status == "실행 중" else 5, QTableWidgetItem(time_info))
                break

    def save_config(self) -> None:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=4)
        except:
            pass

    def load_config(self) -> None:
        if not os.path.exists(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config_data = json.load(f)
                for item in config_data:
                    start_date = item.get("start_date", QDate.currentDate().toString("yyyy-MM-dd"))
                    end_date = item.get("end_date", None)
                    macro_name = item.get("macro_name", "")
                    schedule_type = item.get("schedule_type", "weekly")
                    cron_dow = item.get("cron_dow", item.get("day_str", "*"))
                    cron_day = item.get("cron_day", "*")

                    if schedule_type == "weekly":
                        display_cycle = f"매주({cron_dow})"
                    else:
                        display_cycle = f"매월(매월 {cron_day}일)" if cron_day != 'last' else "매월(매월 말일(마지막 날))"

                    self.register_task(item["file_path"], macro_name, start_date, end_date, schedule_type, cron_dow,
                                       cron_day, item["hour"], item["minute"], display_cycle, save=False)
            self.append_log(f"[시스템] 등록된 {len(config_data)}개의 스케줄 설정을 로드했습니다.")
        except:
            pass


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # 마지막 창이 닫혀도 프로그램이 종료되지 않게 설정 (트레이 아이콘 유지를 위해 필수)
    app.setQuitOnLastWindowClosed(False)

    window = MultiSchedulerApp()

    # --hidden 인자가 있으면 창을 띄우지 않고 트레이에서만 실행 (부팅 시 자동 시작용)
    if "--hidden" not in sys.argv:
        window.show()

    sys.exit(app.exec_())