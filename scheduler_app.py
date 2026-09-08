import json
import os
import re
import shutil
import subprocess
import sys
import threading
import uuid
from datetime import datetime

import pythoncom
import win32com.client
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger
from apscheduler.triggers.date import DateTrigger
from PyQt5.QtCore import (
    QDate,
    QLockFile,
    QObject,
    QSettings,
    Qt,
    QTime,
    QTimer,
    pyqtSignal,
)
from PyQt5.QtNetwork import QLocalServer, QLocalSocket
from PyQt5.QtWidgets import (
    QAction,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QStackedWidget,
    QStyle,
    QSystemTrayIcon,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)

CONFIG_FILE = "scheduler_config.json"
LOG_FILE = os.path.join("logs", "scheduler.log")
AUTOSTART_REG_KEY = "PythonMultiScheduler_v6"
INSTANCE_KEY = "CMC_Scheduler_SingleInstance"
UI_SETTINGS_ORG = "CMC"
UI_SETTINGS_APP = "Scheduler"

DAY_KO = {
    "mon": "월",
    "tue": "화",
    "wed": "수",
    "thu": "목",
    "fri": "금",
    "sat": "토",
    "sun": "일",
}
WEEKDAYS = ["mon", "tue", "wed", "thu", "fri"]
WEEKEND = ["sat", "sun"]

def parse_vba_macros(vba_code: str, module_name: str) -> list:
    if not vba_code:
        return []
    match = VB_NAME_RE.search(vba_code)
    if match:
        module_name = match.group(1)
    names = []
    for proc in MACRO_NAME_RE.findall(vba_code):
        if proc.lower() in SKIP_MACRO_NAMES:
            continue
        names.append(f"{module_name}.{proc}" if module_name else proc)
    return names


def list_excel_macros(file_path: str) -> list:
    found = _list_macros_oletools(file_path)
    if not found:
        found = _list_macros_excel_com(file_path)
    unique = []
    seen = set()
    for name in found:
        key = name.lower()
        if key not in seen:
            seen.add(key)
            unique.append(name)
    return unique


def _list_macros_oletools(file_path: str) -> list:
    try:
        from oletools.olevba import VBA_Parser
    except Exception:
        return []

    names = []
    parser = None
    try:
        parser = VBA_Parser(file_path)
        if not parser.detect_vba_macros():
            return []
        for _filename, _stream, vba_filename, vba_code in parser.extract_macros():
            module_name = os.path.splitext(os.path.basename(vba_filename or ""))[0]
            names.extend(parse_vba_macros(vba_code or "", module_name))
    except Exception:
        return []
    finally:
        if parser is not None:
            try:
                parser.close()
            except Exception:
                pass
    return names


def _list_macros_excel_com(file_path: str) -> list:
    names = []
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        wb = excel.Workbooks.Open(os.path.abspath(file_path), ReadOnly=True)
        for component in wb.VBProject.VBComponents:
            try:
                code_module = component.CodeModule
                line_count = code_module.CountOfLines
                source = code_module.Lines(1, line_count) if line_count else ""
            except Exception:
                continue
            names.extend(parse_vba_macros(source, component.Name))
    except Exception:
        return []
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
        except Exception:
            pass
    return names

MACRO_NAME_RE = re.compile(
    r'(?im)^[ \t]*(?:Public |Private |Friend )?(?:Static )?(?:Sub|Function)[ \t]+'
    r'([A-Za-z가-힣_][\w가-힣]*)'
)
VB_NAME_RE = re.compile(r'(?im)^Attribute VB_Name\s*=\s*"([^"]+)"')
SKIP_MACRO_NAMES = {
    "class_initialize",
    "class_terminate",
}

def find_python_executable() -> list:
    if not getattr(sys, "frozen", False):
        return [sys.executable]
    for name in ("python", "pythonw"):
        path = shutil.which(name)
        if path:
            return [path]
    py_launcher = shutil.which("py")
    if py_launcher:
        return [py_launcher, "-3"]
    raise FileNotFoundError(
        "Python 인터프리터를 찾을 수 없습니다. .py 파일을 실행하려면 PC에 Python이 설치되어 있고 PATH에 등록되어 있어야 합니다."
    )


def app_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def ensure_workdir() -> str:
    base = app_dir()
    os.chdir(base)
    os.makedirs("logs", exist_ok=True)
    return base


def local_tz():
    try:
        from tzlocal import get_localzone
        return get_localzone()
    except Exception:
        return None


def weekly_cycle_label(cron_dow: str) -> str:
    parts = [p.strip() for p in cron_dow.split(",") if p.strip()]
    if not parts or parts == ["*"]:
        return "매주"
    labels = [DAY_KO.get(p, p) for p in parts]
    return "매주(" + ", ".join(labels) + ")"


def status_item(text: str) -> QTableWidgetItem:
    item = QTableWidgetItem(text)
    item.setTextAlignment(Qt.AlignCenter)
    return item


def append_log_file(text: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass


class WorkerSignals(QObject):
    log_signal = pyqtSignal(str)
    status_signal = pyqtSignal(str, str, str)


class MultiSchedulerApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("CMC 스케줄러")
        self.setGeometry(100, 100, 1150, 800)

        self.tasks = []
        self.tz = local_tz()
        self.ui_settings = QSettings(UI_SETTINGS_ORG, UI_SETTINGS_APP)
        self.current_selected_file = ""

        scheduler_kwargs = {"timezone": self.tz} if self.tz else {}
        self.scheduler = BackgroundScheduler(**scheduler_kwargs)
        self.scheduler.start()
        self.signals = WorkerSignals()
        self.signals.log_signal.connect(self.append_log)
        self.signals.status_signal.connect(self.update_table_status)

        self.init_ui()
        self.init_tray_icon()
        self.init_single_instance_server()
        self.load_config()
        self.refresh_status_bar()

        self.status_timer = QTimer(self)
        self.status_timer.timeout.connect(self.refresh_status_bar)
        self.status_timer.start(15000)

        self.append_log(f"[시스템] 작업 폴더: {os.getcwd()}")

    def init_single_instance_server(self) -> None:
        QLocalServer.removeServer(INSTANCE_KEY)
        self.instance_server = QLocalServer(self)
        self.instance_server.newConnection.connect(self._on_second_instance)
        self.instance_server.listen(INSTANCE_KEY)

    def _on_second_instance(self) -> None:
        conn = self.instance_server.nextPendingConnection()
        if conn:
            conn.readyRead.connect(lambda: None)
            conn.deleteLater()
        self.show_main_window()

    def show_main_window(self) -> None:
        self.showNormal()
        self.show()
        self.raise_()
        self.activateWindow()

    def init_ui(self) -> None:
        main_layout = QVBoxLayout()

        config_box = QGroupBox("새 작업 등록 및 시스템 설정")
        config_layout = QVBoxLayout()

        sys_row = QHBoxLayout()
        self.autostart_cb = QCheckBox("PC 부팅 시 백그라운드로 자동 시작 (윈도우 시작프로그램 등록)")
        settings = QSettings(
            r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run",
            QSettings.NativeFormat,
        )
        if settings.contains(AUTOSTART_REG_KEY):
            self.autostart_cb.setChecked(True)
        self.autostart_cb.stateChanged.connect(self.toggle_autostart)
        self.minimize_to_tray_cb = QCheckBox("닫기 버튼을 눌러도 종료하지 않고 트레이로 숨기기")
        self.minimize_to_tray_cb.setChecked(
            self.ui_settings.value("minimize_to_tray", True, type=bool)
        )
        self.minimize_to_tray_cb.stateChanged.connect(self.toggle_minimize_to_tray)
        sys_row.addWidget(self.autostart_cb)
        sys_row.addWidget(self.minimize_to_tray_cb)
        sys_row.addStretch(1)
        config_layout.addLayout(sys_row)

        frame = QWidget()
        frame.setFixedHeight(1)
        frame.setStyleSheet("background-color: #ccc;")
        config_layout.addWidget(frame)

        file_row = QHBoxLayout()
        self.file_label = QLabel("선택된 파일 없음")
        btn_browse = QPushButton("파일 찾아보기")
        btn_browse.clicked.connect(self.browse_file)
        file_row.addWidget(btn_browse)
        file_row.addWidget(self.file_label, 1)
        config_layout.addLayout(file_row)

        macro_row = QHBoxLayout()
        macro_row.addWidget(QLabel("VBA 매크로명:"))
        self.macro_input = QComboBox()
        self.macro_input.setEditable(True)
        self.macro_input.setInsertPolicy(QComboBox.NoInsert)
        self.macro_input.lineEdit().setPlaceholderText(
            "엑셀 파일 선택 시 매크로 목록이 표시됩니다. 직접 입력도 가능합니다."
        )
        self.macro_input.setEnabled(False)
        macro_row.addWidget(self.macro_input, 1)
        config_layout.addLayout(macro_row)

        period_row = QHBoxLayout()
        period_row.addWidget(QLabel("실행 기간:"))
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setMinimumWidth(170)
        period_row.addWidget(self.start_date_edit)
        period_row.addWidget(QLabel(" ~ "))
        self.use_end_date_cb = QCheckBox("종료일 지정")
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat("yyyy-MM-dd")
        self.end_date_edit.setDate(QDate.currentDate().addDays(30))
        self.end_date_edit.setEnabled(False)
        self.end_date_edit.setMinimumWidth(170)
        self.use_end_date_cb.toggled.connect(self.end_date_edit.setEnabled)
        period_row.addWidget(self.use_end_date_cb)
        period_row.addWidget(self.end_date_edit)
        period_row.addStretch(1)
        config_layout.addLayout(period_row)

        type_row = QHBoxLayout()
        type_row.addWidget(QLabel("반복 방식:"))
        self.radio_once = QRadioButton("한 번만")
        self.radio_daily = QRadioButton("매일")
        self.radio_weekly = QRadioButton("매주 (특정 요일)")
        self.radio_monthly = QRadioButton("매월 (특정 날짜)")
        self.radio_weekly.setChecked(True)
        type_row.addWidget(self.radio_once)
        type_row.addWidget(self.radio_daily)
        type_row.addWidget(self.radio_weekly)
        type_row.addWidget(self.radio_monthly)
        type_row.addStretch(1)
        config_layout.addLayout(type_row)

        self.schedule_stack = QStackedWidget()

        self.once_widget = QWidget()
        once_layout = QHBoxLayout(self.once_widget)
        once_layout.setContentsMargins(0, 0, 0, 0)
        once_layout.addWidget(QLabel("시작일과 실행 시간에 1회만 실행합니다."))
        once_layout.addStretch(1)
        self.schedule_stack.addWidget(self.once_widget)

        self.daily_widget = QWidget()
        daily_layout = QHBoxLayout(self.daily_widget)
        daily_layout.setContentsMargins(0, 0, 0, 0)
        daily_layout.addWidget(QLabel("지정한 기간 동안 매일 같은 시간에 실행합니다."))
        daily_layout.addStretch(1)
        self.schedule_stack.addWidget(self.daily_widget)

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
        btn_weekdays = QPushButton("평일")
        btn_weekend = QPushButton("주말")
        btn_all_days = QPushButton("전체")
        btn_weekdays.clicked.connect(lambda: self._select_days(WEEKDAYS))
        btn_weekend.clicked.connect(lambda: self._select_days(WEEKEND))
        btn_all_days.clicked.connect(lambda: self._select_days(list(DAY_KO.keys())))
        day_layout.addWidget(btn_weekdays)
        day_layout.addWidget(btn_weekend)
        day_layout.addWidget(btn_all_days)
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

        self.radio_once.toggled.connect(self._sync_schedule_options)
        self.radio_daily.toggled.connect(self._sync_schedule_options)
        self.radio_weekly.toggled.connect(self._sync_schedule_options)
        self.radio_monthly.toggled.connect(self._sync_schedule_options)
        self.schedule_stack.setCurrentWidget(self.weekly_widget)
        config_layout.addWidget(self.schedule_stack)

        time_row = QHBoxLayout()
        time_row.addWidget(QLabel("실행 시간:"))
        self.time_edit = QTimeEdit()
        self.time_edit.setDisplayFormat("HH:mm")
        self.time_edit.setTime(QTime.currentTime())
        self.time_edit.setMinimumWidth(170)
        btn_add = QPushButton("스케줄 추가")
        btn_add.clicked.connect(self.add_task)
        time_row.addWidget(self.time_edit)
        time_row.addWidget(btn_add, 1)
        config_layout.addLayout(time_row)

        config_box.setLayout(config_layout)
        main_layout.addWidget(config_box)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels(
            ["작업 대상 (파일/매크로)", "실행 기간", "실행 주기", "상태", "마지막 실행", "다음 실행", "관리"]
        )
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.setColumnWidth(6, 200)
        self.table.cellDoubleClicked.connect(self.open_task_folder)
        main_layout.addWidget(self.table)

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

    def _select_days(self, keys) -> None:
        selected = set(keys)
        for eng, cb in self.day_checkboxes.items():
            cb.setChecked(eng in selected)

    def _sync_schedule_options(self) -> None:
        is_once = self.radio_once.isChecked()
        self.use_end_date_cb.setEnabled(not is_once)
        if is_once:
            self.use_end_date_cb.setChecked(False)
        if self.radio_once.isChecked():
            self.schedule_stack.setCurrentWidget(self.once_widget)
        elif self.radio_daily.isChecked():
            self.schedule_stack.setCurrentWidget(self.daily_widget)
        elif self.radio_weekly.isChecked():
            self.schedule_stack.setCurrentWidget(self.weekly_widget)
        else:
            self.schedule_stack.setCurrentWidget(self.monthly_widget)

    def init_tray_icon(self) -> None:
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setToolTip("CMC 스케줄러")

        tray_menu = QMenu()
        show_action = QAction("관리자 화면 열기", self)
        show_action.triggered.connect(self.show_main_window)
        tray_menu.addAction(show_action)
        tray_menu.addSeparator()
        quit_action = QAction("완전히 종료", self)
        quit_action.triggered.connect(self.quit_application)
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)
        self.tray_icon.show()

    def on_tray_icon_activated(self, reason) -> None:
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_main_window()

    def toggle_autostart(self, state) -> None:
        settings = QSettings(
            r"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run",
            QSettings.NativeFormat,
        )
        if self.autostart_cb.isChecked():
            if getattr(sys, "frozen", False):
                cmd_path = f'"{sys.executable}" --hidden'
            else:
                cmd_path = f'"{sys.executable}" "{os.path.abspath(sys.argv[0])}" --hidden'
            settings.setValue(AUTOSTART_REG_KEY, cmd_path)
            self.append_log("[시스템] PC 부팅 시 자동 시작이 설정되었습니다.")
        else:
            settings.remove(AUTOSTART_REG_KEY)
            self.append_log("[시스템] 부팅 시 자동 시작이 해제되었습니다.")

    def toggle_minimize_to_tray(self, state) -> None:
        enabled = self.minimize_to_tray_cb.isChecked()
        self.ui_settings.setValue("minimize_to_tray", enabled)
        if enabled:
            self.append_log("[시스템] 닫기 시 트레이로 숨기기가 설정되었습니다.")
        else:
            self.append_log("[시스템] 닫기 시 프로그램이 종료되도록 변경되었습니다.")

    def closeEvent(self, event) -> None:
        if self.minimize_to_tray_cb.isChecked():
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "백그라운드 실행 중",
                "스케줄러가 시스템 트레이에서 계속 작동합니다.\n완전 종료는 트레이 메뉴에서 하세요.",
                QSystemTrayIcon.Information,
                2500,
            )
            return
        event.accept()
        self.quit_application()

    def quit_application(self) -> None:
        if self.scheduler.running:
            self.scheduler.shutdown(wait=False)
        self.tray_icon.hide()
        QApplication.quit()

    def browse_file(self) -> None:
        start_dir = self.ui_settings.value("last_dir", app_dir())
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "파일 선택",
            start_dir,
            "지원하는 파일 (*.py *.pyw *.exe *.xlsm *.xlsb);;"
            "Python Files (*.py *.pyw);;"
            "실행 파일 (*.exe);;"
            "Excel Macro Files (*.xlsm *.xlsb)",
        )
        if not file_path:
            return

        self.current_selected_file = file_path
        self.file_label.setText(os.path.basename(file_path))
        self.file_label.setToolTip(file_path)
        self.ui_settings.setValue("last_dir", os.path.dirname(file_path))

        ext = os.path.splitext(file_path)[1].lower()
        if ext in (".xlsm", ".xlsb"):
            self._load_excel_macros(file_path)
        else:
            self.macro_input.setEnabled(False)
            self.macro_input.clear()
            self.macro_input.setEditText("")

    def _load_excel_macros(self, file_path: str) -> None:
        self.macro_input.setEnabled(True)
        self.macro_input.clear()
        self.macro_input.setEditText("")
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            macros = list_excel_macros(file_path)
        finally:
            QApplication.restoreOverrideCursor()

        if macros:
            self.macro_input.addItems(macros)
            self.macro_input.setCurrentIndex(0)
            self.append_log(f"[시스템] VBA 매크로 {len(macros)}개를 불러왔습니다.")
        else:
            self.append_log(
                "[경고] 매크로 목록을 자동으로 읽지 못했습니다. "
                "이름을 직접 입력하거나, Excel 보안 센터에서 "
                "'VBA 프로젝트 개체 모델에 액세스할 수 있음'을 허용해 주세요."
            )
        self.macro_input.setFocus()

    def add_task(self) -> None:
        if not self.current_selected_file:
            QMessageBox.warning(self, "파일 없음", "먼저 실행할 파일을 선택해주세요.")
            return

        macro_name = self.macro_input.currentText().strip()
        if self.current_selected_file.lower().endswith((".xlsm", ".xlsb")) and not macro_name:
            QMessageBox.warning(self, "매크로 필요", "엑셀 파일을 스케줄링하려면 VBA 매크로 이름을 입력해야 합니다.")
            return

        if self.radio_once.isChecked():
            schedule_type = "once"
        elif self.radio_daily.isChecked():
            schedule_type = "daily"
        elif self.radio_weekly.isChecked():
            schedule_type = "weekly"
        else:
            schedule_type = "monthly"

        cron_dow = "*"
        cron_day = "*"
        display_cycle = ""

        if schedule_type == "once":
            display_cycle = "한 번만"
        elif schedule_type == "daily":
            display_cycle = "매일"
        elif schedule_type == "weekly":
            selected_days = [eng for eng, cb in self.day_checkboxes.items() if cb.isChecked()]
            if not selected_days:
                QMessageBox.warning(self, "요일 필요", "매주 실행의 경우 최소 하나 이상의 요일을 선택해주세요.")
                return
            cron_dow = ",".join(selected_days)
            display_cycle = weekly_cycle_label(cron_dow)
        else:
            cron_day = self.month_day_combo.currentData()
            display_cycle = f"매월({self.month_day_combo.currentText()})"

        time_obj = self.time_edit.time()
        hour = time_obj.hour()
        minute = time_obj.minute()
        start_date_str = self.start_date_edit.date().toString("yyyy-MM-dd")
        end_date_str = self.end_date_edit.date().toString("yyyy-MM-dd") if self.use_end_date_cb.isChecked() else None

        if schedule_type == "once":
            run_dt = datetime.strptime(f"{start_date_str} {hour:02d}:{minute:02d}", "%Y-%m-%d %H:%M")
            if run_dt <= datetime.now():
                QMessageBox.warning(self, "시간 오류", "한 번 실행 시각이 현재보다 과거입니다. 날짜나 시간을 바꿔 주세요.")
                return

        self.register_task(
            self.current_selected_file,
            macro_name,
            start_date_str,
            end_date_str,
            schedule_type,
            cron_dow,
            cron_day,
            hour,
            minute,
            display_cycle,
            save=True,
        )

    def _build_trigger(self, schedule_type: str, cron_dow: str, cron_day: str,
                       hour: int, minute: int, start_date: str, end_date: str):
        tz_kwargs = {"timezone": self.tz} if self.tz else {}
        if schedule_type == "once":
            run_dt = datetime.strptime(f"{start_date} {hour:02d}:{minute:02d}", "%Y-%m-%d %H:%M")
            if self.tz is not None:
                try:
                    run_dt = run_dt.replace(tzinfo=self.tz)
                except Exception:
                    pass
            return DateTrigger(run_date=run_dt, **tz_kwargs)
        if schedule_type == "daily":
            return CronTrigger(hour=hour, minute=minute, start_date=start_date, end_date=end_date, **tz_kwargs)
        if schedule_type == "weekly":
            return CronTrigger(
                day_of_week=cron_dow, hour=hour, minute=minute,
                start_date=start_date, end_date=end_date, **tz_kwargs,
            )
        return CronTrigger(
            day=cron_day, hour=hour, minute=minute,
            start_date=start_date, end_date=end_date, **tz_kwargs,
        )

    def register_task(self, file_path: str, macro_name: str, start_date: str, end_date: str,
                      schedule_type: str, cron_dow: str, cron_day: str, hour: int, minute: int,
                      display_cycle: str, save: bool = True, enabled: bool = True,
                      job_id: str = None) -> None:
        job_id = job_id or str(uuid.uuid4())

        try:
            trigger = self._build_trigger(
                schedule_type, cron_dow, cron_day, hour, minute, start_date, end_date
            )
            job = self.scheduler.add_job(
                self.execute_script, trigger=trigger, id=job_id,
                args=[file_path, macro_name, job_id],
            )
            if not enabled:
                job.pause()
        except ValueError as e:
            self.append_log(f"[설정 오류] 스케줄을 등록할 수 없습니다: {str(e)}")
            if save:
                QMessageBox.warning(self, "설정 오류", f"스케줄을 등록할 수 없습니다.\n{e}")
            return

        row_index = self.table.rowCount()
        self.table.insertRow(row_index)

        display_name = os.path.basename(file_path)
        if macro_name:
            display_name += f" [{macro_name}]"

        file_item = QTableWidgetItem(display_name)
        file_item.setData(Qt.UserRole, job_id)
        file_item.setToolTip(f"{file_path}\n매크로: {macro_name or '-'}\n더블클릭하면 폴더가 열립니다.")
        self.table.setItem(row_index, 0, file_item)

        period_str = f"{start_date} ~ {end_date if end_date else '제한 없음'}"
        if schedule_type == "once":
            period_str = start_date
        self.table.setItem(row_index, 1, QTableWidgetItem(period_str))
        self.table.setItem(row_index, 2, QTableWidgetItem(f"{display_cycle} {hour:02d}:{minute:02d}"))

        initial_status = "대기 중" if enabled else "일시중지"
        self.table.setItem(row_index, 3, status_item(initial_status))
        self.table.setItem(row_index, 4, QTableWidgetItem("-"))

        next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job.next_run_time else "기간 만료"
        self.table.setItem(row_index, 5, QTableWidgetItem(next_run))
        self.table.setCellWidget(row_index, 6, self._make_row_buttons(job_id, enabled))

        self.tasks.append({
            "file_path": file_path,
            "macro_name": macro_name,
            "start_date": start_date,
            "end_date": end_date,
            "schedule_type": schedule_type,
            "cron_dow": cron_dow,
            "cron_day": cron_day,
            "hour": hour,
            "minute": minute,
            "display_cycle": display_cycle,
            "enabled": enabled,
            "job_id": job_id,
        })

        if save:
            self.save_config()
            self.append_log(f"[등록 완료] {display_name} 스케줄이 지정되었습니다.")
            self.refresh_status_bar()

    def _make_row_buttons(self, job_id: str, enabled: bool) -> QWidget:
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(2, 2, 2, 2)
        layout.setSpacing(3)

        btn_run = QPushButton("실행")
        btn_pause = QPushButton("중지" if enabled else "재개")
        btn_delete = QPushButton("삭제")

        for btn in (btn_run, btn_pause, btn_delete):
            btn.setProperty("job_id", job_id)

        btn_run.clicked.connect(self.run_task_now)
        btn_pause.clicked.connect(self.toggle_pause_task)
        btn_delete.clicked.connect(self.delete_task)

        layout.addWidget(btn_run)
        layout.addWidget(btn_pause)
        layout.addWidget(btn_delete)
        return widget

    def _task_by_id(self, job_id: str):
        for task in self.tasks:
            if task["job_id"] == job_id:
                return task
        return None

    def _row_by_job_id(self, job_id: str):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.UserRole) == job_id:
                return row
        return None

    def run_task_now(self) -> None:
        button = self.sender()
        if not button:
            return
        job_id = button.property("job_id")
        task = self._task_by_id(job_id)
        if not task:
            return
        self.append_log(f"[시스템] {os.path.basename(task['file_path'])} 작업을 즉시 실행합니다.")
        threading.Thread(
            target=self.execute_script,
            args=(task["file_path"], task.get("macro_name", ""), job_id),
            daemon=True,
        ).start()

    def toggle_pause_task(self) -> None:
        button = self.sender()
        if not button:
            return
        job_id = button.property("job_id")
        task = self._task_by_id(job_id)
        job = self.scheduler.get_job(job_id)
        if not task or not job:
            return

        name = os.path.basename(task["file_path"])
        if task.get("enabled", True):
            job.pause()
            task["enabled"] = False
            button.setText("재개")
            self.update_table_status(job_id, "일시중지", "")
            self.append_log(f"[시스템] {name} 작업을 일시중지했습니다.")
        else:
            job.resume()
            task["enabled"] = True
            button.setText("중지")
            next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job.next_run_time else "기간 만료"
            self.update_table_status(job_id, "대기 중", next_run)
            self.append_log(f"[시스템] {name} 작업을 재개했습니다.")

        self.save_config()
        self.refresh_status_bar()

    def delete_task(self) -> None:
        button = self.sender()
        if not button:
            return
        job_id = button.property("job_id")
        task = self._task_by_id(job_id)
        file_name = os.path.basename(task["file_path"]) if task else "이 작업"

        reply = QMessageBox.question(
            self,
            "작업 삭제",
            f"'{file_name}' 스케줄을 삭제할까요?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        try:
            if self.scheduler.get_job(job_id):
                self.scheduler.remove_job(job_id)
        except Exception:
            pass

        row = self._row_by_job_id(job_id)
        if row is not None:
            self.table.removeRow(row)

        self.tasks = [t for t in self.tasks if t["job_id"] != job_id]
        self.save_config()
        self.append_log(f"[삭제 완료] {file_name} 스케줄이 삭제되었습니다.")
        self.refresh_status_bar()

    def open_task_folder(self, row: int, _col: int) -> None:
        item = self.table.item(row, 0)
        if not item:
            return
        task = self._task_by_id(item.data(Qt.UserRole))
        if not task:
            return
        folder = os.path.dirname(task["file_path"])
        if os.path.isdir(folder):
            os.startfile(folder)
        else:
            QMessageBox.warning(self, "경로 없음", f"폴더를 찾을 수 없습니다.\n{folder}")

    def open_log_folder(self) -> None:
        log_dir = os.path.abspath("logs")
        os.makedirs(log_dir, exist_ok=True)
        os.startfile(log_dir)

    def clear_log(self) -> None:
        self.log_view.clear()

    def _run_excel_macro(self, file_path: str, macro_name: str) -> None:
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
            self.signals.log_signal.emit(
                f"[{os.path.basename(file_path)}] VBA 매크로({macro_name}) 정상 종료 및 저장 완료"
            )
        finally:
            if wb:
                wb.Close(SaveChanges=False)
            if excel:
                excel.Quit()
            pythoncom.CoUninitialize()

    def _run_python(self, file_path: str) -> None:
        cmd = find_python_executable()
        workdir = os.path.dirname(os.path.abspath(file_path)) or None
        flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        result = subprocess.run(
            cmd + [file_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=True,
            cwd=workdir,
            creationflags=flags,
        )
        if result.stdout and result.stdout.strip():
            self.signals.log_signal.emit(
                f"[{os.path.basename(file_path)} 출력]:\n{result.stdout.strip()}"
            )
        if result.stderr and result.stderr.strip():
            self.signals.log_signal.emit(
                f"[{os.path.basename(file_path)} 경고]:\n{result.stderr.strip()}"
            )

    def _run_executable(self, file_path: str) -> None:
        workdir = os.path.dirname(os.path.abspath(file_path)) or None
        result = subprocess.run(
            [file_path],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            check=True,
            cwd=workdir,
        )
        if result.stdout and result.stdout.strip():
            self.signals.log_signal.emit(
                f"[{os.path.basename(file_path)} 출력]:\n{result.stdout.strip()}"
            )

    def execute_script(self, file_path: str, macro_name: str, job_id: str) -> None:
        file_name = os.path.basename(file_path)
        started = datetime.now()
        start_time = started.strftime("%Y-%m-%d %H:%M:%S")
        self.signals.log_signal.emit(f"[{start_time}] [시작] {file_name} 작업을 실행합니다.")
        self.signals.status_signal.emit(job_id, "실행 중", start_time)

        try:
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"파일이 존재하지 않습니다: {file_path}")

            ext = os.path.splitext(file_path)[1].lower()
            if ext in (".xlsm", ".xlsb"):
                self._run_excel_macro(file_path, macro_name)
            elif ext in (".py", ".pyw"):
                self._run_python(file_path)
            else:
                self._run_executable(file_path)

            elapsed = datetime.now() - started
            end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.signals.log_signal.emit(
                f"[{end_time}] [성공] {file_name} 작업 완료 (소요 {elapsed.total_seconds():.1f}초)"
            )
            job = self.scheduler.get_job(job_id)
            next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job and job.next_run_time else "기간 만료"
            self.signals.status_signal.emit(job_id, "정상 종료", next_run)

        except subprocess.CalledProcessError as e:
            detail = (e.stderr or e.stdout or str(e)).strip() or str(e)
            end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.signals.log_signal.emit(f"[{end_time}] [에러 발생] {file_name} 실패: {detail}")
            job = self.scheduler.get_job(job_id)
            next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job and job.next_run_time else "기간 만료"
            self.signals.status_signal.emit(job_id, "오류 발생", next_run)
        except Exception as e:
            end_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.signals.log_signal.emit(f"[{end_time}] [에러 발생] {file_name} 실패: {str(e)}")
            job = self.scheduler.get_job(job_id)
            next_run = job.next_run_time.strftime("%Y-%m-%d %H:%M:%S") if job and job.next_run_time else "기간 만료"
            self.signals.status_signal.emit(job_id, "오류 발생", next_run)

    def append_log(self, text: str) -> None:
        append_log_file(text)
        self.log_view.append(text)

    def update_table_status(self, job_id: str, status: str, time_info: str) -> None:
        row = self._row_by_job_id(job_id)
        if row is None:
            return
        self.table.setItem(row, 3, status_item(status))
        if status == "실행 중":
            self.table.setItem(row, 4, QTableWidgetItem(time_info))
        elif time_info:
            self.table.setItem(row, 5, QTableWidgetItem(time_info))
        self.refresh_status_bar()

    def refresh_status_bar(self) -> None:
        jobs = self.scheduler.get_jobs()
        upcoming = [j for j in jobs if j.next_run_time]
        upcoming.sort(key=lambda j: j.next_run_time)
        paused = sum(1 for t in self.tasks if not t.get("enabled", True))
        prefix = f"등록 {len(self.tasks)}건"
        if paused:
            prefix += f" · 중지 {paused}건"
        if upcoming:
            job = upcoming[0]
            name = os.path.basename(job.args[0]) if job.args else ""
            self.statusBar().showMessage(
                f"{prefix}  |  다음 실행: {job.next_run_time.strftime('%Y-%m-%d %H:%M')}  ({name})"
            )
        else:
            self.statusBar().showMessage(f"{prefix}  |  예정된 작업 없음")

    def save_config(self) -> None:
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.tasks, f, ensure_ascii=False, indent=4)
        except Exception:
            self.append_log("[경고] 설정 파일을 저장하지 못했습니다.")

    def load_config(self) -> None:
        if not os.path.exists(CONFIG_FILE):
            return
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
                display_cycle = item.get("display_cycle")
                if not display_cycle:
                    if schedule_type == "once":
                        display_cycle = "한 번만"
                    elif schedule_type == "daily":
                        display_cycle = "매일"
                    elif schedule_type == "weekly":
                        display_cycle = weekly_cycle_label(cron_dow)
                    else:
                        display_cycle = (
                            "매월(매월 말일(마지막 날))"
                            if cron_day == "last"
                            else f"매월(매월 {cron_day}일)"
                        )
                self.register_task(
                    item["file_path"],
                    macro_name,
                    start_date,
                    end_date,
                    schedule_type,
                    cron_dow,
                    cron_day,
                    item["hour"],
                    item["minute"],
                    display_cycle,
                    save=False,
                    enabled=item.get("enabled", True),
                    job_id=item.get("job_id"),
                )
            self.append_log(f"[시스템] 등록된 {len(config_data)}개의 스케줄 설정을 로드했습니다.")
        except Exception as e:
            self.append_log(f"[경고] 설정 파일을 읽지 못했습니다: {e}")


def activate_existing_instance() -> bool:
    socket = QLocalSocket()
    socket.connectToServer(INSTANCE_KEY)
    if socket.waitForConnected(300):
        socket.write(b"SHOW")
        socket.waitForBytesWritten(300)
        socket.flush()
        return True
    return False


def main() -> None:
    ensure_workdir()

    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)

    if activate_existing_instance():
        sys.exit(0)

    lock_path = os.path.join(os.environ.get("TEMP", app_dir()), "cmc_scheduler.lock")
    lock_file = QLockFile(lock_path)
    if not lock_file.tryLock(200):
        activate_existing_instance()
        sys.exit(0)

    window = MultiSchedulerApp()
    window._instance_lock = lock_file

    if "--hidden" not in sys.argv:
        window.show()

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
