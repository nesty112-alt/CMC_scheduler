import sqlite3
import pandas as pd
import os
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import warnings

# Pandas의 띄어쓰기 관련 경고 메시지를 숨김 처리 (터미널 출력 방지)
warnings.filterwarnings('ignore', category=UserWarning, module='pandas')

# 설정 파일 경로 (마지막 실행 경로 저장용)
CONFIG_FILE = "config.json"


class ValidationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("데이터 검증 프로그램")
        self.root.geometry("900x650")

        # 기본 설정값
        self.config = {
            "main_csv": "",
            "main_rule": "",
            "sub_csv": "",
            "sub_rule": "",
            "output_dir": ""
        }
        self.load_config()

        # UI 구성
        self.create_widgets()

    def load_config(self):
        """설정 파일(JSON)을 불러와서 기존 경로를 복구합니다."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    saved_config = json.load(f)
                    self.config.update(saved_config)
            except Exception:
                pass

    def save_config(self):
        """현재 UI에 입력된 경로들을 설정 파일에 저장합니다."""
        self.config["main_csv"] = self.main_csv_var.get()
        self.config["main_rule"] = self.main_rule_var.get()
        self.config["sub_csv"] = self.sub_csv_var.get()
        self.config["sub_rule"] = self.sub_rule_var.get()
        self.config["output_dir"] = self.output_dir_var.get()

        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        """UI 화면의 각 요소들을 배치합니다."""
        # --- [1] 파일 경로 설정 영역 ---
        frame_paths = tk.LabelFrame(self.root, text="파일 및 경로 설정", padx=10, pady=10)
        frame_paths.pack(fill="x", padx=10, pady=5)

        self.main_csv_var = tk.StringVar(value=self.config["main_csv"])
        self.main_rule_var = tk.StringVar(value=self.config["main_rule"])
        self.sub_csv_var = tk.StringVar(value=self.config["sub_csv"])
        self.sub_rule_var = tk.StringVar(value=self.config["sub_rule"])
        self.output_dir_var = tk.StringVar(value=self.config["output_dir"])

        # 파일 경로가 변경될 때마다 실행 옵션을 자동으로 선택하도록 이벤트 연결
        for var in (self.main_csv_var, self.main_rule_var, self.sub_csv_var, self.sub_rule_var):
            var.trace_add("write", self.auto_select_run_option)

        def add_file_row(parent, label_text, var, row, is_dir=False):
            tk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", pady=2)
            tk.Entry(parent, textvariable=var, width=65).grid(row=row, column=1, padx=5, pady=2)

            def browse():
                if is_dir:
                    path = filedialog.askdirectory()
                else:
                    path = filedialog.askopenfilename(
                        filetypes=[("Excel/CSV Files", "*.xlsx *.csv"), ("All Files", "*.*")])
                if path:
                    var.set(path)

            tk.Button(parent, text="찾아보기", command=browse).grid(row=row, column=2, pady=2)

        add_file_row(frame_paths, "Main 데이터 (CSV):", self.main_csv_var, 0)
        add_file_row(frame_paths, "Main 검증 룰 (Excel):", self.main_rule_var, 1)
        tk.Label(frame_paths, text="-" * 110).grid(row=2, column=0, columnspan=3)
        add_file_row(frame_paths, "Sub 데이터 (CSV):", self.sub_csv_var, 3)
        add_file_row(frame_paths, "Sub 검증 룰 (Excel):", self.sub_rule_var, 4)
        tk.Label(frame_paths, text="-" * 110).grid(row=5, column=0, columnspan=3)
        add_file_row(frame_paths, "결과 저장 폴더:", self.output_dir_var, 6, is_dir=True)

        # --- [2] 실행 옵션 영역 ---
        frame_options = tk.Frame(self.root)
        frame_options.pack(fill="x", padx=10, pady=5)

        self.run_option_var = tk.IntVar(value=3)
        tk.Radiobutton(frame_options, text="Main만 실행", variable=self.run_option_var, value=1).pack(side="left", padx=10)
        tk.Radiobutton(frame_options, text="Sub만 실행", variable=self.run_option_var, value=2).pack(side="left", padx=10)
        tk.Radiobutton(frame_options, text="모두 실행 (Main + Sub)", variable=self.run_option_var, value=3).pack(
            side="left", padx=10)

        self.run_btn_text = tk.StringVar(value="▶ 검증 실행")
        self.run_btn = tk.Button(frame_options, textvariable=self.run_btn_text, bg="#4CAF50", fg="white",
                                 font=("Arial", 11, "bold"), command=self.run_process)
        self.run_btn.pack(side="right", padx=10)

        # 초기 프로그램 로드 시 입력된 경로를 바탕으로 라디오 버튼 상태 업데이트
        self.auto_select_run_option()

        # --- [3] 결과 출력 영역 (Treeview) ---
        frame_results = tk.LabelFrame(self.root, text="실행 결과 대시보드 (오류 내역 우클릭 시 복사 가능)", padx=10, pady=10)
        frame_results.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("target", "rule_name", "status", "total_cnt", "error_cnt", "error_msg")
        self.tree = ttk.Treeview(frame_results, columns=columns, show="headings")

        self.tree.heading("target", text="타겟")
        self.tree.heading("rule_name", text="검증 룰")
        self.tree.heading("status", text="결과")
        self.tree.heading("total_cnt", text="총 데이터")
        self.tree.heading("error_cnt", text="오류 건수")
        self.tree.heading("error_msg", text="비고 (SQL 오류 등)")

        self.tree.column("target", width=60, anchor="center")
        self.tree.column("rule_name", width=220, anchor="w")
        self.tree.column("status", width=80, anchor="center")
        self.tree.column("total_cnt", width=90, anchor="e")
        self.tree.column("error_cnt", width=80, anchor="e")
        self.tree.column("error_msg", width=300, anchor="w")

        self.tree.pack(fill="both", expand=True)

        # 더블클릭 및 우클릭 이벤트 바인딩
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Button-2>", self.show_context_menu)

    def auto_select_run_option(self, *args):
        """입력된 파일 경로 유무를 감지하여 실행 옵션 라디오 버튼을 자동 선택합니다."""
        main_ready = bool(self.main_csv_var.get().strip() and self.main_rule_var.get().strip())
        sub_ready = bool(self.sub_csv_var.get().strip() and self.sub_rule_var.get().strip())

        if main_ready and sub_ready:
            self.run_option_var.set(3)  # 모두 실행
        elif main_ready and not sub_ready:
            self.run_option_var.set(1)  # Main만 실행
        elif sub_ready and not main_ready:
            self.run_option_var.set(2)  # Sub만 실행

    def show_context_menu(self, event):
        """Treeview에서 우클릭 시 복사 메뉴를 띄웁니다."""
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)

        if row_id:
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)

            item_values = self.tree.item(row_id, "values")

            if column_id == '#6':
                error_msg = item_values[5]

                if error_msg and error_msg != "-":
                    context_menu = tk.Menu(self.root, tearoff=0)
                    context_menu.add_command(
                        label="비고(오류 메시지) 복사",
                        command=lambda: self.copy_to_clipboard_direct(error_msg)
                    )
                    context_menu.post(event.x_root, event.y_root)

    def copy_to_clipboard_direct(self, text):
        """전달받은 텍스트를 클립보드에 팝업 없이 즉시 복사합니다."""
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def on_tree_double_click(self, event):
        """Treeview 항목 더블클릭 시 상세 메시지 팝업을 띄웁니다."""
        selected_item = self.tree.selection()
        if not selected_item:
            return

        item_values = self.tree.item(selected_item[0], "values")
        error_msg = item_values[5]

        if error_msg and error_msg != "-":
            self.show_error_popup(item_values[1], error_msg)

    def show_error_popup(self, rule_name, message):
        """상세 오류 내용을 보여주고 복사할 수 있는 팝업창을 생성합니다."""
        popup = tk.Toplevel(self.root)
        popup.title("상세 오류 메시지 확인")
        popup.geometry("600x300")
        popup.attributes('-topmost', True)

        tk.Label(popup, text=f"검증 룰: {rule_name}", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))

        text_widget = tk.Text(popup, wrap="word", font=("Consolas", 10))
        text_widget.insert("1.0", message)
        text_widget.config(state="disabled")
        text_widget.pack(fill="both", expand=True, padx=10, pady=10)

        def copy_to_clipboard():
            self.root.clipboard_clear()
            self.root.clipboard_append(message)
            btn_copy.config(text="✓ 복사 완료", bg="#4CAF50")

        btn_copy = tk.Button(popup, text="메시지 복사하기", bg="#2196F3", fg="white", command=copy_to_clipboard)
        btn_copy.pack(pady=(0, 10))

    def run_process(self):
        """실행 버튼 클릭 시 유효성 검사 후 작동하는 스레드 트리거 함수입니다."""
        opt = self.run_option_var.get()
        out_dir = self.output_dir_var.get().strip()

        # 1. 저장 폴더 입력 확인
        if not out_dir:
            messagebox.showerror("오류", "결과를 저장할 폴더를 지정해주세요.")
            return

        # 2. 실행 옵션에 따른 필수 파일 경로 입력 확인
        main_c = self.main_csv_var.get().strip()
        main_r = self.main_rule_var.get().strip()
        sub_c = self.sub_csv_var.get().strip()
        sub_r = self.sub_rule_var.get().strip()

        if opt == 1:  # Main만 실행
            if not main_c or not main_r:
                messagebox.showwarning("경로 누락", "Main 실행을 위해 'Main 데이터'와 '검증 룰' 파일 경로를 모두 입력해주세요.")
                return
        elif opt == 2:  # Sub만 실행
            if not sub_c or not sub_r:
                messagebox.showwarning("경로 누락", "Sub 실행을 위해 'Sub 데이터'와 '검증 룰' 파일 경로를 모두 입력해주세요.")
                return
        elif opt == 3:  # 모두 실행
            if not main_c or not main_r or not sub_c or not sub_r:
                messagebox.showwarning("경로 누락",
                                       "모두 실행하려면 Main 및 Sub 데이터와 검증 룰 등 4개의 파일 경로를 전부 입력해야 합니다.\n(일부만 입력하셨다면 해당 옵션으로 변경해 주세요.)")
                return

        # 3. 유효성 검사 통과 시 실제 실행 준비
        self.save_config()

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.run_btn.config(state="disabled")
        self.run_btn_text.set("⏳ 실행 중...")

        thread = threading.Thread(target=self._run_process_thread, args=(opt, out_dir))
        thread.daemon = True
        thread.start()

    def _safe_tree_insert(self, values):
        """스레드에서 안전하게 UI(Treeview)를 업데이트하기 위한 래퍼 함수"""
        self.tree.insert("", "end", values=values)

    def _run_process_thread(self, opt, out_dir):
        """실제 데이터 처리를 담당하는 백그라운드 스레드 로직입니다."""
        conn = sqlite3.connect(':memory:')

        try:
            if opt in (1, 3):
                self._load_and_validate("Main", self.main_csv_var.get(), self.main_rule_var.get(), out_dir, conn)

            if opt in (2, 3):
                self._load_and_validate("Sub", self.sub_csv_var.get(), self.sub_rule_var.get(), out_dir, conn)

        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("실행 오류", f"작업 중 오류가 발생했습니다:\n{str(e)}"))
        finally:
            conn.close()
            # 실행 완료 후 UI 버튼 텍스트가 원래대로 돌아오는 것으로 작업 완료 확인 가능
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            self.root.after(0, lambda: self.run_btn_text.set("▶ 검증 실행"))

    def _load_and_validate(self, target_name, csv_path, rule_path, out_dir, conn):
        """데이터를 로드하고 룰에 따라 검증하는 내부 함수입니다."""
        if not csv_path or not rule_path:
            return

        table_name = target_name.lower()

        try:
            # 1차 시도: utf-8 인코딩으로 읽기
            try:
                df_data = pd.read_csv(csv_path, encoding='utf-8', low_memory=False)
            # 2차 시도: exe 환경에서 발생하는 다양한 인코딩 관련 예외를 모두 포괄하여 처리
            except Exception:
                df_data = pd.read_csv(csv_path, encoding='cp949', low_memory=False)

            df_data.to_sql(table_name, conn, if_exists='replace', index=False)
            total_records = len(df_data)
        except Exception as e:
            raise Exception(f"{target_name} 데이터 로드 실패: {e}")

        try:
            df_rules = pd.read_excel(rule_path)
        except Exception as e:
            raise Exception(f"{target_name} 룰 파일 로드 실패: {e}")

        all_errors = []
        output_excel_path = os.path.join(out_dir, f"{table_name}_audit_result.xlsx")

        for index, row in df_rules.iterrows():
            rule_id = str(row.get('number', ''))
            rule_name = str(row.get('Rule_Name', ''))
            sql_query = row.get('SQL_Query', '')

            if not sql_query or pd.isna(sql_query):
                continue

            display_name = f"[{rule_id}] {rule_name}" if rule_id else rule_name

            try:
                result_df = pd.read_sql_query(sql_query, conn)
                error_count = len(result_df)
                status = "🚨 오류" if error_count > 0 else "✅ 정상"

                self.root.after(0, self._safe_tree_insert,
                                (target_name, display_name, status, f"{total_records:,}건", f"{error_count:,}건", "-"))

                if not result_df.empty:
                    temp_df = result_df.copy()
                    temp_df.insert(0, '오류유형', display_name)
                    all_errors.append(temp_df)

            except Exception as e:
                error_msg = str(e).replace('\n', ' ')
                self.root.after(0, self._safe_tree_insert,
                                (target_name, display_name, "⚠️ SQL실패", f"{total_records:,}건", "N/A", error_msg))

                # [수정됨] exe(noconsole) 환경에서 튕김 현상을 유발할 수 있는 print문 삭제

        if all_errors:
            final_error_df = pd.concat(all_errors, ignore_index=True)
            with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
                final_error_df.to_excel(writer, sheet_name='오류내역조회', index=False)
        else:
            empty_df = pd.DataFrame(columns=['오류유형', '결과메시지'])
            empty_df.loc[0] = ['정상', '발견된 오류 데이터가 없습니다.']
            with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
                empty_df.to_excel(writer, sheet_name='오류내역_없음', index=False)


if __name__ == "__main__":
    root = tk.Tk()
    app = ValidationApp(root)
    root.mainloop()
