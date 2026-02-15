import sys
import threading
import queue
import logging
from dataclasses import dataclass, field
from pathlib import Path  # os.path 대신 사용 (Modern Python)
from typing import List, Optional

import pandas as pd  # type: ignore
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
from openpyxl import load_workbook  # type: ignore
from openpyxl.chart import ScatterChart, Reference  # type: ignore
import plotly.express as px  # type: ignore

# -------------------------
# 1. 설정 (Configuration)
# -------------------------
@dataclass(frozen=True)
class AppConfig:
    """프로그램의 모든 설정을 한곳에서 관리"""
    log_file: Path = Path("merge_excel_log.txt")
    target_column: str = "Line"       # 기준이 되는 컬럼
    date_column: str = "Date"         # 날짜 컬럼
    base_filename: str = "merged_file"

    # 로깅 설정
    def setup_logging(self):
        logging.basicConfig(
            filename=self.log_file,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )
        logging.info("System Initialized.")

# -------------------------
# 2. 데이터 구조 (Data Structures)
# -------------------------
@dataclass
class MergeResult:
    merged_path: Path
    sorted_path: Optional[Path]
    excluded_sheets: List[str]
    report_msg: str

@dataclass
class ProgressEvent:
    percent: int
    message: str

# -------------------------
# 3. 비즈니스 로직 (Business Logic)
# -------------------------
class ExcelProcessor:
    """엑셀 처리 및 차트 생성을 전담하는 전문가 클래스"""
    def __init__(self, config: AppConfig):
        self.cfg = config
        self.logger = logging.getLogger(self.__class__.__name__)

    def process(self, source_files: List[Path], dest_folder: Path, sort_key: str, progress_q: queue.Queue) -> MergeResult:
        def report(pct, msg):
            progress_q.put(ProgressEvent(pct, msg))

        # 1. 병합
        report(10, "파일 읽기 및 병합 중...")
        df_merged, excluded = self._merge_files(source_files)

        # 2. 병합 파일 저장
        report(50, "병합 파일 저장 중...")
        merged_path = dest_folder / f"{self.cfg.base_filename}_1_merged.xlsx"
        df_merged.to_excel(merged_path, index=False)

        # 3. 정렬 및 저장
        sorted_path = None
        if sort_key in df_merged.columns:
            report(70, f"'{sort_key}' 기준으로 정렬 중...")
            df_sorted = self._sort_dataframe(df_merged, sort_key)
            sorted_path = dest_folder / f"{self.cfg.base_filename}_2_sorted.xlsx"
            df_sorted.to_excel(sorted_path, index=False)

            # 4. 엑셀 차트 생성 (OpenPyXL)
            report(85, "엑셀 내부 차트 생성 중...")
            try:
                self._create_excel_chart(sorted_path)
            except Exception as e:
                self.logger.error(f"Excel Chart Error: {e}")

            # 5. 웹 그래프 생성 (Plotly)
            report(95, "웹(HTML) 그래프 생성 중...")
            try:
                self._create_plotly_html(df_sorted, sort_key, dest_folder)
            except Exception as e:
                self.logger.error(f"Plotly Error: {e}")

        report(100, "완료")

        msg = f"병합 완료!\n- 저장 위치: {dest_folder}\n- 제외된 시트: {len(excluded)}개"
        return MergeResult(merged_path, sorted_path, excluded, msg)

    def _merge_files(self, files: List[Path]):
        frames = []
        excluded = []
        for f in files:
            try:
                xls = pd.ExcelFile(f)
                for sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    # 설정된 타겟 컬럼(Line)이 있고, 숫자형인지 확인
                    if self.cfg.target_column in df.columns and pd.api.types.is_numeric_dtype(df[self.cfg.target_column]):
                        # 날짜 컬럼 전처리
                        if self.cfg.date_column in df.columns:
                            df[self.cfg.date_column] = pd.to_datetime(df[self.cfg.date_column], errors='coerce').ffill().bfill()

                        df = df.set_index(self.cfg.target_column)
                        frames.append(df)
                    else:
                        excluded.append(f"{f.name} - {sheet}")
            except Exception as e:
                self.logger.error(f"Error reading {f}: {e}")

        if not frames:
            raise ValueError("합칠 데이터가 없습니다.")

        # 병합 후 인덱스 정렬 및 중복 제거
        result = pd.concat(frames, axis=0)
        result = result.sort_index().groupby(level=0).last().reset_index()
        return result, excluded

    def _sort_dataframe(self, df, key):
        df = df.sort_values(by=key)
        # 정렬 기준 컬럼 결측치 채우기
        df[key] = df[key].ffill().bfill()

        # 컬럼 순서 변경 (기준 컬럼을 맨 앞으로)
        cols = df.columns.tolist()
        if key in cols:
            cols.insert(0, cols.pop(cols.index(key)))
        return df[cols]

    def _create_excel_chart(self, path: Path):
        """OpenPyXL을 이용해 엑셀 파일 내부에 차트 삽입"""
        wb = load_workbook(path)
        ws = wb.active

        if ws.max_row <= 1: return

        chart = ScatterChart()
        chart.title = "Scatter Plot"
        chart.style = 13
        chart.x_axis.title = "X"
        chart.y_axis.title = "Values"

        # 데이터 범위 설정 (1열은 X축, 나머지 열은 Y축)
        x_values = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
        for col in range(2, ws.max_column + 1):
            y_values = Reference(ws, min_col=col, min_row=1, max_row=ws.max_row)
            chart.add_data(y_values, titles_from_data=True)

        chart.set_categories(x_values)

        # 차트 시트 생성
        chart_sheet_name = "Chart"
        if chart_sheet_name in wb.sheetnames:
            del wb[chart_sheet_name]

        cs = wb.create_sheet(chart_sheet_name)
        cs.add_chart(chart, "A1")
        wb.save(path)

    def _create_plotly_html(self, df: pd.DataFrame, x_axis: str, folder: Path):
        """Plotly를 이용해 HTML 그래프 파일 생성"""
        numeric_cols = df.select_dtypes(include="number").columns.tolist()
        if x_axis in numeric_cols:
            numeric_cols.remove(x_axis)

        fig = px.scatter()

        # 숫자 데이터 추가 (왼쪽 Y축)
        for col in numeric_cols:
            fig.add_scatter(x=df[x_axis], y=df[col], mode='lines', name=col, yaxis='y1')

        # 날짜 데이터 추가 (오른쪽 Y축, Date가 있을 경우)
        if self.cfg.date_column in df.columns and x_axis != self.cfg.date_column:
            fig.add_scatter(x=df[x_axis], y=df[self.cfg.date_column], mode='lines', name='Date', yaxis='y2')
            fig.update_layout(
                yaxis2=dict(title="Date", overlaying="y", side="right")
            )

        fig.update_layout(title=f"Chart based on {x_axis}", xaxis_title=x_axis)

        # 파일 저장 (pathlib 경로를 string으로 변환해야 함)
        out_path = folder / f"{self.cfg.base_filename}_{x_axis}_chart.html"
        fig.write_html(str(out_path))

        # (선택 사항) 자동으로 브라우저 띄우기 - 서버 환경이면 주석 처리
        # try:
        #     fig.show()
        # except:
        #     passs

# -------------------------
# 4. UI 클래스 (GUI)
# -------------------------
class MergeApp(tk.Tk):
    """화면을 담당하는 사장님 클래스"""
    def __init__(self, processor: ExcelProcessor):
        super().__init__()
        self.processor = processor
        self.title("Smart Excel Merger")
        self.geometry("400x200")

        # UI 컴포넌트 선언
        self.p_bar: ttk.Progressbar
        self.status_lbl: ttk.Label

        # UI 컴포넌트 초기화
        self._init_ui()


        # 통신용 큐
        self.progress_q = queue.Queue()
        self.result_q = queue.Queue()
        self.error_q = queue.Queue()

    def _init_ui(self):
        # 스타일링 및 위젯 배치
        lbl = ttk.Label(self, text="Excel 병합 도구", font=("Helvetica", 16))
        lbl.pack(pady=20)

        btn = ttk.Button(self, text="작업 시작 (파일 선택)", command=self.start_process)
        btn.pack(pady=10, fill='x', padx=50)

        self.p_bar = ttk.Progressbar(self, length=300, mode='determinate')
        self.p_bar.pack(pady=10)

        self.status_lbl = ttk.Label(self, text="대기 중...")
        self.status_lbl.pack()

    def start_process(self):
        # 1. 파일 선택
        files = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx")])
        if not files: return
        source_paths = [Path(f) for f in files]

        # 2. 폴더 선택
        folder = filedialog.askdirectory()
        if not folder: return
        dest_path = Path(folder)

        # 3. 정렬 기준 입력
        sort_key = simpledialog.askstring("설정", "정렬 기준 컬럼명 (예: Date, ID)")
        if not sort_key: return

        # UI 잠금 및 초기화
        self.p_bar['value'] = 0
        self.status_lbl.config(text="작업 시작...")

        # 스레드 시작
        threading.Thread(
            target=self._run_worker,
            args=(source_paths, dest_path, sort_key),
            daemon=True
        ).start()

        # 감시 시작
        self.after(100, self._poll_queue)  # type: ignore

    def _run_worker(self, sources, dest, sort_key):
        try:
            result = self.processor.process(sources, dest, sort_key, self.progress_q)
            self.result_q.put(result)
        except Exception as e:
            self.error_q.put(e)

    def _poll_queue(self):
        # 진행률 업데이트
        try:
            while True:
                evt = self.progress_q.get_nowait()
                self.p_bar['value'] = evt.percent
                self.status_lbl.config(text=evt.message)
        except queue.Empty:
            pass

        # 에러 처리
        if not self.error_q.empty():
            err = self.error_q.get()
            messagebox.showerror("Error", str(err))
            self.status_lbl.config(text="오류 발생")
            return

        # 완료 처리
        if not self.result_q.empty():
            res = self.result_q.get()
            self.p_bar['value'] = 100
            self.status_lbl.config(text="완료!")
            messagebox.showinfo("성공", res.report_msg)
            return

        # 계속 감시
        self.after(100, self._poll_queue)  # type: ignore

# -------------------------
# 5. 실행부 (Entry Point)
# -------------------------
if __name__ == "__main__":
    # 설정 생성
    config = AppConfig()
    config.setup_logging()

    # 로직 객체 생성 (의존성 주입)
    logic_processor = ExcelProcessor(config)

    # 앱 실행
    app = MergeApp(logic_processor)
    app.mainloop()
