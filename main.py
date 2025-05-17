import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import webbrowser
import time
import logging
import sys
import platform
import subprocess
from PIL import Image, ImageTk
from datetime import datetime
import atexit
import tempfile
import shutil
import json # For history
from typing import Optional, List, Dict, Any, Callable


from pptx import Presentation

# 프로젝트 설정 파일 import
import config

# 프로젝트 루트의 다른 .py 파일들 import
from translator import OllamaTranslator
from pptx_handler import PptxHandler
from ocr_handler import PaddleOcrHandler, EasyOcrHandler
from ollama_service import OllamaService
from chart_xml_handler import ChartXmlHandler
import utils

# --- 로깅 설정 ---
debug_mode = "--debug" in sys.argv
log_level = config.DEBUG_LOG_LEVEL if debug_mode else config.DEFAULT_LOG_LEVEL

root_logger = logging.getLogger()
root_logger.setLevel(log_level)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)
if not any(isinstance(h, logging.StreamHandler) for h in root_logger.handlers):
    root_logger.addHandler(console_handler)

# --- 경로 설정 (config.py에서 가져옴) ---
BASE_DIR_MAIN = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = config.ASSETS_DIR
FONTS_DIR = config.FONTS_DIR
LOGS_DIR = config.LOGS_DIR
HISTORY_DIR = config.HISTORY_DIR # 히스토리 경로

# 로그 및 히스토리 폴더 생성 확인
for folder_path in [LOGS_DIR, HISTORY_DIR]:
    if not os.path.exists(folder_path):
        try: os.makedirs(folder_path)
        except Exception as e: print(f"폴더 생성 실패: {folder_path}, 오류: {e}")

logger = logging.getLogger(__name__)

# --- 전역 변수 및 설정 (config.py에서 가져옴) ---
APP_NAME = config.APP_NAME
DEFAULT_MODEL = config.DEFAULT_OLLAMA_MODEL
SUPPORTED_LANGUAGES = config.SUPPORTED_LANGUAGES
# TRANSLATION_HISTORY 는 파일 기반으로 변경


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title(APP_NAME)
        self.general_file_handler = None
        self._setup_logging_file_handler()

        app_icon_png_path = os.path.join(ASSETS_DIR, "app_icon.png")
        app_icon_ico_path = os.path.join(ASSETS_DIR, "app_icon.ico")
        icon_set = False
        try:
            if platform.system() == "Windows" and os.path.exists(app_icon_ico_path):
                self.master.iconbitmap(app_icon_ico_path); icon_set = True
            if not icon_set and os.path.exists(app_icon_png_path):
                try:
                    icon_image_tk = tk.PhotoImage(file=app_icon_png_path, master=self.master)
                    self.master.iconphoto(True, icon_image_tk); icon_set = True
                except tk.TclError:
                    try:
                        pil_icon = Image.open(app_icon_png_path)
                        icon_image_pil = ImageTk.PhotoImage(pil_icon, master=self.master)
                        self.master.iconphoto(True, icon_image_pil); icon_set = True
                    except Exception as e_pil_icon_fallback: logger.warning(f"Pillow로도 PNG 아이콘 설정 실패: {e_pil_icon_fallback}")
            if not icon_set: logger.warning(f"애플리케이션 아이콘 파일을 찾을 수 없거나 설정 실패.")
        except Exception as e_icon_general: logger.warning(f"애플리케이션 아이콘 설정 중 예외: {e_icon_general}", exc_info=True)

        self.style = ttk.Style()
        current_os = platform.system()
        if current_os == "Windows": self.style.theme_use('vista')
        elif current_os == "Darwin": self.style.theme_use('aqua')
        else: self.style.theme_use('clam')

        self.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.ollama_service = OllamaService()
        self.translator = OllamaTranslator()
        self.pptx_handler = PptxHandler()
        self.chart_xml_handler = ChartXmlHandler(self.translator, self.ollama_service)
        
        self.ocr_handler = None
        self.current_ocr_engine_type = None

        self.translation_thread = None
        self.model_download_thread = None
        self.stop_event = threading.Event()
        self.logo_image_tk_bottom = None
        self.start_time = None

        # 파일 정보 관련 변수 (작업량 표시는 삭제되므로 관련 변수도 일부 삭제)
        self.current_file_slide_count = 0
        # self.current_file_text_elements_count = 0 # 삭제
        # self.current_file_total_text_chars = 0 # 삭제
        # self.current_file_image_elements_count = 0 # 삭제
        # self.current_file_chart_elements_count = 0 # 삭제
        self.total_weighted_work = 0 # 프로그레스바 계산에는 여전히 사용
        self.current_weighted_done = 0

        # 히스토리 파일 경로
        self.history_file_path = os.path.join(HISTORY_DIR, "translation_history.json")
        self.translation_history_data: List[Dict[str, Any]] = []


        self.create_widgets()
        self._load_translation_history() # 히스토리 로드
        self.master.after(100, self.initial_checks)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        atexit.register(self.on_closing)

        log_file_path_msg = self.general_file_handler.baseFilename if self.general_file_handler else '미설정'
        logger.info(f"--- {APP_NAME} 시작됨 (일반 로그 파일: {log_file_path_msg}) ---")

    def _setup_logging_file_handler(self):
        if self.general_file_handler: return
        try:
            general_log_filename = os.path.join(LOGS_DIR, "app_general.log")
            self.general_file_handler = logging.FileHandler(general_log_filename, mode='a', encoding='utf-8')
            self.general_file_handler.setFormatter(formatter)
            if not any(h.baseFilename == os.path.abspath(general_log_filename) for h in root_logger.handlers if isinstance(h, logging.FileHandler)):
                root_logger.addHandler(self.general_file_handler)
        except Exception as e:
            print(f"일반 로그 파일 핸들러 설정 실패: {e}")

    def _destroy_current_ocr_handler(self):
        if self.ocr_handler:
            logger.info(f"기존 OCR 핸들러 ({self.current_ocr_engine_type}) 자원 해제 시도...")
            if hasattr(self.ocr_handler, 'ocr_engine') and self.ocr_handler.ocr_engine:
                try:
                    del self.ocr_handler.ocr_engine
                    logger.debug(f"{self.current_ocr_engine_type} 엔진 객체 참조 제거됨.")
                except Exception as e:
                    logger.warning(f"OCR 엔진 객체('ocr_engine') 삭제 중 오류: {e}")
            self.ocr_handler = None
            self.current_ocr_engine_type = None
            logger.info("기존 OCR 핸들러 자원 해제 완료.")

    def on_closing(self):
        logger.info("애플리케이션 종료 절차 시작...")
        if not self.stop_event.is_set():
            self.stop_event.set()
            if self.translation_thread and self.translation_thread.is_alive():
                logger.info("번역 스레드 종료 대기 중...")
                self.translation_thread.join(timeout=5)
                if self.translation_thread.is_alive(): logger.warning("번역 스레드가 시간 내에 종료되지 않았습니다.")
            if self.model_download_thread and self.model_download_thread.is_alive():
                logger.info("모델 다운로드 스레드 종료 대기 중...")
                self.model_download_thread.join(timeout=2)
                if self.model_download_thread.is_alive(): logger.warning("모델 다운로드 스레드가 시간 내에 정상 종료되지 않았습니다.")
            
            self._destroy_current_ocr_handler()

            if self.general_file_handler:
                logger.debug(f"일반 로그 파일 핸들러({self.general_file_handler.baseFilename}) 닫기 시도.")
                try:
                    self.general_file_handler.close()
                    root_logger.removeHandler(self.general_file_handler)
                    self.general_file_handler = None
                    logger.info("일반 로그 파일 핸들러가 성공적으로 닫혔습니다.")
                except Exception as e_log_close: logger.error(f"일반 로그 파일 핸들러 닫기 중 오류: {e_log_close}")
            else: logger.debug("일반 로그 파일 핸들러가 이미 닫혔거나 설정되지 않았습니다.")
        
        # 사용자 설정 저장 (필요시) - 현재는 구현 안 함

        if hasattr(self, 'master') and self.master.winfo_exists():
            # 종료 확인 메시지 박스는 그대로 두거나, 필요에 따라 자동 종료로 변경 가능
            # if messagebox.askokcancel("종료 확인", f"{APP_NAME}을(를) 종료하시겠습니까?"):
            logger.info("모든 정리 작업 완료. 애플리케이션을 종료합니다.")
            self.master.destroy()
            # else:
            #      logger.info("애플리케이션 종료 취소됨.")
            #      if self.stop_event.is_set(): self.stop_event.clear()
            #      return
        else: logger.info("애플리케이션 윈도우가 이미 없으므로 바로 종료합니다.")


    def initial_checks(self):
        logger.debug("초기 점검 시작: OCR 라이브러리 설치 여부 및 Ollama 상태 확인")
        self.update_ocr_status_display()
        self.check_ollama_status_manual(initial_check=True)
        logger.debug("초기 점검 완료.")

    def create_widgets(self):
            top_frame = ttk.Frame(self)
            top_frame.pack(fill=tk.BOTH, expand=True)
            
            bottom_frame = ttk.Frame(self, height=30)
            bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(5,0))
            bottom_frame.pack_propagate(False)

            main_paned_window = ttk.PanedWindow(top_frame, orient=tk.HORIZONTAL)
            main_paned_window.pack(fill=tk.BOTH, expand=True)

            # 좌측 패널 너비 비율 조정 (예: 2 -> 1.8 또는 1.7)
            left_panel = ttk.Frame(main_paned_window, padding=10)
            main_paned_window.add(left_panel, weight=3) # weight 조정으로 너비 감소

            right_panel = ttk.Frame(main_paned_window, padding=0) # 오른쪽 패널 패딩 제거 또는 조정
            main_paned_window.add(right_panel, weight=2)

            # --- Left Panel ---
            path_frame = ttk.LabelFrame(left_panel, text="파일 경로", padding=5)
            path_frame.pack(padx=5, pady=(0,5), fill=tk.X)
            self.file_path_var = tk.StringVar()
            file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, width=60)
            file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
            browse_button = ttk.Button(path_frame, text="찾아보기", command=self.browse_file)
            browse_button.pack(side=tk.LEFT)

            server_status_frame = ttk.LabelFrame(left_panel, text="서버 상태", padding=5)
            server_status_frame.pack(padx=5, pady=5, fill=tk.X)
            server_status_frame.columnconfigure(1, weight=1)
            
            self.os_label = ttk.Label(server_status_frame, text=f"OS: {platform.system()} {platform.release()}")
            self.os_label.grid(row=0, column=0, columnspan=2, padx=5, pady=2, sticky=tk.W)

            self.ollama_status_label = ttk.Label(server_status_frame, text="Ollama 설치: 미확인")
            self.ollama_status_label.grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
            self.ollama_running_label = ttk.Label(server_status_frame, text="Ollama 실행: 미확인")
            self.ollama_running_label.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
            self.ollama_port_label = ttk.Label(server_status_frame, text="Ollama 포트: -")
            self.ollama_port_label.grid(row=1, column=2, padx=5, pady=2, sticky=tk.W)
            self.ollama_check_button = ttk.Button(server_status_frame, text="Ollama 확인", command=self.check_ollama_status_manual)
            self.ollama_check_button.grid(row=1, column=3, padx=5, pady=2, sticky=tk.E)
            
            self.ocr_status_label = ttk.Label(server_status_frame, text="OCR 상태: 미확인")
            self.ocr_status_label.grid(row=2, column=0, columnspan=4, padx=5, pady=2, sticky=tk.W)

            # 파일 정보 및 진행 상황 프레임 (기존 작업량 표시는 삭제)
            file_progress_outer_frame = ttk.Frame(left_panel)
            file_progress_outer_frame.pack(padx=5, pady=5, fill=tk.X)

            file_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="파일 정보", padding=5)
            file_info_frame.pack(side=tk.LEFT, padx=(0,5), fill=tk.BOTH, expand=True)
            self.file_name_label = ttk.Label(file_info_frame, text="파일 이름: ")
            self.file_name_label.pack(anchor=tk.W, pady=1)
            self.slide_count_label = ttk.Label(file_info_frame, text="슬라이드 수: ")
            self.slide_count_label.pack(anchor=tk.W, pady=1)
            # self.text_elements_label 등 작업량 관련 레이블 삭제

            progress_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="진행 상황", padding=5)
            progress_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.current_slide_label = ttk.Label(progress_info_frame, text="현재 위치: -") # "현재 슬라이드" -> "현재 위치"
            self.current_slide_label.pack(anchor=tk.W, pady=1)
            self.current_work_label = ttk.Label(progress_info_frame, text="현재 작업: 대기 중") # 초기 메시지 변경
            self.current_work_label.pack(anchor=tk.W, pady=1)
            # self.translated_elements_label, self.remaining_elements_label 삭제

            translation_options_frame = ttk.LabelFrame(left_panel, text="번역 옵션", padding=5)
            translation_options_frame.pack(padx=5, pady=5, fill=tk.X)
            translation_options_frame.columnconfigure(1, weight=1)
            translation_options_frame.columnconfigure(4, weight=1)

            ttk.Label(translation_options_frame, text="원본 언어:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
            self.src_lang_var = tk.StringVar(value=SUPPORTED_LANGUAGES[0])
            self.src_lang_combo = ttk.Combobox(translation_options_frame, textvariable=self.src_lang_var, values=SUPPORTED_LANGUAGES, state="readonly", width=12)
            self.src_lang_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
            self.src_lang_combo.bind("<<ComboboxSelected>>", self.on_source_language_change)
            
            self.swap_button = ttk.Button(translation_options_frame, text="↔", command=self.swap_languages, width=3)
            self.swap_button.grid(row=0, column=2, padx=5, pady=5)
            
            ttk.Label(translation_options_frame, text="번역 언어:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
            self.tgt_lang_var = tk.StringVar(value=SUPPORTED_LANGUAGES[1])
            self.tgt_lang_combo = ttk.Combobox(translation_options_frame, textvariable=self.tgt_lang_var, values=SUPPORTED_LANGUAGES, state="readonly", width=12)
            self.tgt_lang_combo.grid(row=0, column=4, padx=5, pady=5, sticky=tk.EW)
            
            model_selection_frame = ttk.Frame(translation_options_frame) # 모델 콤보와 새로고침 버튼을 위한 프레임
            model_selection_frame.grid(row=1, column=1, columnspan=4, padx=0, pady=0, sticky=tk.EW)
            model_selection_frame.columnconfigure(0, weight=1) # 콤보박스가 확장되도록

            ttk.Label(translation_options_frame, text="번역 모델:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
            self.model_var = tk.StringVar(value=DEFAULT_MODEL)
            self.model_combo = ttk.Combobox(model_selection_frame, textvariable=self.model_var, state="disabled")
            self.model_combo.grid(row=0, column=0, padx=(5,0), pady=5, sticky=tk.EW) # 프레임 내에서 grid
            self.model_refresh_button = ttk.Button(model_selection_frame, text="🔄", command=self.load_ollama_models, width=3)
            self.model_refresh_button.grid(row=0, column=1, padx=(2,5), pady=5, sticky=tk.W)


            action_buttons_frame = ttk.Frame(left_panel, padding=(0,5,0,0))
            action_buttons_frame.pack(padx=5, pady=10, fill=tk.X)
            
            # 버튼 스타일 및 크기 변경
            self.style.configure("Big.TButton", font=('TkDefaultFont', 11, 'bold'), foreground="black") # 흰색 -> 검정색, 폰트 크기 증가
            # Accent.TButton 대신 Big.TButton 사용 또는 Accent.TButton 수정
            # self.style.configure("Accent.TButton", font=('Helvetica', 13, 'bold'), foreground="black", background="#0078D7") # 폰트 크기 30% 증가 (예: 10->13), 흰색->검정색

            self.start_button = ttk.Button(action_buttons_frame, text="번역 시작", command=self.start_translation, style="Big.TButton")
            self.start_button.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X, ipady=10) # ipady로 높이 증가 (기존 5 -> 10)
            
            self.stop_button = ttk.Button(action_buttons_frame, text="번역 중지", command=self.stop_translation, state=tk.DISABLED, style="Big.TButton")
            self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=10) # ipady로 높이 증가

            progress_bar_frame = ttk.Frame(left_panel)
            progress_bar_frame.pack(padx=5, pady=5, fill=tk.X)
            self.progress_bar = ttk.Progressbar(progress_bar_frame, orient="horizontal", length=300, mode="determinate")
            self.progress_bar.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
            self.progress_label_var = tk.StringVar(value="0%")
            ttk.Label(progress_bar_frame, textvariable=self.progress_label_var).pack(side=tk.LEFT)

            self.translated_file_path_var = tk.StringVar()
            translated_file_frame = ttk.LabelFrame(left_panel, text="번역 완료 파일", padding=5)
            translated_file_frame.pack(padx=5, pady=5, fill=tk.X)
            self.translated_file_entry = ttk.Entry(translated_file_frame, textvariable=self.translated_file_path_var, state="readonly", width=60)
            self.translated_file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
            self.open_folder_button = ttk.Button(translated_file_frame, text="폴더 열기", command=self.open_translated_folder, state=tk.DISABLED)
            self.open_folder_button.pack(side=tk.LEFT)

            # --- Right Panel (로그, 히스토리, 고급옵션) ---
            # 실행로그/번역히스토리 창 크기 조정 (상하 약 30% 감소)
            # 기존 right_panel_notebook이 right_panel 전체를 차지했으므로,
            # right_panel을 다시 분할하여 notebook과 advanced_options_frame을 배치
            
            right_top_frame = ttk.Frame(right_panel) # 로그와 히스토리용
            right_top_frame.pack(fill=tk.BOTH, expand=True) # 이 프레임이 기존 notebook 역할

            right_bottom_frame = ttk.Frame(right_panel, height=150) # 고급 옵션용 고정 프레임 (높이 조절 가능)
            right_bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(5,0))
            right_bottom_frame.pack_propagate(False) # 높이 고정


            right_panel_notebook = ttk.Notebook(right_top_frame) # 기존 notebook을 right_top_frame에 넣음
            # notebook의 높이를 직접 지정하거나, pack/grid 옵션으로 조절.
            # 예: right_panel_notebook.pack(fill=tk.BOTH, expand=True, ipady=-50) # 음수 ipady로 높이 줄이기 시도 (작동 안할수 있음)
            # 또는 right_top_frame의 높이를 제한하고 notebook은 expand=True
            right_panel_notebook.pack(fill=tk.BOTH, expand=True)


            log_tab_frame = ttk.Frame(right_panel_notebook, padding=5)
            right_panel_notebook.add(log_tab_frame, text="실행 로그")
            self.log_text = tk.Text(log_tab_frame, state=tk.DISABLED, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1, font=("TkFixedFont", 9))
            log_scrollbar_y = ttk.Scrollbar(log_tab_frame, orient="vertical", command=self.log_text.yview)
            self.log_text.config(yscrollcommand=log_scrollbar_y.set)
            log_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            self.log_text.pack(fill=tk.BOTH, expand=True)
            
            text_widget_handler = TextHandler(self.log_text)
            text_widget_handler.setFormatter(formatter)
            if not any(isinstance(h, TextHandler) for h in root_logger.handlers):
                root_logger.addHandler(text_widget_handler)

            history_tab_frame = ttk.Frame(right_panel_notebook, padding=5)
            right_panel_notebook.add(history_tab_frame, text="번역 히스토리")
            # 히스토리 컬럼 추가 (모델명, OCR 온도 등)
            history_columns = ("name", "src", "tgt", "model", "ocr_temp", "status", "time", "path")
            self.history_tree = ttk.Treeview(history_tab_frame, columns=history_columns, show="headings")
            self.history_tree.heading("name", text="문서 이름"); self.history_tree.column("name", width=120, anchor=tk.W, stretch=tk.YES)
            self.history_tree.heading("src", text="원본"); self.history_tree.column("src", width=50, anchor=tk.CENTER)
            self.history_tree.heading("tgt", text="대상"); self.history_tree.column("tgt", width=50, anchor=tk.CENTER)
            self.history_tree.heading("model", text="모델"); self.history_tree.column("model", width=100, anchor=tk.W)
            self.history_tree.heading("ocr_temp", text="OCR온도"); self.history_tree.column("ocr_temp", width=60, anchor=tk.CENTER)
            self.history_tree.heading("status", text="결과"); self.history_tree.column("status", width=60, anchor=tk.CENTER)
            self.history_tree.heading("time", text="번역일시"); self.history_tree.column("time", width=110, anchor=tk.CENTER)
            self.history_tree.heading("path", text="경로"); self.history_tree.column("path", width=0, stretch=tk.NO)

            hist_scrollbar_y = ttk.Scrollbar(history_tab_frame, orient="vertical", command=self.history_tree.yview)
            hist_scrollbar_x = ttk.Scrollbar(history_tab_frame, orient="horizontal", command=self.history_tree.xview)
            self.history_tree.configure(yscrollcommand=hist_scrollbar_y.set, xscrollcommand=hist_scrollbar_x.set)
            hist_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            hist_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            self.history_tree.pack(fill=tk.BOTH, expand=True)
            self.history_tree.bind("<Double-1>", self.on_history_double_click)

            # --- 고급 옵션 프레임 (right_bottom_frame 내부에) ---
            self.advanced_options_visible = tk.BooleanVar(value=False)
            
            advanced_options_toggle_button = ttk.Button(
                right_bottom_frame, text="▼ 고급 옵션 보기",
                command=self.toggle_advanced_options
            )
            advanced_options_toggle_button.pack(fill=tk.X, padx=5, pady=(5,0))

            self.advanced_options_frame = ttk.Frame(right_bottom_frame, padding=5)
            # 초기에는 숨겨져 있음 (pack 안 함)

            # 이미지 번역 온도 설정
            temp_frame = ttk.Frame(self.advanced_options_frame)
            temp_frame.pack(fill=tk.X, pady=2)
            ttk.Label(temp_frame, text="이미지 번역 온도:").pack(side=tk.LEFT, padx=(0,5))
            
            self.ocr_temperature_var = tk.DoubleVar(value=config.DEFAULT_OCR_TEMPERATURE)
            self.ocr_temp_slider = ttk.Scale(
                temp_frame, from_=0.1, to=1.0, variable=self.ocr_temperature_var,
                orient=tk.HORIZONTAL, command=self._update_ocr_temp_label
            )
            self.ocr_temp_slider.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
            self.ocr_temp_value_label = ttk.Label(temp_frame, text=f"{config.DEFAULT_OCR_TEMPERATURE:.1f}")
            self.ocr_temp_value_label.pack(side=tk.LEFT, padx=(5,0))

            temp_description_frame = ttk.Frame(self.advanced_options_frame)
            temp_description_frame.pack(fill=tk.X, pady=(0,5))
            ttk.Label(temp_description_frame, text="0.1 (정직함) <----------------------> 1.0 (창의적)", justify=tk.CENTER).pack(fill=tk.X)
            ttk.Label(temp_description_frame, text="(기본값: 0.4, 이미지 품질이 좋지 않을 경우 수치를 올리는 것이 번역에 도움 될 수 있음)", wraplength=300, justify=tk.LEFT, font=("TkDefaultFont",8)).pack(fill=tk.X)


            # 이미지 번역 실행 여부
            self.image_translation_enabled_var = tk.BooleanVar(value=config.DEFAULT_IMAGE_TRANSLATION_ENABLED)
            image_trans_check = ttk.Checkbutton(
                self.advanced_options_frame, text="이미지 내 텍스트 번역 실행",
                variable=self.image_translation_enabled_var
            )
            image_trans_check.pack(anchor=tk.W, padx=5, pady=2)

            # 이미지 번역 시 GPU 사용 여부 (OCR GPU)
            self.ocr_use_gpu_var = tk.BooleanVar(value=config.DEFAULT_OCR_USE_GPU)
            ocr_gpu_check = ttk.Checkbutton(
                self.advanced_options_frame, text="이미지 번역(OCR) 시 GPU 사용 (지원 시)",
                variable=self.ocr_use_gpu_var, command=self.on_ocr_gpu_setting_change
            )
            ocr_gpu_check.pack(anchor=tk.W, padx=5, pady=2)


            # --- 하단 로고 ---
            logo_path_bottom = os.path.join(ASSETS_DIR, "LINEstudio2.png")
            if os.path.exists(logo_path_bottom):
                try:
                    pil_temp_for_size = Image.open(logo_path_bottom)
                    original_width, original_height = pil_temp_for_size.size
                    pil_temp_for_size.close()
                    target_height_bottom = 20
                    subsample_factor = max(1, int(original_height / target_height_bottom)) if original_height > target_height_bottom and target_height_bottom > 0 else (1 if original_height > 0 else 6)
                    
                    temp_logo_image_bottom = tk.PhotoImage(file=logo_path_bottom, master=self.master)
                    self.logo_image_tk_bottom = temp_logo_image_bottom.subsample(subsample_factor, subsample_factor)
                    logo_label_bottom = ttk.Label(bottom_frame, image=self.logo_image_tk_bottom)
                    logo_label_bottom.pack(side=tk.RIGHT, padx=10, pady=2)
                except Exception as e_general_bottom:
                    logger.warning(f"하단 로고 로드 중 예외: {e_general_bottom}", exc_info=True)
            else:
                logger.warning(f"하단 로고 파일({logo_path_bottom})을 찾을 수 없습니다.")

    def _update_ocr_temp_label(self, value_str):
        try:
            value = float(value_str)
            self.ocr_temp_value_label.config(text=f"{value:.1f}")
        except ValueError:
            pass # 슬라이더 이동 중 일시적으로 발생 가능

    def toggle_advanced_options(self):
        if self.advanced_options_visible.get():
            self.advanced_options_frame.pack_forget()
            self.advanced_options_visible.set(False)
            # 토글 버튼 텍스트 변경 (Frame 찾아서)
            # This way of finding the button is a bit fragile if the UI structure changes a lot.
            # A more robust way would be to keep a reference to the button itself.
            # For now, assuming the structure from create_widgets.
            try:
                right_panel_widget = self.nametowidget(self.winfo_children()[0].panes()[1]) # main_paned_window -> right_panel
                right_bottom_frame_widget = right_panel_widget.winfo_children()[1] # right_bottom_frame
                toggle_button_widget = right_bottom_frame_widget.winfo_children()[0] # The toggle button
                if isinstance(toggle_button_widget, ttk.Button):
                    toggle_button_widget.config(text="▼ 고급 옵션 보기")
            except Exception as e:
                logger.warning(f"고급 옵션 토글 버튼 텍스트 변경 실패 (숨길 때): {e}")

        else:
            # Pack 'before' the logo which is assumed to be the last child of the bottom_frame of the main Application frame.
            # This is also a bit fragile.
            try:
                # Find the main bottom_frame (where the logo is)
                main_bottom_frame_widget = self.winfo_children()[1] # Assumes bottom_frame is the second child of Application
                # If logo exists, it's likely the last child of main_bottom_frame_widget
                logo_widget_or_last_child = main_bottom_frame_widget.winfo_children()[-1] if main_bottom_frame_widget.winfo_children() else None
                
                # Pack the advanced_options_frame inside its parent (right_bottom_frame)
                self.advanced_options_frame.pack(fill=tk.X, padx=5, pady=5) # Simpler packing
                self.advanced_options_visible.set(True)

                right_panel_widget = self.nametowidget(self.winfo_children()[0].panes()[1])
                right_bottom_frame_widget = right_panel_widget.winfo_children()[1]
                toggle_button_widget = right_bottom_frame_widget.winfo_children()[0]
                if isinstance(toggle_button_widget, ttk.Button):
                    toggle_button_widget.config(text="▲ 고급 옵션 숨기기")
            except Exception as e:
                logger.warning(f"고급 옵션 토글 버튼 텍스트 변경 실패 (보일 때) 또는 고급 옵션 프레임 표시 오류: {e}")

    
    def on_ocr_gpu_setting_change(self):
        logger.info(f"OCR GPU 사용 설정 변경: {self.ocr_use_gpu_var.get()}. OCR 핸들러 재초기화 필요할 수 있음.")
        # 다음 번역 시작 시 check_ocr_engine_status에서 이 값을 반영하여 OCR 핸들러를 (재)생성함.
        # 즉시 재초기화보다는, 사용자가 번역 시작 시점에 적용되도록 하는 것이 더 자연스러울 수 있음.
        # 만약 즉시 반영을 원한다면 여기서 self._destroy_current_ocr_handler() 호출 후 self.check_ocr_engine_status() 호출.
        # 현재는 다음 번역 시 적용. 상태 표시는 다음 언어 변경 또는 번역 시작 시 업데이트됨.
        self.update_ocr_status_display() # 상태 표시 업데이트 시도


    def _load_translation_history(self):
        if os.path.exists(self.history_file_path):
            try:
                with open(self.history_file_path, 'r', encoding='utf-8') as f:
                    self.translation_history_data = json.load(f)
                # 최신순으로 정렬 (시간 기준 내림차순)
                self.translation_history_data.sort(key=lambda x: x.get('time', '0'), reverse=True)
                # 최대 개수 제한
                self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
            except json.JSONDecodeError:
                logger.error(f"번역 히스토리 파일({self.history_file_path}) 디코딩 오류. 새 히스토리 시작.")
                self.translation_history_data = []
            except Exception as e:
                logger.error(f"번역 히스토리 로드 중 오류: {e}", exc_info=True)
                self.translation_history_data = []
        else:
            self.translation_history_data = []

        self._populate_history_treeview()

    def _save_translation_history(self):
        try:
            # 저장 전 다시 최신순 정렬 및 개수 제한 (중복 방지 및 일관성)
            self.translation_history_data.sort(key=lambda x: x.get('time', '0'), reverse=True)
            self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
            with open(self.history_file_path, 'w', encoding='utf-8') as f:
                json.dump(self.translation_history_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logger.error(f"번역 히스토리 저장 중 오류: {e}", exc_info=True)

    def _add_history_entry(self, entry: Dict[str, Any]):
        # entry: {"name": ..., "src": ..., "tgt": ..., "model": ..., "ocr_temp": ..., "status": ..., "time": ..., "path": ...}
        self.translation_history_data.insert(0, entry) # 새 항목을 맨 앞에 추가
        self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS] # 개수 제한
        self._save_translation_history()
        self._populate_history_treeview()


    def _populate_history_treeview(self):
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()):
            return
        # 기존 항목 모두 삭제
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        # 새 데이터로 채우기 (self.translation_history_data는 이미 최신순)
        for entry in self.translation_history_data:
            values = (
                entry.get("name", "-"),
                entry.get("src", "-"),
                entry.get("tgt", "-"),
                entry.get("model", "-"),
                f"{entry.get('ocr_temp', '-')}", # 숫자로 나올 수 있으므로 문자열화
                entry.get("status", "-"),
                entry.get("time", "-"),
                entry.get("path", "-")
            )
            self.history_tree.insert("", tk.END, values=values)
        if self.translation_history_data:
            self.history_tree.yview_moveto(0) # 가장 최근 항목이 보이도록 스크롤


    def update_ocr_status_display(self):
        selected_ui_lang = self.src_lang_var.get()
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name_display = "EasyOCR" if use_easyocr else "PaddleOCR"
        
        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get()
        gpu_status_text = "(GPU 사용 예정)" if gpu_enabled_for_ocr else "(CPU 사용 예정)"

        if self.ocr_handler and self.current_ocr_engine_type == engine_name_display.lower():
            current_handler_lang_display = ""
            if self.current_ocr_engine_type == "paddleocr": # 소문자로 비교
                current_handler_lang_display = self.ocr_handler.current_lang_codes # Paddle은 단일 코드
            elif self.current_ocr_engine_type == "easyocr" and self.ocr_handler.current_lang_codes:
                current_handler_lang_display = ", ".join(self.ocr_handler.current_lang_codes) # EasyOCR은 리스트
            
            gpu_in_use_text = "(GPU 사용 중)" if self.ocr_handler.use_gpu else "(CPU 사용 중)"
            self.ocr_status_label.config(text=f"{engine_name_display}: 준비됨 ({current_handler_lang_display}) {gpu_in_use_text}")
        else:
            self.ocr_status_label.config(text=f"{engine_name_display}: ({selected_ui_lang}) 사용 예정 {gpu_status_text} (미확인)")


    def on_source_language_change(self, event=None):
        selected_ui_lang = self.src_lang_var.get()
        logger.info(f"원본 언어 변경됨: {selected_ui_lang}.")
        self.update_ocr_status_display()
        if self.file_path_var.get():
            self.load_file_info(self.file_path_var.get()) # 파일 정보 다시 로드 (가중치 계산 등)

    def browse_file(self):
        file_path = filedialog.askopenfilename(title="파워포인트 파일 선택", filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*")))
        if file_path:
            self.file_path_var.set(file_path); logger.info(f"파일 선택됨: {file_path}")
            self.load_file_info(file_path); self.translated_file_path_var.set("")
            self.open_folder_button.config(state=tk.DISABLED)
            self.current_work_label.config(text="파일 선택됨. 번역 대기 중.") # 상태 업데이트

    def load_file_info(self, file_path):
        self.current_work_label.config(text="파일 분석 중...")
        self.master.update_idletasks() # UI 업데이트 강제
        info = {"slide_count": 0, "text_elements_count": 0, "total_text_char_count":0, "image_elements_count": 0, "chart_elements_count": 0}
        try:
            logger.debug(f"파일 정보 분석 중: {file_path}"); file_name = os.path.basename(file_path)
            # PptxHandler에서 파일 정보를 가져오는 부분은 그대로 사용
            info = self.pptx_handler.get_file_info(file_path)
            
            self.current_file_slide_count = info.get('slide_count', 0)
            # text_elements_count, total_text_char_count 등은 UI에서 삭제되었으므로 내부 변수만 업데이트하거나,
            # 가중치 계산에 필요한 값만 유지합니다.
            _text_chars = info.get('total_text_char_count', 0)
            _images = info.get('image_elements_count', 0)
            _charts = info.get('chart_elements_count', 0)

            self.file_name_label.config(text=f"파일 이름: {file_name}")
            self.slide_count_label.config(text=f"슬라이드 수: {self.current_file_slide_count}")
            # self.text_elements_label 등 삭제

            self.total_weighted_work = (_text_chars * config.WEIGHT_TEXT_CHAR) + \
                                       (_images * config.WEIGHT_IMAGE) + \
                                       (_charts * config.WEIGHT_CHART)
            
            # self.total_weighted_work_label 등 삭제
            # self.remaining_elements_label 등 삭제
            # self.translated_elements_label 등 삭제
            logger.info(f"파일 정보 분석 완료. 총 슬라이드: {self.current_file_slide_count}, 예상 가중 작업량: {self.total_weighted_work}")
            self.current_work_label.config(text="파일 분석 완료. 번역 대기 중.")
        except Exception as e:
            logger.error(f"파일 정보 분석 오류: {e}", exc_info=True)
            self.file_name_label.config(text="파일 이름: - (오류)")
            self.slide_count_label.config(text="슬라이드 수: -")
            self.total_weighted_work = 0
            self.current_work_label.config(text="파일 분석 실패!")
            messagebox.showerror("파일 분석 오류", f"선택한 파일({os.path.basename(file_path)})을 분석하는 중 오류가 발생했습니다.\n파일이 손상되었거나 지원하지 않는 형식일 수 있습니다.\n\n오류: {e}")


    def check_ollama_status_manual(self, initial_check=False):
        logger.info("Ollama 상태 확인 중...")
        self.ollama_check_button.config(state=tk.DISABLED) # 확인 중 버튼 비활성화
        self.master.update_idletasks()

        ollama_installed = self.ollama_service.is_installed()
        self.ollama_status_label.config(text=f"Ollama 설치: {'설치됨' if ollama_installed else '미설치'}")
        if not ollama_installed:
            logger.warning("Ollama가 설치되어 있지 않습니다.")
            if not initial_check and messagebox.askyesno("Ollama 설치 필요", "Ollama가 설치되어 있지 않습니다. Ollama 다운로드 페이지로 이동하시겠습니까?"): webbrowser.open("https://ollama.com/download")
            self.ollama_running_label.config(text="Ollama 실행: 미설치"); self.ollama_port_label.config(text="Ollama 포트: -")
            self.model_combo.config(values=[], state="disabled"); self.model_var.set("");
            self.ollama_check_button.config(state=tk.NORMAL)
            return
        
        ollama_running, port = self.ollama_service.is_running()
        self.ollama_running_label.config(text=f"Ollama 실행: {'실행 중' if ollama_running else '미실행'}"); self.ollama_port_label.config(text=f"Ollama 포트: {port if ollama_running and port else '-'}")
        if ollama_running:
            logger.info(f"Ollama 실행 중 (포트: {port}). 모델 목록 로드 시도.")
            self.load_ollama_models()
        else:
            logger.warning("Ollama가 설치되었으나 실행 중이지 않습니다. 자동 시작을 시도합니다.")
            self.model_combo.config(values=[], state="disabled"); self.model_var.set("")
            if initial_check or messagebox.askyesno("Ollama 실행 필요", "Ollama가 실행 중이지 않습니다. 지금 시작하시겠습니까? (권장)"):
                if self.ollama_service.start_ollama():
                    logger.info("Ollama 자동 시작 성공. 잠시 후 상태를 다시 확인합니다.")
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(3000, lambda: self.check_ollama_status_manual(initial_check=initial_check)) # 대기 시간 증가
                else:
                    logger.error("Ollama 자동 시작 실패. 수동으로 실행해주세요.")
                    if not initial_check: messagebox.showwarning("Ollama 시작 실패", "Ollama를 자동으로 시작할 수 없습니다. 수동으로 실행 후 'Ollama 확인'을 눌러주세요.")
            
        self.ollama_check_button.config(state=tk.NORMAL)


    def load_ollama_models(self):
        logger.debug("Ollama 모델 목록 로드 중...")
        self.model_refresh_button.config(state=tk.DISABLED) # 새로고침 중 버튼 비활성화
        self.master.update_idletasks()

        models = self.ollama_service.get_text_models()
        if models:
            self.model_combo.config(values=models, state="readonly")
            # 현재 선택된 모델이 목록에 있으면 유지, 없으면 기본 모델 또는 첫 번째 모델 선택
            current_selected_model = self.model_var.get()
            if current_selected_model in models:
                self.model_var.set(current_selected_model)
            elif DEFAULT_MODEL in models:
                self.model_var.set(DEFAULT_MODEL)
            elif models: # 목록에 모델이 하나라도 있으면
                self.model_var.set(models[0])
            
            logger.info(f"사용 가능 Ollama 모델: {models}")
            if DEFAULT_MODEL not in models and not self.model_var.get(): # 기본 모델도 없고, 선택된 모델도 없을 때
                self.download_default_model_if_needed(initial_check_from_ollama=True)
        else:
            self.model_combo.config(values=[], state="disabled"); self.model_var.set("")
            logger.warning("Ollama에 로드된 모델이 없습니다.")
            self.download_default_model_if_needed(initial_check_from_ollama=True)
        
        self.model_refresh_button.config(state=tk.NORMAL)


    def download_default_model_if_needed(self, initial_check_from_ollama=False):
        current_models = self.ollama_service.get_text_models() # 최신 모델 목록 가져오기
        if DEFAULT_MODEL not in current_models:
            logger.warning(f"기본 모델 ({DEFAULT_MODEL})이 설치되어 있지 않습니다.")
            if initial_check_from_ollama or messagebox.askyesno("기본 모델 다운로드", f"기본 번역 모델 '{DEFAULT_MODEL}'이(가) 없습니다. 지금 다운로드하시겠습니까? (시간 소요)"):
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드 시작...")
                self.start_button.config(state=tk.DISABLED); self.progress_bar["value"] = 0
                self.current_work_label.config(text=f"모델 다운로드 시작: {DEFAULT_MODEL}") # 상태 업데이트
                self.progress_label_var.set(f"모델 다운로드 시작: {DEFAULT_MODEL}")
                if self.model_download_thread and self.model_download_thread.is_alive():
                    logger.warning("이미 모델 다운로드 스레드가 실행 중입니다."); return
                self.stop_event.clear() # 혹시 모를 중지 이벤트 초기화
                self.model_download_thread = threading.Thread(target=self._model_download_worker, args=(DEFAULT_MODEL, self.stop_event), daemon=True)
                self.model_download_thread.start()
            else: logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드가 취소되었습니다.")
        else: logger.info(f"기본 모델 ({DEFAULT_MODEL})이 이미 설치되어 있습니다.")


    def _model_download_worker(self, model_name, stop_event_ref):
        success = self.ollama_service.pull_model_with_progress(model_name, self.update_model_download_progress, stop_event=stop_event_ref)
        if hasattr(self, 'master') and self.master.winfo_exists():
            self.master.after(0, self._model_download_finished, model_name, success)
        self.model_download_thread = None

    def _model_download_finished(self, model_name, success):
        if success:
            logger.info(f"'{model_name}' 모델 다운로드 완료."); self.load_ollama_models()
            self.current_work_label.config(text=f"모델 '{model_name}' 다운로드 완료.")
        else:
            logger.error(f"'{model_name}' 모델 다운로드 실패.")
            self.current_work_label.config(text=f"모델 '{model_name}' 다운로드 실패.")
            if not self.stop_event.is_set(): # 사용자가 중지한게 아니라면
                messagebox.showerror("모델 다운로드 실패", f"'{model_name}' 모델 다운로드에 실패했습니다.\nOllama 서버 로그 또는 인터넷 연결을 확인해주세요.")
        
        if not (self.translation_thread and self.translation_thread.is_alive()): # 다른 번역 작업이 없다면
            self.start_button.config(state=tk.NORMAL)
            self.progress_bar["value"] = 0; self.progress_label_var.set("0%")
            if not success : self.current_work_label.config(text="모델 다운로드 실패. 재시도 요망.")
            else: self.current_work_label.config(text="대기 중")


    def update_model_download_progress(self, status_text, completed_bytes, total_bytes, is_error=False):
        if self.stop_event.is_set() and "중지됨" not in status_text : return # 중지 이벤트 처리 강화
        
        if total_bytes > 0: percent = (completed_bytes / total_bytes) * 100; progress_str = f"{percent:.1f}%"
        else: percent = 0; progress_str = status_text # 다운로드 시작 전 상태 메시지
        
        def _update():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            if not is_error:
                self.progress_bar["value"] = percent
                self.progress_label_var.set(f"모델 다운로드: {progress_str} ({status_text})")
                self.current_work_label.config(text=f"모델 다운로드 중: {status_text} {progress_str}")
            else:
                self.progress_label_var.set(f"모델 다운로드 오류: {status_text}")
                self.current_work_label.config(text=f"모델 다운로드 오류: {status_text}")

            # 로그는 DEBUG 레벨로 유지하여 너무 많이 쌓이지 않도록 함
            logger.log(logging.DEBUG if not is_error else logging.ERROR,
                       f"모델 다운로드 진행: {status_text} ({completed_bytes}/{total_bytes})")
        if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, _update)


    def check_ocr_engine_status(self, is_called_from_start_translation=False):
        self.current_work_label.config(text="OCR 엔진 확인 중...")
        self.master.update_idletasks()

        selected_ui_lang = self.src_lang_var.get()
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name_display = "EasyOCR" if use_easyocr else "PaddleOCR"
        engine_name_internal = engine_name_display.lower() # 내부 비교용

        ocr_lang_code = None
        if use_easyocr:
            ocr_lang_code = config.UI_LANG_TO_EASYOCR_CODE_MAP.get(selected_ui_lang)
        else:
            ocr_lang_code = config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(selected_ui_lang, config.DEFAULT_PADDLE_OCR_LANG)

        if not ocr_lang_code:
            msg = f"{engine_name_display}: 언어 '{selected_ui_lang}'에 대한 OCR 코드 없음."
            self.ocr_status_label.config(text=msg)
            logger.error(msg)
            if is_called_from_start_translation: messagebox.showerror("OCR 설정 오류", msg)
            self.current_work_label.config(text="OCR 설정 오류!")
            return False

        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get() # 고급 옵션에서 GPU 사용 여부 가져오기

        needs_reinit = False
        if not self.ocr_handler: needs_reinit = True
        elif self.current_ocr_engine_type != engine_name_internal: needs_reinit = True
        elif self.ocr_handler.use_gpu != gpu_enabled_for_ocr: needs_reinit = True # GPU 설정 변경 시 재초기화
        elif engine_name_internal == "paddleocr" and self.ocr_handler.current_lang_codes != ocr_lang_code: needs_reinit = True
        elif engine_name_internal == "easyocr" and (not self.ocr_handler.current_lang_codes or ocr_lang_code not in self.ocr_handler.current_lang_codes):
            needs_reinit = True
        
        if needs_reinit:
            self._destroy_current_ocr_handler()
            logger.info(f"{engine_name_display} 핸들러 (재)초기화 시도 (언어: {ocr_lang_code}, GPU: {gpu_enabled_for_ocr}).")
            self.current_work_label.config(text=f"{engine_name_display} 엔진 로딩 중 (언어: {ocr_lang_code}, GPU: {gpu_enabled_for_ocr})...")
            self.master.update_idletasks()
            try:
                if use_easyocr:
                    if not utils.check_easyocr():
                        self.ocr_status_label.config(text=f"{engine_name_display}: 미설치")
                        if messagebox.askyesno(f"{engine_name_display} 설치 필요", f"{engine_name_display}이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            # 설치 로직은 utils에 있다고 가정
                            if utils.install_easyocr(): messagebox.showinfo(f"{engine_name_display} 설치 완료", f"{engine_name_display}이(가) 설치되었습니다. 애플리케이션을 재시작하거나 다시 시도해주세요.")
                            else: messagebox.showerror(f"{engine_name_display} 설치 실패", f"{engine_name_display} 설치에 실패했습니다.")
                        self.current_work_label.config(text=f"{engine_name_display} 미설치.")
                        return False
                    self.ocr_handler = EasyOcrHandler(lang_codes_list=[ocr_lang_code], debug_enabled=debug_mode, use_gpu=gpu_enabled_for_ocr)
                    self.current_ocr_engine_type = "easyocr"
                else:
                    if not utils.check_paddleocr():
                        self.ocr_status_label.config(text=f"{engine_name_display}: 미설치")
                        if messagebox.askyesno(f"{engine_name_display} 설치 필요", f"{engine_name_display}이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            if utils.install_paddleocr(): messagebox.showinfo(f"{engine_name_display} 설치 완료", f"{engine_name_display}이(가) 설치되었습니다. 애플리케이션을 재시작하거나 다시 시도해주세요.")
                            else: messagebox.showerror(f"{engine_name_display} 설치 실패", f"{engine_name_display} 설치에 실패했습니다.")
                        self.current_work_label.config(text=f"{engine_name_display} 미설치.")
                        return False
                    self.ocr_handler = PaddleOcrHandler(lang_code=ocr_lang_code, debug_enabled=debug_mode, use_gpu=gpu_enabled_for_ocr)
                    self.current_ocr_engine_type = "paddleocr"
                
                logger.info(f"{engine_name_display} 핸들러 초기화 성공 (언어: {ocr_lang_code}, GPU: {gpu_enabled_for_ocr}).")
                self.current_work_label.config(text=f"{engine_name_display} 엔진 로딩 완료.")

            except RuntimeError as e:
                logger.error(f"{engine_name_display} 핸들러 초기화 실패: {e}", exc_info=True)
                self.ocr_status_label.config(text=f"{engine_name_display}: 초기화 실패 ({ocr_lang_code}, GPU:{gpu_enabled_for_ocr})")
                if is_called_from_start_translation: messagebox.showerror(f"{engine_name_display} 오류", f"{engine_name_display} 초기화 중 오류:\n{e}\n\nGPU 관련 문제일 수 있습니다. GPU 사용 옵션을 확인해보세요.")
                self._destroy_current_ocr_handler()
                self.current_work_label.config(text=f"{engine_name_display} 엔진 초기화 실패!")
                return False
            except Exception as e_other:
                 logger.error(f"{engine_name_display} 핸들러 생성 중 예기치 않은 오류: {e_other}", exc_info=True)
                 self.ocr_status_label.config(text=f"{engine_name_display}: 알 수 없는 오류")
                 if is_called_from_start_translation: messagebox.showerror(f"{engine_name_display} 오류", f"{engine_name_display} 처리 중 예기치 않은 오류:\n{e_other}")
                 self._destroy_current_ocr_handler()
                 self.current_work_label.config(text=f"{engine_name_display} 엔진 오류!")
                 return False

        # OCR 핸들러 상태 최종 업데이트
        self.update_ocr_status_display()
        if self.ocr_handler and self.ocr_handler.ocr_engine:
            return True
        else: # 핸들러가 여전히 준비되지 않았다면
            self.ocr_status_label.config(text=f"{engine_name_display} OCR: 준비 안됨 ({selected_ui_lang})")
            if is_called_from_start_translation and not needs_reinit : # 재초기화 시도가 없었는데도 준비 안된 경우
                 messagebox.showwarning("OCR 오류", f"{engine_name_display} OCR 엔진을 사용할 수 없습니다. 이전 로그를 확인해주세요.")
            self.current_work_label.config(text=f"{engine_name_display} OCR 준비 안됨.")
            return False


    def swap_languages(self):
        src, tgt = self.src_lang_var.get(), self.tgt_lang_var.get()
        self.src_lang_var.set(tgt); self.tgt_lang_var.set(src)
        logger.info(f"언어 스왑: {tgt} <-> {src}")
        self.on_source_language_change() # 여기서 OCR 상태 업데이트 및 파일 정보 재로드

    def start_translation(self):
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "번역할 유효한 파워포인트 파일을 선택해주세요.\n'찾아보기' 버튼을 사용하여 파일을 선택할 수 있습니다.")
            return

        # 이미지 번역 실행 여부 (고급 옵션)
        image_translation_really_enabled = self.image_translation_enabled_var.get()
        
        if image_translation_really_enabled: # 이미지 번역 옵션이 켜져 있을 때만 OCR 엔진 체크
            if not self.check_ocr_engine_status(is_called_from_start_translation=True):
                if not messagebox.askyesno("OCR 준비 실패",
                                         "이미지 내 텍스트 번역에 필요한 OCR 기능이 준비되지 않았거나 사용할 수 없습니다.\n"
                                         "이 경우 이미지 안의 글자는 번역되지 않습니다.\n"
                                         "계속 진행하시겠습니까? (텍스트/차트만 번역)"):
                    logger.warning("OCR 준비 실패로 사용자가 번역을 취소했습니다.")
                    self.current_work_label.config(text="번역 취소됨 (OCR 준비 실패).")
                    return
                logger.warning("OCR 핸들러 준비 실패. 이미지 번역 없이 진행합니다.")
                image_translation_really_enabled = False # 실제로는 이미지 번역 안함
        else:
            logger.info("이미지 번역 옵션이 꺼져있으므로 OCR 엔진을 확인하지 않습니다.")
            self._destroy_current_ocr_handler() # OCR 핸들러가 있다면 명시적으로 해제


        src_lang, tgt_lang, model = self.src_lang_var.get(), self.tgt_lang_var.get(), self.model_var.get()
        if not model:
            messagebox.showerror("모델 오류", "번역 모델을 선택해주세요.\nOllama 서버가 실행 중이고 모델이 다운로드되었는지 확인하세요.\n'Ollama 확인' 버튼과 모델 목록 '🔄' 버튼을 사용해볼 수 있습니다.");
            self.check_ollama_status_manual(); return
        if src_lang == tgt_lang:
            messagebox.showwarning("언어 동일", "원본 언어와 번역 언어가 동일합니다.\n다른 언어를 선택해주세요."); return
        
        ollama_running, _ = self.ollama_service.is_running()
        if not ollama_running:
            messagebox.showerror("Ollama 미실행", "Ollama 서버가 실행 중이지 않습니다.\nOllama를 실행한 후 'Ollama 확인' 버튼을 눌러주세요.");
            self.check_ollama_status_manual(); return

        if self.total_weighted_work <= 0:
            logger.info("총 예상 작업량이 0입니다. 파일 정보를 다시 로드하여 확인합니다.")
            self.load_file_info(file_path) # 여기서 오류 발생 시 메시지 박스 표시됨
            if self.total_weighted_work <= 0: # load_file_info 후에도 0이면
                messagebox.showinfo("정보", "번역할 내용이 없거나 작업량을 계산할 수 없습니다.\n파일 내용을 확인해주세요.")
                logger.warning("재확인 후에도 총 예상 작업량이 0 이하입니다. 번역을 시작하지 않습니다.")
                self.current_work_label.config(text="번역할 내용 없음.")
                return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_filename = os.path.basename(file_path)
        safe_original_filename_part = "".join(c if c.isalnum() or c in ['.', '_'] else '_' for c in os.path.splitext(original_filename)[0])
        task_log_filename = f"translation_{timestamp}_{safe_original_filename_part}.log"
        task_log_filepath = os.path.join(LOGS_DIR, task_log_filename)
        
        ocr_engine_for_log = self.current_ocr_engine_type if image_translation_really_enabled and self.ocr_handler else '사용 안 함'
        ocr_temp_for_log = self.ocr_temperature_var.get() if image_translation_really_enabled else 'N/A'
        ocr_gpu_for_log = self.ocr_use_gpu_var.get() if image_translation_really_enabled and self.ocr_handler else 'N/A'

        logger.info(f"번역 시작: '{original_filename}' ({src_lang} -> {tgt_lang}) using {model}. "
                    f"이미지 번역: {'활성' if image_translation_really_enabled else '비활성'}, "
                    f"OCR 엔진: {ocr_engine_for_log}, OCR 온도: {ocr_temp_for_log}, OCR GPU: {ocr_gpu_for_log}")
        
        self.start_button.config(state=tk.DISABLED); self.stop_button.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0; self.progress_label_var.set("0%")
        self.translated_file_path_var.set(""); self.open_folder_button.config(state=tk.DISABLED)
        
        self.current_weighted_done = 0
        # self.translated_elements_label 등 삭제

        self.stop_event.clear()
        
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("이미 번역 스레드가 실행 중입니다."); messagebox.showwarning("번역 중복", "이미 다른 번역 작업이 진행 중입니다.")
            self.start_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED); return

        self.current_work_label.config(text="번역 준비 중...") # 시작 전 상태
        self.master.update_idletasks()

        # 고급 옵션에서 가져온 값들
        ocr_temperature_to_use = self.ocr_temperature_var.get()

        self.translation_thread = threading.Thread(target=self._translation_worker,
                                                   args=(file_path, src_lang, tgt_lang, model, task_log_filepath,
                                                         image_translation_really_enabled, ocr_temperature_to_use),
                                                   daemon=True)
        self.start_time = time.time()
        self.translation_thread.start()
        self.update_progress_timer()


    def _translation_worker(self, file_path, src_lang, tgt_lang, model, task_log_filepath,
                            image_translation_enabled: bool, ocr_temperature: float):
        output_path, translation_result_status = "", "실패"
        prs = None
        
        # 작업 로그 파일에 번역 설정 기록
        try:
            with open(task_log_filepath, 'a', encoding='utf-8') as f_log_init:
                f_log_init.write(f"--- 번역 작업 시작 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                f_log_init.write(f"원본 파일: {os.path.basename(file_path)}\n")
                f_log_init.write(f"원본 언어: {src_lang}, 대상 언어: {tgt_lang}, 번역 모델: {model}\n")
                f_log_init.write(f"이미지 번역 활성화: {image_translation_enabled}\n")
                if image_translation_enabled:
                    f_log_init.write(f"  OCR 엔진: {self.current_ocr_engine_type or '미지정'}\n")
                    f_log_init.write(f"  OCR 번역 온도: {ocr_temperature}\n")
                    f_log_init.write(f"  OCR GPU 사용: {self.ocr_use_gpu_var.get() if self.ocr_handler else 'N/A'}\n")
                f_log_init.write(f"총 예상 가중 작업량: {self.total_weighted_work}\n")
                f_log_init.write("-" * 30 + "\n")
        except Exception as e_log_header:
            logger.error(f"작업 로그 파일 헤더 작성 실패: {e_log_header}")


        def report_item_completed_from_handler(slide_info_or_stage: Any, item_type_str: str,
                                               weighted_work_for_item: int, text_snippet_str: str):
            if self.stop_event.is_set(): return

            self.current_weighted_done += weighted_work_for_item
            self.current_weighted_done = min(self.current_weighted_done, self.total_weighted_work) # 초과 방지
            
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self.update_translation_progress,
                                  slide_info_or_stage, item_type_str, # item_type을 문자열로 통일
                                  self.current_weighted_done,
                                  self.total_weighted_work,
                                  text_snippet_str)
        try:
            if self.total_weighted_work == 0: # 번역할 내용이 없는 경우
                logger.warning("번역할 가중 작업량이 없습니다.")
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     self.master.after(0, lambda: messagebox.showinfo("정보", "파일에 번역할 내용이 없습니다."))
                translation_result_status, output_path = "내용 없음", file_path
                with open(task_log_filepath, 'a', encoding='utf-8') as f_log_empty:
                    f_log_empty.write(f"번역할 내용 없음. 원본 파일: {file_path}\n")
            else:
                font_code_for_render = config.UI_LANG_TO_FONT_CODE_MAP.get(tgt_lang, 'en')
                
                self.master.after(0, lambda: self.current_work_label.config(text="1단계: 텍스트/이미지 번역 준비 중..."))
                self.master.update_idletasks()

                temp_dir_for_pptx_handler_main = tempfile.mkdtemp(prefix="pptx_trans_main_")
                temp_pptx_for_chart_translation_path: Optional[str] = None

                prs = Presentation(file_path) # 여기서 시간 걸릴 수 있음 (파일 로드)
                self.master.after(0, lambda: self.current_work_label.config(text="1단계: 텍스트/이미지 번역 처리 중..."))
                
                stage1_success = self.pptx_handler.translate_presentation_stage1(
                    prs, src_lang, tgt_lang,
                    self.translator, 
                    self.ocr_handler if image_translation_enabled else None, # OCR 핸들러 조건부 전달
                    model, self.ollama_service,
                    font_code_for_render, task_log_filepath,
                    report_item_completed_from_handler,
                    self.stop_event,
                    image_translation_enabled, # 이미지 번역 활성화 여부 전달
                    ocr_temperature # OCR 번역 온도 전달
                )

                if self.stop_event.is_set():
                    logger.warning("1단계 번역 중 중지됨 (사용자 요청).")
                    translation_result_status = "부분 성공 (중지)"
                    try:
                        stopped_filename_s1 = os.path.join(temp_dir_for_pptx_handler_main,
                                                           f"{os.path.splitext(os.path.basename(file_path))[0]}_stage1_stopped.pptx")
                        if prs: prs.save(stopped_filename_s1)
                        output_path = stopped_filename_s1
                        logger.info(f"1단계 중단, 부분 저장: {output_path}")
                    except Exception as e_save_stop:
                        logger.error(f"1단계 중단 후 저장 실패: {e_save_stop}")
                        output_path = file_path # 원본 파일 경로로 대체
                elif not stage1_success:
                    logger.error("1단계 번역 실패.")
                    translation_result_status = "실패 (1단계 오류)"
                    output_path = file_path
                else: # 1단계 성공
                    logger.info("번역 작업자: 1단계 완료. 임시 파일 저장 시도.")
                    self.master.after(0, lambda: self.current_work_label.config(text="1단계 완료. 임시 파일 저장 중..."))
                    self.master.update_idletasks()

                    temp_pptx_for_chart_translation_path = os.path.join(
                        temp_dir_for_pptx_handler_main,
                        f"{os.path.splitext(os.path.basename(file_path))[0]}_temp_for_charts.pptx"
                    )
                    if prs: prs.save(temp_pptx_for_chart_translation_path)
                    logger.info(f"1단계 결과 임시 저장: {temp_pptx_for_chart_translation_path}")

                    # 차트 요소 수 확인 (get_file_info 결과 재사용 또는 다시 호출)
                    # 여기서는 간소화를 위해 self.total_weighted_work 계산 시 사용된 _charts 값을 가정.
                    # 정확하려면 prs 객체에서 차트 수를 다시 세어야 함.
                    # 예: num_charts_in_prs = sum(1 for slide in prs.slides for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.CHART)
                    info_for_charts = self.pptx_handler.get_file_info(temp_pptx_for_chart_translation_path) # 임시 파일 기준
                    num_charts_in_prs = info_for_charts.get('chart_elements_count', 0)


                    if num_charts_in_prs > 0 and not self.stop_event.is_set():
                        self.master.after(0, lambda: self.current_work_label.config(text=f"2단계: 차트 번역 준비 중 ({num_charts_in_prs}개)..."))
                        self.master.update_idletasks()
                        logger.info(f"번역 작업자: 2단계 (차트) 시작. 대상 차트 수: {num_charts_in_prs}")
                        
                        safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                        final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                        final_output_dir = os.path.dirname(file_path)
                        final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)
                        
                        self.master.after(0, lambda: self.current_work_label.config(text="2단계: 차트 데이터 압축 해제 중..."))
                        self.master.update_idletasks()

                        output_path_charts = self.chart_xml_handler.translate_charts_in_pptx(
                            pptx_path=temp_pptx_for_chart_translation_path,
                            src_lang_ui_name=src_lang,
                            tgt_lang_ui_name=tgt_lang,
                            model_name=model,
                            output_path=final_pptx_output_path, # 최종 경로 직접 전달
                            progress_callback_item_completed=report_item_completed_from_handler,
                            stop_event=self.stop_event,
                            task_log_filepath=task_log_filepath
                        )
                        # chart_xml_handler.py에서 파일 저장/압축 관련 피드백을 직접 UI에 주기는 어려움.
                        # 여기서는 차트 핸들러 호출 전후로 메시지 업데이트.
                        self.master.after(0, lambda: self.current_work_label.config(text="2단계: 번역된 차트 데이터 압축 중..."))
                        self.master.update_idletasks()


                        if self.stop_event.is_set():
                            logger.warning("2단계 차트 번역 중 또는 완료 직후 중지됨.")
                            translation_result_status = "부분 성공 (중지)"
                            # output_path_charts가 None일 수 있으므로, 1단계 결과물을 최종 경로로 간주
                            output_path = temp_pptx_for_chart_translation_path if not (output_path_charts and os.path.exists(output_path_charts)) else output_path_charts
                        elif output_path_charts and os.path.exists(output_path_charts):
                            logger.info(f"2단계 차트 번역 완료. 최종 파일: {output_path_charts}")
                            translation_result_status = "성공"
                            output_path = output_path_charts
                        else: # 차트 번역 실패 또는 결과 파일 없음
                            logger.error("2단계 차트 번역 실패 또는 결과 파일 없음. 1단계 결과물 사용 시도.")
                            translation_result_status = "실패 (2단계 오류)"
                            # 1단계 결과물을 최종 경로에 복사
                            if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                                try:
                                    shutil.copy2(temp_pptx_for_chart_translation_path, final_pptx_output_path)
                                    output_path = final_pptx_output_path
                                    logger.info(f"차트 번역 실패로 1단계 결과물을 최종 경로에 복사: {output_path}")
                                except Exception as e_copy_fallback:
                                     logger.error(f"차트 번역 실패 후 1단계 결과물 복사 중 오류: {e_copy_fallback}.")
                                     output_path = temp_pptx_for_chart_translation_path # 복사 실패 시 임시 경로 유지
                            else: # 1단계 임시 파일도 없다면 원본 파일 경로
                                output_path = file_path

                    elif self.stop_event.is_set():
                        logger.info("1단계 후 중단되어 차트 번역은 실행되지 않음.")
                        translation_result_status = "부분 성공 (중지)"
                        output_path = temp_pptx_for_chart_translation_path # 1단계 결과
                    else: # 차트가 없는 경우
                        logger.info("번역할 차트가 없습니다. 1단계 결과물을 최종 결과로 사용합니다.")
                        self.master.after(0, lambda: self.current_work_label.config(text="번역할 차트 없음. 최종 파일 저장 중..."))
                        self.master.update_idletasks()
                        
                        safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                        final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                        final_output_dir = os.path.dirname(file_path)
                        final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)
                        try:
                            if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                                shutil.copy2(temp_pptx_for_chart_translation_path, final_pptx_output_path)
                                output_path = final_pptx_output_path
                                translation_result_status = "성공"
                                logger.info(f"차트 없음. 최종 파일 저장: {output_path}")
                            else:
                                logger.error("차트가 없고, 1단계 임시 파일도 찾을 수 없습니다.")
                                translation_result_status = "실패 (파일 오류)"
                                output_path = file_path # 원본 파일
                        except Exception as e_copy_no_chart:
                            logger.error(f"차트 없는 경우 최종 파일 복사 중 오류: {e_copy_no_chart}")
                            translation_result_status = "실패 (파일 복사 오류)"
                            output_path = temp_pptx_for_chart_translation_path if temp_pptx_for_chart_translation_path else file_path

                # 임시 디렉토리 정리
                if 'temp_dir_for_pptx_handler_main' in locals() and temp_dir_for_pptx_handler_main and os.path.exists(temp_dir_for_pptx_handler_main):
                    try:
                        shutil.rmtree(temp_dir_for_pptx_handler_main)
                        logger.debug(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 완료.")
                    except Exception as e_clean_main_dir:
                        logger.warning(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 중 오류: {e_clean_main_dir}")

            # 최종 번역 완료 처리
            if translation_result_status == "성공" and not self.stop_event.is_set():
                 self.current_weighted_done = self.total_weighted_work # 작업량 100%로 채움
                 if hasattr(self, 'master') and self.master.winfo_exists():
                     self.master.after(0, self.update_translation_progress,
                                   "완료", "종료", self.current_weighted_done, self.total_weighted_work, "번역 완료")
                 
                 # 저장된 파일 경로가 유효한지 한번 더 확인
                 if not (output_path and os.path.exists(output_path)):
                     logger.error(f"번역 '성공'으로 기록되었으나, 최종 결과 파일({output_path})을 찾을 수 없습니다.")
                     translation_result_status = "실패 (결과 파일 없음)"
                     # output_path를 원본으로 돌리거나, 사용자에게 알림
                     output_path = file_path # 안전하게 원본 파일로
                 else: # 성공적으로 파일이 존재하면 폴더 열기 질문
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        self.master.after(100, lambda: self._ask_open_folder(output_path))

            elif "실패" in translation_result_status or "오류" in translation_result_status:
                 self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath)
                 if not output_path: output_path = file_path # output_path가 비어있으면 원본으로 설정

            # 중지된 경우도 output_path는 설정되어 있어야 함 (부분 저장된 파일)

        except Exception as e_worker:
            logger.error(f"번역 작업 중 심각한 오류 발생: {e_worker}", exc_info=True)
            translation_result_status = "치명적 오류 발생"
            if not output_path: output_path = file_path # output_path가 비어있으면 원본으로 설정
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_err:
                    f_err.write(f"\n--- 번역 작업 중 심각한 오류 발생 ---\n오류: {e_worker}\n")
                    import traceback
                    traceback.print_exc(file=f_err)
            except Exception as ef_log: logger.error(f"작업 로그 파일에 오류 기록 실패: {ef_log}")
            
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath, str(e_worker))

        finally:
            if hasattr(self, 'master') and self.master.winfo_exists():
                # 히스토리 저장을 위해 필요한 정보들
                history_entry = {
                    "name": os.path.basename(file_path),
                    "src": src_lang,
                    "tgt": tgt_lang,
                    "model": model,
                    "ocr_temp": ocr_temperature if image_translation_enabled else "N/A",
                    "ocr_gpu": self.ocr_use_gpu_var.get() if image_translation_enabled and self.ocr_handler else "N/A",
                    "img_trans_enabled": image_translation_enabled,
                    "status": translation_result_status,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "path": output_path or file_path, # output_path가 None이면 원본 파일 경로
                    "log_file": task_log_filepath
                }
                self.master.after(0, self.translation_finished, history_entry)
            self.translation_thread = None

    def _handle_translation_failure(self, status, original_file, log_file, error_details=""):
        """번역 실패 시 사용자에게 알림을 강화하는 함수"""
        logger.error(f"번역 실패: {status}, 원본: {original_file}, 로그: {log_file}, 상세: {error_details}")
        self.current_work_label.config(text=f"번역 실패: {status}")
        
        error_title = f"번역 작업 실패 ({status})"
        user_message = f"'{os.path.basename(original_file)}' 파일 번역 중 오류가 발생했습니다.\n\n상태: {status}\n"
        if error_details:
            user_message += f"오류 정보: {error_details[:200]}...\n\n" # 너무 길지 않게
        
        user_message += "다음 사항을 확인해 보세요:\n"
        user_message += "- Ollama 서버가 정상적으로 실행 중인지 ('Ollama 확인' 버튼)\n"
        user_message += "- 선택한 번역 모델이 유효한지 (모델 목록 '🔄' 버튼)\n"
        user_message += "- 원본 파일이 손상되지 않았는지\n"
        if "GPU" in status or "CUDA" in status or "메모리" in status or (self.ocr_use_gpu_var.get() and "OCR" in status) : # GPU 관련 오류 의심 시
            user_message += "- 고급 옵션에서 'GPU 사용'을 해제하고 다시 시도해보세요.\n"
        
        user_message += f"\n자세한 내용은 로그 파일에서 확인할 수 있습니다.\n로그 파일: {log_file}"

        # messagebox 대신 사용자 정의 다이얼로그를 사용하거나, 로그 파일을 여는 버튼을 추가할 수 있음
        # 여기서는 messagebox를 사용하되, 로그 파일 경로를 명시
        
        # 로그 폴더 열기 버튼을 포함한 메시지 박스 (간단 버전)
        if messagebox.askyesno(error_title, user_message + "\n\n오류 로그가 저장된 폴더를 여시겠습니까?", icon='error'):
            try:
                utils.open_folder(os.path.dirname(log_file))
            except Exception as e_open_log_dir:
                logger.warning(f"로그 폴더 열기 실패: {e_open_log_dir}")
                messagebox.showinfo("정보", f"로그 폴더를 열 수 없습니다.\n경로: {os.path.dirname(log_file)}")


    def _ask_open_folder(self, path):
        if path and os.path.exists(path):
            user_choice = messagebox.askyesnocancel("번역 완료",
                                           f"번역이 완료되었습니다.\n저장된 파일: {os.path.basename(path)}\n\n결과 파일이 저장된 폴더를 여시겠습니까?",
                                           icon='info', default=messagebox.YES)
            if user_choice is True: # Yes
                utils.open_folder(os.path.dirname(path))
            elif user_choice is False: # No
                pass # 아무것도 안함
            # Cancel은 처리 안함 (기본 동작)

    def update_translation_progress(self, current_location_info: Any, current_task_type: str,
                                    current_total_weighted_done: int, total_weighted_overall: int,
                                    current_text_snippet: str = ""):
        if self.stop_event.is_set(): return
        
        progress = 0
        if total_weighted_overall > 0:
            progress = (current_total_weighted_done / total_weighted_overall) * 100
        elif current_total_weighted_done == 0 and total_weighted_overall == 0 :
             progress = 100

        progress = min(progress, 100)
        progress_text_val = f"{progress:.1f}%"

        # 작업 유형에 따른 메시지 구체화
        task_description = ""
        if current_task_type == "text": task_description = "텍스트 요소 번역 중"
        elif current_task_type == "table_cell": task_description = "표 내부 텍스트 번역 중"
        elif current_task_type == "image_ocr": task_description = "이미지 내 텍스트 인식(OCR) 중"
        elif current_task_type == "image_render": task_description = "번역된 이미지 텍스트 합성 중"
        elif current_task_type == "chart": task_description = "차트 데이터 번역 중"
        elif current_task_type == "chart-error": task_description = "차트 처리 오류 발생"
        elif current_task_type == "종료": task_description = "번역 완료"
        else: task_description = current_task_type # 그대로 표시

        snippet_display = current_text_snippet.replace('\n', ' ').strip()
        if len(snippet_display) > 25: snippet_display = snippet_display[:22] + "..."


        def _update_ui():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            self.progress_bar["value"] = progress
            self.progress_label_var.set(progress_text_val)
            
            location_display_text = str(current_location_info)
            if isinstance(current_location_info, (int, float)):
                location_display_text = f"슬라이드 {int(current_location_info)}"
            elif not current_location_info or str(current_location_info).upper() == "N/A":
                location_display_text = "전체 파일 처리 중" # 차트 등 슬라이드 번호 특정 어려울 때
            elif str(current_location_info).lower() == "완료":
                 location_display_text = "모든 작업 완료"

            self.current_slide_label.config(text=f"현재 위치: {location_display_text}")
            self.current_work_label.config(text=f"현재 작업: {task_description} - '{snippet_display}'")
            
            # self.translated_elements_label, self.remaining_elements_label 삭제됨

        if hasattr(self, 'master') and self.master.winfo_exists():
            self.master.after(0, _update_ui)

    def update_progress_timer(self):
        if self.translation_thread and self.translation_thread.is_alive() and not self.stop_event.is_set():
            # 특별히 타이머로 업데이트할 내용이 없어짐 (진행률은 콜백 기반)
            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(1000, self.update_progress_timer)


    def stop_translation(self):
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("번역 중지 요청 중..."); self.stop_event.set()
            self.stop_button.config(state=tk.DISABLED)
            self.current_work_label.config(text="번역 중지 요청됨...")


    def translation_finished(self, history_entry: Dict[str, Any]):
        """history_entry를 직접 받도록 변경"""
        if not (hasattr(self, 'start_button') and self.start_button.winfo_exists()): return
        self.start_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED)
        
        result_status = history_entry.get("status", "알 수 없음")
        translated_file_path = history_entry.get("path")
        original_file = history_entry.get("name") # 히스토리에는 파일명만 있으므로, 전체 경로는 file_path_var에서 가져와야 할 수도.
                                                  # _translation_worker에서 file_path를 사용하므로, history_entry에 original_full_path 추가 고려.
                                                  # 현재는 original_file (basename) 사용.

        current_progress_val = self.progress_bar["value"]
        final_progress_text = f"{current_progress_val:.1f}%"

        if result_status == "성공":
            final_progress_text = "100%"
            self.progress_bar["value"] = 100
            self.current_work_label.config(text=f"번역 완료: {os.path.basename(translated_file_path) if translated_file_path else '파일 없음'}")
        elif "중지" in result_status:
            final_progress_text = f"{current_progress_val:.1f}% (중지됨)"
            self.current_work_label.config(text="번역 중지됨.")
        elif result_status == "내용 없음":
            final_progress_text = "100% (내용 없음)"
            self.progress_bar["value"] = 100
            self.current_work_label.config(text="번역할 내용 없음.")
        else: # 실패 또는 오류
            final_progress_text = f"{current_progress_val:.1f}% ({result_status})"
            # _handle_translation_failure 에서 current_work_label을 이미 설정했을 것이므로 여기서는 변경 안 함

        self.progress_label_var.set(final_progress_text)
        
        if translated_file_path and os.path.exists(translated_file_path) and result_status == "성공":
            self.translated_file_path_var.set(translated_file_path); self.open_folder_button.config(state=tk.NORMAL)
        else:
            self.translated_file_path_var.set("번역 실패 또는 파일 없음"); self.open_folder_button.config(state=tk.DISABLED)
            if not (translated_file_path and os.path.exists(translated_file_path)) and result_status == "성공":
                 logger.warning(f"번역은 '성공'으로 기록되었으나, 결과 파일 경로가 유효하지 않음: {translated_file_path}")
        
        # 히스토리 추가
        self._add_history_entry(history_entry)
        
        task_log_filepath = history_entry.get("log_file")
        if task_log_filepath and os.path.exists(os.path.dirname(task_log_filepath)):
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    f_task_log.write(f"\n--- 번역 작업 최종 상태 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                    f_task_log.write(f"최종 상태: {result_status}\n")
                    if self.file_path_var.get(): f_task_log.write(f"원본 파일 (UI 경로): {self.file_path_var.get()}\n")
                    if translated_file_path and os.path.exists(translated_file_path): f_task_log.write(f"번역된 파일: {translated_file_path}\n")
                    elapsed_time_for_log = (time.time() - self.start_time) if self.start_time else 0
                    m, s = divmod(elapsed_time_for_log, 60)
                    f_task_log.write(f"총 소요 시간 (내부 기록용): {int(m):02d}분 {s:05.2f}초\n")
                    f_task_log.write("-" * 30 + "\n")
            except Exception as e_log_finish: logger.error(f"작업 로그 파일에 최종 상태 기록 실패: {e_log_finish}")
        
        self.start_time = None # 타이머 리셋
        if result_status != "성공" and "중지" not in result_status and result_status != "내용 없음":
             # _handle_translation_failure에서 이미 메시지 처리했을 가능성 높음
             pass
        elif result_status == "성공":
             pass # _ask_open_folder에서 메시지 처리


    def open_translated_folder(self):
        path = self.translated_file_path_var.get()
        if path and os.path.exists(path): utils.open_folder(os.path.dirname(path))
        elif path: messagebox.showwarning("폴더 열기 실패", f"경로를 찾을 수 없습니다: {path}")
        else: messagebox.showinfo("정보", "번역된 파일 경로가 없습니다.")

    def on_history_double_click(self, event):
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()): return
        item_id = self.history_tree.identify_row(event.y)
        if item_id:
            item_values = self.history_tree.item(item_id, "values")
            # 히스토리 컬럼 순서에 맞춰 경로 인덱스 확인 (path는 마지막 컬럼)
            if item_values and len(item_values) == len(self.history_tree["columns"]):
                file_path_to_open = item_values[self.history_tree["columns"].index("path")]
                status_val = item_values[self.history_tree["columns"].index("status")]
                
                if file_path_to_open and os.path.exists(file_path_to_open) and "성공" in status_val : # 성공한 경우만 파일 열기 시도
                    if messagebox.askyesno("파일 열기", f"번역된 파일 '{os.path.basename(file_path_to_open)}'을(를) 여시겠습니까?"):
                        try:
                            if platform.system() == "Windows": os.startfile(file_path_to_open)
                            elif platform.system() == "Darwin": subprocess.Popen(["open", file_path_to_open])
                            else: subprocess.Popen(["xdg-open", file_path_to_open])
                        except Exception as e: logger.error(f"히스토리 파일 열기 실패: {e}", exc_info=True)
                elif "성공" not in status_val and file_path_to_open :
                     log_file_path_from_history = ""
                     # 실제 히스토리 데이터에서 로그 파일 경로 찾기
                     for entry_data in self.translation_history_data:
                         if entry_data.get("path") == file_path_to_open and entry_data.get("time") == item_values[self.history_tree["columns"].index("time")]:
                             log_file_path_from_history = entry_data.get("log_file", "")
                             break
                     
                     if log_file_path_from_history and os.path.exists(log_file_path_from_history):
                         if messagebox.askyesno("로그 파일 열기", f"번역 결과가 '{status_val}'입니다.\n관련 로그 파일 '{os.path.basename(log_file_path_from_history)}'을(를) 여시겠습니까?"):
                             try: utils.open_folder(os.path.dirname(log_file_path_from_history)) # 폴더 열기
                             except Exception as e: logger.error(f"히스토리 로그 폴더 열기 실패: {e}")
                     else:
                          messagebox.showwarning("정보", f"번역 결과가 '{status_val}'입니다.\n(관련 로그 파일 정보 없음)")

                elif file_path_to_open: messagebox.showwarning("파일 없음", f"파일을 찾을 수 없습니다: {file_path_to_open}")


class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__(); self.text_widget = text_widget
    def emit(self, record):
        if not (self.text_widget and self.text_widget.winfo_exists()): return
        msg = self.format(record)
        def append_message():
            if not (self.text_widget and self.text_widget.winfo_exists()): return
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
        # UI가 파괴되는 시점에 after 콜백이 실행되려 하면 TclError 발생 가능성 있음
        try:
            if self.text_widget.winfo_exists():
                self.text_widget.after(0, append_message)
        except tk.TclError:
            pass # 위젯이 이미 파괴된 경우 무시


if __name__ == "__main__":
    # 폴더 생성은 Application 클래스 __init__으로 이동
    
    if debug_mode: logger.info("디버그 모드로 실행 중입니다.")
    else: logger.info("일반 모드로 실행 중입니다.")
    
    if not os.path.exists(config.FONTS_DIR):
        logger.critical(f"필수 폰트 디렉토리를 찾을 수 없습니다: {config.FONTS_DIR}")
    else: logger.info(f"폰트 디렉토리 확인: {config.FONTS_DIR}")
    
    if not os.path.exists(config.ASSETS_DIR):
        logger.warning(f"에셋 디렉토리를 찾을 수 없습니다: {config.ASSETS_DIR}")
    else: logger.info(f"에셋 디렉토리 확인: {config.ASSETS_DIR}")

    root = tk.Tk()
    app = Application(master=root)
    root.geometry("1024x768") # 창 크기 약간 증가
    root.update_idletasks()
    min_width = root.winfo_reqwidth()
    min_height = root.winfo_reqheight()
    root.minsize(min_width + 20, min_height + 20)
    try:
        root.mainloop()
    except KeyboardInterrupt:
        logger.info("Ctrl+C로 애플리케이션 종료 중...")
    finally:
        logger.info(f"--- {APP_NAME} 종료됨 ---")