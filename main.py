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
import tempfile # <--- tempfile import 추가
import shutil
from typing import Optional # <--- 이 줄을 추가하거나, 이미 from typing import ... 가 있다면 Optional을 추가하세요.


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
# debug_mode는 sys.argv에서 가져오므로 main.py에서 처리
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
# BASE_DIR_MAIN은 main.py의 위치를 기준으로 하는 것이 적절할 수 있음
BASE_DIR_MAIN = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = config.ASSETS_DIR
FONTS_DIR = config.FONTS_DIR # main.py에서도 폰트 디렉토리 확인에 사용
LOGS_DIR = config.LOGS_DIR

# 로그 폴더 생성 확인 (config의 LOGS_DIR 사용)
if not os.path.exists(LOGS_DIR):
    try: os.makedirs(LOGS_DIR)
    except Exception as e: print(f"로그 폴더 생성 실패: {LOGS_DIR}, 오류: {e}")

logger = logging.getLogger(__name__)

# --- 전역 변수 및 설정 (config.py에서 가져옴) ---
APP_NAME = config.APP_NAME
DEFAULT_MODEL = config.DEFAULT_OLLAMA_MODEL
SUPPORTED_LANGUAGES = config.SUPPORTED_LANGUAGES
TRANSLATION_HISTORY = []


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

        self.ollama_service = OllamaService() # 내부에서 config 사용
        self.translator = OllamaTranslator() # 내부에서 config 사용
        self.pptx_handler = PptxHandler()     # 내부에서 config 사용
        self.chart_xml_handler = ChartXmlHandler(self.translator, self.ollama_service) # 내부에서 config 사용
        
        self.ocr_handler = None 
        self.current_ocr_engine_type = None 

        self.translation_thread = None
        self.model_download_thread = None
        self.stop_event = threading.Event()
        self.logo_image_tk_bottom = None
        self.start_time = None

        self.current_file_slide_count = 0
        self.current_file_text_elements_count = 0
        self.current_file_total_text_chars = 0
        self.current_file_image_elements_count = 0
        self.current_file_chart_elements_count = 0
        self.total_weighted_work = 0
        self.current_weighted_done = 0

        self.create_widgets()
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
        if hasattr(self, 'master') and self.master.winfo_exists():
            if messagebox.askokcancel("종료 확인", f"{APP_NAME}을(를) 종료하시겠습니까?"):
                 logger.info("모든 정리 작업 완료. 애플리케이션을 종료합니다.")
                 self.master.destroy()
            else:
                 logger.info("애플리케이션 종료 취소됨.")
                 if self.stop_event.is_set(): self.stop_event.clear() 
                 return 
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

            left_panel = ttk.Frame(main_paned_window, padding=10)
            main_paned_window.add(left_panel, weight=2) 

            right_panel = ttk.Frame(main_paned_window, padding=10)
            main_paned_window.add(right_panel, weight=1) 

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

            self.ollama_status_label = ttk.Label(server_status_frame, text="Ollama 설치: 미확인")
            self.ollama_status_label.grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
            self.ollama_running_label = ttk.Label(server_status_frame, text="Ollama 실행: 미확인")
            self.ollama_running_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
            self.ollama_port_label = ttk.Label(server_status_frame, text="Ollama 포트: -")
            self.ollama_port_label.grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
            ttk.Button(server_status_frame, text="Ollama 확인", command=self.check_ollama_status_manual).grid(row=0, column=3, padx=5, pady=2, sticky=tk.E)
            
            self.ocr_status_label = ttk.Label(server_status_frame, text="OCR 상태: 미확인") 
            self.ocr_status_label.grid(row=1, column=0, columnspan=4, padx=5, pady=2, sticky=tk.W) 

            file_progress_outer_frame = ttk.Frame(left_panel)
            file_progress_outer_frame.pack(padx=5, pady=5, fill=tk.X)

            file_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="파일 정보", padding=5)
            file_info_frame.pack(side=tk.LEFT, padx=(0,5), fill=tk.BOTH, expand=True)
            self.file_name_label = ttk.Label(file_info_frame, text="파일 이름: ")
            self.file_name_label.pack(anchor=tk.W, pady=1)
            self.slide_count_label = ttk.Label(file_info_frame, text="슬라이드 수: ")
            self.slide_count_label.pack(anchor=tk.W, pady=1)
            self.text_elements_label = ttk.Label(file_info_frame, text="텍스트 요소 수: ")
            self.text_elements_label.pack(anchor=tk.W, pady=1)
            self.total_text_char_label = ttk.Label(file_info_frame, text="총 텍스트 글자 수: ")
            self.total_text_char_label.pack(anchor=tk.W, pady=1)
            self.image_elements_label = ttk.Label(file_info_frame, text="총 이미지 수: ") 
            self.image_elements_label.pack(anchor=tk.W, pady=1)
            self.chart_elements_label = ttk.Label(file_info_frame, text="차트 수: ")
            self.chart_elements_label.pack(anchor=tk.W, pady=1)
            self.total_weighted_work_label = ttk.Label(file_info_frame, text="총 예상 작업량: ")
            self.total_weighted_work_label.pack(anchor=tk.W, pady=1)

            progress_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="진행 상황", padding=5)
            progress_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            self.current_slide_label = ttk.Label(progress_info_frame, text="현재 슬라이드: -")
            self.current_slide_label.pack(anchor=tk.W, pady=1)
            self.current_work_label = ttk.Label(progress_info_frame, text="현재 작업: -")
            self.current_work_label.pack(anchor=tk.W, pady=1)
            self.translated_elements_label = ttk.Label(progress_info_frame, text="완료된 작업량: 0")
            self.translated_elements_label.pack(anchor=tk.W, pady=1)
            self.remaining_elements_label = ttk.Label(progress_info_frame, text="남은 작업량: 0")
            self.remaining_elements_label.pack(anchor=tk.W, pady=1)

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
            
            ttk.Label(translation_options_frame, text="번역 모델:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
            self.model_var = tk.StringVar(value=DEFAULT_MODEL)
            self.model_combo = ttk.Combobox(translation_options_frame, textvariable=self.model_var, state="disabled") 
            self.model_combo.grid(row=1, column=1, columnspan=4, padx=5, pady=5, sticky=tk.EW)

            action_buttons_frame = ttk.Frame(left_panel, padding=(0,5,0,0)) 
            action_buttons_frame.pack(padx=5, pady=10, fill=tk.X)
            self.start_button = ttk.Button(action_buttons_frame, text="번역 시작", command=self.start_translation, style="Accent.TButton")
            self.start_button.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X, ipady=5) 
            self.stop_button = ttk.Button(action_buttons_frame, text="번역 중지", command=self.stop_translation, state=tk.DISABLED)
            self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=5)
            try:
                self.style.configure("Accent.TButton", font=('Helvetica', 10, 'bold'), foreground="white", background="#0078D7") 
            except tk.TclError:
                logger.warning("Accent.TButton 스타일 적용 실패. 시스템 기본 버튼 스타일이 사용됩니다.")

            progress_bar_frame = ttk.Frame(left_panel)
            progress_bar_frame.pack(padx=5, pady=5, fill=tk.X)
            self.progress_bar = ttk.Progressbar(progress_bar_frame, orient="horizontal", length=300, mode="determinate")
            self.progress_bar.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
            self.progress_label_var = tk.StringVar(value="0%") # 시간 관련 텍스트 제거
            ttk.Label(progress_bar_frame, textvariable=self.progress_label_var).pack(side=tk.LEFT)

            self.translated_file_path_var = tk.StringVar()
            translated_file_frame = ttk.LabelFrame(left_panel, text="번역 완료 파일", padding=5)
            translated_file_frame.pack(padx=5, pady=5, fill=tk.X)
            self.translated_file_entry = ttk.Entry(translated_file_frame, textvariable=self.translated_file_path_var, state="readonly", width=60) 
            self.translated_file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
            self.open_folder_button = ttk.Button(translated_file_frame, text="폴더 열기", command=self.open_translated_folder, state=tk.DISABLED)
            self.open_folder_button.pack(side=tk.LEFT)

            right_panel_notebook = ttk.Notebook(right_panel)
            right_panel_notebook.pack(fill=tk.BOTH, expand=True)

            log_tab_frame = ttk.Frame(right_panel_notebook, padding=5)
            right_panel_notebook.add(log_tab_frame, text="실행 로그")
            self.log_text = tk.Text(log_tab_frame, height=15, state=tk.DISABLED, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1, font=("TkFixedFont", 9))
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
            self.history_tree = ttk.Treeview(history_tab_frame, columns=("name", "src", "tgt", "result", "time", "path"), show="headings", height=10)
            self.history_tree.heading("name", text="문서 이름")
            self.history_tree.column("name", width=150, anchor=tk.W, stretch=tk.YES)
            self.history_tree.heading("src", text="원본언어")
            self.history_tree.column("src", width=70, anchor=tk.CENTER)
            self.history_tree.heading("tgt", text="번역언어")
            self.history_tree.column("tgt", width=70, anchor=tk.CENTER)
            self.history_tree.heading("result", text="결과")
            self.history_tree.column("result", width=80, anchor=tk.CENTER)
            self.history_tree.heading("time", text="번역일시")
            self.history_tree.column("time", width=120, anchor=tk.CENTER)
            self.history_tree.heading("path", text="경로") 
            self.history_tree.column("path", width=0, stretch=tk.NO) 

            hist_scrollbar_y = ttk.Scrollbar(history_tab_frame, orient="vertical", command=self.history_tree.yview)
            hist_scrollbar_x = ttk.Scrollbar(history_tab_frame, orient="horizontal", command=self.history_tree.xview)
            self.history_tree.configure(yscrollcommand=hist_scrollbar_y.set, xscrollcommand=hist_scrollbar_x.set)
            hist_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            hist_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            self.history_tree.pack(fill=tk.BOTH, expand=True)
            self.history_tree.bind("<Double-1>", self.on_history_double_click)

            logo_path_bottom = os.path.join(ASSETS_DIR, "LINEstudio2.png")
            if os.path.exists(logo_path_bottom):
                try:
                    pil_temp_for_size = Image.open(logo_path_bottom)
                    original_width, original_height = pil_temp_for_size.size
                    pil_temp_for_size.close()
                    target_height_bottom = 20
                    subsample_factor = 1
                    if original_height > target_height_bottom and target_height_bottom > 0: 
                        subsample_factor = max(1, int(original_height / target_height_bottom))
                    elif original_height > 0 : subsample_factor = 1
                    else: subsample_factor = 6 
                    
                    if original_width > 0 and original_height > 0:
                        target_width_approx = int(target_height_bottom * (original_width / original_height))
                        if target_width_approx > 0 :
                            subsample_factor_w = max(1, int(original_width / target_width_approx))
                            subsample_factor = max(subsample_factor, subsample_factor_w) 
                    
                    if subsample_factor < 1: subsample_factor = 1

                    temp_logo_image_bottom = tk.PhotoImage(file=logo_path_bottom, master=self.master)
                    self.logo_image_tk_bottom = temp_logo_image_bottom.subsample(subsample_factor, subsample_factor)
                    logo_label_bottom = ttk.Label(bottom_frame, image=self.logo_image_tk_bottom)
                    logo_label_bottom.pack(side=tk.RIGHT, padx=10, pady=2)
                except Exception as e_general_bottom: 
                    logger.warning(f"하단 로고 로드 중 예외: {e_general_bottom}", exc_info=True)
            else:
                logger.warning(f"하단 로고 파일({logo_path_bottom})을 찾을 수 없습니다.")

    def update_ocr_status_display(self):
        selected_ui_lang = self.src_lang_var.get()
        engine_type = "EasyOCR" if selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS else "PaddleOCR"
        
        if self.ocr_handler and self.current_ocr_engine_type == engine_type.lower():
            current_handler_lang = ""
            if self.current_ocr_engine_type == "paddle":
                current_handler_lang = self.ocr_handler.current_lang_codes
            elif self.current_ocr_engine_type == "easyocr" and self.ocr_handler.current_lang_codes:
                current_handler_lang = ", ".join(self.ocr_handler.current_lang_codes)
            self.ocr_status_label.config(text=f"{engine_type} OCR: 준비됨 ({current_handler_lang})")
        else:
            self.ocr_status_label.config(text=f"{engine_type} OCR: ({selected_ui_lang}) 사용 예정 (미확인)")

    def on_source_language_change(self, event=None):
        selected_ui_lang = self.src_lang_var.get()
        logger.info(f"원본 언어 변경됨: {selected_ui_lang}.")
        self.update_ocr_status_display() 
        if self.file_path_var.get():
            self.load_file_info(self.file_path_var.get())

    def browse_file(self): 
        file_path = filedialog.askopenfilename(title="파워포인트 파일 선택", filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*")))
        if file_path:
            self.file_path_var.set(file_path); logger.info(f"파일 선택됨: {file_path}")
            self.load_file_info(file_path); self.translated_file_path_var.set("")
            self.open_folder_button.config(state=tk.DISABLED)

    def load_file_info(self, file_path):
        info = {"slide_count": 0, "text_elements_count": 0, "total_text_char_count":0, "image_elements_count": 0, "chart_elements_count": 0}
        try:
            logger.debug(f"파일 정보 분석 중: {file_path}"); file_name = os.path.basename(file_path)
            info = self.pptx_handler.get_file_info(file_path)
            
            self.current_file_slide_count = info.get('slide_count', 0)
            self.current_file_text_elements_count = info.get('text_elements_count', 0)
            self.current_file_total_text_chars = info.get('total_text_char_count', 0)
            self.current_file_image_elements_count = info.get('image_elements_count', 0)
            self.current_file_chart_elements_count = info.get('chart_elements_count', 0)

            self.file_name_label.config(text=f"파일 이름: {file_name}")
            self.slide_count_label.config(text=f"슬라이드 수: {self.current_file_slide_count}")
            self.text_elements_label.config(text=f"텍스트 요소 수: {self.current_file_text_elements_count}")
            self.total_text_char_label.config(text=f"총 텍스트 글자 수: {self.current_file_total_text_chars}")
            self.image_elements_label.config(text=f"총 이미지 수: {self.current_file_image_elements_count}")
            self.chart_elements_label.config(text=f"차트 수: {self.current_file_chart_elements_count}")

            self.total_weighted_work = (self.current_file_total_text_chars * config.WEIGHT_TEXT_CHAR) + \
                                       (self.current_file_image_elements_count * config.WEIGHT_IMAGE) + \
                                       (self.current_file_chart_elements_count * config.WEIGHT_CHART)
            
            self.total_weighted_work_label.config(text=f"총 예상 작업량: {self.total_weighted_work}")
            self.remaining_elements_label.config(text=f"남은 작업량: {self.total_weighted_work}")
            self.translated_elements_label.config(text="완료된 작업량: 0")
            logger.info("파일 정보 분석 및 가중 작업량 계산 완료.")
        except Exception as e: 
            logger.error(f"파일 정보 분석 오류: {e}", exc_info=True)
            self.file_name_label.config(text="파일 이름: -")
            self.slide_count_label.config(text="슬라이드 수: -")
            self.text_elements_label.config(text="텍스트 요소 수: -")
            self.total_text_char_label.config(text="총 텍스트 글자 수: -")
            self.image_elements_label.config(text="총 이미지 수: -")
            self.chart_elements_label.config(text="차트 수: -")
            self.total_weighted_work_label.config(text="총 예상 작업량: -")

    def check_ollama_status_manual(self, initial_check=False): 
        logger.info("Ollama 상태 확인 중...")
        ollama_installed = self.ollama_service.is_installed()
        self.ollama_status_label.config(text=f"Ollama 설치: {'설치됨' if ollama_installed else '미설치'}")
        if not ollama_installed:
            logger.warning("Ollama가 설치되어 있지 않습니다.")
            if not initial_check and messagebox.askyesno("Ollama 설치 필요", "Ollama가 설치되어 있지 않습니다. Ollama 다운로드 페이지로 이동하시겠습니까?"): webbrowser.open("https://ollama.com/download")
            self.ollama_running_label.config(text="Ollama 실행: 미설치"); self.ollama_port_label.config(text="Ollama 포트: -")
            self.model_combo.config(values=[], state="disabled"); self.model_var.set(""); return
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
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(2000, lambda: self.check_ollama_status_manual(initial_check=initial_check))
                else:
                    logger.error("Ollama 자동 시작 실패. 수동으로 실행해주세요.")
                    if not initial_check: messagebox.showwarning("Ollama 시작 실패", "Ollama를 자동으로 시작할 수 없습니다. 수동으로 실행 후 'Ollama 확인'을 눌러주세요.")

    def load_ollama_models(self): 
        logger.debug("Ollama 모델 목록 로드 중...")
        models = self.ollama_service.get_text_models()
        if models:
            self.model_combo.config(values=models, state="readonly")
            if DEFAULT_MODEL in models: self.model_var.set(DEFAULT_MODEL)
            elif models: self.model_var.set(models[0])
            logger.info(f"사용 가능 Ollama 모델: {models}")
            if DEFAULT_MODEL not in models: self.download_default_model_if_needed(initial_check_from_ollama=True)
        else:
            self.model_combo.config(values=[], state="disabled"); self.model_var.set("")
            logger.warning("Ollama에 로드된 모델이 없습니다.")
            self.download_default_model_if_needed(initial_check_from_ollama=True)

    def download_default_model_if_needed(self, initial_check_from_ollama=False): 
        current_models = self.ollama_service.get_text_models()
        if DEFAULT_MODEL not in current_models:
            logger.warning(f"기본 모델 ({DEFAULT_MODEL})이 설치되어 있지 않습니다.")
            if initial_check_from_ollama or messagebox.askyesno("기본 모델 다운로드", f"기본 번역 모델 '{DEFAULT_MODEL}'이(가) 없습니다. 지금 다운로드하시겠습니까? (시간 소요)"):
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드 시작...")
                self.start_button.config(state=tk.DISABLED); self.progress_bar["value"] = 0
                self.progress_label_var.set(f"모델 다운로드 시작: {DEFAULT_MODEL}") # 시간 표시 제거
                if self.model_download_thread and self.model_download_thread.is_alive():
                    logger.warning("이미 모델 다운로드 스레드가 실행 중입니다."); return
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
        else: 
            logger.error(f"'{model_name}' 모델 다운로드 실패.")
            if not self.stop_event.is_set(): messagebox.showerror("모델 다운로드 실패", f"'{model_name}' 모델 다운로드에 실패했습니다. Ollama 로그를 확인해주세요.")
        if not (self.translation_thread and self.translation_thread.is_alive()):
            self.start_button.config(state=tk.NORMAL)
            self.progress_bar["value"] = 0; self.progress_label_var.set("0%") # 시간 표시 제거

    def update_model_download_progress(self, status_text, completed_bytes, total_bytes, is_error=False): 
        if self.stop_event.is_set(): return
        if total_bytes > 0: percent = (completed_bytes / total_bytes) * 100; progress_str = f"{percent:.1f}%"
        else: percent = 0; progress_str = status_text # 다운로드 시작 전 상태 메시지
        
        def _update():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            if not is_error: 
                self.progress_bar["value"] = percent
                # 모델 다운로드 시에는 상세 상태 텍스트 유지 (시간 정보는 없음)
                self.progress_label_var.set(f"모델 다운로드: {progress_str} ({status_text})") 
            else:
                self.progress_label_var.set(f"모델 다운로드 오류: {status_text}")
            logger.log(logging.ERROR if is_error else logging.DEBUG, f"모델 다운로드 진행: {status_text} ({completed_bytes}/{total_bytes})")
        if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, _update)


    def check_ocr_engine_status(self, is_called_from_start_translation=False):
        selected_ui_lang = self.src_lang_var.get()
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name = "EasyOCR" if use_easyocr else "PaddleOCR"
        
        ocr_lang_code = None
        if use_easyocr:
            ocr_lang_code = config.UI_LANG_TO_EASYOCR_CODE_MAP.get(selected_ui_lang)
        else:
            ocr_lang_code = config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(selected_ui_lang, config.DEFAULT_PADDLE_OCR_LANG)

        if not ocr_lang_code:
            msg = f"{engine_name}: 언어 '{selected_ui_lang}'에 대한 코드 없음."
            self.ocr_status_label.config(text=msg)
            logger.error(msg)
            if is_called_from_start_translation: messagebox.showerror("OCR 오류", msg)
            return False

        needs_reinit = False
        if not self.ocr_handler: needs_reinit = True
        elif self.current_ocr_engine_type != engine_name.lower(): needs_reinit = True
        elif engine_name == "PaddleOCR" and self.ocr_handler.current_lang_codes != ocr_lang_code: needs_reinit = True
        elif engine_name == "EasyOCR" and (not self.ocr_handler.current_lang_codes or ocr_lang_code not in self.ocr_handler.current_lang_codes):
            needs_reinit = True
        
        if needs_reinit:
            self._destroy_current_ocr_handler() 
            logger.info(f"{engine_name} 핸들러 (재)초기화 시도 (언어: {ocr_lang_code}).")
            try:
                if use_easyocr:
                    if not utils.check_easyocr():
                        self.ocr_status_label.config(text=f"{engine_name}: 미설치")
                        if messagebox.askyesno(f"{engine_name} 설치 필요", f"{engine_name}이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            if utils.install_easyocr(): messagebox.showinfo(f"{engine_name} 설치 완료", f"{engine_name}이(가) 설치되었습니다. 다시 시도해주세요.")
                            else: messagebox.showerror(f"{engine_name} 설치 실패", f"{engine_name} 설치에 실패했습니다.")
                        return False 
                    self.ocr_handler = EasyOcrHandler(lang_codes_list=[ocr_lang_code], debug_enabled=debug_mode, use_gpu=False) 
                    self.current_ocr_engine_type = "easyocr"
                else: 
                    if not utils.check_paddleocr():
                        self.ocr_status_label.config(text=f"{engine_name}: 미설치")
                        if messagebox.askyesno(f"{engine_name} 설치 필요", f"{engine_name}이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            if utils.install_paddleocr(): messagebox.showinfo(f"{engine_name} 설치 완료", f"{engine_name}이(가) 설치되었습니다. 다시 시도해주세요.")
                            else: messagebox.showerror(f"{engine_name} 설치 실패", f"{engine_name} 설치에 실패했습니다.")
                        return False
                    self.ocr_handler = PaddleOcrHandler(lang_code=ocr_lang_code, debug_enabled=debug_mode, use_gpu=False)
                    self.current_ocr_engine_type = "paddle"
                
                logger.info(f"{engine_name} 핸들러 초기화 성공 (언어: {ocr_lang_code}).")

            except RuntimeError as e: 
                logger.error(f"{engine_name} 핸들러 초기화 실패: {e}", exc_info=True)
                self.ocr_status_label.config(text=f"{engine_name}: 초기화 실패 ({ocr_lang_code})")
                if is_called_from_start_translation: messagebox.showerror(f"{engine_name} 오류", f"{engine_name} 초기화 중 오류:\n{e}")
                self._destroy_current_ocr_handler() 
                return False
            except Exception as e_other: 
                 logger.error(f"{engine_name} 핸들러 생성 중 예기치 않은 오류: {e_other}", exc_info=True)
                 self.ocr_status_label.config(text=f"{engine_name}: 알 수 없는 오류")
                 if is_called_from_start_translation: messagebox.showerror(f"{engine_name} 오류", f"{engine_name} 처리 중 예기치 않은 오류:\n{e_other}")
                 self._destroy_current_ocr_handler()
                 return False

        if self.ocr_handler and self.ocr_handler.ocr_engine:
            current_handler_lang_display = ocr_lang_code
            if use_easyocr and isinstance(self.ocr_handler.current_lang_codes, list):
                current_handler_lang_display = ", ".join(self.ocr_handler.current_lang_codes)

            self.ocr_status_label.config(text=f"{engine_name} OCR: 준비됨 ({current_handler_lang_display})")
            return True
        else:
            self.ocr_status_label.config(text=f"{engine_name} OCR: 준비 안됨 ({selected_ui_lang})")
            if is_called_from_start_translation and not needs_reinit : 
                 messagebox.showwarning("OCR 오류", f"{engine_name} OCR 엔진을 사용할 수 없습니다.")
            return False


    def swap_languages(self): 
        src, tgt = self.src_lang_var.get(), self.tgt_lang_var.get()
        self.src_lang_var.set(tgt); self.tgt_lang_var.set(src)
        logger.info(f"언어 스왑: {tgt} <-> {src}")
        self.on_source_language_change()

    def start_translation(self):
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "번역할 유효한 파워포인트 파일을 선택해주세요.")
            return

        if not self.check_ocr_engine_status(is_called_from_start_translation=True):
            if not messagebox.askyesno("OCR 준비 실패",
                                     "OCR 기능이 준비되지 않았거나 사용할 수 없습니다.\n"
                                     "이 경우 이미지 안의 글자는 번역되지 않습니다.\n"
                                     "계속 진행하시겠습니까? (텍스트/차트만 번역)"):
                logger.warning("OCR 준비 실패로 사용자가 번역을 취소했습니다.")
                return
            logger.warning("OCR 핸들러 준비 실패. 이미지 번역 없이 진행합니다.")
        
        src_lang, tgt_lang, model = self.src_lang_var.get(), self.tgt_lang_var.get(), self.model_var.get()
        if not model: 
            messagebox.showerror("모델 오류", "번역 모델을 선택해주세요."); self.check_ollama_status_manual(); return
        if src_lang == tgt_lang: 
            messagebox.showwarning("언어 동일", "원본 언어와 번역 언어가 동일합니다."); return
        ollama_running, _ = self.ollama_service.is_running()
        if not ollama_running: 
            messagebox.showerror("Ollama 미실행", "Ollama 서버가 실행 중이지 않습니다."); self.check_ollama_status_manual(); return

        if self.total_weighted_work <= 0:
            # 파일 선택 시 load_file_info가 호출되어 total_weighted_work가 계산됨.
            # 여기서 0이라는 것은 번역할 내용이 없거나, 파일 로드에 문제가 있었을 수 있음.
            # 다시 한번 load_file_info를 호출하여 확인
            logger.info("총 예상 작업량이 0입니다. 파일 정보를 다시 로드하여 확인합니다.")
            self.load_file_info(file_path)
            if self.total_weighted_work <= 0:
                messagebox.showinfo("정보", "번역할 내용이 없거나 작업량을 계산할 수 없습니다.")
                logger.warning("재확인 후에도 총 예상 작업량이 0 이하입니다. 번역을 시작하지 않습니다.")
                return


        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_filename = os.path.basename(file_path)
        safe_original_filename_part = "".join(c if c.isalnum() or c in ['.', '_'] else '_' for c in os.path.splitext(original_filename)[0])
        task_log_filename = f"translation_{timestamp}_{safe_original_filename_part}.log"
        task_log_filepath = os.path.join(LOGS_DIR, task_log_filename)
        
        logger.info(f"번역 시작: '{original_filename}' ({src_lang} -> {tgt_lang}) using {model}. OCR 엔진: {self.current_ocr_engine_type if self.ocr_handler else '없음'}")
        self.start_button.config(state=tk.DISABLED); self.stop_button.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0; self.progress_label_var.set("0%") # 시간 표시 제거
        self.translated_file_path_var.set(""); self.open_folder_button.config(state=tk.DISABLED)
        
        self.current_weighted_done = 0 
        self.translated_elements_label.config(text=f"완료된 작업량: {self.current_weighted_done}")
        self.remaining_elements_label.config(text=f"남은 작업량: {self.total_weighted_work}")

        self.stop_event.clear()
        
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("이미 번역 스레드가 실행 중입니다."); messagebox.showwarning("번역 중복", "이미 다른 번역 작업이 진행 중입니다.")
            self.start_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED); return

        self.translation_thread = threading.Thread(target=self._translation_worker, 
                                                   args=(file_path, src_lang, tgt_lang, model, task_log_filepath),
                                                   daemon=True)
        self.start_time = time.time()
        self.translation_thread.start()
        self.update_progress_timer()


    def _translation_worker(self, file_path, src_lang, tgt_lang, model, task_log_filepath):
        output_path, translation_result_status = "", "실패"
        prs = None 

        def report_item_completed_from_handler(slide_info, item_type, weighted_work_for_item, text_snippet):
            if self.stop_event.is_set(): return

            self.current_weighted_done += weighted_work_for_item
            self.current_weighted_done = min(self.current_weighted_done, self.total_weighted_work)
            
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self.update_translation_progress,
                                  slide_info, item_type,
                                  self.current_weighted_done,
                                  self.total_weighted_work,
                                  text_snippet)

        try:
            if self.total_weighted_work == 0:
                logger.warning("번역할 가중 작업량이 없습니다.")
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     self.master.after(0, lambda: messagebox.showinfo("정보", "파일에 번역할 내용이 없습니다."))
                translation_result_status, output_path = "내용 없음", file_path 
                with open(task_log_filepath, 'a', encoding='utf-8') as f_log_empty:
                    f_log_empty.write(f"번역할 내용 없음. 원본 파일: {file_path}\n")
            else:
                font_code_for_render = config.UI_LANG_TO_FONT_CODE_MAP.get(tgt_lang, 'en')
                
                logger.info("번역 작업자: 1단계 (텍스트/이미지) 시작")
                temp_dir_for_pptx_handler_main = tempfile.mkdtemp(prefix="pptx_trans_main_")
                temp_pptx_for_chart_translation_path: Optional[str] = None

                prs = Presentation(file_path)
                stage1_success = self.pptx_handler.translate_presentation_stage1(
                    prs, src_lang, tgt_lang,
                    self.translator, self.ocr_handler,
                    model, self.ollama_service,
                    font_code_for_render, task_log_filepath,
                    report_item_completed_from_handler,
                    self.stop_event
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
                        output_path = file_path
                elif not stage1_success:
                    logger.error("1단계 번역 실패.")
                    translation_result_status = "실패 (1단계 오류)"
                    output_path = file_path
                else:
                    logger.info("번역 작업자: 1단계 완료. 임시 파일 저장 시도.")
                    temp_pptx_for_chart_translation_path = os.path.join(
                        temp_dir_for_pptx_handler_main, 
                        f"{os.path.splitext(os.path.basename(file_path))[0]}_temp_for_charts.pptx"
                    )
                    if prs: prs.save(temp_pptx_for_chart_translation_path)
                    logger.info(f"1단계 결과 임시 저장: {temp_pptx_for_chart_translation_path}")

                    if self.current_file_chart_elements_count > 0 and not self.stop_event.is_set():
                        logger.info(f"번역 작업자: 2단계 (차트) 시작. 대상 차트 수: {self.current_file_chart_elements_count}")
                        safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                        final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                        final_output_dir = os.path.dirname(file_path)
                        final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)

                        output_path = self.chart_xml_handler.translate_charts_in_pptx(
                            pptx_path=temp_pptx_for_chart_translation_path,
                            src_lang_ui_name=src_lang,
                            tgt_lang_ui_name=tgt_lang,
                            model_name=model,
                            output_path=final_pptx_output_path,
                            progress_callback_item_completed=report_item_completed_from_handler,
                            stop_event=self.stop_event,
                            task_log_filepath=task_log_filepath
                        )
                        if self.stop_event.is_set():
                            logger.warning("2단계 차트 번역 중 또는 완료 직후 중지됨.")
                            translation_result_status = "부분 성공 (중지)"
                            if not (output_path and os.path.exists(output_path)): 
                                output_path = temp_pptx_for_chart_translation_path 
                        elif output_path and os.path.exists(output_path):
                            logger.info(f"2단계 차트 번역 완료. 최종 파일: {output_path}")
                            translation_result_status = "성공"
                        else:
                            logger.error("2단계 차트 번역 실패 또는 결과 파일 없음.")
                            translation_result_status = "실패 (2단계 오류)"
                            if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                                try:
                                    shutil.copy2(temp_pptx_for_chart_translation_path, final_pptx_output_path)
                                    output_path = final_pptx_output_path
                                    logger.info(f"차트 번역 실패로 1단계 결과물을 최종 경로에 복사: {output_path}")
                                except Exception as e_copy_fallback:
                                     logger.error(f"차트 번역 실패 후 1단계 결과물 복사 중 오류: {e_copy_fallback}.")
                                     output_path = temp_pptx_for_chart_translation_path
                            else:
                                output_path = file_path

                    elif self.stop_event.is_set():
                        logger.info("1단계 후 중단되어 차트 번역은 실행되지 않음.")
                        translation_result_status = "부분 성공 (중지)"
                        output_path = temp_pptx_for_chart_translation_path
                    else: # 차트가 없는 경우
                        logger.info("번역할 차트가 없습니다. 1단계 결과물을 최종 결과로 사용합니다.")
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
                                output_path = file_path
                        except Exception as e_copy_no_chart:
                            logger.error(f"차트 없는 경우 최종 파일 복사 중 오류: {e_copy_no_chart}")
                            translation_result_status = "실패 (파일 복사 오류)"
                            output_path = temp_pptx_for_chart_translation_path

                if 'temp_dir_for_pptx_handler_main' in locals() and temp_dir_for_pptx_handler_main and os.path.exists(temp_dir_for_pptx_handler_main):
                    try:
                        shutil.rmtree(temp_dir_for_pptx_handler_main)
                        logger.debug(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 완료.")
                    except Exception as e_clean_main_dir:
                        logger.warning(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 중 오류: {e_clean_main_dir}")

            if translation_result_status == "성공" and not self.stop_event.is_set():
                 self.current_weighted_done = self.total_weighted_work
                 if hasattr(self, 'master') and self.master.winfo_exists():
                     self.master.after(0, self.update_translation_progress,
                                   "완료", "종료", self.current_weighted_done, self.total_weighted_work, "번역 완료")
                 if hasattr(self, 'master') and self.master.winfo_exists():
                    self.master.after(100, lambda: self._ask_open_folder(output_path))
            
        except Exception as e: 
            logger.error(f"번역 작업 중 심각한 오류 발생: {e}", exc_info=True)
            translation_result_status = "오류 발생"
            if not output_path: output_path = file_path
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_err:
                    f_err.write(f"\n--- 번역 작업 중 심각한 오류 발생 ---\n오류: {e}\n")
                    import traceback
                    traceback.print_exc(file=f_err)
            except Exception as ef_log: logger.error(f"작업 로그 파일에 오류 기록 실패: {ef_log}")
        finally:
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self.translation_finished, translation_result_status, file_path, src_lang, tgt_lang, output_path, task_log_filepath)
            self.translation_thread = None

    def _ask_open_folder(self, path): 
        if path and os.path.exists(path) and messagebox.askyesno("번역 완료", f"번역이 완료되었습니다.\n저장된 폴더를 여시겠습니까?\n{path}"):
            utils.open_folder(os.path.dirname(path))

    def _format_time(self, seconds): 
        if seconds is None or seconds < 0: return "00:00.00" # 사용되지 않음 (진행률 표시 변경)
        m, s = divmod(seconds, 60); return f"{int(m):02d}:{s:05.2f}"

    def update_translation_progress(self, current_slide_info, current_element_type, 
                                    current_total_weighted_done, total_weighted_overall, 
                                    current_text_snippet=""):
        if self.stop_event.is_set(): return
        
        progress = 0
        if total_weighted_overall > 0:
            progress = (current_total_weighted_done / total_weighted_overall) * 100
        elif current_total_weighted_done == 0 and total_weighted_overall == 0 : # 작업량이 0인 경우
             progress = 100 # 내용 없는 파일은 즉시 100%로 간주

        progress = min(progress, 100) 
        progress_text_val = f"{progress:.1f}%" # 시간 관련 텍스트 제거

        def _update_ui():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            self.progress_bar["value"] = progress
            self.progress_label_var.set(progress_text_val)
            
            slide_display_text = current_slide_info
            if isinstance(current_slide_info, (int, float)):
                slide_display_text = f"슬라이드 {int(current_slide_info)}"
            elif not current_slide_info or str(current_slide_info).upper() == "N/A":
                slide_display_text = "차트 처리 중"
            elif str(current_slide_info).lower() == "완료":
                 slide_display_text = "작업 완료"


            self.current_slide_label.config(text=f"현재 위치: {slide_display_text}")
            
            display_text_snippet = current_text_snippet if len(current_text_snippet) < 30 else current_text_snippet[:27] + "..."
            self.current_work_label.config(text=f"현재 작업: {current_element_type} - '{display_text_snippet}'")
            
            self.translated_elements_label.config(text=f"완료된 작업량: {current_total_weighted_done}")
            self.remaining_elements_label.config(text=f"남은 작업량: {max(0, total_weighted_overall - current_total_weighted_done)}")
        
        if hasattr(self, 'master') and self.master.winfo_exists(): 
            self.master.after(0, _update_ui)

    def update_progress_timer(self): # 이 함수는 이제 시간 표시에 직접 관여하지 않음
        if self.translation_thread and self.translation_thread.is_alive() and not self.stop_event.is_set():
            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(1000, self.update_progress_timer)

    def stop_translation(self): 
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("번역 중지 요청 중..."); self.stop_event.set()
            self.stop_button.config(state=tk.DISABLED)

    def translation_finished(self, result_status, original_file, src_lang, tgt_lang, translated_file_path, task_log_filepath): 
        if not (hasattr(self, 'start_button') and self.start_button.winfo_exists()): return
        self.start_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED)
        
        current_progress_val = self.progress_bar["value"]
        final_progress_text = f"{current_progress_val:.1f}%" # 기본은 현재 %

        if hasattr(self, 'start_time') and self.start_time:
            elapsed_time_for_log = time.time() - self.start_time # 로그용 시간은 계산
        else:
            elapsed_time_for_log = 0


        if result_status == "성공":
            final_progress_text = "100%"
            self.progress_bar["value"] = 100
        elif "중지" in result_status or "취소" in result_status:
            final_progress_text = f"{current_progress_val:.1f}% (중지됨)"
        elif result_status == "내용 없음":
            final_progress_text = "100% (내용 없음)"
            self.progress_bar["value"] = 100 
        else: # 실패 또는 오류
            final_progress_text = f"{current_progress_val:.1f}% (오류)"
        self.progress_label_var.set(final_progress_text)
        
        if translated_file_path and os.path.exists(translated_file_path) and result_status not in ["취소됨", "오류 발생", "실패 (파일 없음)", "실패 (1단계 오류)", "실패 (2단계 오류)", "실패 (파일 오류)"]:
            self.translated_file_path_var.set(translated_file_path); self.open_folder_button.config(state=tk.NORMAL)
        else:
            self.translated_file_path_var.set("번역 실패 또는 파일 없음"); self.open_folder_button.config(state=tk.DISABLED)
            if not (translated_file_path and os.path.exists(translated_file_path)) and result_status == "성공":
                 logger.warning(f"번역은 '성공'으로 기록되었으나, 결과 파일 경로가 유효하지 않음: {translated_file_path}")
        
        file_name = os.path.basename(original_file); current_time_str = time.strftime("%Y-%m-%d %H:%M:%S")
        history_entry_values = (file_name, src_lang, tgt_lang, result_status, current_time_str, translated_file_path or original_file)
        TRANSLATION_HISTORY.append(history_entry_values)
        if hasattr(self, 'history_tree') and self.history_tree.winfo_exists():
            self.history_tree.insert("", tk.END, values=history_entry_values); self.history_tree.yview_moveto(1)
        
        if task_log_filepath and os.path.exists(os.path.dirname(task_log_filepath)):
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    f_task_log.write(f"\n--- 번역 작업 완료 ---\n")
                    f_task_log.write(f"최종 상태: {result_status}\n")
                    f_task_log.write(f"원본 파일: {original_file}\n")
                    if translated_file_path and os.path.exists(translated_file_path): f_task_log.write(f"번역된 파일: {translated_file_path}\n")
                    f_task_log.write(f"총 소요 시간 (내부 기록용): {self._format_time(elapsed_time_for_log)}\n") # 로그에는 시간 기록
            except Exception as e_log_finish: logger.error(f"작업 로그 파일에 최종 상태 기록 실패: {e_log_finish}")
        self.start_time = None

    def open_translated_folder(self): 
        path = self.translated_file_path_var.get()
        if path and os.path.exists(path): utils.open_folder(os.path.dirname(path))
        elif path: messagebox.showwarning("폴더 열기 실패", f"경로를 찾을 수 없습니다: {path}")

    def on_history_double_click(self, event): 
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()): return
        item_id = self.history_tree.identify_row(event.y)
        if item_id:
            item_values = self.history_tree.item(item_id, "values")
            if item_values and len(item_values) > 5:
                file_path_to_open = item_values[5]
                if file_path_to_open and os.path.exists(file_path_to_open):
                    if messagebox.askyesno("파일 열기", f"번역된 파일 '{os.path.basename(file_path_to_open)}'을(를) 여시겠습니까?"):
                        try:
                            if platform.system() == "Windows": os.startfile(file_path_to_open)
                            elif platform.system() == "Darwin": subprocess.Popen(["open", file_path_to_open])
                            else: subprocess.Popen(["xdg-open", file_path_to_open])
                        except Exception as e: logger.error(f"히스토리 파일 열기 실패: {e}", exc_info=True)
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
        self.text_widget.after(0, append_message)

if __name__ == "__main__":
    if not os.path.exists(config.LOGS_DIR): # config 사용
        try: os.makedirs(config.LOGS_DIR); print(f"로그 폴더 생성됨: {config.LOGS_DIR}")
        except Exception as e: print(f"메인 실행부: 로그 폴더 생성 실패: {config.LOGS_DIR}, 오류: {e}")
    
    # debug_mode는 main.py 상단에서 sys.argv로 이미 결정됨
    if debug_mode: logger.info("디버그 모드로 실행 중입니다.")
    else: logger.info("일반 모드로 실행 중입니다.")
    
    if not os.path.exists(config.FONTS_DIR): # config 사용
        logger.critical(f"필수 폰트 디렉토리를 찾을 수 없습니다: {config.FONTS_DIR}")
    else: logger.info(f"폰트 디렉토리 확인: {config.FONTS_DIR}")
    
    if not os.path.exists(config.ASSETS_DIR): # config 사용
        logger.warning(f"에셋 디렉토리를 찾을 수 없습니다: {config.ASSETS_DIR}")
    else: logger.info(f"에셋 디렉토리 확인: {config.ASSETS_DIR}")

    root = tk.Tk()
    app = Application(master=root)
    root.geometry("960x820") 
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