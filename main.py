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
import json
from typing import Optional, List, Dict, Any, Callable # IO 추가 (혹시 필요할 수 있음)
import traceback

from pptx import Presentation

import config
# --- 2단계: DIP 적용 (인터페이스를 통해 의존성 주입) ---
from interfaces import AbsOllamaService, AbsTranslator, AbsPptxProcessor, AbsChartProcessor, AbsOcrHandler, AbsOcrHandlerFactory
# 실제 구현체는 __main__ 블록에서 주입
from ollama_service import OllamaService
from translator import OllamaTranslator
from pptx_handler import PptxHandler
from chart_xml_handler import ChartXmlHandler
from ocr_handler import OcrHandlerFactory # 실제 팩토리 구현체

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

# --- 경로 설정 ---
BASE_DIR_MAIN = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = config.ASSETS_DIR
FONTS_DIR = config.FONTS_DIR
LOGS_DIR = config.LOGS_DIR
HISTORY_DIR = config.HISTORY_DIR
USER_SETTINGS_PATH = os.path.join(BASE_DIR_MAIN, config.USER_SETTINGS_FILENAME)

logger = logging.getLogger(__name__)

# --- 전역 변수 ---
APP_NAME = config.APP_NAME
DEFAULT_MODEL = config.DEFAULT_OLLAMA_MODEL
SUPPORTED_LANGUAGES = config.SUPPORTED_LANGUAGES


class Application(tk.Frame):
    def __init__(self, master=None,
                 ollama_service: AbsOllamaService = None,
                 translator: AbsTranslator = None,
                 pptx_handler: AbsPptxProcessor = None,
                 chart_processor: AbsChartProcessor = None,
                 ocr_handler_factory: AbsOcrHandlerFactory = None # 인터페이스 타입으로 변경
                 ):
        super().__init__(master)
        self.master = master
        self.master.title(APP_NAME)
        self.general_file_handler: Optional[logging.FileHandler] = None
        self._setup_logging_file_handler()

        self.user_settings: Dict[str, Any] = {}
        self._load_user_settings()

        # --- 2단계: 의존성 주입 ---
        self.ollama_service = ollama_service if ollama_service else OllamaService()
        self.translator = translator if translator else OllamaTranslator()
        self.pptx_handler = pptx_handler if pptx_handler else PptxHandler()
        self.chart_xml_handler = chart_processor if chart_processor else ChartXmlHandler(self.translator, self.ollama_service)
        self.ocr_handler_factory = ocr_handler_factory if ocr_handler_factory else OcrHandlerFactory()

        self.ocr_handler: Optional[AbsOcrHandler] = None
        self.current_ocr_engine_type: Optional[str] = None

        self._set_app_icon()

        self.style = ttk.Style()
        current_os = platform.system()
        if current_os == "Windows": self.style.theme_use('vista')
        elif current_os == "Darwin": self.style.theme_use('aqua')
        else: self.style.theme_use('clam')

        self.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.translation_thread: Optional[threading.Thread] = None
        self.model_download_thread: Optional[threading.Thread] = None
        self.stop_event = threading.Event()
        self.logo_image_tk_bottom: Optional[tk.PhotoImage] = None
        self.start_time: Optional[float] = None

        self.current_file_slide_count = 0
        self.current_file_total_text_chars = 0
        self.current_file_image_elements_count = 0
        self.current_file_chart_elements_count = 0
        self.total_weighted_work = 0
        self.current_weighted_done = 0

        # --- 3단계: UI 반응성 개선용 변수 ---
        self.last_progress_update_time = 0.0
        self.min_progress_update_interval = config.UI_PROGRESS_UPDATE_INTERVAL
        self.progress_update_threshold = 0.5
        self.last_reported_progress_percent = 0.0

        self.history_file_path = os.path.join(HISTORY_DIR, "translation_history.json")
        self.translation_history_data: List[Dict[str, Any]] = []

        self.ocr_temperature_var = tk.DoubleVar(
            value=self.user_settings.get("ocr_temperature", config.DEFAULT_ADVANCED_SETTINGS["ocr_temperature"])
        )
        self.image_translation_enabled_var = tk.BooleanVar(
            value=self.user_settings.get("image_translation_enabled", config.DEFAULT_ADVANCED_SETTINGS["image_translation_enabled"])
        )
        self.ocr_use_gpu_var = tk.BooleanVar(
            value=self.user_settings.get("ocr_use_gpu", config.DEFAULT_ADVANCED_SETTINGS["ocr_use_gpu"])
        )

        self.create_widgets()
        self._load_translation_history()
        self.master.after(100, self.initial_checks)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
        atexit.register(self.on_closing)

        log_file_path_msg = self.general_file_handler.baseFilename if self.general_file_handler else '미설정'
        logger.info(f"--- {APP_NAME} 시작됨 (일반 로그 파일: {log_file_path_msg}) ---")
        logger.info(f"로드된 사용자 설정: {self.user_settings}")


    def _set_app_icon(self): # 메서드 분리
        app_icon_png_path = os.path.join(ASSETS_DIR, "app_icon.png")
        app_icon_ico_path = os.path.join(ASSETS_DIR, "app_icon.ico")
        icon_set = False
        try:
            if platform.system() == "Windows" and os.path.exists(app_icon_ico_path):
                self.master.iconbitmap(app_icon_ico_path)
                icon_set = True
            if not icon_set and os.path.exists(app_icon_png_path):
                try:
                    icon_image_tk = tk.PhotoImage(file=app_icon_png_path, master=self.master)
                    self.master.iconphoto(True, icon_image_tk)
                    icon_set = True
                except tk.TclError:
                    try:
                        pil_icon = Image.open(app_icon_png_path)
                        icon_image_pil = ImageTk.PhotoImage(pil_icon, master=self.master)
                        self.master.iconphoto(True, icon_image_pil)
                        icon_set = True
                    except Exception as e_pil_icon_fallback:
                        logger.warning(f"Pillow로도 PNG 아이콘 설정 실패: {e_pil_icon_fallback}")
            if not icon_set:
                logger.warning(f"애플리케이션 아이콘 파일을 찾을 수 없거나 설정 실패: PNG='{app_icon_png_path}', ICO='{app_icon_ico_path}'")
        except Exception as e_icon_general:
            logger.warning(f"애플리케이션 아이콘 설정 중 예외: {e_icon_general}", exc_info=True)


    def _setup_logging_file_handler(self):
        if self.general_file_handler: # 이미 설정되어 있으면 반환
            return
        try:
            os.makedirs(LOGS_DIR, exist_ok=True) # 로그 디렉토리 생성
            general_log_filename = os.path.join(LOGS_DIR, "app_general.log")
            # FileHandler 설정
            self.general_file_handler = logging.FileHandler(general_log_filename, mode='a', encoding='utf-8')
            self.general_file_handler.setFormatter(formatter)
            # 핸들러 중복 추가 방지 (파일 경로 기반으로 확인)
            if not any(h.baseFilename == os.path.abspath(general_log_filename) for h in root_logger.handlers if isinstance(h, logging.FileHandler)):
                root_logger.addHandler(self.general_file_handler)
        except Exception as e:
            # 이 시점에서는 logger가 완전히 설정되지 않았을 수 있으므로 print 사용
            print(f"일반 로그 파일 핸들러 설정 실패: {e}")


    def _load_user_settings(self):
        """사용자 설정을 JSON 파일에서 로드합니다."""
        if os.path.exists(USER_SETTINGS_PATH):
            try:
                with open(USER_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                    if isinstance(loaded_settings, dict):
                        self.user_settings = loaded_settings
                        logger.info(f"사용자 설정 로드 완료: {USER_SETTINGS_PATH}")
                    else:
                        logger.warning(f"사용자 설정 파일({USER_SETTINGS_PATH}) 형식이 올바르지 않아 기본값 사용.")
                        self.user_settings = config.DEFAULT_ADVANCED_SETTINGS.copy() # 기본값으로 초기화
            except json.JSONDecodeError:
                logger.error(f"사용자 설정 파일({USER_SETTINGS_PATH}) 디코딩 오류. 기본값 사용.")
                self.user_settings = config.DEFAULT_ADVANCED_SETTINGS.copy()
            except Exception as e:
                logger.error(f"사용자 설정 로드 중 오류: {e}", exc_info=True)
                self.user_settings = config.DEFAULT_ADVANCED_SETTINGS.copy()
        else:
            logger.info(f"사용자 설정 파일 없음 ({USER_SETTINGS_PATH}). 기본값 사용.")
            self.user_settings = config.DEFAULT_ADVANCED_SETTINGS.copy() # 파일 없으면 기본값

    def _save_user_settings(self):
        """현재 고급 설정을 JSON 파일에 저장합니다."""
        settings_to_save = {
            "ocr_temperature": self.ocr_temperature_var.get(),
            "image_translation_enabled": self.image_translation_enabled_var.get(),
            "ocr_use_gpu": self.ocr_use_gpu_var.get()
            # 추가적인 사용자 설정이 있다면 여기에 포함
        }
        try:
            # 설정 파일이 저장될 디렉토리가 없다면 생성
            os.makedirs(os.path.dirname(USER_SETTINGS_PATH), exist_ok=True)
            with open(USER_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings_to_save, f, ensure_ascii=False, indent=4)
            logger.info(f"사용자 설정 저장 완료: {USER_SETTINGS_PATH}")
            self.user_settings = settings_to_save # 저장 후 내부 상태도 업데이트
        except Exception as e:
            logger.error(f"사용자 설정 저장 중 오류: {e}", exc_info=True)

    def _destroy_current_ocr_handler(self): # 2단계: OCR 핸들러 관리 로직은 팩토리 또는 OCR 관리 서비스로 이동/통합될 수 있음 (현재는 유지)
        if self.ocr_handler:
            logger.info(f"기존 OCR 핸들러 ({self.current_ocr_engine_type}) 자원 해제 시도...")
            if hasattr(self.ocr_handler, 'ocr_engine') and self.ocr_handler.ocr_engine:
                try:
                    # 명시적인 release 함수가 있다면 호출 (예: PaddleOCR의 경우 내부적으로 처리될 수 있음)
                    # if hasattr(self.ocr_handler.ocr_engine, 'release'): self.ocr_handler.ocr_engine.release()
                    del self.ocr_handler.ocr_engine
                    logger.debug(f"{self.current_ocr_engine_type} 엔진 객체 참조 제거됨.")
                except Exception as e:
                    logger.warning(f"OCR 엔진 객체('ocr_engine') 삭제 중 오류: {e}")

            self.ocr_handler = None
            self.current_ocr_engine_type = None
            logger.info("기존 OCR 핸들러 자원 해제 완료.")


    def on_closing(self):
        logger.info("애플리케이션 종료 절차 시작...")
        self._save_user_settings() # 종료 시 사용자 설정 저장

        if not self.stop_event.is_set(): # 중복 호출 방지
            self.stop_event.set() # 모든 백그라운드 작업에 중지 신호

            # 번역 스레드 종료 대기
            if self.translation_thread and self.translation_thread.is_alive():
                logger.info("번역 스레드 종료 대기 중...")
                self.translation_thread.join(timeout=5) # 최대 5초 대기
                if self.translation_thread.is_alive():
                    logger.warning("번역 스레드가 시간 내에 종료되지 않았습니다.")

            # 모델 다운로드 스레드 종료 대기
            if self.model_download_thread and self.model_download_thread.is_alive():
                logger.info("모델 다운로드 스레드 종료 대기 중...")
                self.model_download_thread.join(timeout=2) # 최대 2초 대기
                if self.model_download_thread.is_alive():
                    logger.warning("모델 다운로드 스레드가 시간 내에 정상 종료되지 않았습니다.")

            # OCR 핸들러 자원 해제
            self._destroy_current_ocr_handler()

            # 로깅 핸들러 닫기
            if self.general_file_handler:
                logger.debug(f"일반 로그 파일 핸들러({self.general_file_handler.baseFilename}) 닫기 시도.")
                try:
                    self.general_file_handler.close()
                    root_logger.removeHandler(self.general_file_handler) # 루트 로거에서 제거
                    self.general_file_handler = None # 핸들러 참조 제거
                    logger.info("일반 로그 파일 핸들러가 성공적으로 닫혔습니다.")
                except Exception as e_log_close:
                    logger.error(f"일반 로그 파일 핸들러 닫기 중 오류: {e_log_close}")
            else:
                logger.debug("일반 로그 파일 핸들러가 이미 닫혔거나 설정되지 않았습니다.")

        # Tkinter 윈도우 종료 (존재하는 경우)
        if hasattr(self, 'master') and self.master.winfo_exists():
            logger.info("모든 정리 작업 완료. 애플리케이션을 종료합니다.")
            self.master.destroy()
        else:
            # master가 없거나 이미 destroy된 경우 (atexit에 의해 여러 번 호출될 수 있음)
            logger.info("애플리케이션 윈도우가 이미 없으므로 바로 종료합니다.")
        
        # atexit에 등록된 경우, 이 함수가 다시 호출될 수 있으므로 sys.exit()는 신중히 사용
        # 여기서는 master.destroy() 후 mainloop가 자연스럽게 종료되도록 함


    def initial_checks(self):
        logger.debug("초기 점검 시작: OCR 라이브러리 설치 여부 및 Ollama 상태 확인")
        self.update_ocr_status_display() # OCR 상태 표시 업데이트
        self.check_ollama_status_manual(initial_check=True) # Ollama 서버 상태 확인
        logger.debug("초기 점검 완료.")

    def create_widgets(self):
        # UI 요소 생성 및 배치
        # 이전에 제공된 create_widgets 코드를 기반으로 복원 및 재구성합니다.
        # ttk 스타일 적용 부분을 __init__으로 옮겨도 무방합니다.

        top_frame = ttk.Frame(self)
        top_frame.pack(fill=tk.BOTH, expand=True)

        bottom_frame = ttk.Frame(self, height=30)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(5,0))
        bottom_frame.pack_propagate(False)

        main_paned_window = ttk.PanedWindow(top_frame, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        # --- 왼쪽 패널 ---
        left_panel = ttk.Frame(main_paned_window, padding=10)
        main_paned_window.add(left_panel, weight=3) # 왼쪽 패널 비중 조절

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

        file_progress_outer_frame = ttk.Frame(left_panel)
        file_progress_outer_frame.pack(padx=5, pady=5, fill=tk.X)
        file_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="파일 정보", padding=5)
        file_info_frame.pack(side=tk.LEFT, padx=(0,5), fill=tk.BOTH, expand=True)
        self.file_name_label = ttk.Label(file_info_frame, text="파일 이름: ")
        self.file_name_label.pack(anchor=tk.W, pady=1)
        self.slide_count_label = ttk.Label(file_info_frame, text="슬라이드 수: ")
        self.slide_count_label.pack(anchor=tk.W, pady=1)
        self.total_text_char_label = ttk.Label(file_info_frame, text="텍스트 글자 수: ")
        self.total_text_char_label.pack(anchor=tk.W, pady=1)
        self.image_elements_label = ttk.Label(file_info_frame, text="이미지 수: ")
        self.image_elements_label.pack(anchor=tk.W, pady=1)
        self.chart_elements_label = ttk.Label(file_info_frame, text="차트 수: ")
        self.chart_elements_label.pack(anchor=tk.W, pady=1)

        progress_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="진행 상황", padding=5)
        progress_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.current_slide_label = ttk.Label(progress_info_frame, text="현재 위치: -")
        self.current_slide_label.pack(anchor=tk.W, pady=1)
        self.current_work_label = ttk.Label(progress_info_frame, text="현재 작업: 대기 중")
        self.current_work_label.pack(anchor=tk.W, pady=1)

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

        model_selection_frame = ttk.Frame(translation_options_frame)
        model_selection_frame.grid(row=1, column=1, columnspan=4, padx=0, pady=0, sticky=tk.EW)
        model_selection_frame.columnconfigure(0, weight=1)
        ttk.Label(translation_options_frame, text="번역 모델:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.model_var = tk.StringVar(value=DEFAULT_MODEL)
        self.model_combo = ttk.Combobox(model_selection_frame, textvariable=self.model_var, state="disabled")
        self.model_combo.grid(row=0, column=0, padx=(5,0), pady=5, sticky=tk.EW)
        self.model_refresh_button = ttk.Button(model_selection_frame, text="🔄", command=self.load_ollama_models, width=3)
        self.model_refresh_button.grid(row=0, column=1, padx=(2,5), pady=5, sticky=tk.W)

        action_buttons_frame = ttk.Frame(left_panel, padding=(0,5,0,0))
        action_buttons_frame.pack(padx=5, pady=10, fill=tk.X)
        self.style.configure("Big.TButton", font=('TkDefaultFont', 11, 'bold'), foreground="black")
        self.start_button = ttk.Button(action_buttons_frame, text="번역 시작", command=self.start_translation, style="Big.TButton")
        self.start_button.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X, ipady=10)
        self.stop_button = ttk.Button(action_buttons_frame, text="번역 중지", command=self.stop_translation, state=tk.DISABLED, style="Big.TButton")
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=10)

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

        # --- 오른쪽 패널 ---
        right_panel = ttk.Frame(main_paned_window, padding=10)  # right_panel을 ttk.Frame으로 정의
        main_paned_window.add(right_panel, weight=2)           # right_panel을 main_paned_window에 추가 (weight는 원하는 비율로 조정)
        right_top_frame = ttk.Frame(right_panel)
        right_top_frame.pack(fill=tk.BOTH, expand=True)


        advanced_options_button = ttk.Button(
            right_panel, text="고급 옵션 설정...",
            command=self.open_advanced_options_popup
        )
        advanced_options_button.pack(fill=tk.X, padx=5, pady=(5,0), side=tk.BOTTOM)

        right_panel_notebook = ttk.Notebook(right_top_frame)
        right_panel_notebook.pack(fill=tk.BOTH, expand=True, pady=(0,0))

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

    def open_advanced_options_popup(self):
        popup = tk.Toplevel(self.master)
        popup.title("고급 옵션")
        popup.geometry("450x280")
        popup.resizable(False, False)
        popup.transient(self.master)
        popup.grab_set()

        temp_ocr_temp_var = tk.DoubleVar(value=self.ocr_temperature_var.get())
        temp_img_trans_enabled_var = tk.BooleanVar(value=self.image_translation_enabled_var.get())
        temp_ocr_gpu_var = tk.BooleanVar(value=self.ocr_use_gpu_var.get())

        main_frame = ttk.Frame(popup, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        temp_label_frame = ttk.LabelFrame(main_frame, text="이미지 번역 온도 설정", padding=10)
        temp_label_frame.pack(fill=tk.X, pady=5)

        temp_frame_inner = ttk.Frame(temp_label_frame)
        temp_frame_inner.pack(fill=tk.X, pady=2)

        temp_current_value_label = ttk.Label(temp_frame_inner, text=f"{temp_ocr_temp_var.get():.1f}")

        def _update_popup_temp_label(value_str):
            try:
                value = float(value_str)
                if temp_current_value_label.winfo_exists():
                    temp_current_value_label.config(text=f"{value:.1f}")
            except (ValueError, tk.TclError):
                pass

        ocr_temp_slider_popup = ttk.Scale(
            temp_frame_inner, from_=0.1, to=1.0, variable=temp_ocr_temp_var,
            orient=tk.HORIZONTAL, command=_update_popup_temp_label
        )
        ocr_temp_slider_popup.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,5))
        temp_current_value_label.pack(side=tk.LEFT)

        temp_description_frame = ttk.Frame(temp_label_frame)
        temp_description_frame.pack(fill=tk.X, pady=(0,5))
        ttk.Label(temp_description_frame, text="0.1 (정직함) <----------------------> 1.0 (창의적)", justify=tk.CENTER).pack(fill=tk.X)
        ttk.Label(temp_description_frame, text="(기본값: 0.4, 이미지 품질이 좋지 않을 경우 수치를 올리는 것이 번역에 도움 될 수 있음)", wraplength=400, justify=tk.LEFT, font=("TkDefaultFont",8)).pack(fill=tk.X)

        check_frame = ttk.Frame(main_frame)
        check_frame.pack(fill=tk.X, pady=10)
        image_trans_check_popup = ttk.Checkbutton(
            check_frame, text="이미지 내 텍스트 번역 실행",
            variable=temp_img_trans_enabled_var
        )
        image_trans_check_popup.pack(anchor=tk.W, padx=5, pady=2)

        ocr_gpu_check_popup = ttk.Checkbutton(
            check_frame, text="이미지 번역(OCR) 시 GPU 사용 (지원 시)",
            variable=temp_ocr_gpu_var
        )
        ocr_gpu_check_popup.pack(anchor=tk.W, padx=5, pady=2)

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20,0), side=tk.BOTTOM)

        def apply_settings():
            self.ocr_temperature_var.set(temp_ocr_temp_var.get())
            self.image_translation_enabled_var.set(temp_img_trans_enabled_var.get())

            gpu_setting_changed = self.ocr_use_gpu_var.get() != temp_ocr_gpu_var.get()
            self.ocr_use_gpu_var.set(temp_ocr_gpu_var.get())

            logger.info(f"고급 옵션 적용: 온도={self.ocr_temperature_var.get()}, 이미지번역={self.image_translation_enabled_var.get()}, OCR GPU={self.ocr_use_gpu_var.get()}")
            self._save_user_settings()

            if gpu_setting_changed:
                logger.info("OCR GPU 사용 설정 변경됨. 다음 번역 시 또는 OCR 상태 확인 시 적용됩니다.")
                self._destroy_current_ocr_handler()
                self.update_ocr_status_display()

            if popup.winfo_exists(): popup.destroy()

        def cancel_settings():
            if popup.winfo_exists(): popup.destroy()

        apply_button = ttk.Button(button_frame, text="적용", command=apply_settings)
        apply_button.pack(side=tk.RIGHT, padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=cancel_settings)
        cancel_button.pack(side=tk.RIGHT)

        popup.wait_window()


    def _load_translation_history(self):
        if not os.path.exists(HISTORY_DIR):
            try:
                os.makedirs(HISTORY_DIR, exist_ok=True)
            except Exception as e_mkdir:
                logger.error(f"히스토리 디렉토리({HISTORY_DIR}) 생성 실패: {e_mkdir}")
                self.translation_history_data = []
                return

        if os.path.exists(self.history_file_path):
            try:
                with open(self.history_file_path, 'r', encoding='utf-8') as f:
                    self.translation_history_data = json.load(f)
                self.translation_history_data.sort(key=lambda x: x.get('time', '0'), reverse=True)
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
            os.makedirs(HISTORY_DIR, exist_ok=True)
            self.translation_history_data.sort(key=lambda x: x.get('time', '0'), reverse=True)
            self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
            with open(self.history_file_path, 'w', encoding='utf-8') as f:
                json.dump(self.translation_history_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logger.error(f"번역 히스토리 저장 중 오류: {e}", exc_info=True)


    def _add_history_entry(self, entry: Dict[str, Any]):
        self.translation_history_data.insert(0, entry)
        self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
        self._save_translation_history()
        self._populate_history_treeview()


    def _populate_history_treeview(self):
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()):
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for entry in self.translation_history_data:
            values = (
                entry.get("name", "-"),
                entry.get("src", "-"),
                entry.get("tgt", "-"),
                entry.get("model", "-"),
                f"{entry.get('ocr_temp', '-')}",
                entry.get("status", "-"),
                entry.get("time", "-"),
                entry.get("path", "-")
            )
            self.history_tree.insert("", tk.END, values=values)
        if self.translation_history_data:
            self.history_tree.yview_moveto(0)

    def update_ocr_status_display(self): # 2단계: OCR 핸들러 팩토리 사용에 따른 수정
        selected_ui_lang = self.src_lang_var.get()
        # OcrHandlerFactory를 통해 엔진 이름과 OCR 언어 코드 가져오기
        engine_name_display = self.ocr_handler_factory.get_engine_name_display(selected_ui_lang)
        ocr_lang_code_to_use = self.ocr_handler_factory.get_ocr_lang_code(selected_ui_lang)

        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get()
        gpu_status_text = "(GPU 사용 예정)" if gpu_enabled_for_ocr else "(CPU 사용 예정)"

        if self.ocr_handler and self.current_ocr_engine_type == engine_name_display.lower():
            current_handler_lang_display = ""
            # 현재 핸들러의 언어 코드 가져오기 (AbsOcrHandler 인터페이스의 current_lang_codes 속성 사용)
            handler_langs = self.ocr_handler.current_lang_codes
            if isinstance(handler_langs, list): # EasyOCR의 경우 리스트일 수 있음
                current_handler_lang_display = ", ".join(handler_langs)
            elif isinstance(handler_langs, str): # PaddleOCR의 경우 문자열
                current_handler_lang_display = handler_langs

            gpu_in_use_text = "(GPU 사용 중)" if self.ocr_handler.use_gpu else "(CPU 사용 중)"
            self.ocr_status_label.config(text=f"{engine_name_display}: 준비됨 ({current_handler_lang_display}) {gpu_in_use_text}")
        else:
            self.ocr_status_label.config(text=f"{engine_name_display}: ({ocr_lang_code_to_use or selected_ui_lang}) 사용 예정 {gpu_status_text} (미확인)")



    def on_source_language_change(self, event=None):
        selected_ui_lang = self.src_lang_var.get()
        logger.info(f"원본 언어 변경됨: {selected_ui_lang}.")
        self.update_ocr_status_display()


    def browse_file(self):
        file_path = filedialog.askopenfilename(title="파워포인트 파일 선택", filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*")))
        if file_path:
            self.file_path_var.set(file_path)
            logger.info(f"파일 선택됨: {file_path}")
            self.load_file_info(file_path)
            self.translated_file_path_var.set("")
            self.open_folder_button.config(state=tk.DISABLED)
            self.current_work_label.config(text="파일 선택됨. 번역 대기 중.")

    def load_file_info(self, file_path):
        self.current_work_label.config(text="파일 분석 중...")
        self.master.update_idletasks()
        try:
            logger.debug(f"파일 정보 분석 중: {file_path}")
            file_name = os.path.basename(file_path)
            info = self.pptx_handler.get_file_info(file_path) # PptxHandler의 get_file_info 호출

            # get_file_info가 오류 시에도 딕셔너리를 반환하며, 특정 키의 값으로 오류를 판단한다고 가정
            if info.get("slide_count", -1) == -1 and info.get("total_text_char_count", -1) == -1 :
                self.file_name_label.config(text=f"파일 이름: {file_name} (분석 오류)")
                self.slide_count_label.config(text="슬라이드 수: -")
                self.total_text_char_label.config(text="텍스트 글자 수: -")
                self.image_elements_label.config(text="이미지 수: -")
                self.chart_elements_label.config(text="차트 수: -")
                self.total_weighted_work = 0
                self.current_work_label.config(text="파일 분석 실패!")
                # get_file_info 내부에서 오류 메시지 박스를 띄웠거나, 여기서 띄울 수 있음
                # messagebox.showerror("파일 분석 오류", f"'{file_name}' 파일 분석 중 오류 발생.")
                return

            self.current_file_slide_count = info.get('slide_count', 0)
            self.current_file_total_text_chars = info.get('total_text_char_count', 0)
            self.current_file_image_elements_count = info.get('image_elements_count', 0)
            self.current_file_chart_elements_count = info.get('chart_elements_count', 0)

            self.file_name_label.config(text=f"파일 이름: {file_name}")
            self.slide_count_label.config(text=f"슬라이드 수: {self.current_file_slide_count}")
            self.total_text_char_label.config(text=f"텍스트 글자 수: {self.current_file_total_text_chars}")
            self.image_elements_label.config(text=f"이미지 수: {self.current_file_image_elements_count}")
            self.chart_elements_label.config(text=f"차트 수: {self.current_file_chart_elements_count}")

            self.total_weighted_work = (self.current_file_total_text_chars * config.WEIGHT_TEXT_CHAR) + \
                                       (self.current_file_image_elements_count * config.WEIGHT_IMAGE) + \
                                       (self.current_file_chart_elements_count * config.WEIGHT_CHART)

            logger.info(f"파일 정보 분석 완료. 총 슬라이드: {self.current_file_slide_count}, 예상 가중 작업량: {self.total_weighted_work}")
            self.current_work_label.config(text="파일 분석 완료. 번역 대기 중.")

        except FileNotFoundError:
            logger.error(f"파일 찾기 오류 (UI): {file_path}")
            self.file_name_label.config(text="파일 이름: - (파일 없음)")
            self.slide_count_label.config(text="슬라이드 수: -")
            self.total_text_char_label.config(text="텍스트 글자 수: -")
            self.image_elements_label.config(text="이미지 수: -")
            self.chart_elements_label.config(text="차트 수: -")
            self.current_work_label.config(text="파일을 찾을 수 없습니다.")
            messagebox.showerror("파일 오류", f"선택한 파일({os.path.basename(file_path)})을 찾을 수 없습니다.")
        except Exception as e:
            logger.error(f"파일 정보 분석 중 UI에서 예외 발생: {e}", exc_info=True)
            self.file_name_label.config(text="파일 이름: - (오류)")
            self.slide_count_label.config(text="슬라이드 수: -")
            self.total_text_char_label.config(text="텍스트 글자 수: -")
            self.image_elements_label.config(text="이미지 수: -")
            self.chart_elements_label.config(text="차트 수: -")
            self.current_work_label.config(text="파일 분석 중 오류 발생.")
            messagebox.showerror("파일 분석 오류", f"선택한 파일({os.path.basename(file_path)})을 분석하는 중 오류가 발생했습니다.\n파일이 손상되었거나 지원하지 않는 형식일 수 있습니다.\n\n오류: {e}")


    def check_ollama_status_manual(self, initial_check=False):
        logger.info("Ollama 상태 확인 중...")
        self.ollama_check_button.config(state=tk.DISABLED)
        self.master.update_idletasks()

        ollama_installed = self.ollama_service.is_installed()
        self.ollama_status_label.config(text=f"Ollama 설치: {'설치됨' if ollama_installed else '미설치'}")

        if not ollama_installed:
            logger.warning("Ollama가 설치되어 있지 않습니다.")
            if not initial_check:
                if messagebox.askyesno("Ollama 설치 필요", "Ollama가 설치되어 있지 않습니다. Ollama 다운로드 페이지로 이동하시겠습니까?"):
                    webbrowser.open("https://ollama.com/download")
            self.ollama_running_label.config(text="Ollama 실행: 미설치")
            self.ollama_port_label.config(text="Ollama 포트: -")
            self.model_combo.config(values=[], state="disabled")
            self.model_var.set("")
            self.ollama_check_button.config(state=tk.NORMAL)
            return

        ollama_running, port = self.ollama_service.is_running()
        self.ollama_running_label.config(text=f"Ollama 실행: {'실행 중' if ollama_running else '미실행'}")
        self.ollama_port_label.config(text=f"Ollama 포트: {port if ollama_running and port else '-'}")

        if ollama_running:
            logger.info(f"Ollama 실행 중 (포트: {port}). 모델 목록 로드 시도.")
            self.load_ollama_models()
        else:
            logger.warning("Ollama가 설치되었으나 실행 중이지 않습니다. 자동 시작을 시도합니다.")
            self.model_combo.config(values=[], state="disabled")
            self.model_var.set("")
            if initial_check or messagebox.askyesno("Ollama 실행 필요", "Ollama가 실행 중이지 않습니다. 지금 시작하시겠습니까? (권장)"):
                if self.ollama_service.start_ollama():
                    logger.info("Ollama 자동 시작 성공. 잠시 후 상태를 다시 확인합니다.")
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        self.master.after(3000, lambda: self.check_ollama_status_manual(initial_check=initial_check))
                else:
                    logger.error("Ollama 자동 시작 실패. 수동으로 실행해주세요.")
                    if not initial_check:
                        messagebox.showwarning("Ollama 시작 실패", "Ollama를 자동으로 시작할 수 없습니다. 수동으로 실행 후 'Ollama 확인'을 눌러주세요.")
        self.ollama_check_button.config(state=tk.NORMAL)


    def load_ollama_models(self):
        logger.debug("Ollama 모델 목록 로드 중 (UI 요청)...")
        self.model_refresh_button.config(state=tk.DISABLED)
        self.master.update_idletasks()

        self.ollama_service.invalidate_models_cache()
        models = self.ollama_service.get_text_models()

        if models:
            self.model_combo.config(values=models, state="readonly")
            current_selected_model = self.model_var.get()
            if current_selected_model in models:
                self.model_var.set(current_selected_model)
            elif DEFAULT_MODEL in models:
                self.model_var.set(DEFAULT_MODEL)
            elif models:
                self.model_var.set(models[0])
            else:
                self.model_var.set("")
            logger.info(f"사용 가능 Ollama 모델: {models}")
            if DEFAULT_MODEL not in models and not self.model_var.get():
                self.download_default_model_if_needed(initial_check_from_ollama=True)
        else:
            self.model_combo.config(values=[], state="disabled")
            self.model_var.set("")
            logger.warning("Ollama에 로드된 모델이 없습니다.")
            self.download_default_model_if_needed(initial_check_from_ollama=True)
        self.model_refresh_button.config(state=tk.NORMAL)


    def download_default_model_if_needed(self, initial_check_from_ollama=False):
        current_models = self.ollama_service.get_text_models()
        if DEFAULT_MODEL not in current_models:
            logger.warning(f"기본 모델 ({DEFAULT_MODEL})이 설치되어 있지 않습니다.")
            if initial_check_from_ollama or messagebox.askyesno("기본 모델 다운로드", f"기본 번역 모델 '{DEFAULT_MODEL}'이(가) 없습니다. 지금 다운로드하시겠습니까? (시간 소요)"):
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드 시작...")
                self.start_button.config(state=tk.DISABLED)
                self.progress_bar["value"] = 0
                self.current_work_label.config(text=f"모델 다운로드 시작: {DEFAULT_MODEL}")
                self.progress_label_var.set(f"모델 다운로드 시작: {DEFAULT_MODEL}")

                if self.model_download_thread and self.model_download_thread.is_alive():
                    logger.warning("이미 모델 다운로드 스레드가 실행 중입니다.")
                    return

                self.stop_event.clear()
                self.model_download_thread = threading.Thread(target=self._model_download_worker, args=(DEFAULT_MODEL, self.stop_event), daemon=True)
                self.model_download_thread.start()
            else:
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드가 취소되었습니다.")
        else:
            logger.info(f"기본 모델 ({DEFAULT_MODEL})이 이미 설치되어 있습니다.")


    def _model_download_worker(self, model_name, stop_event_ref):
        success = self.ollama_service.pull_model_with_progress(model_name, self.update_model_download_progress, stop_event=stop_event_ref)
        if hasattr(self, 'master') and self.master.winfo_exists():
            self.master.after(0, self._model_download_finished, model_name, success)
        self.model_download_thread = None


    def _model_download_finished(self, model_name, success):
        if success:
            logger.info(f"'{model_name}' 모델 다운로드 완료.")
            self.load_ollama_models()
            self.current_work_label.config(text=f"모델 '{model_name}' 다운로드 완료.")
        else:
            logger.error(f"'{model_name}' 모델 다운로드 실패.")
            self.current_work_label.config(text=f"모델 '{model_name}' 다운로드 실패.")
            if not self.stop_event.is_set():
                messagebox.showerror("모델 다운로드 실패", f"'{model_name}' 모델 다운로드에 실패했습니다.\nOllama 서버 로그 또는 인터넷 연결을 확인해주세요.")

        if not (self.translation_thread and self.translation_thread.is_alive()):
            self.start_button.config(state=tk.NORMAL)
            self.progress_bar["value"] = 0
            self.progress_label_var.set("0%")
            if not success :
                self.current_work_label.config(text="모델 다운로드 실패. 재시도 요망.")
            else:
                self.current_work_label.config(text="대기 중")


    def update_model_download_progress(self, status_text, completed_bytes, total_bytes, is_error=False):
        if self.stop_event.is_set() and "중지됨" not in status_text : return

        percent = 0
        progress_str = status_text
        if total_bytes > 0:
            percent = (completed_bytes / total_bytes) * 100
            progress_str = f"{percent:.1f}%"

        def _update():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            if not is_error:
                self.progress_bar["value"] = percent
                self.progress_label_var.set(f"모델 다운로드: {progress_str} ({status_text})")
                self.current_work_label.config(text=f"모델 다운로드 중: {status_text} {progress_str}")
            else:
                self.progress_label_var.set(f"모델 다운로드 오류: {status_text}")
                self.current_work_label.config(text=f"모델 다운로드 오류: {status_text}")
            logger.log(logging.DEBUG if not is_error else logging.ERROR,
                       f"모델 다운로드 진행: {status_text} ({completed_bytes}/{total_bytes})")

        if hasattr(self, 'master') and self.master.winfo_exists():
            self.master.after(0, _update)

    def check_ocr_engine_status(self, is_called_from_start_translation=False): # 2단계: OCR 핸들러 팩토리 사용 및 관리 로직 변경
        self.current_work_label.config(text="OCR 엔진 확인 중...")
        self.master.update_idletasks()

        selected_ui_lang = self.src_lang_var.get()
        engine_name_display = self.ocr_handler_factory.get_engine_name_display(selected_ui_lang)
        ocr_lang_code_to_use = self.ocr_handler_factory.get_ocr_lang_code(selected_ui_lang)

        if not ocr_lang_code_to_use:
            msg = f"{engine_name_display}: UI 언어 '{selected_ui_lang}'에 대한 OCR 코드가 설정되지 않았습니다."
            self.ocr_status_label.config(text=msg)
            logger.error(msg)
            if is_called_from_start_translation:
                messagebox.showerror("OCR 설정 오류", msg)
            self.current_work_label.config(text="OCR 설정 오류!")
            return False

        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get()
        needs_reinit = False

        # OCR 핸들러 (재)초기화 필요 조건 검사
        if not self.ocr_handler:
            needs_reinit = True
            logger.debug("OCR 핸들러 없음, 재초기화 필요.")
        elif self.current_ocr_engine_type != engine_name_display.lower():
            needs_reinit = True
            logger.debug(f"OCR 엔진 타입 변경됨 ({self.current_ocr_engine_type} -> {engine_name_display.lower()}), 재초기화 필요.")
        elif self.ocr_handler.use_gpu != gpu_enabled_for_ocr:
            needs_reinit = True
            logger.debug(f"OCR GPU 설정 변경됨 (현재: {self.ocr_handler.use_gpu}, 요청: {gpu_enabled_for_ocr}), 재초기화 필요.")
        else:
            # 엔진 타입과 GPU 설정이 동일한 경우, 언어 코드 확인
            current_handler_langs = self.ocr_handler.current_lang_codes
            if isinstance(current_handler_langs, list): # EasyOCR
                if ocr_lang_code_to_use not in current_handler_langs:
                    needs_reinit = True
                    logger.debug(f"EasyOCR 언어 코드 변경됨 (현재: {current_handler_langs}, 요청: {ocr_lang_code_to_use}), 재초기화 필요.")
            elif isinstance(current_handler_langs, str): # PaddleOCR
                if current_handler_langs != ocr_lang_code_to_use:
                    needs_reinit = True
                    logger.debug(f"PaddleOCR 언어 코드 변경됨 (현재: {current_handler_langs}, 요청: {ocr_lang_code_to_use}), 재초기화 필요.")

        if needs_reinit:
            self._destroy_current_ocr_handler() # 기존 핸들러 자원 해제
            logger.info(f"{engine_name_display} 핸들러 (재)초기화 시도 (언어: {ocr_lang_code_to_use}, GPU: {gpu_enabled_for_ocr}).")
            self.current_work_label.config(text=f"{engine_name_display} 엔진 로딩 중 (언어: {ocr_lang_code_to_use}, GPU: {gpu_enabled_for_ocr})...")
            self.master.update_idletasks()

            # 팩토리를 통해 핸들러 생성 시도
            # utils.check_easyocr/paddleocr는 팩토리 내부에서 호출될 수 있으므로, 여기서는 직접 호출 안 함
            self.ocr_handler = self.ocr_handler_factory.get_ocr_handler(
                lang_code_ui=selected_ui_lang,
                use_gpu=gpu_enabled_for_ocr,
                debug_enabled=debug_mode
            )

            if self.ocr_handler:
                self.current_ocr_engine_type = engine_name_display.lower()
                logger.info(f"{engine_name_display} 핸들러 초기화 성공 (언어: {ocr_lang_code_to_use}, GPU: {gpu_enabled_for_ocr}).")
                self.current_work_label.config(text=f"{engine_name_display} 엔진 로딩 완료.")
            else: # 핸들러 생성 실패 (팩토리 내부에서 로깅 및 오류 처리)
                self.ocr_status_label.config(text=f"{engine_name_display}: 초기화 실패 ({ocr_lang_code_to_use}, GPU:{gpu_enabled_for_ocr})")
                # 사용자에게는 팩토리에서 반환된 None을 기반으로 좀 더 일반적인 메시지 표시
                if is_called_from_start_translation:
                    messagebox.showerror(f"{engine_name_display} 오류", f"{engine_name_display} 엔진 초기화 중 오류가 발생했습니다.\n설치 상태 및 설정을 확인해주세요.\n자세한 내용은 로그 파일에서 확인할 수 있습니다.")
                self.current_work_label.config(text=f"{engine_name_display} 엔진 초기화 실패!")
                return False # OCR 준비 실패

        self.update_ocr_status_display() # 최종 OCR 상태 UI 업데이트
        # 핸들러와 엔진 객체가 모두 정상적으로 생성되었는지 확인
        if self.ocr_handler and hasattr(self.ocr_handler, 'ocr_engine') and self.ocr_handler.ocr_engine:
            return True
        else:
            self.ocr_status_label.config(text=f"{engine_name_display} OCR: 준비 안됨 ({selected_ui_lang})")
            if is_called_from_start_translation and not needs_reinit : # 재초기화 시도가 아니었는데도 준비 안된 경우
                 messagebox.showwarning("OCR 오류", f"{engine_name_display} OCR 엔진을 사용할 수 없습니다.\n(엔진 객체 생성 실패 또는 내부 오류)\n이전 로그를 확인해주세요.")
            self.current_work_label.config(text=f"{engine_name_display} OCR 준비 안됨.")
            return False
        
    def swap_languages(self):
        src = self.src_lang_var.get()
        tgt = self.tgt_lang_var.get()
        self.src_lang_var.set(tgt)
        self.tgt_lang_var.set(src)
        logger.info(f"언어 스왑: {tgt} <-> {src}")
        self.on_source_language_change()

    def start_translation(self): # 2단계: OCR 핸들러 관리 로직 변경, 3단계: UI 피드백 세분화 관련 로직 추가
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "번역할 유효한 파워포인트 파일을 선택해주세요.\n'찾아보기' 버튼을 사용하여 파일을 선택할 수 있습니다.")
            return

        image_translation_really_enabled = self.image_translation_enabled_var.get()
        ocr_temperature_to_use = self.ocr_temperature_var.get()

        # 이미지 번역이 활성화된 경우에만 OCR 엔진 상태 확인 및 준비
        if image_translation_really_enabled:
            if not self.check_ocr_engine_status(is_called_from_start_translation=True): # 번역 시작 시점에서 호출됨을 알림
                # check_ocr_engine_status 내부에서 사용자에게 메시지 박스를 띄우므로, 여기서는 추가 메시지 최소화
                if not messagebox.askyesno("OCR 준비 실패", "이미지 내 텍스트 번역에 필요한 OCR 기능이 준비되지 않았거나 사용할 수 없습니다.\n이 경우 이미지 안의 글자는 번역되지 않습니다.\n\n계속 진행하시겠습니까? (텍스트/차트만 번역)"):
                    logger.warning("OCR 준비 실패로 사용자가 번역을 취소했습니다.")
                    self.current_work_label.config(text="번역 취소됨 (OCR 준비 실패).")
                    return
                logger.warning("OCR 핸들러 준비 실패. 이미지 번역 없이 진행합니다.")
                image_translation_really_enabled = False # OCR 실패 시 이미지 번역 기능 내부적으로 비활성화
        else: # 이미지 번역 옵션이 꺼져있다면
            logger.info("이미지 번역 옵션이 꺼져있으므로 OCR 엔진을 확인하지 않습니다.")
            self._destroy_current_ocr_handler() # 사용 안 할 OCR 핸들러는 자원 해제

        src_lang, tgt_lang, model = self.src_lang_var.get(), self.tgt_lang_var.get(), self.model_var.get()
        if not model:
            messagebox.showerror("모델 오류", "번역 모델을 선택해주세요.\nOllama 서버가 실행 중이고 모델이 다운로드되었는지 확인하세요.\n'Ollama 확인' 버튼과 모델 목록 '🔄' 버튼을 사용해볼 수 있습니다.")
            self.check_ollama_status_manual() # 상태 재확인 유도
            return
        if src_lang == tgt_lang:
            messagebox.showwarning("언어 동일", "원본 언어와 번역 언어가 동일합니다.\n다른 언어를 선택해주세요.")
            return

        ollama_running, _ = self.ollama_service.is_running()
        if not ollama_running:
            messagebox.showerror("Ollama 미실행", "Ollama 서버가 실행 중이지 않습니다.\nOllama를 실행한 후 'Ollama 확인' 버튼을 눌러주세요.")
            self.check_ollama_status_manual() # 상태 재확인 및 자동 시작 시도
            return

        if self.total_weighted_work <= 0: # 파일 정보 로드 후 작업량이 0인 경우
            logger.info("총 예상 작업량이 0입니다. 파일 정보를 다시 로드하여 확인합니다.")
            self.load_file_info(file_path) # 파일 정보 재로드
            if self.total_weighted_work <= 0: # 그래도 0이면
                messagebox.showinfo("정보", "번역할 내용이 없거나 작업량을 계산할 수 없습니다.\n파일 내용을 확인해주세요.")
                logger.warning("재확인 후에도 총 예상 작업량이 0 이하입니다. 번역을 시작하지 않습니다.")
                self.current_work_label.config(text="번역할 내용 없음.")
                return

        # 작업 로그 파일 이름 및 경로 설정
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_filename = os.path.basename(file_path)
        # 파일명에서 확장자 제외하고, 유효한 문자만 남기기 (로그 파일명 오류 방지)
        safe_original_filename_part = "".join(c if c.isalnum() or c in ['.', '_'] else '_' for c in os.path.splitext(original_filename)[0])
        task_log_filename = f"translation_{timestamp}_{safe_original_filename_part}.log"
        task_log_filepath = os.path.join(LOGS_DIR, task_log_filename)

        # 로깅 정보 구성 (실제 사용된 OCR 핸들러 정보 포함)
        ocr_engine_for_log = self.current_ocr_engine_type if image_translation_really_enabled and self.ocr_handler else '사용 안 함'
        ocr_temp_for_log = ocr_temperature_to_use if image_translation_really_enabled else 'N/A'
        ocr_gpu_for_log = 'N/A'
        if image_translation_really_enabled and self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu'):
            ocr_gpu_for_log = self.ocr_handler.use_gpu # 실제 핸들러의 GPU 사용 여부

        logger.info(f"번역 시작: '{original_filename}' ({src_lang} -> {tgt_lang}) using {model}. "
                    f"이미지 번역: {'활성' if image_translation_really_enabled else '비활성'}, "
                    f"OCR 엔진: {ocr_engine_for_log}, OCR 온도: {ocr_temp_for_log}, OCR GPU (실제 사용): {ocr_gpu_for_log}")

        # UI 상태 변경 및 스레드 시작
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.progress_label_var.set("0%")
        self.translated_file_path_var.set("") # 이전 결과 초기화
        self.open_folder_button.config(state=tk.DISABLED)
        self.current_weighted_done = 0 # 누적 진행량 초기화
        self.last_reported_progress_percent = 0.0 # 3단계: UI 업데이트 제어용 변수 초기화
        self.last_progress_update_time = 0.0      # 3단계: UI 업데이트 제어용 변수 초기화

        self.stop_event.clear() # 중지 이벤트 초기화

        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("이미 번역 스레드가 실행 중입니다.")
            messagebox.showwarning("번역 중복", "이미 다른 번역 작업이 진행 중입니다.")
            self.start_button.config(state=tk.NORMAL) # 버튼 상태 복원
            self.stop_button.config(state=tk.DISABLED)
            return

        self.current_work_label.config(text="번역 준비 중...")
        self.master.update_idletasks()

        self.translation_thread = threading.Thread(target=self._translation_worker,
                                                   args=(file_path, src_lang, tgt_lang, model, task_log_filepath,
                                                         image_translation_really_enabled, ocr_temperature_to_use),
                                                   daemon=True)
        self.start_time = time.time() # 번역 시작 시간 기록
        self.translation_thread.start()
        self.update_progress_timer() # 주기적 UI 업데이트 타이머 시작 (선택적)


    def _translation_worker(self, file_path, src_lang, tgt_lang, model, task_log_filepath,
                            image_translation_enabled: bool, ocr_temperature: float):
        output_path: Optional[str] = None
        translation_result_status = "실패"
        prs: Optional[Presentation] = None
        temp_dir_for_pptx_handler_main: Optional[str] = None

        try:
            # --- 3단계: 공통 유틸리티 함수/클래스 추출 (task_log_filepath 사용) ---
            # 작업 로그 파일 헤더 작성 (utils.setup_task_logging은 chart_xml_handler에서 사용)
            # main의 _translation_worker는 전체 흐름을 관장하므로, 여기서 직접 로그 파일에 초기 정보 기록
            initial_log_messages = [
                f"--- 번역 작업 시작 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---",
                f"원본 파일: {os.path.basename(file_path)}",
                f"원본 언어: {src_lang}, 대상 언어: {tgt_lang}, 번역 모델: {model}",
                f"이미지 번역 활성화: {image_translation_enabled}"
            ]
            if image_translation_enabled and self.ocr_handler:
                initial_log_messages.append(f"  OCR 엔진: {self.current_ocr_engine_type or '미지정'}")
                initial_log_messages.append(f"  OCR 번역 온도: {ocr_temperature}")
                gpu_in_use_log = self.ocr_handler.use_gpu if hasattr(self.ocr_handler, 'use_gpu') else 'N/A'
                initial_log_messages.append(f"  OCR GPU 사용 (실제): {gpu_in_use_log}")
            elif image_translation_enabled and not self.ocr_handler:
                initial_log_messages.append(f"  OCR 엔진: 사용 불가 (핸들러 준비 안됨)")
            initial_log_messages.append(f"총 예상 가중 작업량: {self.total_weighted_work}")
            initial_log_messages.append("-" * 30)

            # utils.setup_task_logging을 사용하지 않고 직접 파일에 기록 (더 유연한 제어 가능)
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_log_init:
                    for line in initial_log_messages:
                        f_log_init.write(line + "\n")
                    f_log_init.flush()
            except Exception as e_log_header:
                logger.error(f"작업 로그 파일 헤더 작성 실패 ({task_log_filepath}): {e_log_header}")


            # --- 3단계: UI 반응성 개선 (진행 상황 콜백 호출 형식) ---
            def report_item_completed_from_handler(slide_info_or_stage: Any, item_type_str: str, weighted_work_for_item: int, text_snippet_str: str):
                if self.stop_event.is_set(): return
                self.current_weighted_done += weighted_work_for_item
                self.current_weighted_done = min(self.current_weighted_done, self.total_weighted_work if self.total_weighted_work > 0 else weighted_work_for_item)
                if hasattr(self, 'master') and self.master.winfo_exists():
                    self.master.after(0, self.update_translation_progress,
                                      slide_info_or_stage, item_type_str,
                                      self.current_weighted_done, self.total_weighted_work,
                                      text_snippet_str)
            if self.total_weighted_work == 0:
                logger.warning("번역할 가중 작업량이 없습니다 (파일 내용 부재 또는 분석 오류).")
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     self.master.after(0, lambda: messagebox.showinfo("정보", "파일에 번역할 내용이 없습니다."))
                translation_result_status, output_path = "내용 없음", file_path
                try:
                    with open(task_log_filepath, 'a', encoding='utf-8') as f_log_empty:
                        f_log_empty.write(f"번역할 내용 없음. 원본 파일: {file_path}\n")
                except Exception as e_log_empty_write: logger.error(f"내용 없음 로그 기록 실패: {e_log_empty_write}")

            else:
                font_code_for_render = config.UI_LANG_TO_FONT_CODE_MAP.get(tgt_lang, 'en')
                if hasattr(self, 'master') and self.master.winfo_exists():
                    self.master.after(0, lambda: self.current_work_label.config(text="파일 로드 중..."))
                    self.master.update_idletasks()

                temp_dir_for_pptx_handler_main = tempfile.mkdtemp(prefix="pptx_trans_main_")
                temp_pptx_for_chart_translation_path: Optional[str] = None

                try:
                    prs = Presentation(file_path)
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        self.master.after(0, lambda: self.current_work_label.config(text="1단계 (텍스트/이미지) 처리 시작..."))

                    stage1_success = self.pptx_handler.translate_presentation_stage1(
                        prs, src_lang, tgt_lang, self.translator,
                        self.ocr_handler if image_translation_enabled else None,
                        model, self.ollama_service, font_code_for_render, task_log_filepath,
                        report_item_completed_from_handler,
                        self.stop_event,
                        image_translation_enabled, ocr_temperature
                    )

                    if self.stop_event.is_set():
                        logger.warning("1단계 번역 중 중지됨 (사용자 요청).")
                        translation_result_status = "부분 성공 (중지)"
                        try:
                            stopped_filename_s1 = os.path.join(temp_dir_for_pptx_handler_main, f"{os.path.splitext(os.path.basename(file_path))[0]}_stage1_stopped.pptx")
                            if prs: prs.save(stopped_filename_s1)
                            output_path = stopped_filename_s1
                            logger.info(f"1단계 중단, 부분 저장: {output_path}")
                        except Exception as e_save_stop:
                            logger.error(f"1단계 중단 후 저장 실패: {e_save_stop}")
                            output_path = file_path
                    elif not stage1_success:
                        logger.error("1단계 번역 실패 (PptxHandler 반환값 False).")
                        translation_result_status = "실패 (1단계 오류)"
                        output_path = file_path
                    else:
                        logger.info("번역 작업자: 1단계 (텍스트/이미지) 완료. 임시 파일 저장 시도.")
                        if hasattr(self, 'master') and self.master.winfo_exists():
                            self.master.after(0, lambda: self.current_work_label.config(text="1단계 완료. 차트 처리 준비 중..."))
                            self.master.update_idletasks()

                        temp_pptx_for_chart_translation_path = os.path.join(
                            temp_dir_for_pptx_handler_main,
                            f"{os.path.splitext(os.path.basename(file_path))[0]}_temp_for_charts.pptx"
                        )
                        if prs: prs.save(temp_pptx_for_chart_translation_path)
                        logger.info(f"1단계 결과 임시 저장 (차트 처리용): {temp_pptx_for_chart_translation_path}")

                        info_for_charts = self.pptx_handler.get_file_info(temp_pptx_for_chart_translation_path)
                        num_charts_in_prs = info_for_charts.get('chart_elements_count', 0) # 여기서 num_charts_in_prs 할당

                        if num_charts_in_prs > 0 and not self.stop_event.is_set():
                            if hasattr(self, 'master') and self.master.winfo_exists():
                                self.master.after(0, lambda: self.current_work_label.config(text=f"2단계 (차트) 처리 시작 ({num_charts_in_prs}개)..."))
                                self.master.update_idletasks()
                            logger.info(f"번역 작업자: 2단계 (차트) 시작. 대상 차트 수: {num_charts_in_prs}")

                            safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                            final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                            final_output_dir = os.path.dirname(file_path)
                            final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)

                            output_path_charts = self.chart_xml_handler.translate_charts_in_pptx(
                                pptx_path=temp_pptx_for_chart_translation_path,
                                src_lang_ui_name=src_lang, tgt_lang_ui_name=tgt_lang, model_name=model,
                                output_path=final_pptx_output_path,
                                progress_callback_item_completed=report_item_completed_from_handler,
                                stop_event=self.stop_event,
                                task_log_filepath=task_log_filepath
                            )

                            if self.stop_event.is_set():
                                logger.warning("2단계 차트 번역 중 또는 완료 직후 중지됨.")
                                translation_result_status = "부분 성공 (중지)"
                                output_path = output_path_charts if (output_path_charts and os.path.exists(output_path_charts)) else temp_pptx_for_chart_translation_path
                            elif output_path_charts and os.path.exists(output_path_charts):
                                logger.info(f"2단계 차트 번역 완료. 최종 파일: {output_path_charts}")
                                translation_result_status = "성공"
                                output_path = output_path_charts
                            else:
                                logger.error("2단계 차트 번역 실패 또는 결과 파일 없음. 1단계 결과물 사용 시도.")
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
                        else:
                            logger.info("번역할 차트가 없습니다. 1단계 결과물을 최종 결과로 사용합니다.")
                            if hasattr(self, 'master') and self.master.winfo_exists():
                                self.master.after(0, lambda: self.current_work_label.config(text="최종 파일 저장 중..."))
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
                                    output_path = file_path
                            except Exception as e_copy_no_chart:
                                logger.error(f"차트 없는 경우 최종 파일 복사 중 오류: {e_copy_no_chart}")
                                translation_result_status = "실패 (파일 복사 오류)"
                                output_path = temp_pptx_for_chart_translation_path if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path) else file_path
                finally:
                    if temp_dir_for_pptx_handler_main and os.path.exists(temp_dir_for_pptx_handler_main):
                        try:
                            shutil.rmtree(temp_dir_for_pptx_handler_main)
                            logger.debug(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 완료.")
                        except Exception as e_clean_main_dir:
                            logger.warning(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 중 오류: {e_clean_main_dir}")

            if translation_result_status == "성공" and not self.stop_event.is_set():
                 self.current_weighted_done = self.total_weighted_work
                 if hasattr(self, 'master') and self.master.winfo_exists():
                     self.master.after(0, self._force_update_translation_progress,
                                      "완료", "번역 완료됨",
                                      self.current_weighted_done, self.total_weighted_work,
                                      "최종 저장 완료")
                 if not (output_path and os.path.exists(output_path)):
                     logger.error(f"번역 '성공'으로 기록되었으나, 최종 결과 파일({output_path})을 찾을 수 없습니다.")
                     translation_result_status = "실패 (결과 파일 없음)"
                     output_path = file_path
                 else:
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        self.master.after(100, lambda: self._ask_open_folder(output_path))
            elif "실패" in translation_result_status or "오류" in translation_result_status:
                 if hasattr(self, 'master') and self.master.winfo_exists():
                     self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath)
                 if not output_path: output_path = file_path

        except Exception as e_worker:
            logger.error(f"번역 작업 중 심각한 오류 발생: {e_worker}", exc_info=True)
            translation_result_status = "치명적 오류 발생"
            output_path = output_path or file_path
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_err:
                    f_err.write(f"\n--- 번역 작업 중 심각한 오류 발생 ---\n오류: {e_worker}\n{traceback.format_exc()}");
            except Exception as ef_log:
                logger.error(f"작업 로그 파일에 오류 기록 실패: {ef_log}")
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath, str(e_worker))
        finally:
            if hasattr(self, 'master') and self.master.winfo_exists():
                history_entry = {
                    "name": os.path.basename(file_path), "src": src_lang, "tgt": tgt_lang, "model": model,
                    "ocr_temp": ocr_temperature if image_translation_enabled else "N/A",
                    "ocr_gpu": self.ocr_use_gpu_var.get() if image_translation_enabled and self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu') else "N/A",
                    "img_trans_enabled": image_translation_enabled, "status": translation_result_status,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "path": output_path or file_path,
                    "log_file": task_log_filepath
                }
                self.master.after(0, self.translation_finished, history_entry)
            self.translation_thread = None

    def _force_update_translation_progress(self, current_location_info: Any, current_task_type: str,
                                      current_total_weighted_done: int, total_weighted_overall: int,
                                      current_text_snippet: str = ""):
        """ UI 업데이트 조건(시간, 변화량)을 무시하고 강제로 진행률을 업데이트합니다. (3단계 추가) """
        # 이 함수는 _update_ui_progress를 직접 호출하여 UI를 즉시 업데이트
        if hasattr(self, 'master') and self.master.winfo_exists():
            self.master.after(0, self._update_ui_progress,
                              current_location_info, current_task_type,
                              current_total_weighted_done, total_weighted_overall,
                              current_text_snippet)

    def _update_ui_progress(self, current_location_info: Any, current_task_type: str,
                           current_total_weighted_done: int, total_weighted_overall: int,
                           current_text_snippet: str = ""):
        """실제 UI를 업데이트하는 내부 함수입니다. (3단계 추가 또는 기존 update_translation_progress에서 분리)"""
        if self.stop_event.is_set() or not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()):
            return

        progress = 0
        if total_weighted_overall > 0:
            progress = (current_total_weighted_done / total_weighted_overall) * 100
        elif current_total_weighted_done == 0 and total_weighted_overall == 0 : # 작업량 0일 때 완료로 간주
            progress = 100 # 이 경우에도 100%로 표시
        progress = min(max(0, progress), 100) # 0~100% 범위 보장
        progress_text_val = f"{progress:.1f}%"

        # --- 3단계: UI 피드백 구체화 ---
        # PptxHandler, ChartXmlHandler 등에서 전달된 구체적인 작업 타입과 위치 정보를 사용
        task_description = current_task_type
        location_display_text = str(current_location_info) # 예: "슬라이드 5", "차트 제목 번역" 등

        snippet_display = current_text_snippet.replace('\n', ' ').strip()
        if len(snippet_display) > 30: # UI에 표시될 텍스트 길이 제한
            snippet_display = snippet_display[:27] + "..."

        # UI 업데이트
        self.progress_bar["value"] = progress
        self.progress_label_var.set(progress_text_val)
        self.current_slide_label.config(text=f"현재 위치: {location_display_text}")
        self.current_work_label.config(text=f"현재 작업: {task_description} - '{snippet_display}'")

        self.last_reported_progress_percent = progress # 마지막으로 UI에 보고된 진행률 업데이트



    def _ask_open_folder(self, path: Optional[str]): # 3단계: _translation_worker에서 분리 (새 메서드)
        """번역 완료 후 사용자에게 폴더 열기 여부를 묻는 메서드입니다."""
        if path and os.path.exists(path):
            user_choice = messagebox.askyesnocancel(
                "번역 완료",
                f"번역이 완료되었습니다.\n저장된 파일: {os.path.basename(path)}\n\n결과 파일이 저장된 폴더를 여시겠습니까?",
                icon='info', default=messagebox.YES
            )
            if user_choice is True: # Yes
                utils.open_folder(os.path.dirname(path))
            # No 또는 Cancel은 아무 작업 안 함
        elif path: # 경로가 있지만 존재하지 않는 경우 (이론상 발생하면 안됨)
            logger.warning(f"_ask_open_folder 호출되었으나 파일 경로({path})가 존재하지 않음.")
            messagebox.showwarning("파일 오류", f"번역된 파일을 찾을 수 없습니다: {path}")
        # 경로가 없는 경우는 translation_finished에서 이미 처리하므로 여기서는 무시

    def _handle_translation_failure(self, status: str, original_file: str, log_file: str, error_details: str = ""): # 3단계: _translation_worker에서 분리 (새 메서드)
        """번역 실패 시 사용자에게 알림을 표시하는 메서드입니다."""
        logger.error(f"번역 실패: {status}, 원본: {original_file}, 로그: {log_file}, 상세: {error_details}")
        if hasattr(self, 'current_work_label') and self.current_work_label.winfo_exists():
            self.current_work_label.config(text=f"번역 실패: {status}")

        error_title = f"번역 작업 실패 ({status})"
        user_message = f"'{os.path.basename(original_file)}' 파일 번역 중 오류가 발생했습니다.\n\n상태: {status}\n"
        if error_details:
            user_message += f"오류 정보: {error_details[:200]}...\n\n" # 너무 길면 잘라서 표시
        user_message += "다음 사항을 확인해 보세요:\n- Ollama 서버가 정상적으로 실행 중인지 ('Ollama 확인' 버튼)\n- 선택한 번역 모델이 유효한지 (모델 목록 '🔄' 버튼)\n- 원본 파일이 손상되지 않았는지\n"
        if "GPU" in status.upper() or "CUDA" in status.upper() or "메모리 부족" in status or \
           (self.ocr_use_gpu_var.get() and ("OCR" in status.upper() or "엔진" in status.upper())): # 대소문자 구분 없이 GPU/엔진 관련 오류 감지
            user_message += "- 고급 옵션에서 'GPU 사용'을 해제하고 다시 시도해보세요.\n"
        user_message += f"\n자세한 내용은 로그 파일에서 확인할 수 있습니다.\n로그 파일: {log_file}"

        if messagebox.askyesno(error_title, user_message + "\n\n오류 로그가 저장된 폴더를 여시겠습니까?", icon='error'):
            try:
                utils.open_folder(os.path.dirname(log_file))
            except Exception as e_open_log_dir:
                logger.warning(f"로그 폴더 열기 실패: {e_open_log_dir}")
                messagebox.showinfo("정보", f"로그 폴더를 열 수 없습니다.\n경로: {os.path.dirname(log_file)}")

    def update_translation_progress(self, current_location_info: Any, current_task_type: str,
                                    current_total_weighted_done: int, total_weighted_overall: int,
                                    current_text_snippet: str = ""):
        """진행률 업데이트 요청을 처리하고, UI 업데이트 빈도를 제어합니다. (3단계 수정)"""
        if self.stop_event.is_set() or not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()):
            return

        # --- 3단계: UI 반응성 개선 (업데이트 빈도 제어) ---
        now = time.time()
        current_progress_percent = (current_total_weighted_done / total_weighted_overall) * 100 if total_weighted_overall > 0 else 0
        progress_diff = abs(current_progress_percent - self.last_reported_progress_percent)

        # 시간 간격 또는 진행률 변화량이 충분할 때, 또는 작업 완료 시에만 UI 업데이트
        should_update_ui = (now - self.last_progress_update_time >= self.min_progress_update_interval) or \
                           (progress_diff >= self.progress_update_threshold) or \
                           (current_total_weighted_done == total_weighted_overall and total_weighted_overall >= 0) # 작업량이 0이어도 완료 시 업데이트

        if should_update_ui:
            if hasattr(self, 'master') and self.master.winfo_exists():
                # _update_ui_progress를 직접 호출하거나, after를 통해 메인 스레드에서 호출
                self.master.after(0, self._update_ui_progress,
                                  current_location_info, current_task_type,
                                  current_total_weighted_done, total_weighted_overall,
                                  current_text_snippet)
            self.last_progress_update_time = now # 마지막 업데이트 시간 갱신
            # self.last_reported_progress_percent는 _update_ui_progress 내부에서 갱신


    def update_progress_timer(self):
        if self.translation_thread and self.translation_thread.is_alive() and not self.stop_event.is_set():
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(1000, self.update_progress_timer)


    def stop_translation(self):
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("번역 중지 요청 중..."); self.stop_event.set(); self.stop_button.config(state=tk.DISABLED); self.current_work_label.config(text="번역 중지 요청됨...")
        elif self.model_download_thread and self.model_download_thread.is_alive():
            logger.warning("모델 다운로드 중지 요청 중..."); self.stop_event.set(); self.stop_button.config(state=tk.DISABLED)


    def translation_finished(self, history_entry: Dict[str, Any]): # 3단계: UI 업데이트 로직 단순화 (콜백에서 처리)
        if not (hasattr(self, 'start_button') and self.start_button.winfo_exists()):
            logger.warning("translation_finished 호출 시 UI 요소(start_button) 없음. UI 업데이트 건너뜀.")
            return

        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        result_status = history_entry.get("status", "알 수 없음")
        translated_file_path = history_entry.get("path")
        current_progress_val_str = self.progress_label_var.get().replace('%', '')
        try:
            current_progress_val = float(current_progress_val_str)
        except ValueError:
            current_progress_val = 0.0 # 파싱 실패 시 기본값

        final_progress_text = f"{current_progress_val:.1f}% ({result_status})" # 기본적으로 현재 진행률과 상태 표시

        if result_status == "성공" and not self.stop_event.is_set():
            # _translation_worker에서 이미 100%로 업데이트했을 것이므로, 여기서는 상태 텍스트만 정리
            final_progress_text = "100% (완료)"
            self.current_work_label.config(text=f"번역 완료: {os.path.basename(translated_file_path) if translated_file_path else '파일 없음'}")
            self.current_slide_label.config(text="모든 작업 완료")
        elif "중지" in result_status:
            self.current_work_label.config(text="번역 중지됨.")
            # final_progress_text는 현재값 유지
        elif result_status == "내용 없음":
            final_progress_text = "100% (내용 없음)" # 내용 없어도 100%로
            self.current_work_label.config(text="번역할 내용 없음.")
        else: # 실패 또는 기타 오류
            self.current_work_label.config(text=f"번역 실패: {result_status}")
            # final_progress_text는 현재값 유지

        self.progress_label_var.set(final_progress_text)

        # 번역된 파일 경로 및 폴더 열기 버튼 상태 업데이트
        if translated_file_path and os.path.exists(translated_file_path) and result_status == "성공":
            self.translated_file_path_var.set(translated_file_path)
            self.open_folder_button.config(state=tk.NORMAL)
        else:
            self.translated_file_path_var.set("번역 실패 또는 파일 없음")
            self.open_folder_button.config(state=tk.DISABLED)
            if result_status == "성공" and not (translated_file_path and os.path.exists(translated_file_path)):
                logger.warning(f"번역은 '성공'으로 기록되었으나, 결과 파일 경로가 유효하지 않음: {translated_file_path}")

        self._add_history_entry(history_entry) # 히스토리 추가

        # 작업 로그 파일에 최종 상태 기록
        task_log_filepath = history_entry.get("log_file")
        if task_log_filepath and os.path.exists(os.path.dirname(task_log_filepath)): # 로그 파일 디렉토리 존재 확인
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    f_task_log.write(f"\n--- 번역 작업 최종 상태 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                    f_task_log.write(f"최종 상태: {result_status}\n")
                    if self.file_path_var.get(): # UI에 표시된 원본 파일 경로
                        f_task_log.write(f"원본 파일 (UI 경로): {self.file_path_var.get()}\n")
                    if translated_file_path and os.path.exists(translated_file_path):
                        f_task_log.write(f"번역된 파일: {translated_file_path}\n")

                    elapsed_time_for_log = (time.time() - self.start_time) if self.start_time else 0
                    m, s = divmod(elapsed_time_for_log, 60)
                    f_task_log.write(f"총 소요 시간 (내부 기록용): {int(m):02d}분 {s:05.2f}초\n")
                    f_task_log.write("-" * 30 + "\n")
            except Exception as e_log_finish:
                logger.error(f"작업 로그 파일에 최종 상태 기록 실패: {e_log_finish}")

        self.start_time = None # 번역 시작 시간 초기화

    def open_translated_folder(self):
        path = self.translated_file_path_var.get()
        if path and os.path.exists(path): utils.open_folder(os.path.dirname(path))
        elif path and not os.path.exists(path): messagebox.showwarning("폴더 열기 실패", f"경로를 찾을 수 없습니다: {path}")
        else: messagebox.showinfo("정보", "번역된 파일 경로가 없습니다.")


    def on_history_double_click(self, event):
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()): return
        item_id = self.history_tree.identify_row(event.y)
        if item_id:
            item_values = self.history_tree.item(item_id, "values")
            if item_values and len(item_values) == len(self.history_tree["columns"]):
                try: path_idx = self.history_tree["columns"].index("path"); status_idx = self.history_tree["columns"].index("status"); time_idx = self.history_tree["columns"].index("time")
                except ValueError: logger.error("히스토리 Treeview 컬럼명 오류. 'path' 또는 'status' 컬럼을 찾을 수 없습니다."); return
                file_path_to_open = item_values[path_idx]; status_val = item_values[status_idx]; time_val = item_values[time_idx]
                if file_path_to_open and os.path.exists(file_path_to_open) and "성공" in status_val :
                    if messagebox.askyesno("파일 열기", f"번역된 파일 '{os.path.basename(file_path_to_open)}'을(를) 여시겠습니까?"):
                        try:
                            if platform.system() == "Windows": os.startfile(file_path_to_open)
                            elif platform.system() == "Darwin": subprocess.Popen(["open", file_path_to_open])
                            else: subprocess.Popen(["xdg-open", file_path_to_open])
                        except Exception as e: logger.error(f"히스토리 파일 열기 실패: {e}", exc_info=True); messagebox.showerror("파일 열기 오류", f"파일을 여는 중 오류가 발생했습니다:\n{e}")
                elif "성공" not in status_val and file_path_to_open :
                     log_file_path_from_history = ""
                     for entry_data in self.translation_history_data:
                         if entry_data.get("path") == file_path_to_open and entry_data.get("time") == time_val: log_file_path_from_history = entry_data.get("log_file", ""); break
                     if log_file_path_from_history and os.path.exists(log_file_path_from_history):
                         if messagebox.askyesno("로그 파일 열기", f"번역 결과가 '{status_val}'입니다.\n관련 로그 파일 '{os.path.basename(log_file_path_from_history)}'이(가) 저장된 폴더를 여시겠습니까?"):
                             try: utils.open_folder(os.path.dirname(log_file_path_from_history))
                             except Exception as e: logger.error(f"히스토리 로그 폴더 열기 실패: {e}"); messagebox.showerror("폴더 열기 오류", f"로그 폴더를 여는 중 오류가 발생했습니다:\n{e}")
                     else: messagebox.showwarning("정보", f"번역 결과가 '{status_val}'입니다.\n(관련 로그 파일 정보 없음 또는 찾을 수 없음)")
                elif file_path_to_open and not os.path.exists(file_path_to_open): messagebox.showwarning("파일 없음", f"파일을 찾을 수 없습니다: {file_path_to_open}")


# Text 위젯으로 로그를 보내는 핸들러
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        if not (self.text_widget and self.text_widget.winfo_exists()): return
        msg = self.format(record)
        def append_message():
            if not (self.text_widget and self.text_widget.winfo_exists()): return
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END) # 자동 스크롤
            self.text_widget.config(state=tk.DISABLED)
        try:
            # Tkinter 위젯 관련 작업은 메인 스레드에서 실행되도록 함
            if self.text_widget.winfo_exists():
                self.text_widget.after(0, append_message)
        except tk.TclError: # 위젯이 파괴된 후 호출될 경우 대비
            pass


if __name__ == "__main__":
    # 필수 디렉토리 생성
    for dir_path in [LOGS_DIR, FONTS_DIR, ASSETS_DIR, HISTORY_DIR, os.path.dirname(USER_SETTINGS_PATH)]:
        try:
            if dir_path:
                os.makedirs(dir_path, exist_ok=True)
        except Exception as e_mkdir_main:
            print(f"필수 디렉토리 생성 실패 ({dir_path}): {e_mkdir_main}") # 로거 설정 전일 수 있으므로 print 사용
            logger.critical(f"필수 디렉토리 생성 실패 ({dir_path}): {e_mkdir_main}", exc_info=True)

    if debug_mode: logger.info("디버그 모드로 실행 중입니다.")
    else: logger.info("일반 모드로 실행 중입니다.")

    # 필수 디렉토리 존재 여부 재확인 (로깅용)
    if not os.path.exists(config.FONTS_DIR) or not os.listdir(config.FONTS_DIR):
        logger.critical(f"필수 폰트 디렉토리({config.FONTS_DIR})를 찾을 수 없거나 비어있습니다. 애플리케이션이 정상 동작하지 않을 수 있습니다.")
    else: logger.info(f"폰트 디렉토리 확인: {config.FONTS_DIR}")

    if not os.path.exists(config.ASSETS_DIR):
        logger.warning(f"에셋 디렉토리를 찾을 수 없습니다: {config.ASSETS_DIR}")
    else: logger.info(f"에셋 디렉토리 확인: {config.ASSETS_DIR}")

    root = tk.Tk()

    # --- 2단계: 의존성 주입 구성 ---
    # 애플리케이션 시작 지점에서 실제 핸들러 구현체들을 생성
    ollama_service_instance = OllamaService()
    translator_instance = OllamaTranslator()
    pptx_handler_instance = PptxHandler()
    # ChartXmlHandler는 OllamaTranslator와 OllamaService를 필요로 함
    chart_processor_instance = ChartXmlHandler(translator_instance, ollama_service_instance)
    ocr_handler_factory_instance = OcrHandlerFactory() # 실제 팩토리 구현체 사용

    # Application 클래스에 주입
    app = Application(master=root,
                      ollama_service=ollama_service_instance,
                      translator=translator_instance,
                      pptx_handler=pptx_handler_instance,
                      chart_processor=chart_processor_instance,
                      ocr_handler_factory=ocr_handler_factory_instance)

    root.geometry("1024x768") # 기본 창 크기
    # UI 요소들이 모두 생성된 후 최소 크기 계산 및 설정
    root.update_idletasks()
    min_width = root.winfo_reqwidth()
    min_height = root.winfo_reqheight()
    root.minsize(min_width + 20, min_height + 20) # 여유 공간 추가

    try:
        root.mainloop()
    except KeyboardInterrupt:
        logger.info("Ctrl+C로 애플리케이션 종료 중...")
    finally:
        # on_closing이 atexit으로도 등록되어 있으므로, 여기서 중복 호출될 수 있음
        # on_closing 내부에서 중복 실행 방지 로직이 중요
        logger.info(f"--- {APP_NAME} 종료됨 (mainloop 이후) ---")