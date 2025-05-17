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
import json # For history and user settings
from typing import Optional, List, Dict, Any, Callable
import traceback


from pptx import Presentation

# 프로젝트 설정 파일 import
import config

# 프로젝트 루트의 다른 .py 파일들 import
from translator import OllamaTranslator
from pptx_handler import PptxHandler
from ocr_handler import PaddleOcrHandler, EasyOcrHandler # BaseOcrHandler는 여기서 직접 사용 안 함
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
HISTORY_DIR = config.HISTORY_DIR # 번역 히스토리 저장 경로 (config.py에서 정의)
USER_SETTINGS_PATH = os.path.join(BASE_DIR_MAIN, config.USER_SETTINGS_FILENAME)


logger = logging.getLogger(__name__)

# --- 전역 변수 및 설정 (config.py에서 가져옴) ---
APP_NAME = config.APP_NAME
DEFAULT_MODEL = config.DEFAULT_OLLAMA_MODEL
SUPPORTED_LANGUAGES = config.SUPPORTED_LANGUAGES


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title(APP_NAME)
        self.general_file_handler = None # 파일 로깅 핸들러
        self._setup_logging_file_handler() # 로깅 핸들러 먼저 설정

        self.user_settings: Dict[str, Any] = {}
        self._load_user_settings() # 사용자 설정 로드

        # 서비스/핸들러 인스턴스 생성 (2단계에서 인터페이스 기반 주입으로 변경 예정)
        self.ollama_service = OllamaService()
        self.translator = OllamaTranslator()
        self.pptx_handler = PptxHandler()
        # ChartXmlHandler는 translator와 ollama_service에 의존하므로, 해당 인스턴스 전달
        self.chart_xml_handler = ChartXmlHandler(self.translator, self.ollama_service)
        self.ocr_handler = None # 동적으로 생성 (PaddleOcrHandler 또는 EasyOcrHandler)
        self.current_ocr_engine_type = None # 현재 사용 중인 OCR 엔진 ("paddleocr" 또는 "easyocr")


        # 아이콘 설정
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
                except tk.TclError: # Tkinter PhotoImage가 PNG 직접 지원 못하는 경우
                    try:
                        pil_icon = Image.open(app_icon_png_path)
                        icon_image_pil = ImageTk.PhotoImage(pil_icon, master=self.master)
                        self.master.iconphoto(True, icon_image_pil)
                        icon_set = True
                    except Exception as e_pil_icon_fallback:
                        logger.warning(f"Pillow로도 PNG 아이콘 설정 실패: {e_pil_icon_fallback}")
            if not icon_set:
                logger.warning(f"애플리케이션 아이콘 파일을 찾을 수 없거나 설정 실패.")
        except Exception as e_icon_general:
            logger.warning(f"애플리케이션 아이콘 설정 중 예외: {e_icon_general}", exc_info=True)


        # 스타일 설정
        self.style = ttk.Style()
        current_os = platform.system()
        if current_os == "Windows":
            self.style.theme_use('vista')
        elif current_os == "Darwin": # macOS
            self.style.theme_use('aqua')
        else: # Linux 등 기타
            self.style.theme_use('clam') # 'clam' 또는 'alt', 'default', 'classic' 등 사용 가능

        self.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)


        # UI 관련 변수 및 상태 변수
        self.translation_thread = None
        self.model_download_thread = None
        self.stop_event = threading.Event()
        self.logo_image_tk_bottom = None # 하단 로고 이미지
        self.start_time = None # 번역 시작 시간

        # 현재 파일 정보 및 진행률 관련 변수
        self.current_file_slide_count = 0
        self.current_file_total_text_chars = 0
        self.current_file_image_elements_count = 0
        self.current_file_chart_elements_count = 0
        self.total_weighted_work = 0 # 총 예상 작업량 (가중치 적용)
        self.current_weighted_done = 0 # 현재까지 완료된 작업량 (가중치 적용)

        # 번역 히스토리 관련
        self.history_file_path = os.path.join(HISTORY_DIR, "translation_history.json")
        self.translation_history_data: List[Dict[str, Any]] = []


        # 고급 옵션 UI 변수 (tk.BooleanVar, tk.DoubleVar 등)
        # Application 생성자에서 tk.BooleanVar 등의 초기값을 저장된 설정 또는 config.py의 기본값으로 설정
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
        self._load_translation_history() # 번역 히스토리 로드
        self.master.after(100, self.initial_checks) # 초기 상태 점검 (Ollama, OCR 등)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing) # 종료 시 처리
        atexit.register(self.on_closing) # 비정상 종료 시에도 호출되도록

        log_file_path_msg = self.general_file_handler.baseFilename if self.general_file_handler else '미설정'
        logger.info(f"--- {APP_NAME} 시작됨 (일반 로그 파일: {log_file_path_msg}) ---")
        logger.info(f"로드된 사용자 설정: {self.user_settings}")


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


    def _destroy_current_ocr_handler(self):
        if self.ocr_handler:
            logger.info(f"기존 OCR 핸들러 ({self.current_ocr_engine_type}) 자원 해제 시도...")
            # OCR 엔진 객체가 ocr_engine 속성에 저장되어 있다고 가정
            if hasattr(self.ocr_handler, 'ocr_engine') and self.ocr_handler.ocr_engine:
                try:
                    # PaddleOCR/EasyOCR의 명시적인 자원 해제 함수가 있다면 호출
                    # 예: if hasattr(self.ocr_handler.ocr_engine, 'release'): self.ocr_handler.ocr_engine.release()
                    del self.ocr_handler.ocr_engine # 참조 제거로 GC 유도
                    logger.debug(f"{self.current_ocr_engine_type} 엔진 객체 참조 제거됨.")
                except Exception as e:
                    logger.warning(f"OCR 엔진 객체('ocr_engine') 삭제 중 오류: {e}")

            self.ocr_handler = None
            self.current_ocr_engine_type = None
            # 강제 GC (메모리 회수에 도움될 수 있으나, 남용 주의)
            # import gc
            # gc.collect()
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
        # 이전에 제공된 create_widgets 코드를 기반으로 복원합니다.
        # 고급 옵션 팝업에서 변수 초기화 시 self.user_settings 또는 config.DEFAULT_ADVANCED_SETTINGS 사용
        # self.ocr_temperature_var, self.image_translation_enabled_var, self.ocr_use_gpu_var는
        # __init__에서 이미 사용자 설정/기본값으로 초기화되었으므로, create_widgets에서는 해당 변수 사용.

        top_frame = ttk.Frame(self)
        top_frame.pack(fill=tk.BOTH, expand=True)

        bottom_frame = ttk.Frame(self, height=30)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(5,0))
        bottom_frame.pack_propagate(False) # 높이 고정

        # 메인 화면을 좌우로 나누는 PanedWindow
        main_paned_window = ttk.PanedWindow(top_frame, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)

        # 왼쪽 패널 (입력, 옵션, 진행상황 등)
        left_panel = ttk.Frame(main_paned_window, padding=10)
        main_paned_window.add(left_panel, weight=3) # 왼쪽 패널이 더 넓게

        # 오른쪽 패널 (로그, 히스토리, 고급옵션 버튼 등)
        right_panel = ttk.Frame(main_paned_window, padding=0) # 오른쪽은 패딩 최소화
        main_paned_window.add(right_panel, weight=2)


        # --- Left Panel ---
        # 파일 경로 프레임
        path_frame = ttk.LabelFrame(left_panel, text="파일 경로", padding=5)
        path_frame.pack(padx=5, pady=(0,5), fill=tk.X)
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var, width=60)
        file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        browse_button = ttk.Button(path_frame, text="찾아보기", command=self.browse_file)
        browse_button.pack(side=tk.LEFT)

        # 서버 상태 프레임
        server_status_frame = ttk.LabelFrame(left_panel, text="서버 상태", padding=5)
        server_status_frame.pack(padx=5, pady=5, fill=tk.X)
        server_status_frame.columnconfigure(1, weight=1) # Ollama 실행 상태 레이블이 공간 차지하도록

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


        # 파일 정보 및 진행 상황 표시를 위한 프레임 (좌우로 나눔)
        file_progress_outer_frame = ttk.Frame(left_panel)
        file_progress_outer_frame.pack(padx=5, pady=5, fill=tk.X)

        # 파일 정보 표시 프레임 (왼쪽)
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


        # 진행 상황 정보 표시 프레임 (오른쪽)
        progress_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="진행 상황", padding=5)
        progress_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.current_slide_label = ttk.Label(progress_info_frame, text="현재 위치: -")
        self.current_slide_label.pack(anchor=tk.W, pady=1)
        self.current_work_label = ttk.Label(progress_info_frame, text="현재 작업: 대기 중")
        self.current_work_label.pack(anchor=tk.W, pady=1)


        # 번역 옵션 프레임
        translation_options_frame = ttk.LabelFrame(left_panel, text="번역 옵션", padding=5)
        translation_options_frame.pack(padx=5, pady=5, fill=tk.X)
        translation_options_frame.columnconfigure(1, weight=1) # 원본 언어 콤보박스 확장
        translation_options_frame.columnconfigure(4, weight=1) # 번역 언어 콤보박스 확장

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

        # 모델 선택 부분 (콤보박스와 새로고침 버튼을 한 프레임에)
        model_selection_frame = ttk.Frame(translation_options_frame) # 패딩 제거
        model_selection_frame.grid(row=1, column=1, columnspan=4, padx=0, pady=0, sticky=tk.EW) # columnspan=4로 확장
        model_selection_frame.columnconfigure(0, weight=1) # 콤보박스가 남은 공간 모두 차지

        ttk.Label(translation_options_frame, text="번역 모델:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.model_var = tk.StringVar(value=DEFAULT_MODEL)
        self.model_combo = ttk.Combobox(model_selection_frame, textvariable=self.model_var, state="disabled") # 초기 비활성화
        self.model_combo.grid(row=0, column=0, padx=(5,0), pady=5, sticky=tk.EW)
        self.model_refresh_button = ttk.Button(model_selection_frame, text="🔄", command=self.load_ollama_models, width=3)
        self.model_refresh_button.grid(row=0, column=1, padx=(2,5), pady=5, sticky=tk.W) # 오른쪽 끝에 붙임


        # 시작/중지 버튼 프레임
        action_buttons_frame = ttk.Frame(left_panel, padding=(0,5,0,0)) # 버튼 간격 조절
        action_buttons_frame.pack(padx=5, pady=10, fill=tk.X)

        self.style.configure("Big.TButton", font=('TkDefaultFont', 11, 'bold'), foreground="black") # 버튼 스타일

        self.start_button = ttk.Button(action_buttons_frame, text="번역 시작", command=self.start_translation, style="Big.TButton")
        self.start_button.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X, ipady=10)

        self.stop_button = ttk.Button(action_buttons_frame, text="번역 중지", command=self.stop_translation, state=tk.DISABLED, style="Big.TButton")
        self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=10)


        # 진행률 표시 바 프레임
        progress_bar_frame = ttk.Frame(left_panel)
        progress_bar_frame.pack(padx=5, pady=5, fill=tk.X)
        self.progress_bar = ttk.Progressbar(progress_bar_frame, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        self.progress_label_var = tk.StringVar(value="0%")
        ttk.Label(progress_bar_frame, textvariable=self.progress_label_var).pack(side=tk.LEFT)


        # 번역 완료 파일 경로 프레임
        self.translated_file_path_var = tk.StringVar()
        translated_file_frame = ttk.LabelFrame(left_panel, text="번역 완료 파일", padding=5)
        translated_file_frame.pack(padx=5, pady=5, fill=tk.X)
        self.translated_file_entry = ttk.Entry(translated_file_frame, textvariable=self.translated_file_path_var, state="readonly", width=60)
        self.translated_file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        self.open_folder_button = ttk.Button(translated_file_frame, text="폴더 열기", command=self.open_translated_folder, state=tk.DISABLED)
        self.open_folder_button.pack(side=tk.LEFT)


        # --- Right Panel (로그, 히스토리, 고급옵션 버튼) ---
        right_top_frame = ttk.Frame(right_panel) # 로그/히스토리용 노트북이 들어갈 프레임
        right_top_frame.pack(fill=tk.BOTH, expand=True) # 위쪽 공간 모두 차지

        # 고급 옵션 버튼 (팝업으로 변경)
        advanced_options_button = ttk.Button(
            right_panel, text="고급 옵션 설정...",
            command=self.open_advanced_options_popup
        )
        advanced_options_button.pack(fill=tk.X, padx=5, pady=(5,0), side=tk.BOTTOM) # 노트북 아래에 배치


        # 로그 및 히스토리 탭을 위한 Notebook 위젯
        right_panel_notebook = ttk.Notebook(right_top_frame) # 오른쪽 패널 상단에 위치
        right_panel_notebook.pack(fill=tk.BOTH, expand=True, pady=(0,0))


        # 실행 로그 탭
        log_tab_frame = ttk.Frame(right_panel_notebook, padding=5)
        right_panel_notebook.add(log_tab_frame, text="실행 로그")
        self.log_text = tk.Text(log_tab_frame, state=tk.DISABLED, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1, font=("TkFixedFont", 9))
        log_scrollbar_y = ttk.Scrollbar(log_tab_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_scrollbar_y.set)
        log_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 로깅 핸들러 설정 (Text 위젯으로 로그 출력)
        text_widget_handler = TextHandler(self.log_text)
        text_widget_handler.setFormatter(formatter)
        if not any(isinstance(h, TextHandler) for h in root_logger.handlers): # 중복 추가 방지
            root_logger.addHandler(text_widget_handler)


        # 번역 히스토리 탭
        history_tab_frame = ttk.Frame(right_panel_notebook, padding=5)
        right_panel_notebook.add(history_tab_frame, text="번역 히스토리")
        history_columns = ("name", "src", "tgt", "model", "ocr_temp", "status", "time", "path") # 컬럼 정의
        self.history_tree = ttk.Treeview(history_tab_frame, columns=history_columns, show="headings") # 헤더만 표시
        # 각 컬럼 설정
        self.history_tree.heading("name", text="문서 이름"); self.history_tree.column("name", width=120, anchor=tk.W, stretch=tk.YES)
        self.history_tree.heading("src", text="원본"); self.history_tree.column("src", width=50, anchor=tk.CENTER)
        self.history_tree.heading("tgt", text="대상"); self.history_tree.column("tgt", width=50, anchor=tk.CENTER)
        self.history_tree.heading("model", text="모델"); self.history_tree.column("model", width=100, anchor=tk.W)
        self.history_tree.heading("ocr_temp", text="OCR온도"); self.history_tree.column("ocr_temp", width=60, anchor=tk.CENTER)
        self.history_tree.heading("status", text="결과"); self.history_tree.column("status", width=60, anchor=tk.CENTER)
        self.history_tree.heading("time", text="번역일시"); self.history_tree.column("time", width=110, anchor=tk.CENTER)
        self.history_tree.heading("path", text="경로"); self.history_tree.column("path", width=0, stretch=tk.NO) # 경로는 숨김 (더블클릭 시 사용)

        hist_scrollbar_y = ttk.Scrollbar(history_tab_frame, orient="vertical", command=self.history_tree.yview)
        hist_scrollbar_x = ttk.Scrollbar(history_tab_frame, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=hist_scrollbar_y.set, xscrollcommand=hist_scrollbar_x.set)
        hist_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        hist_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.history_tree.pack(fill=tk.BOTH, expand=True)
        self.history_tree.bind("<Double-1>", self.on_history_double_click) # 더블클릭 이벤트 바인딩


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

    def update_ocr_status_display(self):
        selected_ui_lang = self.src_lang_var.get()
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name_display = "EasyOCR" if use_easyocr else "PaddleOCR"
        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get()
        gpu_status_text = "(GPU 사용 예정)" if gpu_enabled_for_ocr else "(CPU 사용 예정)"

        if self.ocr_handler and self.current_ocr_engine_type == engine_name_display.lower():
            current_handler_lang_display = ""
            if self.current_ocr_engine_type == "paddleocr" and hasattr(self.ocr_handler, 'current_lang_codes'):
                current_handler_lang_display = self.ocr_handler.current_lang_codes
            elif self.current_ocr_engine_type == "easyocr" and hasattr(self.ocr_handler, 'current_lang_codes') and self.ocr_handler.current_lang_codes:
                current_handler_lang_display = ", ".join(self.ocr_handler.current_lang_codes)

            gpu_in_use_text = "(GPU 사용 중)" if self.ocr_handler.use_gpu else "(CPU 사용 중)"
            self.ocr_status_label.config(text=f"{engine_name_display}: 준비됨 ({current_handler_lang_display}) {gpu_in_use_text}")
        else:
            ocr_lang_code_to_use = ""
            if use_easyocr:
                ocr_lang_code_to_use = config.UI_LANG_TO_EASYOCR_CODE_MAP.get(selected_ui_lang, "")
            else:
                ocr_lang_code_to_use = config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(selected_ui_lang, config.DEFAULT_PADDLE_OCR_LANG)
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

    def check_ocr_engine_status(self, is_called_from_start_translation=False):
        self.current_work_label.config(text="OCR 엔진 확인 중...")
        self.master.update_idletasks()

        selected_ui_lang = self.src_lang_var.get()
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name_display = "EasyOCR" if use_easyocr else "PaddleOCR"
        engine_name_internal = engine_name_display.lower()
        ocr_lang_code = None
        if use_easyocr:
            ocr_lang_code = config.UI_LANG_TO_EASYOCR_CODE_MAP.get(selected_ui_lang)
        else:
            ocr_lang_code = config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(selected_ui_lang, config.DEFAULT_PADDLE_OCR_LANG)

        if not ocr_lang_code:
            msg = f"{engine_name_display}: 언어 '{selected_ui_lang}'에 대한 OCR 코드가 설정되지 않았습니다."
            self.ocr_status_label.config(text=msg)
            logger.error(msg)
            if is_called_from_start_translation:
                messagebox.showerror("OCR 설정 오류", msg)
            self.current_work_label.config(text="OCR 설정 오류!")
            return False

        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get()
        needs_reinit = False
        if not self.ocr_handler: needs_reinit = True
        elif self.current_ocr_engine_type != engine_name_internal: needs_reinit = True
        elif self.ocr_handler.use_gpu != gpu_enabled_for_ocr: needs_reinit = True
        elif engine_name_internal == "paddleocr" and self.ocr_handler.current_lang_codes != ocr_lang_code: needs_reinit = True
        elif engine_name_internal == "easyocr" and (not self.ocr_handler.current_lang_codes or ocr_lang_code not in self.ocr_handler.current_lang_codes): needs_reinit = True

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
                            if utils.install_easyocr(): messagebox.showinfo(f"{engine_name_display} 설치 완료", f"{engine_name_display}이(가) 설치되었습니다. 애플리케이션을 재시작하거나 다시 시도해주세요.")
                            else: messagebox.showerror(f"{engine_name_display} 설치 실패", f"{engine_name_display} 설치에 실패했습니다.")
                        self.current_work_label.config(text=f"{engine_name_display} 미설치.")
                        return False
                    self.ocr_handler = EasyOcrHandler(lang_codes_list=[ocr_lang_code], debug_enabled=debug_mode, use_gpu=gpu_enabled_for_ocr)
                    self.current_ocr_engine_type = "easyocr"
                else:
                    if not utils.check_paddleocr():
                        self.ocr_status_label.config(text=f"{engine_name_display}: 미설치")
                        if messagebox.askyesno(f"{engine_name_display} 설치 필요", f"{engine_name_display}(paddlepaddle)이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
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

        self.update_ocr_status_display()
        if self.ocr_handler and self.ocr_handler.ocr_engine: return True
        else:
            self.ocr_status_label.config(text=f"{engine_name_display} OCR: 준비 안됨 ({selected_ui_lang})")
            if is_called_from_start_translation and not needs_reinit : messagebox.showwarning("OCR 오류", f"{engine_name_display} OCR 엔진을 사용할 수 없습니다. 이전 로그를 확인해주세요.")
            self.current_work_label.config(text=f"{engine_name_display} OCR 준비 안됨.")
            return False


    def swap_languages(self):
        src = self.src_lang_var.get()
        tgt = self.tgt_lang_var.get()
        self.src_lang_var.set(tgt)
        self.tgt_lang_var.set(src)
        logger.info(f"언어 스왑: {tgt} <-> {src}")
        self.on_source_language_change()

    def start_translation(self):
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "번역할 유효한 파워포인트 파일을 선택해주세요.\n'찾아보기' 버튼을 사용하여 파일을 선택할 수 있습니다.")
            return

        image_translation_really_enabled = self.image_translation_enabled_var.get()
        ocr_temperature_to_use = self.ocr_temperature_var.get()

        if image_translation_really_enabled:
            if not self.check_ocr_engine_status(is_called_from_start_translation=True):
                if not messagebox.askyesno("OCR 준비 실패", "이미지 내 텍스트 번역에 필요한 OCR 기능이 준비되지 않았거나 사용할 수 없습니다.\n이 경우 이미지 안의 글자는 번역되지 않습니다.\n계속 진행하시겠습니까? (텍스트/차트만 번역)"):
                    logger.warning("OCR 준비 실패로 사용자가 번역을 취소했습니다.")
                    self.current_work_label.config(text="번역 취소됨 (OCR 준비 실패).")
                    return
                logger.warning("OCR 핸들러 준비 실패. 이미지 번역 없이 진행합니다.")
                image_translation_really_enabled = False
        else:
            logger.info("이미지 번역 옵션이 꺼져있으므로 OCR 엔진을 확인하지 않습니다.")
            self._destroy_current_ocr_handler()

        src_lang, tgt_lang, model = self.src_lang_var.get(), self.tgt_lang_var.get(), self.model_var.get()
        if not model:
            messagebox.showerror("모델 오류", "번역 모델을 선택해주세요.\nOllama 서버가 실행 중이고 모델이 다운로드되었는지 확인하세요.\n'Ollama 확인' 버튼과 모델 목록 '🔄' 버튼을 사용해볼 수 있습니다.")
            self.check_ollama_status_manual()
            return
        if src_lang == tgt_lang:
            messagebox.showwarning("언어 동일", "원본 언어와 번역 언어가 동일합니다.\n다른 언어를 선택해주세요.")
            return

        ollama_running, _ = self.ollama_service.is_running()
        if not ollama_running:
            messagebox.showerror("Ollama 미실행", "Ollama 서버가 실행 중이지 않습니다.\nOllama를 실행한 후 'Ollama 확인' 버튼을 눌러주세요.")
            self.check_ollama_status_manual()
            return

        if self.total_weighted_work <= 0:
            logger.info("총 예상 작업량이 0입니다. 파일 정보를 다시 로드하여 확인합니다.")
            self.load_file_info(file_path)
            if self.total_weighted_work <= 0:
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
        ocr_temp_for_log = ocr_temperature_to_use if image_translation_really_enabled else 'N/A'
        ocr_gpu_for_log = self.ocr_use_gpu_var.get() if image_translation_really_enabled and self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu') and self.ocr_handler.use_gpu else 'N/A'

        logger.info(f"번역 시작: '{original_filename}' ({src_lang} -> {tgt_lang}) using {model}. "
                    f"이미지 번역: {'활성' if image_translation_really_enabled else '비활성'}, "
                    f"OCR 엔진: {ocr_engine_for_log}, OCR 온도: {ocr_temp_for_log}, OCR GPU: {ocr_gpu_for_log}")

        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.progress_label_var.set("0%")
        self.translated_file_path_var.set("")
        self.open_folder_button.config(state=tk.DISABLED)
        self.current_weighted_done = 0
        self.stop_event.clear()

        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("이미 번역 스레드가 실행 중입니다.")
            messagebox.showwarning("번역 중복", "이미 다른 번역 작업이 진행 중입니다.")
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            return

        self.current_work_label.config(text="번역 준비 중...")
        self.master.update_idletasks()

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
        temp_dir_for_pptx_handler_main = None # 여기에 초기화

        try:
            with open(task_log_filepath, 'a', encoding='utf-8') as f_log_init:
                f_log_init.write(f"--- 번역 작업 시작 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                f_log_init.write(f"원본 파일: {os.path.basename(file_path)}\n")
                f_log_init.write(f"원본 언어: {src_lang}, 대상 언어: {tgt_lang}, 번역 모델: {model}\n")
                f_log_init.write(f"이미지 번역 활성화: {image_translation_enabled}\n")
                if image_translation_enabled:
                    f_log_init.write(f"  OCR 엔진: {self.current_ocr_engine_type or '미지정'}\n")
                    f_log_init.write(f"  OCR 번역 온도: {ocr_temperature}\n")
                    gpu_in_use_log = 'N/A'
                    if self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu'): gpu_in_use_log = self.ocr_handler.use_gpu
                    f_log_init.write(f"  OCR GPU 사용 (실제): {gpu_in_use_log}\n")
                f_log_init.write(f"총 예상 가중 작업량: {self.total_weighted_work}\n")
                f_log_init.write("-" * 30 + "\n")
        except Exception as e_log_header: logger.error(f"작업 로그 파일 헤더 작성 실패: {e_log_header}")

        def report_item_completed_from_handler(slide_info_or_stage: Any, item_type_str: str, weighted_work_for_item: int, text_snippet_str: str):
            if self.stop_event.is_set(): return
            self.current_weighted_done += weighted_work_for_item
            self.current_weighted_done = min(self.current_weighted_done, self.total_weighted_work if self.total_weighted_work > 0 else weighted_work_for_item)
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self.update_translation_progress, slide_info_or_stage, item_type_str, self.current_weighted_done, self.total_weighted_work, text_snippet_str)
        try:
            if self.total_weighted_work == 0:
                logger.warning("번역할 가중 작업량이 없습니다.")
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     self.master.after(0, lambda: messagebox.showinfo("정보", "파일에 번역할 내용이 없습니다."))
                translation_result_status, output_path = "내용 없음", file_path
                with open(task_log_filepath, 'a', encoding='utf-8') as f_log_empty: f_log_empty.write(f"번역할 내용 없음. 원본 파일: {file_path}\n")
            else:
                font_code_for_render = config.UI_LANG_TO_FONT_CODE_MAP.get(tgt_lang, 'en')
                if hasattr(self, 'master') and self.master.winfo_exists():
                    self.master.after(0, lambda: self.current_work_label.config(text="파일 로드 중..."))
                    self.master.update_idletasks()

                temp_dir_for_pptx_handler_main = tempfile.mkdtemp(prefix="pptx_trans_main_") # 여기서 할당
                temp_pptx_for_chart_translation_path: Optional[str] = None

                try:
                    prs = Presentation(file_path)
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, lambda: self.current_work_label.config(text="1단계 (텍스트/이미지) 처리 시작..."))

                    stage1_success = self.pptx_handler.translate_presentation_stage1(
                        prs, src_lang, tgt_lang, self.translator,
                        self.ocr_handler if image_translation_enabled else None,
                        model, self.ollama_service, font_code_for_render, task_log_filepath,
                        report_item_completed_from_handler, self.stop_event,
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
                        except Exception as e_save_stop: logger.error(f"1단계 중단 후 저장 실패: {e_save_stop}"); output_path = file_path
                    elif not stage1_success:
                        logger.error("1단계 번역 실패."); translation_result_status = "실패 (1단계 오류)"; output_path = file_path
                    else:
                        logger.info("번역 작업자: 1단계 완료. 임시 파일 저장 시도.")
                        if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, lambda: self.current_work_label.config(text="1단계 완료. 임시 파일 저장 중...")); self.master.update_idletasks()
                        temp_pptx_for_chart_translation_path = os.path.join(temp_dir_for_pptx_handler_main, f"{os.path.splitext(os.path.basename(file_path))[0]}_temp_for_charts.pptx")
                        if prs: prs.save(temp_pptx_for_chart_translation_path)
                        logger.info(f"1단계 결과 임시 저장: {temp_pptx_for_chart_translation_path}")
                        info_for_charts = self.pptx_handler.get_file_info(temp_pptx_for_chart_translation_path)
                        num_charts_in_prs = info_for_charts.get('chart_elements_count', 0)

                        if num_charts_in_prs > 0 and not self.stop_event.is_set():
                            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, lambda: self.current_work_label.config(text=f"2단계 (차트) 처리 시작 ({num_charts_in_prs}개)...")); self.master.update_idletasks()
                            logger.info(f"번역 작업자: 2단계 (차트) 시작. 대상 차트 수: {num_charts_in_prs}")
                            safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                            final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                            final_output_dir = os.path.dirname(file_path)
                            final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)
                            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, lambda: self.current_work_label.config(text="2단계: 차트 XML 압축 해제 중...")); self.master.update_idletasks()
                            output_path_charts = self.chart_xml_handler.translate_charts_in_pptx(
                                pptx_path=temp_pptx_for_chart_translation_path, src_lang_ui_name=src_lang, tgt_lang_ui_name=tgt_lang, model_name=model,
                                output_path=final_pptx_output_path, progress_callback_item_completed=report_item_completed_from_handler,
                                stop_event=self.stop_event, task_log_filepath=task_log_filepath
                            )
                            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, lambda: self.current_work_label.config(text="2단계: 번역된 차트 XML 압축 중...")); self.master.update_idletasks()
                            if self.stop_event.is_set():
                                logger.warning("2단계 차트 번역 중 또는 완료 직후 중지됨."); translation_result_status = "부분 성공 (중지)"
                                output_path = output_path_charts if (output_path_charts and os.path.exists(output_path_charts)) else temp_pptx_for_chart_translation_path
                            elif output_path_charts and os.path.exists(output_path_charts):
                                logger.info(f"2단계 차트 번역 완료. 최종 파일: {output_path_charts}"); translation_result_status = "성공"; output_path = output_path_charts
                            else:
                                logger.error("2단계 차트 번역 실패 또는 결과 파일 없음. 1단계 결과물 사용 시도."); translation_result_status = "실패 (2단계 오류)"
                                if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                                    try: shutil.copy2(temp_pptx_for_chart_translation_path, final_pptx_output_path); output_path = final_pptx_output_path; logger.info(f"차트 번역 실패로 1단계 결과물을 최종 경로에 복사: {output_path}")
                                    except Exception as e_copy_fallback: logger.error(f"차트 번역 실패 후 1단계 결과물 복사 중 오류: {e_copy_fallback}."); output_path = temp_pptx_for_chart_translation_path
                                else: output_path = file_path
                        elif self.stop_event.is_set():
                            logger.info("1단계 후 중단되어 차트 번역은 실행되지 않음."); translation_result_status = "부분 성공 (중지)"; output_path = temp_pptx_for_chart_translation_path
                        else:
                            logger.info("번역할 차트가 없습니다. 1단계 결과물을 최종 결과로 사용합니다.")
                            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, lambda: self.current_work_label.config(text="최종 파일 저장 중...")); self.master.update_idletasks()
                            safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                            final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                            final_output_dir = os.path.dirname(file_path)
                            final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)
                            try:
                                if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                                    shutil.copy2(temp_pptx_for_chart_translation_path, final_pptx_output_path); output_path = final_pptx_output_path; translation_result_status = "성공"; logger.info(f"차트 없음. 최종 파일 저장: {output_path}")
                                else: logger.error("차트가 없고, 1단계 임시 파일도 찾을 수 없습니다."); translation_result_status = "실패 (파일 오류)"; output_path = file_path
                            except Exception as e_copy_no_chart: logger.error(f"차트 없는 경우 최종 파일 복사 중 오류: {e_copy_no_chart}"); translation_result_status = "실패 (파일 복사 오류)"; output_path = temp_pptx_for_chart_translation_path if temp_pptx_for_chart_translation_path else file_path
                finally:
                    if temp_dir_for_pptx_handler_main and os.path.exists(temp_dir_for_pptx_handler_main): # finally 블록 전에 temp_dir_for_pptx_handler_main이 할당되었는지 확인
                        try: shutil.rmtree(temp_dir_for_pptx_handler_main); logger.debug(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 완료.")
                        except Exception as e_clean_main_dir: logger.warning(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 중 오류: {e_clean_main_dir}")

            if translation_result_status == "성공" and not self.stop_event.is_set():
                 self.current_weighted_done = self.total_weighted_work
                 if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, self.update_translation_progress, "완료", "번역 완료됨", self.current_weighted_done, self.total_weighted_work, "최종 저장 완료")
                 if not (output_path and os.path.exists(output_path)):
                     logger.error(f"번역 '성공'으로 기록되었으나, 최종 결과 파일({output_path})을 찾을 수 없습니다."); translation_result_status = "실패 (결과 파일 없음)"; output_path = file_path
                 else:
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(100, lambda: self._ask_open_folder(output_path))
            elif "실패" in translation_result_status or "오류" in translation_result_status:
                 if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath)
                 if not output_path: output_path = file_path
        except Exception as e_worker:
            logger.error(f"번역 작업 중 심각한 오류 발생: {e_worker}", exc_info=True)
            translation_result_status = "치명적 오류 발생"; output_path = output_path or file_path
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_err: f_err.write(f"\n--- 번역 작업 중 심각한 오류 발생 ---\n오류: {e_worker}\n{traceback.format_exc()}");
            except Exception as ef_log: logger.error(f"작업 로그 파일에 오류 기록 실패: {ef_log}")
            if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath, str(e_worker))
        finally:
            if hasattr(self, 'master') and self.master.winfo_exists():
                history_entry = {
                    "name": os.path.basename(file_path), "src": src_lang, "tgt": tgt_lang, "model": model,
                    "ocr_temp": ocr_temperature if image_translation_enabled else "N/A",
                    "ocr_gpu": self.ocr_use_gpu_var.get() if image_translation_enabled and self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu') else "N/A",
                    "img_trans_enabled": image_translation_enabled, "status": translation_result_status,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "path": output_path or file_path,
                    "log_file": task_log_filepath
                }
                self.master.after(0, self.translation_finished, history_entry)
            self.translation_thread = None


    def _handle_translation_failure(self, status, original_file, log_file, error_details=""):
        logger.error(f"번역 실패: {status}, 원본: {original_file}, 로그: {log_file}, 상세: {error_details}")
        if hasattr(self, 'current_work_label') and self.current_work_label.winfo_exists():
            self.current_work_label.config(text=f"번역 실패: {status}")
        error_title = f"번역 작업 실패 ({status})"
        user_message = f"'{os.path.basename(original_file)}' 파일 번역 중 오류가 발생했습니다.\n\n상태: {status}\n"
        if error_details: user_message += f"오류 정보: {error_details[:200]}...\n\n"
        user_message += "다음 사항을 확인해 보세요:\n- Ollama 서버가 정상적으로 실행 중인지 ('Ollama 확인' 버튼)\n- 선택한 번역 모델이 유효한지 (모델 목록 '🔄' 버튼)\n- 원본 파일이 손상되지 않았는지\n"
        if "GPU" in status.upper() or "CUDA" in status.upper() or "메모리 부족" in status or \
           (self.ocr_use_gpu_var.get() and ("OCR" in status.upper() or "엔진" in status)):
            user_message += "- 고급 옵션에서 'GPU 사용'을 해제하고 다시 시도해보세요.\n"
        user_message += f"\n자세한 내용은 로그 파일에서 확인할 수 있습니다.\n로그 파일: {log_file}"
        if messagebox.askyesno(error_title, user_message + "\n\n오류 로그가 저장된 폴더를 여시겠습니까?", icon='error'):
            try: utils.open_folder(os.path.dirname(log_file))
            except Exception as e_open_log_dir: logger.warning(f"로그 폴더 열기 실패: {e_open_log_dir}"); messagebox.showinfo("정보", f"로그 폴더를 열 수 없습니다.\n경로: {os.path.dirname(log_file)}")


    def _ask_open_folder(self, path):
        if path and os.path.exists(path):
            user_choice = messagebox.askyesnocancel("번역 완료", f"번역이 완료되었습니다.\n저장된 파일: {os.path.basename(path)}\n\n결과 파일이 저장된 폴더를 여시겠습니까?", icon='info', default=messagebox.YES)
            if user_choice is True: utils.open_folder(os.path.dirname(path))


    def update_translation_progress(self, current_location_info: Any, current_task_type: str,
                                    current_total_weighted_done: int, total_weighted_overall: int,
                                    current_text_snippet: str = ""):
        if self.stop_event.is_set(): return
        progress = 0
        if total_weighted_overall > 0: progress = (current_total_weighted_done / total_weighted_overall) * 100
        elif current_total_weighted_done == 0 and total_weighted_overall == 0 : progress = 100
        progress = min(max(0, progress), 100); progress_text_val = f"{progress:.1f}%"
        task_description = current_task_type
        location_display_text = str(current_location_info)
        if isinstance(current_location_info, (int, float)):
            location_display_text = f"슬라이드 {int(current_location_info)} / {self.current_file_slide_count}"
            if "텍스트" in task_description: task_description = "1단계: 텍스트 요소 번역"
            elif "이미지" in task_description: task_description = "1단계: 이미지 처리"
            elif "표" in task_description: task_description = "1단계: 표 내부 텍스트 번역"
            else: task_description = f"1단계: {task_description}"
        elif not current_location_info or str(current_location_info).upper() == "N/A":
            location_display_text = "전체 파일 처리"
            if "차트" in task_description or "chart" in task_description.lower(): task_description = f"2단계: {task_description}"
        elif str(current_location_info).lower() == "완료": location_display_text = "모든 슬라이드 완료"; task_description = "번역 완료됨"
        snippet_display = current_text_snippet.replace('\n', ' ').strip();
        if len(snippet_display) > 25: snippet_display = snippet_display[:22] + "..."
        def _update_ui():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            self.progress_bar["value"] = progress; self.progress_label_var.set(progress_text_val)
            self.current_slide_label.config(text=f"현재 위치: {location_display_text}")
            self.current_work_label.config(text=f"현재 작업: {task_description} - '{snippet_display}'")
        if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(0, _update_ui)


    def update_progress_timer(self):
        if self.translation_thread and self.translation_thread.is_alive() and not self.stop_event.is_set():
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(1000, self.update_progress_timer)


    def stop_translation(self):
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("번역 중지 요청 중..."); self.stop_event.set(); self.stop_button.config(state=tk.DISABLED); self.current_work_label.config(text="번역 중지 요청됨...")
        elif self.model_download_thread and self.model_download_thread.is_alive():
            logger.warning("모델 다운로드 중지 요청 중..."); self.stop_event.set(); self.stop_button.config(state=tk.DISABLED)


    def translation_finished(self, history_entry: Dict[str, Any]):
        if not (hasattr(self, 'start_button') and self.start_button.winfo_exists()): return
        self.start_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED)
        result_status = history_entry.get("status", "알 수 없음"); translated_file_path = history_entry.get("path"); current_progress_val = self.progress_bar["value"]
        if result_status == "성공" and not self.stop_event.is_set(): final_progress_text = "100%"; self.progress_bar["value"] = 100; self.current_work_label.config(text=f"번역 완료: {os.path.basename(translated_file_path) if translated_file_path else '파일 없음'}"); self.current_slide_label.config(text="모든 작업 완료")
        elif "중지" in result_status: final_progress_text = f"{current_progress_val:.1f}% (중지됨)"; self.current_work_label.config(text="번역 중지됨.")
        elif result_status == "내용 없음": final_progress_text = "100% (내용 없음)"; self.progress_bar["value"] = 100; self.current_work_label.config(text="번역할 내용 없음.")
        else: final_progress_text = f"{current_progress_val:.1f}% ({result_status})"
        self.progress_label_var.set(final_progress_text)
        if translated_file_path and os.path.exists(translated_file_path) and result_status == "성공": self.translated_file_path_var.set(translated_file_path); self.open_folder_button.config(state=tk.NORMAL)
        else: self.translated_file_path_var.set("번역 실패 또는 파일 없음"); self.open_folder_button.config(state=tk.DISABLED);
        if result_status == "성공" and not (translated_file_path and os.path.exists(translated_file_path)): logger.warning(f"번역은 '성공'으로 기록되었으나, 결과 파일 경로가 유효하지 않음: {translated_file_path}")
        self._add_history_entry(history_entry)
        task_log_filepath = history_entry.get("log_file")
        if task_log_filepath and os.path.exists(os.path.dirname(task_log_filepath)):
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    f_task_log.write(f"\n--- 번역 작업 최종 상태 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                    f_task_log.write(f"최종 상태: {result_status}\n")
                    if self.file_path_var.get(): f_task_log.write(f"원본 파일 (UI 경로): {self.file_path_var.get()}\n")
                    if translated_file_path and os.path.exists(translated_file_path): f_task_log.write(f"번역된 파일: {translated_file_path}\n")
                    elapsed_time_for_log = (time.time() - self.start_time) if self.start_time else 0; m, s = divmod(elapsed_time_for_log, 60)
                    f_task_log.write(f"총 소요 시간 (내부 기록용): {int(m):02d}분 {s:05.2f}초\n"); f_task_log.write("-" * 30 + "\n")
            except Exception as e_log_finish: logger.error(f"작업 로그 파일에 최종 상태 기록 실패: {e_log_finish}")
        self.start_time = None


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
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
        try:
            if self.text_widget.winfo_exists(): self.text_widget.after(0, append_message)
        except tk.TclError: pass


if __name__ == "__main__":
    for dir_path in [LOGS_DIR, FONTS_DIR, ASSETS_DIR, HISTORY_DIR, os.path.dirname(USER_SETTINGS_PATH)]:
        try:
            if dir_path: os.makedirs(dir_path, exist_ok=True)
        except Exception as e_mkdir_main: print(f"디렉토리 생성 실패 ({dir_path}): {e_mkdir_main}")

    if debug_mode: logger.info("디버그 모드로 실행 중입니다.")
    else: logger.info("일반 모드로 실행 중입니다.")

    if not os.path.exists(config.FONTS_DIR) or not os.listdir(config.FONTS_DIR):
        logger.critical(f"필수 폰트 디렉토리({config.FONTS_DIR})를 찾을 수 없거나 비어있습니다. 애플리케이션이 정상 동작하지 않을 수 있습니다.")
    else: logger.info(f"폰트 디렉토리 확인: {config.FONTS_DIR}")
    if not os.path.exists(config.ASSETS_DIR): logger.warning(f"에셋 디렉토리를 찾을 수 없습니다: {config.ASSETS_DIR}")
    else: logger.info(f"에셋 디렉토리 확인: {config.ASSETS_DIR}")

    root = tk.Tk()
    app = Application(master=root)
    root.geometry("1024x768")
    root.update_idletasks()
    min_width = root.winfo_reqwidth()
    min_height = root.winfo_reqheight()
    root.minsize(min_width + 20, min_height + 20)

    try:
        root.mainloop()
    except KeyboardInterrupt: logger.info("Ctrl+C로 애플리케이션 종료 중...")
    finally: logger.info(f"--- {APP_NAME} 종료됨 (mainloop 이후) ---")