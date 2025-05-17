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

# ... (로깅 설정, 경로 설정 등은 기존과 유사) ...
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
USER_SETTINGS_PATH = os.path.join(BASE_DIR_MAIN, config.USER_SETTINGS_FILENAME) # --- 1단계 개선: 사용자 설정 파일 경로 ---


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
        self.general_file_handler = None
        self._setup_logging_file_handler() # 로깅 핸들러 먼저 설정
        
        self.user_settings: Dict[str, Any] = {} # --- 1단계 개선: 사용자 설정 저장용 딕셔너리 ---
        self._load_user_settings() # --- 1단계 개선: 사용자 설정 로드 ---

        # 서비스/핸들러 인스턴스 생성 (2단계에서 인터페이스 기반 주입으로 변경 예정)
        self.ollama_service = OllamaService()
        self.translator = OllamaTranslator()
        self.pptx_handler = PptxHandler()
        self.chart_xml_handler = ChartXmlHandler(self.translator, self.ollama_service)
        self.ocr_handler = None # 동적으로 생성
        self.current_ocr_engine_type = None # 현재 사용 중인 OCR 엔진 ("paddleocr" 또는 "easyocr")

        # ... (아이콘, 스타일 설정 등 기존 코드 유지) ...
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


        # --- 1단계 개선: 고급 옵션 변수 초기화 시 사용자 설정 또는 기본값 사용 ---
        self.ocr_temperature_var = tk.DoubleVar(
            value=self.user_settings.get("ocr_temperature", config.DEFAULT_ADVANCED_SETTINGS["ocr_temperature"])
        )
        self.image_translation_enabled_var = tk.BooleanVar(
            value=self.user_settings.get("image_translation_enabled", config.DEFAULT_ADVANCED_SETTINGS["image_translation_enabled"])
        )
        self.ocr_use_gpu_var = tk.BooleanVar(
            value=self.user_settings.get("ocr_use_gpu", config.DEFAULT_ADVANCED_SETTINGS["ocr_use_gpu"])
        )
        # --- 1단계 개선 끝 ---

        self.create_widgets()
        self._load_translation_history() # 번역 히스토리 로드
        self.master.after(100, self.initial_checks) # 초기 상태 점검 (Ollama, OCR 등)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing) # 종료 시 처리
        atexit.register(self.on_closing) # 비정상 종료 시에도 호출되도록

        log_file_path_msg = self.general_file_handler.baseFilename if self.general_file_handler else '미설정'
        logger.info(f"--- {APP_NAME} 시작됨 (일반 로그 파일: {log_file_path_msg}) ---")
        logger.info(f"로드된 사용자 설정: {self.user_settings}")

    def _setup_logging_file_handler(self):
        # ... (기존과 동일) ...
        if self.general_file_handler: return
        try:
            os.makedirs(LOGS_DIR, exist_ok=True) # 로그 디렉토리 생성
            general_log_filename = os.path.join(LOGS_DIR, "app_general.log")
            self.general_file_handler = logging.FileHandler(general_log_filename, mode='a', encoding='utf-8')
            self.general_file_handler.setFormatter(formatter)
            # 핸들러 중복 추가 방지
            if not any(h.baseFilename == os.path.abspath(general_log_filename) for h in root_logger.handlers if isinstance(h, logging.FileHandler)):
                root_logger.addHandler(self.general_file_handler)
        except Exception as e:
            # 이 시점에서는 logger가 완전히 설정되지 않았을 수 있으므로 print 사용
            print(f"일반 로그 파일 핸들러 설정 실패: {e}")


    # --- 1단계 개선: 사용자 설정 로드 및 저장 메소드 추가 ---
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
            os.makedirs(os.path.dirname(USER_SETTINGS_PATH), exist_ok=True) # 설정 파일 디렉토리 생성
            with open(USER_SETTINGS_PATH, 'w', encoding='utf-8') as f:
                json.dump(settings_to_save, f, ensure_ascii=False, indent=4)
            logger.info(f"사용자 설정 저장 완료: {USER_SETTINGS_PATH}")
            self.user_settings = settings_to_save # 저장 후 내부 상태도 업데이트
        except Exception as e:
            logger.error(f"사용자 설정 저장 중 오류: {e}", exc_info=True)
    # --- 1단계 개선 끝 ---

    def _destroy_current_ocr_handler(self):
        # ... (기존과 동일) ...
        if self.ocr_handler:
            logger.info(f"기존 OCR 핸들러 ({self.current_ocr_engine_type}) 자원 해제 시도...")
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
        # --- 1단계 개선: 종료 시 사용자 설정 저장 ---
        self._save_user_settings()
        # --- 1단계 개선 끝 ---

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
                    self.general_file_handler = None
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
            # master가 없거나 이미 destroy된 경우
            logger.info("애플리케이션 윈도우가 이미 없으므로 바로 종료합니다.")
        
        # atexit에 등록된 경우, 이 함수가 다시 호출될 수 있으므로 sys.exit()는 신중히 사용
        # 여기서는 master.destroy() 후 mainloop가 자연스럽게 종료되도록 함


    def initial_checks(self):
        # ... (기존과 동일) ...
        logger.debug("초기 점검 시작: OCR 라이브러리 설치 여부 및 Ollama 상태 확인")
        self.update_ocr_status_display() # OCR 상태 표시 업데이트
        self.check_ollama_status_manual(initial_check=True) # Ollama 서버 상태 확인
        logger.debug("초기 점검 완료.")

    def create_widgets(self):
        # ... (기존 위젯 생성 코드와 대부분 동일) ...
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
        right_panel_notebook.pack(fill=tk.BOTH, expand=True, pady=(0,0)) # ipady 제거, pady 조정


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
                # Pillow를 사용하여 이미지 크기 얻고, Tkinter PhotoImage로 로드 시 크기 조절
                pil_temp_for_size = Image.open(logo_path_bottom)
                original_width, original_height = pil_temp_for_size.size
                pil_temp_for_size.close() # 이미지 파일 핸들 닫기

                # 목표 높이에 맞춰 subsample 계수 계산 (너무 작아지지 않도록 최소 1)
                target_height_bottom = 20 # 하단 로고 목표 높이
                # subsample_factor는 정수여야 함
                subsample_factor = max(1, int(original_height / target_height_bottom)) if original_height > target_height_bottom and target_height_bottom > 0 else (1 if original_height > 0 else 6) # 0으로 나누는 것 방지 및 기본값

                # PhotoImage는 master 인자 필요할 수 있음 (Tk 윈도우 파괴 시 관련 오류 방지)
                temp_logo_image_bottom = tk.PhotoImage(file=logo_path_bottom, master=self.master)
                self.logo_image_tk_bottom = temp_logo_image_bottom.subsample(subsample_factor, subsample_factor)
                logo_label_bottom = ttk.Label(bottom_frame, image=self.logo_image_tk_bottom)
                logo_label_bottom.pack(side=tk.RIGHT, padx=10, pady=2)
            # except tk.TclError as e_logo_tk: # PhotoImage 관련 오류
            #     logger.warning(f"하단 로고 로드 중 Tkinter 오류: {e_logo_tk}. Pillow 대체 시도 안 함 (subsample 문제일 수 있음).")
            except Exception as e_general_bottom: # 기타 모든 예외
                logger.warning(f"하단 로고 로드 중 예외: {e_general_bottom}", exc_info=True)
        else:
            logger.warning(f"하단 로고 파일({logo_path_bottom})을 찾을 수 없습니다.")


    # --- 1단계 개선: 고급 옵션 팝업에서 사용자 설정/기본값 사용 ---
    def open_advanced_options_popup(self):
        popup = tk.Toplevel(self.master)
        popup.title("고급 옵션")
        popup.geometry("450x280") # 팝업 크기
        popup.resizable(False, False) # 크기 조절 불가
        popup.transient(self.master) # 부모 창 위에 항상 표시
        popup.grab_set() # 팝업이 떠 있는 동안 다른 창 비활성화

        # 현재 설정값을 임시 변수에 저장 (취소 시 복원 위함이 아니라, 팝업 내에서만 사용)
        # self.ocr_temperature_var 등은 이미 __init__에서 사용자 설정/기본값으로 초기화됨
        temp_ocr_temp_var = tk.DoubleVar(value=self.ocr_temperature_var.get())
        temp_img_trans_enabled_var = tk.BooleanVar(value=self.image_translation_enabled_var.get())
        temp_ocr_gpu_var = tk.BooleanVar(value=self.ocr_use_gpu_var.get())

        main_frame = ttk.Frame(popup, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # OCR 온도 설정 프레임
        temp_label_frame = ttk.LabelFrame(main_frame, text="이미지 번역 온도 설정", padding=10)
        temp_label_frame.pack(fill=tk.X, pady=5)

        temp_frame_inner = ttk.Frame(temp_label_frame) # 슬라이더와 값 표시 레이블을 위한 내부 프레임
        temp_frame_inner.pack(fill=tk.X, pady=2)

        temp_current_value_label = ttk.Label(temp_frame_inner, text=f"{temp_ocr_temp_var.get():.1f}") # 초기값 표시

        # 슬라이더 값 변경 시 레이블 업데이트 함수
        def _update_popup_temp_label(value_str): # ttk.Scale의 command는 문자열 값을 전달
            try:
                value = float(value_str)
                if temp_current_value_label.winfo_exists(): # 위젯 존재 확인
                    temp_current_value_label.config(text=f"{value:.1f}")
            except ValueError: pass # float 변환 실패 시 무시
            except tk.TclError: pass # 위젯 파괴 후 호출 시 오류 방지

        ocr_temp_slider_popup = ttk.Scale(
            temp_frame_inner, from_=0.1, to=1.0, variable=temp_ocr_temp_var,
            orient=tk.HORIZONTAL, command=_update_popup_temp_label
        )
        ocr_temp_slider_popup.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0,5))
        temp_current_value_label.pack(side=tk.LEFT) # 슬라이더 오른쪽에 값 표시


        # 온도 설명 레이블
        temp_description_frame = ttk.Frame(temp_label_frame)
        temp_description_frame.pack(fill=tk.X, pady=(0,5))
        ttk.Label(temp_description_frame, text="0.1 (정직함) <----------------------> 1.0 (창의적)", justify=tk.CENTER).pack(fill=tk.X)
        ttk.Label(temp_description_frame, text="(기본값: 0.4, 이미지 품질이 좋지 않을 경우 수치를 올리는 것이 번역에 도움 될 수 있음)", wraplength=400, justify=tk.LEFT, font=("TkDefaultFont",8)).pack(fill=tk.X)


        # 체크박스 프레임
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

        # 버튼 프레임 (적용, 취소)
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20,0), side=tk.BOTTOM) # 하단에 배치

        def apply_settings():
            # 임시 변수의 값을 실제 설정 변수에 반영
            self.ocr_temperature_var.set(temp_ocr_temp_var.get())
            self.image_translation_enabled_var.set(temp_img_trans_enabled_var.get())

            gpu_setting_changed = self.ocr_use_gpu_var.get() != temp_ocr_gpu_var.get()
            self.ocr_use_gpu_var.set(temp_ocr_gpu_var.get())

            logger.info(f"고급 옵션 적용: 온도={self.ocr_temperature_var.get()}, 이미지번역={self.image_translation_enabled_var.get()}, OCR GPU={self.ocr_use_gpu_var.get()}")
            
            # --- 1단계 개선: 설정 변경 시 즉시 저장 ---
            self._save_user_settings() 
            
            if gpu_setting_changed:
                logger.info("OCR GPU 사용 설정 변경됨. 다음 번역 시 또는 OCR 상태 확인 시 적용됩니다.")
                self._destroy_current_ocr_handler() # GPU 설정 변경 시 기존 OCR 핸들러 해제
                self.update_ocr_status_display() # OCR 상태 표시 업데이트

            if popup.winfo_exists(): popup.destroy()

        def cancel_settings():
            if popup.winfo_exists(): popup.destroy()

        apply_button = ttk.Button(button_frame, text="적용", command=apply_settings)
        apply_button.pack(side=tk.RIGHT, padx=5)
        cancel_button = ttk.Button(button_frame, text="취소", command=cancel_settings)
        cancel_button.pack(side=tk.RIGHT)

        popup.wait_window() # 팝업이 닫힐 때까지 대기
    # --- 1단계 개선 끝 ---


    def _load_translation_history(self):
        # ... (기존과 동일) ...
        if not os.path.exists(HISTORY_DIR):
            try: os.makedirs(HISTORY_DIR, exist_ok=True)
            except Exception as e_mkdir:
                logger.error(f"히스토리 디렉토리({HISTORY_DIR}) 생성 실패: {e_mkdir}")
                self.translation_history_data = []
                return

        if os.path.exists(self.history_file_path):
            try:
                with open(self.history_file_path, 'r', encoding='utf-8') as f:
                    self.translation_history_data = json.load(f)
                # 시간순 정렬 (최신이 위로) 및 최대 개수 제한
                self.translation_history_data.sort(key=lambda x: x.get('time', '0'), reverse=True)
                self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
            except json.JSONDecodeError:
                logger.error(f"번역 히스토리 파일({self.history_file_path}) 디코딩 오류. 새 히스토리 시작.")
                self.translation_history_data = []
            except Exception as e:
                logger.error(f"번역 히스토리 로드 중 오류: {e}", exc_info=True)
                self.translation_history_data = []
        else:
            self.translation_history_data = [] # 파일 없으면 빈 리스트로 시작
        self._populate_history_treeview()


    def _save_translation_history(self):
        # ... (기존과 동일) ...
        try:
            os.makedirs(HISTORY_DIR, exist_ok=True) # 히스토리 디렉토리 생성 (없을 경우 대비)
            # 시간순 정렬 (최신이 위로) 및 최대 개수 제한 (저장 직전에 한 번 더)
            self.translation_history_data.sort(key=lambda x: x.get('time', '0'), reverse=True)
            self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
            with open(self.history_file_path, 'w', encoding='utf-8') as f:
                json.dump(self.translation_history_data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logger.error(f"번역 히스토리 저장 중 오류: {e}", exc_info=True)


    def _add_history_entry(self, entry: Dict[str, Any]):
        # ... (기존과 동일) ...
        self.translation_history_data.insert(0, entry) # 새 항목을 맨 앞에 추가
        # 최대 개수 유지 (정렬은 _save_translation_history 또는 _load_translation_history에서 담당)
        self.translation_history_data = self.translation_history_data[:config.MAX_HISTORY_ITEMS]
        self._save_translation_history() # 변경 시마다 저장
        self._populate_history_treeview() # Treeview 업데이트


    def _populate_history_treeview(self):
        # ... (기존과 동일) ...
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()):
            return # Treeview 위젯이 없으면 아무것도 안 함
        # 기존 항목 모두 삭제
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        # 새 히스토리 데이터로 채우기
        for entry in self.translation_history_data:
            values = (
                entry.get("name", "-"),
                entry.get("src", "-"),
                entry.get("tgt", "-"),
                entry.get("model", "-"),
                f"{entry.get('ocr_temp', '-')}", # OCR 온도는 문자열로 표시
                entry.get("status", "-"),
                entry.get("time", "-"),
                entry.get("path", "-") # 경로는 숨겨져 있지만 값은 유지
            )
            self.history_tree.insert("", tk.END, values=values)
        if self.translation_history_data: # 데이터가 있으면 맨 위로 스크롤
            self.history_tree.yview_moveto(0)

    def update_ocr_status_display(self):
        # ... (기존과 동일) ...
        selected_ui_lang = self.src_lang_var.get() # 현재 선택된 원본 언어
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name_display = "EasyOCR" if use_easyocr else "PaddleOCR"

        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get() # 현재 GPU 사용 설정
        gpu_status_text = "(GPU 사용 예정)" if gpu_enabled_for_ocr else "(CPU 사용 예정)"

        # OCR 핸들러가 이미 초기화되었고, 현재 설정과 일치하는 경우
        if self.ocr_handler and self.current_ocr_engine_type == engine_name_display.lower():
            current_handler_lang_display = ""
            if self.current_ocr_engine_type == "paddleocr" and hasattr(self.ocr_handler, 'current_lang_codes'):
                current_handler_lang_display = self.ocr_handler.current_lang_codes # Paddle은 단일 코드
            elif self.current_ocr_engine_type == "easyocr" and hasattr(self.ocr_handler, 'current_lang_codes') and self.ocr_handler.current_lang_codes:
                current_handler_lang_display = ", ".join(self.ocr_handler.current_lang_codes) # EasyOCR은 리스트

            gpu_in_use_text = "(GPU 사용 중)" if self.ocr_handler.use_gpu else "(CPU 사용 중)"
            self.ocr_status_label.config(text=f"{engine_name_display}: 준비됨 ({current_handler_lang_display}) {gpu_in_use_text}")
        else: # OCR 핸들러가 없거나, 설정이 변경되어 재초기화가 필요한 경우
            ocr_lang_code_to_use = ""
            if use_easyocr:
                ocr_lang_code_to_use = config.UI_LANG_TO_EASYOCR_CODE_MAP.get(selected_ui_lang, "")
            else:
                ocr_lang_code_to_use = config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(selected_ui_lang, config.DEFAULT_PADDLE_OCR_LANG)

            self.ocr_status_label.config(text=f"{engine_name_display}: ({ocr_lang_code_to_use or selected_ui_lang}) 사용 예정 {gpu_status_text} (미확인)")


    def on_source_language_change(self, event=None):
        # ... (기존과 동일) ...
        selected_ui_lang = self.src_lang_var.get()
        logger.info(f"원본 언어 변경됨: {selected_ui_lang}.")
        self.update_ocr_status_display() # OCR 상태 표시 업데이트
        # 파일이 선택되어 있다면, 해당 파일 정보 다시 로드 (필요시, 현재는 불필요해 보임)
        # if self.file_path_var.get():
        #     self.load_file_info(self.file_path_var.get())


    def browse_file(self):
        # ... (기존과 동일) ...
        file_path = filedialog.askopenfilename(title="파워포인트 파일 선택", filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*")))
        if file_path:
            self.file_path_var.set(file_path)
            logger.info(f"파일 선택됨: {file_path}")
            self.load_file_info(file_path) # 파일 정보 로드
            self.translated_file_path_var.set("") # 이전 번역 완료 경로 초기화
            self.open_folder_button.config(state=tk.DISABLED) # 폴더 열기 버튼 비활성화
            self.current_work_label.config(text="파일 선택됨. 번역 대기 중.")

    def load_file_info(self, file_path):
        # ... (기존과 동일, get_file_info의 반환값 사용) ...
        self.current_work_label.config(text="파일 분석 중...")
        self.master.update_idletasks() # UI 즉시 업데이트
        
        # --- 1단계 개선: PptxHandler.get_file_info의 반환값 형식 일관성 가정 ---
        # info = self.pptx_handler.get_file_info(file_path) -> 반환값은 Dict[str, int]
        # 오류 발생 시에도 동일한 키를 가지되, 값은 0 또는 음수 등으로 처리하는 것을 가정
        try:
            logger.debug(f"파일 정보 분석 중: {file_path}")
            file_name = os.path.basename(file_path)
            
            # PptxHandler의 get_file_info 호출
            info = self.pptx_handler.get_file_info(file_path)
            
            # --- 1단계 개선: get_file_info가 오류 시에도 딕셔너리 반환 가정, 오류 처리는 내부에서 ---
            if info.get("slide_count", -1) == -1 and info.get("total_text_char_count", -1) == -1 : # 예시: 오류 발생 시 특정 값으로 반환
                # get_file_info 내부에서 이미 오류 로깅 및 사용자 알림 처리했다고 가정
                # 여기서는 UI만 초기화
                self.file_name_label.config(text=f"파일 이름: {file_name} (분석 오류)")
                self.slide_count_label.config(text="슬라이드 수: -")
                self.total_text_char_label.config(text="텍스트 글자 수: -")
                self.image_elements_label.config(text="이미지 수: -")
                self.chart_elements_label.config(text="차트 수: -")
                self.total_weighted_work = 0
                self.current_work_label.config(text="파일 분석 실패!")
                # messagebox.showerror("파일 분석 오류", ...) -> get_file_info에서 직접 처리하거나, 여기서 처리하려면 반환값으로 구분
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


            # 총 예상 작업량 계산 (가중치 기반)
            self.total_weighted_work = (self.current_file_total_text_chars * config.WEIGHT_TEXT_CHAR) + \
                                       (self.current_file_image_elements_count * config.WEIGHT_IMAGE) + \
                                       (self.current_file_chart_elements_count * config.WEIGHT_CHART)

            logger.info(f"파일 정보 분석 완료. 총 슬라이드: {self.current_file_slide_count}, 예상 가중 작업량: {self.total_weighted_work}")
            self.current_work_label.config(text="파일 분석 완료. 번역 대기 중.")

        except FileNotFoundError: # 혹시 모를 경우 대비 (PptxHandler에서 처리했더라도)
            logger.error(f"파일 찾기 오류 (UI): {file_path}")
            self.file_name_label.config(text="파일 이름: - (파일 없음)")
            # ... (UI 초기화) ...
            messagebox.showerror("파일 오류", f"선택한 파일({os.path.basename(file_path)})을 찾을 수 없습니다.")
        except Exception as e: # 기타 예외 (예: pptx 파일 손상)
            logger.error(f"파일 정보 분석 중 UI에서 예외 발생: {e}", exc_info=True)
            self.file_name_label.config(text="파일 이름: - (오류)")
            # ... (UI 초기화) ...
            messagebox.showerror("파일 분석 오류", f"선택한 파일({os.path.basename(file_path)})을 분석하는 중 오류가 발생했습니다.\n파일이 손상되었거나 지원하지 않는 형식일 수 있습니다.\n\n오류: {e}")
    # --- 1단계 개선 끝 (에러 처리 관련) ---


    def check_ollama_status_manual(self, initial_check=False):
        # ... (기존과 동일) ...
        logger.info("Ollama 상태 확인 중...")
        self.ollama_check_button.config(state=tk.DISABLED) # 확인 중 버튼 비활성화
        self.master.update_idletasks()

        ollama_installed = self.ollama_service.is_installed()
        self.ollama_status_label.config(text=f"Ollama 설치: {'설치됨' if ollama_installed else '미설치'}")

        if not ollama_installed:
            logger.warning("Ollama가 설치되어 있지 않습니다.")
            if not initial_check: # 초기 점검이 아닐 때만 메시지 박스 표시
                if messagebox.askyesno("Ollama 설치 필요", "Ollama가 설치되어 있지 않습니다. Ollama 다운로드 페이지로 이동하시겠습니까?"):
                    webbrowser.open("https://ollama.com/download")
            self.ollama_running_label.config(text="Ollama 실행: 미설치")
            self.ollama_port_label.config(text="Ollama 포트: -")
            self.model_combo.config(values=[], state="disabled") # 모델 목록 비우고 비활성화
            self.model_var.set("")
            self.ollama_check_button.config(state=tk.NORMAL) # 버튼 다시 활성화
            return

        # Ollama 설치된 경우, 실행 상태 확인
        ollama_running, port = self.ollama_service.is_running()
        self.ollama_running_label.config(text=f"Ollama 실행: {'실행 중' if ollama_running else '미실행'}")
        self.ollama_port_label.config(text=f"Ollama 포트: {port if ollama_running and port else '-'}")

        if ollama_running:
            logger.info(f"Ollama 실행 중 (포트: {port}). 모델 목록 로드 시도.")
            self.load_ollama_models() # 모델 목록 로드
        else: # 설치는 되었으나 실행 중이지 않은 경우
            logger.warning("Ollama가 설치되었으나 실행 중이지 않습니다. 자동 시작을 시도합니다.")
            self.model_combo.config(values=[], state="disabled") # 모델 목록 비우고 비활성화
            self.model_var.set("")
            # 초기 점검 시에는 자동으로, 그 외에는 사용자에게 물어보고 시작
            if initial_check or messagebox.askyesno("Ollama 실행 필요", "Ollama가 실행 중이지 않습니다. 지금 시작하시겠습니까? (권장)"):
                if self.ollama_service.start_ollama():
                    logger.info("Ollama 자동 시작 성공. 잠시 후 상태를 다시 확인합니다.")
                    # 잠시 후 상태를 다시 확인하여 UI 업데이트
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        self.master.after(3000, lambda: self.check_ollama_status_manual(initial_check=initial_check))
                else:
                    logger.error("Ollama 자동 시작 실패. 수동으로 실행해주세요.")
                    if not initial_check:
                        messagebox.showwarning("Ollama 시작 실패", "Ollama를 자동으로 시작할 수 없습니다. 수동으로 실행 후 'Ollama 확인'을 눌러주세요.")

        self.ollama_check_button.config(state=tk.NORMAL) # 버튼 다시 활성화

    # ... (이하 나머지 Application 클래스 메소드들은 이전과 유사하거나,
    #      위에서 변경된 변수/메소드를 호출하는 부분이 자연스럽게 반영될 것입니다.
    #      예: _translation_worker에서 ocr_temperature 전달 시 self.ocr_temperature_var.get() 사용 등)

    def load_ollama_models(self):
        # ... (기존과 동일) ...
        logger.debug("Ollama 모델 목록 로드 중 (UI 요청)...")
        self.model_refresh_button.config(state=tk.DISABLED) # 새로고침 중 버튼 비활성화
        self.master.update_idletasks()

        self.ollama_service.invalidate_models_cache() # 사용자가 새로고침을 눌렀으므로 캐시 무효화

        models = self.ollama_service.get_text_models() # 캐시 적용된 함수 호출
        if models:
            self.model_combo.config(values=models, state="readonly") # 모델 목록 설정, 읽기 전용
            current_selected_model = self.model_var.get()
            # 현재 선택된 모델이 목록에 있으면 유지, 없으면 기본 모델, 그것도 없으면 첫 번째 모델 선택
            if current_selected_model in models:
                self.model_var.set(current_selected_model)
            elif DEFAULT_MODEL in models:
                self.model_var.set(DEFAULT_MODEL)
            elif models: # 목록에 모델이 하나라도 있으면 첫 번째 모델 선택
                self.model_var.set(models[0])
            else: # 목록이 비어있지만 models가 None이 아닌 경우 (빈 리스트)
                self.model_var.set("")

            logger.info(f"사용 가능 Ollama 모델: {models}")
            # 기본 모델이 없고, 현재 선택된 모델도 없는 경우 (목록은 있으나 기본 모델이 없는 경우)
            if DEFAULT_MODEL not in models and not self.model_var.get():
                self.download_default_model_if_needed(initial_check_from_ollama=True) # 기본 모델 다운로드 시도
        else: # 모델 목록을 가져오지 못한 경우 (빈 리스트 또는 None)
            self.model_combo.config(values=[], state="disabled")
            self.model_var.set("")
            logger.warning("Ollama에 로드된 모델이 없습니다.")
            self.download_default_model_if_needed(initial_check_from_ollama=True) # 기본 모델 다운로드 시도

        self.model_refresh_button.config(state=tk.NORMAL) # 새로고침 버튼 다시 활성화


    def download_default_model_if_needed(self, initial_check_from_ollama=False):
        # ... (기존과 동일) ...
        current_models = self.ollama_service.get_text_models() # 최신 모델 목록 확인
        if DEFAULT_MODEL not in current_models:
            logger.warning(f"기본 모델 ({DEFAULT_MODEL})이 설치되어 있지 않습니다.")
            # 초기 Ollama 확인 시 또는 사용자가 동의한 경우 다운로드
            if initial_check_from_ollama or messagebox.askyesno("기본 모델 다운로드", f"기본 번역 모델 '{DEFAULT_MODEL}'이(가) 없습니다. 지금 다운로드하시겠습니까? (시간 소요)"):
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드 시작...")
                self.start_button.config(state=tk.DISABLED) # 번역 시작 버튼 비활성화
                self.progress_bar["value"] = 0
                self.current_work_label.config(text=f"모델 다운로드 시작: {DEFAULT_MODEL}")
                self.progress_label_var.set(f"모델 다운로드 시작: {DEFAULT_MODEL}")

                if self.model_download_thread and self.model_download_thread.is_alive():
                    logger.warning("이미 모델 다운로드 스레드가 실행 중입니다.")
                    return

                self.stop_event.clear() # 중지 이벤트 초기화
                self.model_download_thread = threading.Thread(target=self._model_download_worker, args=(DEFAULT_MODEL, self.stop_event), daemon=True)
                self.model_download_thread.start()
            else:
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드가 취소되었습니다.")
        else:
            logger.info(f"기본 모델 ({DEFAULT_MODEL})이 이미 설치되어 있습니다.")


    def _model_download_worker(self, model_name, stop_event_ref):
        # ... (기존과 동일) ...
        success = self.ollama_service.pull_model_with_progress(model_name, self.update_model_download_progress, stop_event=stop_event_ref)
        if hasattr(self, 'master') and self.master.winfo_exists(): # UI 스레드에서 호출 보장
            self.master.after(0, self._model_download_finished, model_name, success)
        self.model_download_thread = None # 스레드 완료 후 참조 제거

    def _model_download_finished(self, model_name, success):
        # ... (기존과 동일) ...
        if success:
            logger.info(f"'{model_name}' 모델 다운로드 완료.")
            self.load_ollama_models() # 모델 목록 새로고침
            self.current_work_label.config(text=f"모델 '{model_name}' 다운로드 완료.")
        else:
            logger.error(f"'{model_name}' 모델 다운로드 실패.")
            self.current_work_label.config(text=f"모델 '{model_name}' 다운로드 실패.")
            if not self.stop_event.is_set(): # 사용자가 중지한 것이 아니라면 오류 메시지 표시
                messagebox.showerror("모델 다운로드 실패", f"'{model_name}' 모델 다운로드에 실패했습니다.\nOllama 서버 로그 또는 인터넷 연결을 확인해주세요.")

        # 번역 스레드가 실행 중이 아니면 UI 상태 복원
        if not (self.translation_thread and self.translation_thread.is_alive()):
            self.start_button.config(state=tk.NORMAL)
            self.progress_bar["value"] = 0
            self.progress_label_var.set("0%")
            if not success :
                self.current_work_label.config(text="모델 다운로드 실패. 재시도 요망.")
            else:
                self.current_work_label.config(text="대기 중")


    def update_model_download_progress(self, status_text, completed_bytes, total_bytes, is_error=False):
        # ... (기존과 동일) ...
        if self.stop_event.is_set() and "중지됨" not in status_text : return # 중지 요청 시 업데이트 안 함 (단, 중지 완료 메시지는 표시)

        percent = 0
        progress_str = status_text # 기본적으로 상태 텍스트 사용
        if total_bytes > 0:
            percent = (completed_bytes / total_bytes) * 100
            progress_str = f"{percent:.1f}%" # 진행률 퍼센트 표시

        def _update():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            if not is_error:
                self.progress_bar["value"] = percent
                self.progress_label_var.set(f"모델 다운로드: {progress_str} ({status_text})")
                self.current_work_label.config(text=f"모델 다운로드 중: {status_text} {progress_str}")
            else: # 오류 발생 시
                self.progress_label_var.set(f"모델 다운로드 오류: {status_text}")
                self.current_work_label.config(text=f"모델 다운로드 오류: {status_text}")

            # 로그 레벨에 따라 로그 기록 (너무 빈번할 수 있으므로 DEBUG 레벨 권장)
            logger.log(logging.DEBUG if not is_error else logging.ERROR,
                       f"모델 다운로드 진행: {status_text} ({completed_bytes}/{total_bytes})")

        if hasattr(self, 'master') and self.master.winfo_exists(): # UI 스레드에서 호출 보장
            self.master.after(0, _update)

    def check_ocr_engine_status(self, is_called_from_start_translation=False):
        # ... (기존과 동일) ...
        # 이 함수 내부에서 self.ocr_use_gpu_var.get() 등을 통해 현재 설정을 사용
        self.current_work_label.config(text="OCR 엔진 확인 중...")
        self.master.update_idletasks()

        selected_ui_lang = self.src_lang_var.get()
        use_easyocr = selected_ui_lang in config.EASYOCR_SUPPORTED_UI_LANGS
        engine_name_display = "EasyOCR" if use_easyocr else "PaddleOCR"
        engine_name_internal = engine_name_display.lower() # 내부 비교용 (소문자)

        ocr_lang_code = None
        if use_easyocr:
            ocr_lang_code = config.UI_LANG_TO_EASYOCR_CODE_MAP.get(selected_ui_lang)
        else: # PaddleOCR 사용
            ocr_lang_code = config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(selected_ui_lang, config.DEFAULT_PADDLE_OCR_LANG)

        if not ocr_lang_code: # 매핑되는 OCR 코드가 없는 경우
            msg = f"{engine_name_display}: 언어 '{selected_ui_lang}'에 대한 OCR 코드가 설정되지 않았습니다."
            self.ocr_status_label.config(text=msg)
            logger.error(msg)
            if is_called_from_start_translation: # 번역 시작 시 호출된 경우만 메시지 박스
                messagebox.showerror("OCR 설정 오류", msg)
            self.current_work_label.config(text="OCR 설정 오류!")
            return False

        gpu_enabled_for_ocr = self.ocr_use_gpu_var.get() # 현재 GPU 사용 설정

        # OCR 핸들러 재초기화 필요 여부 판단
        needs_reinit = False
        if not self.ocr_handler: # 핸들러가 아예 없는 경우
            needs_reinit = True
        elif self.current_ocr_engine_type != engine_name_internal: # 엔진 종류가 바뀐 경우
            needs_reinit = True
        elif self.ocr_handler.use_gpu != gpu_enabled_for_ocr: # GPU 사용 설정이 바뀐 경우
            needs_reinit = True
        # 언어 코드가 바뀐 경우 (PaddleOCR은 단일 코드, EasyOCR은 리스트에 포함 여부)
        elif engine_name_internal == "paddleocr" and self.ocr_handler.current_lang_codes != ocr_lang_code:
            needs_reinit = True
        elif engine_name_internal == "easyocr" and (not self.ocr_handler.current_lang_codes or ocr_lang_code not in self.ocr_handler.current_lang_codes):
            needs_reinit = True # EasyOCR은 여러 언어 동시 지원 가능. 현재 요청 언어가 없으면 추가 필요.

        if needs_reinit:
            self._destroy_current_ocr_handler() # 기존 핸들러 자원 해제
            logger.info(f"{engine_name_display} 핸들러 (재)초기화 시도 (언어: {ocr_lang_code}, GPU: {gpu_enabled_for_ocr}).")
            self.current_work_label.config(text=f"{engine_name_display} 엔진 로딩 중 (언어: {ocr_lang_code}, GPU: {gpu_enabled_for_ocr})...")
            self.master.update_idletasks()
            try:
                if use_easyocr:
                    if not utils.check_easyocr(): # EasyOCR 라이브러리 설치 확인
                        self.ocr_status_label.config(text=f"{engine_name_display}: 미설치")
                        if messagebox.askyesno(f"{engine_name_display} 설치 필요", f"{engine_name_display}이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            if utils.install_easyocr():
                                messagebox.showinfo(f"{engine_name_display} 설치 완료", f"{engine_name_display}이(가) 설치되었습니다. 애플리케이션을 재시작하거나 다시 시도해주세요.")
                            else:
                                messagebox.showerror(f"{engine_name_display} 설치 실패", f"{engine_name_display} 설치에 실패했습니다.")
                        self.current_work_label.config(text=f"{engine_name_display} 미설치.")
                        return False
                    # EasyOCR은 언어 코드 리스트를 받음
                    self.ocr_handler = EasyOcrHandler(lang_codes_list=[ocr_lang_code], debug_enabled=debug_mode, use_gpu=gpu_enabled_for_ocr)
                    self.current_ocr_engine_type = "easyocr"
                else: # PaddleOCR 사용
                    if not utils.check_paddleocr(): # PaddleOCR 라이브러리 설치 확인
                        self.ocr_status_label.config(text=f"{engine_name_display}: 미설치")
                        if messagebox.askyesno(f"{engine_name_display} 설치 필요", f"{engine_name_display}(paddlepaddle)이(가) 설치되어 있지 않습니다. 지금 설치하시겠습니까?"):
                            if utils.install_paddleocr():
                                messagebox.showinfo(f"{engine_name_display} 설치 완료", f"{engine_name_display}이(가) 설치되었습니다. 애플리케이션을 재시작하거나 다시 시도해주세요.")
                            else:
                                messagebox.showerror(f"{engine_name_display} 설치 실패", f"{engine_name_display} 설치에 실패했습니다.")
                        self.current_work_label.config(text=f"{engine_name_display} 미설치.")
                        return False
                    self.ocr_handler = PaddleOcrHandler(lang_code=ocr_lang_code, debug_enabled=debug_mode, use_gpu=gpu_enabled_for_ocr)
                    self.current_ocr_engine_type = "paddleocr"

                logger.info(f"{engine_name_display} 핸들러 초기화 성공 (언어: {ocr_lang_code}, GPU: {gpu_enabled_for_ocr}).")
                self.current_work_label.config(text=f"{engine_name_display} 엔진 로딩 완료.")

            except RuntimeError as e: # OCR 핸들러 초기화 실패 (라이브러리 내부 오류 등)
                logger.error(f"{engine_name_display} 핸들러 초기화 실패: {e}", exc_info=True)
                self.ocr_status_label.config(text=f"{engine_name_display}: 초기화 실패 ({ocr_lang_code}, GPU:{gpu_enabled_for_ocr})")
                if is_called_from_start_translation:
                    messagebox.showerror(f"{engine_name_display} 오류", f"{engine_name_display} 초기화 중 오류:\n{e}\n\nGPU 관련 문제일 수 있습니다. GPU 사용 옵션을 확인해보세요.")
                self._destroy_current_ocr_handler() # 실패 시 핸들러 다시 제거
                self.current_work_label.config(text=f"{engine_name_display} 엔진 초기화 실패!")
                return False
            except Exception as e_other: # 기타 예상치 못한 오류
                 logger.error(f"{engine_name_display} 핸들러 생성 중 예기치 않은 오류: {e_other}", exc_info=True)
                 self.ocr_status_label.config(text=f"{engine_name_display}: 알 수 없는 오류")
                 if is_called_from_start_translation:
                     messagebox.showerror(f"{engine_name_display} 오류", f"{engine_name_display} 처리 중 예기치 않은 오류:\n{e_other}")
                 self._destroy_current_ocr_handler() # 실패 시 핸들러 다시 제거
                 self.current_work_label.config(text=f"{engine_name_display} 엔진 오류!")
                 return False

        # OCR 상태 UI 업데이트 (재초기화 되었거나, 원래 문제 없었거나)
        self.update_ocr_status_display()

        # 최종적으로 핸들러와 엔진이 준비되었는지 확인
        if self.ocr_handler and self.ocr_handler.ocr_engine:
            return True
        else: # 준비 안 됨
            self.ocr_status_label.config(text=f"{engine_name_display} OCR: 준비 안됨 ({selected_ui_lang})")
            # 번역 시작 시 호출되었는데, 재초기화도 필요 없었지만 여전히 준비 안 된 경우 (이전 오류 등)
            if is_called_from_start_translation and not needs_reinit :
                 messagebox.showwarning("OCR 오류", f"{engine_name_display} OCR 엔진을 사용할 수 없습니다. 이전 로그를 확인해주세요.")
            self.current_work_label.config(text=f"{engine_name_display} OCR 준비 안됨.")
            return False


    def swap_languages(self):
        # ... (기존과 동일) ...
        src = self.src_lang_var.get()
        tgt = self.tgt_lang_var.get()
        self.src_lang_var.set(tgt)
        self.tgt_lang_var.set(src)
        logger.info(f"언어 스왑: {tgt} <-> {src}")
        self.on_source_language_change() # 원본 언어 변경 시 처리 호출 (OCR 상태 업데이트 등)

    def start_translation(self):
        # ... (기존과 동일, ocr_temperature 전달 부분 수정) ...
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("파일 오류", "번역할 유효한 파워포인트 파일을 선택해주세요.\n'찾아보기' 버튼을 사용하여 파일을 선택할 수 있습니다.")
            return

        # --- 1단계 개선: 고급 옵션에서 가져온 변수 사용 ---
        image_translation_really_enabled = self.image_translation_enabled_var.get()
        ocr_temperature_to_use = self.ocr_temperature_var.get()
        # --- 1단계 개선 끝 ---

        if image_translation_really_enabled: # 이미지 번역이 활성화된 경우에만 OCR 엔진 확인
            if not self.check_ocr_engine_status(is_called_from_start_translation=True):
                # OCR 준비 실패 시 사용자에게 계속 진행할지 확인
                if not messagebox.askyesno("OCR 준비 실패",
                                         "이미지 내 텍스트 번역에 필요한 OCR 기능이 준비되지 않았거나 사용할 수 없습니다.\n"
                                         "이 경우 이미지 안의 글자는 번역되지 않습니다.\n"
                                         "계속 진행하시겠습니까? (텍스트/차트만 번역)"):
                    logger.warning("OCR 준비 실패로 사용자가 번역을 취소했습니다.")
                    self.current_work_label.config(text="번역 취소됨 (OCR 준비 실패).")
                    return
                logger.warning("OCR 핸들러 준비 실패. 이미지 번역 없이 진행합니다.")
                image_translation_really_enabled = False # OCR 실패 시 이미지 번역 비활성화
        else: # 이미지 번역 비활성화 시
            logger.info("이미지 번역 옵션이 꺼져있으므로 OCR 엔진을 확인하지 않습니다.")
            self._destroy_current_ocr_handler() # 기존 OCR 핸들러가 있다면 자원 해제

        # 번역 언어 및 모델 선택 확인
        src_lang, tgt_lang, model = self.src_lang_var.get(), self.tgt_lang_var.get(), self.model_var.get()
        if not model:
            messagebox.showerror("모델 오류", "번역 모델을 선택해주세요.\nOllama 서버가 실행 중이고 모델이 다운로드되었는지 확인하세요.\n'Ollama 확인' 버튼과 모델 목록 '🔄' 버튼을 사용해볼 수 있습니다.")
            self.check_ollama_status_manual() # Ollama 상태 다시 확인
            return
        if src_lang == tgt_lang:
            messagebox.showwarning("언어 동일", "원본 언어와 번역 언어가 동일합니다.\n다른 언어를 선택해주세요.")
            return

        # Ollama 서버 실행 상태 확인
        ollama_running, _ = self.ollama_service.is_running()
        if not ollama_running:
            messagebox.showerror("Ollama 미실행", "Ollama 서버가 실행 중이지 않습니다.\nOllama를 실행한 후 'Ollama 확인' 버튼을 눌러주세요.")
            self.check_ollama_status_manual() # Ollama 상태 다시 확인
            return

        # 번역할 내용이 있는지 확인 (total_weighted_work 기반)
        if self.total_weighted_work <= 0:
            logger.info("총 예상 작업량이 0입니다. 파일 정보를 다시 로드하여 확인합니다.")
            self.load_file_info(file_path) # 파일 정보 강제 재로드
            if self.total_weighted_work <= 0: # 그래도 0이면
                messagebox.showinfo("정보", "번역할 내용이 없거나 작업량을 계산할 수 없습니다.\n파일 내용을 확인해주세요.")
                logger.warning("재확인 후에도 총 예상 작업량이 0 이하입니다. 번역을 시작하지 않습니다.")
                self.current_work_label.config(text="번역할 내용 없음.")
                return

        # 작업 로그 파일 설정
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_filename = os.path.basename(file_path)
        # 파일명에 포함될 수 없는 문자 제거
        safe_original_filename_part = "".join(c if c.isalnum() or c in ['.', '_'] else '_' for c in os.path.splitext(original_filename)[0])
        task_log_filename = f"translation_{timestamp}_{safe_original_filename_part}.log"
        task_log_filepath = os.path.join(LOGS_DIR, task_log_filename)

        # 로그 기록용 정보
        ocr_engine_for_log = self.current_ocr_engine_type if image_translation_really_enabled and self.ocr_handler else '사용 안 함'
        ocr_temp_for_log = ocr_temperature_to_use if image_translation_really_enabled else 'N/A'
        ocr_gpu_for_log = self.ocr_use_gpu_var.get() if image_translation_really_enabled and self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu') and self.ocr_handler.use_gpu else 'N/A'


        logger.info(f"번역 시작: '{original_filename}' ({src_lang} -> {tgt_lang}) using {model}. "
                    f"이미지 번역: {'활성' if image_translation_really_enabled else '비활성'}, "
                    f"OCR 엔진: {ocr_engine_for_log}, OCR 온도: {ocr_temp_for_log}, OCR GPU: {ocr_gpu_for_log}")

        # UI 상태 변경: 번역 시작 준비
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.progress_label_var.set("0%")
        self.translated_file_path_var.set("") # 이전 결과 초기화
        self.open_folder_button.config(state=tk.DISABLED)

        self.current_weighted_done = 0 # 완료된 작업량 초기화
        self.stop_event.clear() # 중지 이벤트 초기화

        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("이미 번역 스레드가 실행 중입니다.")
            messagebox.showwarning("번역 중복", "이미 다른 번역 작업이 진행 중입니다.")
            self.start_button.config(state=tk.NORMAL) # 시작 버튼 다시 활성화
            self.stop_button.config(state=tk.DISABLED) # 중지 버튼 비활성화
            return

        self.current_work_label.config(text="번역 준비 중...")
        self.master.update_idletasks()


        # 번역 작업을 위한 스레드 생성 및 시작
        self.translation_thread = threading.Thread(target=self._translation_worker,
                                                   args=(file_path, src_lang, tgt_lang, model, task_log_filepath,
                                                         image_translation_really_enabled, ocr_temperature_to_use), # ocr_temperature 전달
                                                   daemon=True)
        self.start_time = time.time() # 번역 시작 시간 기록
        self.translation_thread.start()
        self.update_progress_timer() # 진행률 업데이트 타이머 시작 (필요시)

    def _translation_worker(self, file_path, src_lang, tgt_lang, model, task_log_filepath,
                            image_translation_enabled: bool, ocr_temperature: float): # ocr_temperature 인자 추가
        # ... (기존과 동일, PptxHandler.translate_presentation_stage1 호출 시 ocr_temperature 전달)
        output_path, translation_result_status = "", "실패"
        prs = None # Presentation 객체 참조

        try:
            # 작업 로그 파일 헤더 작성
            with open(task_log_filepath, 'a', encoding='utf-8') as f_log_init:
                f_log_init.write(f"--- 번역 작업 시작 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                f_log_init.write(f"원본 파일: {os.path.basename(file_path)}\n")
                f_log_init.write(f"원본 언어: {src_lang}, 대상 언어: {tgt_lang}, 번역 모델: {model}\n")
                f_log_init.write(f"이미지 번역 활성화: {image_translation_enabled}\n")
                if image_translation_enabled:
                    f_log_init.write(f"  OCR 엔진: {self.current_ocr_engine_type or '미지정'}\n")
                    f_log_init.write(f"  OCR 번역 온도: {ocr_temperature}\n") # 전달받은 ocr_temperature 사용
                    # OCR GPU 사용 여부 로깅 (ocr_handler가 있고, use_gpu 속성이 있는 경우)
                    gpu_in_use_log = 'N/A'
                    if self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu'):
                        gpu_in_use_log = self.ocr_handler.use_gpu
                    f_log_init.write(f"  OCR GPU 사용 (실제): {gpu_in_use_log}\n")
                f_log_init.write(f"총 예상 가중 작업량: {self.total_weighted_work}\n")
                f_log_init.write("-" * 30 + "\n")
        except Exception as e_log_header:
            logger.error(f"작업 로그 파일 헤더 작성 실패: {e_log_header}")


        # 진행 상황 콜백 함수 (UI 업데이트용)
        def report_item_completed_from_handler(slide_info_or_stage: Any, item_type_str: str,
                                               weighted_work_for_item: int, text_snippet_str: str):
            if self.stop_event.is_set(): return # 중지 요청 시 콜백 무시

            self.current_weighted_done += weighted_work_for_item
            # 완료된 작업량이 전체 작업량을 넘지 않도록 제한
            self.current_weighted_done = min(self.current_weighted_done, self.total_weighted_work if self.total_weighted_work > 0 else weighted_work_for_item)

            if hasattr(self, 'master') and self.master.winfo_exists(): # UI 스레드에서 실행 보장
                self.master.after(0, self.update_translation_progress,
                                  slide_info_or_stage, item_type_str,
                                  self.current_weighted_done,
                                  self.total_weighted_work,
                                  text_snippet_str)
        try:
            # 번역할 내용이 없는 경우
            if self.total_weighted_work == 0:
                logger.warning("번역할 가중 작업량이 없습니다.")
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     # UI 스레드에서 메시지 박스 표시
                     self.master.after(0, lambda: messagebox.showinfo("정보", "파일에 번역할 내용이 없습니다."))
                translation_result_status, output_path = "내용 없음", file_path # 원본 파일 경로 반환
                with open(task_log_filepath, 'a', encoding='utf-8') as f_log_empty: # 로그 기록
                    f_log_empty.write(f"번역할 내용 없음. 원본 파일: {file_path}\n")
            else:
                # 번역 대상 언어에 맞는 폰트 코드 준비 (OCR 렌더링용)
                font_code_for_render = config.UI_LANG_TO_FONT_CODE_MAP.get(tgt_lang, 'en') # 기본 영어

                # UI 업데이트: 파일 로드 중
                if hasattr(self, 'master') and self.master.winfo_exists():
                    self.master.after(0, lambda: self.current_work_label.config(text="파일 로드 중..."))
                    self.master.update_idletasks()

                # 임시 디렉토리 생성 (1단계 결과 저장용)
                temp_dir_for_pptx_handler_main = tempfile.mkdtemp(prefix="pptx_trans_main_")
                temp_pptx_for_chart_translation_path: Optional[str] = None # 차트 번역을 위한 임시 파일 경로

                try:
                    prs = Presentation(file_path) # Presentation 객체 로드
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        self.master.after(0, lambda: self.current_work_label.config(text="1단계 (텍스트/이미지) 처리 시작..."))

                    # 1단계: 텍스트 및 이미지 번역 (차트 제외)
                    stage1_success = self.pptx_handler.translate_presentation_stage1(
                        prs, src_lang, tgt_lang,
                        self.translator,
                        self.ocr_handler if image_translation_enabled else None, # OCR 핸들러 조건부 전달
                        model, self.ollama_service,
                        font_code_for_render, task_log_filepath,
                        report_item_completed_from_handler, # 진행 상황 콜백
                        self.stop_event, # 중지 이벤트
                        image_translation_enabled,
                        ocr_temperature # OCR 번역 온도 전달
                    )

                    # 1단계 처리 후 중지 요청 확인
                    if self.stop_event.is_set():
                        logger.warning("1단계 번역 중 중지됨 (사용자 요청).")
                        translation_result_status = "부분 성공 (중지)"
                        # 중지 시 현재까지의 결과 저장 시도
                        try:
                            stopped_filename_s1 = os.path.join(temp_dir_for_pptx_handler_main,
                                                               f"{os.path.splitext(os.path.basename(file_path))[0]}_stage1_stopped.pptx")
                            if prs: prs.save(stopped_filename_s1)
                            output_path = stopped_filename_s1
                            logger.info(f"1단계 중단, 부분 저장: {output_path}")
                        except Exception as e_save_stop:
                            logger.error(f"1단계 중단 후 저장 실패: {e_save_stop}")
                            output_path = file_path # 저장 실패 시 원본 경로
                    elif not stage1_success: # 1단계 실패
                        logger.error("1단계 번역 실패.")
                        translation_result_status = "실패 (1단계 오류)"
                        output_path = file_path
                    else: # 1단계 성공
                        logger.info("번역 작업자: 1단계 완료. 임시 파일 저장 시도.")
                        if hasattr(self, 'master') and self.master.winfo_exists():
                            self.master.after(0, lambda: self.current_work_label.config(text="1단계 완료. 임시 파일 저장 중..."))
                            self.master.update_idletasks()

                        # 차트 번역을 위해 1단계 결과물을 임시 파일로 저장
                        temp_pptx_for_chart_translation_path = os.path.join(
                            temp_dir_for_pptx_handler_main,
                            f"{os.path.splitext(os.path.basename(file_path))[0]}_temp_for_charts.pptx"
                        )
                        if prs: prs.save(temp_pptx_for_chart_translation_path)
                        logger.info(f"1단계 결과 임시 저장: {temp_pptx_for_chart_translation_path}")

                        # 임시 저장된 파일에서 차트 정보 다시 가져오기 (정확한 차트 수 파악)
                        info_for_charts = self.pptx_handler.get_file_info(temp_pptx_for_chart_translation_path)
                        num_charts_in_prs = info_for_charts.get('chart_elements_count', 0)


                        if num_charts_in_prs > 0 and not self.stop_event.is_set(): # 차트가 있고 중지 요청이 없는 경우
                            if hasattr(self, 'master') and self.master.winfo_exists():
                                self.master.after(0, lambda: self.current_work_label.config(text=f"2단계 (차트) 처리 시작 ({num_charts_in_prs}개)..."))
                                self.master.update_idletasks()
                            logger.info(f"번역 작업자: 2단계 (차트) 시작. 대상 차트 수: {num_charts_in_prs}")

                            # 최종 출력 파일 경로 설정
                            safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang) # 파일명 안전 문자 처리
                            final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                            final_output_dir = os.path.dirname(file_path) # 원본 파일과 같은 디렉토리
                            final_pptx_output_path = os.path.join(final_output_dir, final_output_filename_base)

                            if hasattr(self, 'master') and self.master.winfo_exists():
                                self.master.after(0, lambda: self.current_work_label.config(text="2단계: 차트 XML 압축 해제 중..."))
                                self.master.update_idletasks()

                            # 2단계: 차트 번역
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
                            if hasattr(self, 'master') and self.master.winfo_exists():
                                self.master.after(0, lambda: self.current_work_label.config(text="2단계: 번역된 차트 XML 압축 중..."))
                                self.master.update_idletasks()


                            if self.stop_event.is_set(): # 차트 번역 중 중지
                                logger.warning("2단계 차트 번역 중 또는 완료 직후 중지됨.")
                                translation_result_status = "부분 성공 (중지)"
                                # 중지 시, 차트 번역 결과 파일이 있으면 그것을, 없으면 1단계 결과물 사용
                                output_path = output_path_charts if (output_path_charts and os.path.exists(output_path_charts)) else temp_pptx_for_chart_translation_path
                            elif output_path_charts and os.path.exists(output_path_charts): # 차트 번역 성공
                                logger.info(f"2단계 차트 번역 완료. 최종 파일: {output_path_charts}")
                                translation_result_status = "성공"
                                output_path = output_path_charts
                            else: # 차트 번역 실패
                                logger.error("2단계 차트 번역 실패 또는 결과 파일 없음. 1단계 결과물 사용 시도.")
                                translation_result_status = "실패 (2단계 오류)"
                                # 1단계 결과물을 최종 경로에 복사 시도
                                if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                                    try:
                                        shutil.copy2(temp_pptx_for_chart_translation_path, final_pptx_output_path)
                                        output_path = final_pptx_output_path
                                        logger.info(f"차트 번역 실패로 1단계 결과물을 최종 경로에 복사: {output_path}")
                                    except Exception as e_copy_fallback:
                                         logger.error(f"차트 번역 실패 후 1단계 결과물 복사 중 오류: {e_copy_fallback}.")
                                         output_path = temp_pptx_for_chart_translation_path # 복사 실패 시 임시 파일 경로
                                else: # 1단계 결과물도 없는 경우 (매우 드문 상황)
                                    output_path = file_path

                        elif self.stop_event.is_set(): # 1단계 후 중단되어 차트 번역 스킵
                            logger.info("1단계 후 중단되어 차트 번역은 실행되지 않음.")
                            translation_result_status = "부분 성공 (중지)"
                            output_path = temp_pptx_for_chart_translation_path # 1단계 결과물
                        else: # 번역할 차트가 없는 경우
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
                                else: # 1단계 임시 파일도 없는 경우
                                    logger.error("차트가 없고, 1단계 임시 파일도 찾을 수 없습니다.")
                                    translation_result_status = "실패 (파일 오류)"
                                    output_path = file_path
                            except Exception as e_copy_no_chart:
                                logger.error(f"차트 없는 경우 최종 파일 복사 중 오류: {e_copy_no_chart}")
                                translation_result_status = "실패 (파일 복사 오류)"
                                output_path = temp_pptx_for_chart_translation_path if temp_pptx_for_chart_translation_path else file_path
                finally: # 임시 디렉토리 정리
                    if 'temp_dir_for_pptx_handler_main' in locals() and temp_dir_for_pptx_handler_main and os.path.exists(temp_dir_for_pptx_handler_main):
                        try:
                            shutil.rmtree(temp_dir_for_pptx_handler_main)
                            logger.debug(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 완료.")
                        except Exception as e_clean_main_dir:
                            logger.warning(f"메인 임시 디렉토리 '{temp_dir_for_pptx_handler_main}' 삭제 중 오류: {e_clean_main_dir}")

            # 번역 성공 및 중지되지 않은 경우 최종 처리
            if translation_result_status == "성공" and not self.stop_event.is_set():
                 self.current_weighted_done = self.total_weighted_work # 진행률 100%로 설정
                 if hasattr(self, 'master') and self.master.winfo_exists():
                     self.master.after(0, self.update_translation_progress,
                                   "완료", "번역 완료됨", self.current_weighted_done, self.total_weighted_work, "최종 저장 완료")

                 if not (output_path and os.path.exists(output_path)): # 최종 결과 파일 확인
                     logger.error(f"번역 '성공'으로 기록되었으나, 최종 결과 파일({output_path})을 찾을 수 없습니다.")
                     translation_result_status = "실패 (결과 파일 없음)"
                     output_path = file_path # 원본 파일 경로로 대체
                 else: # 성공 시 폴더 열기 옵션 제공
                    if hasattr(self, 'master') and self.master.winfo_exists():
                        # 바로 열지 않고, translation_finished에서 사용자에게 물어보도록 변경 가능
                        self.master.after(100, lambda: self._ask_open_folder(output_path))


            elif "실패" in translation_result_status or "오류" in translation_result_status: # 실패 또는 오류 시
                 if hasattr(self, 'master') and self.master.winfo_exists():
                     self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath)
                 if not output_path: output_path = file_path # 출력 경로 없으면 원본으로

        except Exception as e_worker: # _translation_worker 전체를 감싸는 예외 처리
            logger.error(f"번역 작업 중 심각한 오류 발생: {e_worker}", exc_info=True)
            translation_result_status = "치명적 오류 발생"
            if not output_path: output_path = file_path # 출력 경로 없으면 원본으로
            try: # 작업 로그에 치명적 오류 기록
                with open(task_log_filepath, 'a', encoding='utf-8') as f_err:
                    f_err.write(f"\n--- 번역 작업 중 심각한 오류 발생 ---\n오류: {e_worker}\n")
                    import traceback
                    traceback.print_exc(file=f_err)
            except Exception as ef_log: logger.error(f"작업 로그 파일에 오류 기록 실패: {ef_log}")

            if hasattr(self, 'master') and self.master.winfo_exists(): # UI 스레드에서 오류 처리
                self.master.after(0, self._handle_translation_failure, translation_result_status, file_path, task_log_filepath, str(e_worker))

        finally: # 스레드 종료 전 항상 실행
            if hasattr(self, 'master') and self.master.winfo_exists():
                # 히스토리 항목 생성
                history_entry = {
                    "name": os.path.basename(file_path),
                    "src": src_lang,
                    "tgt": tgt_lang,
                    "model": model,
                    "ocr_temp": ocr_temperature if image_translation_enabled else "N/A",
                    "ocr_gpu": self.ocr_use_gpu_var.get() if image_translation_enabled and self.ocr_handler and hasattr(self.ocr_handler, 'use_gpu') else "N/A",
                    "img_trans_enabled": image_translation_enabled,
                    "status": translation_result_status,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "path": output_path or file_path, # output_path가 없으면 원본 경로
                    "log_file": task_log_filepath
                }
                self.master.after(0, self.translation_finished, history_entry) # UI 스레드에서 후처리
            self.translation_thread = None # 스레드 참조 제거

    def _handle_translation_failure(self, status, original_file, log_file, error_details=""):
        # ... (기존과 동일) ...
        logger.error(f"번역 실패: {status}, 원본: {original_file}, 로그: {log_file}, 상세: {error_details}")
        if hasattr(self, 'current_work_label') and self.current_work_label.winfo_exists():
            self.current_work_label.config(text=f"번역 실패: {status}")

        error_title = f"번역 작업 실패 ({status})"
        user_message = f"'{os.path.basename(original_file)}' 파일 번역 중 오류가 발생했습니다.\n\n상태: {status}\n"
        if error_details:
            user_message += f"오류 정보: {error_details[:200]}...\n\n" # 오류 상세 정보 일부 표시

        user_message += "다음 사항을 확인해 보세요:\n"
        user_message += "- Ollama 서버가 정상적으로 실행 중인지 ('Ollama 확인' 버튼)\n"
        user_message += "- 선택한 번역 모델이 유효한지 (모델 목록 '🔄' 버튼)\n"
        user_message += "- 원본 파일이 손상되지 않았는지\n"
        # GPU 관련 오류 메시지 예시 추가
        if "GPU" in status.upper() or "CUDA" in status.upper() or "메모리 부족" in status or \
           (self.ocr_use_gpu_var.get() and ("OCR" in status.upper() or "엔진" in status)):
            user_message += "- 고급 옵션에서 'GPU 사용'을 해제하고 다시 시도해보세요.\n"

        user_message += f"\n자세한 내용은 로그 파일에서 확인할 수 있습니다.\n로그 파일: {log_file}"

        if messagebox.askyesno(error_title, user_message + "\n\n오류 로그가 저장된 폴더를 여시겠습니까?", icon='error'):
            try:
                utils.open_folder(os.path.dirname(log_file))
            except Exception as e_open_log_dir:
                logger.warning(f"로그 폴더 열기 실패: {e_open_log_dir}")
                messagebox.showinfo("정보", f"로그 폴더를 열 수 없습니다.\n경로: {os.path.dirname(log_file)}")


    def _ask_open_folder(self, path):
        # ... (기존과 동일) ...
        if path and os.path.exists(path):
            user_choice = messagebox.askyesnocancel("번역 완료",
                                           f"번역이 완료되었습니다.\n저장된 파일: {os.path.basename(path)}\n\n결과 파일이 저장된 폴더를 여시겠습니까?",
                                           icon='info', default=messagebox.YES)
            if user_choice is True: # 사용자가 '예'를 선택한 경우
                utils.open_folder(os.path.dirname(path))
            # '아니오' 또는 '취소'는 아무 작업 안 함


    def update_translation_progress(self, current_location_info: Any, current_task_type: str,
                                    current_total_weighted_done: int, total_weighted_overall: int,
                                    current_text_snippet: str = ""):
        # ... (기존과 동일) ...
        if self.stop_event.is_set(): return # 중지 요청 시 업데이트 안 함

        progress = 0
        if total_weighted_overall > 0:
            progress = (current_total_weighted_done / total_weighted_overall) * 100
        elif current_total_weighted_done == 0 and total_weighted_overall == 0 : # 작업량이 0인 경우 (내용 없음)
             progress = 100 # 100%로 간주

        progress = min(max(0, progress), 100) # 진행률은 0~100 사이
        progress_text_val = f"{progress:.1f}%"

        task_description = current_task_type # 기본 작업 설명

        # 현재 위치 정보 가공
        location_display_text = str(current_location_info)
        if isinstance(current_location_info, (int, float)): # 슬라이드 번호로 온 경우
            location_display_text = f"슬라이드 {int(current_location_info)} / {self.current_file_slide_count}"
            # 1단계 특정 작업 표시 (예시)
            if "텍스트" in task_description: task_description = "1단계: 텍스트 요소 번역"
            elif "이미지" in task_description: task_description = "1단계: 이미지 처리"
            elif "표" in task_description: task_description = "1단계: 표 내부 텍스트 번역"
            else: task_description = f"1단계: {task_description}" # 기타 1단계 작업
        elif not current_location_info or str(current_location_info).upper() == "N/A": # 슬라이드 정보 없을 때 (예: 차트 전체 처리)
            location_display_text = "전체 파일 처리"
            if "차트" in task_description or "chart" in task_description.lower():
                task_description = f"2단계: {task_description}" # 차트 작업은 2단계로 명시
        elif str(current_location_info).lower() == "완료":
             location_display_text = "모든 슬라이드 완료"
             task_description = "번역 완료됨"


        # 현재 처리 중인 텍스트 스니펫 가공 (줄바꿈 제거, 길이 제한)
        snippet_display = current_text_snippet.replace('\n', ' ').strip()
        if len(snippet_display) > 25: snippet_display = snippet_display[:22] + "..."


        def _update_ui(): # UI 업데이트는 항상 이 함수를 통해 호출
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return # 위젯 존재 확인
            self.progress_bar["value"] = progress
            self.progress_label_var.set(progress_text_val)
            self.current_slide_label.config(text=f"현재 위치: {location_display_text}")
            self.current_work_label.config(text=f"현재 작업: {task_description} - '{snippet_display}'")

        if hasattr(self, 'master') and self.master.winfo_exists(): # UI 스레드에서 호출 보장
            self.master.after(0, _update_ui)


    def update_progress_timer(self):
        # ... (기존과 동일) ...
        # 이 함수는 현재 명시적으로 사용되지 않는 것으로 보이나,
        # 만약 주기적인 업데이트가 필요하다면 _translation_worker 내부에서 호출하거나,
        # 또는 self.after를 이용한 주기적 호출 로직이 필요.
        # 현재는 report_item_completed_from_handler 가 이벤트 기반으로 UI를 업데이트하므로,
        # 이 타이머는 추가적인 용도(예: 경과 시간 표시)가 없다면 제거 가능.
        # 만약 유지한다면, 번역 스레드가 살아있고 중지 요청이 없을 때만 재귀 호출하도록.
        if self.translation_thread and self.translation_thread.is_alive() and \
           not self.stop_event.is_set():
            # 여기에 주기적으로 업데이트할 내용 추가 (예: 경과 시간)
            # elapsed_time = time.time() - self.start_time if self.start_time else 0
            # self.elapsed_time_label.config(text=f"경과 시간: {elapsed_time:.0f}초")
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(1000, self.update_progress_timer) # 1초마다 호출


    def stop_translation(self):
        # ... (기존과 동일) ...
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("번역 중지 요청 중...")
            self.stop_event.set() # 중지 이벤트 설정
            self.stop_button.config(state=tk.DISABLED) # 중지 버튼 비활성화 (이미 눌렸으므로)
            self.current_work_label.config(text="번역 중지 요청됨...")
        elif self.model_download_thread and self.model_download_thread.is_alive(): # 모델 다운로드 중지
            logger.warning("모델 다운로드 중지 요청 중...")
            self.stop_event.set() # 중지 이벤트 설정
            self.stop_button.config(state=tk.DISABLED)
            # 모델 다운로드 중지 시 UI 메시지 업데이트는 update_model_download_progress 에서 처리


    def translation_finished(self, history_entry: Dict[str, Any]):
        # ... (기존과 동일) ...
        if not (hasattr(self, 'start_button') and self.start_button.winfo_exists()): return # UI 위젯 확인
        self.start_button.config(state=tk.NORMAL) # 시작 버튼 활성화
        self.stop_button.config(state=tk.DISABLED) # 중지 버튼 비활성화

        result_status = history_entry.get("status", "알 수 없음")
        translated_file_path = history_entry.get("path")
        current_progress_val = self.progress_bar["value"] # 현재 진행률 값

        if result_status == "성공" and not self.stop_event.is_set(): # 성공 & 사용자 중지 아님
            final_progress_text = "100%"
            self.progress_bar["value"] = 100
            # self.current_weighted_done = self.total_weighted_work # 이미 _translation_worker에서 처리됨
            self.current_work_label.config(text=f"번역 완료: {os.path.basename(translated_file_path) if translated_file_path else '파일 없음'}")
            self.current_slide_label.config(text="모든 작업 완료")
        elif "중지" in result_status: # 사용자가 중지한 경우
            final_progress_text = f"{current_progress_val:.1f}% (중지됨)"
            self.current_work_label.config(text="번역 중지됨.")
        elif result_status == "내용 없음": # 번역할 내용이 없었던 경우
            final_progress_text = "100% (내용 없음)"
            self.progress_bar["value"] = 100
            self.current_work_label.config(text="번역할 내용 없음.")
        else: # 기타 실패/오류
            final_progress_text = f"{current_progress_val:.1f}% ({result_status})"
            # current_work_label은 _handle_translation_failure 에서 이미 설정되었을 수 있음

        self.progress_label_var.set(final_progress_text)

        # 번역된 파일 경로 UI 업데이트 및 폴더 열기 버튼 상태 변경
        if translated_file_path and os.path.exists(translated_file_path) and result_status == "성공":
            self.translated_file_path_var.set(translated_file_path)
            self.open_folder_button.config(state=tk.NORMAL)
        else: # 실패했거나, 성공했으나 파일 경로가 유효하지 않은 경우
            self.translated_file_path_var.set("번역 실패 또는 파일 없음")
            self.open_folder_button.config(state=tk.DISABLED)
            if result_status == "성공" and not (translated_file_path and os.path.exists(translated_file_path)):
                 logger.warning(f"번역은 '성공'으로 기록되었으나, 결과 파일 경로가 유효하지 않음: {translated_file_path}")

        self._add_history_entry(history_entry) # 번역 히스토리 추가

        # 작업 로그 파일에 최종 상태 기록
        task_log_filepath = history_entry.get("log_file")
        if task_log_filepath and os.path.exists(os.path.dirname(task_log_filepath)): # 로그 파일 경로 유효성 확인
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    f_task_log.write(f"\n--- 번역 작업 최종 상태 ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}) ---\n")
                    f_task_log.write(f"최종 상태: {result_status}\n")
                    # UI에 표시된 원본 파일 경로 (사용자가 선택한 경로)
                    if self.file_path_var.get():
                        f_task_log.write(f"원본 파일 (UI 경로): {self.file_path_var.get()}\n")
                    # 실제 번역된 파일 경로 (성공 시)
                    if translated_file_path and os.path.exists(translated_file_path):
                        f_task_log.write(f"번역된 파일: {translated_file_path}\n")

                    # 총 소요 시간 기록
                    elapsed_time_for_log = (time.time() - self.start_time) if self.start_time else 0
                    m, s = divmod(elapsed_time_for_log, 60)
                    f_task_log.write(f"총 소요 시간 (내부 기록용): {int(m):02d}분 {s:05.2f}초\n")
                    f_task_log.write("-" * 30 + "\n")
            except Exception as e_log_finish:
                logger.error(f"작업 로그 파일에 최종 상태 기록 실패: {e_log_finish}")

        self.start_time = None # 번역 시작 시간 초기화

        # 실패 시 추가적인 사용자 알림 (이미 _handle_translation_failure 에서 처리했을 수 있음)
        # if result_status != "성공" and "중지" not in result_status and result_status != "내용 없음":
        #      pass # 이미 _handle_translation_failure에서 메시지박스 표시
        # elif result_status == "성공":
        #      pass # _ask_open_folder 에서 메시지박스 표시


    def open_translated_folder(self):
        # ... (기존과 동일) ...
        path = self.translated_file_path_var.get()
        if path and os.path.exists(path):
            utils.open_folder(os.path.dirname(path)) # 파일이므로 부모 디렉토리 열기
        elif path and not os.path.exists(path): # 경로 정보는 있으나 실제 파일이 없는 경우
            messagebox.showwarning("폴더 열기 실패", f"경로를 찾을 수 없습니다: {path}")
        else: # 경로 정보 자체가 없는 경우
            messagebox.showinfo("정보", "번역된 파일 경로가 없습니다.")


    def on_history_double_click(self, event):
        # ... (기존과 동일) ...
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()): return
        item_id = self.history_tree.identify_row(event.y) # 클릭된 아이템 ID 가져오기
        if item_id:
            item_values = self.history_tree.item(item_id, "values")
            if item_values and len(item_values) == len(self.history_tree["columns"]):
                # 경로와 상태 값 가져오기 (인덱스 기반)
                try:
                    path_idx = self.history_tree["columns"].index("path")
                    status_idx = self.history_tree["columns"].index("status")
                    time_idx = self.history_tree["columns"].index("time") # 로그 파일 식별용
                except ValueError:
                    logger.error("히스토리 Treeview 컬럼명 오류. 'path' 또는 'status' 컬럼을 찾을 수 없습니다.")
                    return

                file_path_to_open = item_values[path_idx]
                status_val = item_values[status_idx]
                time_val = item_values[time_idx]

                if file_path_to_open and os.path.exists(file_path_to_open) and "성공" in status_val :
                    if messagebox.askyesno("파일 열기", f"번역된 파일 '{os.path.basename(file_path_to_open)}'을(를) 여시겠습니까?"):
                        try:
                            if platform.system() == "Windows": os.startfile(file_path_to_open)
                            elif platform.system() == "Darwin": subprocess.Popen(["open", file_path_to_open])
                            else: subprocess.Popen(["xdg-open", file_path_to_open])
                        except Exception as e:
                            logger.error(f"히스토리 파일 열기 실패: {e}", exc_info=True)
                            messagebox.showerror("파일 열기 오류", f"파일을 여는 중 오류가 발생했습니다:\n{e}")
                elif "성공" not in status_val and file_path_to_open : # 성공이 아닌 경우 로그 파일 열기 시도
                     log_file_path_from_history = ""
                     # 히스토리 데이터에서 해당 항목의 로그 파일 경로 찾기
                     for entry_data in self.translation_history_data:
                         # 경로와 시간으로 특정 항목 식별 (동일 파일 여러 번 번역 가능성)
                         if entry_data.get("path") == file_path_to_open and entry_data.get("time") == time_val:
                             log_file_path_from_history = entry_data.get("log_file", "")
                             break

                     if log_file_path_from_history and os.path.exists(log_file_path_from_history):
                         if messagebox.askyesno("로그 파일 열기", f"번역 결과가 '{status_val}'입니다.\n관련 로그 파일 '{os.path.basename(log_file_path_from_history)}'이(가) 저장된 폴더를 여시겠습니까?"):
                             try: utils.open_folder(os.path.dirname(log_file_path_from_history))
                             except Exception as e:
                                 logger.error(f"히스토리 로그 폴더 열기 실패: {e}")
                                 messagebox.showerror("폴더 열기 오류", f"로그 폴더를 여는 중 오류가 발생했습니다:\n{e}")
                     else:
                          messagebox.showwarning("정보", f"번역 결과가 '{status_val}'입니다.\n(관련 로그 파일 정보 없음 또는 찾을 수 없음)")
                elif file_path_to_open and not os.path.exists(file_path_to_open): # 파일 경로가 있으나 존재하지 않는 경우
                    messagebox.showwarning("파일 없음", f"파일을 찾을 수 없습니다: {file_path_to_open}")


# Text 위젯으로 로그를 보내는 핸들러
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        if not (self.text_widget and self.text_widget.winfo_exists()): return # 위젯이 없으면 무시
        msg = self.format(record)
        def append_message(): # UI 스레드에서 실행될 함수
            if not (self.text_widget and self.text_widget.winfo_exists()): return
            self.text_widget.config(state=tk.NORMAL) # 쓰기 가능 상태로 변경
            self.text_widget.insert(tk.END, msg + '\n') # 메시지 추가
            self.text_widget.see(tk.END) # 마지막 줄로 스크롤
            self.text_widget.config(state=tk.DISABLED) # 다시 읽기 전용으로
        try:
            # Tkinter 위젯이 다른 스레드에서 직접 조작될 수 없으므로, after 사용
            if self.text_widget.winfo_exists():
                self.text_widget.after(0, append_message)
        except tk.TclError: # 위젯이 파괴된 후 호출될 경우의 예외 처리
            pass


if __name__ == "__main__":
    # 디렉토리 생성 (애플리케이션 시작 시)
    for dir_path in [LOGS_DIR, FONTS_DIR, ASSETS_DIR, HISTORY_DIR, os.path.dirname(USER_SETTINGS_PATH)]:
        try:
            if dir_path: # 빈 문자열이 아닌 경우에만 생성 시도
                os.makedirs(dir_path, exist_ok=True)
        except Exception as e_mkdir_main:
            # 시작 시점에는 logger가 완전히 설정되지 않았을 수 있으므로 print 사용
            print(f"디렉토리 생성 실패 ({dir_path}): {e_mkdir_main}")


    if debug_mode: logger.info("디버그 모드로 실행 중입니다.")
    else: logger.info("일반 모드로 실행 중입니다.")

    if not os.path.exists(config.FONTS_DIR) or not os.listdir(config.FONTS_DIR): # 폰트 디렉토리 존재 및 내용 확인
        logger.critical(f"필수 폰트 디렉토리({config.FONTS_DIR})를 찾을 수 없거나 비어있습니다. 애플리케이션이 정상 동작하지 않을 수 있습니다.")
        # messagebox.showerror("치명적 오류", f"필수 폰트 디렉토리({config.FONTS_DIR})를 찾을 수 없거나 비어있습니다.\n애플리케이션을 종료합니다.")
        # sys.exit(1) # 폰트 없으면 실행 불가 처리 (선택적)
    else:
        logger.info(f"폰트 디렉토리 확인: {config.FONTS_DIR}")

    if not os.path.exists(config.ASSETS_DIR):
        logger.warning(f"에셋 디렉토리를 찾을 수 없습니다: {config.ASSETS_DIR}")
    else:
        logger.info(f"에셋 디렉토리 확인: {config.ASSETS_DIR}")

    root = tk.Tk()
    app = Application(master=root) # Application 인스턴스 생성
    root.geometry("1024x768") # 기본 창 크기
    root.update_idletasks() # 창 크기 계산 위해 필요
    min_width = root.winfo_reqwidth() # 최소 너비
    min_height = root.winfo_reqheight() # 최소 높이
    root.minsize(min_width + 20, min_height + 20) # 최소 창 크기 설정 (패딩 고려)

    try:
        root.mainloop()
    except KeyboardInterrupt: # Ctrl+C로 종료 시
        logger.info("Ctrl+C로 애플리케이션 종료 중...")
    finally:
        # on_closing이 atexit에 의해 호출될 것이므로, 여기서는 추가 작업 불필요
        # 다만, mainloop 이후의 명시적인 정리 작업이 있다면 여기에 추가
        logger.info(f"--- {APP_NAME} 종료됨 (mainloop 이후) ---")