import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import webbrowser
import time
import logging # 로깅 모듈
import sys
import platform
import subprocess
from PIL import Image, ImageTk
from datetime import datetime # For timestamped log files

# 프로젝트 루트의 다른 .py 파일들 import
from translator import OllamaTranslator
from pptx_handler import PptxHandler
from ocr_handler import PaddleOcrHandler
from ollama_service import OllamaService
import utils

# --- 로깅 설정 ---
debug_mode = "--debug" in sys.argv
log_level = logging.DEBUG if debug_mode else logging.INFO

# 기본 로거 가져오기 (모든 모듈에서 이 로거를 사용)
root_logger = logging.getLogger()
root_logger.setLevel(log_level)

# 포매터 생성
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# 콘솔 핸들러 설정
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setFormatter(formatter)
if not any(isinstance(h, logging.StreamHandler) for h in root_logger.handlers): # 중복 방지
    root_logger.addHandler(console_handler)

# --- 기본 경로 설정 ---
BASE_DIR_MAIN = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR_MAIN, "assets")
FONTS_DIR = os.path.join(BASE_DIR_MAIN, "fonts")
LOGS_DIR = os.path.join(BASE_DIR_MAIN, "logs") # 작업별 로그 저장 폴더

# 로그 폴더 생성
if not os.path.exists(LOGS_DIR):
    try:
        os.makedirs(LOGS_DIR)
    except Exception as e:
        print(f"로그 폴더 생성 실패: {LOGS_DIR}, 오류: {e}") # 로거 설정 전이므로 print

# 일반 앱 로그 파일 핸들러 설정 (app_general.log)
GENERAL_LOG_FILENAME = os.path.join(LOGS_DIR, "app_general.log") # 일반 로그도 logs 폴더 하위로 이동
try:
    general_file_handler = logging.FileHandler(GENERAL_LOG_FILENAME, mode='a', encoding='utf-8') # 이어쓰기 모드
    general_file_handler.setFormatter(formatter)
    if not any(h.baseFilename == os.path.abspath(GENERAL_LOG_FILENAME) for h in root_logger.handlers if isinstance(h, logging.FileHandler)): # 중복 방지
        root_logger.addHandler(general_file_handler)
except Exception as e:
    print(f"일반 로그 파일 핸들러 설정 실패: {e}") # 로거 설정 전이므로 print 사용

logger = logging.getLogger(__name__) # main.py용 로거

# --- 전역 변수 및 설정 ---
APP_NAME = "Powerpoint Document Translator"
DEFAULT_MODEL = "gemma3:12b" # 사용자님 요청으로 gemma3:12b 로 변경
SUPPORTED_LANGUAGES = ["한국어", "일본어", "영어", "중국어", "대만어", "태국어", "스페인어"]
UI_LANG_TO_PADDLEOCR_CODE = {
    "한국어": "korean", "일본어": "japan", "영어": "en",
    "중국어": "ch_doc", "대만어": "chinese_cht", "태국어": "th", "스페인어": "es",
}
DEFAULT_PADDLE_OCR_LANG = "korean"
UI_LANG_TO_FONT_CODE_MAP = {
    "한국어": "korean", "일본어": "japan", "영어": "en",
    "중국어": "ch_doc", "대만어": "chinese_cht", "태국어": "th", "스페인어": "es",
}
TRANSLATION_HISTORY = []


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title(APP_NAME)

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
                except tk.TclError as e_tk_icon:
                    logger.warning(f"tk.PhotoImage로 PNG 아이콘 설정 실패: {e_tk_icon}. Pillow 시도.")
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
        self.ocr_handler = None
        
        self.translation_thread = None
        self.model_download_thread = None # 모델 다운로드 스레드 관리를 위해 추가
        self.stop_event = threading.Event()
        
        self.logo_image_tk_bottom = None
        self.start_time = None
        
        self.create_widgets()
        self.master.after(100, self.initial_checks)
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing) # 자원 해제 로직 연결

        logger.info(f"--- {APP_NAME} 시작됨 (일반 로그 파일: {os.path.abspath(GENERAL_LOG_FILENAME)}) ---")
        logger.debug(f"디버그 모드: {debug_mode}")
        logger.debug(f"Assets 디렉토리: {ASSETS_DIR}, 존재 여부: {os.path.exists(ASSETS_DIR)}")
        logger.debug(f"Fonts 디렉토리: {FONTS_DIR}, 존재 여부: {os.path.exists(FONTS_DIR)}")
        logger.debug(f"Logs 디렉토리: {LOGS_DIR}, 존재 여부: {os.path.exists(LOGS_DIR)}")

    def on_closing(self):
        """애플리케이션 종료 시 호출되는 함수. 자원 해제 및 스레드 정리."""
        logger.info("애플리케이션 종료 절차 시작...")
        if messagebox.askokcancel("종료 확인", f"{APP_NAME}을(를) 종료하시겠습니까? 진행 중인 작업이 있다면 중단될 수 있습니다."):
            self.stop_event.set() # 모든 활성 스레드에 중지 신호 전송

            if self.translation_thread and self.translation_thread.is_alive():
                logger.info("번역 스레드 종료 대기 중...")
                self.translation_thread.join(timeout=5) # 최대 5초 대기
                if self.translation_thread.is_alive():
                    logger.warning("번역 스레드가 시간 내에 종료되지 않았습니다.")
            
            if self.model_download_thread and self.model_download_thread.is_alive():
                logger.info("모델 다운로드 스레드 종료 대기 중... (Ollama는 자체적으로 중단 처리할 수 있음)")
                # 모델 다운로드의 경우 Ollama 프로세스에 의해 관리되므로,
                # 스레드 자체를 강제로 중단하기보다는 join으로 완료를 기다리거나 타임아웃 처리
                self.model_download_thread.join(timeout=2) # 짧게 대기
                if self.model_download_thread.is_alive():
                    logger.warning("모델 다운로드 스레드가 시간 내에 정상 종료되지 않았습니다.")
            
            # PaddleOCR 핸들러 등 다른 명시적 해제가 필요한 자원이 있다면 여기에 추가
            if self.ocr_handler and hasattr(self.ocr_handler, 'ocr') and self.ocr_handler.ocr:
                logger.debug("PaddleOCR 자원 해제 시도 (del self.ocr_handler.ocr)")
                try:
                    del self.ocr_handler.ocr # PaddleOCR 객체 해제 시도
                    # 일부 환경에서는 명시적 해제가 필요 없을 수 있으나, 시도해볼 수 있음
                except Exception as e_ocr_del:
                    logger.warning(f"PaddleOCR 자원 해제 중 오류: {e_ocr_del}")
                self.ocr_handler = None

            logger.info("모든 정리 작업 완료. 애플리케이션을 종료합니다.")
            self.master.destroy()
        else:
            logger.info("애플리케이션 종료 취소됨.")


    def initial_checks(self):
        """GUI가 표시된 후 실행되는 초기 점검들"""
        logger.debug("초기 점검 시작: PaddleOCR 및 Ollama 상태 확인")
        self.check_paddleocr_status_manual(initial_check=True)
        self.check_ollama_status_manual(initial_check=True)
        logger.debug("초기 점검 완료.")

    def create_widgets(self):
        top_frame = ttk.Frame(self); top_frame.pack(fill=tk.BOTH, expand=True)
        bottom_frame = ttk.Frame(self, height=30); bottom_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(5,0)); bottom_frame.pack_propagate(False)
        main_paned_window = ttk.PanedWindow(top_frame, orient=tk.HORIZONTAL); main_paned_window.pack(fill=tk.BOTH, expand=True)
        left_panel = ttk.Frame(main_paned_window, padding=10); main_paned_window.add(left_panel, weight=2)
        right_panel = ttk.Frame(main_paned_window, padding=10); main_paned_window.add(right_panel, weight=1)

        path_frame = ttk.LabelFrame(left_panel, text="파일 경로", padding=5); path_frame.pack(padx=5, pady=(0,5), fill=tk.X)
        self.file_path_var = tk.StringVar()
        file_entry = ttk.Entry(path_frame, textvariable=self.file_path_var); file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        browse_button = ttk.Button(path_frame, text="찾아보기", command=self.browse_file); browse_button.pack(side=tk.LEFT)

        server_status_frame = ttk.LabelFrame(left_panel, text="서버 상태", padding=5); server_status_frame.pack(padx=5, pady=5, fill=tk.X)
        server_status_frame.columnconfigure(1, weight=1)
        self.ollama_status_label = ttk.Label(server_status_frame, text="Ollama 설치: 미확인"); self.ollama_status_label.grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.ollama_running_label = ttk.Label(server_status_frame, text="Ollama 실행: 미확인"); self.ollama_running_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        self.ollama_port_label = ttk.Label(server_status_frame, text="Ollama 포트: -"); self.ollama_port_label.grid(row=0, column=2, padx=5, pady=2, sticky=tk.W)
        ttk.Button(server_status_frame, text="Ollama 확인", command=self.check_ollama_status_manual).grid(row=0, column=3, padx=5, pady=2, sticky=tk.E)
        self.paddleocr_status_label = ttk.Label(server_status_frame, text="PaddleOCR 상태: 미확인"); self.paddleocr_status_label.grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        ttk.Button(server_status_frame, text="PaddleOCR 확인", command=self.check_paddleocr_status_manual).grid(row=1, column=1, columnspan=3, padx=5, pady=2, sticky=tk.W)

        file_progress_outer_frame = ttk.Frame(left_panel); file_progress_outer_frame.pack(padx=5, pady=5, fill=tk.X)
        file_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="파일 정보", padding=5); file_info_frame.pack(side=tk.LEFT, padx=(0,5), fill=tk.BOTH, expand=True)
        self.file_name_label = ttk.Label(file_info_frame, text="파일 이름: "); self.file_name_label.pack(anchor=tk.W, pady=1)
        self.slide_count_label = ttk.Label(file_info_frame, text="슬라이드 수: "); self.slide_count_label.pack(anchor=tk.W, pady=1)
        self.text_elements_label = ttk.Label(file_info_frame, text="텍스트 요소 수: "); self.text_elements_label.pack(anchor=tk.W, pady=1)
        # '이미지 내 텍스트 수' 레이블의 텍스트를 '총 이미지 수'로 변경
        self.image_elements_label = ttk.Label(file_info_frame, text="총 이미지 수: "); self.image_elements_label.pack(anchor=tk.W, pady=1)
        self.total_elements_label = ttk.Label(file_info_frame, text="총 번역 요소 수: "); self.total_elements_label.pack(anchor=tk.W, pady=1)
        progress_info_frame = ttk.LabelFrame(file_progress_outer_frame, text="진행 상황", padding=5); progress_info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.current_slide_label = ttk.Label(progress_info_frame, text="현재 슬라이드: -"); self.current_slide_label.pack(anchor=tk.W, pady=1)
        self.current_work_label = ttk.Label(progress_info_frame, text="현재 작업: -"); self.current_work_label.pack(anchor=tk.W, pady=1)
        self.translated_elements_label = ttk.Label(progress_info_frame, text="번역된 요소: 0"); self.translated_elements_label.pack(anchor=tk.W, pady=1)
        self.remaining_elements_label = ttk.Label(progress_info_frame, text="남은 요소: 0"); self.remaining_elements_label.pack(anchor=tk.W, pady=1)

        translation_options_frame = ttk.LabelFrame(left_panel, text="번역 옵션", padding=5); translation_options_frame.pack(padx=5, pady=5, fill=tk.X)
        translation_options_frame.columnconfigure(1, weight=1); translation_options_frame.columnconfigure(4, weight=1)
        ttk.Label(translation_options_frame, text="원본 언어:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.src_lang_var = tk.StringVar(value=SUPPORTED_LANGUAGES[0])
        self.src_lang_combo = ttk.Combobox(translation_options_frame, textvariable=self.src_lang_var, values=SUPPORTED_LANGUAGES, state="readonly", width=12); self.src_lang_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.EW)
        self.src_lang_combo.bind("<<ComboboxSelected>>", self.on_source_language_change)
        self.swap_button = ttk.Button(translation_options_frame, text="↔", command=self.swap_languages, width=3); self.swap_button.grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(translation_options_frame, text="번역 언어:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.tgt_lang_var = tk.StringVar(value=SUPPORTED_LANGUAGES[1])
        self.tgt_lang_combo = ttk.Combobox(translation_options_frame, textvariable=self.tgt_lang_var, values=SUPPORTED_LANGUAGES, state="readonly", width=12); self.tgt_lang_combo.grid(row=0, column=4, padx=5, pady=5, sticky=tk.EW)
        ttk.Label(translation_options_frame, text="번역 모델:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.model_var = tk.StringVar(value=DEFAULT_MODEL)
        self.model_combo = ttk.Combobox(translation_options_frame, textvariable=self.model_var, state="disabled"); self.model_combo.grid(row=1, column=1, columnspan=4, padx=5, pady=5, sticky=tk.EW)

        action_buttons_frame = ttk.Frame(left_panel, padding=(0,5,0,0)); action_buttons_frame.pack(padx=5, pady=10, fill=tk.X)
        self.start_button = ttk.Button(action_buttons_frame, text="번역 시작", command=self.start_translation, style="Accent.TButton"); self.start_button.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X, ipady=5)
        self.stop_button = ttk.Button(action_buttons_frame, text="번역 중지", command=self.stop_translation, state=tk.DISABLED); self.stop_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=5)
        try: self.style.configure("Accent.TButton", font=('Helvetica', 10, 'bold'), foreground="white", background="#0078D7")
        except tk.TclError: logger.warning("Accent.TButton 스타일 적용 실패.")

        progress_bar_frame = ttk.Frame(left_panel); progress_bar_frame.pack(padx=5, pady=5, fill=tk.X)
        self.progress_bar = ttk.Progressbar(progress_bar_frame, orient="horizontal", length=300, mode="determinate"); self.progress_bar.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        self.progress_label_var = tk.StringVar(value="0% (총 소요시간: 00:00.00)"); ttk.Label(progress_bar_frame, textvariable=self.progress_label_var).pack(side=tk.LEFT)

        self.translated_file_path_var = tk.StringVar()
        translated_file_frame = ttk.LabelFrame(left_panel, text="번역 완료 파일", padding=5); translated_file_frame.pack(padx=5, pady=5, fill=tk.X)
        self.translated_file_entry = ttk.Entry(translated_file_frame, textvariable=self.translated_file_path_var, state="readonly", width=60); self.translated_file_entry.pack(side=tk.LEFT, padx=(0,5), expand=True, fill=tk.X)
        self.open_folder_button = ttk.Button(translated_file_frame, text="폴더 열기", command=self.open_translated_folder, state=tk.DISABLED); self.open_folder_button.pack(side=tk.LEFT)

        right_panel_notebook = ttk.Notebook(right_panel); right_panel_notebook.pack(fill=tk.BOTH, expand=True)
        log_tab_frame = ttk.Frame(right_panel_notebook, padding=5); right_panel_notebook.add(log_tab_frame, text="실행 로그") # 이 탭은 일반 앱 로그를 보여줌
        self.log_text = tk.Text(log_tab_frame, height=15, state=tk.DISABLED, wrap=tk.WORD, relief=tk.SOLID, borderwidth=1, font=("TkFixedFont", 9))
        log_scrollbar_y = ttk.Scrollbar(log_tab_frame, orient="vertical", command=self.log_text.yview); self.log_text.config(yscrollcommand=log_scrollbar_y.set); log_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # TextHandler를 log_text 위젯에 연결 (일반 로그용)
        text_widget_handler = TextHandler(self.log_text)
        text_widget_handler.setFormatter(formatter)
        # root_logger에 직접 추가하기보다, main.py의 logger 또는 특정 logger에만 추가하는 것을 고려할 수 있음
        # 여기서는 UI에 표시되는 로그이므로, root_logger에 추가하여 모든 로그를 보여주도록 함.
        # 단, 파일 핸들러와 콘솔 핸들러는 이미 root_logger에 있으므로, 중복 포맷팅은 발생하지 않음.
        if not any(isinstance(h, TextHandler) for h in root_logger.handlers): # 중복 방지
            root_logger.addHandler(text_widget_handler)

        history_tab_frame = ttk.Frame(right_panel_notebook, padding=5); right_panel_notebook.add(history_tab_frame, text="번역 히스토리")
        self.history_tree = ttk.Treeview(history_tab_frame, columns=("name", "src", "tgt", "result", "time", "path"), show="headings", height=10)
        self.history_tree.heading("name", text="문서 이름"); self.history_tree.heading("src", text="원본언어"); self.history_tree.heading("tgt", text="번역언어"); self.history_tree.heading("result", text="결과"); self.history_tree.heading("time", text="번역일시"); self.history_tree.heading("path", text="경로")
        self.history_tree.column("name", width=150, anchor=tk.W, stretch=tk.YES); self.history_tree.column("src", width=70, anchor=tk.CENTER); self.history_tree.column("tgt", width=70, anchor=tk.CENTER); self.history_tree.column("result", width=80, anchor=tk.CENTER); self.history_tree.column("time", width=120, anchor=tk.CENTER); self.history_tree.column("path", width=0, stretch=tk.NO)
        hist_scrollbar_y = ttk.Scrollbar(history_tab_frame, orient="vertical", command=self.history_tree.yview); hist_scrollbar_x = ttk.Scrollbar(history_tab_frame, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=hist_scrollbar_y.set, xscrollcommand=hist_scrollbar_x.set); hist_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y); hist_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.history_tree.pack(fill=tk.BOTH, expand=True); self.history_tree.bind("<Double-1>", self.on_history_double_click)

        logo_path_bottom = os.path.join(ASSETS_DIR, "LINEstudio2.png")
        if os.path.exists(logo_path_bottom):
            try:
                logger.debug(f"하단 로고 로드 시도 (tkinter.PhotoImage): {logo_path_bottom}")
                pil_temp_for_size = Image.open(logo_path_bottom)
                original_width, original_height = pil_temp_for_size.size
                pil_temp_for_size.close()
                target_height_bottom = 20
                subsample_factor = 1
                if original_height > target_height_bottom and target_height_bottom > 0: subsample_factor = max(1, int(original_height / target_height_bottom))
                elif original_height > 0 : subsample_factor = 1
                else: subsample_factor = 6 # Default if original_height is 0 or negative
                
                if original_width > 0 and original_height > 0:
                    target_width_approx = int(target_height_bottom * (original_width / original_height))
                    if target_width_approx > 0 :
                         subsample_factor_w = max(1, int(original_width / target_width_approx))
                         subsample_factor = max(subsample_factor, subsample_factor_w) # Use the larger subsample factor
                
                if subsample_factor < 1: subsample_factor = 1

                temp_logo_image_bottom = tk.PhotoImage(file=logo_path_bottom, master=self.master)
                self.logo_image_tk_bottom = temp_logo_image_bottom.subsample(subsample_factor, subsample_factor)
                logo_label_bottom = ttk.Label(bottom_frame, image=self.logo_image_tk_bottom)
                logo_label_bottom.pack(side=tk.RIGHT, padx=10, pady=2)
                logger.info(f"하단 로고 로드 성공 (tkinter.PhotoImage, 1/{subsample_factor} 크기)")
            except tk.TclError as e_tk_bottom: logger.warning(f"하단 로고 로드 실패 (tkinter.PhotoImage - TclError): {e_tk_bottom}", exc_info=True)
            except Exception as e_general_bottom: logger.warning(f"하단 로고 로드 중 예외: {e_general_bottom}", exc_info=True)
        else: logger.warning(f"하단 로고 파일({logo_path_bottom})을 찾을 수 없습니다.")
    
    def on_source_language_change(self, event=None):
        selected_ui_lang = self.src_lang_var.get()
        logger.info(f"원본 언어 변경됨: {selected_ui_lang}.") # 일반 로거 사용
        self.check_paddleocr_status_manual(initial_check=False) # False 전달 확인
        if self.file_path_var.get(): self.load_file_info(self.file_path_var.get())

    # log_message 함수는 이제 사용되지 않음. logging 모듈 직접 사용.
    # def log_message(self, message, level="INFO", exc_info=False):
    #     # ... (기존 로직은 TextHandler와 logging 모듈로 대체됨)

    def browse_file(self):
        file_path = filedialog.askopenfilename(title="파워포인트 파일 선택", filetypes=(("PowerPoint files", "*.pptx"), ("All files", "*.*")))
        if file_path:
            self.file_path_var.set(file_path); logger.info(f"파일 선택됨: {file_path}")
            self.load_file_info(file_path); self.translated_file_path_var.set("")
            self.open_folder_button.config(state=tk.DISABLED)

    def load_file_info(self, file_path):
        # 기본값 설정 (image_elements_with_text는 get_file_info에서 0으로 반환됨)
        info = {"slide_count": 0, "text_elements": 0, "image_elements": 0, "image_elements_with_text": 0}
        
        try:
            logger.debug(f"파일 정보 분석 중 (OCR 미수행): {file_path}"); file_name = os.path.basename(file_path)
            # pptx_handler.get_file_info는 이제 ocr_handler를 내부적으로 사용하지 않음 (카운팅 목적에서)
            info = self.pptx_handler.get_file_info(file_path, self.ocr_handler) 
            
            text_elements_count = info.get('text_elements',0)
            image_elements_count = info.get('image_elements',0) # 총 이미지 수

            self.file_name_label.config(text=f"파일 이름: {file_name}")
            self.slide_count_label.config(text=f"슬라이드 수: {info.get('slide_count',0)}")
            self.text_elements_label.config(text=f"텍스트 요소 수: {text_elements_count}")
            
            # self.image_elements_label은 이제 "총 이미지 수"를 표시 (create_widgets에서 텍스트 변경됨)
            self.image_elements_label.config(text=f"총 이미지 수: {image_elements_count}")
            
            # 총 번역 요소 수는 텍스트 요소와 모든 이미지 요소의 합으로 표시 (이미지는 OCR 시도 대상이 됨)
            total_elements_for_display = text_elements_count + image_elements_count
            self.total_elements_label.config(text=f"총 번역 요소 수: {total_elements_for_display}")
            self.remaining_elements_label.config(text=f"남은 요소: {total_elements_for_display}") # 초기 상태

            logger.info("파일 정보 분석 완료 (OCR 카운팅 미수행).")
        except Exception as e: 
            logger.error(f"파일 정보 분석 오류: {e}", exc_info=True)
            # 오류 발생 시 레이블 초기화 또는 오류 메시지 표시도 고려 가능
            self.file_name_label.config(text="파일 이름: 분석 오류")
            self.slide_count_label.config(text="슬라이드 수: -")
            self.text_elements_label.config(text="텍스트 요소 수: -")
            self.image_elements_label.config(text="총 이미지 수: -")
            self.total_elements_label.config(text="총 번역 요소 수: -")
            self.remaining_elements_label.config(text="남은 요소: -")

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
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(2000, lambda: self.check_ollama_status_manual(initial_check=initial_check)) # initial_check 전달
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
            self.download_default_model_if_needed(initial_check_from_ollama=True) # 모델이 없을 때도 다운로드 시도

    def download_default_model_if_needed(self, initial_check_from_ollama=False):
        current_models = self.ollama_service.get_text_models() # 최신 모델 목록 다시 확인
        if DEFAULT_MODEL not in current_models:
            logger.warning(f"기본 모델 ({DEFAULT_MODEL})이 설치되어 있지 않습니다.")
            if initial_check_from_ollama or messagebox.askyesno("기본 모델 다운로드", f"기본 번역 모델 '{DEFAULT_MODEL}'이(가) 없습니다. 지금 다운로드하시겠습니까? (시간 소요)"):
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드 시작...")
                self.start_button.config(state=tk.DISABLED)
                self.progress_bar["value"] = 0
                self.progress_label_var.set(f"모델 다운로드 시작: {DEFAULT_MODEL}")
                
                # 스레드 관리 추가
                if self.model_download_thread and self.model_download_thread.is_alive():
                    logger.warning("이미 모델 다운로드 스레드가 실행 중입니다. 새로운 다운로드를 시작하지 않습니다.")
                    return
                self.model_download_thread = threading.Thread(target=self._model_download_worker, args=(DEFAULT_MODEL,), daemon=True)
                self.model_download_thread.start()
            else:
                logger.info(f"'{DEFAULT_MODEL}' 모델 다운로드가 취소되었습니다.")
        else:
            logger.info(f"기본 모델 ({DEFAULT_MODEL})이 이미 설치되어 있습니다.")

    def _model_download_worker(self, model_name):
        success = self.ollama_service.pull_model_with_progress(model_name, self.update_model_download_progress)
        if hasattr(self, 'master') and self.master.winfo_exists(): 
            self.master.after(0, self._model_download_finished, model_name, success)
        self.model_download_thread = None # 스레드 종료 후 참조 제거

    def _model_download_finished(self, model_name, success):
        if success: 
            logger.info(f"'{model_name}' 모델 다운로드 완료.")
            self.load_ollama_models() # 모델 목록 다시 로드
        else: 
            logger.error(f"'{model_name}' 모델 다운로드 실패.")
            if not self.stop_event.is_set(): # 사용자가 앱을 종료하는 중이 아닐 때만 메시지 박스
                 messagebox.showerror("모델 다운로드 실패", f"'{model_name}' 모델 다운로드에 실패했습니다. Ollama 로그를 확인해주세요.")
        
        # 번역 스레드가 실행 중이 아닐 때만 UI 업데이트 (번역 중 모델 다운로드는 UI 건드리지 않음)
        if not (self.translation_thread and self.translation_thread.is_alive()):
            self.start_button.config(state=tk.NORMAL)
            self.progress_bar["value"] = 0
            self.progress_label_var.set("0% (총 소요시간: 00:00.00)")

    def update_model_download_progress(self, status_text, completed_bytes, total_bytes, is_error=False):
        if self.stop_event.is_set(): return # 앱 종료 중이면 UI 업데이트 안 함

        if total_bytes > 0: percent = (completed_bytes / total_bytes) * 100; progress_str = f"{percent:.1f}%"
        else: percent = 0; progress_str = status_text
        
        def _update():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            if not is_error: 
                self.progress_bar["value"] = percent
                self.progress_label_var.set(f"모델 다운로드: {progress_str} ({status_text})")
            logger.log(logging.ERROR if is_error else logging.DEBUG, f"모델 다운로드 진행: {status_text} ({completed_bytes}/{total_bytes})")
        
        if hasattr(self, 'master') and self.master.winfo_exists(): 
            self.master.after(0, _update)


    def check_paddleocr_status_manual(self, initial_check=False):
        logger.info("PaddleOCR 상태 확인 중...")
        selected_ui_lang = self.src_lang_var.get()
        paddle_ocr_code = UI_LANG_TO_PADDLEOCR_CODE.get(selected_ui_lang, DEFAULT_PADDLE_OCR_LANG)
        if selected_ui_lang not in UI_LANG_TO_PADDLEOCR_CODE:
            logger.warning(f"UI 언어 '{selected_ui_lang}'에 대한 PaddleOCR 코드 매핑 없음. 기본 '{DEFAULT_PADDLE_OCR_LANG}' 사용.")

        if utils.check_paddleocr():
            try:
                if not self.ocr_handler or \
                   (self.ocr_handler and self.ocr_handler.current_lang != paddle_ocr_code) or \
                   (self.ocr_handler and self.ocr_handler.debug_mode != debug_mode):
                    log_msg_init = f"PaddleOCR 핸들러 재초기화 시도 (요청 언어: {paddle_ocr_code}, 현재: {self.ocr_handler.current_lang if self.ocr_handler else 'N/A'})."
                    logger.info(log_msg_init)
                    if self.ocr_handler and hasattr(self.ocr_handler, 'ocr'): # 기존 핸들러 자원 해제 시도
                        del self.ocr_handler.ocr
                    self.ocr_handler = PaddleOcrHandler(lang=paddle_ocr_code, debug_enabled=debug_mode)
                
                self.paddleocr_status_label.config(text=f"PaddleOCR 상태: 준비됨 ({self.ocr_handler.current_lang})") # 실제 초기화된 언어 사용
                logger.info(f"PaddleOCR 준비 완료 (언어: {self.ocr_handler.current_lang}).")
                if self.file_path_var.get() and not initial_check: self.load_file_info(self.file_path_var.get())

            except RuntimeError as re_ocr: # PaddleOCR 초기화 시 발생할 수 있는 명시적 오류
                 logger.error(f"PaddleOCR 초기화 실패 (런타임 오류 - 요청 언어: {paddle_ocr_code}): {re_ocr}", exc_info=True)
                 self.paddleocr_status_label.config(text=f"PaddleOCR 상태: 초기화 실패 ({paddle_ocr_code})")
                 if not initial_check: messagebox.showerror("PaddleOCR 오류", f"PaddleOCR 초기화 중 오류가 발생했습니다 ({paddle_ocr_code}):\n{re_ocr}\n\n지원되지 않는 언어이거나 모델 파일 문제일 수 있습니다. 프로그램을 재시작하거나 다른 원본 언어를 선택해보세요.")
                 self.ocr_handler = None
            except Exception as e_ocr:
                logger.error(f"PaddleOCR 초기화 중 예외 발생 (요청 언어: {paddle_ocr_code}): {e_ocr}", exc_info=True)
                self.paddleocr_status_label.config(text="PaddleOCR 상태: 알 수 없는 오류")
                self.ocr_handler = None
                if not initial_check: messagebox.showerror("PaddleOCR 오류", f"PaddleOCR 처리 중 예기치 않은 오류 ({paddle_ocr_code}): {e_ocr}")
        else: # PaddleOCR 미설치
            self.paddleocr_status_label.config(text="PaddleOCR 상태: 미설치")
            logger.warning("PaddleOCR이 설치되어 있지 않습니다.")
            self.ocr_handler = None # 핸들러 참조 제거
            
            install_prompt_message = "PaddleOCR이(가) 설치되어 있지 않습니다. 지금 자동으로 설치하시겠습니까? (권장)\n설치에는 다소 시간이 소요될 수 있습니다."
            critical_fail_message = ("PaddleOCR 자동 설치 실패. 수동으로 설치해주세요.\n"
                                     "터미널/명령 프롬프트에서 다음 명령어를 실행하세요:\n"
                                     "pip install paddlepaddle paddleocr\n"
                                     "그 후 프로그램을 재시작해주세요.")

            if initial_check: # 앱 시작 시 자동 설치 시도 (사용자 확인 없이)
                logger.info("최초 실행: PaddleOCR 자동 설치를 시도합니다.")
                if utils.install_paddleocr():
                    logger.info("PaddleOCR 자동 설치 성공. 잠시 후 다시 확인합니다.")
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(1000, lambda: self.check_paddleocr_status_manual(initial_check=True))
                else:
                    logger.critical("PaddleOCR 자동 설치 실패 (최초 실행). 수동 설치 필요.")
                    # 치명적 오류로 간주하고 사용자에게 알림 (단, 프로그램 강제 종료는 하지 않음)
                    messagebox.showerror("PaddleOCR 설치 실패", critical_fail_message)
            elif messagebox.askyesno("PaddleOCR 설치 필요", install_prompt_message): # 사용자 확인 후 설치
                if utils.install_paddleocr():
                    logger.info("PaddleOCR 자동 설치 성공. 다시 확인합니다.")
                    if hasattr(self, 'master') and self.master.winfo_exists(): self.master.after(100, self.check_paddleocr_status_manual) # initial_check=False 기본값
                else:
                    logger.error("PaddleOCR 자동 설치 실패 (사용자 동의). 수동 설치 필요.")
                    messagebox.showwarning("PaddleOCR 설치 실패", critical_fail_message)


    def swap_languages(self):
        src, tgt = self.src_lang_var.get(), self.tgt_lang_var.get()
        self.src_lang_var.set(tgt); self.tgt_lang_var.set(src)
        logger.info(f"언어 스왑: {tgt} <-> {src}")
        self.on_source_language_change() # Trigger OCR check and file info reload

    def start_translation(self):
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path): 
            messagebox.showerror("파일 오류", "번역할 유효한 파워포인트 파일을 선택해주세요.")
            return
        
        if not self.ocr_handler:
            logger.warning("PaddleOCR 미준비 상태로 번역 시도.")
            if not messagebox.askyesno("OCR 미준비", "PaddleOCR이 준비되지 않아 이미지 내 텍스트는 번역되지 않습니다.\n그래도 계속하시겠습니까?"): 
                self.check_paddleocr_status_manual() # 사용자에게 OCR 확인/설치 유도
                return
            logger.info("OCR 미준비 상태로 번역 진행 (이미지 내 텍스트 제외).")

        src_lang, tgt_lang, model = self.src_lang_var.get(), self.tgt_lang_var.get(), self.model_var.get()
        if not model: 
            messagebox.showerror("모델 오류", "번역 모델을 선택해주세요.")
            self.check_ollama_status_manual() # Ollama 상태 확인 및 모델 로드 유도
            return
        if src_lang == tgt_lang: 
            messagebox.showwarning("언어 동일", "원본 언어와 번역 언어가 동일합니다.")
            return
        
        ollama_running, _ = self.ollama_service.is_running()
        if not ollama_running: 
            messagebox.showerror("Ollama 미실행", "Ollama 서버가 실행 중이지 않습니다. Ollama를 실행 후 다시 시도해주세요.")
            self.check_ollama_status_manual()
            return

        # --- 작업별 로그 파일 경로 생성 ---
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_filename = os.path.basename(file_path)
        # 파일 이름에서 확장자를 제외하고, 안전한 문자만 남김
        safe_original_filename_part = "".join(c if c.isalnum() or c in ['.', '_'] else '_' for c in os.path.splitext(original_filename)[0])
        task_log_filename = f"translation_{timestamp}_{safe_original_filename_part}.log"
        task_log_filepath = os.path.join(LOGS_DIR, task_log_filename)
        logger.info(f"번역 작업 로그 파일 생성: {task_log_filepath}")
        # --- 작업별 로그 파일 경로 생성 완료 ---

        logger.info(f"번역 시작: '{original_filename}' ({src_lang} -> {tgt_lang}) using {model}")
        self.start_button.config(state=tk.DISABLED); self.stop_button.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0; self.progress_label_var.set("0% (시작 중...)")
        self.translated_file_path_var.set(""); self.open_folder_button.config(state=tk.DISABLED)
        self.translated_elements_label.config(text="번역된 요소: 0"); self.stop_event.clear()
        
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("이미 번역 스레드가 실행 중입니다. 새로운 번역을 시작하지 않습니다.")
            messagebox.showwarning("번역 중복", "이미 다른 번역 작업이 진행 중입니다.")
            self.start_button.config(state=tk.NORMAL); self.stop_button.config(state=tk.DISABLED) # 버튼 상태 복원
            return

        self.translation_thread = threading.Thread(target=self._translation_worker, 
                                                   args=(file_path, src_lang, tgt_lang, model, task_log_filepath), # task_log_filepath 전달
                                                   daemon=True)
        self.start_time = time.time()
        self.translation_thread.start()
        self.update_progress_timer()

    def _translation_worker(self, file_path, src_lang, tgt_lang, model, task_log_filepath): # task_log_filepath 매개변수 추가
        output_path, translation_result_status = "", "실패"
        try:
            logger.debug("번역 작업자: 파일 정보 재확인...")
            # get_file_info는 OCR 핸들러가 필요함
            info_for_translation = self.pptx_handler.get_file_info(file_path, self.ocr_handler if self.ocr_handler else None)
            
            if not info_for_translation or (info_for_translation.get('text_elements',0) == 0 and info_for_translation.get('image_elements_with_text',0) == 0):
                logger.warning("번역할 텍스트 요소가 없습니다.")
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     self.master.after(0, lambda: messagebox.showinfo("정보", "파일에 번역할 텍스트 요소가 없습니다."))
                translation_result_status, output_path = "내용 없음", file_path # 원본 경로를 output으로 간주
                if hasattr(self, 'master') and self.master.winfo_exists():
                    self.master.after(0, self.translation_finished, translation_result_status, file_path, src_lang, tgt_lang, output_path, task_log_filepath) # task_log_filepath 전달
                return

            total_elements = info_for_translation.get('text_elements',0) + info_for_translation.get('image_elements_with_text',0)
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self.remaining_elements_label.config, {"text": f"남은 요소: {total_elements}"})
            
            font_code_for_render = UI_LANG_TO_FONT_CODE_MAP.get(tgt_lang, 'en') # 대상 언어에 맞는 폰트 코드
            
            output_path = self.pptx_handler.translate_presentation(
                file_path, src_lang, tgt_lang, 
                self.translator, self.ocr_handler, # ocr_handler 전달 (None일 수 있음)
                model, self.ollama_service, 
                font_code_for_render, task_log_filepath, # task_log_filepath 전달
                self.update_translation_progress, self.stop_event
            )

            if self.stop_event.is_set():
                logger.warning("번역 중지됨 (사용자 요청).")
                translation_result_status = "부분 성공 (중지)" if output_path and os.path.exists(output_path) else "취소됨"
            elif output_path and os.path.exists(output_path):
                elapsed_time = time.time() - self.start_time
                logger.info(f"번역 완료! 총 소요시간: {self._format_time(elapsed_time)}")
                translation_result_status = "성공"
                if hasattr(self, 'master') and self.master.winfo_exists() and not self.stop_event.is_set():
                     self.master.after(0, self._ask_open_folder, output_path)
            else: 
                logger.error("번역 실패 또는 결과 파일이 생성되지 않았습니다.")
                translation_result_status = "실패 (파일 없음)"
        except Exception as e: 
            logger.error(f"번역 작업 중 심각한 오류 발생: {e}", exc_info=True)
            translation_result_status = "오류 발생"
            # 작업 로그 파일에 오류 기록 (PptxHandler 내부에서도 기록되지만, 여기서도 간단히 남길 수 있음)
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_err:
                    f_err.write(f"\n--- 번역 작업 중 심각한 오류 발생 ---\n오류: {e}\n")
                    import traceback
                    traceback.print_exc(file=f_err)
            except Exception as ef_log:
                logger.error(f"작업 로그 파일에 오류 기록 실패: {ef_log}")

        finally:
            if hasattr(self, 'master') and self.master.winfo_exists():
                self.master.after(0, self.translation_finished, translation_result_status, file_path, src_lang, tgt_lang, output_path, task_log_filepath) # task_log_filepath 전달
            self.translation_thread = None # 스레드 종료 후 참조 제거


    def _ask_open_folder(self, path):
        if messagebox.askyesno("번역 완료", f"번역이 완료되었습니다.\n저장된 폴더를 여시겠습니까?\n{path}"):
            utils.open_folder(os.path.dirname(path))

    def _format_time(self, seconds):
        if seconds is None or seconds < 0: return "00:00.00"
        m, s = divmod(seconds, 60); return f"{int(m):02d}:{s:05.2f}"

    def update_translation_progress(self, current_slide, current_element_type, translated_count, total_elements, current_text=""):
        if self.stop_event.is_set(): return
        if total_elements > 0: progress = (translated_count / total_elements) * 100
        else: progress = 0
        
        elapsed_time = time.time() - (self.start_time if self.start_time else time.time())
        estimated_total_time = (elapsed_time / progress * 100) if progress > 0.1 else 0 # progress가 0이면 ZeroDivisionError 방지
        remaining_time = estimated_total_time - elapsed_time if estimated_total_time > elapsed_time else 0
        
        progress_text_val = f"{progress:.1f}% (진행: {self._format_time(elapsed_time)} / 남은예상: {self._format_time(remaining_time if remaining_time > 0 else 0)})"
        
        def _update_ui():
            if not (hasattr(self, 'progress_bar') and self.progress_bar.winfo_exists()): return
            self.progress_bar["value"] = progress
            self.progress_label_var.set(progress_text_val)
            self.current_slide_label.config(text=f"현재 슬라이드: {current_slide}")
            display_text = current_text if len(current_text) < 30 else current_text[:27] + "..."
            self.current_work_label.config(text=f"현재 작업: {current_element_type} - '{display_text}'")
            self.translated_elements_label.config(text=f"번역된 요소: {translated_count}")
            self.remaining_elements_label.config(text=f"남은 요소: {total_elements - translated_count}")
        
        if hasattr(self, 'master') and self.master.winfo_exists(): 
            self.master.after(0, _update_ui)

    def update_progress_timer(self):
        if self.translation_thread and self.translation_thread.is_alive() and not self.stop_event.is_set():
            # 타이머 업데이트는 진행 상황 콜백에서 이미 충분히 자주 발생하므로,
            # 별도의 타이머 업데이트 로직은 제거하거나 빈도를 줄여도 됨. 여기서는 유지.
            if hasattr(self, 'master') and self.master.winfo_exists(): 
                self.master.after(1000, self.update_progress_timer)


    def stop_translation(self):
        if self.translation_thread and self.translation_thread.is_alive():
            logger.warning("번역 중지 요청 중...")
            self.stop_event.set()
            self.stop_button.config(state=tk.DISABLED) # 중지 버튼 비활성화 (처리 중 표시)

    def translation_finished(self, result_status, original_file, src_lang, tgt_lang, translated_file_path, task_log_filepath): # task_log_filepath 추가
        if not (hasattr(self, 'start_button') and self.start_button.winfo_exists()): return
        
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        
        final_progress_text = self.progress_label_var.get() # 기본값
        if hasattr(self, 'start_time') and self.start_time:
            elapsed_time = time.time() - self.start_time
            if result_status == "성공": 
                final_progress_text = f"100% (총 소요시간: {self._format_time(elapsed_time)})"
            elif "중지" in result_status or "취소" in result_status: 
                final_progress_text = f"{self.progress_bar['value']:.1f}% (중지됨 - 소요시간: {self._format_time(elapsed_time)})"
            elif result_status == "내용 없음":
                final_progress_text = f"번역할 내용 없음 (소요시간: {self._format_time(elapsed_time)})"
            else: # 실패 또는 오류
                final_progress_text = f"오류 ({self.progress_bar['value']:.1f}% 진행 - 소요시간: {self._format_time(elapsed_time)})"
            self.progress_label_var.set(final_progress_text)

        if translated_file_path and os.path.exists(translated_file_path) and result_status not in ["취소됨", "오류 발생", "실패 (파일 없음)"]:
            self.translated_file_path_var.set(translated_file_path)
            self.open_folder_button.config(state=tk.NORMAL)
        else:
            self.translated_file_path_var.set("번역 실패 또는 파일 없음")
            self.open_folder_button.config(state=tk.DISABLED)
            if not (translated_file_path and os.path.exists(translated_file_path)) and result_status == "성공":
                 logger.warning(f"번역은 '성공'으로 기록되었으나, 결과 파일 경로가 유효하지 않음: {translated_file_path}")


        file_name = os.path.basename(original_file)
        current_time_str = time.strftime("%Y-%m-%d %H:%M:%S")
        # 히스토리에는 실제 번역된 파일 경로 또는 원본 파일 경로(실패 시) 저장, 작업 로그 경로도 추가할 수 있음
        history_entry_values = (file_name, src_lang, tgt_lang, result_status, current_time_str, translated_file_path or original_file) # 작업 로그 경로도 추가 가능
        TRANSLATION_HISTORY.append(history_entry_values)
        
        if hasattr(self, 'history_tree') and self.history_tree.winfo_exists():
            self.history_tree.insert("", tk.END, values=history_entry_values)
            self.history_tree.yview_moveto(1) # 마지막 항목으로 스크롤

        # 작업 로그 파일에 최종 상태 기록
        if task_log_filepath and os.path.exists(os.path.dirname(task_log_filepath)): # 로그 폴더가 있어야 함
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    f_task_log.write(f"\n--- 번역 작업 완료 ---\n")
                    f_task_log.write(f"최종 상태: {result_status}\n")
                    f_task_log.write(f"원본 파일: {original_file}\n")
                    if translated_file_path and os.path.exists(translated_file_path):
                        f_task_log.write(f"번역된 파일: {translated_file_path}\n")
                    f_task_log.write(f"총 소요 시간: {self._format_time(time.time() - self.start_time if self.start_time else 0)}\n")
                    logger.info(f"번역 최종 결과 '{result_status}'를 작업 로그 '{task_log_filepath}'에 기록했습니다.")
            except Exception as e_log_finish:
                logger.error(f"작업 로그 파일에 최종 상태 기록 실패: {e_log_finish}")
        
        # self.stop_event.clear() # _translation_worker 시작 시 clear하므로 여기서 필요X
        # self.translation_thread = None # _translation_worker finally에서 처리
        self.start_time = None


    def open_translated_folder(self):
        path = self.translated_file_path_var.get()
        if path and os.path.exists(path): 
            utils.open_folder(os.path.dirname(path))
        elif path: 
            messagebox.showwarning("폴더 열기 실패", f"경로를 찾을 수 없습니다: {path}")

    def on_history_double_click(self, event):
        if not (hasattr(self, 'history_tree') and self.history_tree.winfo_exists()): return
        item_id = self.history_tree.identify_row(event.y)
        if item_id:
            item_values = self.history_tree.item(item_id, "values")
            if item_values and len(item_values) > 5: # 경로 인덱스는 5
                file_path_to_open = item_values[5]
                if file_path_to_open and os.path.exists(file_path_to_open):
                    if messagebox.askyesno("파일 열기", f"번역된 파일 '{os.path.basename(file_path_to_open)}'을(를) 여시겠습니까?"):
                        try:
                            if platform.system() == "Windows": os.startfile(file_path_to_open)
                            elif platform.system() == "Darwin": subprocess.Popen(["open", file_path_to_open])
                            else: subprocess.Popen(["xdg-open", file_path_to_open])
                        except Exception as e: logger.error(f"히스토리 파일 열기 실패: {e}", exc_info=True)
                elif file_path_to_open: 
                    messagebox.showwarning("파일 없음", f"파일을 찾을 수 없습니다: {file_path_to_open}")

# tkinter Text 위젯으로 로그를 보내는 핸들러
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        if not (self.text_widget and self.text_widget.winfo_exists()):
            return # 위젯이 없거나 파괴된 경우 아무것도 하지 않음

        msg = self.format(record)
        def append_message():
            # 위젯이 파괴되었는지 다시 한번 확인 (after 콜백 실행 시점)
            if not (self.text_widget and self.text_widget.winfo_exists()):
                return
            self.text_widget.config(state=tk.NORMAL)
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.see(tk.END)
            self.text_widget.config(state=tk.DISABLED)
        
        # Tkinter 위젯 업데이트는 메인 스레드에서만 안전하게 수행
        # after(0)을 사용하여 현재 이벤트 루프가 끝난 후 즉시 실행하도록 예약
        self.text_widget.after(0, append_message)


if __name__ == "__main__":
    # 로그 폴더 생성 (Application 클래스 이전에도 확인/생성)
    if not os.path.exists(LOGS_DIR):
        try:
            os.makedirs(LOGS_DIR)
            print(f"로그 폴더 생성됨: {LOGS_DIR}")
        except Exception as e:
            print(f"메인 실행부: 로그 폴더 생성 실패: {LOGS_DIR}, 오류: {e}")
            
    if debug_mode: logger.info("디버그 모드로 실행 중입니다.")
    else: logger.info("일반 모드로 실행 중입니다.")
    
    if not os.path.exists(FONTS_DIR): logger.critical(f"필수 폰트 디렉토리를 찾을 수 없습니다: {FONTS_DIR}")
    else: logger.info(f"폰트 디렉토리 확인: {FONTS_DIR}")
    
    if not os.path.exists(ASSETS_DIR): logger.warning(f"에셋 디렉토리를 찾을 수 없습니다: {ASSETS_DIR}")
    else: logger.info(f"에셋 디렉토리 확인: {ASSETS_DIR}")

    root = tk.Tk()
    app = Application(master=root)
    root.geometry("960x780") # 창 크기 조정
    root.update_idletasks() # 실제 크기 계산
    min_width = root.winfo_reqwidth()
    min_height = root.winfo_reqheight()
    root.minsize(min_width + 20, min_height + 20) # 최소 크기 설정
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        logger.info("Ctrl+C로 애플리케이션 종료 중...")
        # on_closing이 이미 WM_DELETE_WINDOW에 연결되어 있으므로,
        # Tkinter의 기본 Ctrl+C 핸들링이 이를 호출할 수 있음.
        # 명시적으로 호출하려면 app.on_closing()을 호출할 수 있으나,
        # mainloop가 중단되면 Tk 객체가 불안정할 수 있으므로 주의.
        # 보통은 OS 레벨에서 프로세스가 종료됨.
    finally:
        logger.info(f"--- {APP_NAME} 종료됨 ---")