# config.py
import os
import logging

# --- Base Directories ---
# 이 프로젝트의 루트 디렉터리를 기준으로 합니다.
# config.py 파일이 프로젝트 루트에 있다고 가정합니다.
PROJECT_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Application Settings ---
APP_NAME = "Powerpoint Document Translator"
DEFAULT_OLLAMA_MODEL = "gemma3:12b" # Ollama 번역 모델
SUPPORTED_LANGUAGES = ["한국어", "일본어", "영어", "중국어", "대만어", "태국어", "스페인어"]

# --- Path Settings (PROJECT_ROOT_DIR 기준 상대 경로) ---
ASSETS_DIR_NAME = "assets"
FONTS_DIR_NAME = "fonts"
LOGS_DIR_NAME = "logs"
HISTORY_DIR_NAME = "hist" # 번역 히스토리 저장 폴더명

ASSETS_DIR = os.path.join(PROJECT_ROOT_DIR, ASSETS_DIR_NAME)
FONTS_DIR = os.path.join(PROJECT_ROOT_DIR, FONTS_DIR_NAME)
LOGS_DIR = os.path.join(PROJECT_ROOT_DIR, LOGS_DIR_NAME)
HISTORY_DIR = os.path.join(PROJECT_ROOT_DIR, HISTORY_DIR_NAME) # 번역 히스토리 저장 경로

# --- Logging Configuration ---
DEFAULT_LOG_LEVEL = logging.INFO
DEBUG_LOG_LEVEL = logging.DEBUG

# --- Translation Weights ---
WEIGHT_TEXT_CHAR = 1
WEIGHT_IMAGE = 100
WEIGHT_CHART = 15

# --- OCR Configuration ---
EASYOCR_SUPPORTED_UI_LANGS = ["일본어", "태국어", "스페인어"]
UI_LANG_TO_PADDLEOCR_CODE_MAP = {
    "한국어": "korean", "영어": "en",
    "중국어": "ch_doc",
    "대만어": "chinese_cht",
}
UI_LANG_TO_EASYOCR_CODE_MAP = {
    "일본어": "ja", "태국어": "th", "스페인어": "es"
}
DEFAULT_PADDLE_OCR_LANG = "korean"

OCR_LANGUAGE_FONT_MAP = {
    'korean': 'NotoSansCJK-Regular.ttc', 'japan': 'NotoSansCJK-Regular.ttc',
    'ch': 'NotoSansCJK-Regular.ttc', 'chinese_cht': 'NotoSansCJK-Regular.ttc',
    'en': 'NotoSansCJK-Regular.ttc', 'th': 'NotoSansThai-VariableFont_wdth,wght.ttf',
    'es': 'NotoSansCJK-Regular.ttc',
    'korean_bold': 'NotoSansCJK-Bold.ttc', 'japan_bold': 'NotoSansCJK-Bold.ttc',
    'ch_bold': 'NotoSansCJK-Bold.ttc', 'chinese_cht_bold': 'NotoSansCJK-Bold.ttc',
    'en_bold': 'NotoSansCJK-Bold.ttc', 'th_bold': 'NotoSansThai-VariableFont_wdth,wght.ttf',
    'es_bold': 'NotoSansCJK-Bold.ttc',
}
OCR_DEFAULT_FONT_FILENAME = 'NotoSansCJK-Regular.ttc'
OCR_DEFAULT_BOLD_FONT_FILENAME = 'NotoSansCJK-Bold.ttc'

# --- Ollama Service Configuration (for ollama_service.py) ---
DEFAULT_OLLAMA_URL = "http://localhost:11434"
OLLAMA_CONNECT_TIMEOUT = 5  # seconds
OLLAMA_READ_TIMEOUT = 180   # seconds for general API calls
OLLAMA_PULL_READ_TIMEOUT = None # 모델 다운로드는 매우 오래 걸릴 수 있음 (None은 무제한 대기)
MODELS_CACHE_TTL_SECONDS = 300 # 모델 목록 API 결과 캐시 시간 (초), 예: 5분

# --- Translator Configuration (for translator.py) ---
TRANSLATOR_TEMPERATURE_GENERAL = 0.2 # 텍스트 번역 기본 온도
# MAX_TRANSLATION_WORKERS 값을 더 넉넉하게 설정 (예: 8 또는 16).
# 로컬 Ollama 서버의 성능과 사용자 시스템 환경에 따라 조절 필요.
# 너무 높으면 Ollama 서버에 부담을 주거나 컨텍스트 스위칭 오버헤드 발생 가능.
# 여기서는 예시로 8로 설정.
MAX_TRANSLATION_WORKERS = 8
# OCR 병렬 처리를 위한 워커 수 (번역 워커와 별도로 설정하거나 공유 가능)
# 여기서는 번역 워커와 동일하게 설정
MAX_OCR_WORKERS = MAX_TRANSLATION_WORKERS


# --- PPTX Handler Configuration (for pptx_handler.py) ---
MIN_MEANINGFUL_CHAR_RATIO_SKIP = 0.1
MIN_MEANINGFUL_CHAR_RATIO_OCR = 0.1

# --- Main UI Configuration ---
UI_LANG_TO_FONT_CODE_MAP = {
    "한국어": "korean", "일본어": "japan", "영어": "en",
    "중국어": "ch_doc", "대만어": "chinese_cht", "태국어": "th", "스페인어": "es",
}
MAX_HISTORY_ITEMS = 50 # 번역 히스토리 최대 저장 개수

# --- Advanced Options Defaults ---
DEFAULT_OCR_TEMPERATURE = 0.4
DEFAULT_IMAGE_TRANSLATION_ENABLED = True
DEFAULT_OCR_USE_GPU = False