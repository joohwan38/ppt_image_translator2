# config.py
import os
import logging

# --- Base Directories ---
PROJECT_ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Application Settings ---
APP_NAME = "Powerpoint Document Translator"
DEFAULT_OLLAMA_MODEL = "gemma3:12b"
SUPPORTED_LANGUAGES = ["한국어", "일본어", "영어", "중국어", "대만어", "태국어", "스페인어"]
USER_SETTINGS_FILENAME = "user_settings.json"

# --- Path Settings ---
ASSETS_DIR_NAME = "assets"
FONTS_DIR_NAME = "fonts"
LOGS_DIR_NAME = "logs"
HISTORY_DIR_NAME = "hist"

ASSETS_DIR = os.path.join(PROJECT_ROOT_DIR, ASSETS_DIR_NAME)
FONTS_DIR = os.path.join(PROJECT_ROOT_DIR, FONTS_DIR_NAME)
LOGS_DIR = os.path.join(PROJECT_ROOT_DIR, LOGS_DIR_NAME)
HISTORY_DIR = os.path.join(PROJECT_ROOT_DIR, HISTORY_DIR_NAME)

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

# --- Ollama Service Configuration ---
DEFAULT_OLLAMA_URL = "http://localhost:11434"
OLLAMA_CONNECT_TIMEOUT = 5
OLLAMA_READ_TIMEOUT = 180
OLLAMA_PULL_READ_TIMEOUT = None
MODELS_CACHE_TTL_SECONDS = 300

# --- Translator Configuration ---
TRANSLATOR_TEMPERATURE_GENERAL = 0.2
MAX_TRANSLATION_WORKERS = 4
MAX_OCR_WORKERS = MAX_TRANSLATION_WORKERS

# --- PPTX Handler Configuration ---
MIN_MEANINGFUL_CHAR_RATIO_SKIP = 0.1
MIN_MEANINGFUL_CHAR_RATIO_OCR = 0.1

# --- Main UI Configuration ---
UI_LANG_TO_FONT_CODE_MAP = {
    "한국어": "korean", "일본어": "japan", "영어": "en",
    "중국어": "ch_doc", "대만어": "chinese_cht", "태국어": "th", "스페인어": "es",
}
MAX_HISTORY_ITEMS = 50
UI_PROGRESS_UPDATE_INTERVAL = 0.2 # 초 단위, UI 진행률 업데이트 최소 간격

# --- Advanced Options Defaults ---
DEFAULT_ADVANCED_SETTINGS = {
    "ocr_temperature": 0.4,
    "image_translation_enabled": True,
    "ocr_use_gpu": False
}

# --- XML Namespaces (공통 사용 가능성) ---
# 필요시 ChartXmlHandler에서 여기로 옮기거나 utils.py에 함수로 정의
# XML_NAMESPACES = {
#     'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
#     'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
#     # ... 기타 네임스페이스
# }