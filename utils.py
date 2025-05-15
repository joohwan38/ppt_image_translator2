import subprocess
import os
import platform
import sys
import logging

logger = logging.getLogger(__name__)

def check_paddleocr():
    """PaddleOCR 설치 여부를 확인합니다."""
    try:
        import paddleocr # 패키지명 paddleocr (라이브러리 paddlepaddle)
        logger.debug("paddleocr 모듈 import 성공.")
        return True
    except ImportError:
        logger.warning("paddleocr 모듈을 찾을 수 없습니다. (미설치)")
        return False
    except Exception as e:
        logger.error(f"PaddleOCR 확인 중 예상치 못한 오류: {e}", exc_info=True)
        return False

def install_paddleocr():
    """PaddleOCR을 pip를 사용하여 설치합니다."""
    try:
        logger.info("PaddleOCR 자동 설치 시도 중... (paddlepaddle, paddleocr)")
        # CPU 버전 우선 설치
        subprocess.check_call([sys.executable, "-m", "pip", "install", "paddlepaddle"])
        subprocess.check_call([sys.executable, "-m", "pip", "install", "paddleocr"])
        logger.info("PaddleOCR 설치 성공 (또는 이미 설치됨).")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"PaddleOCR 설치 실패 (pip 오류): {e}", exc_info=True)
        return False
    except Exception as e:
        logger.error(f"PaddleOCR 설치 중 예기치 않은 오류: {e}", exc_info=True)
        return False

def check_easyocr():
    """EasyOCR 설치 여부를 확인합니다."""
    try:
        import easyocr
        logger.debug("easyocr 모듈 import 성공.")
        return True
    except ImportError:
        logger.warning("easyocr 모듈을 찾을 수 없습니다. (미설치)")
        return False
    except Exception as e:
        logger.error(f"EasyOCR 확인 중 예상치 못한 오류: {e}", exc_info=True)
        return False

def install_easyocr():
    """EasyOCR을 pip를 사용하여 설치합니다."""
    try:
        logger.info("EasyOCR 자동 설치 시도 중... (easyocr)")
        # EasyOCR은 PyTorch 등 의존성이 있을 수 있음.
        # 설치 시 필요한 의존성이 자동으로 설치되지만, 환경에 따라 추가 설정이 필요할 수 있음.
        subprocess.check_call([sys.executable, "-m", "pip", "install", "easyocr"])
        logger.info("EasyOCR 설치 성공 (또는 이미 설치됨).")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"EasyOCR 설치 실패 (pip 오류): {e}", exc_info=True)
        return False
    except Exception as e:
        logger.error(f"EasyOCR 설치 중 예기치 않은 오류: {e}", exc_info=True)
        return False

def open_folder(path):
    # (이전 답변과 동일)
    try:
        if not os.path.isdir(path):
            path = os.path.dirname(path)
            if not os.path.isdir(path):
                logger.warning(f"폴더 열기 실패: 유효한 디렉토리 경로가 아님 - {path}")
                return
        logger.info(f"폴더 열기: {path}")
        if platform.system() == "Windows": os.startfile(path)
        elif platform.system() == "Darwin": subprocess.Popen(["open", path])
        else: subprocess.Popen(["xdg-open", path])
    except Exception as e:
        logger.error(f"폴더 열기 중 오류: {e}", exc_info=True)