import subprocess
import os
import platform
import sys
import logging

logger = logging.getLogger(__name__)

def check_paddleocr():
    """PaddleOCR 설치 여부를 확인합니다."""
    try:
        import paddleocr # import 시도
        # 간단히 PaddleOCR 클래스 인스턴스 생성 시도 (실제 사용 가능 여부 확인)
        # 초기화가 오래 걸릴 수 있으므로, 더 빠른 확인 방법 고려 가능 (예: 특정 파일 존재 여부)
        # 여기서는 import 성공 여부만 우선 확인
        logger.debug("paddleocr 모듈 import 성공.")
        return True
    except ImportError:
        logger.warning("paddleocr 모듈을 찾을 수 없습니다. (미설치)")
        return False
    except Exception as e:
        logger.error(f"PaddleOCR 확인 중 예상치 못한 오류 (설치는 되어있으나 초기화 문제 가능성): {e}", exc_info=True)
        return False

def install_paddleocr():
    """PaddleOCR을 pip를 사용하여 설치합니다."""
    try:
        logger.info("PaddleOCR 자동 설치 시도 중... (paddlepaddle, paddleocr)")
        # CPU 버전 우선 설치. GPU 버전은 사용자 환경에 따라 별도 안내 필요.
        # Tsinghua mirror 사용 (중국 지역 사용자에게 빠를 수 있음)
        # subprocess.check_call([sys.executable, "-m", "pip", "install", "paddlepaddle", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"])
        # subprocess.check_call([sys.executable, "-m", "pip", "install", "paddleocr", "-i", "https://pypi.tuna.tsinghua.edu.cn/simple"])
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

def open_folder(path):
    """지정된 폴더를 운영체제의 파일 탐색기로 엽니다."""
    try:
        if not os.path.isdir(path): # 폴더가 아닌 파일 경로가 올 경우 디렉토리만 추출
            path = os.path.dirname(path)
            if not os.path.isdir(path): # 그래도 폴더가 아니면
                logger.warning(f"폴더 열기 실패: 유효한 디렉토리 경로가 아님 - {path}")
                return

        logger.info(f"폴더 열기: {path}")
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin": # macOS
            subprocess.Popen(["open", path])
        else: # Linux and other Unix-like
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        logger.error(f"폴더 열기 중 오류: {e}", exc_info=True)