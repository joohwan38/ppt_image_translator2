# utils.py
import subprocess
import os
import platform
import sys
import logging
from typing import Optional, Callable, IO # IO 추가

logger = logging.getLogger(__name__)

def check_paddleocr():
    """PaddleOCR 설치 여부를 확인합니다."""
    try:
        import paddleocr
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
        subprocess.check_call([sys.executable, "-m", "pip", "install", "easyocr"])
        logger.info("EasyOCR 설치 성공 (또는 이미 설치됨).")
        return True
    except subprocess.CalledProcessError as e:
        logger.error(f"EasyOCR 설치 실패 (pip 오류): {e}", exc_info=True)
        return False
    except Exception as e:
        logger.error(f"EasyOCR 설치 중 예기치 않은 오류: {e}", exc_info=True)
        return False

def open_folder(path: str):
    """주어진 경로의 폴더를 엽니다."""
    try:
        if not os.path.isdir(path):
            path = os.path.dirname(path)
            if not os.path.isdir(path):
                logger.warning(f"폴더 열기 실패: 유효한 디렉토리 경로가 아님 - {path}")
                return
        logger.info(f"폴더 열기: {path}")
        if platform.system() == "Windows":
            os.startfile(path)
        elif platform.system() == "Darwin": # macOS
            subprocess.Popen(["open", path])
        else: # Linux 등
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        logger.error(f"폴더 열기 중 오류: {e}", exc_info=True)

def setup_task_logging(task_log_filepath: str,
                       initial_message_lines: Optional[list[str]] = None
                       ) -> tuple[Optional[IO[str]], Optional[Callable[[str], None]]]:
    """
    작업별 로그 파일을 설정하고 로깅 함수를 반환합니다.

    Args:
        task_log_filepath: 작업 로그 파일의 전체 경로.
        initial_message_lines: 로그 파일 생성 시 초기에 기록할 메시지 목록.

    Returns:
        Tuple (파일 객체, 로그 함수). 파일 열기 실패 시 (None, None).
    """
    f_task_log = None
    log_func = None
    try:
        # 로그 파일 디렉토리 생성 (필요시)
        log_dir = os.path.dirname(task_log_filepath)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)

        f_task_log = open(task_log_filepath, 'a', encoding='utf-8')
        if initial_message_lines:
            for line in initial_message_lines:
                f_task_log.write(line + "\n")
            f_task_log.flush()

        def write_log(message: str):
            if f_task_log and not f_task_log.closed:
                f_task_log.write(message + "\n")
                f_task_log.flush()
        log_func = write_log
        logger.info(f"작업 로그 파일 설정 완료: {task_log_filepath}")

    except Exception as e_log_open:
        logger.error(f"작업 로그 파일 ({task_log_filepath}) 열기/설정 실패: {e_log_open}")
        if f_task_log: # 만약 파일은 열렸으나 다른 오류 발생 시 닫기 시도
            try: f_task_log.close()
            except Exception: pass
        f_task_log = None
        log_func = None

    return f_task_log, log_func