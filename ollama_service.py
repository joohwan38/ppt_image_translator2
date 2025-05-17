import shutil
import platform
import os
import subprocess
import requests
import psutil
import time
import logging
import json
from typing import Tuple, Optional, List
import threading

# 설정 파일 import
import config

logger = logging.getLogger(__name__)

# Ollama URL 및 타임아웃은 config에서 가져옴
# DEFAULT_OLLAMA_URL = config.DEFAULT_OLLAMA_URL
# OLLAMA_CONNECT_TIMEOUT = config.OLLAMA_CONNECT_TIMEOUT
# OLLAMA_READ_TIMEOUT = config.OLLAMA_READ_TIMEOUT
# OLLAMA_PULL_READ_TIMEOUT = config.OLLAMA_PULL_READ_TIMEOUT

class OllamaService:
    def __init__(self, url: str = None): # config에서 가져오므로 기본값 None 처리
        self.url = url if url is not None else config.DEFAULT_OLLAMA_URL
        self.connect_timeout = config.OLLAMA_CONNECT_TIMEOUT
        self.read_timeout = config.OLLAMA_READ_TIMEOUT
        self.pull_read_timeout = config.OLLAMA_PULL_READ_TIMEOUT
        logger.debug(f"OllamaService initialized with URL: {self.url}")


    def is_installed(self) -> bool:
        try:
            if shutil.which('ollama'):
                logger.debug("Ollama found in PATH via shutil.which")
                return True
            system = platform.system()
            if system == "Windows":
                paths_to_check = [
                    "C:\\Program Files\\Ollama\\ollama.exe",
                    os.path.expanduser("~\\AppData\\Local\\Ollama\\ollama.exe")
                ]
                for path in paths_to_check:
                    if os.path.exists(path):
                        logger.debug(f"Ollama found at: {path}")
                        return True
            elif system == "Darwin": # macOS
                paths_to_check = [
                    "/usr/local/bin/ollama",
                    "/opt/homebrew/bin/ollama",
                    "/Applications/Ollama.app/Contents/Resources/ollama"
                ]
                for path in paths_to_check:
                    if os.path.exists(path):
                        logger.debug(f"Ollama found at: {path}")
                        return True
            elif system == "Linux":
                paths_to_check = [
                    "/usr/local/bin/ollama",
                    "/usr/bin/ollama",
                    "/bin/ollama",
                    os.path.expanduser("~/.local/bin/ollama")
                ]
                for path in paths_to_check:
                    if os.path.exists(path):
                        logger.debug(f"Ollama found at {path}")
                        return True
            logger.debug("Ollama executable not found in common locations or PATH.")
            return False
        except Exception as e:
            logger.error(f"Ollama 설치 확인 오류: {e}", exc_info=True)
            return False


    def is_running(self) -> Tuple[bool, Optional[str]]:
        try:
            response = requests.get(f"{self.url}/api/tags", timeout=self.connect_timeout)
            if response.status_code == 200:
                port = self.url.split(':')[-1].split('/')[0]
                logger.debug(f"Ollama running, confirmed via API on port {port}")
                return True, port
        except requests.exceptions.RequestException as e:
            logger.debug(f"Ollama API check failed (this is okay, will try process check): {e}")

        try:
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                proc_info = proc.info
                if proc_info:
                    proc_name = proc_info.get('name', '').lower()
                    cmdline = proc_info.get('cmdline')
                    is_ollama_in_cmd = False
                    if cmdline and isinstance(cmdline, list):
                        is_ollama_in_cmd = any('ollama' in c.lower() for c in cmdline if isinstance(c, str))

                    if 'ollama' in proc_name or is_ollama_in_cmd:
                        logger.debug(f"Ollama process found: {proc_name} (PID: {proc_info.get('pid')}). Assuming default port if API failed.")
                        try:
                            port_from_url = self.url.split(':')[-1].split('/')[0]
                            if port_from_url.isdigit():
                                return True, port_from_url
                        except Exception:
                            pass
                        return True, "11434" # Default port if not extractable
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass # These are expected errors if the process doesn't exist or access is denied.
        except Exception as e:
            logger.error(f"Ollama 상태 확인 중 psutil 오류: {e}", exc_info=True)
        logger.debug("Ollama not detected as running by API or process check.")
        return False, None


    def start_ollama(self) -> bool:
        if not self.is_installed():
            logger.warning("Ollama가 설치되어 있지 않아 시작할 수 없습니다.")
            return False
        is_already_running, _ = self.is_running()
        if is_already_running:
            logger.info("Ollama가 이미 실행 중입니다.")
            return True
        try:
            logger.info("Ollama 시작 시도 ('ollama serve')...")
            cmd = ["ollama", "serve"]
            process_options = {
                'stdout': subprocess.DEVNULL,
                'stderr': subprocess.DEVNULL
            }
            if platform.system() == "Windows":
                process_options['creationflags'] = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
            else:
                process_options['start_new_session'] = True # For Unix-like systems
            
            subprocess.Popen(cmd, **process_options)
            
            # Wait a bit for Ollama to start
            for attempt in range(10): # Try for up to 10 seconds
                time.sleep(1)
                running, _ = self.is_running()
                if running:
                    logger.info(f"Ollama 시작 성공 (시도: {attempt + 1})")
                    return True
            logger.warning("Ollama 시작 시간 초과 (10초). 상태를 다시 확인해주세요.")
            return False
        except FileNotFoundError:
            logger.error("Ollama 실행 파일을 찾을 수 없습니다. PATH 설정을 확인하거나 Ollama를 올바르게 설치해주세요.")
            return False
        except Exception as e:
            logger.error(f"Ollama 시작 오류: {e}", exc_info=True)
            return False


    def get_text_models(self) -> List[str]:
        running, _ = self.is_running()
        if not running:
            logger.warning("Ollama가 실행 중이지 않아 모델 목록을 가져올 수 없습니다.")
            return []
        
        models = []
        try:
            response = requests.get(f"{self.url}/api/tags", timeout=(self.connect_timeout, self.read_timeout))
            response.raise_for_status() # Raise an exception for HTTP errors
            models_data = response.json()
            if 'models' in models_data and isinstance(models_data['models'], list):
                models = [model['name'] for model in models_data['models'] if isinstance(model, dict) and 'name' in model]
                logger.debug(f"Ollama 모델 목록 (API): {models}")
                return models
            else:
                logger.warning(f"Ollama 모델 목록 API 응답 형식이 올바르지 않음: {models_data}")
        except requests.exceptions.RequestException as e:
            logger.warning(f"Ollama 모델 목록 API 요청 중 예외 발생 (CLI 시도): {e}")
        except json.JSONDecodeError as e: # If response is not valid JSON
            logger.warning(f"Ollama 모델 목록 API 응답 JSON 디코딩 오류 (CLI 시도): {e}")

        # Fallback to CLI if API fails or returns unexpected format
        if self.is_installed(): # Check again if installed, as API might fail for other reasons
            try:
                result = subprocess.run(["ollama", "list"], capture_output=True, text=True, check=False, timeout=15)
                if result.returncode == 0:
                    lines = result.stdout.strip().split('\n')
                    # Expecting header line, then model lines
                    if len(lines) > 1: # Header + at least one model
                        # Assuming the first word of each line (after header) is the model name
                        cli_models = [line.split()[0] for line in lines[1:] if line.strip() and line.split()]
                        logger.debug(f"Ollama 모델 목록 (CLI): {cli_models}")
                        return cli_models
                else:
                    logger.warning(f"Ollama list 명령어 실행 실패 (종료 코드: {result.returncode}): {result.stderr.strip()}")
            except subprocess.TimeoutExpired:
                logger.warning("Ollama list 명령어 실행 시간 초과.")
            except FileNotFoundError: # Should have been caught by is_installed, but as a safeguard
                logger.warning("Ollama 명령어를 찾을 수 없어 CLI로 모델 목록을 가져올 수 없습니다.")
            except Exception as e:
                logger.error(f"Ollama list 명령어 실행 중 예외 발생: {e}", exc_info=True)
        
        if not models: # If models list is still empty after API and CLI attempts
            logger.warning("Ollama에서 모델 목록을 가져오지 못했습니다.")
        return models


    def pull_model_with_progress(self, model_name: str,
                                 progress_callback=None,
                                 stop_event: Optional[threading.Event] = None):
        running, _ = self.is_running()
        if not running:
            logger.warning(f"Ollama 미실행. {model_name} 모델 다운로드 불가.")
            if progress_callback: progress_callback("Ollama 서버 미실행", 0, 0, is_error=True)
            return False

        response = None
        try:
            logger.info(f"{model_name} 모델 다운로드 시작...")
            if progress_callback: progress_callback(f"{model_name} 다운로드 시작...", 0, 0)

            # For model pulling, use OLLAMA_PULL_READ_TIMEOUT
            current_pull_timeout = self.pull_read_timeout
            
            response = requests.post(
                f"{self.url}/api/pull",
                json={"name": model_name, "stream": True},
                stream=True,
                timeout=(self.connect_timeout, current_pull_timeout) # Use specific timeout for pull
            )
            response.raise_for_status()

            for line in response.iter_lines():
                if stop_event and stop_event.is_set():
                    logger.info(f"{model_name} 모델 다운로드 중지됨 (사용자 요청).")
                    if progress_callback: progress_callback("다운로드 중지됨", 0, 0, is_error=True)
                    return False

                if line:
                    try:
                        data = json.loads(line.decode('utf-8'))
                        status = data.get("status", "")
                        completed = data.get("completed", 0)
                        total = data.get("total", 0)

                        if "error" in data:
                            error_msg = f"모델 다운로드 오류 ({model_name}): {data['error']}"
                            logger.error(error_msg)
                            if progress_callback: progress_callback(error_msg, completed, total, is_error=True)
                            return False

                        if progress_callback:
                            progress_text = status
                            if total > 0 and "downloading" in status.lower(): # More robust check
                                progress_text = f"{status} ({completed/1024/1024:.1f}MB / {total/1024/1024:.1f}MB)"
                            elif "digest" in data and "completed" in data and "total" in data : # For layers
                                progress_text = f"레이어 처리 중... ({completed/1024/1024:.1f}MB / {total/1024/1024:.1f}MB)"
                            progress_callback(progress_text, completed, total)

                        if status.lower() == "success": # Case-insensitive check for success
                            logger.info(f"{model_name} 모델 다운로드 성공.")
                            if progress_callback: progress_callback("다운로드 완료", total if total else completed, total if total else completed)
                            return True

                    except json.JSONDecodeError:
                        logger.debug(f"JSON 디코딩 오류 (무시 가능, 스트림 라인): {line.decode('utf-8', errors='ignore')}")
                    except Exception as e_stream_proc:
                        error_msg = f"모델 다운로드 스트림 처리 중 예외 ({model_name}): {e_stream_proc}"
                        logger.error(error_msg, exc_info=True)
                        if progress_callback: progress_callback(error_msg, 0,0, is_error=True)
                        return False

            logger.warning(f"{model_name} 모델 다운로드 확인 실패 (스트림 종료, 'success' 메시지 없음).")
            if progress_callback: progress_callback("다운로드 확인 실패", 0, 0, is_error=True)
            return False

        except requests.exceptions.Timeout as e_timeout: # Specific timeout handling
            error_msg = f"Ollama 모델 다운로드 요청 시간 초과 ({model_name}): {e_timeout}"
            logger.error(error_msg, exc_info=True)
            if progress_callback: progress_callback(error_msg, 0, 0, is_error=True)
            return False
        except requests.exceptions.RequestException as e_req:
            error_msg = f"Ollama 모델 다운로드 요청 오류 ({model_name}): {e_req}"
            logger.error(error_msg, exc_info=True)
            if progress_callback: progress_callback(error_msg, 0, 0, is_error=True)
            return False
        except Exception as e_pull:
            error_msg = f"Ollama 모델 다운로드 중 예측하지 못한 오류 ({model_name}): {e_pull}"
            logger.error(error_msg, exc_info=True)
            if progress_callback: progress_callback(error_msg, 0, 0, is_error=True)
            return False
        finally:
            if response:
                try:
                    response.close()
                    logger.debug(f"Ollama pull API 응답 스트림 닫힘 ({model_name}).")
                except Exception as e_close:
                    logger.warning(f"Ollama pull API 응답 스트림 닫기 중 오류 ({model_name}): {e_close}")