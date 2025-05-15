import shutil
import platform
import os
import subprocess
import requests
import psutil # pip install psutil
import time
import logging
import json
from typing import Tuple, Optional, List

logger = logging.getLogger(__name__)

DEFAULT_OLLAMA_URL = "http://localhost:11434"
OLLAMA_CONNECT_TIMEOUT = 5
OLLAMA_READ_TIMEOUT = 180
OLLAMA_PULL_READ_TIMEOUT = None # 모델 다운로드는 매우 오래 걸릴 수 있음

class OllamaService:
    def __init__(self, url: str = DEFAULT_OLLAMA_URL):
        self.url = url
        self.connect_timeout = OLLAMA_CONNECT_TIMEOUT
        self.read_timeout = OLLAMA_READ_TIMEOUT
        self.pull_read_timeout = OLLAMA_PULL_READ_TIMEOUT

    def is_installed(self) -> bool:
        """Ollama 설치 여부 확인"""
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
                    "/opt/homebrew/bin/ollama", # Homebrew on Apple Silicon
                    "/Applications/Ollama.app/Contents/Resources/ollama" # Official app
                ]
                for path in paths_to_check:
                    if os.path.exists(path):
                        logger.debug(f"Ollama found at: {path}")
                        return True
            elif system == "Linux":
                # Common paths for Linux
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
        """Ollama 실행 상태 및 포트 확인"""
        try:
            # API를 통한 확인이 가장 확실함
            response = requests.get(f"{self.url}/api/tags", timeout=self.connect_timeout)
            if response.status_code == 200:
                port = self.url.split(':')[-1].split('/')[0]
                logger.debug(f"Ollama running, confirmed via API on port {port}")
                return True, port
        except requests.exceptions.RequestException as e:
            logger.debug(f"Ollama API check failed (this is okay, will try process check): {e}")
            # API 실패는 일반적인 상황일 수 있으므로, 다음 프로세스 확인으로 넘어감

        # API 확인 실패 시, 프로세스 목록 확인 (차선책)
        try:
            for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
                proc_info = proc.info # 한 번만 호출하여 None 가능성 줄임
                if proc_info: # proc.info가 None이 아닌지 확인
                    proc_name = proc_info.get('name', '').lower()
                    cmdline = proc_info.get('cmdline') # 기본값을 빈 리스트로 하지 않고, None 가능성 확인

                    # cmdline이 None이거나 빈 리스트가 아닐 경우에만 반복문 실행
                    is_ollama_in_cmd = False
                    if cmdline and isinstance(cmdline, list): # cmdline이 유효한 리스트인지 확인
                        is_ollama_in_cmd = any('ollama' in c.lower() for c in cmdline if isinstance(c, str))

                    if 'ollama' in proc_name or is_ollama_in_cmd:
                        logger.debug(f"Ollama process found: {proc_name} (PID: {proc_info.get('pid')}). Assuming default port if API failed.")
                        # API가 실패했으므로 URL에서 포트 추출 시도 또는 기본값 사용
                        try:
                            port_from_url = self.url.split(':')[-1].split('/')[0]
                            if port_from_url.isdigit():
                                return True, port_from_url
                        except Exception:
                            pass # URL에서 포트 추출 실패 시 기본 포트 반환 고려
                        return True, "11434" # 기본 포트 반환 또는 None
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # 이러한 예외는 프로세스 반복 중에 발생할 수 있으며, 무시하고 계속 진행
            pass
        except Exception as e:
            # 여기서 발생한 오류는 psutil 사용 자체의 문제일 수 있음
            logger.error(f"Ollama 상태 확인 중 psutil 오류: {e}", exc_info=True)
            # 이 경우, psutil을 통한 확인은 실패로 간주

        logger.debug("Ollama not detected as running by API or process check.")
        return False, None

    def start_ollama(self) -> bool:
        """Ollama 서버 시작 (백그라운드 'ollama serve' 실행 시도)"""
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
            
            # 백그라운드 실행 및 출력 숨김 설정
            process_options = {
                'stdout': subprocess.DEVNULL,
                'stderr': subprocess.DEVNULL
            }
            if platform.system() == "Windows":
                # CREATE_NEW_PROCESS_GROUP: 부모 프로세스와 독립된 새 프로세스 그룹 생성
                # DETACHED_PROCESS: 부모 콘솔에서 분리 (터미널 창이 뜨지 않도록)
                process_options['creationflags'] = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS
            else: # macOS, Linux
                # start_new_session=True: 새 세션에서 프로세스 시작 (데몬화와 유사)
                process_options['start_new_session'] = True

            subprocess.Popen(cmd, **process_options)
            
            # Ollama 서버가 시작될 시간을 줌 (최대 10초, 1초 간격으로 확인)
            for attempt in range(10):
                time.sleep(1) 
                running, _ = self.is_running() # API 및 프로세스 재확인
                if running:
                    logger.info(f"Ollama 시작 성공 (시도: {attempt + 1})")
                    return True
            logger.warning("Ollama 시작 시간 초과 (10초). 상태를 다시 확인해주세요.")
            return False # 시간 초과 시 실패로 간주
        except FileNotFoundError:
            logger.error("Ollama 실행 파일을 찾을 수 없습니다. PATH 설정을 확인하거나 Ollama를 올바르게 설치해주세요.")
            return False
        except Exception as e:
            logger.error(f"Ollama 시작 오류: {e}", exc_info=True)
            return False

    def get_text_models(self) -> List[str]:
        """설치된 (Ollama에 로드된) 텍스트 모델 목록 가져오기"""
        running, _ = self.is_running()
        if not running:
            logger.warning("Ollama가 실행 중이지 않아 모델 목록을 가져올 수 없습니다.")
            return []
        
        models = []
        # 1. API를 통한 모델 목록 가져오기 (가장 선호)
        try:
            response = requests.get(f"{self.url}/api/tags", timeout=(self.connect_timeout, self.read_timeout))
            response.raise_for_status() # 오류 발생 시 예외 발생
            models_data = response.json()
            if 'models' in models_data and isinstance(models_data['models'], list):
                models = [model['name'] for model in models_data['models'] if isinstance(model, dict) and 'name' in model]
                logger.debug(f"Ollama 모델 목록 (API): {models}")
                return models
            else:
                logger.warning(f"Ollama 모델 목록 API 응답 형식이 올바르지 않음: {models_data}")
        except requests.exceptions.RequestException as e:
            logger.warning(f"Ollama 모델 목록 API 요청 중 예외 발생 (CLI 시도): {e}")
        except json.JSONDecodeError as e:
            logger.warning(f"Ollama 모델 목록 API 응답 JSON 디코딩 오류 (CLI 시도): {e}")

        # 2. API 실패 시 CLI를 통한 모델 목록 가져오기 (차선책)
        if self.is_installed(): # Ollama가 설치되어 있어야 CLI 사용 가능
            try:
                result = subprocess.run(["ollama", "list"], capture_output=True, text=True, check=False, timeout=15) # 타임아웃 증가
                if result.returncode == 0:
                    lines = result.stdout.strip().split('\n')
                    if len(lines) > 1: # 첫 줄은 헤더
                        # NAME            ID              SIZE    MODIFIED
                        # gemma:7b        f50c60f258e7    5.0 GB  3 weeks ago
                        # 공백으로 분리된 첫 번째 항목이 모델 이름
                        cli_models = [line.split()[0] for line in lines[1:] if line.strip() and line.split()]
                        logger.debug(f"Ollama 모델 목록 (CLI): {cli_models}")
                        return cli_models
                else:
                    logger.warning(f"Ollama list 명령어 실행 실패 (종료 코드: {result.returncode}): {result.stderr.strip()}")
            except subprocess.TimeoutExpired:
                logger.warning("Ollama list 명령어 실행 시간 초과.")
            except FileNotFoundError: # 'ollama' 명령어를 찾을 수 없는 경우
                logger.warning("Ollama 명령어를 찾을 수 없어 CLI로 모델 목록을 가져올 수 없습니다.")
            except Exception as e:
                logger.error(f"Ollama list 명령어 실행 중 예외 발생: {e}", exc_info=True)
        
        if not models: # API와 CLI 모두 실패한 경우
            logger.warning("Ollama에서 모델 목록을 가져오지 못했습니다.")
        return models # 빈 리스트 반환

    def pull_model_with_progress(self, model_name, progress_callback=None):
        running, _ = self.is_running()
        if not running:
            logger.warning(f"Ollama 미실행. {model_name} 모델 다운로드 불가.")
            if progress_callback: progress_callback("Ollama 서버 미실행", 0, 0, is_error=True)
            return False
        try:
            logger.info(f"{model_name} 모델 다운로드 시작...")
            if progress_callback: progress_callback(f"{model_name} 다운로드 시작...", 0, 0)
            
            # 스트리밍 요청
            response = requests.post(
                f"{self.url}/api/pull",
                json={"name": model_name, "stream": True},
                stream=True,
                timeout=(self.connect_timeout, self.pull_read_timeout) # 연결 타임아웃, 읽기 타임아웃(스트림은 길게)
            )
            response.raise_for_status() # HTTP 오류 발생 시 예외

            for line in response.iter_lines():
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
                            return False # 오류 발생 시 즉시 반환
                        
                        if progress_callback:
                            progress_text = status
                            # "pulling manifest", "verifying sha256", "writing manifest", "removing any unused layers", "success"
                            if total > 0 and "downloading" in status: # 다운로드 중일 때만 용량 표시
                                progress_text = f"{status} ({completed/1024/1024:.1f}MB / {total/1024/1024:.1f}MB)"
                            elif "digest" in data and "completed" in data and "total" in data : # 레이어 다운로드 시
                                progress_text = f"레이어 다운로드 중... ({completed/1024/1024:.1f}MB / {total/1024/1024:.1f}MB)"

                            progress_callback(progress_text, completed, total)
                        
                        # "status": "success"는 최종 성공 메시지
                        if status == "success":
                            logger.info(f"{model_name} 모델 다운로드 성공.")
                            if progress_callback: progress_callback("다운로드 완료", total if total else completed, total if total else completed) # 최종 상태
                            return True
                            
                    except json.JSONDecodeError:
                        # 간혹 빈 줄이나 불완전한 JSON이 올 수 있으므로, 디버그 레벨로 로깅하고 무시
                        logger.debug(f"JSON 디코딩 오류 (무시 가능, 스트림 라인): {line.decode('utf-8', errors='ignore')}")
                    except Exception as e_stream_proc: # 스트림 처리 중 예외
                        error_msg = f"모델 다운로드 스트림 처리 중 예외 ({model_name}): {e_stream_proc}"
                        logger.error(error_msg, exc_info=True)
                        if progress_callback: progress_callback(error_msg, 0,0, is_error=True)
                        return False # 스트림 처리 중 문제 발생 시 실패로 간주

            # 스트림이 정상적으로 종료되었으나 "success" 메시지를 못 받은 경우 (이론상 발생하기 어려움)
            logger.warning(f"{model_name} 모델 다운로드 확인 실패 (스트림 종료, 'success' 메시지 없음).")
            if progress_callback: progress_callback("다운로드 확인 실패", 0, 0, is_error=True)
            return False

        except requests.exceptions.RequestException as e_req: # 요청 관련 예외 (연결, 타임아웃 등)
            error_msg = f"Ollama 모델 다운로드 요청 오류 ({model_name}): {e_req}"
            logger.error(error_msg, exc_info=True)
            if progress_callback: progress_callback(error_msg, 0, 0, is_error=True)
            return False
        except Exception as e_pull: # 그 외 모든 예외
            error_msg = f"Ollama 모델 다운로드 중 예측하지 못한 오류 ({model_name}): {e_pull}"
            logger.error(error_msg, exc_info=True)
            if progress_callback: progress_callback(error_msg, 0, 0, is_error=True)
            return False