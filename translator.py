# translator.py
import logging
import time
from typing import TYPE_CHECKING, Optional, List, Dict
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import hashlib

# 설정 파일 import
import config

if TYPE_CHECKING:
    from ollama_service import OllamaService

logger = logging.getLogger(__name__)

MAX_TRANSLATION_WORKERS = config.MAX_TRANSLATION_WORKERS

class OllamaTranslator:
    def __init__(self):
        self.translation_cache: Dict[str, str] = {}
        logger.info(f"OllamaTranslator 초기화됨. 번역 작업자 수: {MAX_TRANSLATION_WORKERS}")

    def _get_cache_key(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str) -> str:
        key_string = f"{src_lang_ui_name}|{tgt_lang_ui_name}|{model_name}|{text_to_translate}"
        return hashlib.md5(key_string.encode('utf-8')).hexdigest()

    def translate_text(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str,
                       model_name: str, ollama_service_instance: 'OllamaService',
                       is_ocr_text: bool = False, ocr_temperature: Optional[float] = None) -> str:
        if not text_to_translate or not text_to_translate.strip():
            return text_to_translate if text_to_translate else ""

        cache_key = self._get_cache_key(text_to_translate, src_lang_ui_name, tgt_lang_ui_name, model_name)
        if cache_key in self.translation_cache:
            cached_result = self.translation_cache[cache_key]
            # logger.debug(f"번역 캐시 사용 (키: {cache_key}): '{text_to_translate[:30]}...' -> '{cached_result[:30]}...'")
            return cached_result

        prompt = f"Translate the following text from {src_lang_ui_name} to {tgt_lang_ui_name}. Provide only the translated text itself, without any additional explanations, introductory phrases, or quotation marks around the translation. Text to translate:\n\n{text_to_translate}"

        current_temperature = config.TRANSLATOR_TEMPERATURE_GENERAL
        if is_ocr_text and ocr_temperature is not None:
            current_temperature = ocr_temperature
            logger.debug(f"OCR 텍스트 번역, 온도: {current_temperature} 적용 (원본: {text_to_translate[:20]}...)")


        # --- 1단계 개선: 에러 메시지 형식 통일화 ---
        text_snippet_for_error = text_to_translate[:20].replace('\n', ' ') + "..."

        try:
            if not ollama_service_instance or not ollama_service_instance.is_running()[0]:
                error_msg_server = f"Ollama 서버 미실행. {model_name} 모델로 번역 불가."
                logger.error(error_msg_server)
                return f"오류: Ollama 서버 연결 실패 - \"{text_snippet_for_error}\"" # 형식 변경

            api_url = f"{ollama_service_instance.url}/api/generate"
            payload = {
                "model": model_name,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": current_temperature
                }
            }

            start_time = time.time()
            response = requests.post(api_url, json=payload,
                                     timeout=(ollama_service_instance.connect_timeout, ollama_service_instance.read_timeout))
            response.raise_for_status()
            response_data = response.json()
            end_time = time.time()

            elapsed_time = end_time - start_time
            if response_data and "response" in response_data:
                translated_text = response_data["response"].strip()
                if translated_text and "오류:" not in translated_text :
                    self.translation_cache[cache_key] = translated_text
                    logger.debug(f"번역 완료 (모델: {model_name}, 온도: {current_temperature}, 소요: {elapsed_time:.2f}s): '{text_to_translate[:30]}...' -> '{translated_text[:30]}...'")
                else: # 오류 포함 또는 빈 결과
                    logger.warning(f"번역 결과에 오류 포함 또는 빈 결과 (모델: {model_name}, 소요: {elapsed_time:.2f}s): 원본='{text_to_translate[:30]}...', 결과='{translated_text[:30]}...'")
                    # 빈 결과도 그대로 반환 (호출한 쪽에서 처리하도록)
                return translated_text
            else: # 응답 형식 오류
                error_msg_format = f"번역 API 응답 형식 오류 (모델: {model_name}, 소요시간: {elapsed_time:.2f}s): {response_data}"
                logger.error(error_msg_format)
                return f"오류: API 응답 형식 이상 - \"{text_snippet_for_error}\"" # 형식 변경

        except requests.exceptions.Timeout:
            logger.error(f"Ollama API 요청 시간 초과 (모델: {model_name})")
            return f"오류: API 시간 초과 - \"{text_snippet_for_error}\"" # 형식 변경
        except requests.exceptions.RequestException as e_req:
            logger.error(f"Ollama API 요청 오류 (모델: {model_name}): {e_req}")
            return f"오류: API 요청 실패 ({e_req.__class__.__name__}) - \"{text_snippet_for_error}\"" # 형식 변경
        except Exception as e_generic:
            logger.error(f"번역 중 예외 발생 (모델: {model_name}): {e_generic}", exc_info=True)
            return f"오류: 번역 중 예외 ({e_generic.__class__.__name__}) - \"{text_snippet_for_error}\"" # 형식 변경
        # --- 1단계 개선 끝 ---

    def translate_texts_batch(self, texts_to_translate: List[str], src_lang_ui_name: str, tgt_lang_ui_name: str,
                              model_name: str, ollama_service_instance: 'OllamaService',
                              is_ocr_text: bool = False, ocr_temperature: Optional[float] = None,
                              stop_event: Optional[threading.Event] = None) -> List[str]:
        # ... (기존과 동일, 내부에서 translate_text 호출 시 변경된 오류 메시지 형식이 반영됨) ...
        if not texts_to_translate:
            return []

        translated_results = [""] * len(texts_to_translate) # 결과 저장용 리스트 초기화

        tasks_to_submit_with_indices = [] # 실제 API 호출이 필요한 작업들

        # 1. 캐시 확인 및 API 호출 대상 작업 분리
        for i, text in enumerate(texts_to_translate):
            if not text or not text.strip(): # 비거나 공백만 있는 텍스트
                translated_results[i] = text if text else ""
                continue
            if text.startswith("오류:"): # 이미 오류로 표시된 텍스트는 그대로 반환
                 translated_results[i] = text
                 continue

            cache_key = self._get_cache_key(text, src_lang_ui_name, tgt_lang_ui_name, model_name)
            if cache_key in self.translation_cache:
                translated_results[i] = self.translation_cache[cache_key]
                # logger.debug(f"배치 번역 캐시 사용 (키: {cache_key}): '{text[:30]}...'")
            else: # 캐시에 없으면 API 호출 대상
                tasks_to_submit_with_indices.append({'text': text, 'original_index': i})

        if not tasks_to_submit_with_indices: # 모든 텍스트가 캐시되었거나 비어있는 경우
            return translated_results

        # 2. 병렬 API 호출
        futures_map = {} # {Future 객체: 원래 인덱스} 매핑
        with ThreadPoolExecutor(max_workers=MAX_TRANSLATION_WORKERS) as executor:
            # 작업 제출
            for item_data in tasks_to_submit_with_indices:
                if stop_event and stop_event.is_set(): # 제출 중 중지 요청 감지
                    logger.info("배치 번역 제출 중 중단 요청 감지됨.")
                    # 아직 제출되지 않은 작업은 원본 텍스트로 채움
                    for remaining_item in tasks_to_submit_with_indices:
                        if translated_results[remaining_item['original_index']] == "": # 아직 처리 안 된 경우
                            translated_results[remaining_item['original_index']] = remaining_item['text'] # 원본으로
                    break # 작업 제출 중단
                
                future = executor.submit(self.translate_text, # 캐싱 로직이 포함된 translate_text 호출
                                         item_data['text'],
                                         src_lang_ui_name,
                                         tgt_lang_ui_name,
                                         model_name,
                                         ollama_service_instance,
                                         is_ocr_text,
                                         ocr_temperature)
                futures_map[future] = item_data['original_index']
            
            if stop_event and stop_event.is_set() and not futures_map: # 제출 전 중단된 경우
                 return translated_results # 이미 원본으로 채워졌거나, 빈 결과 반환


            # 결과 취합
            for future in as_completed(futures_map):
                original_idx = futures_map[future]
                text_snippet_for_log = texts_to_translate[original_idx][:20].replace('\n',' ') + "..."

                if stop_event and stop_event.is_set(): # 결과 취합 중 중지 요청 감지
                    if not future.done() or future.cancelled(): # 아직 완료 안됐거나 취소된 작업
                        translated_results[original_idx] = texts_to_translate[original_idx] # 원본으로
                        logger.debug(f"배치 번역 중단으로 인덱스 {original_idx} 텍스트 원본 유지: '{text_snippet_for_log}'")
                    else: # 이미 완료된 작업은 결과 사용
                        try:
                            translated_text = future.result()
                            translated_results[original_idx] = translated_text
                        except Exception as e:
                            logger.error(f"배치 번역 중단 시 완료된 작업 결과 가져오는 중 오류 (인덱스 {original_idx}): {e}")
                            translated_results[original_idx] = f"오류: 중단 시 결과 처리 실패 - \"{text_snippet_for_log}\""
                    continue # 다음 future 처리 (이미 완료된 것들 마저 처리)

                try:
                    translated_text = future.result() # translate_text 내부에서 성공 시 캐시에 저장됨
                    translated_results[original_idx] = translated_text
                except Exception as e: # Future.result()에서 발생할 수 있는 예외 (translate_text 내부 예외 포함)
                    logger.error(f"배치 번역 중 '{text_snippet_for_log}' 처리 오류: {e}")
                    translated_results[original_idx] = f"오류: 배치 처리 중 예외 - \"{text_snippet_for_log}\""
        
        # 3. 최종적으로 중단 요청 시 처리되지 않은 부분 확인 (루프 후)
        if stop_event and stop_event.is_set():
            for i in range(len(translated_results)):
                if translated_results[i] == "": # 어떤 이유로든 비어있다면 (예: 제출 전 중단, 또는 예외 발생 후 빈 문자열)
                    is_task_identified = False
                    for item_data in tasks_to_submit_with_indices: # 이 작업이 API 호출 대상이었는지 확인
                        if item_data['original_index'] == i:
                             translated_results[i] = texts_to_translate[i] # 원본으로 복구
                             is_task_identified = True
                             break
                    # API 호출 대상이 아니었거나 (즉, 캐시 사용 예정이었거나 빈 텍스트),
                    # 또는 호출 대상이었으나 위에서 원본으로 복구된 경우 외에,
                    # texts_to_translate[i]가 실제로 내용이 있고, 오류로 시작하지 않는다면 원본으로 채움.
                    # (이 경우는 거의 발생하지 않아야 함)
                    if not is_task_identified and texts_to_translate[i] and \
                       not texts_to_translate[i].strip().startswith("오류:"):
                         translated_results[i] = texts_to_translate[i]

        return translated_results

    def clear_translation_cache(self):
        logger.info(f"실행 중 번역 캐시({len(self.translation_cache)} 항목)가 비워졌습니다.")
        self.translation_cache.clear()