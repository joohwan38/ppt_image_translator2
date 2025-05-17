# translator.py
import logging
import time
from typing import Optional, List, Dict, Any # TYPE_CHECKING 제거, Any 추가
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import hashlib

# 설정 파일 import
import config
from interfaces import AbsTranslator, AbsOllamaService # 인터페이스 import

logger = logging.getLogger(__name__)

MAX_TRANSLATION_WORKERS = config.MAX_TRANSLATION_WORKERS

class OllamaTranslator(AbsTranslator): # AbsTranslator 상속
    def __init__(self):
        self.translation_cache: Dict[str, str] = {}
        logger.info(f"OllamaTranslator 초기화됨. 번역 작업자 수: {MAX_TRANSLATION_WORKERS}")

    def _get_cache_key(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str) -> str:
        key_string = f"{src_lang_ui_name}|{tgt_lang_ui_name}|{model_name}|{text_to_translate}"
        return hashlib.md5(key_string.encode('utf-8')).hexdigest()

    def translate_text(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str,
                       model_name: str, ollama_service_instance: AbsOllamaService, # 타입 힌트 변경
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

            api_url = f"{ollama_service_instance.url}/api/generate" # AbsOllamaService.url 사용
            payload = {
                "model": model_name,
                "prompt": prompt,
                "stream": False,
                "options": {
                    "temperature": current_temperature
                }
            }

            start_time = time.time()
            # ollama_service_instance의 timeout 속성 직접 접근 대신 config 값 사용 (인터페이스에 timeout 정의 안함)
            response = requests.post(api_url, json=payload,
                                     timeout=(config.OLLAMA_CONNECT_TIMEOUT, config.OLLAMA_READ_TIMEOUT))
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
                              model_name: str, ollama_service_instance: AbsOllamaService, # 타입 힌트 변경
                              is_ocr_text: bool = False, ocr_temperature: Optional[float] = None,
                              stop_event: Optional[threading.Event] = None) -> List[str]: # Any 대신 threading.Event 명시
        if not texts_to_translate:
            return []

        translated_results = [""] * len(texts_to_translate) # 결과 저장용 리스트 초기화
        tasks_to_submit_with_indices = [] # 실제 API 호출이 필요한 작업들

        for i, text in enumerate(texts_to_translate):
            if not text or not text.strip():
                translated_results[i] = text if text else ""
                continue
            if text.startswith("오류:"):
                 translated_results[i] = text
                 continue

            cache_key = self._get_cache_key(text, src_lang_ui_name, tgt_lang_ui_name, model_name)
            if cache_key in self.translation_cache:
                translated_results[i] = self.translation_cache[cache_key]
            else:
                tasks_to_submit_with_indices.append({'text': text, 'original_index': i})

        if not tasks_to_submit_with_indices:
            return translated_results

        futures_map = {}
        with ThreadPoolExecutor(max_workers=MAX_TRANSLATION_WORKERS) as executor:
            for item_data in tasks_to_submit_with_indices:
                if stop_event and stop_event.is_set():
                    logger.info("배치 번역 제출 중 중단 요청 감지됨.")
                    for remaining_item in tasks_to_submit_with_indices[tasks_to_submit_with_indices.index(item_data):]: # 현재 이후 모든 작업
                        if translated_results[remaining_item['original_index']] == "":
                            translated_results[remaining_item['original_index']] = remaining_item['text']
                    break
                
                future = executor.submit(self.translate_text,
                                         item_data['text'],
                                         src_lang_ui_name,
                                         tgt_lang_ui_name,
                                         model_name,
                                         ollama_service_instance,
                                         is_ocr_text,
                                         ocr_temperature)
                futures_map[future] = item_data['original_index']
            
            if stop_event and stop_event.is_set() and not futures_map : # 제출 전에 중단된 경우
                 # 모든 작업 원본으로 채우기 (이미 위에서 처리됨)
                 return translated_results

            for future in as_completed(futures_map):
                original_idx = futures_map[future]
                text_snippet_for_log = texts_to_translate[original_idx][:20].replace('\n',' ') + "..."

                if stop_event and stop_event.is_set():
                    if not future.done() or future.cancelled():
                        translated_results[original_idx] = texts_to_translate[original_idx]
                        logger.debug(f"배치 번역 중단으로 인덱스 {original_idx} 텍스트 원본 유지: '{text_snippet_for_log}'")
                    else:
                        try:
                            translated_text = future.result()
                            translated_results[original_idx] = translated_text
                        except Exception as e:
                            logger.error(f"배치 번역 중단 시 완료된 작업 결과 가져오는 중 오류 (인덱스 {original_idx}): {e}")
                            translated_results[original_idx] = f"오류: 중단 시 결과 처리 실패 - \"{text_snippet_for_log}\""
                    continue

                try:
                    translated_text = future.result()
                    translated_results[original_idx] = translated_text
                except Exception as e:
                    logger.error(f"배치 번역 중 '{text_snippet_for_log}' 처리 오류: {e}")
                    translated_results[original_idx] = f"오류: 배치 처리 중 예외 - \"{text_snippet_for_log}\""
        
        if stop_event and stop_event.is_set():
            for i in range(len(translated_results)):
                if translated_results[i] == "": 
                    # tasks_to_submit_with_indices에 있었던 것만 원본으로 복구
                    is_api_target = any(task['original_index'] == i for task in tasks_to_submit_with_indices)
                    if is_api_target:
                        translated_results[i] = texts_to_translate[i]
                    # 캐시 대상이었거나, 원래 비어있던 텍스트는 그대로 유지됨 (공백 또는 빈 문자열)
                    elif not texts_to_translate[i] or not texts_to_translate[i].strip():
                         translated_results[i] = texts_to_translate[i]
                    # 드물지만, API 대상이 아니었고, 내용이 있었는데 결과가 없는 경우 (이론상 발생 안 함)
                    elif texts_to_translate[i] and not texts_to_translate[i].startswith("오류:"):
                        translated_results[i] = texts_to_translate[i]


        return translated_results

    def clear_translation_cache(self):
        logger.info(f"실행 중 번역 캐시({len(self.translation_cache)} 항목)가 비워졌습니다.")
        self.translation_cache.clear()