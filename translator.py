# translator.py
import logging
import time
from typing import TYPE_CHECKING, Optional, List, Dict # Dict 추가
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import hashlib # 캐시 키 생성에 사용 가능 (선택적)

# 설정 파일 import
import config

if TYPE_CHECKING:
    from ollama_service import OllamaService

logger = logging.getLogger(__name__)

MAX_TRANSLATION_WORKERS = config.MAX_TRANSLATION_WORKERS # config.py에서 가져옴

class OllamaTranslator:
    def __init__(self):
        # 실행 중 번역 캐시: {(src_lang, tgt_lang, model, text_hash): translated_text}
        self.translation_cache: Dict[str, str] = {}
        logger.info(f"OllamaTranslator 초기화됨. 번역 작업자 수: {MAX_TRANSLATION_WORKERS}")

    def _get_cache_key(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str) -> str:
        """번역 캐시를 위한 고유 키 생성"""
        # 텍스트가 매우 길 경우 해시 사용 고려, 짧으면 직접 사용
        # 여기서는 간단하게 조합 후 해시 사용 (매우 긴 텍스트 키 방지)
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
            logger.debug(f"번역 캐시 사용 (키: {cache_key}): '{text_to_translate[:30]}...' -> '{cached_result[:30]}...'")
            return cached_result

        prompt = f"Translate the following text from {src_lang_ui_name} to {tgt_lang_ui_name}. Provide only the translated text itself, without any additional explanations, introductory phrases, or quotation marks around the translation. Text to translate:\n\n{text_to_translate}"
        
        current_temperature = config.TRANSLATOR_TEMPERATURE_GENERAL
        if is_ocr_text and ocr_temperature is not None:
            current_temperature = ocr_temperature

        try:
            if not ollama_service_instance or not ollama_service_instance.is_running()[0]:
                error_msg_server = f"Ollama 서버 미실행. {model_name} 모델로 번역 불가."
                logger.error(error_msg_server)
                return f"오류: Ollama 서버 미실행 - {text_to_translate[:20]}..."

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
            response.raise_for_status() # HTTP 오류 발생 시 예외 발생
            response_data = response.json()
            end_time = time.time()
            
            elapsed_time = end_time - start_time
            if response_data and "response" in response_data:
                translated_text = response_data["response"].strip()
                # 성공적인 번역 결과만 캐시 (오류 메시지 등은 캐시하지 않음)
                if translated_text and "오류:" not in translated_text :
                    self.translation_cache[cache_key] = translated_text
                    logger.debug(f"번역 완료 및 캐시 저장 (키: {cache_key}, 모델: {model_name}, 소요시간: {elapsed_time:.2f}s): '{text_to_translate[:30]}...' -> '{translated_text[:30]}...'")
                else:
                    logger.debug(f"번역 결과에 오류 포함 또는 빈 결과로 캐시하지 않음 (모델: {model_name}, 소요시간: {elapsed_time:.2f}s): '{text_to_translate[:30]}...' -> '{translated_text[:30]}...'")
                return translated_text
            else:
                error_msg_format = f"번역 API 응답 형식 오류 (모델: {model_name}, 소요시간: {elapsed_time:.2f}s): {response_data}"
                logger.error(error_msg_format)
                return f"오류: 번역 API 응답 없음 - {text_to_translate[:20]}..."

        except requests.exceptions.Timeout as e_timeout:
            error_msg_timeout = f"Ollama API 요청 시간 초과 (모델: {model_name}): {e_timeout}"
            logger.error(error_msg_timeout)
            return f"오류: API 시간 초과 - {text_to_translate[:20]}..."
        except requests.exceptions.RequestException as e_req:
            error_msg_req = f"Ollama API 요청 오류 (모델: {model_name}): {e_req}"
            logger.error(error_msg_req)
            return f"오류: API 요청 실패 - {text_to_translate[:20]}..."
        except Exception as e_generic:
            error_msg_generic = f"번역 중 예외 발생 (모델: {model_name}): {e_generic}"
            logger.error(error_msg_generic, exc_info=True) # 더 자세한 오류 로깅
            return f"오류: 번역 중 예외 - {text_to_translate[:20]}..."

    def translate_texts_batch(self, texts_to_translate: List[str], src_lang_ui_name: str, tgt_lang_ui_name: str,
                              model_name: str, ollama_service_instance: 'OllamaService',
                              is_ocr_text: bool = False, ocr_temperature: Optional[float] = None,
                              stop_event: Optional[threading.Event] = None) -> List[str]:
        if not texts_to_translate:
            return []

        translated_results = [""] * len(texts_to_translate)
        
        # 제출할 작업과 원래 인덱스 매핑
        tasks_to_submit_with_indices = []
        
        for i, text in enumerate(texts_to_translate):
            if not text or not text.strip():
                translated_results[i] = text if text else ""
                continue
            if text.startswith("오류:"): # 이미 오류인 경우 그대로 반환
                 translated_results[i] = text
                 continue

            cache_key = self._get_cache_key(text, src_lang_ui_name, tgt_lang_ui_name, model_name)
            if cache_key in self.translation_cache:
                translated_results[i] = self.translation_cache[cache_key]
                logger.debug(f"배치 번역 캐시 사용 (키: {cache_key}): '{text[:30]}...'")
            else:
                tasks_to_submit_with_indices.append({'text': text, 'original_index': i})

        if not tasks_to_submit_with_indices: # 모든 텍스트가 캐시되었거나 비어있는 경우
            return translated_results

        futures_map = {}
        # MAX_TRANSLATION_WORKERS는 config.py에서 가져온 값을 사용
        with ThreadPoolExecutor(max_workers=MAX_TRANSLATION_WORKERS) as executor:
            for item_data in tasks_to_submit_with_indices:
                if stop_event and stop_event.is_set():
                    logger.info("배치 번역 제출 중 중단 요청 감지됨.")
                    # 중단 시 아직 제출되지 않은 작업은 원본 텍스트로 채움
                    for remaining_item in tasks_to_submit_with_indices:
                        if translated_results[remaining_item['original_index']] == "": # 아직 처리 안 된 경우
                            translated_results[remaining_item['original_index']] = remaining_item['text'] # 원본으로
                    break 
                
                future = executor.submit(self.translate_text, # 캐싱 로직이 포함된 translate_text 호출
                                         item_data['text'],
                                         src_lang_ui_name,
                                         tgt_lang_ui_name,
                                         model_name,
                                         ollama_service_instance,
                                         is_ocr_text,
                                         ocr_temperature)
                futures_map[future] = item_data['original_index']
            
            if stop_event and stop_event.is_set() and not futures_map: # 제출 전에 중단된 경우
                 return translated_results


            for future in as_completed(futures_map):
                original_idx = futures_map[future]
                if stop_event and stop_event.is_set():
                    # 이미 완료된 작업은 결과를 사용하고, 진행 중이거나 대기 중인 작업은 원본으로 처리
                    if not future.done() or future.cancelled():
                        translated_results[original_idx] = texts_to_translate[original_idx]
                        logger.debug(f"배치 번역 중단으로 인덱스 {original_idx} 텍스트 원본 유지: '{texts_to_translate[original_idx][:20]}...'")
                    else: # 이미 완료된 작업
                        try:
                            translated_text = future.result()
                            translated_results[original_idx] = translated_text
                        except Exception as e:
                            logger.error(f"배치 번역 완료된 작업 결과 가져오는 중 오류 (인덱스 {original_idx}): {e}")
                            translated_results[original_idx] = f"오류: 결과 처리 실패 - {texts_to_translate[original_idx][:20]}..."
                    continue # 다음 future 처리 (이미 완료된 것들 마저 처리)

                try:
                    translated_text = future.result() # translate_text 내부에서 성공 시 캐시에 저장됨
                    translated_results[original_idx] = translated_text
                except Exception as e:
                    logger.error(f"배치 번역 중 '{texts_to_translate[original_idx][:20]}...' 처리 오류: {e}")
                    translated_results[original_idx] = f"오류: 배치 처리 중 예외 - {texts_to_translate[original_idx][:20]}..."
        
        # 최종적으로 중단 요청 시 처리되지 않은 부분 확인 (루프 후)
        if stop_event and stop_event.is_set():
            for i in range(len(translated_results)):
                if translated_results[i] == "": # 어떤 이유로든 비어있다면 (예: 제출 전 중단)
                    is_processed = False
                    for item_data in tasks_to_submit_with_indices:
                        if item_data['original_index'] == i: # 처리 대상이었으나 결과가 없는 경우
                             translated_results[i] = texts_to_translate[i]
                             is_processed = True
                             break
                    if not is_processed and texts_to_translate[i] and not texts_to_translate[i].strip().startswith("오류:"):
                        # 캐시에도 없었고, 처리 대상 목록에도 없었지만(이미 캐시된 것으로 간주되었으나 결과가 비어있는 이상한 경우)
                        # 또는 처리 대상이었으나 결과가 비어있는 경우. 안전하게 원본으로.
                         translated_results[i] = texts_to_translate[i]


        return translated_results

    def clear_translation_cache(self):
        """
        실행 중인 번역 캐시를 비웁니다.
        새 문서 번역 시작 시 호출하여 이전 문서의 캐시가 영향을 주지 않도록 할 수 있습니다.
        """
        logger.info(f"실행 중 번역 캐시({len(self.translation_cache)} 항목)가 비워졌습니다.")
        self.translation_cache.clear()