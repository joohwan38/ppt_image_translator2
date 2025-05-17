# translator.py
import logging
import time
from typing import Optional, List, Dict, Any
import requests
# ThreadPoolExecutor와 as_completed는 배치 API를 직접 사용하면 필요 없을 수 있습니다.
# from concurrent.futures import ThreadPoolExecutor, as_completed 
import threading
import hashlib

import config
from interfaces import AbsTranslator, AbsOllamaService

logger = logging.getLogger(__name__)

# MAX_TRANSLATION_WORKERS는 배치 API 사용 시 의미가 달라지거나 불필요해질 수 있습니다.
# MAX_TRANSLATION_WORKERS = config.MAX_TRANSLATION_WORKERS

class OllamaTranslator(AbsTranslator):
    def __init__(self):
        self.translation_cache: Dict[str, str] = {}
        # logger.info(f"OllamaTranslator 초기화됨. 번역 작업자 수: {MAX_TRANSLATION_WORKERS}") # 주석 처리 또는 수정
        logger.info("OllamaTranslator 초기화됨.")


    def _get_cache_key(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str) -> str:
        key_string = f"{src_lang_ui_name}|{tgt_lang_ui_name}|{model_name}|{text_to_translate}"
        return hashlib.md5(key_string.encode('utf-8')).hexdigest()

    def translate_text(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str,
                       model_name: str, ollama_service_instance: AbsOllamaService,
                       is_ocr_text: bool = False, ocr_temperature: Optional[float] = None) -> str:
        # 이 메서드는 배치 API 사용 시에는 직접 호출되지 않거나,
        # 매우 적은 수의 텍스트(예: 1개)에 대한 폴백(fallback)용으로 남겨둘 수 있습니다.
        # 또는 내부적으로 단일 텍스트를 배치 API 형식으로 감싸서 호출하도록 수정할 수도 있습니다.
        # 현재 코드는 그대로 두거나, 필요시 수정합니다.
        # ... (기존 translate_text 로직) ...
        if not text_to_translate or not text_to_translate.strip():
            return text_to_translate if text_to_translate else ""

        cache_key = self._get_cache_key(text_to_translate, src_lang_ui_name, tgt_lang_ui_name, model_name)
        if cache_key in self.translation_cache:
            cached_result = self.translation_cache[cache_key]
            return cached_result

        prompt = f"Translate the following text from {src_lang_ui_name} to {tgt_lang_ui_name}. Provide only the translated text itself, without any additional explanations, introductory phrases, or quotation marks around the translation. Text to translate:\n\n{text_to_translate}"
        current_temperature = config.TRANSLATOR_TEMPERATURE_GENERAL
        if is_ocr_text and ocr_temperature is not None:
            current_temperature = ocr_temperature
        
        text_snippet_for_error = text_to_translate[:20].replace('\n', ' ') + "..."
        try:
            if not ollama_service_instance or not ollama_service_instance.is_running()[0]:
                # ... (오류 처리)
                return f"오류: Ollama 서버 연결 실패 - \"{text_snippet_for_error}\""

            api_url = f"{ollama_service_instance.url}/api/generate" # 개별 API 엔드포인트
            payload = {
                "model": model_name, "prompt": prompt, "stream": False,
                "options": {"temperature": current_temperature}
            }
            response = requests.post(api_url, json=payload, timeout=(config.OLLAMA_CONNECT_TIMEOUT, config.OLLAMA_READ_TIMEOUT))
            response.raise_for_status()
            response_data = response.json()
            if response_data and "response" in response_data:
                translated_text = response_data["response"].strip()
                if translated_text and "오류:" not in translated_text:
                    self.translation_cache[cache_key] = translated_text
                return translated_text
            # ... (오류 처리)
            return f"오류: API 응답 형식 이상 - \"{text_snippet_for_error}\""
        # ... (예외 처리)
        except Exception as e_generic:
            return f"오류: 번역 중 예외 ({e_generic.__class__.__name__}) - \"{text_snippet_for_error}\""


    def translate_texts_batch(self, texts_to_translate: List[str], src_lang_ui_name: str, tgt_lang_ui_name: str,
                              model_name: str, ollama_service_instance: AbsOllamaService,
                              is_ocr_text: bool = False, ocr_temperature: Optional[float] = None,
                              stop_event: Optional[threading.Event] = None) -> List[str]:
        if not texts_to_translate:
            return []

        translated_results = [""] * len(texts_to_translate)
        
        # 1. 캐시 확인 및 실제 API 호출 대상 필터링
        tasks_for_batch_api_call = [] # (original_text, original_index) 저장
        texts_to_send_to_api = []     # 실제 API로 보낼 텍스트 목록 (프롬프트 포함 가능)
        
        for i, text in enumerate(texts_to_translate):
            if stop_event and stop_event.is_set(): # 작업 제출 전 중단 확인
                logger.info("배치 번역 작업 준비 중 중단 요청 감지.")
                # 남은 텍스트는 원본으로 채우거나 특정 오류 메시지 반환
                for j in range(i, len(texts_to_translate)):
                    if translated_results[j] == "": # 아직 처리 안 된 경우
                        translated_results[j] = texts_to_translate[j]
                return translated_results

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
                # 배치 API로 보낼 작업 정보 저장
                tasks_for_batch_api_call.append({'original_text': text, 'original_index': i})
                
                # Ollama 배치 API가 각 텍스트별 프롬프트를 어떻게 받을지에 따라 구성 변경
                # 예시: API가 프롬프트 자체를 받는다고 가정
                prompt_for_this_text = f"Translate the following text from {src_lang_ui_name} to {tgt_lang_ui_name}. Provide only the translated text itself, without any additional explanations, introductory phrases, or quotation marks around the translation. Text to translate:\n\n{text}"
                texts_to_send_to_api.append(prompt_for_this_text) # 또는 text 자체만 보낼 수도 있음

        if not tasks_for_batch_api_call: # 모든 텍스트가 캐시되었거나 비어있음
            return translated_results

        if stop_event and stop_event.is_set(): # 모든 캐시 확인 후 최종 중단 확인
            logger.info("배치 API 호출 직전 중단 요청 감지.")
            for task_info in tasks_for_batch_api_call:
                translated_results[task_info['original_index']] = task_info['original_text']
            return translated_results
            
        # 2. Ollama 배치 API 호출
        try:
            if not ollama_service_instance or not ollama_service_instance.is_running()[0]:
                error_msg_server = f"Ollama 서버 미실행. {model_name} 모델로 배치 번역 불가."
                logger.error(error_msg_server)
                for task_info in tasks_for_batch_api_call:
                    translated_results[task_info['original_index']] = f"오류: Ollama 서버 연결 실패 - \"{task_info['original_text'][:20]}...\""
                return translated_results

            # ===== Ollama 배치 API 호출 로직 (실제 API 사양에 따라 크게 달라짐) =====
            # 가정: `/api/generate_batch` 엔드포인트가 있고, 아래와 유사한 JSON 형식을 사용한다고 가정
            # 요청: {"model": "...", "prompts": ["prompt1", "prompt2", ...], "options": {...}, "stream": false}
            # 응답: {"results": [{"response": "translated1"}, {"response": "translated2"}, ...]} 또는 오류 정보 포함
            
            api_batch_url = f"{ollama_service_instance.url}/api/generate_batch" # 가상의 배치 엔드포인트
            
            current_batch_temperature = config.TRANSLATOR_TEMPERATURE_GENERAL
            if is_ocr_text and ocr_temperature is not None:
                current_batch_temperature = ocr_temperature

            # Ollama의 /api/generate는 한 번에 하나의 프롬프트만 처리합니다.
            # 만약 Ollama가 실제 배치 API (예: 여러 'prompt'를 리스트로 받는 기능)를 
            # `/api/generate` 엔드포인트에서 지원하거나 별도 엔드포인트를 제공한다면,
            # 해당 API 명세를 따라야 합니다.
            #
            # 현재로서는 Ollama에 그런 표준 배치 API가 없으므로,
            # 이 부분은 "만약 있다면"에 대한 개념적인 코드입니다.
            # 실제로는 이전의 ThreadPoolExecutor 방식이 현재 Ollama API 하에서는 더 현실적일 수 있습니다.
            #
            # 이 예시에서는 /api/generate가 prompts 리스트를 받는다고 "가정"해봅니다. (실제론 아님)
            # 또는, Ollama 서비스에 배치 처리를 위한 래퍼 함수를 만들 수도 있습니다.
            
            # << 경고: 아래는 Ollama의 현재 API와는 다를 수 있는 가상적인 배치 요청입니다. >>
            batch_payload = {
                "model": model_name,
                "prompts": texts_to_send_to_api, # 각 텍스트에 대해 생성된 프롬프트 목록
                "stream": False,
                "options": { # 공통 옵션. 각 프롬프트별 옵션이 가능하다면 페이로드 구조 변경
                    "temperature": current_batch_temperature
                }
            }
            
            # 만약 Ollama가 /api/generate 엔드포인트에서 "prompt" 필드에 리스트를 허용하지 않는다면,
            # 이 방식은 작동하지 않습니다. 그 경우, 이전처럼 ThreadPoolExecutor를 사용하거나,
            # ollama_service.py에 여러 요청을 내부적으로 처리하는 로직을 만들어야 합니다.

            logger.info(f"Ollama 배치 API 호출 시작 ({len(texts_to_send_to_api)}개 텍스트, 모델: {model_name})...")
            start_time = time.time()
            
            # 타임아웃은 (연결 타임아웃, 읽기 타임아웃)입니다.
            # 배치 작업이므로 읽기 타임아웃을 텍스트 수에 비례하여 늘리는 것을 고려할 수 있습니다.
            # 예: config.OLLAMA_READ_TIMEOUT * max(1, len(texts_to_send_to_api) // 5)
            # 하지만 너무 길게 잡으면 다른 문제가 생길 수 있으므로 적절한 값 설정이 중요합니다.
            # 여기서는 우선 고정된 타임아웃을 사용합니다.
            effective_read_timeout = config.OLLAMA_READ_TIMEOUT 
            if len(texts_to_send_to_api) > 10: # 예시: 10개 초과 시 타임아웃 증가
                 effective_read_timeout = config.OLLAMA_READ_TIMEOUT * (len(texts_to_send_to_api) / 10.0)
            effective_read_timeout = min(effective_read_timeout, 600) # 최대 10분으로 제한


            # 실제로는 Ollama가 `/api/generate`에서 `prompt` 필드에 단일 문자열만 받습니다.
            # 따라서, 이 "batch_payload"를 직접 보내는 것은 작동하지 않습니다.
            # 만약 Ollama가 그런 기능을 추가하거나, 사용자가 Ollama를 수정하여 사용한다면 가능합니다.
            # 현재 상태에서는 이 부분 대신 ThreadPoolExecutor 로직이 맞습니다.
            # 이 답변은 "가장 큰 성능 향상을 기대할 수 있는 부분"에 대한 이론적 접근입니다.

            # --- 만약 Ollama가 위와 같은 배치 API를 지원한다면 아래와 같이 진행 ---
            # response = requests.post(
            #     f"{ollama_service_instance.url}/api/generate", # 또는 실제 배치 엔드포인트
            #     json=batch_payload, # 이 페이로드 형식이 지원된다고 가정
            #     timeout=(config.OLLAMA_CONNECT_TIMEOUT, effective_read_timeout)
            # )
            # response.raise_for_status() # 오류 시 예외 발생
            # batch_response_data = response.json()
            # end_time = time.time()
            # logger.info(f"Ollama 배치 API 응답 받음 (소요 시간: {end_time - start_time:.2f}s).")

            # # 3. API 응답 파싱 및 결과 적용
            # # (batch_response_data의 실제 구조에 따라 파싱 로직이 달라집니다)
            # # 예시: 응답이 {"results": [{"response": "번역1"}, {"response": "번역2"}, ...]} 형태라고 가정
            # if batch_response_data and "results" in batch_response_data and isinstance(batch_response_data["results"], list):
            #     api_translated_texts = batch_response_data["results"]
            #     if len(api_translated_texts) == len(tasks_for_batch_api_call):
            #         for i, res_item in enumerate(api_translated_texts):
            #             original_idx = tasks_for_batch_api_call[i]['original_index']
            #             original_text_for_cache = tasks_for_batch_api_call[i]['original_text']
            #             text_snippet_log = original_text_for_cache[:20].replace('\n',' ') + "..."
                            
            #             if isinstance(res_item, dict) and "response" in res_item and res_item.get("response"):
            #                 translated_text = res_item["response"].strip()
            #                 translated_results[original_idx] = translated_text
            #                 if "오류:" not in translated_text: # 캐시 저장
            #                     self.translation_cache[self._get_cache_key(original_text_for_cache, src_lang_ui_name, tgt_lang_ui_name, model_name)] = translated_text
            #                 logger.debug(f"배치 결과 적용 (인덱스 {original_idx}): '{text_snippet_log}' -> '{translated_text[:30]}...'")
            #             elif isinstance(res_item, dict) and "error" in res_item: # 개별 항목 오류 처리
            #                 err_msg = f"오류: 배치 API 개별 항목 오류 ({res_item['error']}) - \"{text_snippet_log}\""
            #                 translated_results[original_idx] = err_msg
            #                 logger.warning(err_msg)
            #             else: # 알 수 없는 응답 형식
            #                 err_msg_unknown = f"오류: 배치 API 알 수 없는 형식의 응답 - \"{text_snippet_log}\""
            #                 translated_results[original_idx] = err_msg_unknown
            #                 logger.warning(f"{err_msg_unknown} 응답: {res_item}")
            #     else:
            #         logger.error("Ollama 배치 API 응답 수와 요청 수가 불일치합니다.")
            #         # 실패 시 원본으로 채우거나 오류 메시지 설정
            #         for task_info in tasks_for_batch_api_call:
            #             translated_results[task_info['original_index']] = f"오류: 배치 응답 불일치 - \"{task_info['original_text'][:20]}...\""
            # else:
            #     logger.error(f"Ollama 배치 API 응답 형식이 올바르지 않습니다: {batch_response_data}")
            #     for task_info in tasks_for_batch_api_call:
            #         translated_results[task_info['original_index']] = f"오류: 배치 응답 형식 이상 - \"{task_info['original_text'][:20]}...\""
            # --- 가상 배치 API 처리 로직 종료 ---


            # <<중요>>: Ollama의 현재 표준 API(/api/generate)는 한 번에 하나의 프롬프트만 처리합니다.
            # 따라서, 실제 "배치" API가 없다면, 이전 답변에서 설명드렸던 것처럼
            # ThreadPoolExecutor를 사용하여 개별 API 호출을 병렬로 수행하는 방식이 여전히 유효합니다.
            # 아래는 ThreadPoolExecutor를 사용하는 기존 방식의 코드입니다.
            # 만약 진정한 배치 API가 없다면 이 코드가 사용되어야 합니다.
            
            futures_map = {}
            # ThreadPoolExecutor는 import 되어 있어야 합니다.
            from concurrent.futures import ThreadPoolExecutor, as_completed

            with ThreadPoolExecutor(max_workers=config.MAX_TRANSLATION_WORKERS) as executor:
                for task_info in tasks_for_batch_api_call: # tasks_for_batch_api_call 사용
                    if stop_event and stop_event.is_set():
                        logger.info("배치 번역(ThreadPool) 제출 중 중단 요청 감지됨.")
                        # 현재 task_info 이후의 것들은 원본으로 채움
                        current_task_index = tasks_for_batch_api_call.index(task_info)
                        for i in range(current_task_index, len(tasks_for_batch_api_call)):
                            unprocessed_task = tasks_for_batch_api_call[i]
                            if translated_results[unprocessed_task['original_index']] == "":
                                translated_results[unprocessed_task['original_index']] = unprocessed_task['original_text']
                        break 

                    # self.translate_text를 사용하여 개별 번역 작업 제출
                    future = executor.submit(self.translate_text, # self.translate_text 사용
                                             task_info['original_text'],
                                             src_lang_ui_name,
                                             tgt_lang_ui_name,
                                             model_name,
                                             ollama_service_instance,
                                             is_ocr_text,
                                             ocr_temperature)
                    futures_map[future] = task_info['original_index']

                if stop_event and stop_event.is_set() and not futures_map:
                     return translated_results # 제출된 작업이 없으면 바로 반환

                for future in as_completed(futures_map):
                    original_idx = futures_map[future]
                    # 원본 텍스트 가져오기 (로깅용)
                    original_text_for_log = ""
                    for task in tasks_for_batch_api_call: # tasks_for_batch_api_call에서 검색
                        if task['original_index'] == original_idx:
                            original_text_for_log = task['original_text']
                            break
                    text_snippet_for_log = original_text_for_log[:20].replace('\n',' ') + "..."


                    if stop_event and stop_event.is_set():
                        if not future.done() or future.cancelled():
                            translated_results[original_idx] = original_text_for_log # 원본으로 복원
                            logger.debug(f"배치 번역(ThreadPool) 중단. 인덱스 {original_idx} 원본 유지: '{text_snippet_for_log}'")
                        else: 
                            try:
                                translated_results[original_idx] = future.result()
                            except Exception as e:
                                translated_results[original_idx] = f"오류: 중단 시 결과 처리 실패 - \"{text_snippet_for_log}\""
                        continue
                    
                    try:
                        translated_text_result = future.result()
                        translated_results[original_idx] = translated_text_result
                        # self.translate_text 내부에서 캐싱이 이미 처리됩니다.
                    except Exception as e:
                        logger.error(f"배치 번역(ThreadPool) 중 '{text_snippet_for_log}' 처리 오류: {e}")
                        translated_results[original_idx] = f"오류: 배치 처리 중 예외 - \"{text_snippet_for_log}\""
            
            # 스레드 풀 종료 후, 만약 stop_event가 설정되었고 아직 빈 결과가 있다면 원본으로 채움
            if stop_event and stop_event.is_set():
                for i in range(len(translated_results)):
                    if translated_results[i] == "":
                        # 이 인덱스가 API 호출 대상이었는지 확인
                        is_api_target = any(task['original_index'] == i for task in tasks_for_batch_api_call)
                        if is_api_target:
                             # tasks_for_batch_api_call에서 원본 텍스트를 찾아 채움
                             original_text_to_fill = ""
                             for task in tasks_for_batch_api_call:
                                 if task['original_index'] == i:
                                     original_text_to_fill = task['original_text']
                                     break
                             translated_results[i] = original_text_to_fill


        except requests.exceptions.Timeout: # 배치 API 자체에 대한 타임아웃 (위 가상 로직에서 발생 가능)
            logger.error(f"Ollama 배치 API 요청 시간 초과 (모델: {model_name})")
            for task_info in tasks_for_batch_api_call:
                translated_results[task_info['original_index']] = f"오류: 배치 API 시간 초과 - \"{task_info['original_text'][:20]}...\""
        except requests.exceptions.RequestException as e_req: # 배치 API 자체에 대한 요청 오류
            logger.error(f"Ollama 배치 API 요청 오류 (모델: {model_name}): {e_req}")
            for task_info in tasks_for_batch_api_call:
                translated_results[task_info['original_index']] = f"오류: 배치 API 요청 실패 ({e_req.__class__.__name__}) - \"{task_info['original_text'][:20]}...\""
        except Exception as e_generic_outer: # ThreadPoolExecutor 외부의 일반 예외
            logger.error(f"배치 번역 준비/실행 중 주요 예외 발생 (모델: {model_name}): {e_generic_outer}", exc_info=True)
            for task_info in tasks_for_batch_api_call:
                if translated_results[task_info['original_index']] == "": # 아직 결과가 없는 경우에만 오류 메시지 설정
                    translated_results[task_info['original_index']] = f"오류: 배치 처리 중 주요 예외 ({e_generic_outer.__class__.__name__}) - \"{task_info['original_text'][:20]}...\""
        
        return translated_results

    def clear_translation_cache(self):
        logger.info(f"실행 중 번역 캐시({len(self.translation_cache)} 항목)가 비워졌습니다.")
        self.translation_cache.clear()