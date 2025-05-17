# translator.py
import logging
import time
from typing import TYPE_CHECKING, Optional, List
import requests # Ensure requests is imported
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# 설정 파일 import
import config

if TYPE_CHECKING:
    from ollama_service import OllamaService

logger = logging.getLogger(__name__)

# 동시에 Ollama에 요청할 최대 작업자 수
# 너무 높으면 로컬 Ollama 서버에 부담을 줄 수 있습니다.
# config 파일에 추가하거나 여기서 조정 가능합니다.
MAX_TRANSLATION_WORKERS = config.MAX_TRANSLATION_WORKERS if hasattr(config, 'MAX_TRANSLATION_WORKERS') else 5


class OllamaTranslator:
    def translate_text(self, text_to_translate: str, src_lang_ui_name: str, tgt_lang_ui_name: str,
                       model_name: str, ollama_service_instance: 'OllamaService',
                       is_ocr_text: bool = False, ocr_temperature: Optional[float] = None) -> str:
        """
        텍스트를 번역합니다.
        is_ocr_text가 True이고 ocr_temperature가 제공되면 해당 온도를 사용합니다.
        그렇지 않으면 config의 TRANSLATOR_TEMPERATURE_GENERAL을 사용합니다.
        OllamaService에 API 요청 메서드가 없다는 가정 하에 직접 요청합니다.
        """
        if not text_to_translate or not text_to_translate.strip():
            # 빈 텍스트나 공백만 있는 텍스트는 원본 또는 빈 문자열 반환
            return text_to_translate if text_to_translate else ""


        prompt = f"Translate the following text from {src_lang_ui_name} to {tgt_lang_ui_name}. Provide only the translated text itself, without any additional explanations, introductory phrases, or quotation marks around the translation. Text to translate:\n\n{text_to_translate}"
        
        current_temperature = config.TRANSLATOR_TEMPERATURE_GENERAL
        if is_ocr_text and ocr_temperature is not None:
            current_temperature = ocr_temperature
            # logger.debug(f"OCR 텍스트 번역에 사용자 지정 온도 적용: {current_temperature}") # 로그 레벨 조정 필요시 주석 해제
        # else:
            # logger.debug(f"일반 텍스트 번역에 기본 온도 적용: {current_temperature}") # 로그 레벨 조정 필요시 주석 해제

        try:
            # is_running 호출 전에 ollama_service_instance가 None이 아닌지 확인
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
            response.raise_for_status()
            response_data = response.json()
            end_time = time.time()
            
            if response_data and "response" in response_data:
                translated_text = response_data["response"].strip()
                # logger.debug(f"번역 완료 (모델: {model_name}, 온도: {current_temperature}, 소요시간: {end_time - start_time:.2f}s): '{text_to_translate[:30]}...' -> '{translated_text[:30]}...'")
                return translated_text
            else:
                error_msg_format = f"번역 API 응답 형식 오류 (모델: {model_name}): {response_data}"
                logger.error(error_msg_format)
                return f"오류: 번역 API 응답 없음 - {text_to_translate[:20]}..."

        except requests.exceptions.RequestException as e_req:
            error_msg_req = f"Ollama API 요청 오류 (모델: {model_name}): {e_req}"
            logger.error(error_msg_req, exc_info=False) # 너무 많은 로그를 피하기 위해 exc_info=False
            return f"오류: API 요청 실패 - {text_to_translate[:20]}..."
        except Exception as e_generic:
            error_msg_generic = f"번역 중 예외 발생 (모델: {model_name}): {e_generic}"
            logger.error(error_msg_generic, exc_info=False) # 너무 많은 로그를 피하기 위해 exc_info=False
            return f"오류: 번역 중 예외 - {text_to_translate[:20]}..."

    def translate_texts_batch(self, texts_to_translate: List[str], src_lang_ui_name: str, tgt_lang_ui_name: str,
                              model_name: str, ollama_service_instance: 'OllamaService',
                              is_ocr_text: bool = False, ocr_temperature: Optional[float] = None,
                              stop_event: Optional[threading.Event] = None) -> List[str]:
        """
        여러 텍스트를 병렬로 번역합니다.
        결과는 입력된 텍스트 리스트의 순서와 동일하게 반환됩니다.
        """
        if not texts_to_translate:
            return []

        translated_results = ["" for _ in texts_to_translate] # 결과를 저장할 리스트 (순서 보장)
        
        # (개선) 이미 번역된 텍스트(오류 메시지 포함)나 빈 텍스트는 미리 처리
        futures_map = {} # executor.submit 결과를 원래 인덱스와 매핑하기 위함
        actual_texts_to_process_with_indices = []

        for i, text in enumerate(texts_to_translate):
            if not text or not text.strip() or text.startswith("오류:"):
                translated_results[i] = text if text else ""
            else:
                actual_texts_to_process_with_indices.append({'text': text, 'original_index': i})

        if not actual_texts_to_process_with_indices: # 실제 번역할 텍스트가 없으면 바로 반환
            return translated_results

        with ThreadPoolExecutor(max_workers=MAX_TRANSLATION_WORKERS) as executor:
            # Future 객체와 원래 인덱스를 매핑
            for item_data in actual_texts_to_process_with_indices:
                if stop_event and stop_event.is_set():
                    logger.info("배치 번역 중 중단 요청 감지됨.")
                    # 이미 제출된 작업은 계속 진행될 수 있으나, 새 작업은 제출하지 않음
                    # 결과 리스트에서 아직 처리되지 않은 부분은 원본 텍스트나 오류 메시지로 채워야 할 수 있음
                    # 여기서는 단순화를 위해 이미 제출된 작업의 결과만 기다리고, 나머지는 기본값 유지
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

            for future in as_completed(futures_map):
                if stop_event and stop_event.is_set():
                     # future.cancel()은 이미 실행 중인 작업은 취소 못할 수 있음
                     # 필요시 더 강력한 중단 로직 고려
                    pass

                original_idx = futures_map[future]
                try:
                    translated_text = future.result()
                    translated_results[original_idx] = translated_text
                except Exception as e:
                    logger.error(f"배치 번역 중 '{texts_to_translate[original_idx][:20]}...' 처리 오류: {e}")
                    translated_results[original_idx] = f"오류: 배치 처리 중 예외 - {texts_to_translate[original_idx][:20]}..."
        
        # 중단 요청 시 처리되지 않은 항목에 대한 후처리
        if stop_event and stop_event.is_set():
            for i, res in enumerate(translated_results):
                if res == "" and any(item['original_index'] == i for item in actual_texts_to_process_with_indices if futures_map.get(next((f for f, idx in futures_map.items() if idx ==i), None)) and not futures_map[next((f for f, idx in futures_map.items() if idx ==i), None)].done() ): # 아직 결과가 없고, 처리 대상이었던 경우
                    translated_results[i] = texts_to_translate[i] # 원본으로 복원 또는 "중단됨" 메시지
                    logger.debug(f"배치 번역 중단으로 인덱스 {i}의 텍스트 원본 유지: '{texts_to_translate[i][:20]}...'")


        return translated_results