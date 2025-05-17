# translator.py
import logging
import time
from typing import TYPE_CHECKING, Optional
import requests # Ensure requests is imported

# 설정 파일 import
import config

if TYPE_CHECKING:
    from ollama_service import OllamaService

logger = logging.getLogger(__name__)

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
            return ""

        prompt = f"Translate the following text from {src_lang_ui_name} to {tgt_lang_ui_name}. Do not provide any explanations or introductory phrases, only the translated text itself. Text to translate:\n\n{text_to_translate}"
        
        current_temperature = config.TRANSLATOR_TEMPERATURE_GENERAL
        if is_ocr_text and ocr_temperature is not None:
            current_temperature = ocr_temperature
            logger.debug(f"OCR 텍스트 번역에 사용자 지정 온도 적용: {current_temperature}")
        else:
            logger.debug(f"일반 텍스트 번역에 기본 온도 적용: {current_temperature}")

        try:
            if not ollama_service_instance.is_running()[0]: # Check if Ollama is running via the service instance
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
            # Use timeouts from the ollama_service_instance
            response = requests.post(api_url, json=payload, 
                                     timeout=(ollama_service_instance.connect_timeout, ollama_service_instance.read_timeout))
            response.raise_for_status()
            response_data = response.json()
            end_time = time.time()
            
            if response_data and "response" in response_data:
                translated_text = response_data["response"].strip()
                logger.debug(f"번역 완료 (모델: {model_name}, 온도: {current_temperature}, 소요시간: {end_time - start_time:.2f}s): '{text_to_translate[:30]}...' -> '{translated_text[:30]}...'")
                return translated_text
            else:
                error_msg_format = f"번역 API 응답 형식 오류 (모델: {model_name}): {response_data}"
                logger.error(error_msg_format)
                return f"오류: 번역 API 응답 없음 - {text_to_translate[:20]}..."

        except requests.exceptions.RequestException as e_req:
            error_msg_req = f"Ollama API 요청 오류 (모델: {model_name}): {e_req}"
            logger.error(error_msg_req, exc_info=True)
            return f"오류: API 요청 실패 - {text_to_translate[:20]}..."
        except Exception as e_generic:
            error_msg_generic = f"번역 중 예외 발생 (모델: {model_name}): {e_generic}"
            logger.error(error_msg_generic, exc_info=True)
            return f"오류: 번역 중 예외 - {text_to_translate[:20]}..."