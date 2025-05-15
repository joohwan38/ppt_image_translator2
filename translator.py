import requests
import json
import logging
import re

logger = logging.getLogger(__name__) # 전역 로거 사용 (main.py에서 설정)

class OllamaTranslator:
    def _post_process_translation(self, text: str, target_lang_ui_name: str = None) -> str:
        if not text:
            logger.debug("후처리 입력: 비어 있음")
            return ""

        logger.debug(f"후처리 입력 (원본 repr): {repr(text)}") # repr()로 로깅

        processed_text = text.strip()
        logger.debug(f"후처리 (strip 후 repr): {repr(processed_text)}")

        # 시작 부분의 명시적인 줄바꿈 문자들 제거 (반복 처리)
        temp_text_for_debug = processed_text # 디버깅용 임시 변수
        while processed_text.startswith('\n') or processed_text.startswith('\r'):
            processed_text = processed_text.lstrip('\n\r \t')
        
        if temp_text_for_debug != processed_text: # 변경이 있었다면 로그
            logger.debug(f"후처리 (앞 줄바꿈 제거 후 repr): {repr(processed_text)}")

        common_introductions_conclusions = [
            r"Here is the translation(?: into \w+)?:", r"The translation is as follows:",
            r"Translation:", r"Translated text:",
            r"\w+ translation:", 
            r"以下是翻译.*?:", r"翻译如下:", r"翻译:", 
            r"以下為翻譯.*?:", 
            r"以下は翻訳です.*?:", r"翻訳は以下の通りです:", r"翻訳:", 
            r"다음은 번역입니다.*?:", r"번역 결과는 다음과 같습니다:", r"번역:",
            r"Note:",
            r"I am an AI language model.*", r"As an AI language model.*" ,
            r"I hope this helps.*", r"Please let me know if you have other questions.*",
            r"Sure, here'?s the translation:", r"Okay, here is the translation:",
            r"^```(?:\w+\n)?", r"```$", 
            r"^'''", r"'''$",
            r'^"""', r'"""$',
            r"^<translation>", r"</translation>$"
        ]
        processed_text = re.sub(r'\n\s*\n', '\n\n', processed_text)
        processed_text = re.sub(r'\n{3,}', '\n\n', processed_text).strip()
        logger.debug(f"후처리 (연속 줄바꿈 정리 후 repr): {repr(processed_text)}")
            
        if processed_text == '"' or processed_text == "''" or processed_text == '""' or processed_text == "'''" or processed_text == '"""':
            processed_text = ""
            logger.debug(f"후처리 (따옴표만 있는 경우 비움): {repr(processed_text)}")
            
        final_result = processed_text.strip() # 최종적으로 한 번 더 strip
        logger.debug(f"후처리 최종 결과 (repr): {repr(final_result)}")
        return final_result

    def translate_text(self, text_to_translate:str, src_lang_ui_name:str, tgt_lang_ui_name:str, model_name:str, ollama_service_instance, is_ocr_text:bool=False):
        """
        텍스트를 번역합니다.
        is_ocr_text: 입력 텍스트가 OCR로 추출된 텍스트인지 여부 (프롬프트 분기용)
        """
        is_running, _ = ollama_service_instance.is_running()
        if not is_running:
            logger.error("Ollama 서버 미실행. 번역 불가.")
            return f"오류: Ollama 서버 미실행"
        
        if not text_to_translate.strip(): # 번역할 텍스트가 비어있으면 바로 반환
            logger.warning("번역할 텍스트가 비어있습니다.")
            return ""

        try:
            lang_map = {
                "한국어": "Korean", "일본어": "Japanese", "대만어": "Traditional Chinese",
                "중국어": "Simplified Chinese", "태국어": "Thai", "영어": "English", "스페인어": "Spanish"
            }
            source_lang_name = lang_map.get(src_lang_ui_name, src_lang_ui_name)
            target_lang_name = lang_map.get(tgt_lang_ui_name, tgt_lang_ui_name)

            if is_ocr_text:
                # OCR 텍스트용 프롬프트 (Claude 제안 반영)
                prompt = (
                    f"Translate the following OCR-extracted text from {source_lang_name} to {target_lang_name}.\n"
                    f"This text was extracted from an image and may contain errors or unusual formatting.\n"
                    f"Your goal is to produce a natural and accurate translation in {target_lang_name}, preserving the original meaning.\n"
                    f"Provide ONLY the translated text. Do not add any extra words, explanations, or introductory phrases.\n"
                    f"If the input is nonsensical, return it as is without any comment.\n\n"
                    f"Original OCR text:\n{text_to_translate}\n\n"
                    f"Translated text:\n"
                )
                temperature_setting = 0.5 # OCR 텍스트는 약간 더 높은 temperature (Claude 제안)
            else:
                # 일반 텍스트용 프롬프트 (이전 답변의 수정된 버전)
                prompt = (
                    f"Translate the following text from {source_lang_name} to {target_lang_name}.\n"
                    f"Provide ONLY the translated text. Do not add any extra words, explanations, or introductory phrases.\n"
                    f"Maintain relevant formatting like line breaks if they are part of the meaning.\n"
                    f"If the input is nonsensical, return it as is without any comment.\n\n"
                    f"Original text:\n{text_to_translate}\n\n"
                    f"Translated text:\n"
                )
                temperature_setting = 0.1 # 일반 텍스트는 낮은 temperature

            logger.debug(f"번역 프롬프트 ({model_name}, OCR: {is_ocr_text}): {prompt[:350]}...")
            payload = {
                "model": model_name, "prompt": prompt, "stream": False,
                "options": {"temperature": temperature_setting, "num_ctx": 4096, "top_k": 30, "top_p": 0.85} # 파라미터 조정
            }
            api_generate_url = f"{ollama_service_instance.url}/api/generate"
            connect_timeout = ollama_service_instance.connect_timeout
            # 번역 API 타임아웃은 더 길게
            translate_read_timeout = ollama_service_instance.read_timeout * (6 if is_ocr_text else 4) # OCR은 더 길게
            if translate_read_timeout <= 0 : translate_read_timeout = 300 
            
            response = requests.post(api_generate_url, json=payload, timeout=(connect_timeout, translate_read_timeout))
            response.raise_for_status()
            response_data = response.json()
            raw_translated_text = response_data.get("response", "")
            
            log_context = f"(OCR, {src_lang_ui_name}->{tgt_lang_ui_name})" if is_ocr_text else f"({src_lang_ui_name}->{tgt_lang_ui_name})"
            logger.info(f"모델 원본 응답 {log_context}: '{text_to_translate[:30].replace(chr(10),' ')}' -> '{raw_translated_text[:50].replace(chr(10),' ')}'")
            
            final_translated_text = self._post_process_translation(raw_translated_text, target_lang_ui_name=tgt_lang_ui_name)
            
            logger.info(f"최종 번역 결과 {log_context}: '{text_to_translate[:20].replace(chr(10),' ')}' -> '{final_translated_text[:30].replace(chr(10),' ')}'")
            
            if not final_translated_text and raw_translated_text.strip() and len(text_to_translate.strip()) > 0:
                logger.warning("후처리 후 번역 결과가 비었으나 원본 응답은 내용이 있어, 원본 응답(strip)을 사용합니다.")
                return raw_translated_text.strip()
            return final_translated_text
        except requests.exceptions.Timeout:
            logger.error(f"Ollama 번역 API 시간 초과: {text_to_translate[:30]}...")
            return f"오류: 번역 API 시간 초과"
        except requests.exceptions.RequestException as e:
            logger.error(f"Ollama 번역 API 오류: {e}", exc_info=True)
            return f"오류: 번역 API 오류 ({e})"
        except json.JSONDecodeError as e:
            logger.error(f"Ollama 응답 JSON 디코딩 오류: {e}", exc_info=True)
            return f"오류: 번역 응답 처리 오류"
        except Exception as e:
            logger.error(f"번역 중 예기치 않은 오류: {e}", exc_info=True)
            return f"오류: 번역 중 내부 오류"