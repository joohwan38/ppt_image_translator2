import requests
import json
import logging
import re

# 설정 파일 import
import config

logger = logging.getLogger(__name__)

class OllamaTranslator:
    def _post_process_translation(self, text: str, target_lang_ui_name: str = None) -> str:
        if not text: return ""
        processed_text = text
        # 패턴에서 target_lang_ui_name을 사용할 때 None 체크 추가
        target_lang_regex_part = target_lang_ui_name if target_lang_ui_name else r"\w+"
        
        common_patterns = [
            r"^\s*Here is the translation(?: into " + target_lang_regex_part + r")?:?\s*",
            r"^\s*The translation is as follows:?\s*",
            r"^\s*Translation:?\s*", r"^\s*Translated text:?\s*",
            r"^\s*Sure, here'?s the translation:?\s*", r"^\s*Okay, here is the translation:?\s*",
            r"^\s*(?:In " + target_lang_regex_part + r"|The " + target_lang_regex_part + r" version is|As requested|Here you go)(?:, the translation is)?:?\s*",
            r"^\s*以下是翻译.*?[:：]?\s*", r"^\s*翻译如下[:：]?\s*", r"^\s*翻译[:：]?\s*",
            r"^\s*以下為翻譯.*?[:：]?\s*", r"^\s*以下は翻訳です.*?[:：]?\s*", r"^\s*翻訳は以下の通りです[:：]?\s*", r"^\s*翻訳[:：]?\s*",
            r"^\s*다음은 번역입니다.*?[:：]?\s*", r"^\s*번역 결과는 다음과 같습니다[:：]?\s*", r"^\s*번역[:：]?\s*",
            r"^\s*```(?:\w+\n)?", r"```\s*$", # 코드 블록 마커
            r"^\s*'''", r"'''\s*$", r'^\s*"""', r'"""\s*$', # 여러 줄 문자열 마커
            r"^\s*<translation>", r"</translation>\s*$", # XML 태그 유사 마커
            r"\s*I am an AI language model.*", r"\s*As an AI language model.*", # AI 자기소개 문구
            r"\s*I hope this helps.*", r"\s*Please let me know if you have other questions.*", # 마무리 문구
            r"\(Note: .*\)\s*$", r"Note:.*\n?", # 노트/주의 문구
        ]
        for pattern in common_patterns:
            processed_text = re.sub(pattern, "", processed_text, flags=re.IGNORECASE).strip()
        
        final_result = processed_text.strip()
        # 번역 결과가 인용 부호만 남은 경우 빈 문자열로 처리
        if re.fullmatch(r"['\"`‘’“”„‚‹›«»]+", final_result):
            return ""
        # 최종적으로 비어있으면 빈 문자열 반환
        if not final_result.strip():
            return ""
        return final_result

    def translate_text(self, text_to_translate:str, src_lang_ui_name:str, tgt_lang_ui_name:str, 
                       model_name:str, ollama_service_instance, is_ocr_text:bool=False):
        is_running, _ = ollama_service_instance.is_running()
        if not is_running:
            logger.error("Ollama 서버 미실행. 번역 불가.")
            return f"오류: Ollama 서버 미실행"
        if not text_to_translate.strip():
            logger.warning("번역할 텍스트가 비어있습니다.")
            return "" # 빈 텍스트는 빈 텍스트로 반환

        try:
            # 언어 매핑은 다양한 언어 지원을 위해 유지하거나 확장 가능
            lang_map = {
                "한국어": "Korean", "일본어": "Japanese", "대만어": "Traditional Chinese",
                "중국어": "Simplified Chinese", "태국어": "Thai", "영어": "English", "스페인어": "Spanish"
                # 필요에 따라 더 많은 언어 추가
            }
            source_lang_name = lang_map.get(src_lang_ui_name, src_lang_ui_name) # UI 이름을 영어 이름으로 변환
            target_lang_name = lang_map.get(tgt_lang_ui_name, tgt_lang_ui_name)

            # 프롬프트 구성
            prompt_instruction = (
                f"Translate the following text from {source_lang_name} to {target_lang_name}.\n"
                f"Provide ONLY the translated text. Do not add any extra words, explanations, or introductory phrases.\n"
                f"Aim for a natural and accurate translation. For proper nouns or specific names, provide an appropriate translation or transliteration in {target_lang_name} if one is commonly understood.\n"
                f"Maintain relevant formatting like line breaks if they are part of the meaning.\n"
            )

            if is_ocr_text:
                prompt = (
                    f"{prompt_instruction}"
                    f"The source text is from OCR and may contain errors. Interpret and translate it as best as possible.\n\n"
                    f"Original OCR text ({source_lang_name}):\n{text_to_translate}\n\n"
                    f"Translated text ({target_lang_name}):\n"
                )
                temperature_setting = config.TRANSLATOR_TEMPERATURE_OCR # config 사용
            else: # 일반 텍스트
                prompt = (
                    f"{prompt_instruction}\n"
                    f"Original text ({source_lang_name}):\n{text_to_translate}\n\n"
                    f"Translated text ({target_lang_name}):\n"
                )
                temperature_setting = config.TRANSLATOR_TEMPERATURE_GENERAL # config 사용
            
            logger.debug(f"번역 프롬프트 ({model_name}, OCR: {is_ocr_text}, Temp: {temperature_setting}): {prompt[:400]}...")
            
            payload = {
                "model": model_name,
                "prompt": prompt,
                "stream": False, # 스트리밍 방식 미사용
                "options": {
                    "temperature": temperature_setting,
                    "num_ctx": 4096, # 컨텍스트 윈도우 크기 (모델에 따라 조절)
                    "top_k": 40,     # Top-k 샘플링
                    "top_p": 0.9     # Top-p (뉴클리어스) 샘플링
                }
            }
            
            # Ollama API 호출
            api_generate_url = f"{ollama_service_instance.url}/api/generate" # ollama_service_instance에 url 속성 사용
            connect_timeout = ollama_service_instance.connect_timeout # config에서 설정된 값 사용
            # 번역 작업의 읽기 타임아웃을 OCR 여부에 따라 동적으로 설정
            # 기본 읽기 타임아웃(ollama_service_instance.read_timeout)을 기준으로 OCR은 더 길게
            translate_read_timeout = ollama_service_instance.read_timeout * (8 if is_ocr_text else 6) 
            if translate_read_timeout <= 0 : translate_read_timeout = 400 # 최소 타임아웃 보장

            response = requests.post(api_generate_url, json=payload, 
                                     timeout=(connect_timeout, translate_read_timeout))
            response.raise_for_status() # HTTP 오류 발생 시 예외 발생
            
            response_data = response.json()
            raw_translated_text = response_data.get("response", "") # 모델의 실제 응답
            
            log_context = f"(OCR, {src_lang_ui_name}->{tgt_lang_ui_name})" if is_ocr_text else f"({src_lang_ui_name}->{tgt_lang_ui_name})"
            logger.info(f"모델 원본 응답 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{raw_translated_text.replace(chr(10),' ')[:70]}'")

            final_translated_text = self._post_process_translation(raw_translated_text, target_lang_ui_name=target_lang_name)
            
            logger.info(f"최종 번역 결과 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{final_translated_text.replace(chr(10),' ')[:70]}'")

            # 후처리 후 결과가 비었으나, 원본 응답은 내용이 있는 경우, 원본 응답(strip) 사용
            if not final_translated_text and raw_translated_text.strip() and text_to_translate.strip():
                logger.warning("후처리 후 번역 결과가 비었으나 원본 응답은 내용이 있어, 원본 응답(strip)을 사용합니다.")
                return raw_translated_text.strip()
                
            return final_translated_text

        except requests.exceptions.Timeout:
            logger.error(f"Ollama 번역 API 시간 초과: {text_to_translate[:30]}...")
            return f"오류: 번역 API 시간 초과"
        except requests.exceptions.RequestException as e:
            logger.error(f"Ollama 번역 API 오류: {e}", exc_info=True)
            return f"오류: 번역 API 오류 ({type(e).__name__})"
        except json.JSONDecodeError as e: # 응답이 JSON이 아닐 경우
            response_text_snippet = response.text[:200] if 'response' in locals() and hasattr(response, 'text') else 'N/A'
            logger.error(f"Ollama 응답 JSON 디코딩 오류: {e}. 응답 일부: {response_text_snippet}", exc_info=True)
            return f"오류: 번역 응답 처리 오류"
        except Exception as e:
            logger.error(f"번역 중 예기치 않은 오류: {e}", exc_info=True)
            return f"오류: 번역 중 내부 오류"