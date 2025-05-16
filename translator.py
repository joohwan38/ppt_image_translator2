import requests
import json
import logging
import re

logger = logging.getLogger(__name__)

class OllamaTranslator:
    def _post_process_translation(self, text: str, target_lang_ui_name: str = None) -> str:
        # (이전 답변의 안정적인 후처리 로직 유지)
        if not text: return ""
        processed_text = text
        common_patterns = [
            r"^\s*Here is the translation(?: into " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r")?:?\s*",
            r"^\s*The translation is as follows:?\s*",
            r"^\s*Translation:?\s*", r"^\s*Translated text:?\s*",
            r"^\s*Sure, here'?s the translation:?\s*", r"^\s*Okay, here is the translation:?\s*",
            r"^\s*(?:In " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r"|The " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r" version is|As requested|Here you go)(?:, the translation is)?:?\s*",
            r"^\s*以下是翻译.*?[:：]?\s*", r"^\s*翻译如下[:：]?\s*", r"^\s*翻译[:：]?\s*",
            r"^\s*以下為翻譯.*?[:：]?\s*", r"^\s*以下は翻訳です.*?[:：]?\s*", r"^\s*翻訳は以下の通りです[:：]?\s*", r"^\s*翻訳[:：]?\s*",
            r"^\s*다음은 번역입니다.*?[:：]?\s*", r"^\s*번역 결과는 다음과 같습니다[:：]?\s*", r"^\s*번역[:：]?\s*",
            r"^\s*```(?:\w+\n)?", r"```\s*$",
            r"^\s*'''", r"'''\s*$", r'^\s*"""', r'"""\s*$',
            r"^\s*<translation>", r"</translation>\s*$",
            r"\s*I am an AI language model.*", r"\s*As an AI language model.*",
            r"\s*I hope this helps.*", r"\s*Please let me know if you have other questions.*",
            r"\(Note: .*\)\s*$", r"Note:.*\n?",
        ]
        for pattern in common_patterns:
            processed_text = re.sub(pattern, "", processed_text, flags=re.IGNORECASE).strip()
        final_result = processed_text.strip()
        if re.fullmatch(r"['\"`‘’“”„‚‹›«»]+", final_result): return ""
        if not final_result.strip(): return ""
        return final_result

    def translate_text(self, text_to_translate:str, src_lang_ui_name:str, tgt_lang_ui_name:str, model_name:str, ollama_service_instance, is_ocr_text:bool=False):
        # ... (함수 시작 부분 및 lang_map 등은 이전과 동일)
        is_running, _ = ollama_service_instance.is_running()
        if not is_running:
            logger.error("Ollama 서버 미실행. 번역 불가.")
            return f"오류: Ollama 서버 미실행"
        if not text_to_translate.strip():
            logger.warning("번역할 텍스트가 비어있습니다.")
            return ""
        try:
            lang_map = {
                "한국어": "Korean", "일본어": "Japanese", "대만어": "Traditional Chinese",
                "중국어": "Simplified Chinese", "태국어": "Thai", "영어": "English", "스페인어": "Spanish"
            }
            source_lang_name = lang_map.get(src_lang_ui_name, src_lang_ui_name)
            target_lang_name = lang_map.get(tgt_lang_ui_name, tgt_lang_ui_name)

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
                temperature_setting = 0.4 # 사용자 요청: OCR 0.5
            else: # 일반 텍스트
                prompt = (
                    f"{prompt_instruction}\n"
                    f"Original text ({source_lang_name}):\n{text_to_translate}\n\n"
                    f"Translated text ({target_lang_name}):\n"
                )
                temperature_setting = 0.2 # 사용자 요청: 일반 0.3
            
            logger.debug(f"번역 프롬프트 ({model_name}, OCR: {is_ocr_text}, Temp: {temperature_setting}): {prompt[:400]}...")
            payload = {
                "model": model_name, "prompt": prompt, "stream": False,
                "options": {"temperature": temperature_setting, "num_ctx": 4096, "top_k": 40, "top_p": 0.9}
            }
            # ... (API 호출 및 응답 처리 로직은 이전과 동일) ...
            api_generate_url = f"{ollama_service_instance.url}/api/generate"
            connect_timeout = ollama_service_instance.connect_timeout
            translate_read_timeout = ollama_service_instance.read_timeout * (8 if is_ocr_text else 6) 
            if translate_read_timeout <= 0 : translate_read_timeout = 400 
            response = requests.post(api_generate_url, json=payload, timeout=(connect_timeout, translate_read_timeout))
            response.raise_for_status()
            response_data = response.json()
            raw_translated_text = response_data.get("response", "")
            log_context = f"(OCR, {src_lang_ui_name}->{tgt_lang_ui_name})" if is_ocr_text else f"({src_lang_ui_name}->{tgt_lang_ui_name})"
            logger.info(f"모델 원본 응답 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{raw_translated_text.replace(chr(10),' ')[:70]}'")
            final_translated_text = self._post_process_translation(raw_translated_text, target_lang_ui_name=target_lang_name)
            logger.info(f"최종 번역 결과 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{final_translated_text.replace(chr(10),' ')[:70]}'")
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
        except json.JSONDecodeError as e:
            response_text_snippet = response.text[:200] if 'response' in locals() and hasattr(response, 'text') else 'N/A'
            logger.error(f"Ollama 응답 JSON 디코딩 오류: {e}. 응답 일부: {response_text_snippet}", exc_info=True)
            return f"오류: 번역 응답 처리 오류"
        except Exception as e:
            logger.error(f"번역 중 예기치 않은 오류: {e}", exc_info=True)
            return f"오류: 번역 중 내부 오류"