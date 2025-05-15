import requests
import json
import logging
import re

logger = logging.getLogger(__name__)

class OllamaTranslator:
    def _post_process_translation(self, text: str, target_lang_ui_name: str = None) -> str:
        # (이 함수는 이전 답변과 동일하게 유지 - 안정적인 후처리 로직)
        if not text:
            logger.debug("후처리 입력: 비어 있음")
            return ""
        processed_text = text
        common_patterns = [
            r"^\s*Here is the translation(?: into " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r")?:?\s*",
            r"^\s*The translation is as follows:?\s*",
            r"^\s*Translation:?\s*", r"^\s*Translated text:?\s*",
            r"^\s*Sure, here'?s the translation:?\s*", r"^\s*Okay, here is the translation:?\s*",
            r"^\s*(?:In " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r"|The " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r" version is|As requested|Here you go)(?:, the translation is)?:?\s*",
            r"^\s*以下是翻译.*?[:：]?\s*", r"^\s*翻译如下[:：]?\s*", r"^\s*翻译[:：]?\s*",
            r"^\s*以下為翻譯.*?[:：]?\s*",
            r"^\s*以下は翻訳です.*?[:：]?\s*", r"^\s*翻訳は以下の通りです[:：]?\s*", r"^\s*翻訳[:：]?\s*",
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
        if re.fullmatch(r"['\"`‘’“”„‚‹›«»]+", final_result):
            logger.debug(f"후처리 결과가 따옴표류만 있어 빈 문자열로 처리: {repr(final_result)}")
            return ""
        if not final_result.strip():
            logger.debug(f"후처리 최종 결과가 실질적으로 비어있어 빈 문자열로 처리: {repr(final_result)}")
            return ""
        logger.debug(f"후처리 최종 결과 (길이: {len(final_result)}): '{final_result[:100].replace(chr(10), '/')}...'")
        return final_result

    def translate_text(self, text_to_translate:str, src_lang_ui_name:str, tgt_lang_ui_name:str, model_name:str, ollama_service_instance, is_ocr_text:bool=False):
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

            # --- "최초 안정 프롬프트" 기준으로 복구 및 최소한의 고유명사 처리 요청 ---
            # 이전에 안정적이었던 기본 프롬프트 구조를 사용합니다.
            # (정확한 "최초" 프롬프트를 알 수 없으므로, 일반적이고 안정적인 형태로 구성)

            # 기본 지시: 간결하게 번역 요청, 부가 설명 없이 번역문만 출력.
            # 고유명사에 대한 언급은 하되, 강제성을 낮춤.
            prompt_instruction = (
                f"Translate the following text from {source_lang_name} to {target_lang_name}.\n"
                f"Provide ONLY the translated text. Do not add any extra words, explanations, or introductory phrases.\n"
                f"Aim for a natural and accurate translation. For proper nouns or specific names, provide an appropriate translation or transliteration in {target_lang_name} if one is commonly understood.\n" # 고유명사 처리 요청 (부드럽게)
                f"Maintain relevant formatting like line breaks if they are part of the meaning.\n"
                # "번역 불가 시 원문 반환" 같은 강제 지침은 제거하거나 매우 완화. 모델의 자연스러운 판단에 맡기는 부분 증가.
                # 필요하다면, 매우 짧고 의미없는 텍스트에 대해서만 원문 반환을 고려할 수 있으나, 우선은 제외.
            )

            if is_ocr_text:
                prompt = (
                    f"{prompt_instruction}"
                    f"The source text is from OCR and may contain errors. Interpret and translate it as best as possible.\n\n"
                    f"Original OCR text ({source_lang_name}):\n{text_to_translate}\n\n"
                    f"Translated text ({target_lang_name}):\n"
                )
                temperature_setting = 0.5 # OCR 텍스트는 약간의 유연성 (이전 0.3과 0.45 사이)
            else: # 일반 텍스트
                prompt = (
                    f"{prompt_instruction}\n"
                    f"Original text ({source_lang_name}):\n{text_to_translate}\n\n"
                    f"Translated text ({target_lang_name}):\n"
                )
                temperature_setting = 0.3 # 일반 텍스트는 안정성 중시 (이전 0.15와 0.25 사이)
            # --- 프롬프트 수정 끝 ---

            logger.debug(f"번역 프롬프트 ({model_name}, OCR: {is_ocr_text}, Temp: {temperature_setting}): {prompt[:400]}...")
            payload = {
                "model": model_name, "prompt": prompt, "stream": False,
                "options": {"temperature": temperature_setting, "num_ctx": 4096, "top_k": 40, "top_p": 0.9}
            }
            api_generate_url = f"{ollama_service_instance.url}/api/generate"
            connect_timeout = ollama_service_instance.connect_timeout
            # 번역 API 타임아웃은 충분히 길게 설정
            translate_read_timeout = ollama_service_instance.read_timeout * (8 if is_ocr_text else 6) # OCR은 더 길게
            if translate_read_timeout <= 0 : translate_read_timeout = 400 # 최소 타임아웃 보장
            
            response = requests.post(api_generate_url, json=payload, timeout=(connect_timeout, translate_read_timeout))
            response.raise_for_status() # HTTP 오류 발생 시 예외 발생
            response_data = response.json()
            raw_translated_text = response_data.get("response", "") # 모델의 순수 응답
            
            log_context = f"(OCR, {src_lang_ui_name}->{tgt_lang_ui_name})" if is_ocr_text else f"({src_lang_ui_name}->{tgt_lang_ui_name})"
            # 로그 출력 시 줄바꿈 문자를 공백으로 치환하여 한 줄로 보기 쉽게
            logger.info(f"모델 원본 응답 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{raw_translated_text.replace(chr(10),' ')[:70]}'")
            
            # 후처리 함수는 모델의 부가 설명 제거에 주로 사용
            final_translated_text = self._post_process_translation(raw_translated_text, target_lang_ui_name=target_lang_name)
            
            logger.info(f"최종 번역 결과 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{final_translated_text.replace(chr(10),' ')[:70]}'")
            
            # 후처리 후 비었으나, 원본 응답에는 내용이 있고, 번역할 원본 텍스트도 내용이 있는 경우,
            # 후처리가 너무 과도하게 작용했을 수 있으므로, 원본 응답(strip)을 사용
            if not final_translated_text and raw_translated_text.strip() and text_to_translate.strip():
                logger.warning("후처리 후 번역 결과가 비었으나 원본 응답은 내용이 있어, 원본 응답(strip)을 사용합니다.")
                return raw_translated_text.strip()
            
            return final_translated_text # 후처리된 최종 번역 결과 반환
        
        except requests.exceptions.Timeout:
            logger.error(f"Ollama 번역 API 시간 초과: {text_to_translate[:30]}...")
            return f"오류: 번역 API 시간 초과"
        except requests.exceptions.RequestException as e: # 모든 requests 관련 예외 포함
            logger.error(f"Ollama 번역 API 오류: {e}", exc_info=True)
            return f"오류: 번역 API 오류 ({type(e).__name__})"
        except json.JSONDecodeError as e: # JSON 파싱 실패 시
            response_text_snippet = response.text[:200] if 'response' in locals() and hasattr(response, 'text') else 'N/A'
            logger.error(f"Ollama 응답 JSON 디코딩 오류: {e}. 응답 일부: {response_text_snippet}", exc_info=True)
            return f"오류: 번역 응답 처리 오류"
        except Exception as e: # 그 외 예측하지 못한 모든 오류
            logger.error(f"번역 중 예기치 않은 오류: {e}", exc_info=True)
            return f"오류: 번역 중 내부 오류"