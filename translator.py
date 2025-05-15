import requests
import json
import logging
import re

logger = logging.getLogger(__name__)

class OllamaTranslator:
    def _post_process_translation(self, text: str, target_lang_ui_name: str = None) -> str:
        if not text:
            logger.debug("후처리 입력: 비어 있음")
            return ""

        # logger.debug(f"후처리 입력 (원본 repr): {repr(text)}")
        processed_text = text

        # 일반적인 모델의 안내문구/부가설명 제거 (정규식 사용, 대소문자 무시)
        # 순서 중요: 구체적인 패턴 먼저, 일반적인 패턴 나중에
        common_patterns = [
            r"^\s*Here is the translation(?: into " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r")?:?\s*",
            r"^\s*The translation is as follows:?\s*",
            r"^\s*Translation:?\s*",
            r"^\s*Translated text:?\s*",
            r"^\s*Sure, here'?s the translation:?\s*",
            r"^\s*Okay, here is the translation:?\s*",
            r"^\s*(?:In " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r"|The " + (target_lang_ui_name if target_lang_ui_name else r"\w+") + r" version is|As requested|Here you go)(?:, the translation is)?:?\s*",
            # 중국어/일본어/한국어 특정 패턴
            r"^\s*以下是翻译.*?[:：]?\s*", r"^\s*翻译如下[:：]?\s*", r"^\s*翻译[:：]?\s*",
            r"^\s*以下為翻譯.*?[:：]?\s*",
            r"^\s*以下は翻訳です.*?[:：]?\s*", r"^\s*翻訳は以下の通りです[:：]?\s*", r"^\s*翻訳[:：]?\s*",
            r"^\s*다음은 번역입니다.*?[:：]?\s*", r"^\s*번역 결과는 다음과 같습니다[:：]?\s*", r"^\s*번역[:：]?\s*",
            # 코드 블록 마커 제거
            r"^\s*```(?:\w+\n)?", r"```\s*$",
            r"^\s*'''", r"'''\s*$",
            r'^\s*"""', r'"""\s*$',
            # XML 스타일 태그 제거 (매우 기본적인 형태만)
            r"^\s*<translation>", r"</translation>\s*$",
            # 모델 자체 설명 문구
            r"\s*I am an AI language model.*", r"\s*As an AI language model.*",
            r"\s*I hope this helps.*", r"\s*Please let me know if you have other questions.*",
            r"\(Note: .*\)\s*$", # 끝에 오는 (Note: ...) 제거
            r"Note:.*\n?", # 시작에 오는 Note: 제거
        ]

        for pattern in common_patterns:
            # re.IGNORECASE 플래그 사용
            processed_text = re.sub(pattern, "", processed_text, flags=re.IGNORECASE).strip()
            # logger.debug(f"패턴 '{pattern}' 적용 후 (repr): {repr(processed_text)}")


        # 시작/끝의 연속적인 줄바꿈, 공백, 탭 제거
        # .strip()은 양쪽 끝의 모든 공백 문자(줄바꿈 포함)를 제거
        final_result = processed_text.strip()
        # logger.debug(f"최종 strip 후 (repr): {repr(final_result)}")

        # 만약 최종 결과가 따옴표, 어포스트로피 등으로만 이루어져 있다면 비움
        # (모델이 번역 대신 입력값을 그대로 따옴표로 감싸 반환하는 경우 대비)
        if re.fullmatch(r"['\"`‘’“”„‚‹›«»]+", final_result):
            logger.debug(f"후처리 결과가 따옴표류만 있어 빈 문자열로 처리: {repr(final_result)}")
            return ""

        # 최종적으로 문자열 전체가 공백 또는 줄바꿈으로만 구성되어 있다면 빈 문자열 반환
        if not final_result.strip(): # 다시 한번 .strip()으로 확인
            logger.debug(f"후처리 최종 결과가 실질적으로 비어있어 빈 문자열로 처리: {repr(final_result)}")
            return ""
        
        # logger.info(f"후처리 최종 결과 (repr): {repr(final_result)}") # 너무 길 수 있으므로 info 대신 debug
        logger.debug(f"후처리 최종 결과 (길이: {len(final_result)}): '{final_result[:100].replace(chr(10), '/')}...'")
        return final_result

    def translate_text(self, text_to_translate:str, src_lang_ui_name:str, tgt_lang_ui_name:str, model_name:str, ollama_service_instance, is_ocr_text:bool=False):
        # (이전 답변과 동일 - 변경 없음)
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

            if is_ocr_text:
                prompt = (
                    f"Translate the following OCR-extracted text from {source_lang_name} to {target_lang_name}.\n"
                    f"This text was extracted from an image and may contain errors or unusual formatting.\n"
                    f"Your goal is to produce a natural and accurate translation in {target_lang_name}, preserving the original meaning.\n"
                    f"Provide ONLY the translated text. Do not add any extra words, explanations, or introductory phrases like \"Here is the translation:\".\n"
                    f"If the input is nonsensical or untranslatable, return the original text as is without any comment or changes.\n\n" # 원본 반환 조건 명확화
                    f"Original OCR text:\n{text_to_translate}\n\n"
                    f"Translated text in {target_lang_name}:\n"
                )
                temperature_setting = 0.4 # OCR은 약간 더 유연하게
            else:
                prompt = (
                    f"Translate the following text from {source_lang_name} to {target_lang_name}.\n"
                    f"Provide ONLY the translated text. Do not add any extra words, explanations, or introductory phrases like \"Here is the translation:\".\n"
                    f"Maintain relevant formatting like line breaks if they are part of the meaning and structure.\n"
                    f"If the input is nonsensical or untranslatable, return the original text as is without any comment or changes.\n\n" # 원본 반환 조건 명확화
                    f"Original text:\n{text_to_translate}\n\n"
                    f"Translated text in {target_lang_name}:\n"
                )
                temperature_setting = 0.2 # 일반 텍스트는 더 정확하게

            logger.debug(f"번역 프롬프트 ({model_name}, OCR: {is_ocr_text}, Temp: {temperature_setting}): {prompt[:350]}...")
            payload = {
                "model": model_name, "prompt": prompt, "stream": False,
                "options": {"temperature": temperature_setting, "num_ctx": 4096, "top_k": 40, "top_p": 0.9}
            }
            api_generate_url = f"{ollama_service_instance.url}/api/generate"
            connect_timeout = ollama_service_instance.connect_timeout
            translate_read_timeout = ollama_service_instance.read_timeout * (6 if is_ocr_text else 4)
            if translate_read_timeout <= 0 : translate_read_timeout = 300 
            
            response = requests.post(api_generate_url, json=payload, timeout=(connect_timeout, translate_read_timeout))
            response.raise_for_status()
            response_data = response.json()
            raw_translated_text = response_data.get("response", "")
            
            log_context = f"(OCR, {src_lang_ui_name}->{tgt_lang_ui_name})" if is_ocr_text else f"({src_lang_ui_name}->{tgt_lang_ui_name})"
            logger.info(f"모델 원본 응답 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{raw_translated_text.replace(chr(10),' ')[:70]}'")
            
            final_translated_text = self._post_process_translation(raw_translated_text, target_lang_ui_name=target_lang_name)
            
            logger.info(f"최종 번역 결과 {log_context}: '{text_to_translate.replace(chr(10),' ')[:50]}' -> '{final_translated_text.replace(chr(10),' ')[:70]}'")
            
            # 후처리 후 비었으나, 원본 응답에 내용이 있고, 원본 텍스트도 내용이 있는 경우, 원본 응답(strip) 사용
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
            logger.error(f"Ollama 응답 JSON 디코딩 오류: {e}. 응답: {response.text if 'response' in locals() else 'N/A'}", exc_info=True)
            return f"오류: 번역 응답 처리 오류"
        except Exception as e:
            logger.error(f"번역 중 예기치 않은 오류: {e}", exc_info=True)
            return f"오류: 번역 중 내부 오류"