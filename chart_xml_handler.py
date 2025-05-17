# ChartXmlHandler.py
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
import logging
from typing import TYPE_CHECKING, Callable, Any, Optional
import traceback
import re 

# 설정 파일 import
import config

if TYPE_CHECKING:
    from translator import OllamaTranslator
    from ollama_service import OllamaService

logger = logging.getLogger(__name__)

class ChartXmlHandler:
    def __init__(self, translator_instance: 'OllamaTranslator', ollama_service_instance: 'OllamaService'):
        self.translator = translator_instance
        self.ollama_service = ollama_service_instance
        self.WEIGHT_CHART = config.WEIGHT_CHART

    def _translate_text_via_translator(self, text: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str) -> str:
        if not text or not text.strip():
            return text
        try:
            translated_text = self.translator.translate_text(
                text_to_translate=text,
                src_lang_ui_name=src_lang_ui_name,
                tgt_lang_ui_name=tgt_lang_ui_name,
                model_name=model_name,
                ollama_service_instance=self.ollama_service,
                is_ocr_text=False 
            )
            # logger.debug(f"  - 차트 XML 내 텍스트 번역 (Translator 사용): '{text}' -> '{translated_text}'") # 로그 너무 많을 수 있어 주석 처리
            return translated_text
        except Exception as e:
            logger.error(f"  - 차트 XML 내 텍스트 번역 중 오류 (Translator 사용): {e}")
            return text 

    def _is_numeric_or_simple_symbols(self, text: str) -> bool:
        """주어진 텍스트가 숫자, 일반적인 기호, 또는 매우 짧은 (영어 기준) 문자열인지 확인"""
        if not text:
            return True
        # 순수 숫자 (소수점, 쉼표, 퍼센트, 통화 기호 등 포함 가능성)
        if re.fullmatch(r"[\d.,\s%+\-/*:$€£¥₩#\(\)]+", text):
            return True
        # # (사용자 요청에 따라 이 부분은 일단 제거) 한글/한자/일본어가 아닌 1글자 (예: A, B, C ...) - 범례 등에서 사용될 수 있음
        # if len(text) == 1 and not re.search(r'[가-힣一-龠ぁ-んァ-ヶ]', text):
        #     return True
        return False

    def translate_charts_in_pptx(self, pptx_path: str, src_lang_ui_name: str, tgt_lang_ui_name: str, 
                                 model_name: str, output_path: str = None,
                                 progress_callback_item_completed: Optional[Callable[[Any, str, int, str], None]] = None,
                                 stop_event: Optional[Any] = None,
                                 task_log_filepath: Optional[str] = None) -> Optional[str]:
        
        if output_path is None:
            base_name = os.path.splitext(pptx_path)[0]
            output_path = f"{base_name}_chart_translated.pptx"

        log_func = None
        f_task_log_chart_local = None
        if task_log_filepath:
            try:
                f_task_log_chart_local = open(task_log_filepath, 'a', encoding='utf-8')
                def write_log_chart(message):
                    if f_task_log_chart_local and not f_task_log_chart_local.closed:
                        f_task_log_chart_local.write(message + "\n")
                        f_task_log_chart_local.flush()
                log_func = write_log_chart
            except Exception as e_log_open:
                logger.error(f"ChartXmlHandler: 작업 로그 파일 ({task_log_filepath}) 열기 실패: {e_log_open}")
        
        if log_func:
            log_func(f"\n--- 2단계: 차트 XML 번역 시작 (ChartXmlHandler) ---")
            log_func(f"입력 파일: {os.path.basename(pptx_path)}, 출력 파일: {os.path.basename(output_path)}")
            log_func(f"언어: {src_lang_ui_name} -> {tgt_lang_ui_name}, 모델: {model_name}")
        else:
            logger.info(f"PPTX 내 차트 XML 번역 시작: {os.path.basename(pptx_path)} -> {os.path.basename(output_path)}")
            logger.info(f"소스 언어(UI): {src_lang_ui_name}, 대상 언어(UI): {tgt_lang_ui_name}, 모델: {model_name}")

        temp_dir_for_xml_processing = tempfile.mkdtemp(prefix="chart_xml_")
        
        SCHEMA_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        SCHEMA_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        SCHEMA_CHARTEX = 'http://schemas.microsoft.com/office/drawing/2014/chartex'

        namespaces_to_register = {
            'c': SCHEMA_CHART, 'a': SCHEMA_MAIN,
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'c14': 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
            'c15': 'http://schemas.microsoft.com/office/drawing/2012/chart',
            'c16': 'http://schemas.microsoft.com/office/drawing/2014/chart',
            'c16r2': 'http://schemas.microsoft.com/office/drawing/2015/06/chart',
            'c16r3': 'http://schemas.microsoft.com/office/drawing/2017/03/chart',
            'cx': SCHEMA_CHARTEX
        }
        for prefix, uri in namespaces_to_register.items():
            try: ET.register_namespace(prefix, uri)
            except ValueError: pass 

        ns_map_for_xpath = {'a': SCHEMA_MAIN, 'c': SCHEMA_CHART, 'cx': SCHEMA_CHARTEX}

        try:
            with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
                chart_files = [f for f in zip_ref.namelist() if f.startswith('ppt/charts/') and f.endswith('.xml')]
                msg_chart_files_found = f"총 {len(chart_files)}개의 차트 XML 파일을 발견했습니다."
                if log_func: log_func(msg_chart_files_found)
                else: logger.info(msg_chart_files_found)
                
                modified_charts_data = {} 

                for chart_xml_idx, chart_xml_path_in_zip in enumerate(chart_files):
                    if stop_event and stop_event.is_set():
                        msg_stop = "차트 XML 처리 중 중단 요청 감지."
                        if log_func: log_func(msg_stop)
                        else: logger.info(msg_stop)
                        return None

                    msg_processing_chart = f"\n차트 XML 처리 중 ({chart_xml_idx + 1}/{len(chart_files)}): {chart_xml_path_in_zip}"
                    if log_func: log_func(msg_processing_chart)
                    else: logger.info(msg_processing_chart)
                    
                    xml_content_bytes = zip_ref.read(chart_xml_path_in_zip)
                    content_str = xml_content_bytes.decode('utf-8', errors='ignore')
                    if content_str.lstrip().startswith('<?xml'):
                        content_str = re.sub(r'^\s*<\?xml[^>]*\?>', '', content_str, count=1).strip()
                    
                    try:
                        root = ET.fromstring(content_str)
                    except ET.ParseError as e_parse:
                        err_msg_parse = f"  오류: 차트 XML 파싱 실패 ({chart_xml_path_in_zip}). 건너뜀. 원인: {e_parse}"
                        if log_func: log_func(err_msg_parse)
                        else: logger.error(err_msg_parse)
                        modified_charts_data[chart_xml_path_in_zip] = xml_content_bytes 
                        if progress_callback_item_completed:
                            progress_callback_item_completed(f"차트 {chart_xml_idx + 1}", "차트 오류", self.WEIGHT_CHART / len(chart_files) if chart_files else self.WEIGHT_CHART , f"파싱 오류: {os.path.basename(chart_xml_path_in_zip)}")
                        continue
                    
                    # XPath 표현식 단순화 및 순회 방식으로 변경
                    # 번역할 텍스트를 담고 있을 가능성이 있는 모든 태그를 찾음
                    # ElementTree는 완전한 XPath 2.0을 지원하지 않으므로, 가능한 모든 텍스트 노드를 찾는 방식으로 접근
                    
                    num_translated_in_chart = 0
                    translated_texts_cache_chart = {}

                    for elem in root.iter(): # 모든 하위 요소 순회
                        if stop_event and stop_event.is_set(): break
                        
                        # a:t, c:v, cx:v 태그의 텍스트를 주요 대상으로 하되, 다른 태그도 검사 가능
                        # 여기서는 명시적으로 많이 사용되는 태그의 텍스트를 확인
                        if elem.tag.endswith('}t') or elem.tag.endswith('}v'): # 네임스페이스 무관하게 t 또는 v로 끝나는 태그
                            original_text = elem.text
                            if original_text and original_text.strip():
                                original_text_stripped = original_text.strip()
                                
                                if self._is_numeric_or_simple_symbols(original_text_stripped):
                                    # logger.debug(f"    숫자/기호로 판단되어 번역 스킵: '{original_text_stripped}' (태그: {elem.tag})")
                                    continue
                                # 사용자 요청: "매우 짧은 (1글자) 비대상 언어 텍스트는 번역에서 제외하는 조건" 제거
                                # if len(original_text_stripped) < 2 and not re.search(r'[가-힣一-龠ぁ-んァ-ヶ]', original_text_stripped):
                                #     logger.debug(f"    매우 짧은 비대상 언어 텍스트로 판단되어 번역 스킵: '{original_text_stripped}' (태그: {elem.tag})")
                                #     continue

                                if original_text_stripped in translated_texts_cache_chart:
                                    elem.text = translated_texts_cache_chart[original_text_stripped]
                                else:
                                    translated = self._translate_text_via_translator(original_text_stripped, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                    if "오류:" not in translated and translated.strip() and translated.strip() != original_text_stripped :
                                        elem.text = translated
                                        translated_texts_cache_chart[original_text_stripped] = translated
                                        num_translated_in_chart +=1
                                        log_msg_detail = f"    차트 요소 번역됨 (태그: {elem.tag}): '{original_text_stripped}' -> '{translated}'"
                                        if log_func: log_func(log_msg_detail)
                                        else: logger.debug(log_msg_detail)
                                    elif "오류:" in translated:
                                         log_msg_err = f"    차트 요소 번역 오류 (태그: {elem.tag}): '{original_text_stripped}' -> {translated}"
                                         if log_func: log_func(log_msg_err)
                                         else: logger.warning(log_msg_err)
                        if stop_event and stop_event.is_set(): break
                    
                    if num_translated_in_chart > 0:
                         logger.info(f"  {chart_xml_path_in_zip} 에서 {num_translated_in_chart}개의 텍스트 요소 번역됨.")
                    else:
                         logger.info(f"  {chart_xml_path_in_zip} 에서 번역된 텍스트 요소 없음 (또는 숫자/기호 등으로 스킵됨).")


                    xml_declaration_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    xml_string_unicode = ET.tostring(root, encoding='unicode', method='xml')
                    modified_charts_data[chart_xml_path_in_zip] = xml_declaration_bytes + xml_string_unicode.encode('utf-8')
                    
                    if progress_callback_item_completed:
                        progress_callback_item_completed(f"차트 {chart_xml_idx + 1}", "차트 파일 처리", self.WEIGHT_CHART / len(chart_files) if chart_files else self.WEIGHT_CHART, f"{os.path.basename(chart_xml_path_in_zip)} ({num_translated_in_chart}개 번역)")
                
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for item_name_in_zip in zip_ref.namelist():
                        if item_name_in_zip in modified_charts_data:
                            zip_out.writestr(item_name_in_zip, modified_charts_data[item_name_in_zip])
                        else:
                            zip_out.writestr(item_name_in_zip, zip_ref.read(item_name_in_zip))
            
            final_msg = f"\nPPTX 내 차트 XML 번역 완료! 최종 파일 저장됨: {output_path}"
            if log_func: log_func(final_msg)
            else: logger.info(final_msg)
            return output_path
        
        except FileNotFoundError:
            err_msg_fnf = f"오류: 원본 PPTX 파일 '{pptx_path}'를 찾을 수 없습니다."
            if log_func: log_func(err_msg_fnf)
            else: logger.error(err_msg_fnf)
            return None
        except zipfile.BadZipFile:
            err_msg_zip = f"오류: 파일 '{pptx_path}'는 유효한 ZIP 파일(PPTX)이 아닙니다."
            if log_func: log_func(err_msg_zip)
            else: logger.error(err_msg_zip)
            return None
        except Exception as e_general: # SyntaxError 포함 모든 예외 처리
            err_msg_gen = f"PPTX 내 차트 XML 번역 중 예기치 않은 오류 발생: {e_general}"
            if log_func: log_func(err_msg_gen + f"\n{traceback.format_exc()}")
            else: logger.error(err_msg_gen, exc_info=True)
            return None
        finally:
            if os.path.exists(temp_dir_for_xml_processing):
                try:
                    shutil.rmtree(temp_dir_for_xml_processing)
                    logger.debug(f"임시 디렉토리 '{temp_dir_for_xml_processing}' 삭제 완료.")
                except Exception as e_clean:
                    logger.warning(f"임시 디렉토리 '{temp_dir_for_xml_processing}' 삭제 중 오류: {e_clean}")
            if f_task_log_chart_local and not f_task_log_chart_local.closed:
                f_task_log_chart_local.close()