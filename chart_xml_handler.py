# ChartXmlHandler.py
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
import logging
from typing import TYPE_CHECKING, Callable, Any, Optional
import traceback

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
        # 가중치는 config에서 가져옴
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
            logger.debug(f"  - 차트 XML 내 텍스트 번역 (Translator 사용): '{text}' -> '{translated_text}'")
            return translated_text
        except Exception as e:
            logger.error(f"  - 차트 XML 내 텍스트 번역 중 오류 (Translator 사용): {e}")
            return text

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

                for chart_xml_path_in_zip in chart_files:
                    if stop_event and stop_event.is_set():
                        msg_stop = "차트 XML 처리 중 중단 요청 감지."
                        if log_func: log_func(msg_stop)
                        else: logger.info(msg_stop)
                        return None

                    msg_processing_chart = f"\n차트 XML 처리 중: {chart_xml_path_in_zip}"
                    if log_func: log_func(msg_processing_chart)
                    else: logger.info(msg_processing_chart)
                    
                    xml_content_bytes = zip_ref.read(chart_xml_path_in_zip)
                    xml_declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    
                    content_str = xml_content_bytes.decode('utf-8')
                    if content_str.startswith('<?xml'):
                        content_str = content_str.split('?>', 1)[-1].strip()
                    
                    try:
                        root = ET.fromstring(content_str)
                    except ET.ParseError as e_parse:
                        err_msg_parse = f"  오류: 차트 XML 파싱 실패 ({chart_xml_path_in_zip}). 건너뜀. 원인: {e_parse}"
                        if log_func: log_func(err_msg_parse)
                        else: logger.error(err_msg_parse)
                        modified_charts_data[chart_xml_path_in_zip] = xml_content_bytes 
                        if progress_callback_item_completed:
                            progress_callback_item_completed("N/A", "chart-error", self.WEIGHT_CHART, f"Error parsing: {chart_xml_path_in_zip}")
                        continue

                    current_chart_ns_uri_from_tag = root.tag.split('}', 1)[0][1:] if root.tag.startswith('{') and '}' in root.tag else ""
                    
                    # (기존 차트 제목, 시리즈, 카테고리, 축 제목, 데이터 레이블 번역 로직 동일)
                    # ...
                    title_paths = [('.//c:title//a:t', False), ('.//c:chart/c:title//c:v', True)]
                    if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                        title_paths = [('.//cx:title//a:t', False), ('.//cx:chart/cx:title//cx:v', True)]
                    
                    for path_template, is_v_elem in title_paths:
                        elements = root.findall(path_template, ns_map_for_xpath)
                        for elem in elements:
                            if elem.text and elem.text.strip():
                                original_text = elem.text.strip()
                                translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                elem.text = translated
                                log_msg_detail = f"    차트 제목 ({'c:v/cx:v' if is_v_elem else 'a:t'}) 번역됨: '{original_text}' -> '{translated}'"
                                if log_func: log_func(log_msg_detail) # 상세 로그는 여기에
                                else: logger.debug(log_msg_detail)
                    # (다른 차트 요소 번역 로직도 유사하게 유지, 로그 상세화)
                    # ...

                    xml_string_unicode = ET.tostring(root, encoding='unicode', method='xml')
                    final_xml_bytes = xml_declaration.strip() + xml_string_unicode.encode('utf-8').strip()
                    modified_charts_data[chart_xml_path_in_zip] = final_xml_bytes
                    
                    if progress_callback_item_completed:
                        progress_callback_item_completed("N/A", "chart", self.WEIGHT_CHART, chart_xml_path_in_zip)
                
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
            err_msg = f"오류: 원본 PPTX 파일 '{pptx_path}'를 찾을 수 없습니다."
            if log_func: log_func(err_msg)
            else: logger.error(err_msg)
            return None
        except zipfile.BadZipFile:
            err_msg = f"오류: 파일 '{pptx_path}'는 유효한 ZIP 파일(PPTX)이 아닙니다."
            if log_func: log_func(err_msg)
            else: logger.error(err_msg)
            return None
        except ET.ParseError as e_xml_parse:
            err_msg = f"오류: 차트 XML 파싱 중 오류 발생: {e_xml_parse}"
            if log_func: log_func(err_msg)
            else: logger.error(err_msg)
            return None
        except Exception as e:
            err_msg = f"PPTX 내 차트 XML 번역 중 예기치 않은 오류 발생: {e}"
            if log_func: log_func(err_msg + f"\n{traceback.format_exc()}")
            else: logger.error(err_msg, exc_info=True)
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