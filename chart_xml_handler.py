# ChartXmlHandler.py
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
import logging
from typing import Callable, Any, Optional, List, Dict, TYPE_CHECKING # TYPE_CHECKING 추가
import traceback
import re 

# 설정 파일 import
import config
from interfaces import AbsChartProcessor, AbsTranslator, AbsOllamaService # 인터페이스 import

# TYPE_CHECKING 블록은 translator, ollama_service가 Abs... 타입일 경우 제거 가능
# if TYPE_CHECKING:
#     from translator import OllamaTranslator # 이제 AbsTranslator 사용
#     from ollama_service import OllamaService # 이제 AbsOllamaService 사용

logger = logging.getLogger(__name__)

class ChartXmlHandler(AbsChartProcessor): # AbsChartProcessor 상속
    def __init__(self, translator_instance: AbsTranslator, ollama_service_instance: AbsOllamaService): # 타입 힌트 변경
        self.translator = translator_instance
        self.ollama_service = ollama_service_instance
        self.WEIGHT_CHART = config.WEIGHT_CHART

    def _is_numeric_or_simple_symbols(self, text: str) -> bool:
        if not text: return True
        if re.fullmatch(r"[\d.,\s%+\-/*:$€£¥₩#\(\)]+", text): return True
        if len(text) == 1 and not re.search(r'[가-힣一-龠ぁ-んァ-ヶ]', text): return True
        return False

    def translate_charts_in_pptx(self, pptx_path: str, src_lang_ui_name: str, tgt_lang_ui_name: str, 
                                 model_name: str, output_path: str = None,
                                 progress_callback_item_completed: Optional[Callable[[Any, str, int, str], None]] = None,
                                 stop_event: Optional[Any] = None, # threading.Event로 변경 가능
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
                        f_task_log_chart_local.write(message + "\n"); f_task_log_chart_local.flush()
                log_func = write_log_chart
            except Exception as e_log_open: logger.error(f"ChartXmlHandler: 작업 로그 파일 ({task_log_filepath}) 열기 실패: {e_log_open}")
        
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
        namespaces_to_register = {'c': SCHEMA_CHART, 'a': SCHEMA_MAIN, 'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006', 'c14': 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart', 'c15': 'http://schemas.microsoft.com/office/drawing/2012/chart', 'c16': 'http://schemas.microsoft.com/office/drawing/2014/chart', 'c16r2': 'http://schemas.microsoft.com/office/drawing/2015/06/chart', 'c16r3': 'http://schemas.microsoft.com/office/drawing/2017/03/chart', 'cx': SCHEMA_CHARTEX}
        for prefix, uri in namespaces_to_register.items():
            try: ET.register_namespace(prefix, uri)
            except ValueError: pass 

        try:
            unique_texts_to_translate_all_charts: Dict[str, None] = {}
            chart_xml_contents_map: Dict[str, bytes] = {}

            with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
                chart_files = [f for f in zip_ref.namelist() if f.startswith('ppt/charts/') and f.endswith('.xml')]
                msg_chart_files_found = f"총 {len(chart_files)}개의 차트 XML 파일을 발견했습니다. 텍스트 수집 중..."
                if log_func: log_func(msg_chart_files_found)
                else: logger.info(msg_chart_files_found)

                for chart_xml_path_in_zip in chart_files:
                    if stop_event and stop_event.is_set(): break
                    xml_content_bytes = zip_ref.read(chart_xml_path_in_zip)
                    chart_xml_contents_map[chart_xml_path_in_zip] = xml_content_bytes
                    content_str = xml_content_bytes.decode('utf-8', errors='ignore')
                    if content_str.lstrip().startswith('<?xml'): content_str = re.sub(r'^\s*<\?xml[^>]*\?>', '', content_str, count=1).strip()
                    try:
                        root = ET.fromstring(content_str)
                        for elem in root.iter():
                            if stop_event and stop_event.is_set(): break
                            if elem.tag.endswith('}t') or elem.tag.endswith('}v'):
                                original_text = elem.text
                                if original_text and original_text.strip():
                                    original_text_stripped = original_text.strip()
                                    if not self._is_numeric_or_simple_symbols(original_text_stripped):
                                        unique_texts_to_translate_all_charts[original_text_stripped] = None
                    except ET.ParseError as e_parse:
                        err_msg_parse = f"  오류: 차트 XML 파싱 실패 ({chart_xml_path_in_zip}) - 텍스트 수집 건너뜀. 원인: {e_parse}"
                        if log_func: log_func(err_msg_parse)
                        else: logger.error(err_msg_parse)
                
                if stop_event and stop_event.is_set():
                    msg_stop_collect = "차트 텍스트 수집 중 중단 요청 감지."
                    if log_func: log_func(msg_stop_collect); return None
                    else: logger.info(msg_stop_collect); return None

                texts_list_for_batch = list(unique_texts_to_translate_all_charts.keys())
                translation_map: Dict[str, str] = {}

                if texts_list_for_batch:
                    msg_batch_start = f"차트 내 고유 텍스트 {len(texts_list_for_batch)}개 일괄 번역 시작..."
                    if log_func: log_func(msg_batch_start)
                    else: logger.info(msg_batch_start)
                    translated_texts_batch = self.translator.translate_texts_batch(texts_list_for_batch, src_lang_ui_name, tgt_lang_ui_name, model_name, self.ollama_service, is_ocr_text=False, stop_event=stop_event)
                    if stop_event and stop_event.is_set():
                        msg_stop_batch = "차트 텍스트 일괄 번역 중 중단 요청 감지."
                        if log_func: log_func(msg_stop_batch); return None
                        else: logger.info(msg_stop_batch); return None
                    if len(texts_list_for_batch) == len(translated_texts_batch):
                        for original, translated in zip(texts_list_for_batch, translated_texts_batch): translation_map[original] = translated
                        msg_batch_done = f"차트 내 고유 텍스트 일괄 번역 완료. {len(translation_map)}개 매핑 생성."
                        if log_func: log_func(msg_batch_done)
                        else: logger.info(msg_batch_done)
                    else:
                        warn_msg_mismatch = f"경고: 차트 원본 텍스트 수({len(texts_list_for_batch)})와 번역 결과 수({len(translated_texts_batch)}) 불일치!"
                        if log_func: log_func(warn_msg_mismatch)
                        else: logger.warning(warn_msg_mismatch)
                        return None

                modified_charts_data: Dict[str, bytes] = {} 
                total_charts = len(chart_files)
                processed_charts_count = 0

                for chart_xml_idx, chart_xml_path_in_zip in enumerate(chart_files):
                    if stop_event and stop_event.is_set(): break
                    msg_processing_chart = f"\n차트 XML 적용 중 ({chart_xml_idx + 1}/{total_charts}): {chart_xml_path_in_zip}"
                    if log_func: log_func(msg_processing_chart)
                    else: logger.info(msg_processing_chart)
                    xml_content_bytes = chart_xml_contents_map[chart_xml_path_in_zip]
                    content_str = xml_content_bytes.decode('utf-8', errors='ignore')
                    if content_str.lstrip().startswith('<?xml'): content_str = re.sub(r'^\s*<\?xml[^>]*\?>', '', content_str, count=1).strip()
                    num_translated_in_chart = 0
                    try:
                        root = ET.fromstring(content_str)
                        for elem in root.iter():
                            if stop_event and stop_event.is_set(): break
                            if elem.tag.endswith('}t') or elem.tag.endswith('}v'):
                                original_text = elem.text
                                if original_text and original_text.strip():
                                    original_text_stripped = original_text.strip()
                                    if original_text_stripped in translation_map:
                                        translated = translation_map[original_text_stripped]
                                        if "오류:" not in translated and translated.strip() and translated.strip() != original_text_stripped:
                                            elem.text = translated; num_translated_in_chart += 1
                                            log_msg_detail = f"    차트 요소 번역됨 (태그: {elem.tag}): '{original_text_stripped}' -> '{translated}'"
                                            if log_func: log_func(log_msg_detail)
                                        elif "오류:" in translated:
                                            log_msg_err = f"    차트 요소 번역 오류 (태그: {elem.tag}, 원본: '{original_text_stripped}') -> {translated}"
                                            if log_func: log_func(log_msg_err)
                                            else: logger.warning(log_msg_err)
                        if num_translated_in_chart > 0: logger.info(f"  {chart_xml_path_in_zip} 에서 {num_translated_in_chart}개의 텍스트 요소 번역됨.")
                        else: logger.info(f"  {chart_xml_path_in_zip} 에서 번역된 텍스트 요소 없음 (또는 숫자/기호 등으로 스킵됨 / 이미 번역됨).")
                        xml_declaration_bytes = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        xml_string_unicode = ET.tostring(root, encoding='unicode', method='xml')
                        modified_charts_data[chart_xml_path_in_zip] = xml_declaration_bytes + xml_string_unicode.encode('utf-8')
                    except ET.ParseError as e_parse_apply:
                        err_msg_parse_apply = f"  오류: 차트 XML 재파싱/적용 실패 ({chart_xml_path_in_zip}). 원본 사용. 원인: {e_parse_apply}"
                        if log_func: log_func(err_msg_parse_apply)
                        else: logger.error(err_msg_parse_apply)
                        modified_charts_data[chart_xml_path_in_zip] = xml_content_bytes
                    
                    processed_charts_count +=1
                    if progress_callback_item_completed:
                        progress_callback_item_completed(f"차트 {chart_xml_idx + 1}", "차트 파일 처리", self.WEIGHT_CHART / total_charts if total_charts > 0 else self.WEIGHT_CHART, f"{os.path.basename(chart_xml_path_in_zip)} ({num_translated_in_chart}개 번역)")
                
                if stop_event and stop_event.is_set():
                    msg_stop_apply = "차트 XML 적용 중 중단 요청 감지."
                    if log_func: log_func(msg_stop_apply); return None
                    else: logger.info(msg_stop_apply); return None

                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for item_name_in_zip in zip_ref.namelist():
                        if item_name_in_zip in modified_charts_data:
                            zip_out.writestr(item_name_in_zip, modified_charts_data[item_name_in_zip])
                        else:
                            zip_out.writestr(item_name_in_zip, zip_ref.read(item_name_in_zip))
            
            final_msg = f"\nPPTX 내 차트 XML 번역 완료! ({processed_charts_count}/{total_charts}개 차트 처리) 최종 파일 저장됨: {output_path}"
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
        except Exception as e_general:
            err_msg_gen = f"PPTX 내 차트 XML 번역 중 예기치 않은 오류 발생: {e_general}"
            if log_func: log_func(err_msg_gen + f"\n{traceback.format_exc()}")
            else: logger.error(err_msg_gen, exc_info=True)
            return None
        finally:
            if os.path.exists(temp_dir_for_xml_processing):
                try: shutil.rmtree(temp_dir_for_xml_processing)
                except Exception as e_clean: logger.warning(f"임시 디렉토리 '{temp_dir_for_xml_processing}' 삭제 중 오류: {e_clean}")
            if f_task_log_chart_local and not f_task_log_chart_local.closed: f_task_log_chart_local.close()