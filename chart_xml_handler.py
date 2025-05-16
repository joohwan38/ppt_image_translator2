# ChartXmlHandler.py
import zipfile
import xml.etree.ElementTree as ET
import os
import tempfile
import shutil
import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from translator import OllamaTranslator
    from ollama_service import OllamaService # OllamaTranslator가 이를 필요로 하므로, 타입 힌팅에 추가

logger = logging.getLogger(__name__)

class ChartXmlHandler:
    def __init__(self, translator_instance: 'OllamaTranslator', ollama_service_instance: 'OllamaService'):
        """
        ChartXmlHandler를 초기화합니다.

        Args:
            translator_instance: 번역 작업을 수행할 OllamaTranslator의 인스턴스입니다.
            ollama_service_instance: Ollama 서비스 상태 확인 및 모델 정보 관리를 위한 OllamaService의 인스턴스입니다.
                                     (OllamaTranslator가 내부적으로 사용할 수 있습니다.)
        """
        self.translator = translator_instance
        self.ollama_service = ollama_service_instance # Translator가 OllamaService를 직접 참조/사용할 경우를 대비

    def _translate_text_via_translator(self, text: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str) -> str:
        """
        주입된 translator 인스턴스를 사용하여 텍스트를 번역합니다.
        Ollama API를 직접 호출하는 대신 중앙화된 번역 로직을 사용합니다.
        """
        if not text or not text.strip():
            return text
        try:
            # translator.translate_text는 is_ocr_text 파라미터를 받지만, 차트 텍스트는 OCR 텍스트가 아니므로 False로 설정
            translated_text = self.translator.translate_text(
                text_to_translate=text,
                src_lang_ui_name=src_lang_ui_name,
                tgt_lang_ui_name=tgt_lang_ui_name,
                model_name=model_name,
                ollama_service_instance=self.ollama_service, # OllamaTranslator의 요구사항에 따라 전달
                is_ocr_text=False # 차트 텍스트는 OCR 텍스트가 아님
            )
            logger.debug(f"  - 차트 XML 내 텍스트 번역 (Translator 사용): '{text}' -> '{translated_text}'")
            return translated_text
        except Exception as e:
            logger.error(f"  - 차트 XML 내 텍스트 번역 중 오류 (Translator 사용): {e}")
            return text # 오류 발생 시 원본 텍스트 반환

    def translate_charts_in_pptx(self, pptx_path: str, src_lang_ui_name: str, tgt_lang_ui_name: str, model_name: str, output_path: str = None) -> str:
        """
        PowerPoint 파일 내의 차트 텍스트를 번역하고, 번역된 파일을 지정된 경로 또는 기본 경로에 저장합니다.

        Args:
            pptx_path: 번역할 원본 PowerPoint 파일 경로입니다.
            src_lang_ui_name: 원본 언어의 UI 표시 이름 (예: "영어", "한국어").
            tgt_lang_ui_name: 대상 언어의 UI 표시 이름 (예: "영어", "한국어").
            model_name: 번역에 사용할 Ollama 모델 이름입니다.
            output_path: 번역된 파일을 저장할 경로입니다. None이면 원본 파일명에 "_chart_translated"를 붙여 저장합니다.

        Returns:
            번역된 PowerPoint 파일의 경로를 반환합니다. 오류 발생 시 None을 반환할 수 있습니다.
        """
        if output_path is None:
            base_name = os.path.splitext(pptx_path)[0]
            output_path = f"{base_name}_chart_translated.pptx" # PptxHandler에서 임시 파일명을 사용할 것이므로, 이 파일명은 최종 결과 파일명으로 고려

        logger.info(f"PPTX 내 차트 XML 번역 시작: {os.path.basename(pptx_path)} -> {os.path.basename(output_path)}")
        logger.info(f"소스 언어(UI): {src_lang_ui_name}, 대상 언어(UI): {tgt_lang_ui_name}, 모델: {model_name}")

        temp_dir_for_xml_processing = tempfile.mkdtemp(prefix="chart_xml_")
        
        # 네임스페이스 URI 정의
        SCHEMA_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        SCHEMA_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
        SCHEMA_CHARTEX = 'http://schemas.microsoft.com/office/drawing/2014/chartex' # 확장 차트

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
            except ValueError: pass # 이미 등록된 경우 무시

        ns_map_for_xpath = {'a': SCHEMA_MAIN, 'c': SCHEMA_CHART, 'cx': SCHEMA_CHARTEX}

        try:
            with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
                chart_files = [f for f in zip_ref.namelist() if f.startswith('ppt/charts/') and f.endswith('.xml')]
                logger.info(f"총 {len(chart_files)}개의 차트 XML 파일을 발견했습니다.")
                
                modified_charts_data = {} 

                for chart_xml_path_in_zip in chart_files:
                    logger.info(f"\n차트 XML 처리 중: {chart_xml_path_in_zip}")
                    
                    xml_content_bytes = zip_ref.read(chart_xml_path_in_zip)
                    xml_declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    
                    content_str = xml_content_bytes.decode('utf-8')
                    if content_str.startswith('<?xml'):
                        # XML 선언부가 있다면 제거 (ElementTree.fromstring은 선언부 없이 순수 XML 내용을 기대)
                        content_str = content_str.split('?>', 1)[-1].strip()
                    
                    try:
                        root = ET.fromstring(content_str)
                    except ET.ParseError as e_parse:
                        logger.error(f"  오류: 차트 XML 파싱 실패 ({chart_xml_path_in_zip}). 건너뜀. 원인: {e_parse}")
                        modified_charts_data[chart_xml_path_in_zip] = xml_content_bytes # 원본 데이터 그대로 유지
                        continue

                    current_chart_ns_uri_from_tag = root.tag.split('}', 1)[0][1:] if root.tag.startswith('{') and '}' in root.tag else ""

                    # 1. 차트 제목 번역 (c:title//a:t 또는 cx:title//a:t, c:chart/c:title//c:v 또는 cx용)
                    title_paths = [
                        ('.//c:title//a:t', False), # (path, is_v_element)
                        ('.//c:chart/c:title//c:v', True)
                    ]
                    if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                        title_paths = [
                            ('.//cx:title//a:t', False), # cx 네임스페이스 사용
                            ('.//cx:chart/cx:title//cx:v', True) # cx의 v 요소 (구조 확인 필요)
                        ]
                    
                    for path_template, is_v_elem in title_paths:
                        elements = root.findall(path_template, ns_map_for_xpath)
                        for elem in elements:
                            if elem.text and elem.text.strip():
                                original_text = elem.text.strip()
                                translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                elem.text = translated
                                if is_v_elem: logger.debug(f"    차트 제목 (c:v) 번역됨: '{original_text}' -> '{translated}'")
                                else: logger.debug(f"    차트 제목 (a:t) 번역됨: '{original_text}' -> '{translated}'")
                    
                    # 2. 시리즈 이름(범례) 번역
                    ser_path_template = './/c:ser'
                    tx_path_in_ser = './/c:tx'
                    v_path_in_tx = './/c:v' # 직접적인 값
                    str_cache_path_in_tx = './/c:strRef/c:strCache//c:pt/c:v' # 캐시된 값
                    tx_data_v_path_in_tx = './/cx:txData//cx:v' # cx의 txData 내 v (chartex)

                    if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                        ser_path_template = './/cx:series' # 또는 cx:ser
                        # cx:tx, cx:v, cx:strRef/cx:strCache//cx:pt/cx:v 등은 일반 차트와 유사할 수 있음
                        # 네임스페이스 접두사만 변경하여 경로 재활용 시도
                        tx_path_in_ser = './/cx:tx'
                        v_path_in_tx = './/cx:v'
                        str_cache_path_in_tx = './/cx:strRef/cx:strCache//cx:pt/cx:v'


                    series_elements = root.findall(ser_path_template, ns_map_for_xpath)
                    for ser_idx, ser_elem in enumerate(series_elements):
                        tx_node = ser_elem.find(tx_path_in_ser, ns_map_for_xpath)
                        if tx_node is not None:
                            # 직접적인 값 (c:tx/c:v 또는 cx:tx/cx:v)
                            v_direct = tx_node.find(v_path_in_tx, ns_map_for_xpath)
                            if v_direct is not None and v_direct.text and v_direct.text.strip():
                                original_text = v_direct.text.strip()
                                translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                v_direct.text = translated
                                logger.debug(f"    시리즈 {ser_idx} 이름 (직접 c:v/cx:v) 번역됨: '{original_text}' -> '{translated}'")
                            
                            # 캐시된 값 (c:tx/c:strRef/c:strCache//c:pt/c:v 또는 cx 버전)
                            for v_cache_idx, v_cache_elem in enumerate(tx_node.findall(str_cache_path_in_tx, ns_map_for_xpath)):
                                if v_cache_elem.text and v_cache_elem.text.strip():
                                    original_text = v_cache_elem.text.strip()
                                    translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                    v_cache_elem.text = translated
                                    logger.debug(f"    시리즈 {ser_idx} 이름 (캐시 pt {v_cache_idx} c:v/cx:v) 번역됨: '{original_text}' -> '{translated}'")
                            
                            # Chartex의 경우 txData 내의 v 요소도 확인 (cx:tx/cx:txData//cx:v)
                            if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                                tx_data_v_elem = tx_node.find(tx_data_v_path_in_tx, ns_map_for_xpath)
                                if tx_data_v_elem is not None and tx_data_v_elem.text and tx_data_v_elem.text.strip():
                                    original_text = tx_data_v_elem.text.strip()
                                    translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                    tx_data_v_elem.text = translated
                                    logger.debug(f"    시리즈 {ser_idx} 이름 (cx:txData//cx:v) 번역됨: '{original_text}' -> '{translated}'")

                    # 3. 카테고리 레이블(X축) 번역
                    cat_v_path_template = './/c:cat//c:strRef/c:strCache//c:pt/c:v'
                    if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                        # 확장 차트의 카테고리 레이블은 다른 구조일 수 있음. 예: cx:strDim[@type="cat"]//cx:pt (text child of cx:pt)
                        # 또는 cx:catAx//.../cx:tx//cx:v 등. 여기서는 일반 차트와 유사한 구조로 먼저 시도.
                        cat_v_path_template = './/cx:cat//cx:strRef/cx:strCache//cx:pt/cx:v' 
                        # 만약 위 경로가 안 맞으면, './/cx:strDim[@type="cat"]//cx:pt' 와 같이 cx:pt 요소의 text를 직접 수정해야 할 수 있음.
                        # 이 경우, cx:pt 요소의 text_content를 번역하고, 자식 a:t가 있다면 그것을, 없다면 cx:pt.text를 수정.
                        # 더 단순하게, _xml_extract.py 처럼 text() 노드를 찾는 방식을 여기서는 ElementTree 특성상 elem.text로 접근.

                    cat_v_elements = root.findall(cat_v_path_template, ns_map_for_xpath)
                    for cat_idx, cat_v_elem in enumerate(cat_v_elements):
                        if cat_v_elem.text and cat_v_elem.text.strip():
                            original_text = cat_v_elem.text.strip()
                            translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                            cat_v_elem.text = translated
                            logger.debug(f"    카테고리 레이블 {cat_idx} (c:v/cx:v) 번역됨: '{original_text}' -> '{translated}'")
                    
                    # 4. 축 제목 번역 (a:t 또는 c:v)
                    axis_title_paths_map = { # (is_v_element, path_list)
                        SCHEMA_CHART: [
                            (False, ['.//c:valAx//c:title//a:t', './/c:catAx//c:title//a:t', './/c:serAx//c:title//a:t', './/c:dateAx//c:title//a:t']),
                            (True, ['.//c:valAx//c:title//c:v', './/c:catAx//c:title//c:v'])
                        ],
                        SCHEMA_CHARTEX: [
                            (False, ['.//cx:axis//cx:title//a:t']), # cx는 주로 a:t를 사용
                            (True, ['.//cx:axis//cx:title//cx:v']) # cx의 v 요소도 확인 (드물 수 있음)
                        ]
                    }
                    paths_for_current_chart_type = axis_title_paths_map.get(current_chart_ns_uri_from_tag, axis_title_paths_map[SCHEMA_CHART])
                    
                    for is_v_elem, path_list in paths_for_current_chart_type:
                        for path_template in path_list:
                            axis_title_elements = root.findall(path_template, ns_map_for_xpath)
                            for title_elem in axis_title_elements:
                                if title_elem.text and title_elem.text.strip():
                                    original_text = title_elem.text.strip()
                                    translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                    title_elem.text = translated
                                    if is_v_elem: logger.debug(f"    축 제목 (c:v/cx:v) 번역됨 ('{path_template}'): '{original_text}' -> '{translated}'")
                                    else: logger.debug(f"    축 제목 (a:t) 번역됨 ('{path_template}'): '{original_text}' -> '{translated}'")

                    # 5. 데이터 레이블 번역 (주로 a:t, 때로는 c:v)
                    # 일반 차트: .//c:dLbls//a:t (또는 c:dLbl//c:tx//a:t, c:dLbl//c:tx//c:v)
                    # 확장 차트: .//cx:dataLabels//a:t (또는 cx:dataLabels//cx:tx//a:t, cx:txBody//a:p//a:r//a:t)
                    # 단순화된 접근: a:t를 먼저 찾고, 없으면 c:v를 찾는 방식도 고려 가능.
                    # 여기서는 _xml_extract.py의 접근과 유사하게 일반적인 경로 위주로 탐색.
                    
                    data_label_paths = [
                        './/c:dLbls//a:t', './/c:dLbl//c:tx//a:t' # 일반 차트 a:t
                    ]
                    if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                        data_label_paths = [
                            './/cx:dataLabels//a:t', './/cx:dataLabels//cx:tx//a:t', # 확장 차트 a:t
                            './/cx:dataLabels//cx:txBody//a:p//a:r//a:t' # 확장 차트 복잡한 구조
                        ]

                    for path_template in data_label_paths:
                        data_label_elements = root.findall(path_template, ns_map_for_xpath)
                        for lbl_idx, label_elem in enumerate(data_label_elements):
                            if label_elem.text and label_elem.text.strip():
                                original_text = label_elem.text.strip()
                                translated = self._translate_text_via_translator(original_text, src_lang_ui_name, tgt_lang_ui_name, model_name)
                                label_elem.text = translated
                                logger.debug(f"    데이터 레이블 {lbl_idx} (a:t, path: '{path_template}') 번역됨: '{original_text}' -> '{translated}'")
                    
                    # ElementTree는 tostring 시 XML 선언을 자동으로 추가하지 않으므로, 수동으로 추가
                    # 또한, ElementTree.tostring은 bytes를 반환하므로 decode 후 다시 encode 필요 없음.
                    # unicode 문자열로 얻은 후 utf-8로 인코딩.
                    xml_string_unicode = ET.tostring(root, encoding='unicode', method='xml')
                    
                    # 최종 바이트열: XML 선언부 + UTF-8 인코딩된 XML 문자열
                    final_xml_bytes = xml_declaration.strip() + xml_string_unicode.encode('utf-8').strip()
                    modified_charts_data[chart_xml_path_in_zip] = final_xml_bytes
                
                # 원본 ZIP 파일의 모든 항목을 새 ZIP 파일에 복사 (수정된 차트 XML은 교체)
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                    for item_name_in_zip in zip_ref.namelist():
                        if item_name_in_zip in modified_charts_data:
                            zip_out.writestr(item_name_in_zip, modified_charts_data[item_name_in_zip])
                        else:
                            zip_out.writestr(item_name_in_zip, zip_ref.read(item_name_in_zip))
            
            logger.info(f"\nPPTX 내 차트 XML 번역 완료! 최종 파일 저장됨: {output_path}")
            return output_path
        
        except FileNotFoundError:
            logger.error(f"오류: 원본 PPTX 파일 '{pptx_path}'를 찾을 수 없습니다.")
            return None
        except zipfile.BadZipFile:
            logger.error(f"오류: 파일 '{pptx_path}'는 유효한 ZIP 파일(PPTX)이 아닙니다.")
            return None
        except ET.ParseError as e_xml_parse:
            logger.error(f"오류: 차트 XML 파싱 중 오류 발생: {e_xml_parse}")
            # 이 경우, 부분적으로 처리되었을 수 있으므로 temp_dir 정리가 필요할 수 있음
            return None
        except Exception as e:
            logger.error(f"PPTX 내 차트 XML 번역 중 예기치 않은 오류 발생: {e}", exc_info=True)
            return None
        finally:
            if os.path.exists(temp_dir_for_xml_processing):
                try:
                    shutil.rmtree(temp_dir_for_xml_processing)
                    logger.debug(f"임시 디렉토리 '{temp_dir_for_xml_processing}' 삭제 완료.")
                except Exception as e_clean:
                    logger.warning(f"임시 디렉토리 '{temp_dir_for_xml_processing}' 삭제 중 오류: {e_clean}")