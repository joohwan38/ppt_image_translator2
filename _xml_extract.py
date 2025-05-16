import zipfile
import xml.etree.ElementTree as ET
import os
import re
import requests
import tempfile
import shutil
from pathlib import Path

def translate_text(text, target_lang="en"):
    """Ollama API를 사용하여 텍스트 번역"""
    if not text or not text.strip():
        return text
    
    url = "http://localhost:11434/api/generate"
    prompt = f"Translate the following text to {target_lang}. Return only the translation without explanations: \"{text}\""
    data = {
        "model": "gemma3:12b", # 모델명은 필요에 따라 변경하세요.
        "prompt": prompt,
        "stream": False
    }
    
    try:
        response = requests.post(url, json=data)
        response.raise_for_status()
        result = response.json()
        translated_text = result.get("response", "").strip()
        if translated_text.startswith('"') and translated_text.endswith('"'):
            translated_text = translated_text[1:-1]
        print(f"  - 번역: '{text}' -> '{translated_text}'")
        return translated_text
    except Exception as e:
        print(f"  - 번역 오류: {e}")
        return text

def translate_pptx_charts(pptx_path, target_lang="en", output_path=None):
    """PowerPoint 파일의 차트 텍스트를 추출하고 번역한 후 다시 삽입"""
    if output_path is None:
        base_name = os.path.splitext(pptx_path)[0]
        output_path = f"{base_name}_translated.pptx"
    
    print(f"PowerPoint 차트 번역 시작: {os.path.basename(pptx_path)} -> {os.path.basename(output_path)}")
    print(f"대상 언어: {target_lang}")
    
    temp_dir = tempfile.mkdtemp()

    # 네임스페이스 URI 정의
    SCHEMA_MAIN = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    SCHEMA_CHART = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    SCHEMA_CHARTEX = 'http://schemas.microsoft.com/office/drawing/2014/chartex' # 확장 차트

    # ElementTree에 네임스페이스 등록 (XML 생성 시 원하는 접두사 사용 유도)
    # 애플리케이션 실행 시 한 번만 등록해도 되지만, 여기서는 함수 호출 시마다 등록
    namespaces_to_register = {
        'c': SCHEMA_CHART,
        'a': SCHEMA_MAIN,
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        'c14': 'http://schemas.microsoft.com/office/drawing/2007/8/2/chart',
        'c16': 'http://schemas.microsoft.com/office/drawing/2014/chart',
        'c16r2': 'http://schemas.microsoft.com/office/drawing/2015/06/chart',
        'c16r3': 'http://schemas.microsoft.com/office/drawing/2017/03/chart',
        'cx': SCHEMA_CHARTEX 
    }
    for prefix, uri in namespaces_to_register.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError as e:
            if "prefix already registered" not in str(e).lower():
                 print(f"Warning: Could not register namespace {prefix}={uri}. Error: {e}")
    
    # findall을 위한 네임스페이스 맵 (ElementTree XPath 용)
    # 이 맵은 element.findall(path, ns_map_for_xpath) 형태로 사용됩니다.
    # XPath에서는 여기서 정의한 접두사를 사용합니다. 예: 'c:title'
    ns_map_for_xpath = {
        'a': SCHEMA_MAIN,
        'c': SCHEMA_CHART,
        'cx': SCHEMA_CHARTEX
    }


    try:
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            chart_files = [f for f in zip_ref.namelist() if f.startswith('ppt/charts/') and f.endswith('.xml')]
            print(f"총 {len(chart_files)}개의 차트 파일을 발견했습니다.")
            
            modified_charts_data = {} 

            for chart_path in chart_files:
                print(f"\n처리 중: {chart_path}")
                
                with zip_ref.open(chart_path) as f:
                    content = f.read()
                    xml_declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    
                    content_str = content.decode('utf-8')
                    if content_str.startswith('<?xml'):
                        content_str = content_str.split('?>', 1)[-1].strip()
                    
                    root = ET.fromstring(content_str)

                # 현재 XML의 루트 네임스페이스가 일반 차트인지 확장 차트인지 판단
                # root.tag는 "{URI}localname" 형태
                current_chart_ns_uri_from_tag = ""
                if root.tag.startswith('{') and '}' in root.tag:
                    current_chart_ns_uri_from_tag = root.tag.split('}', 1)[0][1:]

                # XPath 검색 시 사용할 네임스페이스 맵 결정
                # (기본적으로 ns_map_for_xpath를 사용하고, findall 호출 시 이 맵을 전달)
                
                # 1. 차트 제목 번역
                # 일반 차트: .//c:title//a:t
                # 확장 차트: .//cx:title//a:t (또는 cx:tx//a:t 등 구조에 따라 다름)
                # 주의: title_elements = root.findall('.//c:title//a:t', ns_map_for_xpath)
                # 이렇게 ns_map_for_xpath를 전달해야 합니다.
                
                title_path_template = './/c:title//a:t'
                if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                    title_path_template = './/cx:title//a:t' # cx 네임스페이스 사용

                title_elements = root.findall(title_path_template, ns_map_for_xpath)
                for elem in title_elements:
                    if elem.text and elem.text.strip():
                        elem.text = translate_text(elem.text.strip(), target_lang)
                
                # 참조 텍스트로 저장된 제목
                title_v_path_template = './/c:chart/c:title//c:v'
                if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                     title_v_path_template = './/cx:chart/cx:title//cx:v' # 또는 cx는 다른 구조일 수 있음

                title_v_elements = root.findall(title_v_path_template, ns_map_for_xpath)
                for v_elem in title_v_elements:
                    if v_elem.text and v_elem.text.strip():
                        v_elem.text = translate_text(v_elem.text.strip(), target_lang)
                
                # 2. 시리즈 이름(범례) 번역
                ser_path_template = './/c:ser'
                tx_path_in_ser = './/c:tx'
                v_path_in_tx = './/c:v'
                str_cache_path_in_tx = './/c:strRef//c:strCache//c:pt//c:v'
                tx_data_path_in_tx = './/cx:txData' # cx용

                if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                    ser_path_template = './/cx:series' # 또는 cx:ser
                    tx_path_in_ser = './/cx:tx'
                    v_path_in_tx = './/cx:v'
                    str_cache_path_in_tx = './/cx:strRef//cx:strCache//cx:pt//cx:v' # cx도 유사 구조 가능

                series_elements = root.findall(ser_path_template, ns_map_for_xpath)
                for ser in series_elements:
                    tx_node = ser.find(tx_path_in_ser, ns_map_for_xpath)
                    if tx_node is not None:
                        v_direct = tx_node.find(v_path_in_tx, ns_map_for_xpath)
                        if v_direct is not None and v_direct.text and v_direct.text.strip():
                            v_direct.text = translate_text(v_direct.text.strip(), target_lang)
                        
                        for v_cache in tx_node.findall(str_cache_path_in_tx, ns_map_for_xpath):
                            if v_cache.text and v_cache.text.strip():
                                v_cache.text = translate_text(v_cache.text.strip(), target_lang)
                        
                        if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX: # chartex의 txData/v
                            tx_data_node = tx_node.find(tx_data_path_in_tx, ns_map_for_xpath)
                            if tx_data_node is not None:
                                v_tx_data = tx_data_node.find(v_path_in_tx, ns_map_for_xpath) # v_path_in_tx 재활용
                                if v_tx_data is not None and v_tx_data.text and v_tx_data.text.strip():
                                    v_tx_data.text = translate_text(v_tx_data.text.strip(), target_lang)

                # 3. 카테고리 레이블(X축) 번역
                cat_v_path_template = './/c:cat//c:strRef//c:strCache//c:pt//c:v'
                if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                    cat_v_path_template = './/cx:strDim[@type="cat"]//cx:pt' # cx:pt의 텍스트

                cat_v_elements = root.findall(cat_v_path_template, ns_map_for_xpath)
                for cat_v in cat_v_elements:
                    if cat_v.text and cat_v.text.strip():
                        cat_v.text = translate_text(cat_v.text.strip(), target_lang)
                
                # 4. 축 제목 번역
                axis_title_t_paths = []
                axis_title_v_paths = []
                if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                    axis_title_t_paths = ['.//cx:axis//cx:title//a:t']
                else: # 일반 차트
                    axis_title_t_paths = [
                        './/c:valAx//c:title//a:t',
                        './/c:catAx//c:title//a:t',
                        './/c:serAx//c:title//a:t',
                        './/c:dateAx//c:title//a:t'
                    ]
                    axis_title_v_paths = [ # c:v 형태의 축 제목 (일반 차트)
                        './/c:valAx//c:title//c:v',
                        './/c:catAx//c:title//c:v'
                    ]
                
                all_axis_titles_t = []
                for path_t in axis_title_t_paths:
                    all_axis_titles_t.extend(root.findall(path_t, ns_map_for_xpath))
                
                for t_elem in all_axis_titles_t:
                    if t_elem.text and t_elem.text.strip():
                        t_elem.text = translate_text(t_elem.text.strip(), target_lang)

                all_axis_titles_v = []
                for path_v in axis_title_v_paths: # 일반 차트의 경우에만 v 검색
                    all_axis_titles_v.extend(root.findall(path_v, ns_map_for_xpath))

                for v_elem in all_axis_titles_v:
                    if v_elem.text and v_elem.text.strip():
                        v_elem.text = translate_text(v_elem.text.strip(), target_lang)
                
                # 5. 데이터 레이블 번역
                data_label_path_template = './/c:dLbls//a:t'
                if current_chart_ns_uri_from_tag == SCHEMA_CHARTEX:
                    # 확장 차트의 데이터 레이블은 구조가 더 복잡할 수 있습니다.
                    # 예: .//cx:dataLabels//cx:txBody//a:p//a:r//a:t 또는 .//cx:dataLabels//cx:tx//a:t
                    # 여기서는 단순화된 예시를 사용하며, 실제 구조에 맞게 조정 필요.
                    data_label_path_template = './/cx:dataLabels//a:t' # 이 경로는 예시이며, 실제 cx 파일 구조 확인 필요

                data_label_elements = root.findall(data_label_path_template, ns_map_for_xpath)
                for label_elem in data_label_elements:
                    if label_elem.text and label_elem.text.strip():
                        label_elem.text = translate_text(label_elem.text.strip(), target_lang)
                
                xml_string_unicode = ET.tostring(root, encoding='unicode', method='xml')
                
                final_xml_bytes = xml_declaration.strip() + xml_string_unicode.encode('utf-8').strip()
                
                modified_charts_data[chart_path] = final_xml_bytes
            
            all_files_in_zip = zip_ref.namelist()
            
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for item_name in all_files_in_zip:
                    if item_name in modified_charts_data:
                        zip_out.writestr(item_name, modified_charts_data[item_name])
                    else:
                        zip_out.writestr(item_name, zip_ref.read(item_name))
        
        print(f"\n번역 완료! 파일 저장됨: {output_path}")
    
    finally:
        shutil.rmtree(temp_dir)

if __name__ == "__main__":
    input_file = 'files/xml_test.pptx' 
    target_language = 'ja' 
    
    if not os.path.exists(input_file):
        print(f"입력 파일 '{input_file}'을 찾을 수 없습니다. 경로를 확인해주세요.")
    else:
        translate_pptx_charts(input_file, target_language)