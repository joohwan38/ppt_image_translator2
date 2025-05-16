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
    
    # Ollama API 엔드포인트
    url = "http://localhost:11434/api/generate"
    
    # 번역 프롬프트 작성
    prompt = f"Translate the following text to {target_lang}. Return only the translation without explanations: \"{text}\""
    
    # API 요청 데이터
    data = {
        "model": "gemma3:12b",
        "prompt": prompt,
        "stream": False
    }
    
    try:
        # API 호출
        response = requests.post(url, json=data)
        response.raise_for_status()
        result = response.json()
        
        # 결과에서 번역된 텍스트 추출
        translated_text = result.get("response", "").strip()
        
        # 따옴표 제거 (API 응답에 따옴표가 포함될 수 있음)
        if translated_text.startswith('"') and translated_text.endswith('"'):
            translated_text = translated_text[1:-1]
        
        print(f"  - 번역: '{text}' -> '{translated_text}'")
        return translated_text
    except Exception as e:
        print(f"  - 번역 오류: {e}")
        return text  # 오류 시 원본 반환

def translate_pptx_charts(pptx_path, target_lang="en", output_path=None):
    """PowerPoint 파일의 차트 텍스트를 추출하고 번역한 후 다시 삽입"""
    if output_path is None:
        # 출력 파일 경로 생성 (원본_translated.pptx)
        base_name = os.path.splitext(pptx_path)[0]
        output_path = f"{base_name}_translated.pptx"
    
    print(f"PowerPoint 차트 번역 시작: {os.path.basename(pptx_path)} -> {os.path.basename(output_path)}")
    print(f"대상 언어: {target_lang}")
    
    # 임시 디렉토리 생성
    temp_dir = tempfile.mkdtemp()
    chart_dir = os.path.join(temp_dir, "charts")
    os.makedirs(chart_dir, exist_ok=True)
    
    try:
        # 차트 파일 분석 및 번역
        with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
            # 네임스페이스 정의
            namespaces = {
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
            }
            
            # 차트 XML 파일 목록 찾기
            chart_files = [f for f in zip_ref.namelist() if f.startswith('ppt/charts/') and f.endswith('.xml')]
            print(f"총 {len(chart_files)}개의 차트 파일을 발견했습니다.")
            
            # 차트 파일 저장 - 수정할 파일 목록
            modified_charts = set()
            
            # 각 차트 파일 처리
            for chart_path in chart_files:
                print(f"\n처리 중: {chart_path}")
                
                # 차트 XML 파일 내용 읽기
                with zip_ref.open(chart_path) as f:
                    content = f.read()
                    # XML 파싱 전에 선언 저장
                    xml_declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    if content.startswith(b'<?xml'):
                        xml_end = content.find(b'?>')
                        if xml_end > 0:
                            xml_declaration = content[:xml_end+2]
                    
                    tree = ET.ElementTree(ET.fromstring(content))
                    root = tree.getroot()
                
                # 1. 차트 제목 번역
                title_elements = root.findall('.//c:title//a:t', namespaces)
                for elem in title_elements:
                    if elem.text and elem.text.strip():
                        elem.text = translate_text(elem.text.strip(), target_lang)
                
                # 참조 텍스트로 저장된 제목
                title_elem = root.find('.//c:chart/c:title', namespaces)
                if title_elem is not None:
                    for v in title_elem.findall('.//c:v', namespaces):
                        if v.text and v.text.strip():
                            v.text = translate_text(v.text.strip(), target_lang)
                
                # 2. 시리즈 이름(범례) 번역
                for ser in root.findall('.//c:ser', namespaces):
                    tx = ser.find('.//c:tx', namespaces)
                    if tx is not None:
                        # 직접 값
                        v = tx.find('.//c:v', namespaces)
                        if v is not None and v.text:
                            v.text = translate_text(v.text, target_lang)
                        
                        # 참조 값
                        for v in tx.findall('.//c:strRef//c:strCache//c:pt//c:v', namespaces):
                            if v.text:
                                v.text = translate_text(v.text, target_lang)
                
                # 3. 카테고리 레이블(X축) 번역
                for cat in root.findall('.//c:cat//c:strRef//c:strCache//c:pt//c:v', namespaces):
                    if cat.text:
                        cat.text = translate_text(cat.text, target_lang)
                
                # 4. 축 제목 번역
                for axis in root.findall('.//c:axis', namespaces):
                    title = axis.find('.//c:title', namespaces)
                    if title is not None:
                        for t in title.findall('.//a:t', namespaces):
                            if t.text:
                                t.text = translate_text(t.text, target_lang)
                        
                        for v in title.findall('.//c:v', namespaces):
                            if v.text:
                                v.text = translate_text(v.text, target_lang)
                
                # 5. 데이터 레이블 번역
                for label in root.findall('.//c:dLbls//a:t', namespaces):
                    if label.text:
                        label.text = translate_text(label.text, target_lang)
                
                # 수정된 XML 문자열 생성 (원본 XML 선언 유지)
                xml_string = ET.tostring(root, encoding='UTF-8')
                
                # 원본 선언 추가
                final_xml = xml_declaration + b'\n' + xml_string
                
                # 수정된 차트 파일 저장
                temp_file = os.path.join(chart_dir, os.path.basename(chart_path))
                with open(temp_file, 'wb') as f:
                    f.write(final_xml)
                
                # 수정된 파일 목록에 추가
                modified_charts.add(chart_path)
            
            # 모든 파일 목록 (수정된 차트 파일 제외)
            all_files = [f for f in zip_ref.namelist() if f not in modified_charts]
            
            # 새 PPTX 파일 생성
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                # 1. 수정되지 않은 파일 복사
                for file_path in all_files:
                    zip_out.writestr(file_path, zip_ref.read(file_path))
                
                # 2. 수정된 차트 파일 추가
                for chart_path in modified_charts:
                    temp_file = os.path.join(chart_dir, os.path.basename(chart_path))
                    with open(temp_file, 'rb') as f:
                        zip_out.writestr(chart_path, f.read())
        
        print(f"\n번역 완료! 파일 저장됨: {output_path}")
    
    finally:
        # 임시 디렉토리 정리
        shutil.rmtree(temp_dir)

# 예시 사용법
if __name__ == "__main__":
    # 예시: 파워포인트 파일의 차트를 한국어로 번역
    input_file = 'files/xml_test.pptx'
    target_language = 'ja'
    translate_pptx_charts(input_file, target_language)