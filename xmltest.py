import zipfile
import xml.etree.ElementTree as ET
import os
import re

def extract_chart_text(pptx_path):
    print(f"Analyzing file: {os.path.basename(pptx_path)}")
    
    # PPTX 파일을 ZIP으로 열기
    with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
        # 필요한 네임스페이스 정의
        namespaces = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
        }
        
        # 슬라이드 목록 찾기
        slides = sorted([f for f in zip_ref.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')])
        slide_count = 0
        chart_count = 0
        
        # 각 슬라이드 분석
        for slide_path in slides:
            slide_count += 1
            slide_num = re.search(r'slide(\d+)\.xml', slide_path).group(1)
            print(f"\nSlide {slide_num}:")
            
            # 슬라이드 파일 열기
            with zip_ref.open(slide_path) as f:
                slide_tree = ET.parse(f)
                slide_root = slide_tree.getroot()
                
                # 슬라이드의 관계 파일 경로
                slide_rel_path = f'ppt/slides/_rels/slide{slide_num}.xml.rels'
                
                if slide_rel_path in zip_ref.namelist():
                    # 관계 파일 열기
                    with zip_ref.open(slide_rel_path) as rel_file:
                        rel_tree = ET.parse(rel_file)
                        rel_root = rel_tree.getroot()
                        
                        # 차트 관계 찾기
                        chart_rels = {}
                        for rel in rel_root.findall('.//{*}Relationship'):
                            if 'Type' in rel.attrib and 'chart' in rel.attrib['Type']:
                                chart_rels[rel.attrib['Id']] = rel.attrib['Target']
                        
                        # 슬라이드에서 차트 찾기
                        chart_refs = slide_root.findall('.//p:graphicFrame//c:chart', namespaces)
                        
                        if not chart_refs:
                            print("  No charts in this slide.")
                            continue
                        
                        for chart_ref in chart_refs:
                            chart_count += 1
                            
                            # 차트 ID 가져오기 (r:id 속성)
                            chart_rid = None
                            for attrib_name in chart_ref.attrib:
                                if attrib_name.endswith('id'):
                                    chart_rid = chart_ref.attrib[attrib_name]
                                    break
                            
                            if chart_rid and chart_rid in chart_rels:
                                # 차트 파일 경로 생성
                                chart_target = chart_rels[chart_rid]
                                
                                # 상대 경로를 절대 경로로 변환
                                if chart_target.startswith('../'):
                                    chart_path = 'ppt/' + chart_target[3:]
                                else:
                                    chart_path = os.path.dirname(slide_path) + '/' + chart_target
                                
                                print(f"  Chart {chart_count} found! (Path: {chart_path})")
                                
                                # 차트 파일이 존재하는지 확인
                                if chart_path in zip_ref.namelist():
                                    # 차트 XML 파일 분석
                                    with zip_ref.open(chart_path) as chart_file:
                                        chart_tree = ET.parse(chart_file)
                                        chart_root = chart_tree.getroot()
                                        
                                        # 1. 차트 제목 (다양한 방식으로 시도)
                                        title_text = None
                                        
                                        # 표준 차트 제목
                                        title_elements = chart_root.findall('.//c:title//a:t', namespaces)
                                        if title_elements:
                                            for title in title_elements:
                                                if title.text:
                                                    title_text = title.text
                                                    break
                                        
                                        # 대체 방법: 텍스트 상자로 된 제목
                                        if not title_text:
                                            alt_title = chart_root.find('.//c:chart//c:title', namespaces)
                                            if alt_title is not None:
                                                tx_elements = alt_title.findall('.//a:t', namespaces)
                                                for tx in tx_elements:
                                                    if tx.text:
                                                        title_text = tx.text
                                                        break
                                        
                                        # 차트 요소 외부에 있는 제목 (슬라이드 내 텍스트)
                                        if not title_text:
                                            chart_parent = None
                                            for elem in slide_root.findall('.//p:graphicFrame', namespaces):
                                                for c in elem.findall('.//c:chart', namespaces):
                                                    for attr in c.attrib:
                                                        if attr.endswith('id') and c.attrib[attr] == chart_rid:
                                                            chart_parent = elem
                                                            break
                                            
                                            if chart_parent:
                                                # 차트 위에 있는 텍스트 상자 찾기
                                                for shape in slide_root.findall('.//p:sp', namespaces):
                                                    txBody = shape.find('.//p:txBody', namespaces)
                                                    if txBody:
                                                        text_elems = txBody.findall('.//a:t', namespaces)
                                                        for t in text_elems:
                                                            if t.text and len(t.text.strip()) > 0:
                                                                title_text = t.text
                                                                break
                                        
                                        # 가능한 모든 방법으로 제목을 찾았으면 출력
                                        if title_text:
                                            print(f"  • Chart Title: {title_text}")
                                        
                                        # 2. 범례(시리즈 이름) - 개선된 방법
                                        # 시리즈 레이블을 저장할 집합(중복 방지)
                                        series_names = set()
                                        
                                        # 방법 1: 시리즈 텍스트에서 직접 찾기
                                        for ser in chart_root.findall('.//c:ser', namespaces):
                                            tx = ser.find('.//c:tx', namespaces)
                                            if tx is not None:
                                                # 방법 1-1: 직접 텍스트
                                                v = tx.find('.//c:v', namespaces)
                                                if v is not None and v.text:
                                                    series_names.add(v.text)
                                                    continue
                                                
                                                # 방법 1-2: 참조 문자열
                                                strRef = tx.find('.//c:strRef', namespaces)
                                                if strRef is not None:
                                                    pt = strRef.find('.//c:pt//c:v', namespaces)
                                                    if pt is not None and pt.text:
                                                        series_names.add(pt.text)
                                                        continue
                                                
                                                # 방법 1-3: 서식 있는 텍스트
                                                rich = tx.find('.//a:rich', namespaces)
                                                if rich is not None:
                                                    t = rich.find('.//a:t', namespaces)
                                                    if t is not None and t.text:
                                                        series_names.add(t.text)
                                        
                                        # 방법 2: 범례 요소에서 찾기
                                        legend = chart_root.find('.//c:legend', namespaces)
                                        if legend is not None:
                                            txPrs = legend.findall('.//a:t', namespaces)
                                            for t in txPrs:
                                                if t.text:
                                                    series_names.add(t.text)
                                        
                                        # 결과 출력
                                        if series_names:
                                            print("  • Series Names (Legend):")
                                            for i, name in enumerate(series_names, 1):
                                                print(f"    - Series {i}: {name}")
                                        
                                        # 3. 축 제목 (더 다양한 방법으로 시도)
                                        axes = chart_root.findall('.//c:axis', namespaces)
                                        for axis in axes:
                                            axis_id = axis.find('.//c:axId', namespaces)
                                            axis_type = ""
                                            if axis_id is not None:
                                                axis_type = "X-axis" if axis.get('{http://schemas.openxmlformats.org/drawingml/2006/chart}axId') == "1" else "Y-axis"
                                            
                                            # 축 제목 요소 찾기
                                            title = axis.find('.//c:title', namespaces)
                                            if title is not None:
                                                # 다양한 방법으로 텍스트 추출 시도
                                                axis_title_text = None
                                                
                                                # 방법 1: 직접 텍스트
                                                for t in title.findall('.//a:t', namespaces):
                                                    if t.text:
                                                        axis_title_text = t.text
                                                        break
                                                
                                                # 방법 2: 참조 문자열
                                                if not axis_title_text:
                                                    strRef = title.find('.//c:strRef//c:strCache//c:pt//c:v', namespaces)
                                                    if strRef is not None and strRef.text:
                                                        axis_title_text = strRef.text
                                                
                                                if axis_title_text:
                                                    print(f"  • {axis_type} Title: {axis_title_text}")
                                        
                                        # 4. 카테고리 레이블 (중복 제거)
                                        # X축 카테고리를 저장할 리스트 (순서 유지)
                                        categories = []
                                        category_set = set()  # 중복 확인용
                                        
                                        # 방법 1: 일반적인 카테고리 (문자열)
                                        for cat in chart_root.findall('.//c:cat//c:strRef//c:strCache//c:pt', namespaces):
                                            idx = cat.get('idx')
                                            v = cat.find('.//c:v', namespaces)
                                            if v is not None and v.text and v.text not in category_set:
                                                categories.append((int(idx) if idx else len(categories), v.text))
                                                category_set.add(v.text)
                                        
                                        # 방법 2: 숫자 카테고리
                                        if not categories:
                                            for cat in chart_root.findall('.//c:cat//c:numRef//c:numCache//c:pt', namespaces):
                                                idx = cat.get('idx')
                                                v = cat.find('.//c:v', namespaces)
                                                if v is not None and v.text and v.text not in category_set:
                                                    categories.append((int(idx) if idx else len(categories), v.text))
                                                    category_set.add(v.text)
                                        
                                        # 순서대로 정렬하고 출력
                                        if categories:
                                            categories.sort(key=lambda x: x[0])
                                            print("  • Category Labels (X-axis):")
                                            for i, (_, cat) in enumerate(categories, 1):
                                                print(f"    - Category {i}: {cat}")
                                        
                                        # 5. 데이터 레이블
                                        data_labels = set()
                                        for label in chart_root.findall('.//c:dLbls//a:t', namespaces):
                                            if label.text:
                                                data_labels.add(label.text)
                                        
                                        if data_labels:
                                            print("  • Data Labels:")
                                            for i, label in enumerate(data_labels, 1):
                                                print(f"    - Label {i}: {label}")
                                else:
                                    print(f"  Error: Chart file not found ({chart_path})")
                                
                                print("-" * 50)
                else:
                    print("  No relationships file found for this slide.")
        
        if chart_count == 0:
            print("\nNo charts found in the presentation.")
        else:
            print(f"\nTotal: {chart_count} charts found in {slide_count} slides.")

# 요청하신 대로 수정된 부분
if __name__ == "__main__":
    extract_chart_text('files/번역테스트1.pptx')