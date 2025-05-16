from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.lang import MSO_LANGUAGE_ID # 명시적 임포트

import os
import io
import logging
import re
from datetime import datetime
import hashlib
import traceback
from PIL import Image # OCR 이미지 처리를 위해 추가

logger = logging.getLogger(__name__)

# (should_skip_translation, is_ocr_text_valid, get_file_info, _get_text_style, _apply_text_style 함수는 이전 답변과 동일하게 유지)
MIN_MEANINGFUL_CHAR_RATIO_SKIP = 0.1
MIN_MEANINGFUL_CHAR_RATIO_OCR = 0.1
MEANINGFUL_CHAR_PATTERN = re.compile(
    r'[a-zA-Z'                 # 영어
    r'\u00C0-\u024F'           # 베트남어 포함한 Latin Extended-A, B 일부
    r'\u1E00-\u1EFF'           # Latin Extended Additional (베트남어 악센트 등)
    r'\u0600-\u06FF'           # 아랍어 기본
    r'\u0750-\u077F'           # 아랍어 보충
    r'\u08A0-\u08FF'           # 아랍어 확장-A
    r'\u3040-\u30ff'           # 일본어
    r'\u3131-\uD79D'           # 한글
    r'\u4e00-\u9fff'           # 한자
    r'\u0E00-\u0E7F'           # 태국어
    r']'
)



def should_skip_translation(text: str) -> bool:
    if not text: return True
    stripped_text = text.strip()
    if not stripped_text:
        logger.debug(f"번역 스킵 (공백만 존재): '{text[:50]}...'")
        return True
    if not MEANINGFUL_CHAR_PATTERN.search(stripped_text):
        logger.debug(f"번역 스킵 (의미 있는 문자 없음): '{stripped_text[:50]}...'")
        return True
    text_len = len(stripped_text)
    if text_len <= 3:
        logger.debug(f"번역 시도 (짧은 문자열, 의미 문자 포함): '{stripped_text}'")
        return False
    meaningful_chars_count = len(MEANINGFUL_CHAR_PATTERN.findall(stripped_text))
    if (meaningful_chars_count / text_len) < MIN_MEANINGFUL_CHAR_RATIO_SKIP:
        logger.debug(f"번역 스킵 (의미 문자 비율 낮음 {meaningful_chars_count / text_len:.2f}, 임계값: {MIN_MEANINGFUL_CHAR_RATIO_SKIP}): '{stripped_text[:50]}...'")
        return True
    logger.debug(f"번역 시도 (조건 통과): '{stripped_text[:50]}...'")
    return False

def is_ocr_text_valid(text: str) -> bool:
    if not text: return False
    stripped_text = text.strip()
    if not stripped_text: return False
    if not MEANINGFUL_CHAR_PATTERN.search(stripped_text):
        logger.debug(f"OCR 유효성 스킵 (의미 있는 문자 없음): '{stripped_text[:50]}...'")
        return False
    text_len = len(stripped_text)
    if text_len <= 2:
        logger.debug(f"OCR 유효 (매우 짧은 문자열, 의미 문자 포함): '{stripped_text}'")
        return True
    meaningful_chars_count = len(MEANINGFUL_CHAR_PATTERN.findall(stripped_text))
    if (meaningful_chars_count / text_len) < MIN_MEANINGFUL_CHAR_RATIO_OCR:
        logger.debug(f"OCR 유효성 스킵 (의미 문자 비율 낮음 {meaningful_chars_count / text_len:.2f}, 임계값: {MIN_MEANINGFUL_CHAR_RATIO_OCR}): '{stripped_text[:50]}...'")
        return False
    logger.debug(f"OCR 유효 (조건 통과): '{stripped_text[:50]}...'")
    return True


class PptxHandler:
    def __init__(self):
        pass

    def get_file_info(self, file_path):
        logger.info(f"파일 정보 분석 시작 (OCR 미수행): {file_path}")
        info = {"slide_count": 0, "text_elements": 0, "image_elements": 0}
        try:
            prs = Presentation(file_path)
            info["slide_count"] = len(prs.slides)
            for slide_idx, slide in enumerate(prs.slides):
                for shape_idx, shape in enumerate(slide.shapes):
                    if shape.has_text_frame and shape.text_frame.text and shape.text_frame.text.strip():
                        info["text_elements"] += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        info["image_elements"] += 1
                    elif shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if cell.text_frame.text and cell.text_frame.text.strip():
                                    info["text_elements"] += 1
            logger.info(f"파일 분석 완료 (OCR 미수행): Slides:{info['slide_count']}, Text:{info['text_elements']}, Img:{info['image_elements']}")
        except Exception as e:
            logger.error(f"'{os.path.basename(file_path)}' 파일 정보 분석 오류: {e}", exc_info=True)
        return info

    def _get_style_properties(self, font_object):
        if font_object is None:
            return {}

        style_props = {
            'name': font_object.name,
            'size': font_object.size,
            'bold': font_object.bold,
            'italic': font_object.italic,
            'underline': font_object.underline,
            'color_rgb': None,
            'color_theme_index': None,
            'color_brightness': None,
            'language_id': None
        }
        
        try:
            # font_object.language_id가 MSO_LANGUAGE_ID 멤버를 반환하거나 ValueError를 발생시킬 수 있음
            lang_id = font_object.language_id
            style_props['language_id'] = lang_id
        except ValueError as e:
            # raw_lang_xml = getattr(font_object._element.get_or_add_rPr(), 'lang', "N/A") if hasattr(font_object, '_element') else "N/A"
            # logger.warning(f"Unsupported language_id value from XML ('{raw_lang_xml}') in _get_style_properties. Defaulting to None. Error: {e}")
            logger.warning(f"Unsupported language_id value in _get_style_properties. Defaulting to None. Error: {e}")
            style_props['language_id'] = None # 오류 발생 시 None으로 설정

        if hasattr(font_object.color, 'type') and font_object.color.type is not None:
            color_type = font_object.color.type
            if color_type == MSO_COLOR_TYPE.RGB:
                # MSO_COLOR_TYPE.RGB인 경우 .rgb 속성이 RGBColor 객체이거나 None일 수 있음
                if font_object.color.rgb is not None:
                    style_props['color_rgb'] = tuple(font_object.color.rgb)
            elif color_type == MSO_COLOR_TYPE.SCHEME:
                # MSO_COLOR_TYPE.SCHEME인 경우, .rgb를 직접 접근하는 대신
                # 테마 색상 정보(theme_color 인덱스, brightness)를 저장합니다.
                # 실제 RGB 값은 프레젠테이션 테마에 따라 동적으로 결정되므로,
                # 적용 시점에 이 정보를 사용하는 것이 더 정확할 수 있습니다.
                # 하지만, python-pptx는 때때로 SCHEME color에 대해서도 .rgb로
                # 현재 계산된 값을 제공하려고 시도합니다. 오류가 발생했으므로 직접 접근을 피합니다.
                style_props['color_theme_index'] = font_object.color.theme_color
                if hasattr(font_object.color, 'brightness') and font_object.color.brightness is not None:
                    style_props['color_brightness'] = font_object.color.brightness
                else:
                    style_props['color_brightness'] = 0.0 # 기본 밝기
                
                # 만약 _SchemeColor 객체에서 안전하게 현재 RGB 값을 얻는 방법이 있다면 여기에 추가할 수 있습니다.
                # 예를 들어, try-except로 .rgb를 접근해 볼 수 있지만, 오류가 이미 발생했으므로
                # 일단은 테마 정보만 저장하는 것이 안전합니다.
                # logger.debug(f"Scheme color: Theme index {style_props['color_theme_index']}, Brightness {style_props['color_brightness']}")
        return style_props

    def _get_text_style(self, run): # run은 pptx.text.text._Run 객체
        style = self._get_style_properties(run.font)
        # _Run 객체에는 hyperlink 속성이 직접 있음
        style['hyperlink_address'] = run.hyperlink.address if run.hyperlink and run.hyperlink.address else None
        return style

    def _apply_style_properties(self, target_font_object, style_dict_to_apply):
        if not style_dict_to_apply or target_font_object is None:
            return

        font = target_font_object

        if 'name' in style_dict_to_apply and style_dict_to_apply['name'] is not None:
            font.name = style_dict_to_apply['name']
        
        if 'size' in style_dict_to_apply and style_dict_to_apply['size'] is not None:
            font.size = style_dict_to_apply['size']

        if 'bold' in style_dict_to_apply and style_dict_to_apply['bold'] is not None:
            font.bold = style_dict_to_apply['bold']
        if 'italic' in style_dict_to_apply and style_dict_to_apply['italic'] is not None:
            font.italic = style_dict_to_apply['italic']
        if 'underline' in style_dict_to_apply and style_dict_to_apply['underline'] is not None:
            font.underline = style_dict_to_apply['underline']

        applied_color = False
        if 'color_rgb' in style_dict_to_apply and style_dict_to_apply['color_rgb'] is not None:
            try:
                font.color.rgb = RGBColor(*(int(c) for c in style_dict_to_apply['color_rgb']))
                applied_color = True
            except Exception as e:
                logger.warning(f"RGB 색상 {style_dict_to_apply['color_rgb']} 적용 실패: {e}")
        
        if not applied_color and 'color_theme_index' in style_dict_to_apply and style_dict_to_apply['color_theme_index'] is not None:
            try:
                font.color.theme_color = style_dict_to_apply['color_theme_index']
                if 'color_brightness' in style_dict_to_apply and style_dict_to_apply['color_brightness'] is not None:
                    brightness_val = float(style_dict_to_apply['color_brightness'])
                    font.color.brightness = max(-1.0, min(1.0, brightness_val))
                applied_color = True
            except Exception as e:
                logger.warning(f"테마 색상 {style_dict_to_apply['color_theme_index']} 적용 실패: {e}")
        
        if 'language_id' in style_dict_to_apply and style_dict_to_apply['language_id'] is not None:
            try:
                font.language_id = style_dict_to_apply['language_id']
            except Exception as e_lang:
                logger.warning(f"language_id '{style_dict_to_apply['language_id']}' 적용 실패 (_apply_style_properties): {e_lang}")


    def _apply_text_style(self, run, style_to_apply): # run은 _Run 객체
        if not style_to_apply or run is None:
            return
        self._apply_style_properties(run.font, style_to_apply)
        
        if 'hyperlink_address' in style_to_apply and style_to_apply['hyperlink_address']:
            try:
                # run.hyperlink는 존재하지 않으면 _Hyperlink 객체를 생성함
                hlink = run.hyperlink
                hlink.address = style_to_apply['hyperlink_address']
            except Exception as e:
                logger.warning(f"실행(run)에 하이퍼링크 주소 적용 오류: {e}")

    def translate_presentation(self, file_path, src_lang, tgt_lang, translator, ocr_handler,
                                model_name, ollama_service, font_code_for_render,
                                task_log_filepath,
                                progress_callback=None, stop_event=None):
            # === 최상위 try: 작업 로그 파일 처리 및 전체 프로세스 예외 처리 ===
            try:
                with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                    start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    f_task_log.write(f"--- 프레젠테이션 번역 로그 시작: {os.path.basename(file_path)} ---\n")
                    f_task_log.write(f"시작 시간: {start_time_str}\n원본 파일: {file_path}\n소스 언어: {src_lang}, 대상 언어: {tgt_lang}, 모델: {model_name}\n")
                    f_task_log.write(f"OCR 핸들러 사용: {'예' if ocr_handler else '아니오 (이미지 내 텍스트 번역 안됨)'}\n결과 렌더링 폰트 코드: {font_code_for_render}\n\n")
                    logger.info(f"'{os.path.basename(file_path)}' 프레젠테이션 번역 시작. 로그: {task_log_filepath}")

                    # === 내부 try: 프레젠테이션 처리 로직 자체의 예외 처리 ===
                    try:
                        prs = Presentation(file_path)
                        safe_tgt_lang = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                        output_filename = os.path.splitext(file_path)[0] + f"_{safe_tgt_lang}_translated.pptx"

                        elements_map = []
                        f_task_log.write("--- 번역 대상 요소 분석 시작 (shape_id 또는 table_obj_id 사용) ---\n")
                        for slide_idx, slide in enumerate(prs.slides):
                            if stop_event and stop_event.is_set(): break
                            for shape_idx, shape_in_slide in enumerate(slide.shapes):
                                if stop_event and stop_event.is_set(): break
                                
                                current_shape_id = getattr(shape_in_slide, 'shape_id', None)
                                element_name = shape_in_slide.name or f"S{slide_idx+1}_Shape{shape_idx}_Id{current_shape_id or 'Unknown'}"

                                item_to_add = None
                                # 도형 객체 자체를 저장하여 나중에 ID로 다시 찾지 않도록 함
                                if shape_in_slide.has_text_frame and shape_in_slide.text_frame.text and shape_in_slide.text_frame.text.strip():
                                    item_to_add = {'type': 'text', 'slide_idx': slide_idx, 'shape_id': current_shape_id, 'shape_obj_ref': shape_in_slide, 'name': element_name}
                                elif shape_in_slide.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                    if ocr_handler:
                                        item_to_add = {'type': 'image', 'slide_idx': slide_idx, 'shape_id': current_shape_id, 'shape_obj_ref': shape_in_slide, 'name': element_name}
                                    else:
                                        f_task_log.write(f"  S{slide_idx+1}-Img '{element_name}' (ID:{current_shape_id}): OCR 핸들러 없어 번역 대상에서 제외.\n")
                                elif shape_in_slide.has_table:
                                    table_shape_id_for_log = current_shape_id 
                                    for r_idx, row in enumerate(shape_in_slide.table.rows):
                                        for c_idx, cell in enumerate(row.cells):
                                            if cell.text_frame.text and cell.text_frame.text.strip():
                                                elements_map.append({
                                                    'type': 'table_cell',
                                                    'slide_idx': slide_idx,
                                                    'shape_id': table_shape_id_for_log, # 로깅 및 참조용 테이블 도형 ID
                                                    'table_shape_obj_ref': shape_in_slide, # 테이블 도형 객체 직접 참조
                                                    'row_idx': r_idx,
                                                    'col_idx': c_idx,
                                                    'name': f"{element_name}_R{r_idx}C{c_idx}"
                                                })
                                if item_to_add: elements_map.append(item_to_add)
                        
                        total_elements_to_translate = len(elements_map)
                        f_task_log.write(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견.\n--- 요소 분석 완료 ---\n\n")
                        logger.info(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견.")

                        if not elements_map and not (stop_event and stop_event.is_set()):
                            prs.save(output_filename)
                            msg = "번역할 대상 요소가 없어 원본 파일을 복사본으로 저장합니다."
                            f_task_log.write(msg + "\n")
                            logger.info(msg)
                            return output_filename

                        translated_count = 0
                        f_task_log.write("--- 번역 작업 시작 ---\n")
                        original_shape_paragraph_styles = {}

                        for item_info in elements_map:
                            if stop_event and stop_event.is_set(): break

                            slide_idx = item_info['slide_idx']
                            element_name_for_log = item_info['name']
                            current_shape_obj = None 

                            if 'shape_obj_ref' in item_info:
                                current_shape_obj = item_info['shape_obj_ref']
                            elif 'table_shape_obj_ref' in item_info:
                                current_shape_obj = item_info['table_shape_obj_ref']
                            
                            if current_shape_obj is None:
                                f_task_log.write(f"  오류: S{slide_idx+1} ('{element_name_for_log}') 요소 객체 참조 실패. 건너뜀.\n")
                                logger.warning(f"Failed to reference shape object for '{element_name_for_log}' on slide {slide_idx+1}. Skipping.")
                                translated_count += 1
                                if progress_callback: progress_callback(slide_idx + 1, item_info['type'], translated_count, total_elements_to_translate, "요소 참조 오류")
                                continue
                            
                            current_shape_id_for_log = getattr(current_shape_obj, 'shape_id', 'N/A')
                            left_val = getattr(current_shape_obj, 'left', 0)
                            top_val = getattr(current_shape_obj, 'top', 0)
                            width_val = getattr(current_shape_obj, 'width', 0)
                            height_val = getattr(current_shape_obj, 'height', 0)

                            log_shape_details = (f"[S{slide_idx+1}] 요소 처리 시작: '{element_name_for_log}' (ID: {current_shape_id_for_log}), 타입: {item_info['type']}, "
                                                f"위치: L{left_val/914400:.2f}cm, T{top_val/914400:.2f}cm, W{width_val/914400:.2f}cm, H{height_val/914400:.2f}cm\n")
                            logger.debug(log_shape_details.strip())
                            f_task_log.write(log_shape_details)
                            
                            current_text_for_progress = ""
                            item_type = item_info['type']

                            if item_type == 'text' or item_type == 'table_cell':
                                text_frame = None
                                row_idx = item_info.get('row_idx', -1)
                                col_idx = item_info.get('col_idx', -1)

                                if item_type == 'text':
                                    if current_shape_obj.has_text_frame:
                                        text_frame = current_shape_obj.text_frame
                                elif item_type == 'table_cell':
                                    if current_shape_obj.has_table:
                                        try:
                                            cell = current_shape_obj.table.cell(row_idx, col_idx)
                                            text_frame = cell.text_frame
                                        except IndexError:
                                            logger.error(f"테이블 셀 접근 오류: {element_name_for_log} (S{slide_idx+1} R{row_idx}C{col_idx})")
                                            f_task_log.write(f"  오류: 테이블 셀 접근 실패 {element_name_for_log}\n")
                                
                                if text_frame and hasattr(text_frame, 'text') and text_frame.text and text_frame.text.strip():
                                    original_text = text_frame.text
                                    current_text_for_progress = original_text
                                    
                                    unique_key_for_style = (slide_idx, id(current_shape_obj), row_idx, col_idx)

                                    if unique_key_for_style not in original_shape_paragraph_styles:
                                        collected_para_styles = []
                                        for para in text_frame.paragraphs:
                                            para_default_font_style = self._get_style_properties(para.font)
                                            runs_data = []
                                            if para.runs:
                                                runs_data = [{'text': r.text, 'style': self._get_text_style(r)} for r in para.runs]
                                            elif para.text and para.text.strip(): # 실행은 없지만 단락에 텍스트가 있는 경우
                                                # 단락 기본 스타일을 실행 스타일로 간주
                                                run_style_as_para_default = para_default_font_style.copy()
                                                run_style_as_para_default['hyperlink_address'] = None # 실행이 없으므로 하이퍼링크 없음
                                                runs_data = [{'text': para.text, 'style': run_style_as_para_default}]
                                            
                                            collected_para_styles.append({
                                                'runs': runs_data,
                                                'alignment': para.alignment,
                                                'level': para.level,
                                                'space_before': para.space_before,
                                                'space_after': para.space_after,
                                                'line_spacing': para.line_spacing,
                                                'paragraph_default_run_style': para_default_font_style
                                            })
                                        original_shape_paragraph_styles[unique_key_for_style] = collected_para_styles
                                        f_task_log.write(f"    '{element_name_for_log}'의 원본 단락/실행 스타일 정보 저장됨 ({len(collected_para_styles)}개 단락).\n")

                                    if should_skip_translation(original_text):
                                        f_task_log.write(f"  [스킵됨 - 번역 불필요] \"{original_text.replace(chr(10), ' / ').strip()[:100]}...\"\n")
                                    else:
                                        translated_text = translator.translate_text(original_text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=False)
                                        log_original_text_snippet = original_text.replace(chr(10), ' / ').strip()[:100]
                                        log_translated_text_snippet = translated_text.replace(chr(10), ' / ').strip()[:100]
                                        f_task_log.write(f"  [번역 전] \"{log_original_text_snippet}...\" -> [번역 후] \"{log_translated_text_snippet}...\"\n")

                                        if "오류:" not in translated_text:
                                            stored_para_infos_for_shape = original_shape_paragraph_styles.get(unique_key_for_style, [])
                                            
                                            original_auto_size = text_frame.auto_size if hasattr(text_frame, 'auto_size') else None
                                            original_word_wrap = text_frame.word_wrap if hasattr(text_frame, 'word_wrap') else None
                                            
                                            if original_auto_size is not None and original_auto_size != MSO_AUTO_SIZE.NONE:
                                                text_frame.auto_size = MSO_AUTO_SIZE.NONE
                                                if original_word_wrap is not None:
                                                    text_frame.word_wrap = True

                                            text_frame.clear() 
                                            # 어제 논의된 XML 레벨에서의 추가 정리 (선택적 강화 조치)
                                            # 이 부분은 text_frame.clear()가 완벽하지 않다고 의심될 때 사용합니다.
                                            # 일반적으로 python-pptx의 clear()는 잘 동작하므로, 이 코드가 반드시 필요한 것은 아닐 수 있습니다.
                                            # 문제가 지속될 경우에만 활성화하는 것을 고려해볼 수 있습니다.
                                            if hasattr(text_frame, '_element') and text_frame._element is not None:
                                                txBody = text_frame._element
                                                # 네임스페이스를 정확히 사용하려면 아래와 같이 nsdecls를 사용해야 합니다.
                                                # from pptx.oxml.ns import nsmap
                                                # para_tag_name = '{%s}p' % nsmap['a'] # 'a'는 DrawingML 네임스페이스
                                                # children_to_remove = [child for child in txBody if child.tag == para_tag_name]
                                                
                                                # 간이 방식 (태그 끝이 '}p'로 끝나는 것으로 확인)
                                                children_to_remove = [child for child in txBody if child.tag.endswith('}p')]
                                                if children_to_remove:
                                                    f_task_log.write(f"      (강화된 정리) XML 레벨에서 기존 단락 {len(children_to_remove)}개 제거 시도...\n")
                                                    for child_to_remove_xml in children_to_remove: # 변수명 변경
                                                        txBody.remove(child_to_remove_xml)
                                                    f_task_log.write(f"      (강화된 정리) XML 레벨 기존 단락 제거 완료.\n")
                                            
                                            # 번역기에서 후처리를 잘 했다고 가정하고 lines를 가져옵니다.
                                            # 번역기 결과 자체에서 앞뒤 공백/개행이 제거되어야 합니다.
                                            lines = translated_text.splitlines()

                                            # 번역 결과가 완전히 비었거나, 개행으로만 구성되어 splitlines()가 빈 리스트를 반환한 경우,
                                            # 또는 모든 줄이 공백으로만 이루어진 경우를 처리합니다.
                                            processed_lines = []
                                            if lines:
                                                for line_content_raw in lines:
                                                    stripped_line = line_content_raw.strip()
                                                    if stripped_line: # 내용이 있는 줄만 추가
                                                        processed_lines.append(line_content_raw) # 원본 공백 유지하며 추가 (내부 들여쓰기 등)
                                                    elif processed_lines or i < len(lines) -1 : # 첫 줄이 아니고, 마지막 줄도 아니면서 빈 줄이거나, 이미 내용있는 줄 뒤의 빈줄
                                                        processed_lines.append("") # 의도된 빈 줄일 수 있으므로 ""로 추가 (단락은 생성됨)
                                            
                                            if not processed_lines: # 모든 줄이 비어있거나 내용이 없었으면
                                                processed_lines = [" "] # 공백 한 칸짜리 단락 하나 생성

                                            lines = processed_lines # 최종 사용할 lines

                                            for i, line_content in enumerate(lines):
                                                p = text_frame.add_paragraph()
                                                para_template = stored_para_infos_for_shape[min(i, len(stored_para_infos_for_shape)-1)] if stored_para_infos_for_shape else None
                                                
                                                if para_template:
                                                    if para_template.get('alignment') is not None: p.alignment = para_template['alignment']
                                                    p.level = para_template.get('level', 0) # level은 항상 존재 (기본값 0)
                                                    if para_template.get('space_before') is not None: p.space_before = para_template['space_before']
                                                    if para_template.get('space_after') is not None: p.space_after = para_template['space_after']
                                                    if para_template.get('line_spacing') is not None: p.line_spacing = para_template['line_spacing']

                                                    if 'paragraph_default_run_style' in para_template:
                                                        self._apply_style_properties(p.font, para_template['paragraph_default_run_style'])
                                                        f_task_log.write(f"      단락 {i+1}에 원본 단락 기본 글꼴 스타일 적용됨.\n")
                                                
                                                run = p.add_run()
                                                run.text = line_content if line_content.strip() else " "
                                                
                                                log_added_line = run.text.strip()[:50].replace(chr(10), ' ')
                                                f_task_log.write(f"    번역된 단락 {i+1} 추가: '{log_added_line}...'\n")

                                                if para_template and para_template.get('runs') and para_template['runs']:
                                                    first_original_run_style = para_template['runs'][0]['style']
                                                    self._apply_text_style(run, first_original_run_style)
                                                    f_task_log.write(f"        실행(run)에 원본 첫 실행 스타일 적용됨 (크기: {first_original_run_style.get('size')}, 색상 RGB: {first_original_run_style.get('color_rgb')}).\n")
                                                elif para_template and 'paragraph_default_run_style' in para_template:
                                                    self._apply_style_properties(run.font, para_template['paragraph_default_run_style'])
                                                    f_task_log.write(f"        실행(run)에 단락 기본 글꼴 스타일 적용됨 (실행 스타일 정보 부족).\n")

                                            if original_auto_size is not None and original_auto_size != MSO_AUTO_SIZE.NONE:
                                                text_frame.auto_size = original_auto_size
                                            if original_word_wrap is not None and text_frame.word_wrap != original_word_wrap :
                                                text_frame.word_wrap = original_word_wrap
                                        else:
                                            f_task_log.write(f"  -> 텍스트 번역 실패 또는 빈 결과: {translated_text}\n")
                                else:
                                    f_task_log.write(f"  '{element_name_for_log}' 건너뜀 (텍스트 내용 없음 또는 텍스트 프레임/텍스트 속성 없음).\n")
                            
                            elif item_type == 'image' and ocr_handler:
                                if current_shape_obj is None or not hasattr(current_shape_obj, 'image') or not hasattr(current_shape_obj.image, 'blob'):
                                    f_task_log.write(f"    경고: '{element_name_for_log}' (ID:{getattr(current_shape_obj, 'shape_id', 'N/A')}) 이미지 객체 또는 이미지 데이터 접근 불가. 건너뜀.\n")
                                    logger.warning(f"Cannot access image data for '{element_name_for_log}' (ID:{getattr(current_shape_obj, 'shape_id', 'N/A')}). Skipping.")
                                else:
                                    try:
                                        image_bytes_io = io.BytesIO(current_shape_obj.image.blob)
                                        with Image.open(image_bytes_io) as img_pil_for_ocr_source:
                                            img_pil_for_ocr = img_pil_for_ocr_source.convert("RGB")
                                            try:
                                                hasher = hashlib.md5()
                                                img_byte_arr_for_hash = io.BytesIO()
                                                img_pil_for_ocr.save(img_byte_arr_for_hash, format='PNG')
                                                hasher.update(img_byte_arr_for_hash.getvalue())
                                                img_hash = hasher.hexdigest()
                                                f_task_log.write(f"    OCR 대상 RGB 이미지 해시 (MD5): {img_hash}\n")
                                            except Exception as e_hash: logger.warning(f"    이미지 해시 생성 중 오류: {e_hash}")

                                            ocr_results = ocr_handler.ocr_image(img_pil_for_ocr)
                                            if ocr_results:
                                                f_task_log.write(f"      이미지 내 OCR 텍스트 {len(ocr_results)}개 발견.\n")
                                                image_bytes_io.seek(0) 
                                                with Image.open(image_bytes_io) as edited_image_pil_base:
                                                    edited_image_pil = edited_image_pil_base.copy()
                                                    any_ocr_text_translated_and_rendered = False
                                                    for ocr_idx, ocr_item in enumerate(ocr_results):
                                                        if not (isinstance(ocr_item, (list, tuple)) and len(ocr_item) >= 2): # 최소 2개 요소 확인
                                                            logger.warning(f"      잘못된 OCR 결과 항목 형식: {ocr_item}. 건너뜀.")
                                                            f_task_log.write(f"        잘못된 OCR 결과 항목 형식. 건너뜀.\n")
                                                            continue

                                                        box = ocr_item[0]
                                                        text_confidence_tuple = ocr_item[1]
                                                        ocr_angle = ocr_item[2] if len(ocr_item) > 2 else None # 각도 정보 (선택적)

                                                        if not (isinstance(text_confidence_tuple, (list, tuple)) and len(text_confidence_tuple) == 2):
                                                            logger.warning(f"      잘못된 OCR 텍스트/신뢰도 튜플 형식: {text_confidence_tuple}. 건너뜀.")
                                                            f_task_log.write(f"        잘못된 OCR 텍스트/신뢰도 튜플 형식. 건너뜀.\n")
                                                            continue
                                                        
                                                        text, confidence = text_confidence_tuple
                                                        # 이제 box, text, confidence, ocr_angle 변수 사용 가능
                                                        if stop_event and stop_event.is_set(): break
                                                        current_text_for_progress = text
                                                        log_ocr_text = text.replace(chr(10), ' ').strip()[:50]
                                                        f_task_log.write(f"        OCR Text [{ocr_idx+1}]: \"{log_ocr_text}...\" (Conf: {confidence:.2f})\n")
                                                        if not is_ocr_text_valid(text):
                                                            f_task_log.write(f"          -> [스킵됨 - OCR 유효성 낮음]\n"); continue
                                                        if should_skip_translation(text):
                                                            f_task_log.write(f"          -> [스킵됨 - 번역 불필요]\n"); continue
                                                        translated_ocr_text = translator.translate_text(text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=True)
                                                        log_translated_ocr = translated_ocr_text.replace(chr(10), ' ').strip()[:50]
                                                        f_task_log.write(f"          -> 번역 결과: \"{log_translated_ocr}...\"\n")
                                                        if "오류:" not in translated_ocr_text and translated_ocr_text.strip():
                                                            try:
                                                                edited_image_pil = ocr_handler.render_translated_text_on_image(
                                                                    edited_image_pil, box, translated_ocr_text,
                                                                    font_code_for_render=font_code_for_render, original_text=text
                                                                )
                                                                any_ocr_text_translated_and_rendered = True
                                                                f_task_log.write(f"            -> 렌더링 완료.\n")
                                                            except Exception as e_render:
                                                                f_task_log.write(f"            오류: OCR 텍스트 렌더링 실패: {e_render}\n")
                                                                logger.error(f"OCR 렌더링 실패 (S{slide_idx+1}-ImgID:{getattr(current_shape_obj, 'shape_id', 'N/A')}): {e_render}", exc_info=True)
                                                        else:
                                                            f_task_log.write(f"            -> 번역 실패 또는 빈 결과로 렌더링 안함.\n")
                                                    if stop_event and stop_event.is_set(): break
                                                    if any_ocr_text_translated_and_rendered:
                                                        image_stream_for_replacement = io.BytesIO()
                                                        save_format = edited_image_pil_base.format if edited_image_pil_base.format and edited_image_pil_base.format.upper() in ['JPEG', 'PNG', 'GIF', 'BMP', 'TIFF'] else 'PNG'
                                                        edited_image_pil.save(image_stream_for_replacement, format=save_format)
                                                        image_stream_for_replacement.seek(0)
                                                        
                                                        left_p, top_p = current_shape_obj.left, current_shape_obj.top
                                                        width_p, height_p = current_shape_obj.width, current_shape_obj.height
                                                        current_slide_for_pic = prs.slides[slide_idx]
                                                        pic_element = current_shape_obj.element
                                                        pic_parent = pic_element.getparent()
                                                        if pic_parent is not None: pic_parent.remove(pic_element)
                                                        new_pic_shape = current_slide_for_pic.shapes.add_picture(image_stream_for_replacement, left_p, top_p, width=width_p, height=height_p)
                                                        f_task_log.write(f"      이미지 '{element_name_for_log}' 교체 완료.\n")
                                                    else: f_task_log.write(f"      이미지 '{element_name_for_log}' 변경 없음.\n")
                                            else: f_task_log.write(f"      이미지 '{element_name_for_log}' 내 OCR 텍스트 발견되지 않음.\n")
                                    except AttributeError as e_attr_img:
                                        f_task_log.write(f"    경고: '{element_name_for_log}' 이미지 처리 중 속성 오류: {e_attr_img}. 건너뜀.\n")
                                        logger.warning(f"Attribute error processing image '{element_name_for_log}': {e_attr_img}", exc_info=True)
                                    except IOError as e_io_img:
                                        f_task_log.write(f"    이미지 처리 중 I/O 오류 '{element_name_for_log}': {e_io_img}. 건너뜀.\n")
                                        logger.error(f"I/O error during image processing '{element_name_for_log}': {e_io_img}", exc_info=False)
                                    except Exception as e_img_proc_detail:
                                        f_task_log.write(f"    이미지 '{element_name_for_log}' 처리 중 예기치 않은 오류: {e_img_proc_detail}. 건너뜀.\n")
                                        logger.error(f"Unexpected error processing image '{element_name_for_log}': {e_img_proc_detail}", exc_info=True)
                            
                            translated_count += 1
                            if progress_callback and not (stop_event and stop_event.is_set()):
                                progress_callback(slide_idx + 1, item_type, translated_count, total_elements_to_translate, current_text_for_progress)
                            f_task_log.write("\n") 

                        f_task_log.write("--- 번역 작업 완료 ---\n")
                        if stop_event and stop_event.is_set():
                            stopped_filename = output_filename.replace(".pptx", "_stopped.pptx")
                            prs.save(stopped_filename)
                            f_task_log.write(f"번역 중지됨. 부분 저장 파일: {stopped_filename}\n")
                            logger.info(f"번역 중지됨. 부분 저장 파일: {stopped_filename}")
                            return stopped_filename
                        
                        prs.save(output_filename)
                        f_task_log.write(f"번역된 파일 저장 완료: {output_filename}\n")
                        logger.info(f"번역된 파일 저장 완료: {output_filename}")
                        return output_filename

                    except Exception as e_translate_general:
                        err_msg = f"프레젠테이션 번역 중 심각한 오류 발생: {e_translate_general}"
                        logger.error(err_msg, exc_info=True)
                        f_task_log.write(f"오류: {err_msg}\n상세 정보: {traceback.format_exc()}\n")
                        return None

            except IOError as e_file_io:
                logger.error(f"작업 로그 파일 처리 오류 ({task_log_filepath}): {e_file_io}", exc_info=True)
                return None
            except Exception as e_outer_unexpected:
                logger.critical(f"translate_presentation 함수 외부 수준에서 예기치 않은 오류 발생: {e_outer_unexpected}", exc_info=True)
                return None