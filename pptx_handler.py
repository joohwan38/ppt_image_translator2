from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.lang import MSO_LANGUAGE_ID

import os
import io
import logging
import re
from datetime import datetime
import hashlib
import traceback
from PIL import Image

logger = logging.getLogger(__name__)

MIN_MEANINGFUL_CHAR_RATIO_SKIP = 0.1
MIN_MEANINGFUL_CHAR_RATIO_OCR = 0.1
MEANINGFUL_CHAR_PATTERN = re.compile(
    r'[a-zA-Z'
    r'\u00C0-\u024F'
    r'\u1E00-\u1EFF'
    r'\u0600-\u06FF'
    r'\u0750-\u077F'
    r'\u08A0-\u08FF'
    r'\u3040-\u30ff'
    r'\u3131-\uD79D'
    r'\u4e00-\u9fff'
    r'\u0E00-\u0E7F'
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
        if MEANINGFUL_CHAR_PATTERN.search(stripped_text):
            logger.debug(f"번역 시도 (짧은 문자열, 의미 문자 포함): '{stripped_text}'")
            return False
        else:
            logger.debug(f"번역 스킵 (짧고 의미 있는 문자 없음): '{stripped_text}'")
            return True
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
        if MEANINGFUL_CHAR_PATTERN.search(stripped_text):
             logger.debug(f"OCR 유효 (매우 짧은 문자열, 의미 문자 포함): '{stripped_text}'")
             return True
        else:
            logger.debug(f"OCR 유효성 스킵 (매우 짧고 의미 있는 문자 없음): '{stripped_text}'")
            return False
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
        info = {"slide_count": 0, "text_elements": 0, "image_elements": 0, "chart_elements": 0}
        try:
            prs = Presentation(file_path)
            info["slide_count"] = len(prs.slides)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame and hasattr(shape.text_frame, 'text') and \
                       shape.text_frame.text and shape.text_frame.text.strip():
                        info["text_elements"] += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        info["image_elements"] += 1
                    elif shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if hasattr(cell.text_frame, 'text') and cell.text_frame.text and cell.text_frame.text.strip():
                                    info["text_elements"] += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                        info["chart_elements"] +=1
            logger.info(f"파일 분석 완료 (OCR 미수행): Slides:{info['slide_count']}, Text:{info['text_elements']}, Img:{info['image_elements']}, Chart:{info['chart_elements']}")
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
            lang_id = font_object.language_id
            style_props['language_id'] = lang_id
        except ValueError as e:
            logger.warning(f"Unsupported language_id value in _get_style_properties. Defaulting to None. Error: {e}")
            style_props['language_id'] = None
        except AttributeError:
            logger.warning(f"font_object has no language_id attribute in _get_style_properties. Defaulting to None.")
            style_props['language_id'] = None

        if hasattr(font_object.color, 'type') and font_object.color.type is not None:
            color_type = font_object.color.type
            if color_type == MSO_COLOR_TYPE.RGB:
                if font_object.color.rgb is not None:
                    style_props['color_rgb'] = tuple(font_object.color.rgb)
            elif color_type == MSO_COLOR_TYPE.SCHEME:
                style_props['color_theme_index'] = font_object.color.theme_color
                if hasattr(font_object.color, 'brightness') and font_object.color.brightness is not None:
                    style_props['color_brightness'] = font_object.color.brightness
                else:
                    style_props['color_brightness'] = 0.0
        return style_props

    def _get_text_style(self, run):
        style = self._get_style_properties(run.font)
        style['hyperlink_address'] = run.hyperlink.address if run.hyperlink and hasattr(run.hyperlink, 'address') and run.hyperlink.address else None
        return style

    def _apply_style_properties(self, target_font_object, style_dict_to_apply):
        if not style_dict_to_apply or target_font_object is None:
            return
        font = target_font_object
        if 'name' in style_dict_to_apply and style_dict_to_apply['name'] is not None:
            try: font.name = style_dict_to_apply['name']
            except Exception as e: logger.warning(f"Failed to apply font name '{style_dict_to_apply['name']}': {e}")
        if 'size' in style_dict_to_apply and style_dict_to_apply['size'] is not None:
            try: font.size = style_dict_to_apply['size']
            except Exception as e: logger.warning(f"Failed to apply font size '{style_dict_to_apply['size']}': {e}")
        if 'bold' in style_dict_to_apply and style_dict_to_apply['bold'] is not None:
            try: font.bold = style_dict_to_apply['bold']
            except Exception as e: logger.warning(f"Failed to apply font bold '{style_dict_to_apply['bold']}': {e}")
        if 'italic' in style_dict_to_apply and style_dict_to_apply['italic'] is not None:
            try: font.italic = style_dict_to_apply['italic']
            except Exception as e: logger.warning(f"Failed to apply font italic '{style_dict_to_apply['italic']}': {e}")
        if 'underline' in style_dict_to_apply and style_dict_to_apply['underline'] is not None:
            try: font.underline = style_dict_to_apply['underline']
            except Exception as e: logger.warning(f"Failed to apply font underline '{style_dict_to_apply['underline']}': {e}")

        applied_color = False
        if 'color_rgb' in style_dict_to_apply and style_dict_to_apply['color_rgb'] is not None:
            try:
                rgb_tuple = tuple(int(c) for c in style_dict_to_apply['color_rgb'])
                if len(rgb_tuple) == 3:
                    font.color.rgb = RGBColor(*rgb_tuple)
                    applied_color = True
                else:
                    logger.warning(f"Invalid RGB tuple length: {rgb_tuple}")
            except Exception as e:
                logger.warning(f"RGB 색상 {style_dict_to_apply['color_rgb']} 적용 실패: {e}")
        
        if not applied_color and 'color_theme_index' in style_dict_to_apply and style_dict_to_apply['color_theme_index'] is not None:
            try:
                font.color.theme_color = style_dict_to_apply['color_theme_index']
                if 'color_brightness' in style_dict_to_apply and style_dict_to_apply['color_brightness'] is not None:
                    brightness_val = float(style_dict_to_apply['color_brightness'])
                    font.color.brightness = max(-1.0, min(1.0, brightness_val))
            except Exception as e:
                logger.warning(f"테마 색상 {style_dict_to_apply['color_theme_index']} 적용 실패: {e}")
        
        if 'language_id' in style_dict_to_apply and style_dict_to_apply['language_id'] is not None:
            try:
                font.language_id = style_dict_to_apply['language_id']
            except Exception as e_lang:
                logger.warning(f"language_id '{style_dict_to_apply['language_id']}' 적용 실패 (_apply_style_properties): {e_lang}")

    def _apply_text_style(self, run, style_to_apply):
        if not style_to_apply or run is None:
            return
        self._apply_style_properties(run.font, style_to_apply)
        if 'hyperlink_address' in style_to_apply and style_to_apply['hyperlink_address']:
            try:
                hlink = run.hyperlink
                hlink.address = style_to_apply['hyperlink_address']
            except Exception as e:
                logger.warning(f"실행(run)에 하이퍼링크 주소 적용 오류: {e}")

    def translate_presentation(self, file_path, src_lang, tgt_lang, translator, ocr_handler,
                               model_name, ollama_service, font_code_for_render,
                               task_log_filepath,
                               progress_callback=None, stop_event=None):
        try:
            with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f_task_log.write(f"--- 프레젠테이션 번역 로그 시작: {os.path.basename(file_path)} ---\n")
                f_task_log.write(f"시작 시간: {start_time_str}\n원본 파일: {file_path}\n소스 언어: {src_lang}, 대상 언어: {tgt_lang}, 모델: {model_name}\n")
                f_task_log.write(f"OCR 핸들러 사용: {'예' if ocr_handler else '아니오 (이미지 내 텍스트 번역 안됨)'}\n결과 렌더링 폰트 코드: {font_code_for_render}\n\n")
                logger.info(f"'{os.path.basename(file_path)}' 프레젠테이션 번역 시작. 로그: {task_log_filepath}")

                try:
                    prs = Presentation(file_path)
                    safe_tgt_lang = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                    output_filename = os.path.splitext(file_path)[0] + f"_{safe_tgt_lang}_translated.pptx"

                    elements_map = []
                    f_task_log.write("--- 번역 대상 요소 분석 시작 ---\n")
                    for slide_idx, slide in enumerate(prs.slides):
                        if stop_event and stop_event.is_set(): break
                        f_task_log.write(f"  슬라이드 {slide_idx + 1} 분석 중...\n")
                        for shape_idx, shape_in_slide in enumerate(slide.shapes):
                            if stop_event and stop_event.is_set(): break
                            
                            current_shape_id = getattr(shape_in_slide, 'shape_id', None)
                            element_name = shape_in_slide.name or f"S{slide_idx+1}_Shape{shape_idx}_Id{current_shape_id or 'Unknown'}"
                            item_to_add = None

                            if shape_in_slide.has_text_frame and hasattr(shape_in_slide.text_frame, 'text') and \
                               shape_in_slide.text_frame.text and shape_in_slide.text_frame.text.strip():
                                item_to_add = {'type': 'text', 'slide_idx': slide_idx, 'shape_id': current_shape_id, 'shape_obj_ref': shape_in_slide, 'name': element_name}
                            elif shape_in_slide.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                if ocr_handler:
                                    item_to_add = {'type': 'image', 'slide_idx': slide_idx, 'shape_id': current_shape_id, 'shape_obj_ref': shape_in_slide, 'name': element_name}
                            elif shape_in_slide.has_table:
                                table_shape_id_for_log = current_shape_id 
                                for r_idx, row in enumerate(shape_in_slide.table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        if hasattr(cell.text_frame, 'text') and cell.text_frame.text and cell.text_frame.text.strip():
                                            elements_map.append({
                                                'type': 'table_cell', 'slide_idx': slide_idx,
                                                'shape_id': table_shape_id_for_log,
                                                'table_shape_obj_ref': shape_in_slide,
                                                'row_idx': r_idx, 'col_idx': c_idx,
                                                'name': f"{element_name}_R{r_idx}C{c_idx}"
                                            })
                            elif shape_in_slide.shape_type == MSO_SHAPE_TYPE.CHART:
                                chart_shape_id_for_log = current_shape_id
                                chart = shape_in_slide.chart
                                chart_element_base_name = element_name or f"S{slide_idx+1}_Chart{shape_idx}"
                                f_task_log.write(f"      차트 요소 '{chart_element_base_name}' (ID: {chart_shape_id_for_log}) 분석 중...\n")

                                if chart.has_title :
                                    f_task_log.write(f"        차트 제목(has_title=True) 발견. 텍스트 프레임 확인 중...\n")
                                    if hasattr(chart.chart_title, 'text_frame') and \
                                       hasattr(chart.chart_title.text_frame, 'text') and \
                                       chart.chart_title.text_frame.text:
                                        chart_title_text_val = chart.chart_title.text_frame.text
                                        if chart_title_text_val.strip():
                                            elements_map.append({
                                                'type': 'chart_title', 'slide_idx': slide_idx, 
                                                'shape_id': chart_shape_id_for_log, 'chart_shape_obj_ref': shape_in_slide,
                                                'name': f"{chart_element_base_name}_Title"
                                            })
                                            f_task_log.write(f"          -> 차트 제목 elements_map에 추가: '{chart_title_text_val.strip()[:30]}...'\n")
                                        else:
                                            f_task_log.write(f"          차트 제목 텍스트는 있으나 공백만 포함. 추가 안 함: '{chart_title_text_val[:30]}...'\n")
                                    else:
                                        f_task_log.write(f"          차트 제목에 텍스트 프레임 또는 텍스트 속성 없음. 추가 안 함.\n")
                                else:
                                    f_task_log.write(f"        차트 제목(has_title=False) 없음.\n")
                                
                                if hasattr(chart, 'category_axis') and chart.category_axis is not None:
                                    f_task_log.write(f"        X축(category_axis) 발견. 축 제목 확인 중...\n")
                                    if chart.category_axis.has_title:
                                        f_task_log.write(f"          X축 제목(has_title=True) 발견. 텍스트 프레임 확인 중...\n")
                                        if hasattr(chart.category_axis.axis_title, 'text_frame') and \
                                           hasattr(chart.category_axis.axis_title.text_frame, 'text') and \
                                           chart.category_axis.axis_title.text_frame.text:
                                            xaxis_title_text_val = chart.category_axis.axis_title.text_frame.text
                                            if xaxis_title_text_val.strip():
                                                elements_map.append({
                                                    'type': 'chart_axis_title_category', 'slide_idx': slide_idx,
                                                    'shape_id': chart_shape_id_for_log, 'chart_shape_obj_ref': shape_in_slide,
                                                    'name': f"{chart_element_base_name}_XAxisTitle"
                                                })
                                                f_task_log.write(f"            -> X축 제목 elements_map에 추가: '{xaxis_title_text_val.strip()[:30]}...'\n")
                                            else:
                                                f_task_log.write(f"            X축 제목 텍스트는 있으나 공백만 포함. 추가 안 함: '{xaxis_title_text_val[:30]}...'\n")
                                        else:
                                            f_task_log.write(f"            X축 제목에 텍스트 프레임 또는 텍스트 속성 없음. 추가 안 함.\n")
                                    else:
                                        f_task_log.write(f"          X축 제목(has_title=False) 없음.\n")
                                else:
                                    f_task_log.write(f"        X축(category_axis) 없음.\n")

                                if hasattr(chart, 'value_axis') and chart.value_axis is not None:
                                    f_task_log.write(f"        Y축(value_axis) 발견. 축 제목 확인 중...\n")
                                    if chart.value_axis.has_title:
                                        f_task_log.write(f"          Y축 제목(has_title=True) 발견. 텍스트 프레임 확인 중...\n")
                                        if hasattr(chart.value_axis.axis_title, 'text_frame') and \
                                           hasattr(chart.value_axis.axis_title.text_frame, 'text') and \
                                           chart.value_axis.axis_title.text_frame.text:
                                            yaxis_title_text_val = chart.value_axis.axis_title.text_frame.text
                                            if yaxis_title_text_val.strip():
                                                elements_map.append({
                                                    'type': 'chart_axis_title_value', 'slide_idx': slide_idx,
                                                    'shape_id': chart_shape_id_for_log, 'chart_shape_obj_ref': shape_in_slide,
                                                    'name': f"{chart_element_base_name}_YAxisTitle"
                                                })
                                                f_task_log.write(f"            -> Y축 제목 elements_map에 추가: '{yaxis_title_text_val.strip()[:30]}...'\n")
                                            else:
                                                f_task_log.write(f"            Y축 제목 텍스트는 있으나 공백만 포함. 추가 안 함: '{yaxis_title_text_val[:30]}...'\n")
                                        else:
                                            f_task_log.write(f"            Y축 제목에 텍스트 프레임 또는 텍스트 속성 없음. 추가 안 함.\n")
                                    else:
                                        f_task_log.write(f"          Y축 제목(has_title=False) 없음.\n")
                                else:
                                    f_task_log.write(f"        Y축(value_axis) 없음.\n")
                                
                                if chart.plots and len(chart.plots) > 0 :
                                    current_plot = chart.plots[0]
                                    if current_plot.has_data_labels:
                                        f_task_log.write(f"        Plot 0 데이터 레이블(has_data_labels=True) 발견. Points 순회 시도...\n")
                                        if hasattr(current_plot, 'points') and current_plot.points:
                                            data_labels_in_plot_count = 0
                                            try:
                                                for pt_idx, point in enumerate(current_plot.points):
                                                    if point.has_data_label:
                                                        data_label = point.data_label
                                                        if hasattr(data_label, 'text_frame') and \
                                                           hasattr(data_label.text_frame, 'text') and \
                                                           data_label.text_frame.text:
                                                            data_label_text_val = data_label.text_frame.text
                                                            if data_label_text_val.strip():
                                                                elements_map.append({
                                                                    'type': 'chart_data_label', 'slide_idx': slide_idx,
                                                                    'shape_id': chart_shape_id_for_log, 
                                                                    'chart_shape_obj_ref': shape_in_slide,
                                                                    'plot_idx': 0, 
                                                                    'point_idx': pt_idx,
                                                                    'name': f"{chart_element_base_name}_Plot0_Point{pt_idx}_DLabel"
                                                                })
                                                                f_task_log.write(f"          -> 데이터 레이블 (Plot 0, Point {pt_idx}) 추가: '{data_label_text_val.strip()[:30]}...'\n")
                                                                data_labels_in_plot_count +=1
                                                if data_labels_in_plot_count == 0:
                                                    f_task_log.write(f"          Plot 0의 모든 Point에서 유효한 데이터 레이블 텍스트 발견 못함.\n")
                                            except Exception as e_points_iter:
                                                logger.warning(f"차트의 Points 컬렉션 순회 중 오류 ({chart_element_base_name}): {e_points_iter}. 데이터 레이블 일부 누락 가능.")
                                                f_task_log.write(f"      경고: Plot 0의 Points 컬렉션 순회 오류: {e_points_iter}\n")
                                        else:
                                            f_task_log.write(f"        Plot 0에 points 컬렉션이 없거나 비어있음.\n")
                                    else:
                                        f_task_log.write(f"        Plot 0에 데이터 레이블(has_data_labels=False) 없음.\n")
                                else:
                                    f_task_log.write(f"        차트에 플롯이 없거나 첫 번째 플롯 접근 불가.\n")

                            if item_to_add: elements_map.append(item_to_add)
                    
                    total_elements_to_translate = len(elements_map)
                    f_task_log.write(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견 (상세 분석 후).\n--- 요소 분석 완료 ---\n\n")
                    logger.info(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견 (상세 분석 후).")

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

                        if 'shape_obj_ref' in item_info: current_shape_obj = item_info['shape_obj_ref']
                        elif 'table_shape_obj_ref' in item_info: current_shape_obj = item_info['table_shape_obj_ref']
                        elif 'chart_shape_obj_ref' in item_info: current_shape_obj = item_info['chart_shape_obj_ref']
                        
                        if current_shape_obj is None:
                            f_task_log.write(f"  오류: S{slide_idx+1} ('{element_name_for_log}') 요소 객체 참조 실패. 건너뜀.\n")
                            logger.warning(f"Failed to reference shape object for '{element_name_for_log}' on slide {slide_idx+1}. Skipping.")
                            translated_count += 1
                            if progress_callback: progress_callback(slide_idx + 1, item_info['type'], translated_count, total_elements_to_translate, "요소 참조 오류")
                            continue
                        
                        current_shape_id_for_log = getattr(current_shape_obj, 'shape_id', 'N/A')
                        left_val = getattr(current_shape_obj, 'left', 0); top_val = getattr(current_shape_obj, 'top', 0)
                        width_val = getattr(current_shape_obj, 'width', 0); height_val = getattr(current_shape_obj, 'height', 0)

                        log_shape_details = (f"[S{slide_idx+1}] 요소 처리 시작: '{element_name_for_log}' (ID: {current_shape_id_for_log}), 타입: {item_info['type']}, "
                                             f"위치: L{left_val/914400:.2f}cm, T{top_val/914400:.2f}cm, W{width_val/914400:.2f}cm, H{height_val/914400:.2f}cm\n")
                        f_task_log.write(log_shape_details)
                        
                        current_text_for_progress = ""
                        item_type = item_info['type']
                        text_frame_to_process = None

                        # --- 텍스트 프레임 가져오기 ---
                        if item_type == 'text':
                            if current_shape_obj.has_text_frame: text_frame_to_process = current_shape_obj.text_frame
                        elif item_type == 'table_cell':
                            row_idx_tc = item_info.get('row_idx', -1)
                            col_idx_tc = item_info.get('col_idx', -1)
                            if current_shape_obj.has_table:
                                try: text_frame_to_process = current_shape_obj.table.cell(row_idx_tc, col_idx_tc).text_frame
                                except IndexError: logger.error(f"테이블 셀 접근 오류: {element_name_for_log}")
                        elif item_type.startswith('chart_'):
                            chart_obj_proc = current_shape_obj.chart
                            if item_type == 'chart_title':
                                if chart_obj_proc.has_title and hasattr(chart_obj_proc.chart_title, 'text_frame'):
                                    text_frame_to_process = chart_obj_proc.chart_title.text_frame
                            elif item_type == 'chart_axis_title_category':
                                if hasattr(chart_obj_proc, 'category_axis') and chart_obj_proc.category_axis.has_title and \
                                   hasattr(chart_obj_proc.category_axis.axis_title, 'text_frame'):
                                    text_frame_to_process = chart_obj_proc.category_axis.axis_title.text_frame
                            elif item_type == 'chart_axis_title_value':
                                if hasattr(chart_obj_proc, 'value_axis') and chart_obj_proc.value_axis.has_title and \
                                   hasattr(chart_obj_proc.value_axis.axis_title, 'text_frame'):
                                    text_frame_to_process = chart_obj_proc.value_axis.axis_title.text_frame
                            elif item_type == 'chart_data_label':
                                plot_idx = item_info.get('plot_idx', 0)
                                point_idx = item_info.get('point_idx')
                                if chart_obj_proc.plots and len(chart_obj_proc.plots) > plot_idx:
                                    current_plot_apply = chart_obj_proc.plots[plot_idx]
                                    if current_plot_apply.has_data_labels and hasattr(current_plot_apply, 'points') and \
                                       point_idx is not None and point_idx < len(current_plot_apply.points):
                                        point_obj_apply = current_plot_apply.points[point_idx]
                                        if point_obj_apply.has_data_label and hasattr(point_obj_apply.data_label, 'text_frame'):
                                            text_frame_to_process = point_obj_apply.data_label.text_frame
                            if not text_frame_to_process:
                                f_task_log.write(f"      경고: 차트 요소 '{element_name_for_log}'에 대한 TextFrame 가져오기 실패.\n")
                        
                        # --- 텍스트 프레임 처리 및 번역 (if/elif/else 구조 명확화) ---
                        if text_frame_to_process and hasattr(text_frame_to_process, 'text') and \
                           text_frame_to_process.text and text_frame_to_process.text.strip():
                            original_text = text_frame_to_process.text
                            current_text_for_progress = original_text
                            f_task_log.write(f"    번역 대상 텍스트 발견 ('{element_name_for_log}'): \"{original_text.strip()[:30]}...\"\n")
                            
                            row_idx_key = item_info.get('row_idx', -1)
                            col_idx_key = item_info.get('col_idx', -1)
                            style_key_suffix_for_dict = element_name_for_log if item_type.startswith('chart_') else (row_idx_key, col_idx_key)
                            unique_key_for_style = (slide_idx, id(current_shape_obj), style_key_suffix_for_dict)

                            if unique_key_for_style not in original_shape_paragraph_styles:
                                collected_para_styles = []
                                for para_in_tf in text_frame_to_process.paragraphs:
                                    para_default_font_style = self._get_style_properties(para_in_tf.font)
                                    runs_data = []
                                    if para_in_tf.runs:
                                        runs_data = [{'text': r.text, 'style': self._get_text_style(r)} for r in para_in_tf.runs]
                                    elif para_in_tf.text and para_in_tf.text.strip():
                                         run_style_as_para_default = para_default_font_style.copy()
                                         run_style_as_para_default['hyperlink_address'] = None
                                         runs_data = [{'text': para_in_tf.text, 'style': run_style_as_para_default}]
                                    collected_para_styles.append({
                                        'runs': runs_data, 'alignment': para_in_tf.alignment, 'level': para_in_tf.level,
                                        'space_before': para_in_tf.space_before, 'space_after': para_in_tf.space_after,
                                        'line_spacing': para_in_tf.line_spacing,
                                        'paragraph_default_run_style': para_default_font_style
                                    })
                                original_shape_paragraph_styles[unique_key_for_style] = collected_para_styles
                                f_task_log.write(f"      '{element_name_for_log}'의 원본 스타일 저장됨.\n")

                            if should_skip_translation(original_text):
                                f_task_log.write(f"      [스킵됨 - 번역 불필요]\n")
                            else: # 번역 시도
                                translated_text = translator.translate_text(original_text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=False)
                                log_translated_text_snippet = translated_text.replace(chr(10), ' / ').strip()[:100]
                                f_task_log.write(f"      [번역 전] \"{original_text.strip()[:30]}...\" -> [번역 후] \"{log_translated_text_snippet}...\"\n")

                                if "오류:" not in translated_text:
                                    stored_para_infos_for_item = original_shape_paragraph_styles.get(unique_key_for_style, [])
                                    original_tf_auto_size = getattr(text_frame_to_process, 'auto_size', None)
                                    original_tf_word_wrap = getattr(text_frame_to_process, 'word_wrap', None)
                                    
                                    if original_tf_auto_size is not None and original_tf_auto_size != MSO_AUTO_SIZE.NONE:
                                        text_frame_to_process.auto_size = MSO_AUTO_SIZE.NONE
                                        if original_tf_word_wrap is not None: text_frame_to_process.word_wrap = True

                                    text_frame_to_process.clear() 
                                    if hasattr(text_frame_to_process, '_element') and text_frame_to_process._element is not None:
                                        txBody = text_frame_to_process._element
                                        children_to_remove = [child for child in txBody if child.tag.endswith('}p')]
                                        if children_to_remove:
                                            f_task_log.write(f"        (강화된 정리) XML 레벨에서 기존 단락 {len(children_to_remove)}개 제거 시도...\n")
                                            for child_xml_to_remove in children_to_remove: txBody.remove(child_xml_to_remove)
                                    
                                    lines = translated_text.splitlines()
                                    processed_lines = []
                                    if lines:
                                        for line_idx_for_proc, line_content_raw in enumerate(lines):
                                            stripped_line = line_content_raw.strip()
                                            if stripped_line: processed_lines.append(line_content_raw)
                                            elif processed_lines or line_idx_for_proc < len(lines) -1 : processed_lines.append("")
                                    if not processed_lines: processed_lines = [" "]
                                    lines_to_apply = processed_lines

                                    for i, line_content_to_apply in enumerate(lines_to_apply):
                                        p_new = text_frame_to_process.add_paragraph()
                                        para_template_apply = stored_para_infos_for_item[min(i, len(stored_para_infos_for_item)-1)] if stored_para_infos_for_item else None
                                        if para_template_apply:
                                            if para_template_apply.get('alignment') is not None: p_new.alignment = para_template_apply['alignment']
                                            p_new.level = para_template_apply.get('level', 0)
                                            if para_template_apply.get('space_before') is not None: p_new.space_before = para_template_apply['space_before']
                                            if para_template_apply.get('space_after') is not None: p_new.space_after = para_template_apply['space_after']
                                            if para_template_apply.get('line_spacing') is not None: p_new.line_spacing = para_template_apply['line_spacing']
                                            if 'paragraph_default_run_style' in para_template_apply:
                                                self._apply_style_properties(p_new.font, para_template_apply['paragraph_default_run_style'])
                                        run_new = p_new.add_run()
                                        run_new.text = line_content_to_apply if line_content_to_apply.strip() else " "
                                        if not run_new.text.strip() and run_new.text != "": run_new.text = " " # 순수 공백만 있을 경우 " "

                                        if para_template_apply and para_template_apply.get('runs') and para_template_apply['runs']:
                                            self._apply_text_style(run_new, para_template_apply['runs'][0]['style'])
                                        elif para_template_apply and 'paragraph_default_run_style' in para_template_apply:
                                            self._apply_style_properties(run_new.font, para_template_apply['paragraph_default_run_style'])
                                    
                                    if original_tf_auto_size is not None and original_tf_auto_size != MSO_AUTO_SIZE.NONE:
                                        text_frame_to_process.auto_size = original_tf_auto_size
                                    if original_tf_word_wrap is not None and text_frame_to_process.word_wrap != original_tf_word_wrap :
                                        text_frame_to_process.word_wrap = original_tf_word_wrap
                                    f_task_log.write(f"        '{element_name_for_log}' 번역된 텍스트 적용 완료.\n")
                                else: # "오류:" in translated_text
                                    f_task_log.write(f"      -> 텍스트 번역 실패 또는 빈 결과: {translated_text}\n")
                        
                        elif item_type == 'image' and ocr_handler:
                            # ... (이전과 동일한 OCR 이미지 처리 로직) ...
                            # (구문적으로는 이전 제공 버전과 동일하게 유지, 내용 생략)
                            if current_shape_obj is None or not hasattr(current_shape_obj, 'image') or not hasattr(current_shape_obj.image, 'blob'):
                                f_task_log.write(f"    경고: '{element_name_for_log}' (ID:{getattr(current_shape_obj, 'shape_id', 'N/A')}) 이미지 객체 또는 이미지 데이터 접근 불가. 건너뜀.\n")
                                logger.warning(f"Cannot access image data for '{element_name_for_log}' (ID:{getattr(current_shape_obj, 'shape_id', 'N/A')}). Skipping.")
                            else:
                                try:
                                    image_bytes_io = io.BytesIO(current_shape_obj.image.blob)
                                    with Image.open(image_bytes_io) as img_pil_for_ocr_source:
                                        img_pil_for_ocr = img_pil_for_ocr_source.convert("RGB")
                                        try:
                                            hasher = hashlib.md5(); img_byte_arr_for_hash = io.BytesIO()
                                            img_pil_for_ocr.save(img_byte_arr_for_hash, format='PNG')
                                            hasher.update(img_byte_arr_for_hash.getvalue()); img_hash_val = hasher.hexdigest()
                                            f_task_log.write(f"    OCR 대상 RGB 이미지 해시 (MD5): {img_hash_val}\n")
                                        except Exception as e_hash_ocr: logger.warning(f"    이미지 해시 생성 중 오류: {e_hash_ocr}")

                                        ocr_results = ocr_handler.ocr_image(img_pil_for_ocr)
                                        if ocr_results:
                                            f_task_log.write(f"      이미지 내 OCR 텍스트 {len(ocr_results)}개 발견.\n")
                                            image_bytes_io.seek(0) 
                                            with Image.open(image_bytes_io) as edited_image_pil_base:
                                                edited_image_pil = edited_image_pil_base.copy()
                                                any_ocr_text_translated_and_rendered = False
                                                for ocr_idx, ocr_item in enumerate(ocr_results):
                                                    if not (isinstance(ocr_item, (list, tuple)) and len(ocr_item) >= 2):
                                                        f_task_log.write(f"        잘못된 OCR 결과 항목 형식. 건너뜀.\n"); continue
                                                    box_ocr, text_conf_tuple_ocr = ocr_item[0], ocr_item[1]
                                                    ocr_angle_val = ocr_item[2] if len(ocr_item) > 2 else None
                                                    if not (isinstance(text_conf_tuple_ocr, (list, tuple)) and len(text_conf_tuple_ocr) == 2):
                                                        f_task_log.write(f"        잘못된 OCR 텍스트/신뢰도 튜플 형식. 건너뜀.\n"); continue
                                                    text_ocr, confidence_ocr = text_conf_tuple_ocr
                                                    if stop_event and stop_event.is_set(): break
                                                    current_text_for_progress = text_ocr
                                                    f_task_log.write(f"        OCR Text [{ocr_idx+1}]: \"{text_ocr.strip()[:30]}...\" (Conf: {confidence_ocr:.2f}, Angle: {ocr_angle_val})\n")
                                                    if not is_ocr_text_valid(text_ocr):
                                                        f_task_log.write(f"          -> [스킵됨 - OCR 유효성 낮음]\n"); continue
                                                    if should_skip_translation(text_ocr):
                                                        f_task_log.write(f"          -> [스킵됨 - 번역 불필요]\n"); continue
                                                    translated_ocr_text = translator.translate_text(text_ocr, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=True)
                                                    f_task_log.write(f"          -> 번역 결과: \"{translated_ocr_text.strip()[:30]}...\"\n")
                                                    if "오류:" not in translated_ocr_text and translated_ocr_text.strip():
                                                        try:
                                                            edited_image_pil = ocr_handler.render_translated_text_on_image(
                                                                edited_image_pil, box_ocr, translated_ocr_text,
                                                                font_code_for_render=font_code_for_render, 
                                                                original_text=text_ocr, ocr_angle=ocr_angle_val
                                                            )
                                                            any_ocr_text_translated_and_rendered = True
                                                            f_task_log.write(f"            -> 렌더링 완료.\n")
                                                        except Exception as e_render_ocr:
                                                            f_task_log.write(f"            오류: OCR 텍스트 렌더링 실패: {e_render_ocr}\n")
                                                            logger.error(f"OCR 렌더링 실패: {e_render_ocr}", exc_info=True)
                                                    else:
                                                         f_task_log.write(f"            -> 번역 실패 또는 빈 결과로 렌더링 안함.\n")
                                                if stop_event and stop_event.is_set(): break
                                                if any_ocr_text_translated_and_rendered:
                                                    image_stream_replace = io.BytesIO()
                                                    save_fmt_ocr = edited_image_pil_base.format if edited_image_pil_base.format and edited_image_pil_base.format.upper() in ['JPEG', 'PNG', 'GIF', 'BMP', 'TIFF'] else 'PNG'
                                                    edited_image_pil.save(image_stream_replace, format=save_fmt_ocr)
                                                    image_stream_replace.seek(0)
                                                    left_pic, top_pic, width_pic, height_pic = current_shape_obj.left, current_shape_obj.top, current_shape_obj.width, current_shape_obj.height
                                                    current_slide_obj_for_pic = prs.slides[slide_idx]
                                                    pic_elem_xml = current_shape_obj.element
                                                    pic_parent_xml = pic_elem_xml.getparent()
                                                    if pic_parent_xml is not None: pic_parent_xml.remove(pic_elem_xml)
                                                    new_pic_shape_obj = current_slide_obj_for_pic.shapes.add_picture(image_stream_replace, left_pic, top_pic, width=width_pic, height=height_pic)
                                                    f_task_log.write(f"      이미지 '{element_name_for_log}' 교체 완료.\n")
                                                else: f_task_log.write(f"      이미지 '{element_name_for_log}' 변경 없음.\n")
                                        else: f_task_log.write(f"      이미지 '{element_name_for_log}' 내 OCR 텍스트 발견되지 않음.\n")
                                except AttributeError as e_attr_ocr_img:
                                    f_task_log.write(f"    경고: '{element_name_for_log}' 이미지 처리 중 속성 오류: {e_attr_ocr_img}. 건너뜀.\n")
                                    logger.warning(f"Attribute error processing image '{element_name_for_log}': {e_attr_ocr_img}", exc_info=True)
                                except IOError as e_io_ocr_img:
                                    f_task_log.write(f"    이미지 처리 중 I/O 오류 '{element_name_for_log}': {e_io_ocr_img}. 건너뜀.\n")
                                    logger.error(f"I/O error during image processing '{element_name_for_log}': {e_io_ocr_img}", exc_info=False)
                                except Exception as e_img_proc_ocr_detail:
                                    f_task_log.write(f"    이미지 '{element_name_for_log}' 처리 중 예기치 않은 오류: {e_img_proc_ocr_detail}. 건너뜀.\n")
                                    logger.error(f"Unexpected error processing image '{element_name_for_log}': {e_img_proc_ocr_detail}", exc_info=True)
                        else: # 텍스트, 테이블셀, 차트, 이미지가 아닌 다른 타입의 요소 처리 (또는 아무것도 안 함)
                            f_task_log.write(f"    '{element_name_for_log}' (타입: {item_type})은 현재 번역 처리 대상 아님. 건너뜀.\n")
                            logger.debug(f"Skipping shape '{element_name_for_log}' of type '{item_type}' as it's not targeted for translation processing.")


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