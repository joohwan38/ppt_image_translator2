# pptx_handler.py
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.enum.lang import MSO_LANGUAGE_ID

import os
import io
import logging
import re
from datetime import datetime
import hashlib
import traceback
from PIL import Image
import tempfile
import shutil

# 설정 파일 import
import config

from typing import TYPE_CHECKING, Optional, Dict, Any, List, Tuple, Callable

if TYPE_CHECKING:
    from translator import OllamaTranslator
    from ollama_service import OllamaService
    from ocr_handler import BaseOcrHandler # BaseOcrHandler로 변경

logger = logging.getLogger(__name__)

MEANINGFUL_CHAR_PATTERN = re.compile(
    r'[a-zA-Z'
    r'\u00C0-\u024F'    # Latin Extended-A
    r'\u1E00-\u1EFF'    # Latin Extended Additional
    r'\u0600-\u06FF'    # Arabic
    r'\u0750-\u077F'    # Arabic Supplement
    r'\u08A0-\u08FF'    # Arabic Extended-A
    r'\u3040-\u30ff'    # Hiragana, Katakana
    r'\u3131-\uD79D'    # Hangul Compatibility Jamo, Hangul Syllables
    r'\u4e00-\u9fff'    # CJK Unified Ideographs
    r'\u0E00-\u0E7F'    # Thai
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
    if text_len <= 3: # 3글자 이하도 의미 있으면 번역 시도
        if MEANINGFUL_CHAR_PATTERN.search(stripped_text): # 특히 중국어, 일본어 등
            logger.debug(f"번역 시도 (짧은 문자열, 의미 문자 포함): '{stripped_text}'")
            return False
        else:
            logger.debug(f"번역 스킵 (짧고 의미 있는 문자 없음): '{stripped_text}'")
            return True

    meaningful_chars_count = len(MEANINGFUL_CHAR_PATTERN.findall(stripped_text))
    ratio = meaningful_chars_count / text_len
    if ratio < config.MIN_MEANINGFUL_CHAR_RATIO_SKIP:
        logger.debug(f"번역 스킵 (의미 문자 비율 낮음 {ratio:.2f}, 임계값: {config.MIN_MEANINGFUL_CHAR_RATIO_SKIP}): '{stripped_text[:50]}...'")
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
    if text_len <= 2: # OCR은 2글자 이하도 의미 있으면 유효 처리
        if MEANINGFUL_CHAR_PATTERN.search(stripped_text):
             logger.debug(f"OCR 유효 (매우 짧은 문자열, 의미 문자 포함): '{stripped_text}'")
             return True
        else:
            logger.debug(f"OCR 유효성 스킵 (매우 짧고 의미 있는 문자 없음): '{stripped_text}'")
            return False

    meaningful_chars_count = len(MEANINGFUL_CHAR_PATTERN.findall(stripped_text))
    ratio = meaningful_chars_count / text_len
    if ratio < config.MIN_MEANINGFUL_CHAR_RATIO_OCR:
        logger.debug(f"OCR 유효성 스킵 (의미 문자 비율 낮음 {ratio:.2f}, 임계값: {config.MIN_MEANINGFUL_CHAR_RATIO_OCR}): '{stripped_text[:50]}...'")
        return False
        
    logger.debug(f"OCR 유효 (조건 통과): '{stripped_text[:50]}...'")
    return True


class PptxHandler:
    def __init__(self):
        pass

    def get_file_info(self, file_path: str) -> Dict[str, int]:
        logger.info(f"파일 정보 분석 시작: {file_path}")
        info = {
            "slide_count": 0,
            "text_elements_count": 0, # UI에서 직접 사용 안해도 내부적으론 카운트 가능
            "total_text_char_count": 0, # 가중치 계산에 사용
            "image_elements_count": 0, # 가중치 계산에 사용
            "chart_elements_count": 0  # 가중치 계산에 사용
        }
        try:
            prs = Presentation(file_path)
            info["slide_count"] = len(prs.slides)
            for slide in prs.slides:
                for shape in slide.shapes:
                    is_counted_as_text_element_for_shape = False
                    if shape.has_text_frame and hasattr(shape.text_frame, 'text') and \
                       shape.text_frame.text and shape.text_frame.text.strip():
                        text_content = shape.text_frame.text
                        if not should_skip_translation(text_content):
                            if not is_counted_as_text_element_for_shape:
                                info["text_elements_count"] += 1
                                is_counted_as_text_element_for_shape = True
                            info["total_text_char_count"] += len(text_content)

                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        info["image_elements_count"] += 1
                    elif shape.has_table:
                        for r_idx, row in enumerate(shape.table.rows):
                            for c_idx, cell in enumerate(row.cells):
                                if hasattr(cell.text_frame, 'text') and cell.text_frame.text and cell.text_frame.text.strip():
                                    text_content = cell.text_frame.text
                                    if not should_skip_translation(text_content):
                                        info["text_elements_count"] += 1 # 테이블 셀은 각 유효 텍스트마다 요소 수 증가
                                        info["total_text_char_count"] += len(text_content)
                    elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                        info["chart_elements_count"] +=1
            logger.info(
                f"파일 분석 완료: Slides:{info['slide_count']}, "
                f"TextElements(internal):{info['text_elements_count']} (TotalChars:{info['total_text_char_count']}), "
                f"Images:{info['image_elements_count']}, Charts:{info['chart_elements_count']}"
            )
        except Exception as e:
            logger.error(f"'{os.path.basename(file_path)}' 파일 정보 분석 오류: {e}", exc_info=True)
            # 오류 발생 시 info 딕셔너리는 초기값으로 반환될 것임
        return info

    def _get_style_properties(self, font_object) -> Dict[str, Any]:
        if font_object is None:
            return {}
        
        style_props: Dict[str, Any] = {
            'name': None, 'size': None, 'bold': None, 'italic': None,
            'underline': None, 'color_rgb': None, 'color_theme_index': None,
            'color_brightness': 0.0, 'language_id': None
        }
        
        try: style_props['name'] = font_object.name
        except AttributeError: pass
        try: style_props['size'] = font_object.size
        except AttributeError: pass
        try: style_props['bold'] = font_object.bold
        except AttributeError: pass
        try: style_props['italic'] = font_object.italic
        except AttributeError: pass
        try: style_props['underline'] = font_object.underline
        except AttributeError: pass
        
        try:
            style_props['language_id'] = font_object.language_id
        except (ValueError, AttributeError) as e_lang:
            # logger.debug(f"Font language_id 가져오기 실패: {e_lang}. 기본값 사용.") # 너무 많은 로그 발생 가능성
            pass


        if hasattr(font_object, 'color') and hasattr(font_object.color, 'type'):
            color_type = font_object.color.type
            if color_type == MSO_COLOR_TYPE.RGB:
                try: style_props['color_rgb'] = tuple(font_object.color.rgb)
                except AttributeError: pass
            elif color_type == MSO_COLOR_TYPE.SCHEME:
                try: style_props['color_theme_index'] = font_object.color.theme_color
                except AttributeError: pass
                try: style_props['color_brightness'] = font_object.color.brightness
                except AttributeError: style_props['color_brightness'] = 0.0 # 기본값 0.0
        
        return style_props

    def _get_text_style(self, run) -> Dict[str, Any]:
        style = self._get_style_properties(run.font)
        try:
            style['hyperlink_address'] = run.hyperlink.address if run.hyperlink and run.hyperlink.address else None
        except AttributeError:
            style['hyperlink_address'] = None
        return style

    def _apply_style_properties(self, target_font_object, style_dict_to_apply: Dict[str, Any]):
        if not style_dict_to_apply or target_font_object is None:
            return
        
        font = target_font_object
        
        if style_dict_to_apply.get('name') is not None: font.name = style_dict_to_apply['name']
        if style_dict_to_apply.get('size') is not None: font.size = style_dict_to_apply['size']
        if style_dict_to_apply.get('bold') is not None: font.bold = style_dict_to_apply['bold']
        if style_dict_to_apply.get('italic') is not None: font.italic = style_dict_to_apply['italic']
        if style_dict_to_apply.get('underline') is not None: font.underline = style_dict_to_apply['underline']

        applied_color = False
        if style_dict_to_apply.get('color_rgb') is not None:
            try:
                rgb_tuple = tuple(int(c) for c in style_dict_to_apply['color_rgb'])
                if len(rgb_tuple) == 3:
                    font.color.rgb = RGBColor(*rgb_tuple)
                    applied_color = True
            except Exception as e: logger.warning(f"RGB 색상 {style_dict_to_apply['color_rgb']} 적용 실패: {e}")
        
        if not applied_color and style_dict_to_apply.get('color_theme_index') is not None:
            try:
                # MSO_THEME_COLOR_INDEX enum 값으로 변환 필요
                theme_color_val = style_dict_to_apply['color_theme_index']
                if isinstance(theme_color_val, MSO_THEME_COLOR_INDEX):
                    font.color.theme_color = theme_color_val
                elif isinstance(theme_color_val, int) and theme_color_val in [item.value for item in MSO_THEME_COLOR_INDEX]: # 정수 값으로 올 경우
                    font.color.theme_color = MSO_THEME_COLOR_INDEX(theme_color_val)
                else:
                    logger.warning(f"유효하지 않은 테마 색상 인덱스: {theme_color_val}")


                brightness_val = style_dict_to_apply.get('color_brightness', 0.0)
                font.color.brightness = max(-1.0, min(1.0, float(brightness_val)))
            except Exception as e: logger.warning(f"테마 색상 {style_dict_to_apply.get('color_theme_index')} 적용 실패: {e}")
        
        if style_dict_to_apply.get('language_id') is not None:
            try:
                lang_id_val = style_dict_to_apply['language_id']
                if isinstance(lang_id_val, MSO_LANGUAGE_ID):
                    font.language_id = lang_id_val
                elif isinstance(lang_id_val, int): # 정수 값으로 올 경우 (예: 1042 for Korean)
                    try:
                        font.language_id = MSO_LANGUAGE_ID(lang_id_val)
                    except ValueError: # MSO_LANGUAGE_ID enum에 없는 값일 경우
                        logger.warning(f"MSO_LANGUAGE_ID에 없는 언어 ID: {lang_id_val}. 기본값 사용.")
                        # font.language_id = MSO_LANGUAGE_ID.NONE # 또는 설정 안 함
                else: # 문자열 등으로 올 경우 (예: "ko-KR") - python-pptx는 직접 지원 안 함
                    logger.debug(f"문자열 언어 ID '{lang_id_val}'는 직접 적용 불가. 무시됨.")

            except Exception as e_lang: logger.warning(f"language_id '{style_dict_to_apply['language_id']}' 적용 실패: {e_lang}")


    def _apply_text_style(self, run, style_to_apply: Dict[str, Any]):
        if not style_to_apply or run is None:
            return
        self._apply_style_properties(run.font, style_to_apply)
        if style_to_apply.get('hyperlink_address'):
            try:
                # 기존 run.hyperlink.address = ... 방식이 가끔 오류 발생 가능성 있음
                # 더 안전한 방법: run.hyperlink 객체를 새로 만들거나, 기존 객체에 address 설정
                hlink = run.hyperlink
                if hlink is None: # 하이퍼링크가 없었으면 새로 추가
                    # run.hyperlink는 읽기 전용 속성. paragraph에 추가해야 함.
                    # 이 부분은 복잡. 일단 기존 링크가 있을 때만 주소 변경.
                    # para = run.part.rels[run.rPr.xpath('./a:hlinkClick')[0].get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')].target_part.part_related_length
                    # 위는 너무 복잡. 단순하게 기존 링크가 있을 때만 주소 변경
                    pass # 새 링크 추가는 복잡하므로 여기서는 생략.
                
                # 기존 하이퍼링크가 있다면 주소 변경 시도
                # 주의: run.hyperlink.address = '...' 할당은 python-pptx에서 직접 지원하지 않을 수 있음.
                # _ relación (rId)을 변경해야 할 수 있음.
                # 여기서는 스타일 복원 시 기존 링크가 있었음을 가정하고,
                # 새 텍스트에도 동일한 링크 주소를 (논리적으로) 유지하려는 의도로 해석.
                # 실제로는 번역된 텍스트의 run에 새 하이퍼링크를 걸어야 한다면 다른 접근 필요.
                # 지금은 스타일 복원 목적이므로, 이 부분은 원본 스타일의 'hyperlink_address' 정보를
                # 가지고 있다가, 새 run 생성 시 활용하는 로직으로 변경되어야 함.
                # 여기서는 일단 주석처리. (스타일 적용 시 하이퍼링크는 복잡)
                # if run.hyperlink:
                #    run.hyperlink.address = style_to_apply['hyperlink_address']
                pass
            except Exception as e: logger.warning(f"Run에 하이퍼링크 주소 적용 시도 중 오류 (무시): {e}")


    def translate_presentation_stage1(self, prs: Presentation, src_lang_ui_name: str, tgt_lang_ui_name: str,
                                      translator: 'OllamaTranslator', ocr_handler: Optional['BaseOcrHandler'],
                                      model_name: str, ollama_service: 'OllamaService',
                                      font_code_for_render: str, task_log_filepath: str,
                                      progress_callback_item_completed: Optional[Callable[[Any, str, int, str], None]] = None,
                                      stop_event: Optional[Any] = None,
                                      image_translation_enabled: bool = True, # 이미지 번역 활성화 여부
                                      ocr_temperature: Optional[float] = None # OCR 번역 온도
                                      ) -> bool:
        
        with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
            f_task_log.write("--- 1단계: 차트 외 요소 번역 시작 (translate_presentation_stage1) ---\n")
            logger.info("1단계: 차트 외 요소 (텍스트 상자, 표, OCR 등) 번역 중...")

            elements_to_process_stage1: List[Dict[str, Any]] = []
            for slide_idx, slide in enumerate(prs.slides):
                if stop_event and stop_event.is_set(): break
                for shape_idx, shape in enumerate(slide.shapes):
                    if stop_event and stop_event.is_set(): break
                    
                    shape_id = getattr(shape, 'shape_id', f"slide{slide_idx}_shape{shape_idx}")
                    element_name = shape.name or f"S{slide_idx+1}_Shape{shape_idx}_Id{shape_id}"
                    item_base_info = {'slide_idx': slide_idx, 'shape_obj_ref': shape, 'name': element_name}
                    item_progress_type = "Unknown" # 진행 콜백용 타입

                    if shape.shape_type == MSO_SHAPE_TYPE.CHART:
                        continue

                    if shape.has_text_frame and hasattr(shape.text_frame, 'text') and \
                       shape.text_frame.text and shape.text_frame.text.strip():
                        original_text = shape.text_frame.text
                        char_count = 0
                        if not should_skip_translation(original_text):
                            char_count = len(original_text)
                        item_progress_type = "텍스트"
                        elements_to_process_stage1.append({**item_base_info, 'type': 'text_shape', # 타입 명확화
                                                           'original_text': original_text, 'char_count': char_count,
                                                           'progress_type': item_progress_type})
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        if image_translation_enabled and ocr_handler: # 이미지 번역 옵션 및 핸들러 확인
                            item_progress_type = "이미지 OCR"
                            elements_to_process_stage1.append({**item_base_info, 'type': 'image',
                                                               'progress_type': item_progress_type})
                    elif shape.has_table:
                        item_progress_type = "표 내부 텍스트"
                        for r_idx, row in enumerate(shape.table.rows):
                            for c_idx, cell in enumerate(row.cells):
                                if hasattr(cell.text_frame, 'text') and cell.text_frame.text and cell.text_frame.text.strip():
                                    original_text = cell.text_frame.text
                                    char_count = 0
                                    if not should_skip_translation(original_text):
                                        char_count = len(original_text)
                                    
                                    elements_to_process_stage1.append({
                                        'slide_idx': slide_idx,
                                        'table_shape_obj_ref': shape,
                                        'name': f"{element_name}_R{r_idx}C{c_idx}",
                                        'type': 'table_cell',
                                        'row_idx': r_idx, 'col_idx': c_idx,
                                        'original_text': original_text, 'char_count': char_count,
                                        'progress_type': item_progress_type
                                    })
            
            f_task_log.write(f"1단계 처리 대상 요소 (차트 제외, 맵핑 완료): {len(elements_to_process_stage1)}개.\n")
            logger.info(f"1단계 처리 대상 요소 (차트 제외, 맵핑 완료): {len(elements_to_process_stage1)}개.")

            if not elements_to_process_stage1:
                 is_any_chart_present = any(s.shape_type == MSO_SHAPE_TYPE.CHART for slide in prs.slides for s in slide.shapes)
                 if not is_any_chart_present:
                    msg = "1단계 번역/처리 대상 요소(텍스트/이미지/표)가 없고, 차트도 없어 1단계 처리 스킵."
                    f_task_log.write(msg + "\n")
                    logger.info(msg)
                 return True # 처리할 내용 없으면 성공으로 간주

            original_paragraph_styles_stage1: Dict[Tuple[int, Any, Any], List[Dict[str, Any]]] = {}

            for item_data in elements_to_process_stage1:
                if stop_event and stop_event.is_set():
                    f_task_log.write(f"1단계 처리 중 중단 요청 감지.\n")
                    return False # 중단 시 실패 반환

                slide_idx = item_data['slide_idx']
                item_name_log = item_data['name']
                item_type_internal = item_data['type'] # 내부 로직용 타입 ('text_shape', 'image', 'table_cell')
                item_progress_type_for_ui = item_data['progress_type'] # UI 표시용 타입

                weighted_work_for_this_item = 0
                current_text_for_progress_update = item_data.get('original_text', '')[:30] if item_type_internal in ['text_shape', 'table_cell'] else "[이미지 처리]"
                
                current_shape_object = item_data.get('shape_obj_ref') or item_data.get('table_shape_obj_ref')
                if not current_shape_object:
                    err_msg_no_ref = f"  오류 (1단계): S{slide_idx+1} ('{item_name_log}') 요소 객체 참조 실패. 건너뜀.\n"
                    f_task_log.write(err_msg_no_ref)
                    logger.error(err_msg_no_ref.strip())
                    if progress_callback_item_completed:
                        progress_callback_item_completed(slide_idx + 1, "오류", 0, "요소 참조 불가")
                    continue
                        
                shape_id_log = getattr(current_shape_object, 'shape_id', 'N/A_ID')
                f_task_log.write(f"  [1단계 S{slide_idx+1}] 처리 시작: '{item_name_log}' (ID: {shape_id_log}), 타입: {item_type_internal}\n")

                text_frame_for_processing: Optional[Any] = None
                original_full_text_from_item = item_data.get('original_text', "")

                if item_type_internal == 'text_shape':
                    if current_shape_object.has_text_frame:
                        text_frame_for_processing = current_shape_object.text_frame
                    weighted_work_for_this_item = item_data.get('char_count', 0) * config.WEIGHT_TEXT_CHAR

                elif item_type_internal == 'table_cell':
                    row_idx = item_data['row_idx']
                    col_idx = item_data['col_idx']
                    if current_shape_object.has_table:
                        try:
                            text_frame_for_processing = current_shape_object.table.cell(row_idx, col_idx).text_frame
                        except IndexError:
                            err_msg_tbl_idx = f"1단계 테이블 셀 접근 오류 (IndexError): {item_name_log} at R{row_idx}C{col_idx}"
                            logger.error(err_msg_tbl_idx)
                            f_task_log.write(f"    오류: {err_msg_tbl_idx}. 건너뜀.\n")
                            if progress_callback_item_completed:
                                progress_callback_item_completed(slide_idx + 1, "오류", 0, f"테이블 셀 접근 실패 R{row_idx}C{col_idx}")
                            continue # 다음 아이템으로
                    weighted_work_for_this_item = item_data.get('char_count', 0) * config.WEIGHT_TEXT_CHAR
                
                # 텍스트 프레임 처리 (text_shape, table_cell 공통)
                if text_frame_for_processing and original_full_text_from_item:
                    if item_data.get('char_count', 0) == 0 :
                        f_task_log.write(f"      [1단계 스킵됨 - 번역 불필요 또는 의미 없는 텍스트]\n")
                    else:
                        f_task_log.write(f"    1단계 번역 대상 텍스트 발견 (길이: {item_data.get('char_count',0)}): \"{original_full_text_from_item.strip()[:50]}...\"\n")
                        
                        style_key_suffix_val = (item_data.get('row_idx', -1), item_data.get('col_idx', -1)) if item_type_internal == 'table_cell' else 'shape_text'
                        style_unique_key = (slide_idx, id(current_shape_object), style_key_suffix_val)

                        if style_unique_key not in original_paragraph_styles_stage1:
                            para_styles_collected: List[Dict[str, Any]] = []
                            for para_obj in text_frame_for_processing.paragraphs:
                                para_default_font_style = self._get_style_properties(para_obj.font)
                                runs_info: List[Dict[str, Any]] = []
                                if para_obj.runs:
                                    for run_obj in para_obj.runs:
                                        runs_info.append({'text': run_obj.text, 'style': self._get_text_style(run_obj)})
                                elif para_obj.text and para_obj.text.strip():
                                    # Run이 없는 단락의 경우, 단락 스타일을 run 스타일로 간주
                                    run_style_from_para = para_default_font_style.copy()
                                    run_style_from_para['hyperlink_address'] = None # 단락에는 하이퍼링크 직접 설정 안됨
                                    runs_info.append({'text': para_obj.text, 'style': run_style_from_para})
                                
                                para_styles_collected.append({
                                    'runs': runs_info,
                                    'alignment': para_obj.alignment,
                                    'level': para_obj.level,
                                    'space_before': para_obj.space_before,
                                    'space_after': para_obj.space_after,
                                    'line_spacing': para_obj.line_spacing,
                                    'paragraph_default_run_style': para_default_font_style
                                })
                            original_paragraph_styles_stage1[style_unique_key] = para_styles_collected
                            f_task_log.write(f"      '{item_name_log}'의 원본 단락 스타일 저장 (1단계).\n")
                        
                        translated_text_content = translator.translate_text(
                            original_full_text_from_item, src_lang_ui_name, tgt_lang_ui_name,
                            model_name, ollama_service, is_ocr_text=False # 일반 텍스트이므로 is_ocr_text=False
                        )
                        log_trans_text_snippet = translated_text_content.replace('\n', ' / ').strip()[:100]
                        f_task_log.write(f"      [1단계 번역 전] \"{original_full_text_from_item.strip()[:50]}...\" -> [1단계 번역 후] \"{log_trans_text_snippet}...\"\n")

                        if "오류:" not in translated_text_content and translated_text_content.strip():
                            stored_paras_info_apply = original_paragraph_styles_stage1.get(style_unique_key, [])
                            
                            # 텍스트 프레임의 기존 속성 저장
                            original_tf_auto_sz = getattr(text_frame_for_processing, 'auto_size', None)
                            original_tf_word_wrp = getattr(text_frame_for_processing, 'word_wrap', None)
                            original_tf_v_anchor = getattr(text_frame_for_processing, 'vertical_anchor', None)
                            original_tf_margins = {
                                'left': getattr(text_frame_for_processing, 'margin_left', None),
                                'right': getattr(text_frame_for_processing, 'margin_right', None),
                                'top': getattr(text_frame_for_processing, 'margin_top', None),
                                'bottom': getattr(text_frame_for_processing, 'margin_bottom', None),
                            }

                            # 자동 크기 조정을 임시로 비활성화 (텍스트 삽입 중 크기 변경 방지)
                            if original_tf_auto_sz is not None and original_tf_auto_sz != MSO_AUTO_SIZE.NONE:
                                text_frame_for_processing.auto_size = MSO_AUTO_SIZE.NONE
                            if original_tf_word_wrp is not None: text_frame_for_processing.word_wrap = True # 자동 줄 바꿈 활성화

                            text_frame_for_processing.clear() # 기존 텍스트 및 단락 모두 삭제
                            # clear() 후에도 _element에 <a:p> 태그가 남아있을 수 있으므로 한 번 더 확인
                            if hasattr(text_frame_for_processing, '_element') and text_frame_for_processing._element is not None:
                                txBody_xml = text_frame_for_processing._element
                                p_tags_to_remove = [child for child in txBody_xml if child.tag.endswith('}p')]
                                if p_tags_to_remove:
                                    f_task_log.write(f"        (강화된 정리) XML에서 기존 단락 {len(p_tags_to_remove)}개 제거.\n")
                                    for p_xml_tag in p_tags_to_remove: txBody_xml.remove(p_xml_tag)
                            
                            lines_from_translation = translated_text_content.splitlines()
                            if not lines_from_translation and translated_text_content: # 번역 결과가 있지만 splitlines가 빈 리스트 반환 시
                                lines_from_translation = [translated_text_content]
                            elif not lines_from_translation: # 빈 번역 결과
                                lines_from_translation = [" "] # 빈 단락이라도 하나 만듦 (스타일 유지를 위해)

                            for line_idx, line_txt in enumerate(lines_from_translation):
                                new_para = text_frame_for_processing.add_paragraph()
                                # 원본 단락 스타일 적용 (가장 유사한 원본 단락 스타일 사용)
                                para_style_template = stored_paras_info_apply[min(line_idx, len(stored_paras_info_apply)-1)] if stored_paras_info_apply else {}
                                
                                if para_style_template.get('alignment') is not None: new_para.alignment = para_style_template['alignment']
                                else: new_para.alignment = PP_ALIGN.LEFT # 기본값

                                new_para.level = para_style_template.get('level', 0)
                                if para_style_template.get('space_before') is not None: new_para.space_before = para_style_template['space_before']
                                if para_style_template.get('space_after') is not None: new_para.space_after = para_style_template['space_after']
                                if para_style_template.get('line_spacing') is not None: new_para.line_spacing = para_style_template['line_spacing']
                                
                                # 단락 기본 폰트 스타일 적용
                                if 'paragraph_default_run_style' in para_style_template:
                                    self._apply_style_properties(new_para.font, para_style_template['paragraph_default_run_style'])

                                # Run 추가 및 스타일 적용
                                # 원본에 여러 run이 있었다면, 첫 번째 run의 스타일을 대표로 사용하거나,
                                # 번역된 텍스트 전체를 하나의 run으로 처리하고 대표 스타일 적용.
                                # 여기서는 하나의 run으로 처리하고, 원본 첫 run 또는 단락 기본 스타일 적용.
                                new_run = new_para.add_run()
                                new_run.text = line_txt if line_txt.strip() else " " # 빈 줄은 공백 하나로
                                if not new_run.text.strip() and new_run.text != " ": new_run.text = " " # 정말 빈 텍스트면 공백

                                run_style_to_apply = {}
                                if para_style_template.get('runs') and para_style_template['runs']:
                                    run_style_to_apply = para_style_template['runs'][0]['style'] # 첫 번째 run 스타일 대표 사용
                                elif 'paragraph_default_run_style' in para_style_template: # run 정보 없으면 단락 기본 사용
                                    run_style_to_apply = para_style_template['paragraph_default_run_style'].copy()
                                    run_style_to_apply['hyperlink_address'] = None # 단락 스타일에는 하이퍼링크 없음

                                if run_style_to_apply:
                                    self._apply_text_style(new_run, run_style_to_apply)
                            
                            # 텍스트 프레임 원본 속성 복원
                            if original_tf_auto_sz is not None: text_frame_for_processing.auto_size = original_tf_auto_sz
                            if original_tf_word_wrp is not None: text_frame_for_processing.word_wrap = original_tf_word_wrp
                            if original_tf_v_anchor is not None: text_frame_for_processing.vertical_anchor = original_tf_v_anchor
                            for margin_prop, val in original_tf_margins.items():
                                if val is not None: setattr(text_frame_for_processing, f"margin_{margin_prop}", val)

                            f_task_log.write(f"        '{item_name_log}' 1단계 번역된 텍스트 적용 완료.\n")
                        else:
                            f_task_log.write(f"      -> 1단계 텍스트 번역 실패 또는 빈 결과: {translated_text_content}\n")
                
                elif item_type_internal == 'image' and image_translation_enabled and ocr_handler:
                    weighted_work_for_this_item = config.WEIGHT_IMAGE
                    f_task_log.write(f"    1단계 OCR 처리 시도: '{item_name_log}' (가중치: {weighted_work_for_this_item})\n")
                    current_text_for_progress_update = "[이미지 OCR 처리 중]"
                    # item_progress_type_for_ui는 이미 "이미지 OCR"로 설정됨
                    
                    if current_shape_object is None or not hasattr(current_shape_object, 'image') or \
                       not hasattr(current_shape_object.image, 'blob'):
                        warn_msg_no_blob = f"      경고 (1단계): '{item_name_log}' 이미지 객체 또는 blob 데이터 접근 불가. 건너뜀.\n"
                        f_task_log.write(warn_msg_no_blob)
                        logger.warning(warn_msg_no_blob.strip())
                        if progress_callback_item_completed: # 오류 상황도 콜백으로 알림
                            progress_callback_item_completed(slide_idx + 1, "오류", weighted_work_for_this_item, "이미지 데이터 접근 불가")
                        continue # 다음 아이템으로
                    else:
                        try:
                            img_bytes_io = io.BytesIO(current_shape_object.image.blob)
                            with Image.open(img_bytes_io) as img_pil_original_ocr:
                                img_pil_rgb_ocr = img_pil_original_ocr.convert("RGB")
                                
                                try: # 이미지 해시 로깅 (디버깅용)
                                    hasher_md5 = hashlib.md5()
                                    img_byte_arr_for_hashing = io.BytesIO()
                                    img_pil_rgb_ocr.save(img_byte_arr_for_hashing, format='PNG')
                                    hasher_md5.update(img_byte_arr_for_hashing.getvalue())
                                    img_hash_val_log = hasher_md5.hexdigest()
                                    f_task_log.write(f"      OCR 대상 RGB 이미지 MD5 해시: {img_hash_val_log}\n")
                                except Exception as e_hash: logger.debug(f"이미지 해시 생성 중 오류: {e_hash}")

                                ocr_results_list = ocr_handler.ocr_image(img_pil_rgb_ocr)
                                
                                if ocr_results_list:
                                    f_task_log.write(f"        이미지 내 OCR 텍스트 {len(ocr_results_list)}개 블록 발견.\n")
                                    img_bytes_io.seek(0) # BytesIO 재사용을 위해 포인터 처음으로
                                    with Image.open(img_bytes_io) as img_to_render_on_base:
                                        original_img_format = img_to_render_on_base.format
                                        edited_img_pil = img_to_render_on_base.copy() # 복사본에 렌더링
                                        
                                        any_ocr_text_rendered = False
                                        for ocr_idx, ocr_res_item in enumerate(ocr_results_list):
                                            if stop_event and stop_event.is_set(): break
                                            if not (isinstance(ocr_res_item, (list, tuple)) and len(ocr_res_item) >= 2):
                                                f_task_log.write(f"          잘못된 형식의 OCR 결과 항목 ({ocr_idx+1}). 건너뜀.\n"); continue
                                            
                                            ocr_box_coords = ocr_res_item[0]
                                            ocr_text_conf_pair = ocr_res_item[1]
                                            ocr_angle_info = ocr_res_item[2] if len(ocr_res_item) > 2 else None

                                            if not (isinstance(ocr_text_conf_pair, (list, tuple)) and len(ocr_text_conf_pair) == 2):
                                                f_task_log.write(f"          잘못된 형식의 OCR 텍스트/신뢰도 쌍 ({ocr_idx+1}). 건너뜀.\n"); continue
                                            
                                            ocr_text_original, ocr_confidence = ocr_text_conf_pair
                                            current_text_for_progress_update = ocr_text_original # UI 업데이트용
                                            item_progress_type_for_ui = "이미지 텍스트 번역" # 더 구체적으로

                                            f_task_log.write(f"          OCR Text [{ocr_idx+1}]: \"{ocr_text_original.strip()[:30]}...\" (신뢰도: {ocr_confidence:.2f}, 각도: {ocr_angle_info})\n")

                                            if not is_ocr_text_valid(ocr_text_original):
                                                f_task_log.write(f"            -> [스킵됨 - OCR 유효성 낮음]\n"); continue
                                            if should_skip_translation(ocr_text_original): # 번역 스킵 조건도 확인
                                                f_task_log.write(f"            -> [스킵됨 - 번역 불필요]\n"); continue
                                            
                                            translated_ocr_text_val = translator.translate_text(
                                                ocr_text_original, src_lang_ui_name, tgt_lang_ui_name,
                                                model_name, ollama_service,
                                                is_ocr_text=True, ocr_temperature=ocr_temperature # OCR 온도 전달
                                            )
                                            f_task_log.write(f"            -> 번역 결과: \"{translated_ocr_text_val.strip()[:30]}...\"\n")

                                            if "오류:" not in translated_ocr_text_val and translated_ocr_text_val.strip():
                                                try:
                                                    # 렌더링 시에도 UI 업데이트 콜백 호출 가능 (세부 진행 표시)
                                                    # progress_callback_item_completed(slide_idx + 1, "이미지 텍스트 렌더링", 0, translated_ocr_text_val[:20])
                                                    
                                                    edited_img_pil = ocr_handler.render_translated_text_on_image(
                                                        edited_img_pil, ocr_box_coords, translated_ocr_text_val,
                                                        font_code_for_render=font_code_for_render,
                                                        original_text=ocr_text_original, ocr_angle=ocr_angle_info
                                                    )
                                                    any_ocr_text_rendered = True
                                                    f_task_log.write(f"              -> 렌더링 완료.\n")
                                                except Exception as e_render:
                                                    f_task_log.write(f"              오류: OCR 텍스트 렌더링 실패: {e_render}\n")
                                                    logger.error(f"OCR 텍스트 렌더링 실패 ('{item_name_log}'): {e_render}", exc_info=True)
                                            else:
                                                f_task_log.write(f"            -> 번역 실패 또는 빈 결과로 렌더링 안 함.\n")
                                        
                                        if stop_event and stop_event.is_set(): break # OCR 블록 루프 중단 시

                                        if any_ocr_text_rendered: # 렌더링된 내용이 있을 때만 이미지 교체
                                            output_img_stream = io.BytesIO()
                                            save_format_ocr_img = original_img_format if original_img_format and original_img_format.upper() in ['JPEG', 'PNG', 'GIF', 'BMP', 'TIFF'] else 'PNG'
                                            edited_img_pil.save(output_img_stream, format=save_format_ocr_img)
                                            output_img_stream.seek(0)

                                            # 이미지 교체 로직 (기존 shape_id 유지 시도)
                                            # 기존 이미지 shape의 위치, 크기, 이름 등 속성 저장
                                            left, top, width, height = current_shape_object.left, current_shape_object.top, current_shape_object.width, current_shape_object.height
                                            name = current_shape_object.name
                                            
                                            # 기존 이미지 shape 제거
                                            sp_xml_elem = current_shape_object.element
                                            parent_xml_elem = sp_xml_elem.getparent()
                                            if parent_xml_elem is not None:
                                                parent_xml_elem.remove(sp_xml_elem)
                                            
                                                # 새 이미지 추가 (동일 슬라이드에)
                                                new_pic_shape = prs.slides[slide_idx].shapes.add_picture(
                                                    output_img_stream, left, top, width=width, height=height
                                                )
                                                if name: new_pic_shape.name = name # 이름 복원
                                                f_task_log.write(f"        이미지 '{item_name_log}' 성공적으로 교체됨.\n")
                                            else:
                                                f_task_log.write(f"        경고: 이미지 '{item_name_log}'의 부모 XML 요소 찾지 못해 교체 실패. 원본 유지.\n")
                                        else:
                                            f_task_log.write(f"        이미지 '{item_name_log}'에 번역 및 렌더링된 텍스트가 없어 변경 없음.\n")
                                else: # OCR 결과 없음
                                    f_task_log.write(f"        이미지 '{item_name_log}' 내에서 OCR 텍스트 발견되지 않음.\n")
                        except FileNotFoundError:
                            err_msg_fnf = f"      오류 (1단계): '{item_name_log}' 이미지 파일 처리 중 FileNotFoundError. 건너뜀.\n"
                            f_task_log.write(err_msg_fnf)
                            logger.error(err_msg_fnf.strip())
                            if progress_callback_item_completed:
                                progress_callback_item_completed(slide_idx + 1, "오류", weighted_work_for_this_item, "이미지 파일 접근 불가")
                        except IOError as e_io_img:
                            err_msg_io = f"      오류 (1단계): '{item_name_log}' 이미지 처리 중 I/O 오류: {e_io_img}. 건너뜀.\n"
                            f_task_log.write(err_msg_io)
                            logger.error(f"I/O error during OCR image processing for '{item_name_log}': {e_io_img}", exc_info=False) # 스택 트레이스는 너무 길 수 있음
                            if progress_callback_item_completed:
                                progress_callback_item_completed(slide_idx + 1, "오류", weighted_work_for_this_item, "이미지 I/O 오류")
                        except Exception as e_ocr_general:
                            err_msg_ocr_gen = f"      오류 (1단계): '{item_name_log}' 이미지 처리 중 예기치 않은 오류: {e_ocr_general}. 건너뜀.\n"
                            f_task_log.write(err_msg_ocr_gen)
                            logger.error(f"Unexpected error processing image OCR for '{item_name_log}': {e_ocr_general}", exc_info=True)
                            if progress_callback_item_completed:
                                progress_callback_item_completed(slide_idx + 1, "오류", weighted_work_for_this_item, "이미지 처리 중 알 수 없는 오류")
                else: # 처리 대상 타입이 아니거나, 이미지 번역이 비활성화된 경우
                    if item_type_internal not in ['text_shape', 'table_cell', 'image']:
                         f_task_log.write(f"    '{item_name_log}' (타입: {item_type_internal})은 1단계에서 처리 대상 아님. 건너뜀.\n")
                    elif item_type_internal == 'image' and not image_translation_enabled:
                         f_task_log.write(f"    '{item_name_log}' (타입: 이미지) - 이미지 번역 비활성화로 건너뜀.\n")
                         weighted_work_for_this_item = config.WEIGHT_IMAGE # 작업량은 계산하되, 실제 작업은 안함
                    elif item_data.get('char_count', 0) == 0 and item_type_internal in ['text_shape', 'table_cell']:
                         weighted_work_for_this_item = 0 # 번역 안하므로 작업량 0
                         f_task_log.write(f"    '{item_name_log}' (타입: {item_type_internal}) 빈 텍스트 또는 스킵 대상임. 작업량 0.\n")


                if progress_callback_item_completed and not (stop_event and stop_event.is_set()):
                    progress_callback_item_completed(slide_idx + 1, item_progress_type_for_ui, weighted_work_for_this_item, current_text_for_progress_update)
                f_task_log.write("\n")

            if stop_event and stop_event.is_set():
                f_task_log.write(f"--- 1단계: 차트 외 요소 번역 중단됨 ---\n")
                return False # 중단 시 실패 반환

            f_task_log.write(f"--- 1단계: 차트 외 요소 번역 완료 ---\n\n")
            return True # 모든 요소 처리 완료 시 성공 반환