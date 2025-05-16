# pptx_handler.py
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN # PP_ALIGN 추가
from pptx.enum.lang import MSO_LANGUAGE_ID

import os
import io
import logging
import re
from datetime import datetime
import hashlib
import traceback
from PIL import Image
import tempfile # 추가
import shutil # 추가

# 새로 추가된 클래스 및 타입 힌팅용 import
from chart_xml_handler import ChartXmlHandler
from typing import TYPE_CHECKING, Optional, Dict, Any, List, Tuple

if TYPE_CHECKING:
    from translator import OllamaTranslator
    from ollama_service import OllamaService
    from ocr_handler import BaseOcrHandler

logger = logging.getLogger(__name__)

# 의미있는 문자열 판단을 위한 정규식 및 최소 비율 (기존과 동일)
MIN_MEANINGFUL_CHAR_RATIO_SKIP = 0.1
MIN_MEANINGFUL_CHAR_RATIO_OCR = 0.1
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
    """주어진 텍스트가 번역을 건너뛸 대상인지 판단합니다."""
    if not text: return True
    stripped_text = text.strip()
    if not stripped_text:
        logger.debug(f"번역 스킵 (공백만 존재): '{text[:50]}...'")
        return True
    
    # 숫자, 특수문자만으로 이루어진 경우도 의미있는 문자 검색을 통해 걸러냄
    if not MEANINGFUL_CHAR_PATTERN.search(stripped_text):
        logger.debug(f"번역 스킵 (의미 있는 문자 없음): '{stripped_text[:50]}...'")
        return True

    text_len = len(stripped_text)
    # 매우 짧은 문자열도 의미 있는 문자가 포함되어 있으면 번역 시도
    if text_len <= 3: # 임계값 조정 가능
        if MEANINGFUL_CHAR_PATTERN.search(stripped_text):
            logger.debug(f"번역 시도 (짧은 문자열, 의미 문자 포함): '{stripped_text}'")
            return False
        else:
            logger.debug(f"번역 스킵 (짧고 의미 있는 문자 없음): '{stripped_text}'")
            return True

    meaningful_chars_count = len(MEANINGFUL_CHAR_PATTERN.findall(stripped_text))
    ratio = meaningful_chars_count / text_len
    if ratio < MIN_MEANINGFUL_CHAR_RATIO_SKIP:
        logger.debug(f"번역 스킵 (의미 문자 비율 낮음 {ratio:.2f}, 임계값: {MIN_MEANINGFUL_CHAR_RATIO_SKIP}): '{stripped_text[:50]}...'")
        return True

    logger.debug(f"번역 시도 (조건 통과): '{stripped_text[:50]}...'")
    return False

def is_ocr_text_valid(text: str) -> bool:
    """OCR로 추출된 텍스트가 유효한지 (번역할 가치가 있는지) 판단합니다."""
    if not text: return False
    stripped_text = text.strip()
    if not stripped_text: return False

    if not MEANINGFUL_CHAR_PATTERN.search(stripped_text):
        logger.debug(f"OCR 유효성 스킵 (의미 있는 문자 없음): '{stripped_text[:50]}...'")
        return False
    
    text_len = len(stripped_text)
    if text_len <= 2: # OCR은 더 짧은 문자열도 노이즈 가능성이 높음
        if MEANINGFUL_CHAR_PATTERN.search(stripped_text):
             logger.debug(f"OCR 유효 (매우 짧은 문자열, 의미 문자 포함): '{stripped_text}'")
             return True
        else:
            logger.debug(f"OCR 유효성 스킵 (매우 짧고 의미 있는 문자 없음): '{stripped_text}'")
            return False

    meaningful_chars_count = len(MEANINGFUL_CHAR_PATTERN.findall(stripped_text))
    ratio = meaningful_chars_count / text_len
    if ratio < MIN_MEANINGFUL_CHAR_RATIO_OCR:
        logger.debug(f"OCR 유효성 스킵 (의미 문자 비율 낮음 {ratio:.2f}, 임계값: {MIN_MEANINGFUL_CHAR_RATIO_OCR}): '{stripped_text[:50]}...'")
        return False
        
    logger.debug(f"OCR 유효 (조건 통과): '{stripped_text[:50]}...'")
    return True


class PptxHandler:
    def __init__(self):
        pass

    def get_file_info(self, file_path: str) -> Dict[str, int]:
        """PPTX 파일의 기본 정보를 분석하여 반환합니다 (슬라이드 수, 텍스트 요소, 이미지, 차트 수)."""
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
                        # 테이블 내 텍스트는 각 셀을 하나의 요소로 간주 (개선 가능)
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if hasattr(cell.text_frame, 'text') and cell.text_frame.text and cell.text_frame.text.strip():
                                    info["text_elements"] += 1 # 테이블 셀도 텍스트 요소로 카운트
                    elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                        info["chart_elements"] +=1
            logger.info(f"파일 분석 완료 (OCR 미수행): Slides:{info['slide_count']}, Text:{info['text_elements']}, Img:{info['image_elements']}, Chart:{info['chart_elements']}")
        except Exception as e:
            logger.error(f"'{os.path.basename(file_path)}' 파일 정보 분석 오류: {e}", exc_info=True)
        return info

    def _get_style_properties(self, font_object) -> Dict[str, Any]:
        """Font 객체에서 스타일 속성을 추출합니다."""
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
            logger.debug(f"Font language_id 가져오기 실패: {e_lang}. 기본값 사용.")


        if hasattr(font_object, 'color') and hasattr(font_object.color, 'type'):
            color_type = font_object.color.type
            if color_type == MSO_COLOR_TYPE.RGB:
                try: style_props['color_rgb'] = tuple(font_object.color.rgb)
                except AttributeError: pass
            elif color_type == MSO_COLOR_TYPE.SCHEME:
                try: style_props['color_theme_index'] = font_object.color.theme_color
                except AttributeError: pass
                try: style_props['color_brightness'] = font_object.color.brightness
                except AttributeError: style_props['color_brightness'] = 0.0 # 기본값
        
        return style_props

    def _get_text_style(self, run) -> Dict[str, Any]:
        """Run 객체에서 텍스트 스타일 (폰트 및 하이퍼링크)을 추출합니다."""
        style = self._get_style_properties(run.font)
        try:
            style['hyperlink_address'] = run.hyperlink.address if run.hyperlink and run.hyperlink.address else None
        except AttributeError:
            style['hyperlink_address'] = None
        return style

    def _apply_style_properties(self, target_font_object, style_dict_to_apply: Dict[str, Any]):
        """추출된 스타일 속성을 대상 Font 객체에 적용합니다."""
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
                font.color.theme_color = style_dict_to_apply['color_theme_index']
                # 밝기(brightness)는 theme_color 설정 후에 적용해야 함
                brightness_val = style_dict_to_apply.get('color_brightness', 0.0) # 기본값 0.0
                font.color.brightness = max(-1.0, min(1.0, float(brightness_val)))
            except Exception as e: logger.warning(f"테마 색상 {style_dict_to_apply['color_theme_index']} 적용 실패: {e}")
        
        if style_dict_to_apply.get('language_id') is not None:
            try: font.language_id = style_dict_to_apply['language_id']
            except Exception as e_lang: logger.warning(f"language_id '{style_dict_to_apply['language_id']}' 적용 실패: {e_lang}")

    def _apply_text_style(self, run, style_to_apply: Dict[str, Any]):
        """추출된 텍스트 스타일을 대상 Run 객체에 적용합니다."""
        if not style_to_apply or run is None:
            return
        self._apply_style_properties(run.font, style_to_apply)
        if style_to_apply.get('hyperlink_address'):
            try:
                # 기존 하이퍼링크 제거 후 새로 추가하는 것이 더 안정적일 수 있음
                # 여기서는 주소만 변경 시도
                hlink = run.hyperlink
                if hlink is None: # 하이퍼링크가 없었다면 새로 생성
                    hlink = run.hyperlink # 다시 가져오거나, add_hyperlink 사용 (pptx 라이브러리 버전 확인)
                    # run.add_hyperlink(address=style_to_apply['hyperlink_address']) # 이 방식이 더 나을 수 있음
                if hlink: # 주소만 설정
                    hlink.address = style_to_apply['hyperlink_address']
                else: # 그래도 없으면, 새 run을 만들고 하이퍼링크 추가 후 기존 run 삭제? 복잡도 증가
                    logger.warning(f"Run에 하이퍼링크 설정 실패 (기존 hlink 객체 없음): {run.text[:20]}")
            except Exception as e: logger.warning(f"Run에 하이퍼링크 주소 적용 오류: {e}")


    def translate_presentation(self, file_path: str, src_lang_ui_name: str, tgt_lang_ui_name: str,
                               translator: 'OllamaTranslator', ocr_handler: Optional['BaseOcrHandler'],
                               model_name: str, ollama_service: 'OllamaService',
                               font_code_for_render: str, task_log_filepath: str,
                               progress_callback=None, stop_event=None) -> Optional[str]:
        """
        PPTX 파일 내의 텍스트 요소(텍스트 상자, 표), 이미지 내 텍스트(OCR), 차트 XML을 번역합니다.
        번역은 2단계로 진행됩니다:
        1. 차트 외 요소 번역 후 임시 파일 저장.
        2. ChartXmlHandler를 사용하여 임시 파일 내 차트 번역 후 최종 파일 저장.
        """
        
        temp_dir_for_pptx_handler = tempfile.mkdtemp(prefix="pptx_trans_main_")
        temp_pptx_for_chart_translation_path: Optional[str] = None
        final_output_path_result: Optional[str] = None

        try:
            with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                start_time_log = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f_task_log.write(f"--- 프레젠테이션 번역 로그 (PptxHandler) 시작: {os.path.basename(file_path)} ---\n")
                f_task_log.write(f"시작 시간: {start_time_log}\n원본 파일: {file_path}\n소스 언어: {src_lang_ui_name}, 대상 언어: {tgt_lang_ui_name}, 모델: {model_name}\n")
                f_task_log.write(f"OCR 핸들러: {type(ocr_handler).__name__ if ocr_handler else '없음'}, OCR 렌더링 폰트 코드: {font_code_for_render}\n")
                f_task_log.write(f"임시 작업 폴더 (PptxHandler): {temp_dir_for_pptx_handler}\n\n")
                logger.info(f"'{os.path.basename(file_path)}' 프레젠테이션 번역 시작 (PptxHandler). 로그: {task_log_filepath}")

                prs = Presentation(file_path)
                
                # --- 1단계: 차트 외 요소 번역 (텍스트, 표, OCR) ---
                f_task_log.write("--- 1단계: 차트 외 요소 번역 시작 ---\n")
                logger.info("1단계: 차트 외 요소 (텍스트 상자, 표, OCR 등) 번역 중...")

                elements_map_phase1: List[Dict[str, Any]] = []
                for slide_idx, slide in enumerate(prs.slides):
                    if stop_event and stop_event.is_set(): break
                    for shape_idx, shape in enumerate(slide.shapes):
                        if stop_event and stop_event.is_set(): break
                        
                        shape_id = getattr(shape, 'shape_id', f"slide{slide_idx}_shape{shape_idx}")
                        element_name = shape.name or f"S{slide_idx+1}_Shape{shape_idx}_Id{shape_id}"
                        item_info: Optional[Dict[str, Any]] = None

                        if shape.shape_type == MSO_SHAPE_TYPE.CHART:
                            f_task_log.write(f"  S{slide_idx+1} ('{element_name}') 차트 발견. 1단계에서는 건너뛰고 2단계에서 XML 처리 예정.\n")
                            continue # 차트는 1단계에서 건너뜀

                        if shape.has_text_frame and hasattr(shape.text_frame, 'text') and \
                           shape.text_frame.text and shape.text_frame.text.strip():
                            item_info = {'type': 'text', 'slide_idx': slide_idx, 'shape_id': shape_id, 'shape_obj_ref': shape, 'name': element_name}
                        elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                            if ocr_handler: # OCR 핸들러가 있을 때만 이미지 처리 대상
                                item_info = {'type': 'image', 'slide_idx': slide_idx, 'shape_id': shape_id, 'shape_obj_ref': shape, 'name': element_name}
                        elif shape.has_table:
                            for r_idx, row in enumerate(shape.table.rows):
                                for c_idx, cell in enumerate(row.cells):
                                    if hasattr(cell.text_frame, 'text') and cell.text_frame.text and cell.text_frame.text.strip():
                                        elements_map_phase1.append({
                                            'type': 'table_cell', 'slide_idx': slide_idx,
                                            'shape_id': shape_id, # 부모 테이블 shape의 ID
                                            'table_shape_obj_ref': shape, # 부모 테이블 shape 객체
                                            'row_idx': r_idx, 'col_idx': c_idx,
                                            'name': f"{element_name}_R{r_idx}C{c_idx}"
                                        })
                        if item_info:
                            elements_map_phase1.append(item_info)
                
                total_elements_phase1 = len(elements_map_phase1)
                f_task_log.write(f"1단계 번역 대상 요소 (차트 제외): {total_elements_phase1}개.\n")
                logger.info(f"1단계 번역 대상 요소 (차트 제외): {total_elements_phase1}개.")

                if not elements_map_phase1 and not any(s.shape_type == MSO_SHAPE_TYPE.CHART for slide in prs.slides for s in slide.shapes):
                    # 번역할 텍스트 요소도 없고, 차트도 없으면 원본 복사 후 종료
                    safe_tgt_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang_ui_name)
                    no_translation_output_filename = os.path.splitext(file_path)[0] + f"_{safe_tgt_lang_suffix}_translated.pptx"
                    shutil.copy2(file_path, no_translation_output_filename)
                    msg = "1단계 번역 대상 요소(텍스트/이미지/표)가 없고, 차트도 없어 원본 파일을 복사본으로 저장합니다."
                    f_task_log.write(msg + f"\n저장된 파일: {no_translation_output_filename}\n")
                    logger.info(msg)
                    return no_translation_output_filename


                translated_count_phase1 = 0
                # 각 요소의 원본 단락 스타일을 저장할 딕셔너리
                original_paragraph_styles_phase1: Dict[Tuple[int, Any, Any], List[Dict[str, Any]]] = {}

                for item_data in elements_map_phase1:
                    if stop_event and stop_event.is_set(): break

                    slide_idx = item_data['slide_idx']
                    item_name_log = item_data['name']
                    item_type = item_data['type']
                    
                    # shape_obj_ref 또는 table_shape_obj_ref 사용
                    current_shape_object = item_data.get('shape_obj_ref') or item_data.get('table_shape_obj_ref')
                    if not current_shape_object: # 방어 코드
                        f_task_log.write(f"  오류 (1단계): S{slide_idx+1} ('{item_name_log}') 요소 객체 참조 실패. 건너뜀.\n")
                        translated_count_phase1 +=1
                        if progress_callback: progress_callback(slide_idx + 1, item_type, translated_count_phase1, total_elements_phase1, "오류: 요소 참조 불가")
                        continue
                        
                    shape_id_log = getattr(current_shape_object, 'shape_id', 'N/A_ID')
                    f_task_log.write(f"  [1단계 S{slide_idx+1}] 처리 시작: '{item_name_log}' (ID: {shape_id_log}), 타입: {item_type}\n")
                    current_text_for_progress_update = ""

                    text_frame_for_processing: Optional[Any] = None # TextFrame 객체
                    if item_type == 'text':
                        if current_shape_object.has_text_frame:
                            text_frame_for_processing = current_shape_object.text_frame
                    elif item_type == 'table_cell':
                        row_idx = item_data['row_idx']
                        col_idx = item_data['col_idx']
                        if current_shape_object.has_table:
                            try:
                                text_frame_for_processing = current_shape_object.table.cell(row_idx, col_idx).text_frame
                            except IndexError:
                                logger.error(f"1단계 테이블 셀 접근 오류 (IndexError): {item_name_log} at R{row_idx}C{col_idx}")
                                f_task_log.write(f"    오류: 1단계 테이블 셀 접근 실패 (R{row_idx}C{col_idx}). 건너뜀.\n")

                    # 텍스트 요소 또는 테이블 셀 번역 처리
                    if text_frame_for_processing and hasattr(text_frame_for_processing, 'text') and \
                       text_frame_for_processing.text and text_frame_for_processing.text.strip():
                        
                        original_full_text = text_frame_for_processing.text
                        current_text_for_progress_update = original_full_text
                        f_task_log.write(f"    1단계 번역 대상 텍스트 발견: \"{original_full_text.strip()[:50]}...\"\n")

                        # 스타일 저장을 위한 고유 키 생성
                        style_key_suffix_val = (item_data.get('row_idx', -1), item_data.get('col_idx', -1)) if item_type == 'table_cell' else 'shape_text'
                        style_unique_key = (slide_idx, id(current_shape_object), style_key_suffix_val)

                        if style_unique_key not in original_paragraph_styles_phase1:
                            para_styles_collected: List[Dict[str, Any]] = []
                            for para_obj in text_frame_for_processing.paragraphs:
                                para_default_font_style = self._get_style_properties(para_obj.font)
                                runs_info: List[Dict[str, Any]] = []
                                if para_obj.runs: # Run이 있는 경우
                                    for run_obj in para_obj.runs:
                                        runs_info.append({'text': run_obj.text, 'style': self._get_text_style(run_obj)})
                                elif para_obj.text and para_obj.text.strip(): # Run이 없고 단락에 직접 텍스트가 있는 경우
                                    run_style_from_para = para_default_font_style.copy()
                                    run_style_from_para['hyperlink_address'] = None # 단락 기본 스타일에는 하이퍼링크 없음
                                    runs_info.append({'text': para_obj.text, 'style': run_style_from_para})
                                
                                para_styles_collected.append({
                                    'runs': runs_info,
                                    'alignment': para_obj.alignment,
                                    'level': para_obj.level,
                                    'space_before': para_obj.space_before,
                                    'space_after': para_obj.space_after,
                                    'line_spacing': para_obj.line_spacing,
                                    'paragraph_default_run_style': para_default_font_style # 단락 기본 Run 스타일
                                })
                            original_paragraph_styles_phase1[style_unique_key] = para_styles_collected
                            f_task_log.write(f"      '{item_name_log}'의 원본 단락 스타일 저장 (1단계).\n")
                        
                        if should_skip_translation(original_full_text):
                            f_task_log.write(f"      [1단계 스킵됨 - 번역 불필요 또는 의미 없는 텍스트]\n")
                        else:
                            translated_text_content = translator.translate_text(original_full_text, src_lang_ui_name, tgt_lang_ui_name, model_name, ollama_service, is_ocr_text=False)
                            log_trans_text_snippet = translated_text_content.replace('\n', ' / ').strip()[:100]
                            f_task_log.write(f"      [1단계 번역 전] \"{original_full_text.strip()[:50]}...\" -> [1단계 번역 후] \"{log_trans_text_snippet}...\"\n")

                            if "오류:" not in translated_text_content and translated_text_content.strip():
                                # 번역된 텍스트 적용 및 스타일 복원
                                stored_paras_info_apply = original_paragraph_styles_phase1.get(style_unique_key, [])
                                
                                original_tf_auto_sz = getattr(text_frame_for_processing, 'auto_size', None)
                                original_tf_word_wrp = getattr(text_frame_for_processing, 'word_wrap', None)
                                original_tf_v_anchor = getattr(text_frame_for_processing, 'vertical_anchor', None)
                                original_tf_margins = {
                                    'left': getattr(text_frame_for_processing, 'margin_left', None),
                                    'right': getattr(text_frame_for_processing, 'margin_right', None),
                                    'top': getattr(text_frame_for_processing, 'margin_top', None),
                                    'bottom': getattr(text_frame_for_processing, 'margin_bottom', None),
                                }


                                if original_tf_auto_sz is not None and original_tf_auto_sz != MSO_AUTO_SIZE.NONE:
                                    text_frame_for_processing.auto_size = MSO_AUTO_SIZE.NONE # 자동 크기 조정을 잠시 끔
                                if original_tf_word_wrp is not None: text_frame_for_processing.word_wrap = True # 줄바꿈 활성화

                                text_frame_for_processing.clear() # 기존 단락 모두 제거 (python-pptx 방식)
                                # XML 레벨에서 <a:p> 태그 직접 제거 (더 확실한 정리)
                                if hasattr(text_frame_for_processing, '_element') and text_frame_for_processing._element is not None:
                                    txBody_xml = text_frame_for_processing._element
                                    p_tags_to_remove = [child for child in txBody_xml if child.tag.endswith('}p')]
                                    if p_tags_to_remove:
                                        f_task_log.write(f"        (강화된 정리) XML에서 기존 단락 {len(p_tags_to_remove)}개 제거.\n")
                                        for p_xml_tag in p_tags_to_remove: txBody_xml.remove(p_xml_tag)
                                
                                # 번역된 텍스트를 줄바꿈 기준으로 나누어 단락 생성
                                lines_from_translation = translated_text_content.splitlines()
                                # 빈 줄도 유지하되, 연속된 빈 줄은 하나로 합치는 등의 처리 가능 (현재는 그대로 적용)
                                if not lines_from_translation and translated_text_content: # 줄바꿈 없는 한 줄 텍스트
                                    lines_from_translation = [translated_text_content]
                                elif not lines_from_translation: # 빈 결과
                                    lines_from_translation = [" "] # 공백 한 칸으로 대체 (단락 유지 위함)

                                for line_idx, line_txt in enumerate(lines_from_translation):
                                    new_para = text_frame_for_processing.add_paragraph()
                                    # 원본 단락 스타일 적용 (저장된 정보가 있다면)
                                    para_style_template = stored_paras_info_apply[min(line_idx, len(stored_paras_info_apply)-1)] if stored_paras_info_apply else {}
                                    
                                    if para_style_template.get('alignment') is not None: new_para.alignment = para_style_template['alignment']
                                    else: new_para.alignment = PP_ALIGN.LEFT # 기본 정렬

                                    new_para.level = para_style_template.get('level', 0)
                                    if para_style_template.get('space_before') is not None: new_para.space_before = para_style_template['space_before']
                                    if para_style_template.get('space_after') is not None: new_para.space_after = para_style_template['space_after']
                                    if para_style_template.get('line_spacing') is not None: new_para.line_spacing = para_style_template['line_spacing']
                                    
                                    # 단락 기본 폰트 스타일 적용
                                    if 'paragraph_default_run_style' in para_style_template:
                                        self._apply_style_properties(new_para.font, para_style_template['paragraph_default_run_style'])

                                    new_run = new_para.add_run()
                                    new_run.text = line_txt if line_txt.strip() else " " # 빈 줄은 공백 1칸으로
                                    if not new_run.text.strip() and new_run.text != " ": new_run.text = " " # 순수 공백만 있을 때 " " 보장

                                    # Run 스타일 적용 (원본 첫번째 Run 스타일 또는 단락 기본 스타일)
                                    run_style_to_apply = {}
                                    if para_style_template.get('runs') and para_style_template['runs']:
                                        run_style_to_apply = para_style_template['runs'][0]['style'] # 첫 번째 run 스타일 대표 사용
                                    elif 'paragraph_default_run_style' in para_style_template: # run 정보 없으면 단락 기본 스타일
                                        run_style_to_apply = para_style_template['paragraph_default_run_style'].copy()
                                        run_style_to_apply['hyperlink_address'] = None # 기본에는 하이퍼링크 없음

                                    if run_style_to_apply:
                                        self._apply_text_style(new_run, run_style_to_apply)
                                
                                # TextFrame 원본 설정 복원
                                if original_tf_auto_sz is not None: text_frame_for_processing.auto_size = original_tf_auto_sz
                                if original_tf_word_wrp is not None: text_frame_for_processing.word_wrap = original_tf_word_wrp
                                if original_tf_v_anchor is not None: text_frame_for_processing.vertical_anchor = original_tf_v_anchor
                                for margin_prop, val in original_tf_margins.items():
                                    if val is not None: setattr(text_frame_for_processing, f"margin_{margin_prop}", val)

                                f_task_log.write(f"        '{item_name_log}' 1단계 번역된 텍스트 적용 완료.\n")
                            else:
                                f_task_log.write(f"      -> 1단계 텍스트 번역 실패 또는 빈 결과: {translated_text_content}\n")
                    
                    # 이미지 OCR 처리
                    elif item_type == 'image' and ocr_handler:
                        f_task_log.write(f"    1단계 OCR 처리 시도: '{item_name_log}'\n")
                        current_text_for_progress_update = "[이미지 OCR 처리 중]"
                        if current_shape_object is None or not hasattr(current_shape_object, 'image') or \
                           not hasattr(current_shape_object.image, 'blob'):
                            f_task_log.write(f"      경고 (1단계): '{item_name_log}' 이미지 객체 또는 blob 데이터 접근 불가. 건너뜀.\n")
                            logger.warning(f"Cannot access image data for '{item_name_log}' (ID:{shape_id_log}). Skipping OCR.")
                        else:
                            try:
                                img_bytes_io = io.BytesIO(current_shape_object.image.blob)
                                with Image.open(img_bytes_io) as img_pil_original_ocr:
                                    img_pil_rgb_ocr = img_pil_original_ocr.convert("RGB") # OCR은 RGB로
                                    
                                    # 이미지 해시 로깅 (디버깅용)
                                    try:
                                        hasher_md5 = hashlib.md5()
                                        img_byte_arr_for_hashing = io.BytesIO()
                                        img_pil_rgb_ocr.save(img_byte_arr_for_hashing, format='PNG') # 일관된 형식으로 해시
                                        hasher_md5.update(img_byte_arr_for_hashing.getvalue())
                                        img_hash_val_log = hasher_md5.hexdigest()
                                        f_task_log.write(f"      OCR 대상 RGB 이미지 MD5 해시: {img_hash_val_log}\n")
                                    except Exception as e_hash: logger.debug(f"이미지 해시 생성 중 오류: {e_hash}")

                                    ocr_results_list = ocr_handler.ocr_image(img_pil_rgb_ocr) # [[box, (text, conf), angle], ...]
                                    
                                    if ocr_results_list:
                                        f_task_log.write(f"        이미지 내 OCR 텍스트 {len(ocr_results_list)}개 블록 발견.\n")
                                        # 원본 이미지 복사본에 번역된 텍스트 렌더링
                                        img_bytes_io.seek(0) # BytesIO 재사용 위해 포인터 초기화
                                        with Image.open(img_bytes_io) as img_to_render_on_base:
                                            # 원본 이미지 포맷 유지 시도 (렌더링 후 저장 시)
                                            original_img_format = img_to_render_on_base.format
                                            edited_img_pil = img_to_render_on_base.copy() # 실제 렌더링은 복사본에
                                            
                                            any_ocr_text_rendered = False
                                            for ocr_idx, ocr_res_item in enumerate(ocr_results_list):
                                                if stop_event and stop_event.is_set(): break
                                                if not (isinstance(ocr_res_item, (list, tuple)) and len(ocr_res_item) >= 2):
                                                    f_task_log.write(f"          잘못된 형식의 OCR 결과 항목. 건너뜀.\n"); continue
                                                
                                                ocr_box_coords = ocr_res_item[0]
                                                ocr_text_conf_pair = ocr_res_item[1]
                                                ocr_angle_info = ocr_res_item[2] if len(ocr_res_item) > 2 else None

                                                if not (isinstance(ocr_text_conf_pair, (list, tuple)) and len(ocr_text_conf_pair) == 2):
                                                    f_task_log.write(f"          잘못된 형식의 OCR 텍스트/신뢰도 쌍. 건너뜀.\n"); continue
                                                
                                                ocr_text_original, ocr_confidence = ocr_text_conf_pair
                                                current_text_for_progress_update = ocr_text_original # UI 업데이트용
                                                f_task_log.write(f"          OCR Text [{ocr_idx+1}]: \"{ocr_text_original.strip()[:30]}...\" (신뢰도: {ocr_confidence:.2f}, 각도: {ocr_angle_info})\n")

                                                if not is_ocr_text_valid(ocr_text_original):
                                                    f_task_log.write(f"            -> [스킵됨 - OCR 유효성 낮음]\n"); continue
                                                if should_skip_translation(ocr_text_original):
                                                    f_task_log.write(f"            -> [스킵됨 - 번역 불필요]\n"); continue
                                                
                                                translated_ocr_text_val = translator.translate_text(ocr_text_original, src_lang_ui_name, tgt_lang_ui_name, model_name, ollama_service, is_ocr_text=True)
                                                f_task_log.write(f"            -> 번역 결과: \"{translated_ocr_text_val.strip()[:30]}...\"\n")

                                                if "오류:" not in translated_ocr_text_val and translated_ocr_text_val.strip():
                                                    try:
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
                                            
                                            if stop_event and stop_event.is_set(): break # 내부 루프 중단 시 외부 루프도 중단

                                            if any_ocr_text_rendered:
                                                output_img_stream = io.BytesIO()
                                                save_format_ocr_img = original_img_format if original_img_format and original_img_format.upper() in ['JPEG', 'PNG', 'GIF', 'BMP', 'TIFF'] else 'PNG'
                                                edited_img_pil.save(output_img_stream, format=save_format_ocr_img)
                                                output_img_stream.seek(0)

                                                # 기존 이미지 교체
                                                slide_obj_for_pic_replace = prs.slides[slide_idx]
                                                pic_xml_element = current_shape_object.element
                                                parent_xml_of_pic = pic_xml_element.getparent()
                                                if parent_xml_of_pic is not None:
                                                    parent_xml_of_pic.remove(pic_xml_element) # 기존 <p:pic> 제거
                                                    # 새 이미지 추가 (위치 및 크기 유지)
                                                    new_pic_shape = slide_obj_for_pic_replace.shapes.add_picture(
                                                        output_img_stream, current_shape_object.left, current_shape_object.top,
                                                        width=current_shape_object.width, height=current_shape_object.height
                                                    )
                                                    # (선택적) 새 이미지 shape의 이름 등 속성 복사
                                                    if hasattr(current_shape_object, 'name') and current_shape_object.name:
                                                        new_pic_shape.name = current_shape_object.name

                                                    f_task_log.write(f"        이미지 '{item_name_log}' 성공적으로 교체됨.\n")
                                                else:
                                                    f_task_log.write(f"        경고: 이미지 '{item_name_log}'의 부모 XML 요소 찾지 못해 교체 실패.\n")
                                            else:
                                                f_task_log.write(f"        이미지 '{item_name_log}'에 번역 및 렌더링된 텍스트가 없어 변경 없음.\n")
                                    else:
                                        f_task_log.write(f"        이미지 '{item_name_log}' 내에서 OCR 텍스트 발견되지 않음.\n")
                            except FileNotFoundError: # PIL.Image.open 에서 발생 가능
                                f_task_log.write(f"      오류 (1단계): '{item_name_log}' 이미지 파일 처리 중 FileNotFoundError. 건너뜀.\n")
                                logger.error(f"FileNotFoundError during OCR image processing for '{item_name_log}'. Skipping.")
                            except IOError as e_io_img: # 손상된 이미지 등
                                f_task_log.write(f"      오류 (1단계): '{item_name_log}' 이미지 처리 중 I/O 오류: {e_io_img}. 건너뜀.\n")
                                logger.error(f"I/O error during OCR image processing for '{item_name_log}': {e_io_img}", exc_info=False) # 스택 트레이스 짧게
                            except Exception as e_ocr_general:
                                f_task_log.write(f"      오류 (1단계): '{item_name_log}' 이미지 처리 중 예기치 않은 오류: {e_ocr_general}. 건너뜀.\n")
                                logger.error(f"Unexpected error processing image OCR for '{item_name_log}': {e_ocr_general}", exc_info=True)
                    else: # 텍스트도 아니고, OCR 대상 이미지도 아닌 경우 (또는 텍스트 프레임 비었음)
                        if item_type not in ['text', 'table_cell', 'image']: # 이미 처리된 타입 제외
                             f_task_log.write(f"    '{item_name_log}' (타입: {item_type})은 1단계에서 처리 대상 아님. 건너뜀.\n")

                    translated_count_phase1 += 1
                    if progress_callback and not (stop_event and stop_event.is_set()):
                        # 진행 콜백 호출 (main.py에서 전체 진행률 계산)
                        progress_callback(slide_idx + 1, item_type, translated_count_phase1, total_elements_phase1, current_text_for_progress_update)
                    f_task_log.write("\n") # 각 요소 처리 후 로그 줄바꿈

                if stop_event and stop_event.is_set():
                    # 중지 시 현재까지 작업된 prs 객체를 임시 파일로 저장
                    stopped_filename_p1 = os.path.join(temp_dir_for_pptx_handler, f"{os.path.splitext(os.path.basename(file_path))[0]}_phase1_stopped.pptx")
                    prs.save(stopped_filename_p1)
                    f_task_log.write(f"1단계 번역 중지됨. 현재까지 작업 저장: {stopped_filename_p1}\n")
                    logger.info(f"1단계 번역 중지됨. 부분 저장 파일: {stopped_filename_p1}")
                    return stopped_filename_p1 # 중지 시 임시 파일 경로 반환

                # 1단계 번역 완료 후 임시 파일 저장 (필수)
                temp_pptx_for_chart_translation_path = os.path.join(temp_dir_for_pptx_handler, f"{os.path.splitext(os.path.basename(file_path))[0]}_temp_for_charts.pptx")
                prs.save(temp_pptx_for_chart_translation_path)
                f_task_log.write(f"--- 1단계: 차트 외 요소 번역 완료. 차트 번역용 임시 파일 저장: {temp_pptx_for_chart_translation_path} ---\n\n")
                logger.info(f"1단계 완료. 차트 XML 번역을 위해 임시 파일 저장: {temp_pptx_for_chart_translation_path}")

                # --- 2단계: ChartXmlHandler를 사용하여 차트 번역 ---
                f_task_log.write("--- 2단계: 차트 XML 직접 번역 시작 ---\n")
                logger.info("2단계: ChartXmlHandler를 사용하여 차트 XML 번역 중...")
                
                chart_xml_translator = ChartXmlHandler(translator_instance=translator, ollama_service_instance=ollama_service)
                
                # 최종 출력 파일명 결정 (main.py와 유사한 패턴)
                safe_target_lang_suffix = "".join(c if c.isalnum() else "_" for c in tgt_lang_ui_name)
                final_output_filename_base = f"{os.path.splitext(os.path.basename(file_path))[0]}_{safe_target_lang_suffix}_translated.pptx"
                # 최종 파일은 원본 파일과 같은 디렉토리에 저장 (main.py의 동작과 일치시키기 위함)
                output_directory = os.path.dirname(file_path)
                final_output_path_for_chart_handler_step = os.path.join(output_directory, final_output_filename_base)

                # ChartXmlHandler는 입력 PPTX 경로와 출력 PPTX 경로를 받아 처리
                # (주의: ChartXmlHandler 내부에서 stop_event를 직접 사용하지 않는다면, 여기서 추가적인 중지 로직은 어려움)
                # ChartXmlHandler가 오래 걸릴 수 있으므로, 그 내부에서도 stop_event를 확인하도록 수정하는 것이 이상적.
                # 현재 ChartXmlHandler에는 stop_event 처리가 없으므로, 호출 후 완료될 때까지 기다림.
                final_output_path_result = chart_xml_translator.translate_charts_in_pptx(
                    pptx_path=temp_pptx_for_chart_translation_path, # 1단계 결과물 입력
                    src_lang_ui_name=src_lang_ui_name,
                    tgt_lang_ui_name=tgt_lang_ui_name,
                    model_name=model_name,
                    output_path=final_output_path_for_chart_handler_step # 최종 저장될 경로
                )

                if stop_event and stop_event.is_set(): # ChartXmlHandler 호출 후 중지 체크 (만약 핸들러가 매우 빨리 끝나거나, 내부 중지 로직이 없다면)
                    f_task_log.write(f"2단계 차트 번역 완료 직후 또는 ChartXmlHandler 내부에서 중단 요청 감지 (외부 확인).\n")
                    logger.warning(f"사용자 요청으로 ChartXmlHandler 작업 후 중지됨 (외부 확인).")
                    # ChartXmlHandler가 생성한 파일이 있다면 그것을 반환, 없다면 1단계 결과물 반환
                    return final_output_path_result if final_output_path_result and os.path.exists(final_output_path_result) else temp_pptx_for_chart_translation_path

                if final_output_path_result and os.path.exists(final_output_path_result):
                    f_task_log.write(f"--- 2단계: 차트 XML 직접 번역 완료. 최종 파일 생성됨: {final_output_path_result} ---\n")
                    logger.info(f"2단계 (차트 XML 번역) 완료. 최종 파일: {final_output_path_result}")
                else:
                    f_task_log.write(f"오류: 2단계 차트 XML 번역 실패 또는 결과 파일 없음. 최종 파일 경로: {final_output_path_result}. 1단계 결과물({temp_pptx_for_chart_translation_path})을 최종 위치에 복사 시도.\n")
                    logger.error(f"2단계 차트 XML 번역 실패 또는 결과 파일({final_output_path_result})이 생성되지 않았습니다.")
                    # 차트 번역 실패 시, 1단계 결과물을 최종 파일명으로 복사
                    if os.path.exists(temp_pptx_for_chart_translation_path):
                        try:
                            shutil.copy2(temp_pptx_for_chart_translation_path, final_output_path_for_chart_handler_step)
                            final_output_path_result = final_output_path_for_chart_handler_step # 경로 업데이트
                            logger.info(f"차트 번역 실패로 1단계 결과물을 최종 경로에 복사: {final_output_path_result}")
                            f_task_log.write(f"  1단계 결과물을 최종 파일로 복사 완료: {final_output_path_result}\n")
                        except Exception as e_copy_fallback:
                            logger.error(f"차트 번역 실패 후 1단계 결과물 복사 중 오류: {e_copy_fallback}. 1단계 임시 파일 경로 반환 시도.")
                            final_output_path_result = temp_pptx_for_chart_translation_path # 최후의 수단 (임시 폴더 내 파일)
                            f_task_log.write(f"  오류: 1단계 결과물 최종 파일로 복사 실패. 임시 파일 경로 반환: {final_output_path_result}\n")
                    else:
                         logger.critical("차트 번역 실패 및 1단계 임시 파일도 존재하지 않아 결과 반환 불가.")
                         final_output_path_result = None # 결과 없음
                         f_task_log.write("  오류: 1단계 임시 파일도 존재하지 않아 결과 반환 불가.\n")


                f_task_log.write("\n--- 전체 프레젠테이션 번역 작업 완료 ---\n")
                return final_output_path_result

        except Exception as e_main_handler:
            err_msg_main = f"프레젠테이션 번역 중 PptxHandler 최상위 수준에서 예기치 않은 오류 발생: {e_main_handler}"
            logger.critical(err_msg_main, exc_info=True)
            # 작업 로그 파일이 열려있다면 기록 시도
            if 'f_task_log' in locals() and f_task_log and not f_task_log.closed:
                try:
                    f_task_log.write(f"치명적 오류 (PptxHandler 전역): {err_msg_main}\n상세 정보: {traceback.format_exc()}\n")
                except Exception as e_log_fatal:
                     logger.error(f"치명적 오류 로그 기록 실패: {e_log_fatal}")
            return None # 오류 발생 시 None 반환
        finally:
            # 임시 파일 및 폴더 정리
            if temp_pptx_for_chart_translation_path and os.path.exists(temp_pptx_for_chart_translation_path):
                try:
                    # 최종 결과 파일이 이 임시 파일과 동일한 경로가 아닐 때만 삭제
                    # (예: 차트 번역 실패로 1단계 임시 파일이 최종 결과가 된 경우 삭제 방지)
                    if final_output_path_result != temp_pptx_for_chart_translation_path :
                        os.remove(temp_pptx_for_chart_translation_path)
                        logger.debug(f"1단계 임시 PPTX 파일 '{temp_pptx_for_chart_translation_path}' 삭제 완료.")
                    else:
                        logger.debug(f"1단계 임시 PPTX 파일 '{temp_pptx_for_chart_translation_path}'은 최종 결과이므로 삭제하지 않음.")
                except Exception as e_clean_temp_file:
                    logger.warning(f"1단계 임시 PPTX 파일 '{temp_pptx_for_chart_translation_path}' 삭제 중 오류: {e_clean_temp_file}")
            
            if os.path.exists(temp_dir_for_pptx_handler):
                try:
                    shutil.rmtree(temp_dir_for_pptx_handler)
                    logger.debug(f"PptxHandler 주 임시 디렉토리 '{temp_dir_for_pptx_handler}' 삭제 완료.")
                except Exception as e_clean_main_dir:
                    logger.warning(f"PptxHandler 주 임시 디렉토리 '{temp_dir_for_pptx_handler}' 삭제 중 오류: {e_clean_main_dir}")