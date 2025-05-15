from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
# from pptx.enum.text import MSO_VERTICAL_ANCHOR # 현재 직접 사용 안 함
import os
import io
import logging
import re
from datetime import datetime
import hashlib

logger = logging.getLogger(__name__)

# OCR 텍스트 유효성 검사를 위한 설정 (이전과 동일)
MIN_OCR_TEXT_LENGTH_TO_VALIDATE = 2
MIN_MEANINGFUL_CHAR_RATIO = 0.4 # 의미 있는 문자 비율 임계값
MAX_CONSECUTIVE_STRANGE_SPECIAL_CHARS = 3

# --- 번역 스킵 조건 함수 ---
def should_skip_translation(text: str) -> bool:
    """주어진 텍스트가 번역할 가치가 없는지(숫자, 특수문자, 깨진 문자 등) 확인합니다."""
    if not text or not text.strip():
        return True # 비었거나 공백만 있으면 스킵

    # 1. 의미 있는 문자(알파벳, 한글, 일본어, 한자 등)가 전혀 없는지 확인
    # 정규식: 최소 하나의 한글, 일본어(히라가나/가타카나/한자), 또는 알파벳 문자를 포함하는지
    # \w 는 숫자도 포함하므로, 여기서는 명시적으로 문자 범위를 지정
    meaningful_char_pattern = re.compile(r'[a-zA-Z\u3040-\u30ff\u3131-\uD79D\u4e00-\u9fff]') # 영어 알파벳, 일본어, 한글, 한자
    if not meaningful_char_pattern.search(text):
        logger.info(f"번역 스킵 (의미 있는 문자 없음): '{text}'")
        return True

    # 2. 숫자, 일반적인 특수문자(구두점 등), 공백으로만 구성되어 있는지 확인
    # 숫자, 정의된 특수문자, 공백 외의 문자가 있다면 번역 대상이 될 수 있음.
    # 여기서는 "의미 있는 문자"가 위에서 하나라도 발견되었다면,
    # 이 조건은 덜 엄격하게 적용하거나, 특정 패턴(예: 단순 날짜 "2024-05-15")을 추가로 제외할 수 있음.
    # 현재는 위 meaningful_char_pattern으로 1차 필터링 후, 아래 조건은 좀 더 관대하게.

    # 추가적으로, 전체 문자열에서 의미있는 문자 비율이 너무 낮은 경우도 스킵 (is_ocr_text_valid 와 유사)
    stripped_text = text.strip()
    text_len = len(stripped_text)
    if text_len == 0: return True # 이미 위에서 처리됨

    meaningful_chars_count = 0
    for char_obj in stripped_text: # char -> char_obj로 변경 (의미 명확화)
        if meaningful_char_pattern.search(char_obj):
            meaningful_chars_count += 1
    
    # 의미 있는 문자가 하나도 없으면 위에서 이미 걸러졌으므로, 여기서는 비율만 체크
    if text_len > 0 and (meaningful_chars_count / text_len) < MIN_MEANINGFUL_CHAR_RATIO:
        # 단, 아주 짧은 문자열(예: 2~3글자)은 이 비율만으로 판단하기 어려울 수 있음.
        # 여기서는 is_ocr_text_valid에서 사용한 MIN_OCR_TEXT_LENGTH_TO_VALIDATE 보다 긴 경우에만 엄격히 적용
        if text_len > MAX_CONSECUTIVE_STRANGE_SPECIAL_CHARS + 2: # 임의의 길이 (예: 5글자 이상)
            logger.info(f"번역 스킵 (의미 있는 문자 비율 낮음 {meaningful_chars_count / text_len:.2f}): '{text}'")
            return True
            
    # 예시: 숫자와 특정 특수문자로만 이루어진 경우 (더 정교한 패턴 필요 가능)
    # ^[\d\s.,;:?!()\[\]{}'"‘’“”%+-=*\/\\<>@#$&^|_~]+$  <-- 이 패턴은 너무 많은 것을 스킵할 수 있음
    # 여기서는 meaningful_char_pattern에 걸리지 않으면 위에서 스킵되므로, 이 부분은 생략하거나 매우 제한적으로.

    return False # 위 조건에 해당하지 않으면 번역 대상

def is_ocr_text_valid(text: str) -> bool: # 이 함수는 OCR 결과 자체의 유효성 검사용
    # (이전과 동일하게 유지 - 번역 가치와는 별개로 OCR 인식 결과가 쓰레기값인지 판단)
    stripped_text = text.strip();
    if not stripped_text: return False
    text_len = len(stripped_text)
    if text_len < MIN_OCR_TEXT_LENGTH_TO_VALIDATE:
        if re.search(r'[\w\u3040-\u30ff\u3131-\uD79D\u4e00-\u9fff]', stripped_text) or re.fullmatch(r"""[.,;:?!()\[\]{}'"‘’“”]""", stripped_text): return True
        else: logger.info(f"OCR 유효성 스킵 (짧고 유효문자X): '{stripped_text}'"); return False
    
    meaningful_char_pattern = re.compile(r'[a-zA-Z\u3040-\u30ff\u3131-\uD79D\u4e00-\u9fff]')
    meaningful_chars_count = 0
    for char_obj in stripped_text:
        if meaningful_char_pattern.search(char_obj): # \w 대신 명시적 문자 사용
            meaningful_chars_count += 1

    if text_len > 0 and (meaningful_chars_count / text_len) < MIN_MEANINGFUL_CHAR_RATIO:
        logger.info(f"OCR 유효성 스킵 (의미문자 비율 낮음 {meaningful_chars_count / text_len:.2f}): '{stripped_text}'")
        return False
        
    strange_pattern = f"(?:[^\w\s\u3040-\u30ff\u3131-\uD79D\u4e00-\u9fff.,;:?!()\[\]{{}}'\"‘’“”]){{{MAX_CONSECUTIVE_STRANGE_SPECIAL_CHARS},}}"
    if re.search(strange_pattern, stripped_text):
        logger.info(f"OCR 유효성 스킵 (이상한 특수문자 연속): '{stripped_text}'")
        return False
    return True

class PptxHandler:
    def __init__(self):
        pass

    def get_file_info(self, file_path, ocr_handler): # 단순화하지 않은 원래 버전
        logger.info(f"파일 정보 분석 시작 (OCR 호출 포함): {file_path}")
        info = {"slide_count": 0, "text_elements": 0, "image_elements": 0, "image_elements_with_text": 0}
        try:
            prs = Presentation(file_path)
            info["slide_count"] = len(prs.slides)
            for slide_idx, slide in enumerate(prs.slides):
                for shape_idx, shape in enumerate(slide.shapes): # 사용하지 않는 shape_idx 제거 가능
                    element_name = shape.name or f"S{slide_idx+1}_Id{shape.shape_id}"
                    if shape.has_text_frame and shape.text_frame.text and shape.text_frame.text.strip():
                        info["text_elements"] += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        info["image_elements"] += 1
                        if ocr_handler:
                            try:
                                image_bytes = shape.image.blob
                                if ocr_handler.has_text_in_image_bytes(image_bytes): # 실제 OCR 호출
                                    info["image_elements_with_text"] += 1
                            except AttributeError:
                                logger.warning(f"S{slide_idx+1}-Shape '{element_name}' (ID:{shape.shape_id}): 그림 shape이나 image 속성 없음 (get_file_info).")
                            except Exception as e:
                                logger.warning(f"S{slide_idx+1}-Img '{element_name}' (ID:{shape.shape_id}): OCR 사전 검사 중 오류 (get_file_info): {e}")
                    elif shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if cell.text_frame.text and cell.text_frame.text.strip():
                                    info["text_elements"] += 1
            logger.info(f"파일 분석 완료: Slides:{info['slide_count']}, Text:{info['text_elements']}, Img:{info['image_elements']}, ImgWithText:{info['image_elements_with_text']}")
        except Exception as e:
            logger.error(f"'{file_path}' 파일 정보 분석 오류: {e}", exc_info=True)
        return info

    def _get_text_style(self, run):
        # ... (이전과 동일) ...
        font = run.font; style = {'name': font.name, 'size': font.size if font.size else Pt(11), 'bold': font.bold, 'italic': font.italic, 'underline': font.underline, 'color_rgb': None, 'color_theme_index': None, 'color_brightness': 0.0, 'language_id': font.language_id, 'hyperlink_address': run.hyperlink.address if run.hyperlink else None}
        if hasattr(font.color, 'brightness') and font.color.brightness is not None: style['color_brightness'] = font.color.brightness
        if hasattr(font.color, 'type'):
            color_type = font.color.type
            if color_type == MSO_COLOR_TYPE.RGB: style['color_rgb'] = font.color.rgb
            elif color_type == MSO_COLOR_TYPE.SCHEME: style['color_theme_index'] = font.color.theme_color
        elif font.color.rgb is not None: style['color_rgb'] = font.color.rgb
        return style

    def _apply_text_style(self, run, style):
        # ... (이전과 동일, 줄바꿈 문제 해결된 버전) ...
        font = run.font
        if style.get('name'): font.name = style['name']
        if style.get('size'): font.size = style['size']
        if style.get('bold') is not None: font.bold = style['bold']
        if style.get('italic') is not None: font.italic = style['italic']
        if style.get('underline') is not None: font.underline = style['underline']
        try:
            if style.get('color_rgb'): font.color.rgb = RGBColor(*(int(c) for c in style['color_rgb']))
            elif style.get('color_theme_index') is not None:
                font.color.theme_color = style['color_theme_index']
                brightness_val = float(style.get('color_brightness', 0.0)); font.color.brightness = max(-1.0, min(1.0, brightness_val))
        except Exception as e: logger.warning(f"텍스트 색상 적용 오류: {e}")
        if style.get('language_id') is not None: font.language_id = style['language_id']
        if style.get('hyperlink_address'):
            try:
                hlink = run.hyperlink
                if hlink: hlink.address = style['hyperlink_address']
            except Exception as e: logger.warning(f"하이퍼링크 주소 적용 오류: {e}")

    def translate_presentation(self, file_path, src_lang, tgt_lang, translator, ocr_handler,
                               model_name, ollama_service, font_code_for_render,
                               task_log_filepath,
                               progress_callback=None, stop_event=None):
        try:
            with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f_task_log.write(f"--- 프레젠테이션 번역 로그 시작: {os.path.basename(file_path)} ---\n")
                f_task_log.write(f"시작 시간: {start_time_str}\n원본 파일: {file_path}\n소스 언어: {src_lang}, 대상 언어: {tgt_lang}, 모델: {model_name}\n")
                f_task_log.write(f"OCR 사용: {'예' if ocr_handler else '아니오'}\n결과 렌더링 폰트 코드: {font_code_for_render}\n\n")
                logger.info(f"'{os.path.basename(file_path)}' 프레젠테이션 번역 시작. 로그: {task_log_filepath}")

                try:
                    prs = Presentation(file_path)
                    safe_tgt_lang = "".join(c if c.isalnum() else "_" for c in tgt_lang)
                    output_filename = os.path.splitext(file_path)[0] + f"_{safe_tgt_lang}_translated.pptx"

                    elements_map = []
                    f_task_log.write("--- 번역 대상 요소 분석 시작 (shape_id 사용, OCR 사전 검사 포함) ---\n")
                    for slide_idx, slide in enumerate(prs.slides):
                        if stop_event and stop_event.is_set(): break
                        for shape_in_slide in slide.shapes:
                            if stop_event and stop_event.is_set(): break
                            element_name = shape_in_slide.name or f"S{slide_idx+1}_Id{shape_in_slide.shape_id}"
                            item_to_add = None
                            if shape_in_slide.has_text_frame and shape_in_slide.text_frame.text and shape_in_slide.text_frame.text.strip():
                                item_to_add = {'type': 'text', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'name': element_name}
                            elif shape_in_slide.shape_type == MSO_SHAPE_TYPE.PICTURE and ocr_handler:
                                try:
                                    # elements_map 생성 시점에 OCR 대상 여부 판단 (원래 방식)
                                    if ocr_handler.has_text_in_image_bytes(shape_in_slide.image.blob):
                                        item_to_add = {'type': 'image', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'name': element_name}
                                    else:
                                        f_task_log.write(f"  S{slide_idx+1}-Img '{element_name}' (ID:{shape_in_slide.shape_id}): OCR 사전 검사 결과 텍스트 없음/처리 불가.\n")
                                except AttributeError:
                                     f_task_log.write(f"  S{slide_idx+1}-Shape '{element_name}' (ID:{shape_in_slide.shape_id}): 그림 shape이나 image 속성 없음 (elem_map).\n")
                                except Exception as e_scan: # Pillow OSError 등 포함
                                     f_task_log.write(f"  S{slide_idx+1}-Img '{element_name}' (ID:{shape_in_slide.shape_id}): OCR 사전 검사 중 오류 (elem_map): {e_scan}\n")
                            elif shape_in_slide.has_table:
                                for r_idx, row in enumerate(shape_in_slide.table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        if cell.text_frame.text and cell.text_frame.text.strip():
                                            elements_map.append({'type': 'table_cell', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'row_idx': r_idx, 'col_idx': c_idx,'name': element_name})
                            if item_to_add: elements_map.append(item_to_add)
                    
                    total_elements_to_translate = len(elements_map)
                    f_task_log.write(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견.\n--- 요소 분석 완료 ---\n\n")
                    logger.info(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견.")

                    if total_elements_to_translate == 0 and not (stop_event and stop_event.is_set()):
                        prs.save(output_filename); return output_filename # 내용 없을 시

                    translated_count = 0
                    f_task_log.write("--- 번역 작업 시작 ---\n")
                    for item_info in elements_map:
                        # ... (current_shape_obj 찾기, 기본 정보/상세 정보 로깅은 이전 답변과 동일하게 유지) ...
                        if stop_event and stop_event.is_set(): break
                        slide_idx = item_info['slide_idx']; target_shape_id = item_info['shape_id']; element_name = item_info['name']
                        current_slide_obj = prs.slides[slide_idx]; current_shape_obj = None
                        for s_obj in current_slide_obj.shapes:
                            if hasattr(s_obj, 'shape_id') and s_obj.shape_id == target_shape_id: current_shape_obj = s_obj; break
                        if current_shape_obj is None: f_task_log.write(f"  오류: S{slide_idx+1}-ID{target_shape_id} shape 찾지 못함. 건너뜀.\n"); continue
                        log_shape_details = (f"[S{slide_idx+1}] 요소 처리 시작: '{element_name}' (ID: {target_shape_id}), 타입: {item_info['type']}, 객체: {current_shape_obj}, "
                                             f"위치: L{current_shape_obj.left}, T{current_shape_obj.top}, W{current_shape_obj.width}, H{current_shape_obj.height}\n")
                        logger.debug(log_shape_details.strip()); f_task_log.write(log_shape_details)
                        current_text_for_progress = ""

                        item_type = item_info['type']

                        if item_type == 'text' or item_type == 'table_cell':
                            # --- 이전 답변의 안정화된 텍스트 처리 로직 (줄바꿈 문제 해결된 버전) ---
                            # (이 부분은 길어서 생략, 이전 답변의 텍스트 처리 로직을 여기에 넣어주세요)
                            # ... (이전 답변의 최종 텍스트 처리 코드 사용)
                            text_frame = None 
                            if item_type == 'text' and current_shape_obj.has_text_frame: text_frame = current_shape_obj.text_frame
                            elif item_type == 'table_cell' and current_shape_obj.has_table:
                                try: cell = current_shape_obj.table.cell(item_info['row_idx'], item_info['col_idx']); text_frame = cell.text_frame
                                except IndexError: logger.error(f"테이블 셀 접근 오류: {element_name}"); f_task_log.write(f"  오류: 테이블 셀 접근 실패 {element_name}\n"); text_frame = None
                            if text_frame and text_frame.text and text_frame.text.strip():
                                original_text = text_frame.text; current_text_for_progress = original_text
                                if progress_callback: progress_callback(slide_idx + 1, f"{item_info['type']} ({element_name})", translated_count, total_elements_to_translate, original_text)
                                translated_text = translator.translate_text(original_text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=False)
                                f_task_log.write(f"  [번역 전] \"{original_text.replace(chr(10), ' / ')}\" -> [번역 후] \"{translated_text.replace(chr(10), ' / ')}\"\n")
                                if "오류:" not in translated_text and translated_text.strip():
                                    # 스타일 복원 및 텍스트 삽입 로직 (줄바꿈 해결된 버전)
                                    original_paragraphs_info = []
                                    for para in text_frame.paragraphs: original_paragraphs_info.append({'runs': [{'text': run.text, 'style': self._get_text_style(run)} for run in para.runs], 'alignment': para.alignment, 'level': para.level, 'space_before': para.space_before, 'space_after': para.space_after, 'line_spacing': para.line_spacing})
                                    translated_lines_raw = translated_text.splitlines(); first_content_line_idx = 0
                                    for idx, line in enumerate(translated_lines_raw):
                                        if line.strip(): first_content_line_idx = idx; break
                                    processed_translated_lines = [line.strip() for line in translated_lines_raw[first_content_line_idx:] if line.strip()]
                                    if not processed_translated_lines: processed_translated_lines = [translated_text.strip()] if translated_text.strip() else [" "]
                                    num_new_paras = len(processed_translated_lines); num_orig_paras = len(text_frame.paragraphs)
                                    for i in range(max(num_new_paras, num_orig_paras)):
                                        if i < num_new_paras:
                                            new_line_text = processed_translated_lines[i]
                                            p = text_frame.paragraphs[i] if i < num_orig_paras else text_frame.add_paragraph()
                                            if i < num_orig_paras: # 기존 단락 내용 지우기
                                                for run_obj in list(p.runs): p._p.remove(run_obj._r)
                                            para_template = original_paragraphs_info[0] if i == 0 and original_paragraphs_info else original_paragraphs_info[min(i, len(original_paragraphs_info)-1)] if original_paragraphs_info else None
                                            if para_template: # 스타일 적용
                                                if para_template.get('alignment'): p.alignment = para_template['alignment']
                                                p.level = para_template.get('level', 0)
                                                p.space_before = para_template.get('space_before') if i == 0 and para_template.get('space_before') is not None else Pt(0) if i == 0 else para_template.get('space_before', Pt(0))
                                                p.space_after = para_template.get('space_after', Pt(0))
                                                p.line_spacing = para_template.get('line_spacing')
                                            else: p.space_before = Pt(0); p.space_after = Pt(0) # 기본값
                                            run = p.add_run(); run.text = new_line_text if new_line_text else " "
                                            if para_template and para_template.get('runs'): self._apply_text_style(run, para_template['runs'][0]['style'])
                                        elif i < num_orig_paras: text_frame.paragraphs[i].clear() # 남는 단락 비우기
                                else: f_task_log.write(f"  -> 텍스트 번역 실패 또는 빈 결과: {translated_text}\n")
                            translated_count += 1

                        elif item_type == 'image' and ocr_handler: # 'image' 타입 (elements_map 생성 시 OCR 대상 확정)
                            try:
                                from PIL import Image as PILImage
                                # .image.blob 접근은 elements_map 생성 시 이미 성공했다고 가정
                                image_bytes = current_shape_obj.image.blob
                                img_pil_original = PILImage.open(io.BytesIO(image_bytes))
                                img_pil_for_ocr = img_pil_original.convert("RGB") # Pillow OSError는 ocr_handler에서 잡음

                                # --- 이미지 해시 로깅 ---
                                try:
                                    hasher = hashlib.md5(); img_byte_arr = io.BytesIO()
                                    img_pil_for_ocr.save(img_byte_arr, format='PNG')
                                    hasher.update(img_byte_arr.getvalue()); img_hash = hasher.hexdigest()
                                    hash_log_msg = f"    OCR 대상 RGB 이미지 해시 (MD5): {img_hash} (Shape ID: {target_shape_id})\n"
                                    logger.debug(hash_log_msg.strip()); f_task_log.write(hash_log_msg)
                                except Exception as e_hash: logger.warning(f"    이미지 해시 생성 중 오류: {e_hash}")
                                
                                ocr_results = ocr_handler.ocr_image(img_pil_for_ocr)
                                # ... (이하 이미지 OCR 결과 처리, 번역, 렌더링, 교체 로직은 이전 답변의 shape_id 사용 버전과 동일) ...
                                # (이 부분도 길어서 생략, 이전 답변의 해당 로직을 여기에 넣어주세요)
                                if ocr_results:
                                    f_task_log.write(f"      이미지 내 OCR 텍스트 {len(ocr_results)}개 발견.\n")
                                    edited_image_pil = img_pil_original.copy(); any_ocr_text_translated_and_rendered = False
                                    for ocr_idx, (box, (text, confidence)) in enumerate(ocr_results):
                                        if stop_event and stop_event.is_set(): break
                                        current_text_for_progress = text
                                        if not is_ocr_text_valid(text): f_task_log.write(f"        OCR Text [{ocr_idx+1}]: \"{text.replace(chr(10), ' ')}\" (Conf: {confidence:.2f}) -> [스킵됨 - 유효X]\n"); continue
                                        translated_ocr_text = translator.translate_text(text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=True)
                                        f_task_log.write(f"        OCR Text [{ocr_idx+1}]: \"{text.replace(chr(10), ' ')}\" (Conf: {confidence:.2f}) -> \"{translated_ocr_text.replace(chr(10), ' ')}\"\n")
                                        if "오류:" not in translated_ocr_text and translated_ocr_text.strip():
                                            try: edited_image_pil = ocr_handler.render_translated_text_on_image(edited_image_pil, box, translated_ocr_text, font_code_for_render=font_code_for_render, original_text=text); any_ocr_text_translated_and_rendered = True
                                            except Exception as e_render: f_task_log.write(f"          오류: OCR 텍스트 렌더링 실패: {e_render}\n"); logger.error(f"OCR 렌더링 실패: {e_render}", exc_info=True)
                                    if stop_event and stop_event.is_set(): break
                                    if any_ocr_text_translated_and_rendered:
                                        image_stream = io.BytesIO(); save_format = img_pil_original.format if img_pil_original.format and img_pil_original.format.upper() in ['JPEG', 'PNG'] else 'PNG'
                                        edited_image_pil.save(image_stream, format=save_format); image_stream.seek(0)
                                        left, top, width, height = current_shape_obj.left, current_shape_obj.top, current_shape_obj.width, current_shape_obj.height
                                        shape_element_to_remove = current_shape_obj.element; sp_parent = shape_element_to_remove.getparent()
                                        if sp_parent is not None:
                                            sp_parent.remove(shape_element_to_remove)
                                            new_pic_shape = current_slide_obj.shapes.add_picture(image_stream, left, top, width=width, height=height)
                                            f_task_log.write(f"      이미지 '{element_name}' (ID:{target_shape_id}) 교체 완료. 새 ID: {new_pic_shape.shape_id}\n")
                                        else: f_task_log.write(f"      오류: 이미지 '{element_name}' 부모 못찾아 교체 실패.\n")
                                    else: f_task_log.write(f"      이미지 '{element_name}' 변경 없음 (유효/번역/렌더링된 OCR 텍스트 없음).\n")
                                else: f_task_log.write(f"      이미지 '{element_name}' 내 OCR 텍스트 발견되지 않음.\n")

                            except AttributeError:
                                 f_task_log.write(f"    경고: '{element_name}' (ID:{target_shape_id}) 그림 shape 처리 중 image 속성 또는 blob 접근 불가. 건너뜀.\n")
                            except OSError as e_os_img_proc:
                                f_task_log.write(f"    이미지 처리 중 Pillow OSError (ID: {target_shape_id}): {e_os_img_proc}\n")
                                logger.error(f"이미지 처리 중 Pillow OSError (ID: {target_shape_id}): {e_os_img_proc}", exc_info=False)
                            except Exception as e_img_proc_detail:
                                f_task_log.write(f"    이미지 '{element_name}' (ID:{target_shape_id}) 처리 중 예기치 않은 오류: {e_img_proc_detail}\n")
                                logger.error(f"이미지 처리 오류 (ID: {target_shape_id}): {e_img_proc_detail}", exc_info=True)
                            translated_count += 1
                        
                        f_task_log.write("\n")
                        if progress_callback and not (stop_event and stop_event.is_set()):
                            progress_callback(slide_idx + 1, item_type, translated_count, total_elements_to_translate, current_text_for_progress)
                    
                    f_task_log.write("--- 번역 작업 완료 ---\n")
                    if stop_event and stop_event.is_set(): prs.save(output_filename.replace(".pptx", "_stopped.pptx")); return output_filename.replace(".pptx", "_stopped.pptx")
                    prs.save(output_filename); return output_filename
                except Exception as e_translate_general:
                    logger.error(f"프레젠테이션 번역 중 심각한 오류: {e_translate_general}", exc_info=True); return None
        except IOError as e_file_io:
            logger.error(f"작업 로그 파일 처리 오류 ({task_log_filepath}): {e_file_io}", exc_info=True); return None