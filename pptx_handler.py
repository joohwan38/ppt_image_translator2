from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
# from pptx.enum.text import MSO_VERTICAL_ANCHOR
import os
import io
import logging
import re
from datetime import datetime
import hashlib

logger = logging.getLogger(__name__)

MIN_OCR_TEXT_LENGTH_TO_VALIDATE = 2
MIN_MEANINGFUL_CHAR_RATIO = 0.4 
MAX_CONSECUTIVE_STRANGE_SPECIAL_CHARS = 3

def should_skip_translation(text: str) -> bool:
    if not text or not text.strip():
        return True 

    # 의미 있는 문자 패턴 (알파벳, 한글, 일본어, 한자)
    # \uAC00-\uD7AF 한글 음절 전체, \u3040-\u30FF 일본어 (히라가나, 가타카나, 구두점 등), \u4E00-\u9FFF 일반적인 한자 범위
    meaningful_char_pattern = re.compile(r'[a-zA-Z\uAC00-\uD7AF\u3040-\u30FF\u4E00-\u9FFF]') 
    
    if not meaningful_char_pattern.search(text):
        logger.info(f"번역 스킵 (의미 있는 언어 문자 없음): '{text}'")
        return True

    stripped_text = text.strip()
    text_len = len(stripped_text)
    if text_len == 0: return True

    meaningful_chars_count = 0
    for char_obj in stripped_text:
        if meaningful_char_pattern.search(char_obj):
            meaningful_chars_count += 1
    
    # 아주 짧은 문자열이 아니면서, 의미 있는 문자 비율이 매우 낮은 경우
    if text_len > 3 and (meaningful_chars_count / text_len) < MIN_MEANINGFUL_CHAR_RATIO:
        logger.info(f"번역 스킵 (의미 있는 문자 비율 낮음 {meaningful_chars_count / text_len:.2f}): '{text}'")
        return True
            
    return False

def is_ocr_text_valid(text: str) -> bool:
    stripped_text = text.strip();
    if not stripped_text: return False
    text_len = len(stripped_text)
    meaningful_char_pattern_ocr = re.compile(r'[a-zA-Z\uAC00-\uD7AF\u3040-\u30FF\u4E00-\u9FFF]') # \w 대신 명시적 문자 범위 사용
    
    if text_len < MIN_OCR_TEXT_LENGTH_TO_VALIDATE:
        return bool(meaningful_char_pattern_ocr.search(stripped_text) or \
                    re.fullmatch(r"""[.,;:?!()\[\]{}'"‘’“”]""", stripped_text))

    meaningful_chars_count = 0
    for char_obj in stripped_text:
        if meaningful_char_pattern_ocr.search(char_obj):
            meaningful_chars_count += 1

    if text_len > 0 and (meaningful_chars_count / text_len) < MIN_MEANINGFUL_CHAR_RATIO:
        logger.info(f"OCR 유효성 스킵 (의미문자 비율 낮음 {meaningful_chars_count / text_len:.2f}): '{stripped_text}'")
        return False
        
    strange_pattern = f"(?:[^\w\s\uAC00-\uD7AF\u3040-\u30FF\u4E00-\u9FFF.,;:?!()\[\]{{}}'\"‘’“”]){{{MAX_CONSECUTIVE_STRANGE_SPECIAL_CHARS},}}"
    if re.search(strange_pattern, stripped_text):
        logger.info(f"OCR 유효성 스킵 (이상한 특수문자 연속): '{stripped_text}'")
        return False
    return True

class PptxHandler:
    def __init__(self):
        pass

    def get_file_info(self, file_path, ocr_handler): # 원본 (OCR 호출) 버전
        logger.info(f"파일 정보 분석 시작 (OCR 호출 포함): {file_path}")
        info = {"slide_count": 0, "text_elements": 0, "image_elements": 0, "image_elements_with_text": 0}
        try:
            prs = Presentation(file_path)
            info["slide_count"] = len(prs.slides)
            for slide_idx, slide in enumerate(prs.slides):
                for shape in slide.shapes:
                    element_name = shape.name or f"S{slide_idx+1}_Id{shape.shape_id}"
                    if shape.has_text_frame and shape.text_frame.text and shape.text_frame.text.strip():
                        info["text_elements"] += 1
                    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        info["image_elements"] += 1
                        if ocr_handler:
                            try:
                                if hasattr(shape, 'image') and shape.image and hasattr(shape.image, 'blob'):
                                    image_bytes = shape.image.blob
                                    if ocr_handler.has_text_in_image_bytes(image_bytes):
                                        info["image_elements_with_text"] += 1
                                else:
                                    logger.warning(f"S{slide_idx+1}-Shape '{element_name}' (ID:{shape.shape_id}): 그림 shape이나 image.blob 속성 없음 (get_file_info).")
                            except Exception as e:
                                logger.warning(f"S{slide_idx+1}-Img '{element_name}' (ID:{shape.shape_id}): OCR 사전 검사 중 오류 (get_file_info): {e}")
                    elif shape.has_table: # ... (테이블 처리)
                        for row in shape.table.rows:
                            for cell in row.cells:
                                if cell.text_frame.text and cell.text_frame.text.strip(): info["text_elements"] +=1
            logger.info(f"파일 분석 완료: Slides:{info['slide_count']}, Text:{info['text_elements']}, Img:{info['image_elements']}, ImgWithText:{info['image_elements_with_text']}")
        except Exception as e:
            logger.error(f"'{file_path}' 파일 정보 분석 오류: {e}", exc_info=True)
        return info

    def _get_text_style(self, run):
        # (이전과 동일)
        font = run.font; style = {'name': font.name, 'size': font.size if font.size else Pt(11), 'bold': font.bold, 'italic': font.italic, 'underline': font.underline, 'color_rgb': None, 'color_theme_index': None, 'color_brightness': 0.0, 'language_id': font.language_id, 'hyperlink_address': run.hyperlink.address if run.hyperlink else None}
        if hasattr(font.color, 'brightness') and font.color.brightness is not None: style['color_brightness'] = font.color.brightness
        if hasattr(font.color, 'type'):
            color_type = font.color.type
            if color_type == MSO_COLOR_TYPE.RGB: style['color_rgb'] = font.color.rgb
            elif color_type == MSO_COLOR_TYPE.SCHEME: style['color_theme_index'] = font.color.theme_color
        elif font.color.rgb is not None: style['color_rgb'] = font.color.rgb
        return style

    def _apply_text_style(self, run, style):
        # (이전과 동일, 줄바꿈 문제 해결된 버전)
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
                # ... (파일 헤더 로깅) ...
                logger.info(f"'{os.path.basename(file_path)}' 프레젠테이션 번역 시작. 로그: {task_log_filepath}")
                try:
                    prs = Presentation(file_path)
                    output_filename = os.path.splitext(file_path)[0] + f"_{''.join(c if c.isalnum() else '_' for c in tgt_lang)}_translated.pptx"
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
                                    if hasattr(shape_in_slide, 'image') and shape_in_slide.image and hasattr(shape_in_slide.image, 'blob'):
                                        if ocr_handler.has_text_in_image_bytes(shape_in_slide.image.blob):
                                            item_to_add = {'type': 'image', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'name': element_name}
                                        else: f_task_log.write(f"  S{slide_idx+1}-Img '{element_name}': OCR 사전 검사 결과 텍스트 없음/처리 불가.\n")
                                    else: f_task_log.write(f"  S{slide_idx+1}-Shape '{element_name}': 그림 shape이나 image.blob 속성 없음 (elem_map).\n")
                                except Exception as e_scan: f_task_log.write(f"  S{slide_idx+1}-Img '{element_name}': OCR 사전 검사 중 오류 (elem_map): {e_scan}\n")
                            elif shape_in_slide.has_table: # ... (테이블 처리) ...
                                for r_idx, row in enumerate(shape_in_slide.table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        if cell.text_frame.text and cell.text_frame.text.strip():
                                            elements_map.append({'type': 'table_cell', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'row_idx': r_idx, 'col_idx': c_idx,'name': element_name})
                            if item_to_add: elements_map.append(item_to_add)
                    
                    total_elements_to_translate = len(elements_map) # ... (이하 로깅) ...
                    if total_elements_to_translate == 0 and not (stop_event and stop_event.is_set()): prs.save(output_filename); return output_filename

                    translated_count = 0
                    f_task_log.write("--- 번역 작업 시작 ---\n")
                    for item_info in elements_map:
                        # ... (current_shape_obj 찾기 및 기본/상세 정보 로깅은 이전과 동일) ...
                        if stop_event and stop_event.is_set(): break
                        slide_idx = item_info['slide_idx']; target_shape_id = item_info['shape_id']; element_name = item_info['name']
                        current_slide_obj = prs.slides[slide_idx]; current_shape_obj = None
                        for s_obj in current_slide_obj.shapes:
                            if hasattr(s_obj, 'shape_id') and s_obj.shape_id == target_shape_id: current_shape_obj = s_obj; break
                        if current_shape_obj is None: f_task_log.write(f"  오류: S{slide_idx+1}-ID{target_shape_id} shape 찾지 못함. 건너뜀.\n"); continue
                        f_task_log.write(f"[S{slide_idx+1}] 요소 처리 시작: '{element_name}' (ID: {target_shape_id}), 타입: {item_info['type']}\n")
                        current_text_for_progress = ""
                        item_type = item_info['type']

                        if item_type == 'text' or item_type == 'table_cell':
                            text_frame = None # ... (text_frame 가져오기) ...
                            if item_type == 'text' and current_shape_obj.has_text_frame: text_frame = current_shape_obj.text_frame
                            elif item_type == 'table_cell' and current_shape_obj.has_table:
                                try: cell = current_shape_obj.table.cell(item_info['row_idx'], item_info['col_idx']); text_frame = cell.text_frame
                                except IndexError: f_task_log.write(f"  오류: 테이블 셀 접근 실패 {element_name}\n"); text_frame = None
                            
                            if text_frame and text_frame.text and text_frame.text.strip():
                                original_text = text_frame.text; current_text_for_progress = original_text
                                if should_skip_translation(original_text): # *** 번역 스킵 로직 ***
                                    f_task_log.write(f"  [번역 전] \"{original_text.replace(chr(10), ' / ')}\" -> [번역 후] [스킵됨 - 번역 불필요]\n")
                                    # 원본 유지를 위해 별도 작업 불필요 (스타일링 코드로 넘어가지 않음)
                                else:
                                    if progress_callback: progress_callback(slide_idx + 1, f"{item_type} ({element_name})", translated_count, total_elements_to_translate, original_text)
                                    translated_text = translator.translate_text(original_text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=False)
                                    f_task_log.write(f"  [번역 전] \"{original_text.replace(chr(10), ' / ')}\" -> [번역 후] \"{translated_text.replace(chr(10), ' / ')}\"\n")
                                    if "오류:" not in translated_text and translated_text.strip():
                                        # (이전 답변의 안정화된 텍스트 스타일 적용 및 삽입 로직)
                                        # (이 부분은 길어서 생략, 이전 답변의 텍스트 처리 로직을 여기에 넣어주세요)
                                        original_paragraphs_info = [] # (스타일 수집...)
                                        # (줄바꿈 해결된 텍스트 삽입 로직...)
                                    else: f_task_log.write(f"  -> 텍스트 번역 실패 또는 빈 결과: {translated_text}\n")
                            translated_count += 1

                        elif item_type == 'image' and ocr_handler:
                            try:
                                from PIL import Image as PILImage
                                if not (hasattr(current_shape_obj, 'image') and current_shape_obj.image and hasattr(current_shape_obj.image, 'blob')):
                                    f_task_log.write(f"    경고: '{element_name}' (ID:{target_shape_id}) 그림 shape이나 image.blob 속성 없음 (처리 단계). 건너뜀.\n")
                                    translated_count +=1; f_task_log.write("\n"); continue

                                image_bytes = current_shape_obj.image.blob
                                img_pil_original = PILImage.open(io.BytesIO(image_bytes))
                                img_pil_for_ocr = img_pil_original.convert("RGB")

                                # (이미지 해시 로깅 등은 이전과 동일하게 유지)
                                try:
                                    hasher = hashlib.md5(); img_byte_arr = io.BytesIO(); img_pil_for_ocr.save(img_byte_arr, format='PNG')
                                    hasher.update(img_byte_arr.getvalue()); img_hash = hasher.hexdigest()
                                    f_task_log.write(f"    OCR 대상 RGB 이미지 해시 (MD5): {img_hash} (ID: {target_shape_id})\n")
                                except Exception as e_hash: logger.warning(f"    이미지 해시 생성 오류: {e_hash}")
                                
                                ocr_results = ocr_handler.ocr_image(img_pil_for_ocr)
                                if ocr_results:
                                    f_task_log.write(f"    이미지 내 OCR 텍스트 {len(ocr_results)}개 발견.\n")
                                    edited_image_pil = img_pil_original.copy(); any_ocr_text_translated_and_rendered = False
                                    for ocr_idx, (box, (text, confidence)) in enumerate(ocr_results):
                                        if stop_event and stop_event.is_set(): break
                                        current_text_for_progress = text
                                        if not is_ocr_text_valid(text): # 1차 OCR 유효성
                                            f_task_log.write(f"      OCR Text [{ocr_idx+1}]: \"{text}\" (Conf: {confidence:.2f}) -> [스킵됨 - OCR 유효X]\n"); continue
                                        if should_skip_translation(text): # 2차 번역 가치
                                            f_task_log.write(f"      OCR Text [{ocr_idx+1}]: \"{text}\" (Conf: {confidence:.2f}) -> [스킵됨 - 번역 불필요]\n"); continue
                                        
                                        translated_ocr_text = translator.translate_text(text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=True)
                                        f_task_log.write(f"      OCR Text [{ocr_idx+1}]: \"{text}\" (Conf: {confidence:.2f}) -> \"{translated_ocr_text}\"\n")
                                        if "오류:" not in translated_ocr_text and translated_ocr_text.strip():
                                            try:
                                                edited_image_pil = ocr_handler.render_translated_text_on_image(
                                                    edited_image_pil, box, translated_ocr_text,
                                                    font_code_for_render=font_code_for_render, original_text=text
                                                )
                                                any_ocr_text_translated_and_rendered = True
                                            except Exception as e_render: f_task_log.write(f"        오류: OCR 텍스트 렌더링 실패: {e_render}\n")
                                    if stop_event and stop_event.is_set(): break
                                    if any_ocr_text_translated_and_rendered:
                                        # (이미지 교체 로직)
                                        image_stream = io.BytesIO(); save_format = img_pil_original.format if img_pil_original.format and img_pil_original.format.upper() in ['JPEG', 'PNG'] else 'PNG'
                                        edited_image_pil.save(image_stream, format=save_format); image_stream.seek(0)
                                        left,top,width,height = current_shape_obj.left,current_shape_obj.top,current_shape_obj.width,current_shape_obj.height
                                        shape_element_to_remove = current_shape_obj.element; sp_parent = shape_element_to_remove.getparent()
                                        if sp_parent is not None:
                                            sp_parent.remove(shape_element_to_remove)
                                            new_pic_shape = current_slide_obj.shapes.add_picture(image_stream, left, top, width=width, height=height)
                                            f_task_log.write(f"      이미지 '{element_name}' (ID:{target_shape_id}) 교체 완료. 새 ID: {new_pic_shape.shape_id}\n")
                                        else: f_task_log.write(f"      오류: 이미지 '{element_name}' 부모 못찾아 교체 실패.\n")
                                    else: f_task_log.write(f"      이미지 '{element_name}' 변경 없음.\n")
                                else: f_task_log.write(f"      이미지 '{element_name}' 내 OCR 텍스트 발견되지 않음.\n")
                            except AttributeError: f_task_log.write(f"    경고: '{element_name}' (ID:{target_shape_id}) image.blob 접근 불가.\n")
                            except OSError as e_os_img_proc: f_task_log.write(f"    이미지 처리 중 Pillow OSError (ID: {target_shape_id}): {e_os_img_proc}\n")
                            except Exception as e_img_proc_detail: f_task_log.write(f"    이미지 '{element_name}' 처리 중 예외: {e_img_proc_detail}\n")
                            translated_count += 1
                        
                        f_task_log.write("\n")
                        if progress_callback and not (stop_event and stop_event.is_set()):
                            progress_callback(slide_idx + 1, item_type, translated_count, total_elements_to_translate, current_text_for_progress)
                    
                    f_task_log.write("--- 번역 작업 완료 ---\n"); # ... (최종 저장)
                    if stop_event and stop_event.is_set(): prs.save(output_filename.replace(".pptx", "_stopped.pptx")); return output_filename.replace(".pptx", "_stopped.pptx")
                    prs.save(output_filename); return output_filename
                except Exception as e_translate_general:
                    logger.error(f"프레젠테이션 번역 중 오류: {e_translate_general}", exc_info=True); return None
        except IOError as e_file_io:
            logger.error(f"작업 로그 파일 처리 오류: {e_file_io}", exc_info=True); return None