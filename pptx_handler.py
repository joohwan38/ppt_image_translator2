from pptx import Presentation
from pptx.util import Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR_INDEX
# lxml.etree._Element를 사용하기 위해 추가 (필요시)
# from lxml.etree import _Element as LxmlElement # 실제 사용은 python-pptx 내부에서 이루어짐

import os
import io
import logging
import re
from datetime import datetime
import hashlib
import traceback

logger = logging.getLogger(__name__)

# (should_skip_translation, is_ocr_text_valid, get_file_info, _get_text_style, _apply_text_style 함수는 이전 답변과 동일하게 유지)
MIN_MEANINGFUL_CHAR_RATIO_SKIP = 0.1
MIN_MEANINGFUL_CHAR_RATIO_OCR = 0.1
MEANINGFUL_CHAR_PATTERN = re.compile(r'[a-zA-Z\u3040-\u30ff\u3131-\uD79D\u4e00-\u9fff]')

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

    def _get_text_style(self, run):
        font = run.font
        style = {
            'name': font.name,
            'size': font.size if font.size else Pt(11),
            'bold': font.bold,
            'italic': font.italic,
            'underline': font.underline,
            'color_rgb': None,
            'color_theme_index': None,
            'color_brightness': 0.0,
            'language_id': font.language_id,
            'hyperlink_address': run.hyperlink.address if run.hyperlink else None
        }
        if hasattr(font.color, 'brightness') and font.color.brightness is not None:
            style['color_brightness'] = font.color.brightness
        if hasattr(font.color, 'type') and font.color.type is not None:
            color_type = font.color.type
            if color_type == MSO_COLOR_TYPE.RGB:
                style['color_rgb'] = font.color.rgb
            elif color_type == MSO_COLOR_TYPE.SCHEME:
                style['color_theme_index'] = font.color.theme_color
        elif font.color.rgb is not None:
             style['color_rgb'] = font.color.rgb
        return style

    def _apply_text_style(self, run, style):
        font = run.font
        if style.get('name'): font.name = style['name']
        if style.get('size'): font.size = style['size']
        if style.get('bold') is not None: font.bold = style.get('bold')
        if style.get('italic') is not None: font.italic = style.get('italic')
        if style.get('underline') is not None: font.underline = style.get('underline')
        try:
            if style.get('color_rgb') is not None:
                font.color.rgb = RGBColor(*(int(c) for c in style['color_rgb']))
            elif style.get('color_theme_index') is not None:
                font.color.theme_color = style['color_theme_index']
                brightness_val = float(style.get('color_brightness', 0.0))
                font.color.brightness = max(-1.0, min(1.0, brightness_val))
        except Exception as e:
            logger.warning(f"텍스트 색상 적용 오류: {e}. 스타일: {style}")
        if style.get('language_id') is not None:
            try:
                font.language_id = style['language_id']
            except Exception as e_lang:
                logger.warning(f"폰트 언어 ID ({style['language_id']}) 적용 오류: {e_lang}")
        if style.get('hyperlink_address'):
            try:
                hlink = run.hyperlink
                if hlink:
                    hlink.address = style['hyperlink_address']
            except Exception as e:
                logger.warning(f"하이퍼링크 주소 적용 중 오류: {e}")

    def translate_presentation(self, file_path, src_lang, tgt_lang, translator, ocr_handler,
                               model_name, ollama_service, font_code_for_render,
                               task_log_filepath,
                               progress_callback=None, stop_event=None):
        try:
            with open(task_log_filepath, 'a', encoding='utf-8') as f_task_log:
                # ... (로그 시작 부분 동일) ...
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
                    # ... (elements_map 생성 로직 동일) ...
                    f_task_log.write("--- 번역 대상 요소 분석 시작 (shape_id 사용) ---\n")
                    for slide_idx, slide in enumerate(prs.slides):
                        if stop_event and stop_event.is_set(): break
                        for shape_in_slide in slide.shapes:
                            if stop_event and stop_event.is_set(): break
                            element_name = shape_in_slide.name or f"S{slide_idx+1}_Id{shape_in_slide.shape_id}"
                            item_to_add = None
                            if shape_in_slide.has_text_frame and shape_in_slide.text_frame.text and shape_in_slide.text_frame.text.strip():
                                item_to_add = {'type': 'text', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'name': element_name}
                            elif shape_in_slide.shape_type == MSO_SHAPE_TYPE.PICTURE:
                                if ocr_handler:
                                    item_to_add = {'type': 'image', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'name': element_name}
                                else:
                                    f_task_log.write(f"  S{slide_idx+1}-Img '{element_name}' (ID:{shape_in_slide.shape_id}): OCR 핸들러 없어 번역 대상에서 제외.\n")
                            elif shape_in_slide.has_table:
                                for r_idx, row in enumerate(shape_in_slide.table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        if cell.text_frame.text and cell.text_frame.text.strip():
                                            elements_map.append({'type': 'table_cell', 'slide_idx': slide_idx, 'shape_id': shape_in_slide.shape_id, 'row_idx': r_idx, 'col_idx': c_idx,'name': element_name})
                            if item_to_add: elements_map.append(item_to_add)
                    total_elements_to_translate = len(elements_map)
                    f_task_log.write(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견 (텍스트 및 OCR 시도 이미지).\n--- 요소 분석 완료 ---\n\n")
                    logger.info(f"총 {total_elements_to_translate}개의 번역 대상 요소 발견.")
                    if total_elements_to_translate == 0 and not (stop_event and stop_event.is_set()):
                        prs.save(output_filename)
                        f_task_log.write("번역할 대상 요소가 없어 원본 파일을 복사본으로 저장합니다.\n")
                        return output_filename

                    translated_count = 0
                    f_task_log.write("--- 번역 작업 시작 ---\n")
                    for item_info in elements_map:
                        # ... (shape 찾기 및 기본 로깅 동일) ...
                        if stop_event and stop_event.is_set(): break
                        slide_idx = item_info['slide_idx']; target_shape_id = item_info['shape_id']; element_name = item_info['name']
                        current_slide_obj = prs.slides[slide_idx]; current_shape_obj = None
                        for s_obj in current_slide_obj.shapes:
                            if hasattr(s_obj, 'shape_id') and s_obj.shape_id == target_shape_id:
                                current_shape_obj = s_obj; break
                        if current_shape_obj is None:
                            f_task_log.write(f"  오류: S{slide_idx+1}-ID{target_shape_id} shape 찾지 못함. 건너뜀.\n")
                            logger.warning(f"Shape ID {target_shape_id} on slide {slide_idx+1} not found. Skipping.")
                            continue
                        log_shape_details = (f"[S{slide_idx+1}] 요소 처리 시작: '{element_name}' (ID: {target_shape_id}), 타입: {item_info['type']}, "
                                             f"위치: L{current_shape_obj.left/914400:.2f}cm, T{current_shape_obj.top/914400:.2f}cm, W{current_shape_obj.width/914400:.2f}cm, H{current_shape_obj.height/914400:.2f}cm\n")
                        logger.debug(log_shape_details.strip()); f_task_log.write(log_shape_details)
                        current_text_for_progress = ""
                        item_type = item_info['type']

                        if item_type == 'text' or item_type == 'table_cell':
                            text_frame = None
                            # ... (text_frame 가져오는 로직 동일) ...
                            if item_type == 'text' and current_shape_obj.has_text_frame:
                                text_frame = current_shape_obj.text_frame
                            elif item_type == 'table_cell' and current_shape_obj.has_table:
                                try:
                                    cell = current_shape_obj.table.cell(item_info['row_idx'], item_info['col_idx'])
                                    text_frame = cell.text_frame
                                except IndexError:
                                    logger.error(f"테이블 셀 접근 오류: {element_name} at S{slide_idx+1} R{item_info['row_idx']}C{item_info['col_idx']}")
                                    f_task_log.write(f"  오류: 테이블 셀 접근 실패 {element_name}\n")
                                    text_frame = None


                            if text_frame and text_frame.text and text_frame.text.strip():
                                original_text = text_frame.text
                                current_text_for_progress = original_text
                                if progress_callback: progress_callback(slide_idx + 1, f"{item_type} ({element_name})", translated_count, total_elements_to_translate, original_text)

                                if should_skip_translation(original_text):
                                    f_task_log.write(f"  [스킵됨 - 번역 불필요] \"{original_text.replace(chr(10), ' / ').strip()[:100]}...\"\n")
                                else:
                                    translated_text = translator.translate_text(original_text, src_lang, tgt_lang, model_name, ollama_service, is_ocr_text=False)
                                    log_original_text_snippet = original_text.replace(chr(10), ' / ').strip()[:100]
                                    log_translated_text_snippet = translated_text.replace(chr(10), ' / ').strip()[:100]
                                    f_task_log.write(f"  [번역 전] \"{log_original_text_snippet}...\" -> [번역 후] \"{log_translated_text_snippet}...\"\n")

                                    if "오류:" not in translated_text:
                                        original_paragraphs_info = []
                                        for para_idx, para in enumerate(text_frame.paragraphs):
                                            original_paragraphs_info.append({
                                                'runs': [{'text': run.text, 'style': self._get_text_style(run)} for run in para.runs],
                                                'alignment': para.alignment, 'level': para.level,
                                                'space_before': para.space_before, 'space_after': para.space_after,
                                                'line_spacing': para.line_spacing
                                            })
                                            if para.text.strip():
                                                f_task_log.write(f"    원본 단락 {para_idx+1} (정렬: {para.alignment}, 수준: {para.level}): '{para.text.strip()[:50].replace(chr(10), ' ')}...'\n")
                                        
                                        # << --- Claude 제안 적용 시작 --- >>
                                        text_frame.clear() # 기존 단락 API 레벨에서 제거 시도

                                        if not translated_text: # 번역 결과가 빈 문자열인 경우 (translator.py에서 ""로 반환됨)
                                            p_new = text_frame.add_paragraph()
                                            run_new = p_new.add_run()
                                            run_new.text = " "
                                            f_task_log.write(f"    번역 결과가 비어 공백 단락 1개 추가.\n")
                                        else:
                                            # text_frame._element는 TextFrame의 txBody XML 요소임
                                            # 여기서 하위 <a:p> 태그들을 직접 제거 시도 (Claude 제안)
                                            # 이 작업은 text_frame.clear()가 완벽하지 않을 경우를 대비한 것일 수 있음.
                                            # 또는 clear() 후에도 기본 단락 구조가 남는 경우를 상정한 것일 수 있음.
                                            # python-pptx 0.6.21 기준, clear()는 모든 <a:p>를 제거해야 함.
                                            # 그럼에도 불구하고 제안된 로직을 테스트.
                                            if hasattr(text_frame, '_element') and text_frame._element is not None:
                                                # text_frame._element는 <p:txBody> 요소
                                                # 그 자식들 중 <a:p>를 제거
                                                txBody = text_frame._element
                                                # lxml.etree._Element의 QName 사용 (a:p 에 해당)
                                                # from pptx.oxml.ns import nsdecls # 'a' 네임스페이스 접두사
                                                # para_tag = '{%s}p' % nsdecls('a') # 네임스페이스를 포함한 태그 이름
                                                
                                                # 더 간단하게는 태그 이름의 로컬 부분만 확인 (네임스페이스 무시)
                                                children_to_remove = [child for child in txBody if child.tag.endswith('}p')]
                                                if children_to_remove:
                                                    f_task_log.write(f"    XML 레벨에서 txBody의 기존 <a:p> {len(children_to_remove)}개 제거 시도.\n")
                                                    for child in children_to_remove:
                                                        txBody.remove(child)

                                            lines = translated_text.splitlines()
                                            if not lines and translated_text: # 단일 줄 텍스트인데 splitlines가 빈 리스트 반환? (이론상 translated_text가 "\n" 같은 경우)
                                                lines = [translated_text] # 이 경우는 translator.py에서 ""로 처리되어 위에서 걸러져야 함
                                            
                                            if not lines: # translated_text가 "" 이거나 "\n" 등 이어서 lines가 비었을 경우
                                                p = text_frame.add_paragraph()
                                                r = p.add_run()
                                                r.text = " "
                                                f_task_log.write(f"    번역된 lines가 비어 공백 단락 1개 추가.\n")
                                            else:
                                                for i, line_content in enumerate(lines):
                                                    p = text_frame.add_paragraph() # 새 단락 추가
                                                    
                                                    para_template = original_paragraphs_info[min(i, len(original_paragraphs_info)-1)] if original_paragraphs_info else None
                                                    if para_template:
                                                        if para_template.get('alignment'): p.alignment = para_template['alignment']
                                                        p.level = para_template.get('level', 0)
                                                        # Claude 제안: 모든 단락의 space_before를 0으로 설정.
                                                        # 이는 원본 스타일을 무시할 수 있으므로, 첫 단락은 원본 유지, 나머지는 0으로 하는 이전 로직 유지 또는 선택 필요.
                                                        # 여기서는 Claude 제안대로 모든 space_before를 0으로 우선 테스트.
                                                        p.space_before = Pt(0) # 명시적으로 0으로 설정 (Claude 제안)
                                                        # p.space_before = para_template.get('space_before', Pt(0)) if i == 0 else Pt(0) # 이전 로직
                                                        p.space_after = para_template.get('space_after', Pt(0))
                                                        p.line_spacing = para_template.get('line_spacing')
                                                    else: # 템플릿 없는 경우 기본값
                                                        p.space_before = Pt(0)
                                                        p.space_after = Pt(0)

                                                    run = p.add_run()
                                                    run.text = line_content if line_content.strip() else " "
                                                    
                                                    log_added_line = line_content.strip()[:50].replace(chr(10), ' ')
                                                    f_task_log.write(f"    번역된 단락 {i+1} 추가: '{log_added_line}...'\n")

                                                    if para_template and para_template.get('runs'):
                                                        self._apply_text_style(run, para_template['runs'][0]['style'])
                                        # << --- Claude 제안 적용 끝 --- >>
                                    else:
                                        f_task_log.write(f"  -> 텍스트 번역 실패: {translated_text}\n")
                            translated_count += 1
                        
                        elif item_type == 'image' and ocr_handler:
                            # (이미지 처리 로직은 이전 답변과 동일하게 유지)
                            try:
                                from PIL import Image as PILImage
                                image_bytes = current_shape_obj.image.blob
                                img_pil_original = PILImage.open(io.BytesIO(image_bytes))
                                img_pil_for_ocr = img_pil_original.convert("RGB")

                                try:
                                    hasher = hashlib.md5(); img_byte_arr = io.BytesIO()
                                    img_pil_for_ocr.save(img_byte_arr, format='PNG')
                                    hasher.update(img_byte_arr.getvalue()); img_hash = hasher.hexdigest()
                                    hash_log_msg = f"    OCR 대상 RGB 이미지 해시 (MD5): {img_hash} (Shape ID: {target_shape_id})\n"
                                    logger.debug(hash_log_msg.strip()); f_task_log.write(hash_log_msg)
                                except Exception as e_hash: logger.warning(f"    이미지 해시 생성 중 오류: {e_hash}")

                                ocr_results = ocr_handler.ocr_image(img_pil_for_ocr)

                                if ocr_results:
                                    f_task_log.write(f"      이미지 내 OCR 텍스트 {len(ocr_results)}개 발견.\n")
                                    edited_image_pil = img_pil_original.copy()
                                    any_ocr_text_translated_and_rendered = False

                                    for ocr_idx, (box, (text, confidence)) in enumerate(ocr_results):
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
                                                logger.error(f"OCR 렌더링 실패 (S{slide_idx+1}-ImgID:{target_shape_id}): {e_render}", exc_info=True)
                                        else:
                                             f_task_log.write(f"            -> 번역 실패 또는 빈 결과로 렌더링 안함.\n")
                                    if stop_event and stop_event.is_set(): break
                                    if any_ocr_text_translated_and_rendered:
                                        image_stream = io.BytesIO()
                                        save_format = img_pil_original.format if img_pil_original.format and img_pil_original.format.upper() in ['JPEG', 'PNG', 'GIF'] else 'PNG'
                                        try:
                                            edited_image_pil.save(image_stream, format=save_format)
                                            image_stream.seek(0)
                                            left, top, width, height = current_shape_obj.left, current_shape_obj.top, current_shape_obj.width, current_shape_obj.height
                                            pic_elem = current_shape_obj.element
                                            pic_parent = pic_elem.getparent()
                                            if pic_parent is not None:
                                                pic_parent.remove(pic_elem)
                                            new_pic_shape = current_slide_obj.shapes.add_picture(image_stream, left, top, width=width, height=height)
                                            f_task_log.write(f"      이미지 '{element_name}' (ID:{target_shape_id}) 교체 완료. 새 Shape ID: {new_pic_shape.shape_id}\n")
                                        except Exception as e_save_replace:
                                            f_task_log.write(f"      오류: 수정된 이미지 저장 또는 교체 중 오류: {e_save_replace}\n")
                                            logger.error(f"Error saving/replacing image (S{slide_idx+1}-ImgID:{target_shape_id}): {e_save_replace}", exc_info=True)
                                    else:
                                        f_task_log.write(f"      이미지 '{element_name}' 변경 없음 (유효/번역/렌더링된 OCR 텍스트 없음).\n")
                                else:
                                    f_task_log.write(f"      이미지 '{element_name}' 내 OCR 텍스트 발견되지 않음.\n")
                            except AttributeError as e_attr:
                                f_task_log.write(f"    경고: '{element_name}' (ID:{target_shape_id}) 그림 shape 처리 중 속성 오류: {e_attr}. 건너뜀.\n")
                                logger.warning(f"Attribute error processing image (S{slide_idx+1}-ImgID:{target_shape_id}): {e_attr}")
                            except OSError as e_os_img_proc:
                                f_task_log.write(f"    이미지 처리 중 Pillow OSError (ID: {target_shape_id}): {e_os_img_proc}. 건너뜀.\n")
                                logger.error(f"Pillow OSError during image processing (S{slide_idx+1}-ImgID:{target_shape_id}): {e_os_img_proc}", exc_info=False)
                            except Exception as e_img_proc_detail:
                                f_task_log.write(f"    이미지 '{element_name}' (ID:{target_shape_id}) 처리 중 예기치 않은 오류: {e_img_proc_detail}. 건너뜀.\n")
                                logger.error(f"Unexpected error processing image (S{slide_idx+1}-ImgID:{target_shape_id}): {e_img_proc_detail}", exc_info=True)
                            translated_count += 1


                        f_task_log.write("\n")
                        if progress_callback and not (stop_event and stop_event.is_set()):
                            progress_callback(slide_idx + 1, item_type, translated_count, total_elements_to_translate, current_text_for_progress)

                    # ... (번역 작업 완료 후 저장 로직 동일) ...
                    f_task_log.write("--- 번역 작업 완료 ---\n")
                    if stop_event and stop_event.is_set():
                        stopped_filename = output_filename.replace(".pptx", "_stopped.pptx")
                        prs.save(stopped_filename)
                        f_task_log.write(f"번역 중지됨. 부분 저장 파일: {stopped_filename}\n")
                        return stopped_filename
                    prs.save(output_filename)
                    f_task_log.write(f"번역된 파일 저장 완료: {output_filename}\n")
                    return output_filename

                except Exception as e_translate_general:
                    err_msg = f"프레젠테이션 번역 중 심각한 오류 발생: {e_translate_general}"
                    logger.error(err_msg, exc_info=True)
                    f_task_log.write(f"오류: {err_msg}\n상세 정보: {traceback.format_exc()}\n")
                    return None
        except IOError as e_file_io:
            logger.error(f"작업 로그 파일 처리 오류 ({task_log_filepath}): {e_file_io}", exc_info=True)
            return None