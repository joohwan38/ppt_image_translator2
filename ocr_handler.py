from PIL import Image, ImageDraw, ImageFont, ImageStat # ImageStat 추가
import numpy as np
import cv2 # cv2는 K-Means 외에는 현재 직접 사용 안함 (K-Means는 제거됨)
import os
# import platform # 현재 직접 사용 안 함
import logging
import io
import textwrap
# import math # 직접 계산

logger = logging.getLogger(__name__)

BASE_DIR_OCR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR_OCR, "fonts")

LANGUAGE_FONT_MAP = {
    'korean': 'NotoSansCJK-Regular.ttc', 'japan': 'NotoSansCJK-Regular.ttc',
    'ch': 'NotoSansCJK-Regular.ttc', 'chinese_cht': 'NotoSansCJK-Regular.ttc',
    'en': 'NotoSansCJK-Regular.ttc', 'th': 'NotoSansThai-VariableFont_wdth,wght.ttf',
    'es': 'NotoSansCJK-Regular.ttc',
    'korean_bold': 'NotoSansCJK-Bold.ttc', 'japan_bold': 'NotoSansCJK-Bold.ttc',
    'en_bold': 'NotoSansCJK-Bold.ttc',
}
DEFAULT_FONT_FILENAME = 'NotoSansCJK-Regular.ttc'
DEFAULT_BOLD_FONT_FILENAME = 'NotoSansCJK-Bold.ttc'


def get_quantized_dominant_color(image_roi, num_colors=8):
    """주어진 PIL Image ROI를 양자화하여 주요 색상을 찾습니다."""
    try:
        if image_roi.width == 0 or image_roi.height == 0:
            return (128, 128, 128)

        # 이미지를 작은 크기로 리사이즈하여 계산 속도 향상 (선택적)
        # thumb_width = min(image_roi.width, 50) # 이전 K-Means에서 사용하던 방식
        # thumb_height = int(image_roi.height * thumb_width / image_roi.width)
        # if thumb_height == 0: thumb_height = 1
        # quantizable_image = image_roi.resize((thumb_width, thumb_height), Image.Resampling.LANCZOS)
        
        quantizable_image = image_roi.convert('RGB') # RGB로 변환하여 alpha 채널 문제 방지

        # Pillow의 quantize 메서드 사용 (MEDIANCUT, MAXCOVERAGE 등 알고리즘 선택 가능)
        # OCTREE 알고리즘이 일반적으로 품질이 좋음
        quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.FASTOCTREE) # Pillow 9.1.0+
        # quantized_image = quantizable_image.quantize(colors=num_colors, method=2) # 구버전 (MEDIANCUT)

        # 양자화된 이미지에서 가장 많이 나타나는 색상 추출
        palette = quantized_image.getpalette() # [(R,G,B), (R,G,B), ...] 형태 (실제로는 R,G,B,R,G,B... 1차원 배열)
        color_counts = quantized_image.getcolors(num_colors) # [(count, palette_index), ...]
        
        if not color_counts: # getcolors가 None 반환 시 (이미지 특성에 따라)
            logger.warning("getcolors() returned None. Falling back to simple average.")
            return get_simple_average_color(image_roi) # 단순 평균으로 대체

        # 가장 빈번한 색상의 인덱스 찾기
        dominant_palette_index = max(color_counts, key=lambda item: item[0])[1]
        
        # 팔레트에서 해당 인덱스의 RGB 값 추출
        # 팔레트는 R,G,B 순서로 1차원 배열이므로, 인덱스 * 3 위치부터 3개 값을 가져옴
        if palette:
            r = palette[dominant_palette_index * 3]
            g = palette[dominant_palette_index * 3 + 1]
            b = palette[dominant_palette_index * 3 + 2]
            dominant_color = (r, g, b)
        else: # 팔레트가 없는 경우 (매우 드묾, quantize는 팔레트 생성)
             logger.warning("Quantized image has no palette. Falling back to simple average.")
             return get_simple_average_color(image_roi)

        return dominant_color
    except Exception as e:
        logger.warning(f"Quantized dominant color detection failed: {e}. Falling back to simple average.")
        # 오류 발생 시 단순 평균색으로 대체
        return get_simple_average_color(image_roi)

def get_simple_average_color(image_roi): # 이전 답변의 함수 유지 (fallback용)
    """주어진 PIL Image ROI의 평균 색상을 반환합니다."""
    try:
        if image_roi.width == 0 or image_roi.height == 0:
            return (128, 128, 128)
        avg_color = tuple(int(c) for c in ImageStat.Stat(image_roi).mean[:3])
        return avg_color
    except Exception as e:
        logger.warning(f"Simple average color detection failed: {e}. Returning default gray.")
        return (128, 128, 128)

def get_contrasting_text_color(bg_color_tuple):
    # (이전 답변과 동일)
    r, g, b = bg_color_tuple
    luma = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0
    if luma > 0.5: return (0, 0, 0)
    else: return (255, 255, 255)


class PaddleOcrHandler:
    # __init__, _pil_to_cv2, _cv2_to_pil, _preprocess_image_for_ocr, has_text_in_image_bytes, ocr_image, _get_font
    # 함수들은 이전 답변과 동일하게 유지됩니다. (변경 없음)
    def __init__(self, lang='korean', debug_enabled=False):
        self.current_lang = lang
        self.debug_mode = debug_enabled
        self.ocr = None
        try:
            from paddleocr import PaddleOCR
            logger.info(f"PaddleOCR 초기화 시도 (lang: {self.current_lang}, debug: {self.debug_mode}, use_angle_cls: False)...")
            self.ocr = PaddleOCR(use_angle_cls=False, lang=self.current_lang, use_gpu=False, show_log=self.debug_mode)
            logger.info(f"PaddleOCR 초기화 완료 (lang: {self.current_lang}).")
        except ImportError:
            logger.critical("PaddleOCR 라이브러리를 찾을 수 없습니다. 'pip install paddleocr paddlepaddle'로 설치해주세요.")
            raise RuntimeError("PaddleOCR 라이브러리가 설치되어 있지 않습니다.")
        except AssertionError as ae:
            logger.error(f"PaddleOCR 초기화 실패 (AssertionError - lang: '{self.current_lang}'): {ae}", exc_info=True)
            raise RuntimeError(f"PaddleOCR 언어 '{self.current_lang}' 모델 로드 실패: {ae}")
        except Exception as e:
            logger.error(f"PaddleOCR 초기화 중 심각한 오류 발생 (lang: '{self.current_lang}'): {e}", exc_info=True)
            raise RuntimeError(f"PaddleOCR 초기화 중 예측하지 못한 오류 (lang: '{self.current_lang}'): {e}")

    def _pil_to_cv2(self, pil_image):
        return cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

    def _cv2_to_pil(self, cv2_image):
        return Image.fromarray(cv2.cvtColor(cv2_image, cv2.COLOR_BGR2RGB))

    def _preprocess_image_for_ocr(self, image_cv):
        logger.debug("OCR을 위한 이미지 전처리 시작...")
        processed_img = image_cv.copy(); h, w = processed_img.shape[:2]; target_width = 1000
        if w > 0 and (w < target_width / 2 or w > target_width * 2.5):
            scale_ratio = target_width / w; new_height = int(h * scale_ratio)
            if new_height > 0 and target_width > 0 :
                 try:
                    processed_img = cv2.resize(processed_img, (target_width, new_height), interpolation=cv2.INTER_LANCZOS4)
                    logger.debug(f"OCR 전처리: 이미지 리사이즈됨 (원본 {w}x{h} -> {target_width}x{new_height})")
                 except cv2.error as e_resize: logger.warning(f"OCR 전처리 중 리사이즈 오류: {e_resize}. 원본 사용.")
        try:
            gray_img = cv2.cvtColor(processed_img, cv2.COLOR_BGR2GRAY)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)); clahe_img = clahe.apply(gray_img)
            _, otsu_thresh_img = cv2.threshold(clahe_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_img_for_ocr = otsu_thresh_img; logger.debug("이미지 전처리 완료 (CLAHE + Otsu).")
        except cv2.error as e_cv:
            logger.error(f"OpenCV 전처리 중 오류: {e_cv}. 원본 회색조로 대체.");
            try: processed_img_for_ocr = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
            except: processed_img_for_ocr = image_cv
        if self.debug_mode:
            try:
                debug_img_path = os.path.join(BASE_DIR_OCR, "preprocessed_debug.png")
                cv2.imwrite(debug_img_path, processed_img_for_ocr)
                logger.debug(f"디버그: 전처리 이미지 저장 ({debug_img_path})")
            except Exception as e_imshow: logger.warning(f"디버그 이미지 저장 오류: {e_imshow}")
        return processed_img_for_ocr

    def has_text_in_image_bytes(self, image_bytes):
        if not self.ocr: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if img_pil.width < 10 or img_pil.height < 10: return False
            img_pil_rgb = img_pil.convert("RGB")
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False
            result = self.ocr_image(img_pil_rgb)
            return bool(result)
        except OSError as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.warning(f"이미지 텍스트 확인 중 Pillow OSError ({format_info}), 처리 건너뜀: {e}", exc_info=False)
            return False
        except Exception as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.error(f"이미지 텍스트 확인 중 예기치 않은 오류 ({format_info}): {e}", exc_info=True)
            return False
        finally:
            if img_pil:
                try: img_pil.close()
                except Exception: pass

    def ocr_image(self, image_pil_rgb):
        if not self.ocr or not image_pil_rgb: return []
        try:
            image_cv_original = self._pil_to_cv2(image_pil_rgb)
            preprocessed_image_cv = self._preprocess_image_for_ocr(image_cv_original)
            ocr_output = self.ocr.ocr(preprocessed_image_cv, cls=False)
            actual_results = []
            if ocr_output and isinstance(ocr_output, list) and len(ocr_output) > 0:
                first_item = ocr_output[0]
                if isinstance(first_item, list) and len(first_item) == 2 and isinstance(first_item[1], tuple):
                    actual_results = ocr_output
                elif isinstance(first_item, list) and len(first_item) > 0 and \
                     isinstance(first_item[0], list) and len(first_item[0]) == 4 and \
                     isinstance(first_item[1], tuple) and len(first_item[1]) == 2:
                    actual_results = [ [item[0][0], item[1]] for item in ocr_output if isinstance(item, list) and len(item) >= 2 ]
                else:
                    logger.warning(f"OCR 결과 형식이 예상과 다릅니다. 첫번째 항목: {first_item}")
                    if len(ocr_output) == 1 and isinstance(ocr_output[0], list):
                         potential_results = ocr_output[0]
                         if potential_results and isinstance(potential_results[0], list) and \
                            len(potential_results[0]) == 2 and isinstance(potential_results[0][1], tuple):
                             actual_results = potential_results
            return actual_results
        except Exception as e:
            logger.error(f"OCR 처리 중 심각한 오류: {e}", exc_info=True)
            return []

    def _get_font(self, font_size, lang_code='en', is_bold=False):
        font_filename = None; font_path = None
        if is_bold:
            bold_font_key = lang_code + '_bold'
            font_filename = LANGUAGE_FONT_MAP.get(bold_font_key)
            if not font_filename:
                font_filename = DEFAULT_BOLD_FONT_FILENAME
                logger.debug(f"'{bold_font_key}'에 대한 볼드 폰트 없음. 기본 볼드 '{font_filename}' 사용.")
        if not font_filename:
            font_filename = LANGUAGE_FONT_MAP.get(lang_code, DEFAULT_FONT_FILENAME)
        if font_filename:
            font_path = os.path.join(FONT_DIR, font_filename)
        if font_path and os.path.exists(font_path):
            try: return ImageFont.truetype(font_path, font_size)
            except IOError as e: logger.warning(f"폰트 로드 실패 ('{font_path}'): {e}. Pillow 기본 폰트로 대체.")
        else: logger.warning(f"폰트 파일('{font_path or font_filename}') 없음. Pillow 기본 폰트 사용.")
        try: return ImageFont.load_default(size=font_size)
        except TypeError: return ImageFont.load_default()

    def render_translated_text_on_image(self, image_pil_original, box, translated_text,
                                        font_code_for_render='en', original_text=""):
        img_to_draw_on = image_pil_original.copy()
        draw = ImageDraw.Draw(img_to_draw_on)

        try:
            box_points = [(int(p[0]), int(p[1])) for p in box]
            x_coords = [p[0] for p in box_points]; y_coords = [p[1] for p in box_points]
            min_x, max_x = min(x_coords), max(x_coords)
            min_y, max_y = min(y_coords), max(y_coords)

            if max_x <= min_x or max_y <= min_y:
                logger.warning(f"렌더링 스킵: 유효하지 않은 텍스트 box 크기 ({min_x},{min_y} - {max_x},{max_y})")
                return image_pil_original
            
            img_w, img_h = img_to_draw_on.size
            crop_min_x = max(0, min_x); crop_min_y = max(0, min_y)
            crop_max_x = min(img_w, max_x); crop_max_y = min(img_h, max_y)

            if crop_max_x <= crop_min_x or crop_max_y <= crop_min_y:
                logger.warning(f"렌더링 스킵: 이미지 경계 내 crop 영역 없음.")
                return image_pil_original
        except Exception as e_box:
            logger.error(f"텍스트 box 좌표 처리 중 오류: {e_box}. 원본 이미지 반환.", exc_info=True)
            return image_pil_original

        # --- 스타일 추정 (색상 양자화 사용) ---
        text_roi_pil = image_pil_original.crop((crop_min_x, crop_min_y, crop_max_x, crop_max_y))
        if text_roi_pil.width == 0 or text_roi_pil.height == 0:
            logger.warning(f"렌더링 스킵: 텍스트 ROI 크기가 0입니다.")
            return image_pil_original
            
        # 색상 양자화를 통해 주요 배경색 추정 (num_colors는 조절 가능)
        estimated_bg_color = get_quantized_dominant_color(text_roi_pil, num_colors=8)
        logger.debug(f"추정된 배경색 (양자화): {estimated_bg_color} for box {box_points}")

        try:
            draw.polygon(box_points, fill=estimated_bg_color)
        except Exception as e_poly_fill:
            logger.warning(f"폴리곤 배경 채우기 실패: {e_poly_fill}. 사각형으로 대체 시도.")
            try: draw.rectangle([min_x, min_y, max_x, max_y], fill=estimated_bg_color)
            except Exception as e_rect_fill: logger.error(f"사각형 배경 채우기 마저 실패: {e_rect_fill}.")

        text_color = get_contrasting_text_color(estimated_bg_color)
        logger.debug(f"결정된 텍스트 색상: {text_color} (배경 대비)")
        # --- 스타일 추정 끝 ---

        # (이하 폰트 크기 조절 및 텍스트 렌더링 로직은 이전 답변과 동일하게 유지)
        bbox_width = max_x - min_x; bbox_height = max_y - min_y
        padding = max(1, int(min(bbox_width, bbox_height) * 0.05))
        render_area_x = min_x + padding; render_area_y = min_y + padding
        render_area_width = bbox_width - 2 * padding; render_area_height = bbox_height - 2 * padding

        if render_area_width <= 0 or render_area_height <= 0:
            logger.warning(f"렌더링 스킵: 패딩 적용 후 렌더링 영역 없음.")
            return img_to_draw_on # 배경은 칠해졌을 수 있음

        target_font_size = int(render_area_height * 0.8); min_font_size = 8
        target_font_size = max(target_font_size, min_font_size)
        is_bold_font = '_bold' in font_code_for_render or 'bold' in font_code_for_render.lower()
        
        font = self._get_font(target_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
        wrapped_text = translated_text; line_spacing_render = 4

        while target_font_size >= min_font_size:
            font = self._get_font(target_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            avg_char_width_sample = "Ag"
            try:
                if hasattr(font, 'getlength'): avg_char_width = font.getlength(avg_char_width_sample) / len(avg_char_width_sample)
                elif hasattr(font, 'getsize'): avg_char_width = font.getsize(avg_char_width_sample)[0] / len(avg_char_width_sample)
                else: avg_char_width = target_font_size / 1.8
            except Exception: avg_char_width = target_font_size / 1.8
            if avg_char_width <= 0: avg_char_width = target_font_size / 2
            chars_per_line = int(render_area_width / avg_char_width) if avg_char_width > 0 else 1
            if chars_per_line <= 0: chars_per_line = 1
            wrapper = textwrap.TextWrapper(width=chars_per_line, break_long_words=True, replace_whitespace=False, drop_whitespace=False)
            wrapped_lines = wrapper.wrap(translated_text)
            if not wrapped_lines: wrapped_lines = [" "]
            total_h = 0
            for line_idx, line_render_text in enumerate(wrapped_lines):
                if hasattr(font, 'getbbox'):
                    line_bbox = font.getbbox(line_render_text)
                    total_h += (line_bbox[3] - line_bbox[1])
                elif hasattr(font, 'getsize'): total_h += font.getsize(line_render_text)[1]
                else: total_h += target_font_size
                if line_idx < len(wrapped_lines) - 1: total_h += line_spacing_render
            if total_h <= render_area_height:
                wrapped_text = "\n".join(wrapped_lines); break
            target_font_size -= 1
        else:
            logger.warning(f"텍스트 '{translated_text[:20]}...' 영역에 맞출 수 없음 (폰트 {min_font_size}).")
            font = self._get_font(min_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            # wrapped_lines는 마지막 시도된 값 사용
            wrapped_text = "\n".join(wrapped_lines)


        try:
            if hasattr(draw, 'multiline_text'):
                 draw.multiline_text((render_area_x, render_area_y), wrapped_text, font=font, fill=text_color, spacing=line_spacing_render, align="left")
            else:
                 draw.text((render_area_x, render_area_y), wrapped_text, font=font, fill=text_color, spacing=line_spacing_render)
            logger.debug(f"텍스트 렌더링 완료: '{wrapped_text[:30].replace(chr(10),'/')}'")
        except TypeError as te:
            logger.warning(f"Pillow draw.text TypeError: {te}. 기본 text 렌더링 시도.")
            try: draw.text((render_area_x, render_area_y), wrapped_text, font=font, fill=text_color)
            except Exception as e_draw_fallback: logger.error(f"텍스트 렌더링 최종 실패: {e_draw_fallback}")
        except Exception as e_draw:
            logger.error(f"텍스트 렌더링 중 오류: {e_draw}", exc_info=True)
        return img_to_draw_on