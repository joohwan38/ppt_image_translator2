from PIL import Image, ImageDraw, ImageFont, ImageStat
import numpy as np
# cv2는 PaddleOCR 전처리 시 필요할 수 있으므로 import 유지
# EasyOCR은 내부적으로 OpenCV를 사용하지만, 핸들러 코드에서 직접 cv2를 많이 쓰지는 않음
import cv2
import os
import logging
import io
import textwrap

logger = logging.getLogger(__name__)

BASE_DIR_OCR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR_OCR, "fonts")

# LANGUAGE_FONT_MAP, DEFAULT_FONT_FILENAME 등은 이전 답변과 동일하게 유지
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


# --- 스타일 추정 함수들 (이전 답변의 양자화 방식 유지) ---
def get_quantized_dominant_color(image_roi, num_colors=8):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128)
        quantizable_image = image_roi.convert('RGB')
        try:
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.FASTOCTREE)
        except AttributeError: 
            logger.debug("FASTOCTREE 양자화 실패, MEDIANCUT으로 대체 시도.")
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.MEDIANCUT) # 구버전 Pillow
        except Exception as e_quant:
             logger.warning(f"색상 양자화 중 오류: {e_quant}. 단순 평균색으로 대체합니다.")
             return get_simple_average_color(image_roi) # get_simple_average_color 정의 필요

        palette = quantized_image.getpalette()
        color_counts = quantized_image.getcolors(num_colors)
        
        if not color_counts:
            logger.warning("getcolors()가 None을 반환했습니다. 단순 평균색으로 대체합니다.")
            return get_simple_average_color(image_roi)

        dominant_palette_index = max(color_counts, key=lambda item: item[0])[1]
        
        if palette:
            r = palette[dominant_palette_index * 3]
            g = palette[dominant_palette_index * 3 + 1]
            b = palette[dominant_palette_index * 3 + 2]
            dominant_color = (r, g, b)
        else:
             logger.warning("양자화된 이미지에 팔레트가 없습니다. 단순 평균색으로 대체합니다.")
             return get_simple_average_color(image_roi)
        return dominant_color
    except Exception as e:
        logger.warning(f"양자화된 주요 색상 감지 실패: {e}. 단순 평균색으로 대체합니다.")
        return get_simple_average_color(image_roi)

def get_simple_average_color(image_roi):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128)
        if image_roi.mode == 'RGBA':
            # 알파 채널 무시하고 RGB 평균 계산
            temp_img = Image.new("RGB", image_roi.size, (255, 255, 255))
            temp_img.paste(image_roi, mask=image_roi.split()[3]) # Paste using alpha channel as mask
            avg_color = tuple(int(c) for c in ImageStat.Stat(temp_img).mean[:3])

        else:
            avg_color = tuple(int(c) for c in ImageStat.Stat(image_roi.convert('RGB')).mean[:3])
        return avg_color
    except Exception as e:
        logger.warning(f"단순 평균색 감지 실패: {e}. 기본 회색 반환.")
        return (128, 128, 128)

def get_contrasting_text_color(bg_color_tuple):
    r, g, b = bg_color_tuple
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    threshold = 150 
    if brightness >= threshold: return (0, 0, 0)
    else: return (255, 255, 255)
# --- 스타일 추정 함수들 끝 ---


class BaseOcrHandler:
    def __init__(self, lang_codes, debug_enabled=False, use_gpu=False):
        self.current_lang_codes = lang_codes 
        self.debug_mode = debug_enabled
        self.use_gpu = use_gpu
        self.ocr_engine = None
        self._initialize_engine()

    def _initialize_engine(self):
        raise NotImplementedError

    def ocr_image(self, image_pil_rgb):
        raise NotImplementedError

    def has_text_in_image_bytes(self, image_bytes):
        if not self.ocr_engine: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if img_pil.width < 10 or img_pil.height < 10: return False
            img_pil_rgb = img_pil.convert("RGB") # Convert to RGB for consistency
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False
            results = self.ocr_image(img_pil_rgb)
            return bool(results)
        except OSError as e: # Pillow에서 이미지 열기/변환 실패
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.warning(f"이미지 텍스트 확인 중 Pillow OSError ({format_info}), 처리 건너뜀: {e}", exc_info=False)
            return False
        except Exception as e: # 그 외 모든 예외
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.error(f"이미지 텍스트 확인 중 예기치 않은 오류 ({format_info}): {e}", exc_info=True)
            return False
        finally:
            if img_pil:
                try: img_pil.close()
                except Exception: pass
    
    def _get_font(self, font_size, lang_code='en', is_bold=False):
        font_filename = None; font_path = None
        if is_bold:
            bold_font_key = lang_code + '_bold'
            font_filename = LANGUAGE_FONT_MAP.get(bold_font_key)
            if not font_filename:
                font_filename = DEFAULT_BOLD_FONT_FILENAME
        if not font_filename: # 일반 폰트
            font_filename = LANGUAGE_FONT_MAP.get(lang_code, DEFAULT_FONT_FILENAME)
        if font_filename: font_path = os.path.join(FONT_DIR, font_filename)
        if font_path and os.path.exists(font_path):
            try: return ImageFont.truetype(font_path, font_size)
            except IOError as e: logger.warning(f"폰트 로드 실패 ('{font_path}'): {e}. Pillow 기본 폰트로 대체.")
        else: logger.warning(f"폰트 파일('{font_path or font_filename}') 없음. Pillow 기본 폰트 사용.")
        # Pillow 기본 폰트 반환 (size 인자 시도)
        try: return ImageFont.load_default(size=font_size)
        except TypeError: return ImageFont.load_default()


    def render_translated_text_on_image(self, image_pil_original, box, translated_text,
                                        font_code_for_render='en', original_text=""):
        # (이전 답변의 양자화 기반 스타일 추정 로직 유지)
        img_to_draw_on = image_pil_original.copy()
        draw = ImageDraw.Draw(img_to_draw_on)
        try:
            box_points = [(int(p[0]), int(p[1])) for p in box]
            x_coords = [p[0] for p in box_points]; y_coords = [p[1] for p in box_points]
            min_x, max_x = min(x_coords), max(x_coords)
            min_y, max_y = min(y_coords), max(y_coords)
            if max_x <= min_x or max_y <= min_y: return image_pil_original
            img_w, img_h = img_to_draw_on.size
            crop_min_x = max(0, min_x); crop_min_y = max(0, min_y)
            crop_max_x = min(img_w, max_x); crop_max_y = min(img_h, max_y)
            if crop_max_x <= crop_min_x or crop_max_y <= crop_min_y: return image_pil_original
        except Exception as e_box:
            logger.error(f"텍스트 box 좌표 처리 중 오류: {e_box}. 원본 이미지 반환.", exc_info=True)
            return image_pil_original

        text_roi_pil = image_pil_original.crop((crop_min_x, crop_min_y, crop_max_x, crop_max_y))
        if text_roi_pil.width == 0 or text_roi_pil.height == 0: return image_pil_original
        
        estimated_bg_color = get_quantized_dominant_color(text_roi_pil, num_colors=8)
        logger.debug(f"추정된 배경색 (양자화): {estimated_bg_color} for box {box_points}")
        try: draw.polygon(box_points, fill=estimated_bg_color)
        except Exception as e_poly_fill:
            logger.warning(f"폴리곤 배경 채우기 실패: {e_poly_fill}. 사각형으로 대체 시도.")
            try: draw.rectangle([min_x, min_y, max_x, max_y], fill=estimated_bg_color)
            except Exception as e_rect_fill: logger.error(f"사각형 배경 채우기 마저 실패: {e_rect_fill}.")
        
        text_color = get_contrasting_text_color(estimated_bg_color)
        
        bbox_width = max_x - min_x; bbox_height = max_y - min_y
        padding = max(1, int(min(bbox_width, bbox_height) * 0.05)) 
        render_area_x = min_x + padding; render_area_y = min_y + padding
        render_area_width = bbox_width - 2 * padding; render_area_height = bbox_height - 2 * padding

        if render_area_width <= 0 or render_area_height <= 0: return img_to_draw_on
        
        target_font_size = int(render_area_height * 0.8); min_font_size = 8
        target_font_size = max(target_font_size, min_font_size)
        is_bold_font = '_bold' in font_code_for_render or 'bold' in font_code_for_render.lower()
        
        font = self._get_font(target_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
        wrapped_text = translated_text; line_spacing_render = 4

        # 폰트 크기 조절 루프 (이전과 동일)
        while target_font_size >= min_font_size:
            font = self._get_font(target_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            avg_char_width_sample = "Ag"; avg_char_width = target_font_size / 1.8 
            try:
                if hasattr(font, 'getlength'): avg_char_width = font.getlength(avg_char_width_sample) / len(avg_char_width_sample)
                elif hasattr(font, 'getsize'): avg_char_width = font.getsize(avg_char_width_sample)[0] / len(avg_char_width_sample)
            except Exception: pass 
            if avg_char_width <= 0: avg_char_width = target_font_size / 2
            chars_per_line = int(render_area_width / avg_char_width) if avg_char_width > 0 else 1
            if chars_per_line <= 0: chars_per_line = 1
            wrapper = textwrap.TextWrapper(width=chars_per_line, break_long_words=True, replace_whitespace=False, drop_whitespace=False)
            wrapped_lines = wrapper.wrap(translated_text);
            if not wrapped_lines: wrapped_lines = [" "]
            total_h = 0
            for line_idx, line_render_text in enumerate(wrapped_lines):
                h_line = target_font_size 
                try:
                    if hasattr(font, 'getbbox'): line_bbox = font.getbbox(line_render_text); h_line = line_bbox[3] - line_bbox[1]
                    elif hasattr(font, 'getsize'): h_line = font.getsize(line_render_text)[1]
                except: pass
                total_h += h_line
                if line_idx < len(wrapped_lines) - 1: total_h += line_spacing_render
            if total_h <= render_area_height: wrapped_text = "\n".join(wrapped_lines); break
            target_font_size -= 1
        else: # 루프가 다 돌아도 못 맞춘 경우
            logger.warning(f"텍스트 '{translated_text[:20]}...' 영역에 맞출 수 없음 (폰트 {min_font_size}).")
            font = self._get_font(min_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            # wrapped_lines는 마지막 시도된 값 사용
            wrapped_text = "\n".join(wrapped_lines) 
        
        # 텍스트 렌더링
        try:
            if hasattr(draw, 'multiline_text'):
                 draw.multiline_text((render_area_x, render_area_y), wrapped_text, font=font, fill=text_color, spacing=line_spacing_render, align="left")
            else:
                 draw.text((render_area_x, render_area_y), wrapped_text, font=font, fill=text_color, spacing=line_spacing_render)
        except TypeError: # Pillow 구버전 호환 (spacing 인자 없을 수 있음)
            try: draw.text((render_area_x, render_area_y), wrapped_text, font=font, fill=text_color)
            except Exception as e_draw_fallback: logger.error(f"텍스트 렌더링 최종 실패: {e_draw_fallback}")
        except Exception as e_draw: logger.error(f"텍스트 렌더링 중 오류: {e_draw}", exc_info=True)
        return img_to_draw_on


class PaddleOcrHandler(BaseOcrHandler):
    def __init__(self, lang_code='korean', debug_enabled=False, use_gpu=False):
        super().__init__(lang_codes=lang_code, debug_enabled=debug_enabled, use_gpu=use_gpu)

    def _initialize_engine(self):
        try:
            from paddleocr import PaddleOCR # 실제 엔진 초기화 시점에 import
            logger.info(f"PaddleOCR 초기화 시도 (lang: {self.current_lang_codes}, use_gpu: {self.use_gpu}, debug: {self.debug_mode})...")
            self.ocr_engine = PaddleOCR(use_angle_cls=False, lang=self.current_lang_codes, use_gpu=self.use_gpu, show_log=self.debug_mode)
            logger.info(f"PaddleOCR 초기화 완료 (lang: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("PaddleOCR 라이브러리를 찾을 수 없습니다. 'pip install paddleocr paddlepaddle'로 설치해주세요.")
            raise RuntimeError("PaddleOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e: # 모델 다운로드 실패 등 포함
            logger.error(f"PaddleOCR 초기화 중 오류 (lang: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"PaddleOCR 초기화 실패 (lang: {self.current_lang_codes}): {e}")

    def _preprocess_image_for_ocr(self, image_pil_rgb):
        # (이전 답변과 동일)
        image_cv = cv2.cvtColor(np.array(image_pil_rgb), cv2.COLOR_RGB2BGR)
        logger.debug("PaddleOCR용 이미지 전처리 시작...")
        processed_img = image_cv.copy(); h, w = processed_img.shape[:2]; target_width = 1000
        if w > 0 and (w < target_width / 2 or w > target_width * 2.5):
            scale_ratio = target_width / w; new_height = int(h * scale_ratio)
            if new_height > 0 and target_width > 0 :
                 try: processed_img = cv2.resize(processed_img, (target_width, new_height), interpolation=cv2.INTER_LANCZOS4)
                 except cv2.error as e_resize: logger.warning(f"PaddleOCR 전처리 리사이즈 오류: {e_resize}.")
        try:
            gray_img = cv2.cvtColor(processed_img, cv2.COLOR_BGR2GRAY)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)); clahe_img = clahe.apply(gray_img)
            _, otsu_thresh_img = cv2.threshold(clahe_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_img_for_ocr = otsu_thresh_img
        except cv2.error as e_cv:
            logger.error(f"PaddleOCR용 OpenCV 전처리 오류: {e_cv}. 원본 회색조 사용.");
            try: processed_img_for_ocr = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
            except: processed_img_for_ocr = image_cv 
        if self.debug_mode: # 디버그 이미지 저장 (필요시)
            # cv2.imwrite(os.path.join(BASE_DIR_OCR, "paddle_preprocessed_debug.png"), processed_img_for_ocr)
            pass
        return processed_img_for_ocr

    def ocr_image(self, image_pil_rgb):
        # (이전 답변과 동일 - PaddleOCR 결과 파싱)
        if not self.ocr_engine: return []
        try:
            preprocessed_cv_img = self._preprocess_image_for_ocr(image_pil_rgb)
            ocr_output = self.ocr_engine.ocr(preprocessed_cv_img, cls=False)
            final_parsed_results = []
            if ocr_output and isinstance(ocr_output, list) and len(ocr_output) > 0:
                results_list = ocr_output
                if isinstance(ocr_output[0], list) and \
                   (len(ocr_output[0]) == 0 or (len(ocr_output[0]) > 0 and isinstance(ocr_output[0][0], list))):
                    results_list = ocr_output[0]
                for item in results_list:
                    if isinstance(item, list) and len(item) == 2:
                        box_data, text_conf_tuple = item
                        if isinstance(box_data, list) and len(box_data) == 4 and \
                           all(isinstance(point, list) and len(point) == 2 for point in box_data) and \
                           isinstance(text_conf_tuple, tuple) and len(text_conf_tuple) == 2:
                            final_parsed_results.append([box_data, text_conf_tuple])
                        else: logger.warning(f"PaddleOCR 결과 항목 형식이 다릅니다 (내부): {item}")
                    else: logger.warning(f"PaddleOCR 결과 항목이 리스트가 아니거나 길이가 2가 아닙니다 (외부): {item}")
            return final_parsed_results
        except Exception as e:
            logger.error(f"PaddleOCR ocr_image 중 오류: {e}", exc_info=True)
            return []


class EasyOcrHandler(BaseOcrHandler): # 여기에 클래스 정의가 있어야 합니다.
    def __init__(self, lang_codes_list=['en'], debug_enabled=False, use_gpu=False):
        super().__init__(lang_codes=lang_codes_list, debug_enabled=debug_enabled, use_gpu=use_gpu)

    def _initialize_engine(self):
        try:
            import easyocr # 실제 엔진 초기화 시점에 import
            logger.info(f"EasyOCR 초기화 시도 (langs: {self.current_lang_codes}, gpu: {self.use_gpu})...")
            self.ocr_engine = easyocr.Reader(self.current_lang_codes, gpu=self.use_gpu, verbose=self.debug_mode)
            logger.info(f"EasyOCR 초기화 완료 (langs: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("EasyOCR 라이브러리를 찾을 수 없습니다. 'pip install easyocr'로 설치해주세요.")
            raise RuntimeError("EasyOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e: # 모델 다운로드 실패 등 포함
            logger.error(f"EasyOCR 초기화 중 오류 (langs: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"EasyOCR 초기화 실패 (langs: {self.current_lang_codes}): {e}")

    def ocr_image(self, image_pil_rgb):
        if not self.ocr_engine: return []
        try:
            # EasyOCR은 RGB NumPy 배열을 기대함
            image_np = np.array(image_pil_rgb.convert('RGB'))
            
            ocr_output = self.ocr_engine.readtext(image_np, detail=1, paragraph=False)
            
            formatted_results = []
            for (bbox, text, confidence) in ocr_output:
                # EasyOCR bbox: [[x1,y1],[x2,y1],[x2,y2],[x1,y2]] (좌상단부터 시계방향 순서)
                # 또는 더 일반적인 폴리곤일 수 있음.
                # 결과 형식을 [[box_points, (text, confidence)], ...]로 통일
                if isinstance(bbox, list) and len(bbox) == 4 and \
                   all(isinstance(p, list) and len(p) == 2 for p in bbox):
                    box_points = [[int(coord[0]), int(coord[1])] for coord in bbox]
                    formatted_results.append([box_points, (text, float(confidence))])
                else:
                     logger.warning(f"EasyOCR 결과의 bbox 형식이 예상과 다릅니다: {bbox}")
            return formatted_results
        except Exception as e:
            logger.error(f"EasyOCR ocr_image 중 오류: {e}", exc_info=True)
            return []