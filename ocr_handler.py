# ocr_handler.py
from PIL import Image, ImageDraw, ImageFont, ImageStat, __version__ as PILLOW_VERSION
import numpy as np
import cv2 # paddleocr, easyocr 둘 다 내부적으로 cv2 사용 가능성
import os
import logging
import io
import textwrap
import math
from typing import List, Any, Optional # Optional 추가
import functools

# 설정 파일 import
import config
from interfaces import AbsOcrHandler, AbsOcrHandlerFactory # 인터페이스 import
import utils # OCR 설치 확인용

logger = logging.getLogger(__name__)

# BASE_DIR_OCR은 ocr_handler.py 파일의 위치를 기준으로 하는 것이 적절
BASE_DIR_OCR = os.path.dirname(os.path.abspath(__file__))
# FONT_DIR은 config.py에서 가져온 전역 설정을 사용
FONT_DIR = config.FONTS_DIR

logger.info(f"OCR Handler: Using Pillow version {PILLOW_VERSION}")
PILLOW_VERSION_TUPLE = tuple(map(int, PILLOW_VERSION.split('.')))

def get_quantized_dominant_color(image_roi, num_colors=5):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128)
        quantizable_image = image_roi.convert('RGB')
        try:
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.FASTOCTREE)
        except AttributeError:
            logger.debug("FASTOCTREE 양자화 실패, MEDIANCUT으로 대체 시도 (Pillow < 9.1.0).")
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.MEDIANCUT)
        except Exception as e_quant:
             logger.warning(f"색상 양자화 중 오류: {e_quant}. 단순 평균색으로 대체합니다.")
             return get_simple_average_color(image_roi)

        palette = quantized_image.getpalette()
        color_counts = quantized_image.getcolors(num_colors * 2)

        if not color_counts:
            logger.warning("getcolors()가 None을 반환 (양자화 실패 가능성). 단순 평균색으로 대체.")
            return get_simple_average_color(image_roi)

        dominant_palette_index = max(color_counts, key=lambda item: item[0])[1]

        if palette:
            r = palette[dominant_palette_index * 3]
            g = palette[dominant_palette_index * 3 + 1]
            b = palette[dominant_palette_index * 3 + 2]
            dominant_color = (r, g, b)
        else:
             logger.warning("양자화된 이미지에 팔레트가 없습니다. 단순 평균색으로 대체.")
             return get_simple_average_color(image_roi)
        return dominant_color
    except Exception as e:
        logger.warning(f"양자화된 주요 색상 감지 실패: {e}. 단순 평균색으로 대체.", exc_info=True)
        return get_simple_average_color(image_roi)

def get_simple_average_color(image_roi):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128)
        if image_roi.mode == 'RGBA':
            temp_img = Image.new("RGB", image_roi.size, (255, 255, 255))
            temp_img.paste(image_roi, mask=image_roi.split()[3])
            avg_color_tuple = ImageStat.Stat(temp_img).mean
        else:
            avg_color_tuple = ImageStat.Stat(image_roi.convert('RGB')).mean

        return tuple(int(c) for c in avg_color_tuple[:3])
    except Exception as e:
        logger.warning(f"단순 평균색 감지 실패: {e}. 기본 회색 반환.", exc_info=True)
        return (128, 128, 128)

def get_contrasting_text_color(bg_color_tuple):
    r, g, b = bg_color_tuple
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    threshold = 128
    if brightness >= threshold:
        return (0, 0, 0)
    else:
        return (255, 255, 255)

class BaseOcrHandlerImpl(AbsOcrHandler): # 실제 구현을 위한 기본 클래스 (이름 변경)
    def __init__(self, lang_codes_param, debug_enabled=False, use_gpu_param=False): # 파라미터 이름 명확화
        self._current_lang_codes = lang_codes_param
        self._debug_mode = debug_enabled
        self._use_gpu = use_gpu_param
        self._ocr_engine = None # AbsOcrHandler의 ocr_engine 속성을 위해 _ocr_engine 사용
        self._initialize_engine()

    @property
    def ocr_engine(self) -> Any:
        return self._ocr_engine

    @property
    def use_gpu(self) -> bool:
        return self._use_gpu
    
    @property
    def current_lang_codes(self) -> Any:
        return self._current_lang_codes


    def _initialize_engine(self):
        raise NotImplementedError(" 각 OCR 핸들러는 이 메서드를 구현해야 합니다.")

    def ocr_image(self, image_pil_rgb: Image.Image) -> List[Any] : # PIL.Image.Image로 명시
        raise NotImplementedError("각 OCR 핸들러는 이 메서드를 구현해야 합니다.")

    def has_text_in_image_bytes(self, image_bytes: bytes) -> bool:
        if not self.ocr_engine: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if img_pil.width < 5 or img_pil.height < 5: return False
            img_pil_rgb = img_pil.convert("RGB")
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False

            results = self.ocr_image(img_pil_rgb)
            return bool(results and any(res[1][0].strip() for res in results if len(res) > 1 and len(res[1]) > 0))

        except OSError as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.warning(f"이미지 텍스트 확인 중 Pillow OSError ({format_info}), 건너뜀: {e}", exc_info=False)
            return False
        except Exception as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.error(f"이미지 텍스트 확인 중 예기치 않은 오류 ({format_info}): {e}", exc_info=True)
            return False
        finally:
            if img_pil:
                try: img_pil.close()
                except Exception: pass

    @functools.lru_cache(maxsize=128)
    def _get_font(self, font_size: int, lang_code: str = 'en', is_bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont :
        font_size = max(1, int(font_size))
        font_filename = None
        font_path = None

        language_font_map = config.OCR_LANGUAGE_FONT_MAP
        default_font_filename = config.OCR_DEFAULT_FONT_FILENAME
        default_bold_font_filename = config.OCR_DEFAULT_BOLD_FONT_FILENAME

        if is_bold:
            bold_font_key = lang_code + '_bold'
            font_filename = language_font_map.get(bold_font_key)
            if not font_filename:
                font_filename = default_bold_font_filename

        if not font_filename:
            font_filename = language_font_map.get(lang_code, default_font_filename)

        if not font_filename:
            font_filename = default_font_filename if not is_bold else default_bold_font_filename

        if font_filename:
            font_path = os.path.join(FONT_DIR, font_filename)

        if font_path and os.path.exists(font_path):
            try:
                return ImageFont.truetype(font_path, int(font_size))
            except IOError as e:
                logger.warning(f"트루타입 폰트 로드 실패 ('{font_path}', size:{font_size}): {e}. Pillow 기본 폰트로 대체.")
            except Exception as e_font:
                logger.error(f"폰트 로드 중 예기치 않은 오류 ('{font_path}', size:{font_size}): {e_font}. Pillow 기본 폰트로 대체.", exc_info=True)
        else:
            logger.warning(f"폰트 파일 없음: '{font_path or font_filename}' (요청 코드: {lang_code}, bold: {is_bold}). Pillow 기본 폰트 사용.")

        try:
            if PILLOW_VERSION_TUPLE >= (10, 0, 0):
                 return ImageFont.load_default()
            elif PILLOW_VERSION_TUPLE >= (9, 0, 0):
                 return ImageFont.load_default(size=int(font_size))
            else:
                 return ImageFont.load_default()
        except TypeError:
            try:
                return ImageFont.load_default()
            except Exception as e_default_font_fallback:
                logger.critical(f"Pillow 기본 폰트 로드조차 실패 (size={font_size}): {e_default_font_fallback}. 글꼴 렌더링 불가.", exc_info=True)
                raise RuntimeError(f"기본 폰트 로드 실패: {e_default_font_fallback}")
        except Exception as e_default_font:
            logger.critical(f"Pillow 기본 폰트 로드 실패 (size={font_size}): {e_default_font}. 글꼴 렌더링 불가.", exc_info=True)
            raise RuntimeError(f"기본 폰트 로드 실패: {e_default_font}")

    def _calculate_text_dimensions(self, draw: ImageDraw.ImageDraw, text: str, font_size: int,
                                   render_area_width: int, lang_code: str, is_bold: bool, line_spacing: int) -> tuple[int, int, List[str]]:
        if font_size < 1: font_size = 1
        current_font = self._get_font(font_size, lang_code=lang_code, is_bold=is_bold)
        estimated_chars_per_line = 1
        if render_area_width > 0:
            try:
                char_w_metric = 0
                if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'textlength'):
                    char_w_metric = draw.textlength("W", font=current_font)
                    if char_w_metric <= 0: char_w_metric = draw.textlength("가", font=current_font)
                elif hasattr(current_font, 'getlength'):
                    char_w_metric = current_font.getlength("W")
                    if char_w_metric <= 0: char_w_metric = current_font.getlength("가")
                elif hasattr(current_font, 'getsize'):
                     char_w_metric, _ = current_font.getsize("W")
                     if char_w_metric <= 0 : char_w_metric, _ = current_font.getsize("가")

                if char_w_metric > 0:
                    estimated_chars_per_line = max(1, int(render_area_width / char_w_metric))
                else:
                    estimated_chars_per_line = max(1, int(render_area_width / (font_size * 0.5 + 1)))
            except Exception as e_char_width:
                logger.debug(f"문자 너비 계산 중 예외: {e_char_width}. 근사치 사용.")
                estimated_chars_per_line = max(1, int(render_area_width / (font_size * 0.6)))

        wrapper = textwrap.TextWrapper(width=estimated_chars_per_line, break_long_words=True,
                                       replace_whitespace=False, drop_whitespace=False,
                                       break_on_hyphens=True)
        wrapped_lines = wrapper.wrap(text)
        if not wrapped_lines: wrapped_lines = [" "]

        rendered_text_height = 0
        rendered_text_width = 0

        if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'multiline_textbbox'):
            try:
                text_bbox_args = {'xy': (0,0), 'text': "\n".join(wrapped_lines), 'font': current_font, 'spacing': line_spacing}
                if PILLOW_VERSION_TUPLE >= (9, 3, 0):
                    text_bbox_args['anchor'] = "lt"
                
                text_bbox = draw.multiline_textbbox(**text_bbox_args)
                rendered_text_width = text_bbox[2] - text_bbox[0]
                rendered_text_height = text_bbox[3] - text_bbox[1]
            except Exception as e_mtbox:
                logger.debug(f"multiline_textbbox 사용 중 예외: {e_mtbox}. 수동 계산으로 대체.")
                rendered_text_width, rendered_text_height = self._manual_calculate_multiline_dimensions(draw, wrapped_lines, current_font, line_spacing, font_size)
        else:
            rendered_text_width, rendered_text_height = self._manual_calculate_multiline_dimensions(draw, wrapped_lines, current_font, line_spacing, font_size)

        return int(rendered_text_width), int(rendered_text_height), wrapped_lines

    def _manual_calculate_multiline_dimensions(self, draw: ImageDraw.ImageDraw, wrapped_lines: List[str],
                                             font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
                                             line_spacing: int, fallback_font_size: int) -> tuple[int, int]:
        total_h = 0
        max_w = 0
        for i, line_txt in enumerate(wrapped_lines):
            line_w, line_h = 0, 0
            try:
                if hasattr(draw, 'textbbox'):
                    bbox_args = {'xy': (0,0), 'text': line_txt, 'font': font}
                    if PILLOW_VERSION_TUPLE >= (9, 3, 0):
                        bbox_args['anchor'] = "lt"
                    line_bbox = draw.textbbox(**bbox_args)
                    line_w = line_bbox[2] - line_bbox[0]
                    line_h = line_bbox[3] - line_bbox[1]
                elif hasattr(font, 'getsize'):
                    line_w, line_h = font.getsize(line_txt)
                elif hasattr(font, 'getbbox'):
                    bbox = font.getbbox(line_txt)
                    line_w = bbox[2] - bbox[0]
                    line_h = bbox[3] - bbox[1]
                else:
                    line_w = len(line_txt) * fallback_font_size * 0.6
                    line_h = fallback_font_size
            except Exception as e_line_calc:
                logger.debug(f"개별 라인 크기 계산 중 예외: {e_line_calc}. 근사치 사용.")
                line_w = len(line_txt) * fallback_font_size * 0.6
                line_h = fallback_font_size

            total_h += line_h
            if line_w > max_w:
                max_w = line_w
            if i < len(wrapped_lines) - 1:
                total_h += line_spacing
        return int(max_w), int(total_h)

    def render_translated_text_on_image(self, image_pil_original: Image.Image, box: List[List[int]], translated_text: str, # PIL.Image.Image로 명시
                                        font_code_for_render='en', original_text="", ocr_angle=None) -> Image.Image: # PIL.Image.Image로 명시
        img_to_draw_on = image_pil_original.copy()
        draw = ImageDraw.Draw(img_to_draw_on)

        try:
            x_coords = [p[0] for p in box]
            y_coords = [p[1] for p in box]
            min_x, max_x = min(x_coords), max(x_coords)
            min_y, max_y = min(y_coords), max(y_coords)

            if max_x <= min_x or max_y <= min_y:
                logger.warning(f"렌더링 스킵: 유효하지 않은 바운딩 박스 {box} for '{translated_text[:20]}...'")
                return image_pil_original

            img_w, img_h = img_to_draw_on.size
            render_box_x1 = max(0, int(min_x))
            render_box_y1 = max(0, int(min_y))
            render_box_x2 = min(img_w, int(max_x))
            render_box_y2 = min(img_h, int(max_y))

            if render_box_x2 <= render_box_x1 or render_box_y2 <= render_box_y1:
                logger.warning(f"렌더링 스킵: 크기가 0인 렌더 박스 for '{translated_text[:20]}...'")
                return image_pil_original

            bbox_width_orig = max_x - min_x
            bbox_height_orig = max_y - min_y
            bbox_width_render = render_box_x2 - render_box_x1
            bbox_height_render = render_box_y2 - render_box_y1

        except Exception as e_box_calc:
            logger.error(f"렌더링 바운딩 박스 계산 오류: {e_box_calc}. Box: {box}. 원본 이미지 반환.", exc_info=True)
            return image_pil_original

        try:
            text_roi_pil = image_pil_original.crop((render_box_x1, render_box_y1, render_box_x2, render_box_y2))
            estimated_bg_color = get_quantized_dominant_color(text_roi_pil) if text_roi_pil.width > 0 and text_roi_pil.height > 0 else (200,200,200)
        except Exception as e_bg:
            logger.warning(f"배경색 추정 실패 ({e_bg}), 기본 회색 사용.", exc_info=True)
            estimated_bg_color = (200, 200, 200)

        draw.rectangle([render_box_x1, render_box_y1, render_box_x2, render_box_y2], fill=estimated_bg_color)
        text_color = get_contrasting_text_color(estimated_bg_color)

        padding_x = max(1, int(bbox_width_render * 0.03))
        padding_y = max(1, int(bbox_height_render * 0.03))

        render_area_x_start = render_box_x1 + padding_x
        render_area_y_start = render_box_y1 + padding_y
        render_area_width = bbox_width_render - 2 * padding_x
        render_area_height = bbox_height_render - 2 * padding_y

        if render_area_width <= 1 or render_area_height <= 1:
            return img_to_draw_on

        font_size_correction_factor = 1.0
        text_angle_deg = 0.0
        if ocr_angle is not None and isinstance(ocr_angle, (int, float)):
            text_angle_deg = abs(ocr_angle)
            if 5 < text_angle_deg < 85 or 95 < text_angle_deg < 175:
                font_size_correction_factor = max(0.6, 1.0 - (text_angle_deg / 90.0) * 0.3)
        elif bbox_width_orig > 0 and bbox_height_orig > 0 :
            aspect_ratio_orig = bbox_width_orig / bbox_height_orig
            if aspect_ratio_orig > 2.0 or aspect_ratio_orig < 0.5:
                font_size_correction_factor = 0.80

        initial_target_font_size = int(min(render_area_height * 0.9,
                                    render_area_width * 0.9 / (len(translated_text.splitlines()[0] if translated_text else "A")*0.5 +1)
                                   ) * font_size_correction_factor)
        initial_target_font_size = max(initial_target_font_size, 1)

        min_font_size = 5
        if initial_target_font_size < min_font_size: initial_target_font_size = min_font_size

        is_bold_font = '_bold' in font_code_for_render or 'bold' in font_code_for_render.lower()
        best_fit_size = min_font_size
        best_wrapped_lines: List[str] = []
        best_text_width = 0
        best_text_height = 0
        low = min_font_size
        high = initial_target_font_size
        max_iterations = int(math.log2(high - low + 1)) + 5 if high > low else 5
        current_iteration = 0

        while low <= high and current_iteration < max_iterations:
            current_iteration +=1
            mid_font_size = low + (high - low) // 2
            if mid_font_size < min_font_size : mid_font_size = min_font_size
            if mid_font_size == 0 : break
            current_line_spacing = int(mid_font_size * 0.2)
            w, h, wrapped = self._calculate_text_dimensions(draw, translated_text, mid_font_size,
                                                            render_area_width, font_code_for_render,
                                                            is_bold_font, current_line_spacing)
            if w <= render_area_width and h <= render_area_height:
                best_fit_size = mid_font_size
                best_wrapped_lines = wrapped
                best_text_width = w
                best_text_height = h
                low = mid_font_size + 1
            else:
                high = mid_font_size - 1

        if not best_wrapped_lines:
            final_line_spacing = int(min_font_size * 0.2)
            best_text_width, best_text_height, best_wrapped_lines = self._calculate_text_dimensions(
                draw, translated_text, min_font_size, render_area_width, font_code_for_render, is_bold_font, final_line_spacing
            )
            best_fit_size = min_font_size

        final_font_size = best_fit_size
        final_font = self._get_font(final_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
        final_line_spacing_render = int(final_font_size * 0.2)
        text_x_start = render_area_x_start + (render_area_width - best_text_width) / 2
        text_y_start = render_area_y_start + (render_area_height - best_text_height) / 2
        text_x_start = max(render_area_x_start, text_x_start)
        text_y_start = max(render_area_y_start, text_y_start)

        try:
            if PILLOW_VERSION_TUPLE >= (9,0,0) and hasattr(draw, 'multiline_text'):
                multiline_args = {
                    'xy': (text_x_start, text_y_start),
                    'text': "\n".join(best_wrapped_lines),
                    'font': final_font,
                    'fill': text_color,
                    'spacing': final_line_spacing_render,
                    'align': "left"
                }
                if PILLOW_VERSION_TUPLE >= (9,3,0):
                    multiline_args['anchor'] = "la"
                draw.multiline_text(**multiline_args)
            else:
                 current_y = text_y_start
                 for line_idx, line_txt in enumerate(best_wrapped_lines):
                     line_height_val = final_font_size
                     if hasattr(draw, 'textbbox'):
                         bbox_args = {'xy': (0,0), 'text': line_txt, 'font': final_font}
                         if PILLOW_VERSION_TUPLE >= (9,3,0): bbox_args['anchor'] = "lt"
                         line_bbox = draw.textbbox(**bbox_args)
                         line_height_val = line_bbox[3] - line_bbox[1] if line_bbox else final_font_size
                     elif hasattr(final_font, 'getsize'):
                        _, line_height_val = final_font.getsize(line_txt)
                     draw.text((text_x_start, current_y), line_txt, font=final_font, fill=text_color)
                     current_y += line_height_val + (final_line_spacing_render if line_idx < len(best_wrapped_lines) -1 else 0)
        except Exception as e_draw:
            logger.error(f"텍스트 렌더링 중 오류: {e_draw}", exc_info=True)
        return img_to_draw_on


class PaddleOcrHandler(BaseOcrHandlerImpl): # BaseOcrHandlerImpl 상속
    def __init__(self, lang_code='korean', debug_enabled=False, use_gpu=False):
        self.use_angle_cls_paddle = False
        super().__init__(lang_codes_param=lang_code, debug_enabled=debug_enabled, use_gpu_param=use_gpu)

    def _initialize_engine(self):
        try:
            from paddleocr import PaddleOCR
            logger.info(f"PaddleOCR 초기화 시도 (lang: {self.current_lang_codes}, use_angle_cls: {self.use_angle_cls_paddle}, use_gpu: {self.use_gpu}, debug: {self._debug_mode})...")
            self._ocr_engine = PaddleOCR(use_angle_cls=self.use_angle_cls_paddle, lang=self.current_lang_codes, use_gpu=self.use_gpu, show_log=self._debug_mode)
            logger.info(f"PaddleOCR 초기화 완료 (lang: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("PaddleOCR 라이브러리를 찾을 수 없습니다. 'pip install paddleocr paddlepaddle'로 설치해주세요.")
            raise RuntimeError("PaddleOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e:
            logger.error(f"PaddleOCR 초기화 중 오류 (lang: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"PaddleOCR 초기화 실패 (lang: {self.current_lang_codes}): {e}")

    def ocr_image(self, image_pil_rgb: Image.Image) -> List[Any]:
        if not self.ocr_engine: return []
        try:
            image_np_rgb = np.array(image_pil_rgb.convert('RGB'))
            ocr_output = self.ocr_engine.ocr(image_np_rgb, cls=self.use_angle_cls_paddle)
            final_parsed_results = []
            if ocr_output and isinstance(ocr_output, list) and len(ocr_output) > 0:
                results_list = ocr_output
                if isinstance(ocr_output[0], list) and \
                   (len(ocr_output[0]) == 0 or (len(ocr_output[0]) > 0 and isinstance(ocr_output[0][0], list))):
                     results_list = ocr_output[0]

                for item in results_list:
                    if isinstance(item, list) and len(item) >= 2:
                        box_data = item[0]
                        text_conf_tuple = item[1]
                        ocr_angle = None
                        if isinstance(box_data, list) and len(box_data) == 4 and \
                           all(isinstance(point, list) and len(point) == 2 for point in box_data) and \
                           isinstance(text_conf_tuple, tuple) and len(text_conf_tuple) == 2:
                            box_points_int = [[int(round(coord[0])), int(round(coord[1]))] for coord in box_data]
                            final_parsed_results.append([box_points_int, text_conf_tuple, ocr_angle])
                        else:
                            logger.warning(f"PaddleOCR 결과 항목 형식이 다릅니다 (내부): {item}")
                    else:
                        logger.warning(f"PaddleOCR 결과 항목이 리스트가 아니거나 길이가 2 미만입니다 (외부): {item}")
            return final_parsed_results
        except Exception as e:
            logger.error(f"PaddleOCR ocr_image 중 오류: {e}", exc_info=True)
            return []


class EasyOcrHandler(BaseOcrHandlerImpl): # BaseOcrHandlerImpl 상속
    def __init__(self, lang_codes_list=['en'], debug_enabled=False, use_gpu=False):
        super().__init__(lang_codes_param=lang_codes_list, debug_enabled=debug_enabled, use_gpu_param=use_gpu)

    def _initialize_engine(self):
        try:
            import easyocr
            logger.info(f"EasyOCR 초기화 시도 (langs: {self.current_lang_codes}, gpu: {self.use_gpu}, verbose: {self._debug_mode})...")
            self._ocr_engine = easyocr.Reader(self.current_lang_codes, gpu=self.use_gpu, verbose=self._debug_mode)
            logger.info(f"EasyOCR 초기화 완료 (langs: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("EasyOCR 라이브러리를 찾을 수 없습니다. 'pip install easyocr'로 설치해주세요.")
            raise RuntimeError("EasyOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e:
            logger.error(f"EasyOCR 초기화 중 오류 (langs: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"EasyOCR 초기화 실패 (langs: {self.current_lang_codes}): {e}")

    def ocr_image(self, image_pil_rgb: Image.Image) -> List[Any]:
        if not self.ocr_engine: return []
        try:
            image_np_rgb = np.array(image_pil_rgb.convert('RGB'))
            ocr_output = self.ocr_engine.readtext(image_np_rgb, detail=1, paragraph=False)
            formatted_results = []
            for item_tuple in ocr_output:
                if not (isinstance(item_tuple, (list, tuple)) and len(item_tuple) >= 2):
                    logger.warning(f"EasyOCR 결과 항목 형식이 이상합니다: {item_tuple}")
                    continue
                bbox, text = item_tuple[0], item_tuple[1]
                confidence = item_tuple[2] if len(item_tuple) > 2 else 0.9
                ocr_angle = None
                if isinstance(bbox, list) and len(bbox) == 4 and \
                   all(isinstance(p, (list, np.ndarray)) and len(p) == 2 for p in bbox):
                    box_points = [[int(round(coord[0])), int(round(coord[1]))] for coord in bbox]
                    formatted_results.append([box_points, (text, float(confidence)), ocr_angle])
                elif isinstance(bbox, np.ndarray) and bbox.shape == (4,2):
                    box_points = bbox.astype(int).tolist()
                    formatted_results.append([box_points, (text, float(confidence)), ocr_angle])
                else:
                     logger.warning(f"EasyOCR 결과의 bbox 형식이 예상과 다릅니다: {bbox}")
            return formatted_results
        except Exception as e:
            logger.error(f"EasyOCR ocr_image 중 오류: {e}", exc_info=True)
            return []

class OcrHandlerFactory(AbsOcrHandlerFactory):
    def get_ocr_handler(self, lang_code_ui: str, use_gpu: bool, debug_enabled: bool = False) -> Optional[AbsOcrHandler]:
        engine_name_display = self.get_engine_name_display(lang_code_ui)
        ocr_lang_code = self.get_ocr_lang_code(lang_code_ui)

        if not ocr_lang_code:
            logger.error(f"{engine_name_display}: UI 언어 '{lang_code_ui}'에 대한 OCR 코드가 설정되지 않았습니다.")
            return None
        
        logger.info(f"OCR Handler Factory: '{engine_name_display}' 엔진 요청 (UI 언어: {lang_code_ui}, OCR 코드: {ocr_lang_code}, GPU: {use_gpu})")

        try:
            if engine_name_display == "EasyOCR":
                if not utils.check_easyocr():
                    logger.error("EasyOCR 라이브러리가 설치되어 있지 않아 핸들러를 생성할 수 없습니다.")
                    # 여기서 설치를 시도하거나, None을 반환하여 UI에서 처리하도록 할 수 있습니다.
                    # (main.py의 check_ocr_engine_status에서 이미 설치 시도 로직이 있으므로 여기서는 생성 실패로 간주)
                    return None
                return EasyOcrHandler(lang_codes_list=[ocr_lang_code], debug_enabled=debug_enabled, use_gpu=use_gpu)
            else: # PaddleOCR (기본)
                if not utils.check_paddleocr():
                    logger.error("PaddleOCR 라이브러리가 설치되어 있지 않아 핸들러를 생성할 수 없습니다.")
                    return None
                return PaddleOcrHandler(lang_code=ocr_lang_code, debug_enabled=debug_enabled, use_gpu=use_gpu)
        except RuntimeError as e: # 핸들러 초기화 실패
            logger.error(f"{engine_name_display} 핸들러 생성 실패 (RuntimeError): {e}")
            return None
        except Exception as e_create: # 기타 예외
            logger.error(f"{engine_name_display} 핸들러 생성 중 예기치 않은 오류: {e_create}", exc_info=True)
            return None

    def get_engine_name_display(self, lang_code_ui: str) -> str:
        return "EasyOCR" if lang_code_ui in config.EASYOCR_SUPPORTED_UI_LANGS else "PaddleOCR"

    def get_ocr_lang_code(self, lang_code_ui: str) -> Optional[str]:
        if lang_code_ui in config.EASYOCR_SUPPORTED_UI_LANGS:
            return config.UI_LANG_TO_EASYOCR_CODE_MAP.get(lang_code_ui)
        else:
            return config.UI_LANG_TO_PADDLEOCR_CODE_MAP.get(lang_code_ui, config.DEFAULT_PADDLE_OCR_LANG)