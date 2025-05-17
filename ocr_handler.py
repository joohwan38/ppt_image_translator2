# ocr_handler.py
# (이전 단계에서 ocr_handler.py는 AbsOcrHandler 인터페이스를 구현하고,
# OcrHandlerFactory가 추가되었으므로, 3단계에서 큰 변경은 없습니다.)
# lru_cache, _calculate_text_dimensions 개선 등은 1단계에서 이미 반영된 것으로 간주합니다.
# 추가적으로 검토할 부분은 _calculate_text_dimensions 내부 로직의 명확성,
# Pillow 버전별 로직 분리가 이미 잘 되어 있는지 등입니다. 현재 코드는 해당 부분이
# PILLOW_VERSION_TUPLE을 통해 분기 처리되고 있어 큰 수정은 필요 없어 보입니다.

from PIL import Image, ImageDraw, ImageFont, ImageStat, __version__ as PILLOW_VERSION
import numpy as np
import cv2
import os
import logging
import io
import textwrap
import math
from typing import List, Any, Optional
import functools

import config
from interfaces import AbsOcrHandler, AbsOcrHandlerFactory
import utils

logger = logging.getLogger(__name__)

BASE_DIR_OCR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = config.FONTS_DIR

logger.info(f"OCR Handler: Using Pillow version {PILLOW_VERSION}")
PILLOW_VERSION_TUPLE = tuple(map(int, PILLOW_VERSION.split('.')))

# get_quantized_dominant_color, get_simple_average_color, get_contrasting_text_color 함수 (이전과 동일)
# ... (함수 코드 생략) ...

def get_quantized_dominant_color(image_roi, num_colors=5):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128)
        quantizable_image = image_roi.convert('RGB')
        try:
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.FASTOCTREE)
        except AttributeError: # Pillow < 9.1.0
            logger.debug("FASTOCTREE 양자화 실패, MEDIANCUT으로 대체 시도 (Pillow < 9.1.0).")
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.MEDIANCUT)
        except Exception as e_quant:
             logger.warning(f"색상 양자화 중 오류: {e_quant}. 단순 평균색으로 대체합니다.")
             return get_simple_average_color(image_roi)

        palette = quantized_image.getpalette() # RGBRGB...
        color_counts = quantized_image.getcolors(num_colors * 2) # 넉넉하게 가져옴

        if not color_counts: # 이미지가 단색이거나 매우 단순할 때 발생 가능
            logger.warning("getcolors()가 None을 반환 (양자화 실패 가능성). 단순 평균색으로 대체.")
            return get_simple_average_color(image_roi)

        dominant_palette_index = max(color_counts, key=lambda item: item[0])[1] # 가장 빈도 높은 색상의 팔레트 인덱스

        if palette: # 팔레트가 정상적으로 생성되었다면
            r = palette[dominant_palette_index * 3]
            g = palette[dominant_palette_index * 3 + 1]
            b = palette[dominant_palette_index * 3 + 2]
            dominant_color = (r, g, b)
        else: # 팔레트가 없는 경우 (거의 발생 안 함)
             logger.warning("양자화된 이미지에 팔레트가 없습니다. 단순 평균색으로 대체.")
             return get_simple_average_color(image_roi)
        return dominant_color
    except Exception as e:
        logger.warning(f"양자화된 주요 색상 감지 실패: {e}. 단순 평균색으로 대체.", exc_info=True)
        return get_simple_average_color(image_roi) # 실패 시 단순 평균색 반환

def get_simple_average_color(image_roi):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128) # 크기 0 방지
        # RGBA인 경우, 알파 채널 고려하여 평균색 계산 (흰색 배경에 합성)
        if image_roi.mode == 'RGBA':
            temp_img = Image.new("RGB", image_roi.size, (255, 255, 255))
            temp_img.paste(image_roi, mask=image_roi.split()[3]) # 알파 채널을 마스크로 사용
            avg_color_tuple = ImageStat.Stat(temp_img).mean
        else:
            avg_color_tuple = ImageStat.Stat(image_roi.convert('RGB')).mean

        return tuple(int(c) for c in avg_color_tuple[:3]) # RGB 값만 사용
    except Exception as e:
        logger.warning(f"단순 평균색 감지 실패: {e}. 기본 회색 반환.", exc_info=True)
        return (128, 128, 128) # 실패 시 기본 회색

def get_contrasting_text_color(bg_color_tuple):
    r, g, b = bg_color_tuple
    # YIQ 휘도 계산 (더 정확한 대비를 위해)
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    threshold = 128 # 밝기 임계값 (조정 가능)
    if brightness >= threshold:
        return (0, 0, 0)  # 어두운 배경에는 흰색 텍스트
    else:
        return (255, 255, 255)  # 밝은 배경에는 검은색 텍스트

class BaseOcrHandlerImpl(AbsOcrHandler):
    def __init__(self, lang_codes_param, debug_enabled=False, use_gpu_param=False):
        self._current_lang_codes = lang_codes_param
        self._debug_mode = debug_enabled
        self._use_gpu = use_gpu_param
        self._ocr_engine = None
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

    def ocr_image(self, image_pil_rgb: Image.Image) -> List[Any] :
        raise NotImplementedError("각 OCR 핸들러는 이 메서드를 구현해야 합니다.")

    def has_text_in_image_bytes(self, image_bytes: bytes) -> bool:
        # (이전 코드와 동일)
        if not self.ocr_engine: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if img_pil.width < 5 or img_pil.height < 5: return False # 매우 작은 이미지는 텍스트 가능성 낮음
            img_pil_rgb = img_pil.convert("RGB") # OCR은 RGB 이미지로 처리
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False # 변환 후 크기 0 방지

            results = self.ocr_image(img_pil_rgb) # 실제 OCR 수행
            # 결과가 있고, 각 결과의 텍스트(res[1][0])가 공백이 아닌 경우 True
            return bool(results and any(res[1][0].strip() for res in results if len(res) > 1 and len(res[1]) > 0))

        except OSError as e: # Pillow가 이미지 파일을 열 수 없을 때
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.warning(f"이미지 텍스트 확인 중 Pillow OSError ({format_info}), 건너뜀: {e}", exc_info=False) # 스택 트레이스 없이 경고만
            return False
        except Exception as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            logger.error(f"이미지 텍스트 확인 중 예기치 않은 오류 ({format_info}): {e}", exc_info=True)
            return False
        finally:
            if img_pil:
                try: img_pil.close()
                except Exception: pass


    @functools.lru_cache(maxsize=128) # 1단계에서 이미 적용됨
    def _get_font(self, font_size: int, lang_code: str = 'en', is_bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont :
        # (1단계 코드와 동일)
        font_size = max(1, int(font_size)) # 0 이하 방지
        font_filename = None
        font_path = None

        language_font_map = config.OCR_LANGUAGE_FONT_MAP
        default_font_filename = config.OCR_DEFAULT_FONT_FILENAME
        default_bold_font_filename = config.OCR_DEFAULT_BOLD_FONT_FILENAME

        # 볼드체 우선 확인
        if is_bold:
            bold_font_key = lang_code + '_bold'
            font_filename = language_font_map.get(bold_font_key)
            if not font_filename: # 특정 언어 볼드체가 없으면 기본 볼드체
                font_filename = default_bold_font_filename

        # 볼드체가 아니거나, 볼드체 폰트를 못 찾았으면 일반 폰트 검색
        if not font_filename:
            font_filename = language_font_map.get(lang_code, default_font_filename)

        # 그래도 못 찾았으면 최종 기본값
        if not font_filename:
            font_filename = default_font_filename if not is_bold else default_bold_font_filename

        if font_filename:
            font_path = os.path.join(FONT_DIR, font_filename)

        if font_path and os.path.exists(font_path):
            try:
                return ImageFont.truetype(font_path, int(font_size))
            except IOError as e:
                logger.warning(f"트루타입 폰트 로드 실패 ('{font_path}', size:{font_size}): {e}. Pillow 기본 폰트로 대체.")
            except Exception as e_font: # 기타 예외 (예: 글꼴 파일 손상)
                logger.error(f"폰트 로드 중 예기치 않은 오류 ('{font_path}', size:{font_size}): {e_font}. Pillow 기본 폰트로 대체.", exc_info=True)
        else:
            logger.warning(f"폰트 파일 없음: '{font_path or font_filename}' (요청 코드: {lang_code}, bold: {is_bold}). Pillow 기본 폰트 사용.")

        # Pillow 기본 폰트 사용 (fallback)
        try:
            if PILLOW_VERSION_TUPLE >= (10, 0, 0): # Pillow 10.0.0부터 load_default 인자 없음
                 return ImageFont.load_default()
            elif PILLOW_VERSION_TUPLE >= (9, 0, 0): # Pillow 9.x.x는 size 인자 지원
                 return ImageFont.load_default(size=int(font_size))
            else: # 그 이전 버전
                 return ImageFont.load_default()
        except TypeError: # load_default에서 size 인자 지원 안하는 오래된 Pillow 대비
            try:
                return ImageFont.load_default()
            except Exception as e_default_font_fallback:
                logger.critical(f"Pillow 기본 폰트 로드조차 실패 (size={font_size}): {e_default_font_fallback}. 글꼴 렌더링 불가.", exc_info=True)
                raise RuntimeError(f"기본 폰트 로드 실패: {e_default_font_fallback}") # 이 경우 렌더링 불가
        except Exception as e_default_font:
            logger.critical(f"Pillow 기본 폰트 로드 실패 (size={font_size}): {e_default_font}. 글꼴 렌더링 불가.", exc_info=True)
            raise RuntimeError(f"기본 폰트 로드 실패: {e_default_font}")


    def _calculate_text_dimensions(self, draw: ImageDraw.ImageDraw, text: str, font_size: int,
                                   render_area_width: int, lang_code: str, is_bold: bool, line_spacing: int) -> tuple[int, int, List[str]]:
        # (1단계 코드와 동일 - 내부 로직 분리 검토는 이미 반영된 것으로 간주)
        if font_size < 1: font_size = 1
        current_font = self._get_font(font_size, lang_code=lang_code, is_bold=is_bold)

        # 줄 바꿈을 위한 한 줄당 예상 글자 수 계산 (TextWrapper용)
        estimated_chars_per_line = 1
        if render_area_width > 0: # 너비가 0보다 클 때만 의미 있음
            try:
                # Pillow 버전에 따라 textlength 또는 getlength/getsize 사용
                char_w_metric = 0
                if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'textlength'):
                    char_w_metric = draw.textlength("W", font=current_font) # 영어 대문자 기준
                    if char_w_metric <= 0: char_w_metric = draw.textlength("가", font=current_font) # 한글 기준 (대체)
                elif hasattr(current_font, 'getlength'): # Pillow < 9.2.0
                    char_w_metric = current_font.getlength("W")
                    if char_w_metric <= 0: char_w_metric = current_font.getlength("가")
                elif hasattr(current_font, 'getsize'): # 더 오래된 Pillow
                     char_w_metric, _ = current_font.getsize("W")
                     if char_w_metric <= 0 : char_w_metric, _ = current_font.getsize("가")

                if char_w_metric > 0:
                    estimated_chars_per_line = max(1, int(render_area_width / char_w_metric))
                else: # 글자 너비 측정 실패 시, 폰트 크기 기반으로 대략적 계산
                    estimated_chars_per_line = max(1, int(render_area_width / (font_size * 0.5 + 1))) # 0.5는 평균적 비율, +1은 최소 너비 보장
            except Exception as e_char_width:
                logger.debug(f"문자 너비 계산 중 예외: {e_char_width}. 근사치 사용.")
                estimated_chars_per_line = max(1, int(render_area_width / (font_size * 0.6))) # 0.6은 경험적 비율

        # TextWrapper를 사용하여 자동 줄 바꿈
        wrapper = textwrap.TextWrapper(width=estimated_chars_per_line, break_long_words=True,
                                       replace_whitespace=False, drop_whitespace=False, # 공백 유지 중요
                                       break_on_hyphens=True)
        wrapped_lines = wrapper.wrap(text)
        if not wrapped_lines: wrapped_lines = [" "] # 빈 텍스트라도 한 줄은 유지 (렌더링 오류 방지)


        rendered_text_height = 0
        rendered_text_width = 0

        # Pillow 9.2.0 이상이고 multiline_textbbox 지원 시 사용
        if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'multiline_textbbox'):
            try:
                text_bbox_args = {'xy': (0,0), 'text': "\n".join(wrapped_lines), 'font': current_font, 'spacing': line_spacing}
                if PILLOW_VERSION_TUPLE >= (9, 3, 0): # anchor 인자 지원 버전
                    text_bbox_args['anchor'] = "lt" # left-top 기준
                
                text_bbox = draw.multiline_textbbox(**text_bbox_args)
                rendered_text_width = text_bbox[2] - text_bbox[0]
                rendered_text_height = text_bbox[3] - text_bbox[1]
            except Exception as e_mtbox:
                logger.debug(f"multiline_textbbox 사용 중 예외: {e_mtbox}. 수동 계산으로 대체.")
                # 수동 계산 폴백
                rendered_text_width, rendered_text_height = self._manual_calculate_multiline_dimensions(draw, wrapped_lines, current_font, line_spacing, font_size)
        else: # 구버전 Pillow 또는 multiline_textbbox 실패 시 수동 계산
            rendered_text_width, rendered_text_height = self._manual_calculate_multiline_dimensions(draw, wrapped_lines, current_font, line_spacing, font_size)

        return int(rendered_text_width), int(rendered_text_height), wrapped_lines

    def _manual_calculate_multiline_dimensions(self, draw: ImageDraw.ImageDraw, wrapped_lines: List[str],
                                             font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
                                             line_spacing: int, fallback_font_size: int) -> tuple[int, int]:
        # (1단계 코드와 동일)
        total_h = 0
        max_w = 0
        for i, line_txt in enumerate(wrapped_lines):
            line_w, line_h = 0, 0
            try:
                if hasattr(draw, 'textbbox'): # Pillow 8.0.0+
                    bbox_args = {'xy': (0,0), 'text': line_txt, 'font': font}
                    if PILLOW_VERSION_TUPLE >= (9, 3, 0): bbox_args['anchor'] = "lt"
                    line_bbox = draw.textbbox(**bbox_args)
                    line_w = line_bbox[2] - line_bbox[0]
                    line_h = line_bbox[3] - line_bbox[1]
                elif hasattr(font, 'getsize'): # 구버전
                    line_w, line_h = font.getsize(line_txt)
                elif hasattr(font, 'getbbox'): # 또 다른 구버전 인터페이스 (거의 사용 안 함)
                    bbox = font.getbbox(line_txt)
                    line_w = bbox[2] - bbox[0]
                    line_h = bbox[3] - bbox[1]
                else: # 최후의 수단
                    line_w = len(line_txt) * fallback_font_size * 0.6 # 대략적 계산
                    line_h = fallback_font_size
            except Exception as e_line_calc:
                logger.debug(f"개별 라인 크기 계산 중 예외: {e_line_calc}. 근사치 사용.")
                line_w = len(line_txt) * fallback_font_size * 0.6
                line_h = fallback_font_size

            total_h += line_h
            if line_w > max_w:
                max_w = line_w
            if i < len(wrapped_lines) - 1: # 마지막 줄 제외하고 줄 간격 추가
                total_h += line_spacing
        return int(max_w), int(total_h)

    def render_translated_text_on_image(self, image_pil_original: Image.Image, box: List[List[int]], translated_text: str,
                                        font_code_for_render='en', original_text="", ocr_angle=None) -> Image.Image:
        # (1단계 코드와 거의 동일)
        # ... (코드 생략, 내부 로직은 이전과 동일)
        img_to_draw_on = image_pil_original.copy()
        draw = ImageDraw.Draw(img_to_draw_on)

        try:
            # 바운딩 박스 좌표 추출 및 유효성 검사
            x_coords = [p[0] for p in box]
            y_coords = [p[1] for p in box]
            min_x, max_x = min(x_coords), max(x_coords)
            min_y, max_y = min(y_coords), max(y_coords)

            if max_x <= min_x or max_y <= min_y: # 너비나 높이가 0 이하인 경우
                logger.warning(f"렌더링 스킵: 유효하지 않은 바운딩 박스 {box} for '{translated_text[:20]}...'")
                return image_pil_original # 원본 이미지 그대로 반환

            # 이미지 경계 내로 좌표 조정
            img_w, img_h = img_to_draw_on.size
            render_box_x1 = max(0, int(min_x))
            render_box_y1 = max(0, int(min_y))
            render_box_x2 = min(img_w, int(max_x))
            render_box_y2 = min(img_h, int(max_y))

            if render_box_x2 <= render_box_x1 or render_box_y2 <= render_box_y1: # 조정 후 크기가 0인 경우
                logger.warning(f"렌더링 스킵: 크기가 0인 렌더 박스 for '{translated_text[:20]}...'")
                return img_to_draw_on # 변경된 이미지 반환 (배경색은 칠해졌을 수 있음)

            # 원본 bbox와 실제 렌더링될 bbox 크기
            bbox_width_orig = max_x - min_x
            bbox_height_orig = max_y - min_y
            bbox_width_render = render_box_x2 - render_box_x1
            bbox_height_render = render_box_y2 - render_box_y1

        except Exception as e_box_calc:
            logger.error(f"렌더링 바운딩 박스 계산 오류: {e_box_calc}. Box: {box}. 원본 이미지 반환.", exc_info=True)
            return image_pil_original # 예외 발생 시 원본 반환

        # 배경색 추정 및 텍스트 영역 덮어쓰기
        try:
            text_roi_pil = image_pil_original.crop((render_box_x1, render_box_y1, render_box_x2, render_box_y2))
            estimated_bg_color = get_quantized_dominant_color(text_roi_pil) if text_roi_pil.width > 0 and text_roi_pil.height > 0 else (200,200,200) # ROI 크기 0 방지
        except Exception as e_bg:
            logger.warning(f"배경색 추정 실패 ({e_bg}), 기본 회색 사용.", exc_info=True)
            estimated_bg_color = (200, 200, 200) # 기본값 (밝은 회색)

        draw.rectangle([render_box_x1, render_box_y1, render_box_x2, render_box_y2], fill=estimated_bg_color)
        text_color = get_contrasting_text_color(estimated_bg_color) # 배경색과 대비되는 텍스트 색상 결정

        # 텍스트 렌더링 영역 설정 (패딩 적용)
        padding_x = max(1, int(bbox_width_render * 0.03)) # 너비의 3%, 최소 1px
        padding_y = max(1, int(bbox_height_render * 0.03)) # 높이의 3%, 최소 1px

        render_area_x_start = render_box_x1 + padding_x
        render_area_y_start = render_box_y1 + padding_y
        render_area_width = bbox_width_render - 2 * padding_x
        render_area_height = bbox_height_render - 2 * padding_y

        if render_area_width <= 1 or render_area_height <= 1: # 패딩 적용 후 렌더링 영역이 너무 작으면
            logger.debug(f"렌더링 영역 너무 작음 ({render_area_width}x{render_area_height}), 텍스트 없이 배경만 칠해진 이미지 반환.")
            return img_to_draw_on # 배경만 칠해진 이미지 반환

        # 폰트 크기 동적 조절 (이진 탐색 방식)
        # ... (폰트 크기 조절 로직은 이전과 동일)
        font_size_correction_factor = 1.0
        text_angle_deg = 0.0
        if ocr_angle is not None and isinstance(ocr_angle, (int, float)): # OCR 각도 정보 활용
            text_angle_deg = abs(ocr_angle) # 절대값 사용
            # 수평/수직에 가까운 각도는 보정 덜하고, 대각선에 가까울수록 더 많이 보정
            if 5 < text_angle_deg < 85 or 95 < text_angle_deg < 175: # 약 5도 이상 기울어진 경우
                font_size_correction_factor = max(0.6, 1.0 - (text_angle_deg / 90.0) * 0.3) # 기울기에 따라 최대 30%까지 크기 줄임
        elif bbox_width_orig > 0 and bbox_height_orig > 0 : # 원본 bbox 비율로 기울기 추정 (OCR 각도 없을 때)
            aspect_ratio_orig = bbox_width_orig / bbox_height_orig
            if aspect_ratio_orig > 2.0 or aspect_ratio_orig < 0.5: # 가로/세로로 긴 경우 (기울어졌을 가능성)
                font_size_correction_factor = 0.80 # 20% 줄임

        initial_target_font_size = int(min(render_area_height * 0.9, # 높이의 90%
                                    render_area_width * 0.9 / (len(translated_text.splitlines()[0] if translated_text else "A")*0.5 +1) # 너비와 글자 수 고려, *0.5는 평균적 글자 너비 비율
                                   ) * font_size_correction_factor) # 기울기 보정 적용
        initial_target_font_size = max(initial_target_font_size, 1) # 최소 1

        min_font_size = 5 # 너무 작은 폰트는 가독성 저하
        if initial_target_font_size < min_font_size: initial_target_font_size = min_font_size

        is_bold_font = '_bold' in font_code_for_render or 'bold' in font_code_for_render.lower()
        best_fit_size = min_font_size # 최적 폰트 크기
        best_wrapped_lines: List[str] = []
        best_text_width = 0
        best_text_height = 0
        # 이진 탐색으로 최적 폰트 크기 찾기
        low = min_font_size
        high = initial_target_font_size
        max_iterations = int(math.log2(high - low + 1)) + 5 if high > low else 5 # 최대 반복 횟수 (무한 루프 방지)
        current_iteration = 0

        while low <= high and current_iteration < max_iterations:
            current_iteration +=1
            mid_font_size = low + (high - low) // 2
            if mid_font_size < min_font_size : mid_font_size = min_font_size # 최소 크기 보장
            if mid_font_size == 0 : break # 폰트 크기 0 방지

            current_line_spacing = int(mid_font_size * 0.2) # 줄 간격은 폰트 크기의 20%
            w, h, wrapped = self._calculate_text_dimensions(draw, translated_text, mid_font_size,
                                                            render_area_width, font_code_for_render,
                                                            is_bold_font, current_line_spacing)
            if w <= render_area_width and h <= render_area_height: # 렌더링 영역 안에 들어오면
                best_fit_size = mid_font_size
                best_wrapped_lines = wrapped
                best_text_width = w
                best_text_height = h
                low = mid_font_size + 1 # 더 큰 폰트 시도
            else: # 영역을 벗어나면
                high = mid_font_size - 1 # 더 작은 폰트 시도

        # 최종적으로 찾은 best_fit_size 또는 min_font_size 사용
        if not best_wrapped_lines: # 이진 탐색 실패 또는 매우 작은 영역으로 인해 best_fit_size가 min_font_size보다 작을 때
            final_line_spacing = int(min_font_size * 0.2)
            best_text_width, best_text_height, best_wrapped_lines = self._calculate_text_dimensions(
                draw, translated_text, min_font_size, render_area_width, font_code_for_render, is_bold_font, final_line_spacing
            )
            best_fit_size = min_font_size # 최소 폰트 크기로 강제

        final_font_size = best_fit_size
        final_font = self._get_font(final_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
        final_line_spacing_render = int(final_font_size * 0.2) # 최종 줄간격

        # 텍스트 위치 계산 (가운데 정렬)
        text_x_start = render_area_x_start + (render_area_width - best_text_width) / 2
        text_y_start = render_area_y_start + (render_area_height - best_text_height) / 2
        # 렌더링 영역을 벗어나지 않도록 보정
        text_x_start = max(render_area_x_start, text_x_start)
        text_y_start = max(render_area_y_start, text_y_start)


        # 텍스트 그리기 (Pillow 버전에 따라 multiline_text 또는 수동 라인별 draw.text 사용)
        try:
            if PILLOW_VERSION_TUPLE >= (9,0,0) and hasattr(draw, 'multiline_text'): # Pillow 9.0.0+
                multiline_args = {
                    'xy': (text_x_start, text_y_start),
                    'text': "\n".join(best_wrapped_lines),
                    'font': final_font,
                    'fill': text_color,
                    'spacing': final_line_spacing_render,
                    'align': "left" # 기본 왼쪽 정렬 (가운데 정렬은 xy 좌표로 이미 처리)
                }
                if PILLOW_VERSION_TUPLE >= (9,3,0): # Pillow 9.3.0+ anchor 지원
                    multiline_args['anchor'] = "la" # Left-Ascent 기준 (보다 정확한 수직 정렬)
                draw.multiline_text(**multiline_args)
            else: # 구버전 Pillow: 수동으로 각 라인 그리기
                 current_y = text_y_start
                 for line_idx, line_txt in enumerate(best_wrapped_lines):
                     # 각 라인의 실제 높이 계산 (폰트마다 다를 수 있음)
                     line_height_val = final_font_size # 기본값
                     if hasattr(draw, 'textbbox'): # Pillow 8.0.0+
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
            # 이 경우에도 배경은 칠해진 이미지가 반환될 수 있음

        return img_to_draw_on

class PaddleOcrHandler(BaseOcrHandlerImpl):
    def __init__(self, lang_code='korean', debug_enabled=False, use_gpu=False):
        self.use_angle_cls_paddle = False # PaddleOCR 각도 분류기 사용 여부 (False로 고정)
        super().__init__(lang_codes_param=lang_code, debug_enabled=debug_enabled, use_gpu_param=use_gpu)

    def _initialize_engine(self):
        # (이전 코드와 동일)
        try:
            from paddleocr import PaddleOCR # 여기서 import 시도
            logger.info(f"PaddleOCR 초기화 시도 (lang: {self.current_lang_codes}, use_angle_cls: {self.use_angle_cls_paddle}, use_gpu: {self.use_gpu}, debug: {self._debug_mode})...")
            # PaddleOCR 초기화 시 show_log는 debug 모드일 때만 True로 설정
            self._ocr_engine = PaddleOCR(use_angle_cls=self.use_angle_cls_paddle, lang=self.current_lang_codes, use_gpu=self.use_gpu, show_log=self._debug_mode)
            logger.info(f"PaddleOCR 초기화 완료 (lang: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("PaddleOCR 라이브러리를 찾을 수 없습니다. 'pip install paddleocr paddlepaddle'로 설치해주세요.")
            raise RuntimeError("PaddleOCR 라이브러리가 설치되어 있지 않습니다.") # 호출한 쪽에서 처리하도록 예외 발생
        except Exception as e:
            logger.error(f"PaddleOCR 초기화 중 오류 (lang: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"PaddleOCR 초기화 실패 (lang: {self.current_lang_codes}): {e}")


    def ocr_image(self, image_pil_rgb: Image.Image) -> List[Any]:
        # (이전 코드와 동일)
        if not self.ocr_engine: return []
        try:
            image_np_rgb = np.array(image_pil_rgb.convert('RGB')) # PIL 이미지를 NumPy 배열로 변환
            # PaddleOCR은 NumPy 배열을 입력으로 받음
            ocr_output = self.ocr_engine.ocr(image_np_rgb, cls=self.use_angle_cls_paddle)
            # ocr_output 형식: [[box, (text, confidence)], [box, (text, confidence)], ...]
            # 또는 [[[box, (text, confidence)], ...]] (이미지 배치 처리 시) - 여기서는 단일 이미지이므로 전자 가정
            
            final_parsed_results = []
            if ocr_output and isinstance(ocr_output, list) and len(ocr_output) > 0:
                # PaddleOCR 2.5+ 버전은 결과가 [[box, (text, confidence)], ...] 형식이 아닌
                # [[[box, (text, confidence)], ...]] 형태로 나올 수 있음 (이미지 1개라도)
                # 이를 처리하기 위해 한 번 더 확인
                results_list = ocr_output
                if isinstance(ocr_output[0], list) and \
                   (len(ocr_output[0]) == 0 or (len(ocr_output[0]) > 0 and isinstance(ocr_output[0][0], list))):
                     results_list = ocr_output[0] # 실제 결과 리스트는 한 단계 더 안쪽에 있음

                for item in results_list:
                    # item: [ [[x1,y1],[x2,y2],[x3,y3],[x4,y4]], ('text', confidence_score) ]
                    if isinstance(item, list) and len(item) >= 2:
                        box_data = item[0] # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
                        text_conf_tuple = item[1] # ('text', confidence_score)
                        ocr_angle = None # PaddleOCR은 각도 정보를 직접 반환하지 않음 (cls로 처리)

                        # 데이터 형식 유효성 검사
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


class EasyOcrHandler(BaseOcrHandlerImpl):
    def __init__(self, lang_codes_list=['en'], debug_enabled=False, use_gpu=False):
        super().__init__(lang_codes_param=lang_codes_list, debug_enabled=debug_enabled, use_gpu_param=use_gpu)

    def _initialize_engine(self):
        # (이전 코드와 동일)
        try:
            import easyocr # 여기서 import 시도
            logger.info(f"EasyOCR 초기화 시도 (langs: {self.current_lang_codes}, gpu: {self.use_gpu}, verbose: {self._debug_mode})...")
            # EasyOCR 초기화 시 verbose는 debug 모드일 때만 True로 설정
            self._ocr_engine = easyocr.Reader(self.current_lang_codes, gpu=self.use_gpu, verbose=self._debug_mode)
            logger.info(f"EasyOCR 초기화 완료 (langs: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("EasyOCR 라이브러리를 찾을 수 없습니다. 'pip install easyocr'로 설치해주세요.")
            raise RuntimeError("EasyOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e:
            logger.error(f"EasyOCR 초기화 중 오류 (langs: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"EasyOCR 초기화 실패 (langs: {self.current_lang_codes}): {e}")


    def ocr_image(self, image_pil_rgb: Image.Image) -> List[Any]:
        # (이전 코드와 동일)
        if not self.ocr_engine: return []
        try:
            image_np_rgb = np.array(image_pil_rgb.convert('RGB')) # PIL 이미지를 NumPy 배열로 변환
            # EasyOCR은 NumPy 배열 또는 파일 경로를 입력으로 받음
            # paragraph=False: 개별 텍스트 라인으로 인식 (True로 하면 단락 단위로 묶으려 시도)
            # detail=1: 좌표, 텍스트, 신뢰도 모두 반환 (0은 텍스트만)
            ocr_output = self.ocr_engine.readtext(image_np_rgb, detail=1, paragraph=False)
            # ocr_output 형식: [([[x1,y1],[x2,y2],[x3,y3],[x4,y4]], 'text', confidence_score), ...]

            formatted_results = []
            for item_tuple in ocr_output:
                # item_tuple: (bbox, text, confidence)
                if not (isinstance(item_tuple, (list, tuple)) and len(item_tuple) >= 2): # 최소 bbox와 text는 있어야 함
                    logger.warning(f"EasyOCR 결과 항목 형식이 이상합니다: {item_tuple}")
                    continue

                bbox, text = item_tuple[0], item_tuple[1]
                confidence = item_tuple[2] if len(item_tuple) > 2 else 0.9 # 신뢰도 없으면 0.9로 가정
                ocr_angle = None # EasyOCR은 각도 정보를 직접 반환하지 않음

                # bbox 형식 확인 및 변환
                if isinstance(bbox, list) and len(bbox) == 4 and \
                   all(isinstance(p, (list, np.ndarray)) and len(p) == 2 for p in bbox): # [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
                    box_points = [[int(round(coord[0])), int(round(coord[1]))] for coord in bbox]
                    formatted_results.append([box_points, (text, float(confidence)), ocr_angle])
                elif isinstance(bbox, np.ndarray) and bbox.shape == (4,2): # NumPy 배열 형태
                    box_points = bbox.astype(int).tolist()
                    formatted_results.append([box_points, (text, float(confidence)), ocr_angle])
                else: # 예상치 못한 bbox 형식
                     logger.warning(f"EasyOCR 결과의 bbox 형식이 예상과 다릅니다: {bbox}")
            return formatted_results
        except Exception as e:
            logger.error(f"EasyOCR ocr_image 중 오류: {e}", exc_info=True)
            return []

class OcrHandlerFactory(AbsOcrHandlerFactory):
    # (이전 코드와 동일)
    def get_ocr_handler(self, lang_code_ui: str, use_gpu: bool, debug_enabled: bool = False) -> Optional[AbsOcrHandler]:
        engine_name_display = self.get_engine_name_display(lang_code_ui)
        ocr_lang_code = self.get_ocr_lang_code(lang_code_ui)

        if not ocr_lang_code:
            logger.error(f"{engine_name_display}: UI 언어 '{lang_code_ui}'에 대한 OCR 코드가 설정되지 않았습니다.")
            return None
        
        logger.info(f"OCR Handler Factory: '{engine_name_display}' 엔진 요청 (UI 언어: {lang_code_ui}, OCR 코드: {ocr_lang_code}, GPU: {use_gpu})")

        try:
            if engine_name_display == "EasyOCR":
                if not utils.check_easyocr(): # 설치 확인
                    logger.error("EasyOCR 라이브러리가 설치되어 있지 않아 핸들러를 생성할 수 없습니다.")
                    return None
                return EasyOcrHandler(lang_codes_list=[ocr_lang_code], debug_enabled=debug_enabled, use_gpu=use_gpu)
            else: # PaddleOCR (기본)
                if not utils.check_paddleocr(): # 설치 확인
                    logger.error("PaddleOCR 라이브러리가 설치되어 있지 않아 핸들러를 생성할 수 없습니다.")
                    return None
                return PaddleOcrHandler(lang_code=ocr_lang_code, debug_enabled=debug_enabled, use_gpu=use_gpu)
        except RuntimeError as e: # 핸들러 초기화 실패 (엔진 내부 오류 등)
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