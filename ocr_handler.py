from PIL import Image, ImageDraw, ImageFont, ImageStat, __version__ as PILLOW_VERSION
import numpy as np
import cv2 # PaddleOCR 전처리 등에 필요할 수 있음
import os
import logging
import io
import textwrap
import math # 각도 계산 등 기하학적 연산에 필요할 수 있음

logger = logging.getLogger(__name__)

BASE_DIR_OCR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR_OCR, "fonts")

# 언어별 폰트 매핑 (기존과 동일)
LANGUAGE_FONT_MAP = {
    'korean': 'NotoSansCJK-Regular.ttc', 'japan': 'NotoSansCJK-Regular.ttc',
    'ch': 'NotoSansCJK-Regular.ttc', 'chinese_cht': 'NotoSansCJK-Regular.ttc',
    'en': 'NotoSansCJK-Regular.ttc', 'th': 'NotoSansThai-VariableFont_wdth,wght.ttf',
    'es': 'NotoSansCJK-Regular.ttc', # 스페인어도 NotoSansCJK로 우선 사용 (범용성)
    'korean_bold': 'NotoSansCJK-Bold.ttc', 'japan_bold': 'NotoSansCJK-Bold.ttc',
    'ch_bold': 'NotoSansCJK-Bold.ttc', 'chinese_cht_bold': 'NotoSansCJK-Bold.ttc',
    'en_bold': 'NotoSansCJK-Bold.ttc', 'th_bold': 'NotoSansThai-VariableFont_wdth,wght.ttf', # 굵은 태국어 폰트 필요시 추가
    'es_bold': 'NotoSansCJK-Bold.ttc',
}
DEFAULT_FONT_FILENAME = 'NotoSansCJK-Regular.ttc' # 가장 많은 문자를 커버하는 폰트
DEFAULT_BOLD_FONT_FILENAME = 'NotoSansCJK-Bold.ttc'

logger.info(f"OCR Handler: Using Pillow version {PILLOW_VERSION}")
PILLOW_VERSION_TUPLE = tuple(map(int, PILLOW_VERSION.split('.')))


# --- 스타일 추정 함수 (get_quantized_dominant_color, get_simple_average_color, get_contrasting_text_color - 기존과 동일) ---
def get_quantized_dominant_color(image_roi, num_colors=5): # num_colors 약간 줄여봄
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128) # 기본 회색
        quantizable_image = image_roi.convert('RGB')
        try:
            # Pillow 9.1.0 부터 FASTOCTREE 사용 가능
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.FASTOCTREE)
        except AttributeError: 
            logger.debug("FASTOCTREE 양자화 실패, MEDIANCUT으로 대체 시도 (Pillow < 9.1.0).")
            quantized_image = quantizable_image.quantize(colors=num_colors, method=Image.Quantize.MEDIANCUT)
        except Exception as e_quant:
             logger.warning(f"색상 양자화 중 오류: {e_quant}. 단순 평균색으로 대체합니다.")
             return get_simple_average_color(image_roi)

        palette = quantized_image.getpalette() # RGBRGB...
        color_counts = quantized_image.getcolors(num_colors * 2) # 충분한 수의 색상 카운트 가져오기
        
        if not color_counts:
            logger.warning("getcolors()가 None을 반환 (양자화 실패 가능성). 단순 평균색으로 대체.")
            return get_simple_average_color(image_roi)

        dominant_palette_index = max(color_counts, key=lambda item: item[0])[1]
        
        if palette:
            r = palette[dominant_palette_index * 3]
            g = palette[dominant_palette_index * 3 + 1]
            b = palette[dominant_palette_index * 3 + 2]
            dominant_color = (r, g, b)
        else: # 팔레트가 없는 이미지 (예: 이미 RGB인데 getcolors()를 사용한 경우)
             logger.warning("양자화된 이미지에 팔레트가 없습니다. getcolors() 인덱스를 직접 색상으로 사용 시도.")
             # 이 경우는 드물지만, getcolors()가 (count, color_tuple) 형태를 반환할 수도 있음
             # 여기서는 dominant_palette_index가 실제 색상 튜플이라고 가정하지 않음.
             # 안전하게 평균색으로 대체.
             return get_simple_average_color(image_roi)
        return dominant_color
    except Exception as e:
        logger.warning(f"양자화된 주요 색상 감지 실패: {e}. 단순 평균색으로 대체.", exc_info=True)
        return get_simple_average_color(image_roi)

def get_simple_average_color(image_roi):
    try:
        if image_roi.width == 0 or image_roi.height == 0: return (128, 128, 128)
        # RGBA인 경우 알파 채널을 고려하여 평균색 계산 (흰 배경에 합성)
        if image_roi.mode == 'RGBA':
            temp_img = Image.new("RGB", image_roi.size, (255, 255, 255)) # 흰 배경
            temp_img.paste(image_roi, mask=image_roi.split()[3]) # 알파 채널을 마스크로 사용
            avg_color_tuple = ImageStat.Stat(temp_img).mean
        else: # RGB 또는 기타 모드 (RGB로 변환하여 평균 계산)
            avg_color_tuple = ImageStat.Stat(image_roi.convert('RGB')).mean
        
        # mean 결과는 float일 수 있으므로 int로 변환
        return tuple(int(c) for c in avg_color_tuple[:3])
    except Exception as e:
        logger.warning(f"단순 평균색 감지 실패: {e}. 기본 회색 반환.", exc_info=True)
        return (128, 128, 128)

def get_contrasting_text_color(bg_color_tuple):
    r, g, b = bg_color_tuple
    # YIQ 공식 (밝기 계산)
    brightness = (r * 299 + g * 587 + b * 114) / 1000
    threshold = 128 # 밝기 임계값 (0-255 범위)
    if brightness >= threshold:
        return (0, 0, 0)  # 배경이 밝으면 검은색 텍스트
    else:
        return (255, 255, 255)  # 배경이 어두우면 흰색 텍스트
# --- 스타일 추정 함수 끝 ---


class BaseOcrHandler:
    def __init__(self, lang_codes, debug_enabled=False, use_gpu=False):
        self.current_lang_codes = lang_codes 
        self.debug_mode = debug_enabled
        self.use_gpu = use_gpu
        self.ocr_engine = None
        self._initialize_engine()

    def _initialize_engine(self):
        raise NotImplementedError(" 각 OCR 핸들러는 이 메서드를 구현해야 합니다.")

    def ocr_image(self, image_pil_rgb):
        """
        주어진 PIL Image 객체(RGB)에서 텍스트를 OCR합니다.
        반환 형식: [[box_points, (text, confidence), optional_angle], ...]
        box_points: [[x1,y1],[x2,y1],[x2,y2],[x1,y2]] 또는 4개의 점으로 이루어진 폴리곤
        optional_angle: OCR 엔진이 각도 정보를 제공하는 경우 float 값, 아니면 None
        """
        raise NotImplementedError("각 OCR 핸들러는 이 메서드를 구현해야 합니다.")

    def has_text_in_image_bytes(self, image_bytes):
        # (기존 로직 유지 - 필요시 ocr_image 반환값 변경에 따라 수정)
        if not self.ocr_engine: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if img_pil.width < 5 or img_pil.height < 5: return False # 너무 작은 이미지 스킵
            img_pil_rgb = img_pil.convert("RGB")
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False
            
            # ocr_image 결과가 [[box, (text, conf), angle], ...] 형태이므로, 실제 텍스트가 있는지 확인
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
    
    def _get_font(self, font_size, lang_code='en', is_bold=False):
        # (기존 로직 유지)
        font_filename = None
        font_path = None

        # 1. 볼드체 우선 검색
        if is_bold:
            bold_font_key = lang_code + '_bold'
            font_filename = LANGUAGE_FONT_MAP.get(bold_font_key)
            if not font_filename: # 특정 언어 볼드체가 없으면 기본 볼드체
                font_filename = DEFAULT_BOLD_FONT_FILENAME
        
        # 2. 볼드체가 아니거나, 볼드체 검색 실패 시 일반 폰트 검색
        if not font_filename:
            font_filename = LANGUAGE_FONT_MAP.get(lang_code, DEFAULT_FONT_FILENAME)
        
        # 3. 최종적으로도 못 찾았으면 절대 기본 폰트 사용 (이 경우는 거의 없음)
        if not font_filename:
            font_filename = DEFAULT_FONT_FILENAME if not is_bold else DEFAULT_BOLD_FONT_FILENAME
            logger.warning(f"요청된 폰트 코드 '{lang_code}'(bold:{is_bold})에 대한 폰트 매핑 없음. 기본 폰트 '{font_filename}' 사용.")

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
        
        # Pillow 기본 폰트 반환 (size 인자 시도, Pillow 9. keine size mehr)
        try:
            if PILLOW_VERSION_TUPLE >= (9, 0, 0): # Pillow 9.0.0 이상은 size 인자 없음
                 return ImageFont.load_default()
            else: # 구버전
                 return ImageFont.load_default(size=int(font_size)) # 일부 구버전에서 size 지원
        except TypeError: # size 인자 없는 최신 버전 또는 매우 구 버전
            return ImageFont.load_default()
        except Exception as e_default_font: # 기본 폰트 로드 실패 (드문 경우)
            logger.critical(f"Pillow 기본 폰트 로드조차 실패: {e_default_font}. 글꼴 렌더링 불가.", exc_info=True)
            # 이 경우 None을 반환하거나, 예외를 다시 발생시켜 상위에서 처리하도록 할 수 있음
            raise RuntimeError(f"기본 폰트 로드 실패: {e_default_font}")


    def render_translated_text_on_image(self, image_pil_original, box, translated_text,
                                        font_code_for_render='en', original_text="", ocr_angle=None):
        # ocr_angle: OCR 엔진이 제공한 텍스트 블록의 각도 (선택적)
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
            # Pillow crop/rectangle 등은 정수 좌표를 기대
            render_box_x1 = max(0, int(min_x))
            render_box_y1 = max(0, int(min_y))
            render_box_x2 = min(img_w, int(max_x))
            render_box_y2 = min(img_h, int(max_y))

            if render_box_x2 <= render_box_x1 or render_box_y2 <= render_box_y1:
                logger.warning(f"렌더링 스킵: 크기가 0인 렌더 박스 for '{translated_text[:20]}...'")
                return image_pil_original
            
            bbox_width_orig = max_x - min_x # 원본 OCR 박스 너비 (기울기 판단용)
            bbox_height_orig = max_y - min_y # 원본 OCR 박스 높이 (기울기 판단용)

            bbox_width_render = render_box_x2 - render_box_x1 # 실제 렌더링 대상 박스 너비
            bbox_height_render = render_box_y2 - render_box_y1 # 실제 렌더링 대상 박스 높이

        except Exception as e_box_calc:
            logger.error(f"렌더링 바운딩 박스 계산 오류: {e_box_calc}. Box: {box}. 원본 이미지 반환.", exc_info=True)
            return image_pil_original

        # 배경 채우기
        try:
            text_roi_pil = image_pil_original.crop((render_box_x1, render_box_y1, render_box_x2, render_box_y2))
            estimated_bg_color = get_quantized_dominant_color(text_roi_pil) if text_roi_pil.width > 0 and text_roi_pil.height > 0 else (200,200,200)
        except Exception as e_bg:
            logger.warning(f"배경색 추정 실패 ({e_bg}), 기본 회색 사용.", exc_info=True)
            estimated_bg_color = (200, 200, 200) # 밝은 회색으로 변경 (어두운 글씨와 대비)
        
        draw.rectangle([render_box_x1, render_box_y1, render_box_x2, render_box_y2], fill=estimated_bg_color)
        text_color = get_contrasting_text_color(estimated_bg_color)

        # 렌더링 영역 및 패딩 (패딩 최소화)
        padding_x = max(1, int(bbox_width_render * 0.03)) # 너비의 3%, 최소 1px
        padding_y = max(1, int(bbox_height_render * 0.03)) # 높이의 3%, 최소 1px
        
        render_area_x_start = render_box_x1 + padding_x
        render_area_y_start = render_box_y1 + padding_y
        render_area_width = bbox_width_render - 2 * padding_x
        render_area_height = bbox_height_render - 2 * padding_y

        if render_area_width <= 1 or render_area_height <= 1: # 렌더링 공간이 너무 작으면
            logger.warning(f"텍스트 '{translated_text[:20]}...' 렌더링 영역 너무 작음 (패딩 후). 스킵.")
            return img_to_draw_on # 배경은 채워진 상태로 반환될 수 있음

        # 기울어진 텍스트 판단 및 글꼴 크기 보정
        # 방법 1: OCR 각도 정보 활용 (제공된다면)
        # 방법 2: 바운딩 박스 비율 (ocr_angle이 None일 경우 사용)
        font_size_correction_factor = 1.0
        text_angle_deg = 0.0 # 수평으로 가정
        if ocr_angle is not None and isinstance(ocr_angle, (int, float)):
            text_angle_deg = abs(ocr_angle) # 절대값 사용
            if 5 < text_angle_deg < 85 or 95 < text_angle_deg < 175: # 유의미한 기울기 (수평/수직에서 5도 이상 벗어남)
                # 각도에 따라 보정 계수 동적 조절 (예시: 45도에 가까울수록 더 많이 줄임)
                # sin(각도) 또는 경험적 값 사용 가능
                font_size_correction_factor = max(0.6, 1.0 - (text_angle_deg / 90.0) * 0.3) # 각도가 클수록 조금 더 줄임 (최대 30% 감소)
                logger.debug(f"OCR 제공 각도 {ocr_angle:.1f}도. 글꼴 크기 보정 계수: {font_size_correction_factor:.2f}")
        elif bbox_width_orig > 0 and bbox_height_orig > 0 : # 각도 정보 없을 때 비율로 판단
            aspect_ratio_orig = bbox_width_orig / bbox_height_orig
            if aspect_ratio_orig > 2.0 or aspect_ratio_orig < 0.5: # 가로/세로 비율이 2배 이상 차이나면
                font_size_correction_factor = 0.80 # 20% 줄임
                logger.debug(f"바운딩 박스 비율 ({aspect_ratio_orig:.2f}) 기반 기울기 의심. 글꼴 크기 보정 계수: {font_size_correction_factor:.2f}")

        # 초기 목표 폰트 크기 (렌더링 영역 높이 기준, 보정 적용)
        # 렌더링 영역이 매우 좁고 높다면, 너비도 고려해야 함
        target_font_size = int(min(render_area_height * 0.9, render_area_width * 0.9 / (len(translated_text.splitlines()[0] if translated_text else 1)*0.5 +1) ) * font_size_correction_factor)
        target_font_size = max(target_font_size, 1) # 최소 1 이상
        
        min_font_size = 5 # 렌더링 가능한 최소 폰트 크기
        if target_font_size < min_font_size: target_font_size = min_font_size
        
        is_bold_font = '_bold' in font_code_for_render or 'bold' in font_code_for_render.lower()
        font = self._get_font(target_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
        
        final_font_size = target_font_size
        wrapped_lines = []
        line_spacing_render = int(target_font_size * 0.2) # 줄 간격을 폰트 크기의 20%로 설정

        # 폰트 크기 및 줄 바꿈 최적화 루프
        while final_font_size >= min_font_size:
            current_font = self._get_font(final_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            
            # textwrap을 사용하여 줄 바꿈 시도
            # textwrap.fill 보다 draw.multiline_textbbox 로 직접 계산하는 것이 더 정확할 수 있음
            # 여기서는 우선 textwrap으로 라인 수를 예측하고, 실제 렌더링 크기를 확인
            
            # 대략적인 한 줄당 문자 수 계산 (이전 방식보다 개선 필요)
            # draw.textlength 사용 (Pillow 9.2.0+)
            estimated_chars_per_line = 1
            if render_area_width > 0:
                try:
                    if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'textlength'):
                        # "W" 문자를 사용하여 대략적인 최대 너비 문자 기준 계산
                        char_w = draw.textlength("W", font=current_font)
                        if char_w > 0:
                            estimated_chars_per_line = max(1, int(render_area_width / char_w))
                    else: # 구버전 fallback
                        # getsize는 Pillow 10에서 제거됨. getlength는 단일 라인만.
                        if hasattr(current_font, 'getlength'):
                            char_w = current_font.getlength("W")
                        else: # 최후의 추정
                            char_w = final_font_size * 0.7 # 매우 거친 추정
                        if char_w > 0 :
                            estimated_chars_per_line = max(1, int(render_area_width / char_w))
                except Exception as e_char_w:
                    logger.warning(f"한 줄당 문자 수 추정 오류: {e_char_w}. 기본값 사용.")
                    estimated_chars_per_line = max(1, int(render_area_width / (final_font_size * 0.6)))


            wrapper = textwrap.TextWrapper(width=estimated_chars_per_line, break_long_words=True, 
                                           replace_whitespace=False, drop_whitespace=False,
                                           break_on_hyphens=True)
            temp_wrapped_lines = wrapper.wrap(translated_text)
            if not temp_wrapped_lines: temp_wrapped_lines = [" "] # 빈 텍스트면 공백 한 칸

            # 실제 렌더링될 높이 계산 (multiline_textbbox 사용)
            if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'multiline_textbbox'):
                try:
                    # multiline_textbbox는 (x1, y1, x2, y2) 반환
                    # anchor='la' (left, top) 기준으로 계산 (0,0 에서 시작한다고 가정)
                    text_bbox = draw.multiline_textbbox((0,0), "\n".join(temp_wrapped_lines), font=current_font, spacing=line_spacing_render)
                    rendered_text_height = text_bbox[3] - text_bbox[1]
                    rendered_text_width = text_bbox[2] - text_bbox[0]
                except Exception as e_mtb:
                    logger.warning(f"multiline_textbbox 계산 오류: {e_mtb}. 글꼴 크기 조정에 영향 미칠 수 있음.", exc_info=True)
                    # fallback: 각 줄의 높이를 더하는 방식 (덜 정확)
                    rendered_text_height = 0
                    max_w_line = 0
                    for line_idx_calc, line_calc in enumerate(temp_wrapped_lines):
                         # getbbox는 (left, top, right, bottom)을 반환하며, 기준점(0,0)에 대한 상대 좌표임
                        line_bbox_calc = current_font.getmask(line_calc).getbbox() if hasattr(current_font.getmask(line_calc),'getbbox') else (0,0,0,final_font_size) # fallback
                        rendered_text_height += (line_bbox_calc[3] - line_bbox_calc[1]) if line_bbox_calc else final_font_size
                        current_line_width = (line_bbox_calc[2] - line_bbox_calc[0]) if line_bbox_calc else current_font.getlength(line_calc) if hasattr(current_font, 'getlength') else len(line_calc) * final_font_size * 0.6
                        if current_line_width > max_w_line: max_w_line = current_line_width
                        if line_idx_calc < len(temp_wrapped_lines) - 1:
                            rendered_text_height += line_spacing_render
                    rendered_text_width = max_w_line

            else: # 구버전 Pillow: 각 줄 높이 합산 (getsize 사용 시 주의)
                rendered_text_height = 0
                max_w_line = 0
                for line_idx_calc, line_calc in enumerate(temp_wrapped_lines):
                    if hasattr(current_font, 'getsize'): # Pillow < 10
                        line_w, line_h = current_font.getsize(line_calc)
                        rendered_text_height += line_h
                        if line_w > max_w_line: max_w_line = line_w
                    else: # Pillow 10+ and no textbbox, getmask().getbbox() 사용
                        line_bbox_calc = current_font.getmask(line_calc).getbbox() if hasattr(current_font.getmask(line_calc),'getbbox') else (0,0,0,final_font_size)
                        rendered_text_height += (line_bbox_calc[3] - line_bbox_calc[1]) if line_bbox_calc else final_font_size
                        current_line_width = (line_bbox_calc[2] - line_bbox_calc[0]) if line_bbox_calc else len(line_calc) * final_font_size * 0.6
                        if current_line_width > max_w_line: max_w_line = current_line_width
                    if line_idx_calc < len(temp_wrapped_lines) - 1:
                        rendered_text_height += line_spacing_render
                rendered_text_width = max_w_line


            if rendered_text_height <= render_area_height and rendered_text_width <= render_area_width:
                wrapped_lines = temp_wrapped_lines # 이 설정 사용
                logger.debug(f"최종 글꼴 크기: {final_font_size}pt, 줄 수: {len(wrapped_lines)}, 계산된 높이: {rendered_text_height}, 너비: {rendered_text_width}")
                break # 적절한 크기 찾음
            
            final_font_size -= 1 # 글꼴 크기 1씩 줄여가며 테스트
            line_spacing_render = int(final_font_size * 0.2) # 줄 간격도 업데이트
        else: # 루프가 다 돌았는데도 못 맞춘 경우 (min_font_size에서도 넘침)
            # min_font_size로 설정하고, wrapped_lines는 마지막 시도된 값 사용
            font = self._get_font(min_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            wrapped_lines = temp_wrapped_lines # 마지막 시도 값
            final_font_size = min_font_size
            line_spacing_render = int(final_font_size * 0.2)
            logger.warning(f"텍스트 '{translated_text[:30]}...'가 영역에 맞지 않아 최소 글꼴 크기 {min_font_size}pt로 설정됨. 잘릴 수 있음.")

        # 텍스트 렌더링 (중앙 정렬 시도)
        # 최종 렌더링될 텍스트의 전체 높이와 너비를 다시 계산
        text_to_render_final = "\n".join(wrapped_lines)
        if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'multiline_textbbox'):
            try:
                final_bbox = draw.multiline_textbbox((0,0), text_to_render_final, font=font, spacing=line_spacing_render, anchor="la") # 좌상단 기준
                final_text_width = final_bbox[2] - final_bbox[0]
                final_text_height = final_bbox[3] - final_bbox[1]
            except Exception: # 만약 오류 발생 시
                final_text_width = render_area_width # 최대 너비로 가정
                final_text_height = render_area_height # 최대 높이로 가정 (정확도 떨어짐)
        else: # 구버전 Pillow 또는 textbbox 실패 시 (위의 루프에서 계산된 rendered_text_width/height 재활용)
            final_text_width = rendered_text_width
            final_text_height = rendered_text_height


        # 텍스트를 렌더링 영역의 중앙에 배치하기 위한 시작점 계산
        text_x_start = render_area_x_start + (render_area_width - final_text_width) / 2
        text_y_start = render_area_y_start + (render_area_height - final_text_height) / 2
        
        # 시작점이 음수가 되지 않도록 보정 (박스 좌상단 기준)
        text_x_start = max(render_area_x_start, text_x_start)
        text_y_start = max(render_area_y_start, text_y_start)

        try:
            if hasattr(draw, 'multiline_text'): # Pillow 8.0.0+
                # anchor 기본값은 'la' (left, top of baseline)
                # 중앙 정렬을 위해 xy는 블록의 좌상단, align은 'center'
                draw.multiline_text((text_x_start, text_y_start), 
                                   text_to_render_final, 
                                   font=font, 
                                   fill=text_color, 
                                   spacing=line_spacing_render, 
                                   align="left") # align="center"는 각 줄 내부 정렬. 블록 전체 정렬은 xy로.
                logger.debug(f"텍스트 렌더링: x={text_x_start:.1f}, y={text_y_start:.1f}, font_size={final_font_size}pt, align=left(block centered)")
            else: # 구버전 (multiline_text 없거나 spacing 지원 안 할 수 있음)
                 # 구버전은 align 인자도 없을 수 있음. 한 줄씩 직접 그려야 할 수도.
                 current_y = text_y_start
                 for line in wrapped_lines:
                     # 각 줄의 너비를 계산하여 x 시작점 조정 (만약 각 줄별 중앙 정렬 원하면)
                     # line_width, _ = font.getsize(line) if hasattr(font, 'getsize') else (len(line) * final_font_size * 0.6, final_font_size)
                     # current_x = text_x_start + (final_text_width - line_width) / 2 # 각 줄 중앙 정렬
                     current_x = text_x_start # 블록 좌측 정렬
                     draw.text((current_x, current_y), line, font=font, fill=text_color)
                     line_height_approx = final_font_size + line_spacing_render # 대략적인 줄 높이
                     current_y += line_height_approx
                 logger.debug(f"텍스트 렌더링 (구버전 방식): x={text_x_start:.1f}, y_start={text_y_start:.1f}, font_size={final_font_size}pt")

        except Exception as e_draw:
            logger.error(f"텍스트 렌더링 중 오류: {e_draw}", exc_info=True)
        
        return img_to_draw_on
# --- BaseOcrHandler 끝 ---


# --- PaddleOcrHandler 클래스 ---
class PaddleOcrHandler(BaseOcrHandler):
    def __init__(self, lang_code='korean', debug_enabled=False, use_gpu=False):
        # use_angle_cls는 False로 유지 (사용자 요청)
        self.use_angle_cls_paddle = False
        super().__init__(lang_codes=lang_code, debug_enabled=debug_enabled, use_gpu=use_gpu)

    def _initialize_engine(self):
        try:
            from paddleocr import PaddleOCR
            logger.info(f"PaddleOCR 초기화 시도 (lang: {self.current_lang_codes}, use_angle_cls: {self.use_angle_cls_paddle}, use_gpu: {self.use_gpu}, debug: {self.debug_mode})...")
            self.ocr_engine = PaddleOCR(use_angle_cls=self.use_angle_cls_paddle, lang=self.current_lang_codes, use_gpu=self.use_gpu, show_log=self.debug_mode)
            logger.info(f"PaddleOCR 초기화 완료 (lang: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("PaddleOCR 라이브러리를 찾을 수 없습니다. 'pip install paddleocr paddlepaddle'로 설치해주세요.")
            raise RuntimeError("PaddleOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e:
            logger.error(f"PaddleOCR 초기화 중 오류 (lang: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"PaddleOCR 초기화 실패 (lang: {self.current_lang_codes}): {e}")

    def _preprocess_image_for_ocr(self, image_pil_rgb):
        # (기존 전처리 로직 유지 또는 개선 가능 - 여기서는 동일하게 유지)
        image_cv = cv2.cvtColor(np.array(image_pil_rgb), cv2.COLOR_RGB2BGR)
        # ... (기존 코드와 동일한 전처리 과정) ...
        # 간단한 전처리로 변경 (테스트용)
        gray_img = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
        if self.debug_mode and gray_img is not None and gray_img.size > 0:
            try:
                debug_img_path = os.path.join(BASE_DIR_OCR, "paddle_preprocessed_debug.png")
                cv2.imwrite(debug_img_path, gray_img)
                logger.debug(f"PaddleOCR 전처리 디버그 이미지 저장: {debug_img_path}")
            except Exception as e_dbg_save:
                logger.warning(f"PaddleOCR 디버그 이미지 저장 실패: {e_dbg_save}")
        return gray_img # 또는 이진화된 이미지 등

    def ocr_image(self, image_pil_rgb):
        if not self.ocr_engine: return []
        try:
            preprocessed_cv_img = self._preprocess_image_for_ocr(image_pil_rgb) # NumPy 배열 반환 가정
            # PaddleOCR의 ocr 메서드는 이미지 경로 또는 NumPy 배열을 받음
            ocr_output = self.ocr_engine.ocr(preprocessed_cv_img, cls=self.use_angle_cls_paddle) # cls=False가 기본값 (방향 분류기 미사용)
            
            final_parsed_results = []
            if ocr_output and isinstance(ocr_output, list) and len(ocr_output) > 0:
                # PaddleOCR 결과는 [[box, (text, confidence)], ...] 또는 [[[box, (text, confidence)], ...]] 형태일 수 있음
                results_list = ocr_output
                if isinstance(ocr_output[0], list) and \
                   (len(ocr_output[0]) == 0 or (len(ocr_output[0]) > 0 and isinstance(ocr_output[0][0], list))):
                    # 이 경우 ocr_output은 [ [ result_line_1 ], [ result_line_2 ], ... ] 형태일 수 있고,
                    # 각 result_line_i는 [box, (text, conf)] 또는 [box, (text, conf), angle_info_if_any] 일 수 있음.
                    # 보통은 [[box1, (text1, conf1)], [box2, (text2, conf2)]] 형태의 단일 리스트를 반환.
                    # 만약 결과가 이중 리스트 [[...]] 형태라면, 첫 번째 요소 사용
                     results_list = ocr_output[0]


                for item in results_list:
                    if isinstance(item, list) and len(item) >= 2: # 최소 box와 (text, conf) 포함
                        box_data = item[0]
                        text_conf_tuple = item[1]
                        ocr_angle = None # PaddleOCR에서 각도 정보 추출 방법 확인 필요 (cls=True일 때 주로 나옴)

                        if isinstance(box_data, list) and len(box_data) == 4 and \
                           all(isinstance(point, list) and len(point) == 2 for point in box_data) and \
                           isinstance(text_conf_tuple, tuple) and len(text_conf_tuple) == 2:
                            # box_data를 int로 변환
                            box_points_int = [[int(coord[0]), int(coord[1])] for coord in box_data]
                            final_parsed_results.append([box_points_int, text_conf_tuple, ocr_angle])
                        else:
                            logger.warning(f"PaddleOCR 결과 항목 형식이 다릅니다 (내부): {item}")
                    else:
                        logger.warning(f"PaddleOCR 결과 항목이 리스트가 아니거나 길이가 2 미만입니다 (외부): {item}")
            return final_parsed_results
        except Exception as e:
            logger.error(f"PaddleOCR ocr_image 중 오류: {e}", exc_info=True)
            return []
# --- PaddleOcrHandler 끝 ---


# --- EasyOcrHandler 클래스 ---
class EasyOcrHandler(BaseOcrHandler):
    def __init__(self, lang_codes_list=['en'], debug_enabled=False, use_gpu=False):
        super().__init__(lang_codes=lang_codes_list, debug_enabled=debug_enabled, use_gpu=use_gpu)

    def _initialize_engine(self):
        try:
            import easyocr
            logger.info(f"EasyOCR 초기화 시도 (langs: {self.current_lang_codes}, gpu: {self.use_gpu}, verbose: {self.debug_mode})...")
            # EasyOCR Reader 초기화 시 rotation_info 같은 파라미터가 있는지 확인 필요 (문서 참조)
            # 기본적으로 EasyOCR은 다양한 방향의 텍스트를 잘 감지하는 편임.
            self.ocr_engine = easyocr.Reader(self.current_lang_codes, gpu=self.use_gpu, verbose=self.debug_mode)
            logger.info(f"EasyOCR 초기화 완료 (langs: {self.current_lang_codes}).")
        except ImportError:
            logger.critical("EasyOCR 라이브러리를 찾을 수 없습니다. 'pip install easyocr'로 설치해주세요.")
            raise RuntimeError("EasyOCR 라이브러리가 설치되어 있지 않습니다.")
        except Exception as e:
            logger.error(f"EasyOCR 초기화 중 오류 (langs: {self.current_lang_codes}): {e}", exc_info=True)
            raise RuntimeError(f"EasyOCR 초기화 실패 (langs: {self.current_lang_codes}): {e}")

    def ocr_image(self, image_pil_rgb):
        if not self.ocr_engine: return []
        try:
            image_np = np.array(image_pil_rgb.convert('RGB'))
            
            # EasyOCR의 readtext는 paragraph=False로 설정하면 각 텍스트 블록별로 반환
            # rotation_info 와 같은 파라미터로 각도 정보를 받을 수 있는지 확인 필요
            # 문서에는 x_ths, y_ths, low_text 등의 파라미터가 있으나 직접적인 각도 반환은 명시 안됨
            # 반환값: [[bbox, text, confidence], ...]
            # bbox는 [[x1,y1],[x2,y1],[x2,y2],[x1,y2]] 형태의 꼭짓점 리스트
            ocr_output = self.ocr_engine.readtext(image_np, detail=1, paragraph=False) 
            
            formatted_results = []
            for item_tuple in ocr_output: # item_tuple이 (bbox, text, confidence) 형태
                if not (isinstance(item_tuple, (list, tuple)) and len(item_tuple) >= 2):
                    logger.warning(f"EasyOCR 결과 항목 형식이 이상합니다: {item_tuple}")
                    continue

                bbox, text = item_tuple[0], item_tuple[1]
                confidence = item_tuple[2] if len(item_tuple) > 2 else 0.9 # 기본 신뢰도 (없을 경우)
                ocr_angle = None # EasyOCR에서 직접적인 각도 정보 반환 여부 확인 필요

                if isinstance(bbox, list) and len(bbox) == 4 and \
                   all(isinstance(p, (list, np.ndarray)) and len(p) == 2 for p in bbox): # np.ndarray도 허용
                    box_points = [[int(coord[0]), int(coord[1])] for coord in bbox]
                    formatted_results.append([box_points, (text, float(confidence)), ocr_angle])
                elif isinstance(bbox, np.ndarray) and bbox.shape == (4,2): # NumPy 배열 형태 [[x1,y1],...,[x4,y4]]
                    box_points = bbox.astype(int).tolist()
                    formatted_results.append([box_points, (text, float(confidence)), ocr_angle])
                else:
                     logger.warning(f"EasyOCR 결과의 bbox 형식이 예상과 다릅니다: {bbox}")
            return formatted_results
        except Exception as e:
            logger.error(f"EasyOCR ocr_image 중 오류: {e}", exc_info=True)
            return []
# --- EasyOcrHandler 끝 ---