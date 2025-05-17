from PIL import Image, ImageDraw, ImageFont, ImageStat, __version__ as PILLOW_VERSION
import numpy as np
import cv2
import os
import logging
import io
import textwrap
import math

# 설정 파일 import
import config

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
        raise NotImplementedError("각 OCR 핸들러는 이 메서드를 구현해야 합니다.")

    def has_text_in_image_bytes(self, image_bytes):
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
    
    def _get_font(self, font_size, lang_code='en', is_bold=False):
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
            logger.warning(f"요청된 폰트 코드 '{lang_code}'(bold:{is_bold})에 대한 폰트 매핑 없음. 기본 폰트 '{font_filename}' 사용.")

        if font_filename:
            font_path = os.path.join(FONT_DIR, font_filename) # config.FONTS_DIR 사용

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
            if PILLOW_VERSION_TUPLE >= (9, 0, 0):
                 return ImageFont.load_default()
            else:
                 return ImageFont.load_default(size=int(font_size))
        except TypeError:
            return ImageFont.load_default()
        except Exception as e_default_font:
            logger.critical(f"Pillow 기본 폰트 로드조차 실패: {e_default_font}. 글꼴 렌더링 불가.", exc_info=True)
            raise RuntimeError(f"기본 폰트 로드 실패: {e_default_font}")

    def render_translated_text_on_image(self, image_pil_original, box, translated_text,
                                        font_code_for_render='en', original_text="", ocr_angle=None):
        # (이하 렌더링 로직은 기존과 동일하게 유지)
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
            logger.warning(f"텍스트 '{translated_text[:20]}...' 렌더링 영역 너무 작음 (패딩 후). 스킵.")
            return img_to_draw_on

        font_size_correction_factor = 1.0
        text_angle_deg = 0.0
        if ocr_angle is not None and isinstance(ocr_angle, (int, float)):
            text_angle_deg = abs(ocr_angle)
            if 5 < text_angle_deg < 85 or 95 < text_angle_deg < 175:
                font_size_correction_factor = max(0.6, 1.0 - (text_angle_deg / 90.0) * 0.3)
                logger.debug(f"OCR 제공 각도 {ocr_angle:.1f}도. 글꼴 크기 보정 계수: {font_size_correction_factor:.2f}")
        elif bbox_width_orig > 0 and bbox_height_orig > 0 :
            aspect_ratio_orig = bbox_width_orig / bbox_height_orig
            if aspect_ratio_orig > 2.0 or aspect_ratio_orig < 0.5:
                font_size_correction_factor = 0.80
                logger.debug(f"바운딩 박스 비율 ({aspect_ratio_orig:.2f}) 기반 기울기 의심. 글꼴 크기 보정 계수: {font_size_correction_factor:.2f}")

        target_font_size = int(min(render_area_height * 0.9, render_area_width * 0.9 / (len(translated_text.splitlines()[0] if translated_text else 1)*0.5 +1) ) * font_size_correction_factor)
        target_font_size = max(target_font_size, 1)
        
        min_font_size = 5
        if target_font_size < min_font_size: target_font_size = min_font_size
        
        is_bold_font = '_bold' in font_code_for_render or 'bold' in font_code_for_render.lower()
        font = self._get_font(target_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
        
        final_font_size = target_font_size
        wrapped_lines = []
        line_spacing_render = int(target_font_size * 0.2)

        while final_font_size >= min_font_size:
            current_font = self._get_font(final_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            
            estimated_chars_per_line = 1
            if render_area_width > 0:
                try:
                    if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'textlength'):
                        char_w = draw.textlength("W", font=current_font)
                        if char_w > 0:
                            estimated_chars_per_line = max(1, int(render_area_width / char_w))
                    else: 
                        if hasattr(current_font, 'getlength'):
                            char_w = current_font.getlength("W")
                        else: 
                            char_w = final_font_size * 0.7
                        if char_w > 0 :
                            estimated_chars_per_line = max(1, int(render_area_width / char_w))
                except Exception as e_char_w:
                    logger.warning(f"한 줄당 문자 수 추정 오류: {e_char_w}. 기본값 사용.")
                    estimated_chars_per_line = max(1, int(render_area_width / (final_font_size * 0.6)))

            wrapper = textwrap.TextWrapper(width=estimated_chars_per_line, break_long_words=True, 
                                           replace_whitespace=False, drop_whitespace=False,
                                           break_on_hyphens=True)
            temp_wrapped_lines = wrapper.wrap(translated_text)
            if not temp_wrapped_lines: temp_wrapped_lines = [" "]

            rendered_text_height = 0
            rendered_text_width = 0
            if PILLOW_VERSION_TUPLE >= (9, 2, 0) and hasattr(draw, 'multiline_textbbox'):
                try:
                    text_bbox = draw.multiline_textbbox((0,0), "\n".join(temp_wrapped_lines), font=current_font, spacing=line_spacing_render)
                    rendered_text_height = text_bbox[3] - text_bbox[1]
                    rendered_text_width = text_bbox[2] - text_bbox[0]
                except Exception as e_mtb:
                    logger.warning(f"multiline_textbbox 계산 오류: {e_mtb}. 글꼴 크기 조정에 영향 미칠 수 있음.", exc_info=True)
                    # Fallback or error handling
                    max_w_line = 0
                    for line_idx_calc, line_calc in enumerate(temp_wrapped_lines):
                        line_bbox_calc = current_font.getmask(line_calc).getbbox() if hasattr(current_font.getmask(line_calc),'getbbox') else (0,0,0,final_font_size)
                        rendered_text_height += (line_bbox_calc[3] - line_bbox_calc[1]) if line_bbox_calc else final_font_size
                        current_line_width = (line_bbox_calc[2] - line_bbox_calc[0]) if line_bbox_calc else (len(line_calc) * final_font_size * 0.6 if hasattr(current_font, 'getlength') else len(line_calc) * final_font_size * 0.6)
                        if current_line_width > max_w_line: max_w_line = current_line_width
                        if line_idx_calc < len(temp_wrapped_lines) - 1:
                            rendered_text_height += line_spacing_render
                    rendered_text_width = max_w_line

            else:
                max_w_line = 0
                for line_idx_calc, line_calc in enumerate(temp_wrapped_lines):
                    if hasattr(current_font, 'getsize'):
                        line_w, line_h = current_font.getsize(line_calc)
                        rendered_text_height += line_h
                        if line_w > max_w_line: max_w_line = line_w
                    else:
                        line_bbox_calc = current_font.getmask(line_calc).getbbox() if hasattr(current_font.getmask(line_calc),'getbbox') else (0,0,0,final_font_size)
                        rendered_text_height += (line_bbox_calc[3] - line_bbox_calc[1]) if line_bbox_calc else final_font_size
                        current_line_width = (line_bbox_calc[2] - line_bbox_calc[0]) if line_bbox_calc else len(line_calc) * final_font_size * 0.6
                        if current_line_width > max_w_line: max_w_line = current_line_width
                    if line_idx_calc < len(temp_wrapped_lines) - 1:
                        rendered_text_height += line_spacing_render
                rendered_text_width = max_w_line

            if rendered_text_height <= render_area_height and rendered_text_width <= render_area_width:
                wrapped_lines = temp_wrapped_lines
                logger.debug(f"최종 글꼴 크기: {final_font_size}pt, 줄 수: {len(wrapped_lines)}, 계산된 높이: {rendered_text_height}, 너비: {rendered_text_width}")
                break
            
            final_font_size -= 1
            line_spacing_render = int(final_font_size * 0.2)
        else:
            font = self._get_font(min_font_size, lang_code=font_code_for_render, is_bold=is_bold_font)
            wrapped_lines = temp_wrapped_lines
            final_font_size = min_font_size
            line_spacing_render = int(final_font_size * 0.2)
            logger.warning(f"텍스트 '{translated_text[:30]}...'가 영역에 맞지 않아 최소 글꼴 크기 {min_font_size}pt로 설정됨. 잘릴 수 있음.")

        text_to_render_final = "\n".join(wrapped_lines)
        final_text_width = rendered_text_width # Use calculated width from loop
        final_text_height = rendered_text_height # Use calculated height from loop

        text_x_start = render_area_x_start + (render_area_width - final_text_width) / 2
        text_y_start = render_area_y_start + (render_area_height - final_text_height) / 2
        
        text_x_start = max(render_area_x_start, text_x_start)
        text_y_start = max(render_area_y_start, text_y_start)

        try:
            if hasattr(draw, 'multiline_text'):
                draw.multiline_text((text_x_start, text_y_start), 
                                   text_to_render_final, 
                                   font=font, 
                                   fill=text_color, 
                                   spacing=line_spacing_render, 
                                   align="left")
                logger.debug(f"텍스트 렌더링: x={text_x_start:.1f}, y={text_y_start:.1f}, font_size={final_font_size}pt, align=left(block centered)")
            else:
                 current_y = text_y_start
                 for line in wrapped_lines:
                     current_x = text_x_start
                     draw.text((current_x, current_y), line, font=font, fill=text_color)
                     line_height_approx = final_font_size + line_spacing_render
                     current_y += line_height_approx
                 logger.debug(f"텍스트 렌더링 (구버전 방식): x={text_x_start:.1f}, y_start={text_y_start:.1f}, font_size={final_font_size}pt")

        except Exception as e_draw:
            logger.error(f"텍스트 렌더링 중 오류: {e_draw}", exc_info=True)
        
        return img_to_draw_on

class PaddleOcrHandler(BaseOcrHandler):
    def __init__(self, lang_code='korean', debug_enabled=False, use_gpu=False):
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
        image_cv = cv2.cvtColor(np.array(image_pil_rgb), cv2.COLOR_RGB2BGR)
        gray_img = cv2.cvtColor(image_cv, cv2.COLOR_BGR2GRAY)
        if self.debug_mode and gray_img is not None and gray_img.size > 0:
            try:
                # 디버그 이미지 저장 경로는 BASE_DIR_OCR (ocr_handler.py 위치) 기준으로 생성
                debug_img_path = os.path.join(BASE_DIR_OCR, "paddle_preprocessed_debug.png")
                cv2.imwrite(debug_img_path, gray_img)
                logger.debug(f"PaddleOCR 전처리 디버그 이미지 저장: {debug_img_path}")
            except Exception as e_dbg_save:
                logger.warning(f"PaddleOCR 디버그 이미지 저장 실패: {e_dbg_save}")
        return gray_img

    def ocr_image(self, image_pil_rgb):
        if not self.ocr_engine: return []
        try:
            preprocessed_cv_img = self._preprocess_image_for_ocr(image_pil_rgb)
            ocr_output = self.ocr_engine.ocr(preprocessed_cv_img, cls=self.use_angle_cls_paddle)
            
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

class EasyOcrHandler(BaseOcrHandler):
    def __init__(self, lang_codes_list=['en'], debug_enabled=False, use_gpu=False):
        super().__init__(lang_codes=lang_codes_list, debug_enabled=debug_enabled, use_gpu=use_gpu)

    def _initialize_engine(self):
        try:
            import easyocr
            logger.info(f"EasyOCR 초기화 시도 (langs: {self.current_lang_codes}, gpu: {self.use_gpu}, verbose: {self.debug_mode})...")
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
            ocr_output = self.ocr_engine.readtext(image_np, detail=1, paragraph=False) 
            
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
                    box_points = [[int(coord[0]), int(coord[1])] for coord in bbox]
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