from PIL import Image, ImageDraw, ImageFont, ImageStat
import numpy as np
import cv2
import os
import platform
import logging
import io
import textwrap

logger = logging.getLogger(__name__)

BASE_DIR_OCR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR_OCR, "fonts")
LANGUAGE_FONT_MAP = {
    'korean': 'NotoSansCJK-Regular.ttc', 'japan': 'NotoSansCJK-Regular.ttc',
    'ch': 'NotoSansCJK-Regular.ttc', 'chinese_cht': 'NotoSansCJK-Regular.ttc',
    'en': 'NotoSansCJK-Regular.ttc', 'th': 'NotoSansThai-VariableFont_wdth,wght.ttf',
    'es': 'NotoSansCJK-Regular.ttc',
}
DEFAULT_FONT_FILENAME = 'NotoSansCJK-Regular.ttc'
DEFAULT_BOLD_FONT_FILENAME = 'NotoSansCJK-Bold.ttc'

# --- 색상 관련 상수 ---
MIN_COLOR_DIFFERENCE = 40 # 배경색과 글자색 간의 최소 차이 (더 크면 대비가 뚜렷)
BG_COLOR_SAMPLE_OFFSET = 5 # 텍스트 바운딩 박스 바깥에서 배경색 샘플링할 오프셋

class PaddleOcrHandler:
    def __init__(self, lang='korean', debug_enabled=False):
        self.current_lang = lang
        self.debug_mode = debug_enabled
        self.ocr = None
        try:
            from paddleocr import PaddleOCR
            logger.info(f"PaddleOCR 초기화 시도 (lang: {self.current_lang}, debug: {self.debug_mode}, use_angle_cls: False)...")
            self.ocr = PaddleOCR(use_angle_cls=False, lang=self.current_lang, use_gpu=False, show_log=False)
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
        if w > 0 and (w < target_width / 2 or w > target_width * 2):
            scale_ratio = target_width / w; new_height = int(h * scale_ratio)
            if new_height > 0 and target_width > 0:
                processed_img = cv2.resize(processed_img, (target_width, new_height), interpolation=cv2.INTER_LANCZOS4)
        try:
            gray_img = cv2.cvtColor(processed_img, cv2.COLOR_BGR2GRAY)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8)); clahe_img = clahe.apply(gray_img)
            _, otsu_thresh_img = cv2.threshold(clahe_img, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            processed_img_for_ocr = otsu_thresh_img; logger.debug("이미지 전처리 완료.")
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

    def _is_image_problematic(self, image_pil_instance):
        if image_pil_instance is None: return True
        if hasattr(image_pil_instance, 'format') and image_pil_instance.format:
            fmt_lower = image_pil_instance.format.lower()
            if 'wmf' in fmt_lower or 'emf' in fmt_lower:
                logger.warning(f"문제 가능성 있는 이미지 형식 감지: {image_pil_instance.format}")
                return True
        return False

    def has_text_in_image_bytes(self, image_bytes):
        if not self.ocr: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if self._is_image_problematic(img_pil): return False # WMF/EMF 등 사전 회피
            if img_pil.width < 20 or img_pil.height < 20: return False
            img_pil_rgb = img_pil.convert("RGB")
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False
            result = self.ocr_image(img_pil_rgb)
            return bool(result)
        except OSError as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            if "WMF" in str(e).upper() or "EMF" in str(e).upper() or "cannot find loader" in str(e):
                logger.warning(f"지원되지 않는 이미지 형식(WMF/EMF 등)으로 텍스트 확인 건너뜀 ({format_info}): {e}")
            else: logger.error(f"이미지 텍스트 확인 중 Pillow OSError ({format_info}): {e}", exc_info=False)
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
                # PaddleOCR 결과 형식 유연하게 처리
                if isinstance(ocr_output[0], list) and all(isinstance(item, list) and len(item) == 2 and isinstance(item[1], tuple) for item in ocr_output[0]) :
                    actual_results = ocr_output[0]
                elif all(isinstance(item, list) and len(item) == 2 and isinstance(item[1], tuple) for item in ocr_output):
                    actual_results = ocr_output
            return actual_results
        except Exception as e:
            logger.error(f"OCR 처리 중 심각한 오류: {e}", exc_info=True)
            return []

    def _get_font(self, font_size, lang_code='en', is_bold=False):
        font_filename = None; font_path = None
        if is_bold: font_filename = LANGUAGE_FONT_MAP.get(lang_code + '_bold', DEFAULT_BOLD_FONT_FILENAME)
        else: font_filename = LANGUAGE_FONT_MAP.get(lang_code, DEFAULT_FONT_FILENAME)
        if font_filename: font_path = os.path.join(FONT_DIR, font_filename)
        if font_path and os.path.exists(font_path):
            try: return ImageFont.truetype(font_path, font_size)
            except IOError as e: logger.warning(f"폰트 로드 실패 ({font_path}): {e}. Pillow 기본 사용.")
        else: logger.warning(f"폰트 파일({font_path}) 없음. Pillow 기본 사용.")
        try: return ImageFont.truetype("arial.ttf", font_size) # Arial 시도
        except IOError: return ImageFont.load_default() # 최후의 수단

    def _get_dominant_color(self, image_pil_crop, num_clusters=3):
        """이미지 영역에서 주요 색상을 K-Means 클러스터링으로 찾습니다 (더 정확하지만 느릴 수 있음)"""
        try:
            img_array = np.array(image_pil_crop.convert("RGB"))
            pixels = img_array.reshape((-1, 3))
            pixels = np.float32(pixels)
            criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 200, .1)
            flags = cv2.KMEANS_RANDOM_CENTERS
            _, labels, palette = cv2.kmeans(pixels, num_clusters, None, criteria, 10, flags)
            _, counts = np.unique(labels, return_counts=True)
            dominant = tuple(palette[np.argmax(counts)].astype(int))
            return dominant
        except Exception as e_kmeans:
            logger.warning(f"K-Means로 주요 색상 감지 실패: {e_kmeans}. 평균색 사용.")
            # K-Means 실패 시 평균색으로 대체
            stat = ImageStat.Stat(image_pil_crop)
            return tuple(map(int, stat.mean[:3]))


    def _get_avg_color_around_box(self, image_pil_original, box):
        """텍스트 박스 주변의 평균 배경색을 추정합니다."""
        try:
            img_w, img_h = image_pil_original.size
            x_coords = [p[0] for p in box]; y_coords = [p[1] for p in box]
            min_x, max_x = int(min(x_coords)), int(max(x_coords))
            min_y, max_y = int(min(y_coords)), int(max(y_coords))

            # 박스 바깥쪽 영역 샘플링 (더 넓은 영역에서)
            sample_regions = []
            # Top
            sample_regions.append((max(0, min_x - BG_COLOR_SAMPLE_OFFSET*2), max(0, min_y - BG_COLOR_SAMPLE_OFFSET*2), 
                                   min(img_w, max_x + BG_COLOR_SAMPLE_OFFSET*2), max(0, min_y - BG_COLOR_SAMPLE_OFFSET)))
            # Bottom
            sample_regions.append((max(0, min_x - BG_COLOR_SAMPLE_OFFSET*2), min(img_h, max_y + BG_COLOR_SAMPLE_OFFSET), 
                                   min(img_w, max_x + BG_COLOR_SAMPLE_OFFSET*2), min(img_h, max_y + BG_COLOR_SAMPLE_OFFSET*2)))
            # Left
            sample_regions.append((max(0, min_x - BG_COLOR_SAMPLE_OFFSET*2), max(0, min_y - BG_COLOR_SAMPLE_OFFSET*2), 
                                   max(0, min_x - BG_COLOR_SAMPLE_OFFSET), min(img_h, max_y + BG_COLOR_SAMPLE_OFFSET*2)))
            # Right
            sample_regions.append((min(img_w, max_x + BG_COLOR_SAMPLE_OFFSET), max(0, min_y - BG_COLOR_SAMPLE_OFFSET*2), 
                                   min(img_w, max_x + BG_COLOR_SAMPLE_OFFSET*2), min(img_h, max_y + BG_COLOR_SAMPLE_OFFSET*2)))
            
            avg_colors = []
            for r_idx, (l, t, r, b) in enumerate(sample_regions):
                if r > l and b > t: # 유효한 영역일 때만
                    crop = image_pil_original.crop((l, t, r, b))
                    if crop.width > 0 and crop.height > 0:
                        stat = ImageStat.Stat(crop)
                        avg_colors.append(np.array(stat.mean[:3])) # RGB 평균
            
            if avg_colors:
                # 모든 주변 영역의 평균색들의 평균을 최종 배경색으로 (더 안정적일 수 있음)
                final_avg_color = tuple(np.mean(avg_colors, axis=0).astype(int))
                logger.debug(f"박스 주변 평균 배경색 감지: {final_avg_color}")
                return final_avg_color
        except Exception as e_avg_bg:
            logger.warning(f"주변 배경색 감지 중 오류: {e_avg_bg}")
        
        # 실패 시, 박스 자체의 평균색을 사용 (최후의 수단)
        logger.debug("주변 배경색 감지 실패, 박스 내부 평균색으로 대체.")
        try:
            box_crop = image_pil_original.crop((min_x, min_y, max_x, max_y))
            if box_crop.width > 0 and box_crop.height > 0:
                 return tuple(map(int, ImageStat.Stat(box_crop).mean[:3]))
        except: pass
        return (255, 255, 255) # 기본값 흰색

    def _determine_text_color(self, image_pil_original, box, bg_color):
        """OCR된 텍스트 박스 내에서 배경색과 대비되는 글자색을 추정합니다."""
        try:
            x_coords = [p[0] for p in box]; y_coords = [p[1] for p in box]
            text_box_crop = image_pil_original.crop((int(min(x_coords)), int(min(y_coords)), 
                                                    int(max(x_coords)), int(max(y_coords))))
            if text_box_crop.width == 0 or text_box_crop.height == 0: return (0,0,0) # 기본 검정색

            # K-Means로 주요 색상 2개 (배경, 글자) 추출 시도
            dominant_colors = []
            try:
                img_array = np.array(text_box_crop.convert("RGB"))
                pixels = img_array.reshape((-1, 3)); pixels = np.float32(pixels)
                criteria = (cv2.TERM_CRITERIA_EPS + cv2.TERM_CRITERIA_MAX_ITER, 10, 1.0)
                _, labels, palette = cv2.kmeans(pixels, 2, None, criteria, 5, cv2.KMEANS_PP_CENTERS) # KMEANS_PP_CENTERS 사용
                dominant_colors = [tuple(p.astype(int)) for p in palette]
            except Exception as e_kmeans_text:
                logger.warning(f"텍스트 영역 K-Means 색상 추출 실패: {e_kmeans_text}. 이미지 전체 평균색으로 대체.")
                # 실패 시 이미지 전체 평균색 사용 (최후의 수단)
                stat = ImageStat.Stat(text_box_crop)
                avg_color = tuple(map(int, stat.mean[:3]))
                return (0,0,0) if sum(abs(c1-c2) for c1,c2 in zip(avg_color, bg_color)) < MIN_COLOR_DIFFERENCE else avg_color

            if len(dominant_colors) == 2:
                color1, color2 = dominant_colors[0], dominant_colors[1]
                diff1 = sum(abs(c1-c2) for c1,c2 in zip(color1, bg_color)) # color1과 배경색 차이
                diff2 = sum(abs(c1-c2) for c1,c2 in zip(color2, bg_color)) # color2와 배경색 차이

                # 배경색과 더 많이 다른 색을 글자색으로 선택
                text_color_candidate = color1 if diff1 > diff2 else color2
                # 글자색과 배경색의 차이가 충분한지 확인
                if sum(abs(c1-c2) for c1,c2 in zip(text_color_candidate, bg_color)) >= MIN_COLOR_DIFFERENCE:
                    logger.debug(f"감지된 글자색: {text_color_candidate} (배경: {bg_color})")
                    return text_color_candidate
            elif dominant_colors: # 색상이 하나만 감지된 경우 (단색 이미지 등)
                 # 배경색과 충분히 다르면 그 색을 글자색으로, 아니면 검정색
                single_color = dominant_colors[0]
                if sum(abs(c1-c2) for c1,c2 in zip(single_color, bg_color)) >= MIN_COLOR_DIFFERENCE:
                    logger.debug(f"단일 주요 색상 감지, 글자색으로 사용: {single_color} (배경: {bg_color})")
                    return single_color
        except Exception as e_text_color:
            logger.warning(f"글자색 감지 중 오류: {e_text_color}")
        
        logger.debug(f"글자색 감지 실패 또는 배경과 유사. 기본 검정색(0,0,0) 사용 (배경: {bg_color}).")
        return (0, 0, 0) # 기본 검정색

    def render_translated_text_on_image(self, image_pil_original, box, translated_text, font_code_for_render='en', original_text=""):
        img_to_draw_on = image_pil_original.copy()
        draw = ImageDraw.Draw(img_to_draw_on)

        # 1. 배경색 감지 (텍스트 박스 주변 영역)
        detected_bg_color = self._get_avg_color_around_box(image_pil_original, box)
        logger.debug(f"렌더링 시 사용할 배경색: {detected_bg_color}")

        # 2. 원본 텍스트 영역을 감지된 배경색으로 덮기
        try:
            poly_points = [(int(p[0]), int(p[1])) for p in box]
            draw.polygon(poly_points, fill=detected_bg_color)
        except Exception as e_poly:
            logger.warning(f"원본 텍스트 영역 배경색으로 덮기 실패 (Pillow): {e_poly}. OpenCV 시도.")
            try:
                img_cv_temp = self._pil_to_cv2(img_to_draw_on)
                pts_cv = np.array(box, dtype=np.int32)
                cv2.fillPoly(img_cv_temp, [pts_cv], detected_bg_color) # OpenCV는 BGR 순서, Pillow는 RGB. _pil_to_cv2에서 변환됨.
                img_to_draw_on = self._cv2_to_pil(img_cv_temp)
                draw = ImageDraw.Draw(img_to_draw_on)
            except Exception as e_cv_fill:
                logger.error(f"OpenCV로도 영역 덮기 실패: {e_cv_fill}. 흰색으로 대체.")
                draw.polygon(poly_points, fill=(255,255,255)) # 최후의 수단: 흰색

        # 3. 글자색 감지 (원본 텍스트 박스 내부, 감지된 배경색과의 대비 활용)
        #    원본 이미지와 원본 박스 정보를 사용해야 함.
        detected_text_color = self._determine_text_color(image_pil_original, box, detected_bg_color)
        logger.debug(f"렌더링 시 사용할 글자색: {detected_text_color}")

        # 4. 텍스트 렌더링 (폰트 크기 조절 등은 이전과 동일)
        x_coords = [p[0] for p in box]; y_coords = [p[1] for p in box]
        min_x, max_x = min(x_coords), max(x_coords); min_y, max_y = min(y_coords), max(y_coords)
        bbox_width = max_x - min_x; bbox_height = max_y - min_y
        if bbox_width <=0 or bbox_height <=0 : return image_pil_original # 변경 없음
        
        padding = int(min(bbox_width, bbox_height) * 0.05)
        render_area_x = min_x + padding; render_area_y = min_y + padding
        render_area_width = bbox_width - 2 * padding; render_area_height = bbox_height - 2 * padding
        if render_area_width <=0 or render_area_height <=0: return image_pil_original

        target_font_size = int(render_area_height * 0.8)
        if target_font_size < 8: target_font_size = 8
        wrapped_text = translated_text; line_spacing_render = 4

        while target_font_size >= 8:
            font = self._get_font(target_font_size, lang_code=font_code_for_render)
            # ... (이전 폰트 크기 조절 루프와 동일한 로직) ...
            avg_char_width=font.getsize("Ag")[0]/2 if hasattr(font,'getsize') else target_font_size/1.8
            if avg_char_width<=0: avg_char_width=1
            chars_per_line=int(render_area_width/avg_char_width);
            if chars_per_line<=0: chars_per_line=1
            wrapper=textwrap.TextWrapper(width=chars_per_line,break_long_words=True,replace_whitespace=False,drop_whitespace=False,expand_tabs=True,tabsize=4)
            wrapped_lines=wrapper.wrap(translated_text);
            if not wrapped_lines: wrapped_lines=[" "]
            wrapped_text="\n".join(wrapped_lines);total_h=0
            for line_idx,line_render_text in enumerate(wrapped_lines):
                line_bbox=font.getbbox(line_render_text) if hasattr(font,'getbbox') else (0,0,font.getsize(line_render_text)[0],font.getsize(line_render_text)[1]) if hasattr(font,'getsize') else (0,0,0,target_font_size)
                total_h+=(line_bbox[3]-line_bbox[1])
                if line_idx<len(wrapped_lines)-1:total_h+=line_spacing_render
            if total_h<=render_area_height:break
            target_font_size-=1
        else: logger.warning(f"텍스트 '{translated_text[:20]}...' 영역에 맞출 수 없음 (폰트 8).")
        
        # 감지된 글자색으로 텍스트 그리기
        logger.debug(f"이미지에 텍스트 렌더링: '{wrapped_text[:30].replace(chr(10),' ')}...' at ({render_area_x},{render_area_y}), font size {target_font_size}, color {detected_text_color}")
        try:
            draw.text((render_area_x, render_area_y), wrapped_text, font=font, fill=detected_text_color, spacing=line_spacing_render, align="left")
        except TypeError: # 이전 Pillow 호환성
            draw.text((render_area_x, render_area_y), wrapped_text, font=font, fill=detected_text_color)
        
        return img_to_draw_on