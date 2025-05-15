from PIL import Image, ImageDraw, ImageFont
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
        # (이전 안정화 시점의 전처리 로직)
        logger.debug("OCR을 위한 이미지 전처리 시작...")
        processed_img = image_cv.copy(); h, w = processed_img.shape[:2]; target_width = 1000
        if w > 0 and (w < target_width / 2 or w > target_width * 2):
            scale_ratio = target_width / w; new_height = int(h * scale_ratio)
            if new_height > 0 and target_width > 0:
                processed_img = cv2.resize(processed_img, (target_width, new_height), interpolation=cv2.INTER_LANCZOS4)
        try:
            gray_img = cv2.cvtColor(processed_img, cv2.COLOR_BGR2GRAY)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)); clahe_img = clahe.apply(gray_img)
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

    def has_text_in_image_bytes(self, image_bytes):
        if not self.ocr: return False
        img_pil = None
        try:
            img_pil = Image.open(io.BytesIO(image_bytes))
            if img_pil.width < 20 or img_pil.height < 20: return False
            img_pil_rgb = img_pil.convert("RGB") # OSError 발생 가능 지점
            if img_pil_rgb.width < 1 or img_pil_rgb.height < 1: return False
            result = self.ocr_image(img_pil_rgb)
            return bool(result)
        except OSError as e:
            format_info = f"Format: {img_pil.format if img_pil else 'N/A'}"
            # WMF/EMF 로더 오류 등을 여기서 일반 OSError로 처리
            logger.warning(f"이미지 텍스트 확인 중 Pillow OSError ({format_info}), 처리 건너뜀: {e}", exc_info=False)
            return False # 오류 시 텍스트 없는 것으로 간주하여 안정성 확보
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
                # 결과 형식 처리 (이전 안정화 버전과 동일하게)
                if isinstance(ocr_output[0], list) and all(isinstance(item, list) and len(item) == 2 and isinstance(item[1], tuple) for item in ocr_output[0]):
                    actual_results = ocr_output[0]
                elif all(isinstance(item, list) and len(item) == 2 and isinstance(item[1], tuple) for item in ocr_output):
                    actual_results = ocr_output # 이전 버전 호환
            return actual_results
        except Exception as e:
            logger.error(f"OCR 처리 중 심각한 오류: {e}", exc_info=True)
            return []

    def _get_font(self, font_size, lang_code='en', is_bold=False):
        # ... (이전과 동일) ...
        font_filename = None; font_path = None
        if is_bold: font_filename = LANGUAGE_FONT_MAP.get(lang_code + '_bold', DEFAULT_BOLD_FONT_FILENAME)
        else: font_filename = LANGUAGE_FONT_MAP.get(lang_code, DEFAULT_FONT_FILENAME)
        if font_filename: font_path = os.path.join(FONT_DIR, font_filename)
        if font_path and os.path.exists(font_path):
            try: return ImageFont.truetype(font_path, font_size)
            except IOError as e: logger.warning(f"폰트 로드 실패 ({font_path}): {e}. Pillow 기본 폰트로 대체.")
        else: logger.warning(f"폰트 파일({font_path}) 없음. Pillow 기본 폰트 사용.")
        return ImageFont.load_default()

    def render_translated_text_on_image(self, image_pil_original, box, translated_text, font_code_for_render='en', original_text=""):
        # ... (이전과 동일) ...
        img_to_draw_on = image_pil_original.copy(); draw = ImageDraw.Draw(img_to_draw_on)
        try: draw.polygon([(int(p[0]), int(p[1])) for p in box], fill=(255, 255, 255))
        except Exception as e_poly: logger.warning(f"영역 지우기 실패 (Pillow): {e_poly}. OpenCV 시도."); # OpenCV 대체 로직 필요 시 추가
        x_coords=[p[0] for p in box]; y_coords=[p[1] for p in box]; min_x,max_x=min(x_coords),max(x_coords); min_y,max_y=min(y_coords),max(y_coords)
        bbox_width=max_x-min_x; bbox_height=max_y-min_y
        if bbox_width<=0 or bbox_height<=0: return image_pil_original
        padding=int(min(bbox_width,bbox_height)*0.05); render_area_x=min_x+padding; render_area_y=min_y+padding
        render_area_width=bbox_width-2*padding; render_area_height=bbox_height-2*padding
        if render_area_width<=0 or render_area_height<=0: return image_pil_original
        target_font_size=int(render_area_height*0.8);
        if target_font_size<8: target_font_size=8
        wrapped_text=translated_text; line_spacing_render=4
        while target_font_size>=8:
            font=self._get_font(target_font_size,lang_code=font_code_for_render)
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
        try: draw.text((render_area_x,render_area_y),wrapped_text,font=font,fill=(0,0,0),spacing=line_spacing_render,align="left")
        except TypeError: draw.text((render_area_x,render_area_y),wrapped_text,font=font,fill=(0,0,0))
        return img_to_draw_on