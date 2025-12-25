"""
pdf2md_ocr.py - PDFのOCR処理モジュール

テキスト検出とOCR機能を提供します。
comic-text-detectorを使用してテキスト領域を検出し、
manga-ocrで各領域のテキストを抽出します。
"""

import os
import sys
import time
import urllib.request
from typing import List, Tuple, Optional
from pathlib import Path

import numpy as np
import cv2
from PIL import Image

# モデルのダウンロードURL
MODEL_URL = "https://github.com/zyddnys/manga-image-translator/releases/download/beta-0.3/comictextdetector.pt.onnx"
MODEL_FILENAME = "comictextdetector.pt.onnx"
MODEL_EXPECTED_SIZE = 99000000  # 約94.6MB、最小サイズとして99MBを期待

# 最大検出領域数（これを超える場合はフォールバック）
MAX_REGIONS = 100

# グローバル変数
_VERBOSE = False


def set_verbose(verbose: bool):
    """詳細出力モードを設定"""
    global _VERBOSE
    _VERBOSE = verbose


def debug_print(msg: str):
    """デバッグメッセージを出力"""
    if _VERBOSE:
        print(msg)


class TextRegion:
    """検出されたテキスト領域を表すクラス"""
    
    def __init__(self, bbox: Tuple[int, int, int, int], 
                 confidence: float = 1.0,
                 is_vertical: bool = False):
        """
        Args:
            bbox: バウンディングボックス (x1, y1, x2, y2)
            confidence: 検出信頼度
            is_vertical: 縦書きかどうか
        """
        self.bbox = bbox
        self.confidence = confidence
        self.is_vertical = is_vertical
        self.text = ""
    
    @property
    def x1(self) -> int:
        return self.bbox[0]
    
    @property
    def y1(self) -> int:
        return self.bbox[1]
    
    @property
    def x2(self) -> int:
        return self.bbox[2]
    
    @property
    def y2(self) -> int:
        return self.bbox[3]
    
    @property
    def width(self) -> int:
        return self.x2 - self.x1
    
    @property
    def height(self) -> int:
        return self.y2 - self.y1
    
    @property
    def area(self) -> int:
        return self.width * self.height
    
    def center(self) -> Tuple[float, float]:
        """領域の中心座標を返す"""
        return ((self.x1 + self.x2) / 2, (self.y1 + self.y2) / 2)


class ComicTextDetector:
    """comic-text-detectorを使用したテキスト検出器"""
    
    def __init__(self, model_path: Optional[str] = None, 
                 input_size: int = 1024,
                 conf_thresh: float = 0.4,
                 nms_thresh: float = 0.35):
        """
        Args:
            model_path: ONNXモデルファイルのパス
            input_size: 入力画像サイズ
            conf_thresh: 信頼度閾値
            nms_thresh: NMS閾値
        """
        self.input_size = input_size
        self.conf_thresh = conf_thresh
        self.nms_thresh = nms_thresh
        self.model = None
        self.model_path = model_path
        
        # デフォルトのモデルパスを設定
        if self.model_path is None:
            default_path = Path(__file__).parent / "models" / MODEL_FILENAME
            self.model_path = str(default_path)
    
    def _download_model(self) -> bool:
        """モデルファイルをダウンロード
        
        Returns:
            ダウンロード成功時True
        """
        model_dir = Path(self.model_path).parent
        model_dir.mkdir(parents=True, exist_ok=True)
        
        debug_print(f"[INFO] テキスト検出モデルをダウンロード中: {MODEL_URL}")
        print(f"[INFO] テキスト検出モデルをダウンロード中...")
        
        try:
            urllib.request.urlretrieve(MODEL_URL, self.model_path)
            debug_print(f"[INFO] モデルのダウンロード完了: {self.model_path}")
            print(f"[INFO] モデルのダウンロード完了")
            return True
        except Exception as e:
            debug_print(f"[WARNING] モデルのダウンロードに失敗: {e}")
            print(f"[WARNING] モデルのダウンロードに失敗: {e}")
            return False
    
    def _load_model(self):
        """モデルを遅延読み込み（必要に応じてダウンロード）"""
        if self.model is not None:
            return True
        
        # モデルファイルが存在しない場合はダウンロード
        if not os.path.exists(self.model_path):
            if not self._download_model():
                return False
        
        try:
            self.model = cv2.dnn.readNetFromONNX(self.model_path)
            debug_print(f"[INFO] テキスト検出モデルを読み込みました: {self.model_path}")
            return True
        except Exception as e:
            debug_print(f"[WARNING] モデル読み込みエラー: {e}")
            return False
    
    def _preprocess_image(self, img: np.ndarray) -> Tuple[np.ndarray, float, int, int]:
        """画像を前処理してモデル入力形式に変換
        
        Args:
            img: 入力画像 (BGR形式)
            
        Returns:
            前処理済み画像, リサイズ比率, パディング幅, パディング高さ
        """
        im_h, im_w = img.shape[:2]
        target_size = self.input_size
        
        # アスペクト比を維持してリサイズ
        scale = min(target_size / im_w, target_size / im_h)
        new_w = int(im_w * scale)
        new_h = int(im_h * scale)
        
        resized = cv2.resize(img, (new_w, new_h), interpolation=cv2.INTER_LINEAR)
        
        # パディングを追加
        dw = target_size - new_w
        dh = target_size - new_h
        top = dh // 2
        bottom = dh - top
        left = dw // 2
        right = dw - left
        
        padded = cv2.copyMakeBorder(resized, top, bottom, left, right,
                                     cv2.BORDER_CONSTANT, value=(0, 0, 0))
        
        # BGR -> RGB, HWC -> CHW, 正規化
        blob = cv2.dnn.blobFromImage(padded, 1/255.0, (target_size, target_size),
                                      swapRB=True, crop=False)
        
        return blob, scale, dw // 2, dh // 2
    
    def _apply_nms(self, boxes: np.ndarray, scores: np.ndarray) -> List[int]:
        """Non-Maximum Suppressionを適用
        
        Args:
            boxes: バウンディングボックス配列 [N, 4]
            scores: スコア配列 [N]
            
        Returns:
            保持するインデックスのリスト
        """
        if len(boxes) == 0:
            return []
        
        x1 = boxes[:, 0]
        y1 = boxes[:, 1]
        x2 = boxes[:, 2]
        y2 = boxes[:, 3]
        
        areas = (x2 - x1) * (y2 - y1)
        order = scores.argsort()[::-1]
        
        keep = []
        while order.size > 0:
            i = order[0]
            keep.append(i)
            
            if order.size == 1:
                break
            
            # IoU計算
            xx1 = np.maximum(x1[i], x1[order[1:]])
            yy1 = np.maximum(y1[i], y1[order[1:]])
            xx2 = np.minimum(x2[i], x2[order[1:]])
            yy2 = np.minimum(y2[i], y2[order[1:]])
            
            w = np.maximum(0, xx2 - xx1)
            h = np.maximum(0, yy2 - yy1)
            inter = w * h
            
            iou = inter / (areas[i] + areas[order[1:]] - inter)
            
            inds = np.where(iou <= self.nms_thresh)[0]
            order = order[inds + 1]
        
        return keep
    
    def detect(self, img: np.ndarray) -> List[TextRegion]:
        """画像からテキスト領域を検出
        
        Args:
            img: 入力画像 (BGR形式またはRGB形式)
            
        Returns:
            検出されたテキスト領域のリスト
        """
        if not self._load_model():
            return []
        
        im_h, im_w = img.shape[:2]
        
        # 前処理
        blob, scale, pad_w, pad_h = self._preprocess_image(img)
        
        # 推論実行
        self.model.setInput(blob)
        try:
            output_names = self.model.getUnconnectedOutLayersNames()
            outputs = self.model.forward(output_names)
        except Exception as e:
            debug_print(f"[WARNING] 推論エラー: {e}")
            return []
        
        # 出力を解析
        # comic-text-detectorの出力形式:
        # outputs[0] (blk): [1, N, 7] - テキストブロック検出 (x_center, y_center, w, h, conf, class1, class2)
        # outputs[1] (det): [1, 2, H, W] - 検出マップ
        # outputs[2] (seg): [1, 1, H, W] - セグメンテーションマスク
        if len(outputs) == 0:
            return []
        
        # blk出力を取得（テキストブロック検出結果）
        blk_output = outputs[0]
        if len(blk_output.shape) == 3:
            blk_output = blk_output[0]  # (N, 7)
        
        regions = []
        boxes = []
        scores = []
        
        # リサイズ比率を計算（元画像サイズへの変換用）
        resize_ratio_w = im_w / (self.input_size - pad_w * 2)
        resize_ratio_h = im_h / (self.input_size - pad_h * 2)
        
        for det in blk_output:
            if len(det) < 5:
                continue
            
            conf = det[4]
            if conf < self.conf_thresh:
                continue
            
            # YOLOv5形式: x_center, y_center, w, h -> x1, y1, x2, y2
            x_center, y_center, w, h = det[0], det[1], det[2], det[3]
            
            # パディングを考慮して座標を変換
            x_center_adj = x_center - pad_w
            y_center_adj = y_center - pad_h
            
            # 元の画像サイズにスケール
            x_center_orig = x_center_adj * resize_ratio_w
            y_center_orig = y_center_adj * resize_ratio_h
            w_orig = w * resize_ratio_w
            h_orig = h * resize_ratio_h
            
            # x1, y1, x2, y2に変換
            x1 = x_center_orig - w_orig / 2
            y1 = y_center_orig - h_orig / 2
            x2 = x_center_orig + w_orig / 2
            y2 = y_center_orig + h_orig / 2
            
            # 画像範囲内にクリップ
            x1 = max(0, min(im_w, x1))
            y1 = max(0, min(im_h, y1))
            x2 = max(0, min(im_w, x2))
            y2 = max(0, min(im_h, y2))
            
            if x2 > x1 and y2 > y1:
                boxes.append([x1, y1, x2, y2])
                scores.append(conf)
        
        if len(boxes) == 0:
            return []
        
        boxes = np.array(boxes)
        scores = np.array(scores)
        
        # NMS適用
        keep = self._apply_nms(boxes, scores)
        
        for idx in keep:
            bbox = tuple(int(v) for v in boxes[idx])
            # 縦書き判定（高さ > 幅 * 1.5）
            is_vertical = (bbox[3] - bbox[1]) > (bbox[2] - bbox[0]) * 1.5
            regions.append(TextRegion(bbox, scores[idx], is_vertical))
        
        return regions


class BaseOCRProcessor:
    """OCR処理器の基底クラス"""
    
    def ocr_region(self, img: np.ndarray, region: TextRegion, 
                   padding_ratio: float = 0.1,
                   min_size: int = 32) -> str:
        """テキスト領域からOCRでテキストを抽出（サブクラスで実装）"""
        raise NotImplementedError
    
    def ocr_full_image(self, img: np.ndarray) -> str:
        """画像全体からOCRでテキストを抽出（サブクラスで実装）"""
        raise NotImplementedError


class OCRProcessor(BaseOCRProcessor):
    """manga-ocrを使用したOCR処理器"""
    
    def __init__(self):
        self._ocr = None
    
    def _get_ocr(self):
        """manga-ocrインスタンスを遅延初期化"""
        if self._ocr is None:
            try:
                # tokenizersのスレッドプール生成を抑止（終了時ハング対策）
                os.environ.setdefault("TOKENIZERS_PARALLELISM", "false")
                
                from manga_ocr import MangaOcr
                self._ocr = MangaOcr()
                debug_print("[INFO] manga-ocrを初期化しました")
            except ImportError:
                debug_print("[WARNING] manga-ocrがインストールされていません")
                self._ocr = False
            except Exception as e:
                debug_print(f"[WARNING] manga-ocr初期化エラー: {e}")
                self._ocr = False
        
        return self._ocr if self._ocr else None
    
    def ocr_region(self, img: np.ndarray, region: TextRegion, 
                   padding_ratio: float = 0.1,
                   min_size: int = 32) -> str:
        """テキスト領域からOCRでテキストを抽出
        
        Args:
            img: 入力画像 (BGR形式)
            region: テキスト領域
            padding_ratio: パディング比率
            min_size: 最小サイズ（これより小さい場合は拡大）
            
        Returns:
            抽出されたテキスト
        """
        ocr = self._get_ocr()
        if ocr is None:
            return ""
        
        im_h, im_w = img.shape[:2]
        
        # パディングを追加してクロップ
        pad_x = int(region.width * padding_ratio)
        pad_y = int(region.height * padding_ratio)
        
        x1 = max(0, region.x1 - pad_x)
        y1 = max(0, region.y1 - pad_y)
        x2 = min(im_w, region.x2 + pad_x)
        y2 = min(im_h, region.y2 + pad_y)
        
        cropped = img[y1:y2, x1:x2]
        
        if cropped.size == 0:
            return ""
        
        # 小さすぎる場合は拡大
        crop_h, crop_w = cropped.shape[:2]
        if crop_h < min_size or crop_w < min_size:
            scale = max(min_size / crop_h, min_size / crop_w, 2.0)
            new_w = int(crop_w * scale)
            new_h = int(crop_h * scale)
            cropped = cv2.resize(cropped, (new_w, new_h), 
                                interpolation=cv2.INTER_LANCZOS4)
        
        # BGR -> RGB -> PIL Image
        rgb = cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        
        try:
            text = ocr(pil_img)
            return text.strip() if text else ""
        except Exception as e:
            debug_print(f"[WARNING] OCRエラー: {e}")
            return ""
    
    def ocr_full_image(self, img: np.ndarray) -> str:
        """画像全体からOCRでテキストを抽出
        
        Args:
            img: 入力画像 (BGR形式)
            
        Returns:
            抽出されたテキスト
        """
        ocr = self._get_ocr()
        if ocr is None:
            return ""
        
        # BGR -> RGB -> PIL Image
        rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        
        try:
            text = ocr(pil_img)
            return text.strip() if text else ""
        except Exception as e:
            debug_print(f"[WARNING] OCRエラー: {e}")
            return ""


class TesseractOCRProcessor(BaseOCRProcessor):
    """Tesseractを使用したOCR処理器"""
    
    def __init__(self, lang: str = "jpn+eng", tessdata_dir: Optional[str] = None):
        """
        Args:
            lang: Tesseractの言語設定（デフォルト: jpn+eng）
            tessdata_dir: tessdataディレクトリのパス（tessdata_best使用時に指定）
        """
        self._tesseract = None
        self._lang = lang
        self._tessdata_dir = tessdata_dir
    
    def _get_tesseract(self):
        """pytesseractモジュールを遅延インポート"""
        if self._tesseract is None:
            try:
                import pytesseract
                self._tesseract = pytesseract
                debug_print("[INFO] Tesseractを初期化しました")
            except ImportError:
                debug_print("[WARNING] pytesseractがインストールされていません")
                self._tesseract = False
            except Exception as e:
                debug_print(f"[WARNING] Tesseract初期化エラー: {e}")
                self._tesseract = False
        
        return self._tesseract if self._tesseract else None
    
    def ocr_region(self, img: np.ndarray, region: TextRegion, 
                   padding_ratio: float = 0.1,
                   min_size: int = 32) -> str:
        """テキスト領域からOCRでテキストを抽出
        
        Args:
            img: 入力画像 (BGR形式)
            region: テキスト領域
            padding_ratio: パディング比率
            min_size: 最小サイズ（これより小さい場合は拡大）
            
        Returns:
            抽出されたテキスト
        """
        tesseract = self._get_tesseract()
        if tesseract is None:
            return ""
        
        im_h, im_w = img.shape[:2]
        
        # パディングを追加してクロップ
        pad_x = int(region.width * padding_ratio)
        pad_y = int(region.height * padding_ratio)
        
        x1 = max(0, region.x1 - pad_x)
        y1 = max(0, region.y1 - pad_y)
        x2 = min(im_w, region.x2 + pad_x)
        y2 = min(im_h, region.y2 + pad_y)
        
        cropped = img[y1:y2, x1:x2]
        
        if cropped.size == 0:
            return ""
        
        # BGR -> RGB -> PIL Image
        rgb = cv2.cvtColor(cropped, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        
        try:
            # tessdata_dirが指定されている場合はconfigに追加
            config = ""
            if self._tessdata_dir:
                config = f"--tessdata-dir {self._tessdata_dir}"
            text = tesseract.image_to_string(pil_img, lang=self._lang, config=config)
            return text.strip() if text else ""
        except Exception as e:
            debug_print(f"[WARNING] Tesseract OCRエラー: {e}")
            return ""
    
    def ocr_full_image(self, img: np.ndarray) -> str:
        """画像全体からOCRでテキストを抽出
        
        Args:
            img: 入力画像 (BGR形式)
            
        Returns:
            抽出されたテキスト
        """
        tesseract = self._get_tesseract()
        if tesseract is None:
            return ""
        
        # BGR -> RGB -> PIL Image
        rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(rgb)
        
        try:
            # tessdata_dirが指定されている場合はconfigに追加
            config = ""
            if self._tessdata_dir:
                config = f"--tessdata-dir {self._tessdata_dir}"
            text = tesseract.image_to_string(pil_img, lang=self._lang, config=config)
            return text.strip() if text else ""
        except Exception as e:
            debug_print(f"[WARNING] Tesseract OCRエラー: {e}")
            return ""


# OCRエンジンの種類
OCR_ENGINE_MANGA = "manga-ocr"
OCR_ENGINE_TESSERACT = "tesseract"
OCR_ENGINES = [OCR_ENGINE_MANGA, OCR_ENGINE_TESSERACT]


def create_ocr_processor(engine: str = OCR_ENGINE_TESSERACT, 
                         lang: str = "jpn+eng",
                         tessdata_dir: Optional[str] = None) -> BaseOCRProcessor:
    """OCRエンジンに応じたOCR処理器を作成
    
    Args:
        engine: OCRエンジン名（"manga-ocr" または "tesseract"）
        lang: Tesseractの言語設定（デフォルト: jpn+eng）
        tessdata_dir: tessdataディレクトリのパス（tessdata_best使用時に指定）
        
    Returns:
        OCR処理器インスタンス
    """
    if engine == OCR_ENGINE_TESSERACT:
        return TesseractOCRProcessor(lang=lang, tessdata_dir=tessdata_dir)
    else:
        return OCRProcessor()


class TextDetectorOCR:
    """テキスト検出とOCRを統合したクラス"""
    
    def __init__(self, model_path: Optional[str] = None,
                 ocr_engine: str = OCR_ENGINE_TESSERACT,
                 ocr_lang: str = "jpn+eng",
                 tessdata_dir: Optional[str] = None):
        """
        Args:
            model_path: テキスト検出モデルのパス
            ocr_engine: OCRエンジン名（"manga-ocr" または "tesseract"）
            ocr_lang: Tesseractの言語設定（デフォルト: jpn+eng）
            tessdata_dir: tessdataディレクトリのパス（tessdata_best使用時に指定）
        """
        self.detector = ComicTextDetector(model_path=model_path)
        self.ocr_processor = create_ocr_processor(
            engine=ocr_engine, lang=ocr_lang, tessdata_dir=tessdata_dir
        )
        self.ocr_engine = ocr_engine
    
    def _split_multiline_region(self, img: np.ndarray, region: TextRegion,
                                 min_gap_height: int = 5,
                                 min_region_height: int = 100,
                                 min_line_height: int = 20) -> List[TextRegion]:
        """複数行を含む領域を水平方向の空白で分割
        
        Args:
            img: 入力画像 (BGR形式)
            region: 分割対象の領域
            min_gap_height: 最小ギャップ高さ（ピクセル）
            min_region_height: 分割対象とする最小領域高さ
            min_line_height: 分割後の最小行高さ
            
        Returns:
            分割された領域のリスト
        """
        # 縦書き領域は分割しない（縦書きは別のロジックが必要）
        if region.is_vertical:
            return [region]
        
        # 高さが小さい領域は分割しない
        if region.height < min_region_height:
            return [region]
        
        # 領域をクロップ
        im_h, im_w = img.shape[:2]
        x1 = max(0, region.x1)
        y1 = max(0, region.y1)
        x2 = min(im_w, region.x2)
        y2 = min(im_h, region.y2)
        
        cropped = img[y1:y2, x1:x2]
        if cropped.size == 0:
            return [region]
        
        # グレースケールに変換して2値化
        gray = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
        
        # y方向の投影（各行の黒画素数）
        projection = np.sum(binary, axis=1)
        
        # 空白行を検出（投影値が閾値以下）
        max_proj = np.max(projection)
        if max_proj == 0:
            return [region]
        threshold = max_proj * 0.05
        is_gap = projection < threshold
        
        # 連続する空白行のグループを検出
        gaps = []
        gap_start = None
        for i, is_g in enumerate(is_gap):
            if is_g and gap_start is None:
                gap_start = i
            elif not is_g and gap_start is not None:
                if i - gap_start >= min_gap_height:
                    gaps.append((gap_start, i))
                gap_start = None
        
        # ギャップがなければ分割しない
        if len(gaps) == 0:
            return [region]
        
        # ギャップで分割
        sub_regions = []
        prev_end = 0
        for gap_start, gap_end in gaps:
            if gap_start > prev_end:
                sub_y1 = y1 + prev_end
                sub_y2 = y1 + gap_start
                if sub_y2 - sub_y1 >= min_line_height:
                    sub_regions.append(TextRegion(
                        bbox=(x1, sub_y1, x2, sub_y2),
                        confidence=region.confidence,
                        is_vertical=region.is_vertical
                    ))
            prev_end = gap_end
        
        # 最後の領域
        if prev_end < region.height:
            sub_y1 = y1 + prev_end
            sub_y2 = y2
            if sub_y2 - sub_y1 >= min_line_height:
                sub_regions.append(TextRegion(
                    bbox=(x1, sub_y1, x2, sub_y2),
                    confidence=region.confidence,
                    is_vertical=region.is_vertical
                ))
        
        if len(sub_regions) > 1:
            debug_print(f"[DEBUG] 領域を{len(sub_regions)}行に分割")
            return sub_regions
        
        return [region]
    
    def process_image(self, img: np.ndarray, 
                      sort_regions: bool = True) -> List[TextRegion]:
        """画像からテキストを検出してOCRを実行
        
        Args:
            img: 入力画像 (BGR形式)
            sort_regions: 領域を読み順にソートするか
            
        Returns:
            OCR結果を含むテキスト領域のリスト
        """
        # テキスト領域を検出
        regions = self.detector.detect(img)
        
        if len(regions) == 0:
            debug_print("[DEBUG] テキスト領域が検出されませんでした")
            return []
        
        debug_print(f"[DEBUG] {len(regions)}個のテキスト領域を検出")
        
        # 複数行を含む領域を行ごとに分割
        split_regions = []
        for region in regions:
            split_regions.extend(self._split_multiline_region(img, region))
        
        if len(split_regions) != len(regions):
            debug_print(f"[DEBUG] 分割後: {len(split_regions)}個の領域")
        
        regions = split_regions
        
        # 読み順にソート（上から下、左から右）
        if sort_regions:
            regions = self._sort_regions(regions, img.shape[1], img.shape[0])
        
        # 各領域でOCRを実行
        for region in regions:
            text = self.ocr_processor.ocr_region(img, region)
            region.text = text
            if text:
                debug_print(f"[DEBUG] OCR結果: {text[:50]}...")
        
        return regions
    
    def _sort_regions(self, regions: List[TextRegion], 
                      im_w: int, im_h: int) -> List[TextRegion]:
        """テキスト領域を読み順にソート
        
        横書き文書の場合: 上から下、左から右
        縦書き文書の場合: 右から左、上から下
        
        Args:
            regions: テキスト領域のリスト
            im_w: 画像幅
            im_h: 画像高さ
            
        Returns:
            ソートされたテキスト領域のリスト
        """
        if len(regions) == 0:
            return regions
        
        # 縦書きが多いかどうかを判定
        vertical_count = sum(1 for r in regions if r.is_vertical)
        is_vertical_doc = vertical_count > len(regions) / 2
        
        if is_vertical_doc:
            # 縦書き: 右から左、上から下
            regions.sort(key=lambda r: (-r.center()[0], r.center()[1]))
        else:
            # 横書き: 上から下、左から右
            # 行をグループ化（Y座標が近いものを同じ行とみなす）
            # 閾値は領域の中央値高さの半分を使用
            heights = [r.height for r in regions]
            median_height = sorted(heights)[len(heights) // 2] if heights else 50
            line_threshold = median_height * 0.5
            
            def get_sort_key(r):
                cy = r.center()[1]
                cx = r.center()[0]
                # 行番号を計算（Y座標を行閾値で丸める）
                line_num = int(cy / line_threshold)
                return (line_num, cx)
            
            regions.sort(key=get_sort_key)
        
        return regions
    
    def get_combined_text(self, regions: List[TextRegion], 
                          im_h: int = 0,
                          line_separator: str = "  ",
                          paragraph_separator: str = "\n\n") -> str:
        """テキスト領域のテキストを結合（同じ行は横に並べる）
        
        Args:
            regions: テキスト領域のリスト
            im_h: 画像高さ（行グループ化の閾値計算用）
            line_separator: 同じ行内の区切り文字
            paragraph_separator: 行間の区切り文字
            
        Returns:
            結合されたテキスト
        """
        if not regions:
            return ""
        
        # テキストがある領域のみ
        regions_with_text = [r for r in regions if r.text]
        if not regions_with_text:
            return ""
        
        # 画像高さが指定されていない場合は単純結合
        if im_h == 0:
            texts = [r.text for r in regions_with_text]
            return paragraph_separator.join(texts)
        
        # 行をグループ化（Y座標が近いものを同じ行とみなす）
        # 閾値は領域の中央値高さの半分を使用
        heights = [r.height for r in regions_with_text]
        median_height = sorted(heights)[len(heights) // 2] if heights else 50
        line_threshold = median_height * 0.5
        
        # 領域を行ごとにグループ化
        lines = []
        current_line = []
        current_line_y = None
        
        for region in regions_with_text:
            cy = region.center()[1]
            
            if current_line_y is None:
                current_line_y = cy
                current_line.append(region)
            elif abs(cy - current_line_y) < line_threshold:
                current_line.append(region)
            else:
                # 新しい行
                if current_line:
                    lines.append(current_line)
                current_line = [region]
                current_line_y = cy
        
        # 最後の行を追加
        if current_line:
            lines.append(current_line)
        
        # 各行内をX座標でソートして結合
        result_lines = []
        for line in lines:
            line.sort(key=lambda r: r.center()[0])
            line_texts = [r.text for r in line if r.text]
            result_lines.append(line_separator.join(line_texts))
        
        return paragraph_separator.join(result_lines)


def process_pdf_page_with_detection(page_img: np.ndarray, 
                                    model_path: Optional[str] = None,
                                    ocr_engine: str = OCR_ENGINE_TESSERACT,
                                    ocr_lang: str = "jpn+eng",
                                    tessdata_dir: Optional[str] = None) -> str:
    """PDFページ画像からテキストを検出してOCRを実行
    
    Args:
        page_img: ページ画像 (BGR形式)
        model_path: テキスト検出モデルのパス
        ocr_engine: OCRエンジン名（"manga-ocr" または "tesseract"）
        ocr_lang: Tesseractの言語設定（デフォルト: jpn+eng）
        tessdata_dir: tessdataディレクトリのパス（tessdata_best使用時に指定）
        
    Returns:
        抽出されたテキスト
    """
    processor = TextDetectorOCR(model_path=model_path, 
                                ocr_engine=ocr_engine,
                                ocr_lang=ocr_lang,
                                tessdata_dir=tessdata_dir)
    regions = processor.process_image(page_img)
    
    if len(regions) == 0:
        # テキスト領域が検出されなかった場合は画像全体でOCR
        debug_print("[DEBUG] テキスト領域が検出されなかったため、画像全体でOCRを実行")
        return processor.ocr_processor.ocr_full_image(page_img)
    
    # 画像高さを渡して行グループ化を有効にする
    im_h = page_img.shape[0]
    return processor.get_combined_text(regions, im_h=im_h)
