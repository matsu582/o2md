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


class OCRProcessor:
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


class TextDetectorOCR:
    """テキスト検出とOCRを統合したクラス"""
    
    def __init__(self, model_path: Optional[str] = None):
        """
        Args:
            model_path: テキスト検出モデルのパス
        """
        self.detector = ComicTextDetector(model_path=model_path)
        self.ocr_processor = OCRProcessor()
    
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
            line_threshold = im_h * 0.02  # 画像高さの2%
            
            def get_sort_key(r):
                cy = r.center()[1]
                cx = r.center()[0]
                # 行番号を計算（Y座標を行閾値で丸める）
                line_num = int(cy / line_threshold)
                return (line_num, cx)
            
            regions.sort(key=get_sort_key)
        
        return regions
    
    def get_combined_text(self, regions: List[TextRegion], 
                          separator: str = "\n\n") -> str:
        """テキスト領域のテキストを結合
        
        Args:
            regions: テキスト領域のリスト
            separator: 区切り文字（デフォルトは空行で段落分割）
            
        Returns:
            結合されたテキスト
        """
        texts = [r.text for r in regions if r.text]
        return separator.join(texts)


def process_pdf_page_with_detection(page_img: np.ndarray, 
                                    model_path: Optional[str] = None) -> str:
    """PDFページ画像からテキストを検出してOCRを実行
    
    Args:
        page_img: ページ画像 (BGR形式)
        model_path: テキスト検出モデルのパス
        
    Returns:
        抽出されたテキスト
    """
    processor = TextDetectorOCR(model_path=model_path)
    regions = processor.process_image(page_img)
    
    if len(regions) == 0:
        # テキスト領域が検出されなかった場合は画像全体でOCR
        debug_print("[DEBUG] テキスト領域が検出されなかったため、画像全体でOCRを実行")
        return processor.ocr_processor.ocr_full_image(page_img)
    
    return processor.get_combined_text(regions)
