#!/usr/bin/env python3
"""
img2md.py - 画像ファイルからOCRでテキスト抽出しMarkdownに変換

対応形式:
- JPEG (.jpg, .jpeg)
- PNG (.png)
- GIF (.gif)
- BMP (.bmp)
- TIFF (.tiff, .tif)
- WebP (.webp)

OCRエンジン:
- tesseract (デフォルト)
- manga-ocr
- sarashina

pdf2md_ocrモジュールのOCRエンジンを再利用して画像からテキストを抽出する。
"""

import os
import sys
import shutil
import argparse
from pathlib import Path
from typing import Optional

import cv2
import numpy as np

# 対応画像拡張子
IMAGE_EXTENSIONS = (
    '.jpg', '.jpeg', '.png', '.gif',
    '.bmp', '.tiff', '.tif', '.webp',
)

# グローバル変数
_VERBOSE = False


def set_verbose(verbose: bool):
    """デバッグ出力の有効/無効を設定"""
    global _VERBOSE
    _VERBOSE = verbose


def is_verbose() -> bool:
    """verboseモードかどうかを返す"""
    return _VERBOSE


def debug_print(msg: str):
    """verboseモード時のみ出力"""
    if _VERBOSE:
        print(msg)


class ImageToMarkdownConverter:
    """画像ファイルをOCRでMarkdownに変換するコンバータ

    画像ファイルを読み込み、OCRエンジンでテキストを抽出し、
    画像リンクとOCRテキストを含むMarkdownファイルを出力する。
    """

    def __init__(
        self,
        file_path: str,
        output_dir: Optional[str] = None,
        ocr_engine: str = "tesseract",
        tessdata_dir: Optional[str] = None,
    ):
        """
        Args:
            file_path: 画像ファイルのパス
            output_dir: 出力ディレクトリ(省略時はカレントディレクトリ/output)
            ocr_engine: OCRエンジン名（"tesseract", "manga-ocr", "sarashina"）
            tessdata_dir: tessdataディレクトリのパス（tessdata_best使用時）
        """
        self.file_path = file_path
        self.base_name = Path(file_path).stem
        self.file_ext = Path(file_path).suffix.lower()
        self.ocr_engine = ocr_engine
        self.tessdata_dir = tessdata_dir

        if output_dir:
            self.output_dir = output_dir
        else:
            self.output_dir = os.path.join(os.getcwd(), "output")

        self.images_dir = os.path.join(self.output_dir, "images")
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.images_dir, exist_ok=True)

    def get_auto_generated_patterns(self) -> list:
        """このコンバータが自動付与する見出しの正規表現パターンを返す"""
        import re
        return [
            re.compile(r'^' + re.escape(self.base_name) + r'$'),
            re.compile(r'^抽出テキスト（OCR）$'),
        ]

    def convert(self) -> str:
        """変換メイン処理

        Returns:
            出力Markdownファイルのパス
        """
        print(f"[INFO] 画像OCR変換開始: {self.file_path}")

        # 画像を読み込み
        img = self._load_image()
        if img is None:
            raise ValueError(f"画像ファイルの読み込みに失敗しました: {self.file_path}")

        # 画像を出力ディレクトリにコピー
        image_filename = f"{self.base_name}{self.file_ext}"
        dest_image_path = os.path.join(self.images_dir, image_filename)
        shutil.copy2(self.file_path, dest_image_path)
        debug_print(f"[INFO] 画像をコピー: {dest_image_path}")

        # OCRでテキスト抽出
        ocr_text = self._extract_text_with_ocr(img)

        # Markdown生成
        md_lines = self._build_markdown(image_filename, ocr_text)

        # ファイル出力
        output_path = os.path.join(
            self.output_dir, f"{self.base_name}.md"
        )
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(md_lines))

        print(f"[SUCCESS] 変換完了: {output_path}")
        return output_path

    def _load_image(self) -> Optional[np.ndarray]:
        """画像ファイルを読み込む

        Returns:
            BGR形式のnumpy配列、失敗時はNone
        """
        try:
            # 日本語パス対応: np.fromfileでバイト列読み込み→cv2.imdecodeでデコード
            buf = np.fromfile(self.file_path, dtype=np.uint8)
            img = cv2.imdecode(buf, cv2.IMREAD_COLOR)
            if img is None:
                print(f"[ERROR] 画像のデコードに失敗: {self.file_path}")
                return None
            debug_print(
                f"[INFO] 画像読み込み完了: {img.shape[1]}x{img.shape[0]}"
            )
            return img
        except Exception as e:
            print(f"[ERROR] 画像読み込みエラー: {e}")
            return None

    def _extract_text_with_ocr(self, img: np.ndarray) -> str:
        """OCRエンジンでテキストを抽出する

        Args:
            img: BGR形式の画像

        Returns:
            抽出されたテキスト
        """
        try:
            from pdf2md_ocr import (
                process_pdf_page_with_detection,
                set_verbose as ocr_set_verbose,
            )
            ocr_set_verbose(is_verbose())

            text = process_pdf_page_with_detection(
                img,
                ocr_engine=self.ocr_engine,
                tessdata_dir=self.tessdata_dir,
            )
            if text:
                debug_print(
                    f"[INFO] OCRテキスト抽出完了: {len(text)}文字"
                )
            else:
                print("[WARNING] OCRでテキストが抽出されませんでした")
            return text.strip() if text else ""

        except ImportError as e:
            print(f"[WARNING] pdf2md_ocrモジュールが利用できません: {e}")
            # フォールバック: 直接tesseractを呼び出し
            return self._ocr_fallback(img)
        except Exception as e:
            print(f"[WARNING] OCR処理中にエラー: {e}")
            return self._ocr_fallback(img)

    def _ocr_fallback(self, img: np.ndarray) -> str:
        """フォールバックOCR処理

        pdf2md_ocrが利用できない場合、直接pytesseractを呼び出す。
        """
        try:
            import pytesseract
            from PIL import Image as PILImage

            rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            pil_img = PILImage.fromarray(rgb)
            config = ""
            if self.tessdata_dir:
                config = f"--tessdata-dir {self.tessdata_dir}"
            text = pytesseract.image_to_string(
                pil_img, lang="jpn+eng", config=config
            )
            return text.strip() if text else ""
        except ImportError:
            print("[ERROR] pytesseractがインストールされていません")
            return ""
        except Exception as e:
            print(f"[ERROR] フォールバックOCRエラー: {e}")
            return ""

    def _build_markdown(
        self, image_filename: str, ocr_text: str
    ) -> list[str]:
        """Markdownテキストを構築する

        Args:
            image_filename: 画像ファイル名
            ocr_text: OCR抽出テキスト

        Returns:
            Markdown行リスト
        """
        md = []
        md.append(f"# {self.base_name}")
        md.append("")
        md.append(f"![{self.base_name}](images/{image_filename})")
        md.append("")

        if ocr_text:
            md.append("### 抽出テキスト（OCR）")
            md.append("")
            md.append(ocr_text)
            md.append("")

        return md


def main():
    """コマンドラインエントリポイント"""
    parser = argparse.ArgumentParser(
        description='画像ファイルからOCRでテキスト抽出しMarkdownに変換'
    )
    parser.add_argument(
        'image_file',
        help='変換する画像ファイル (.jpg/.png/.gif/.bmp/.tiff/.webp)'
    )
    parser.add_argument(
        '-o', '--output-dir', type=str,
        help='出力ディレクトリを指定（デフォルト: ./output）'
    )
    parser.add_argument(
        '--ocr-engine',
        choices=['tesseract', 'manga-ocr', 'sarashina'],
        default='tesseract',
        help='OCRエンジンを指定（デフォルト: tesseract）'
    )
    parser.add_argument(
        '--tessdata-dir', type=str,
        help='tessdataディレクトリを指定（tessdata_best使用時）'
    )
    parser.add_argument(
        '-v', '--verbose', action='store_true',
        help='デバッグ情報を出力'
    )
    parser.add_argument(
        '--text', action='store_true',
        help='.mdと.txtの両方を出力（プレーンテキスト変換）'
    )

    args = parser.parse_args()

    set_verbose(args.verbose)

    if not os.path.exists(args.image_file):
        print(f"エラー: ファイル '{args.image_file}' が見つかりません。")
        sys.exit(1)

    ext = Path(args.image_file).suffix.lower()
    if ext not in IMAGE_EXTENSIONS:
        print(
            f"エラー: 対応していない画像形式です: {ext}\n"
            f"対応形式: {', '.join(IMAGE_EXTENSIONS)}"
        )
        sys.exit(1)

    converter = ImageToMarkdownConverter(
        args.image_file,
        output_dir=args.output_dir,
        ocr_engine=args.ocr_engine,
        tessdata_dir=args.tessdata_dir,
    )
    output_file = converter.convert()

    txt_file = None
    if args.text and output_file and output_file.endswith('.md'):
        from o2md import convert_md_to_text
        auto_patterns = {
            'heading_patterns': [],
            'html_tags': [],
            'line_patterns': [],
        }
        auto_patterns['heading_patterns'] = (
            converter.get_auto_generated_patterns()
        )
        txt_file = convert_md_to_text(
            output_file, auto_patterns=auto_patterns
        )

    print(f"\n変換完了!")
    print(f"出力ファイル: {output_file}")
    print(f"画像フォルダ: {converter.images_dir}")
    if txt_file:
        print(f"テキストファイル: {txt_file}")


if __name__ == "__main__":
    main()
