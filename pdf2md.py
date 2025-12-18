#!/usr/bin/env python3
"""
PDFファイルを画像とテキストに変換するツール

このツールは、指定されたディレクトリ内のPDFファイルを処理し、
各ページを画像ファイルに変換し、テキストを抽出します。
"""

import argparse
import os
import sys
from pathlib import Path

import fitz

_VERBOSE = False


def set_verbose(verbose: bool) -> None:
    """詳細ログ出力の設定"""
    global _VERBOSE
    _VERBOSE = verbose


def log_info(message: str) -> None:
    """情報ログを出力"""
    if _VERBOSE:
        print(f"[INFO] {message}")


def log_error(message: str) -> None:
    """エラーログを出力"""
    print(f"[ERROR] {message}", file=sys.stderr)


class PDFConverter:
    """
    PDFファイルを画像とテキストに変換するクラス
    
    属性:
        pdf_path: 変換対象のPDFファイルパス
        output_dir: 出力先ディレクトリ
        image_format: 出力画像フォーマット（png または jpeg）
    """
    
    def __init__(
        self,
        pdf_path: str,
        output_dir: str,
        image_format: str = "png"
    ):
        """
        PDFConverterの初期化
        
        引数:
            pdf_path: 変換対象のPDFファイルパス
            output_dir: 出力先ディレクトリ
            image_format: 出力画像フォーマット（デフォルト: png）
        """
        self.pdf_path = Path(pdf_path)
        self.output_dir = Path(output_dir)
        self.image_format = image_format.lower()
        
        if self.image_format not in ("png", "jpeg", "jpg"):
            raise ValueError(
                f"サポートされていない画像フォーマット: {image_format}"
            )
        
        if self.image_format == "jpg":
            self.image_format = "jpeg"
    
    def convert(self) -> dict:
        """
        PDFファイルを変換する
        
        戻り値:
            変換結果の辞書（画像パスリスト、テキストパス、ページ数）
        """
        pdf_name = self.pdf_path.stem
        pdf_output_dir = self.output_dir / pdf_name
        images_dir = pdf_output_dir / "images"
        
        images_dir.mkdir(parents=True, exist_ok=True)
        
        log_info(f"PDFファイルを処理中: {self.pdf_path}")
        
        try:
            doc = fitz.open(str(self.pdf_path))
        except Exception as e:
            log_error(f"PDFファイルを開けません: {self.pdf_path} - {e}")
            raise
        
        image_paths = []
        all_text = []
        
        try:
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                image_path = self._convert_page_to_image(
                    page, page_num, images_dir
                )
                image_paths.append(image_path)
                
                page_text = self._extract_text_from_page(page, page_num)
                all_text.append(page_text)
                
                log_info(f"  ページ {page_num + 1}/{len(doc)} を処理完了")
        finally:
            doc.close()
        
        text_path = pdf_output_dir / f"{pdf_name}.txt"
        combined_text = "\n\n".join(all_text)
        text_path.write_text(combined_text, encoding="utf-8")
        
        log_info(f"テキストファイルを保存: {text_path}")
        
        return {
            "image_paths": image_paths,
            "text_path": str(text_path),
            "page_count": len(image_paths)
        }
    
    def _convert_page_to_image(
        self,
        page: fitz.Page,
        page_num: int,
        images_dir: Path
    ) -> str:
        """
        PDFページを画像に変換する
        
        引数:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
            images_dir: 画像出力ディレクトリ
        
        戻り値:
            保存された画像ファイルのパス
        """
        matrix = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=matrix)
        
        ext = "png" if self.image_format == "png" else "jpg"
        image_filename = f"page_{page_num + 1:03d}.{ext}"
        image_path = images_dir / image_filename
        
        if self.image_format == "png":
            pix.save(str(image_path))
        else:
            from PIL import Image
            import io
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            img = img.convert("RGB")
            img.save(str(image_path), "JPEG", quality=95)
        
        log_info(f"    画像を保存: {image_path}")
        return str(image_path)
    
    def _extract_text_from_page(
        self,
        page: fitz.Page,
        page_num: int
    ) -> str:
        """
        PDFページからテキストを抽出する
        
        埋め込みテキストを優先的に抽出し、
        テキストが取得できない場合はOCRを使用する。
        
        引数:
            page: PyMuPDFのページオブジェクト
            page_num: ページ番号（0始まり）
        
        戻り値:
            抽出されたテキスト
        """
        text = page.get_text("text").strip()
        
        if text:
            log_info(f"    ページ {page_num + 1}: 埋め込みテキストを抽出")
            return f"--- ページ {page_num + 1} ---\n{text}"
        
        log_info(f"    ページ {page_num + 1}: OCRでテキストを抽出")
        ocr_text = self._ocr_page(page)
        return f"--- ページ {page_num + 1} (OCR) ---\n{ocr_text}"
    
    def _ocr_page(self, page: fitz.Page) -> str:
        """
        OCRを使用してページからテキストを抽出する
        
        引数:
            page: PyMuPDFのページオブジェクト
        
        戻り値:
            OCRで抽出されたテキスト
        """
        try:
            import pytesseract
            from PIL import Image
            import io
            
            matrix = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=matrix)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
            
            text = pytesseract.image_to_string(img, lang="jpn+eng")
            return text.strip()
        except ImportError:
            log_error(
                "pytesseractがインストールされていません。"
                "OCR機能を使用するには 'pip install pytesseract' を実行してください。"
            )
            return "(OCR利用不可)"
        except Exception as e:
            log_error(f"OCR処理中にエラーが発生: {e}")
            return "(OCRエラー)"


def process_directory(
    input_dir: str,
    output_dir: str,
    image_format: str = "png"
) -> list:
    """
    ディレクトリ内のすべてのPDFファイルを処理する
    
    引数:
        input_dir: PDFファイルが格納されたディレクトリ
        output_dir: 出力先ディレクトリ
        image_format: 出力画像フォーマット
    
    戻り値:
        各PDFファイルの変換結果リスト
    """
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    
    if not input_path.exists():
        raise FileNotFoundError(f"入力ディレクトリが存在しません: {input_dir}")
    
    if not input_path.is_dir():
        raise NotADirectoryError(f"入力パスはディレクトリではありません: {input_dir}")
    
    output_path.mkdir(parents=True, exist_ok=True)
    
    pdf_files = list(input_path.glob("*.pdf")) + list(input_path.glob("*.PDF"))
    
    if not pdf_files:
        log_info(f"PDFファイルが見つかりません: {input_dir}")
        return []
    
    log_info(f"{len(pdf_files)} 個のPDFファイルを検出")
    
    results = []
    for pdf_file in sorted(pdf_files):
        try:
            converter = PDFConverter(
                str(pdf_file),
                str(output_path),
                image_format
            )
            result = converter.convert()
            result["source_file"] = str(pdf_file)
            result["status"] = "success"
            results.append(result)
        except Exception as e:
            log_error(f"ファイル処理中にエラー: {pdf_file} - {e}")
            results.append({
                "source_file": str(pdf_file),
                "status": "error",
                "error": str(e)
            })
    
    return results


def main():
    """メインエントリーポイント"""
    parser = argparse.ArgumentParser(
        description="PDFファイルを画像とテキストに変換するツール"
    )
    parser.add_argument(
        "input_dir",
        help="PDFファイルが格納されたディレクトリのパス"
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help="出力先ディレクトリ（デフォルト: 入力ディレクトリ内のoutput）"
    )
    parser.add_argument(
        "-f", "--format",
        choices=["png", "jpeg", "jpg"],
        default="png",
        help="出力画像フォーマット（デフォルト: png）"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="詳細ログを表示"
    )
    
    args = parser.parse_args()
    
    set_verbose(args.verbose)
    
    output_dir = args.output
    if output_dir is None:
        output_dir = os.path.join(args.input_dir, "output")
    
    try:
        results = process_directory(
            args.input_dir,
            output_dir,
            args.format
        )
        
        success_count = sum(1 for r in results if r["status"] == "success")
        error_count = sum(1 for r in results if r["status"] == "error")
        
        print(f"\n処理完了: {success_count} 成功, {error_count} エラー")
        
        if error_count > 0:
            sys.exit(1)
    except Exception as e:
        log_error(f"処理中にエラーが発生: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
