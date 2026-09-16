"""x2md 系で使うテキスト判定ヒューリスティックの共通モジュール

パス/URL/XML らしい文字列の判定と、番号付きリスト（①、1.、(a) 等）の
マーカー判定を一箇所に集約する。これまで x2md.py / x2md_tables.py に
散在していた同種判定を共通化したもの。
"""

import re
import unicodedata

ENUMERATED_LIST_MIN_RATIO = 0.8  # 左列のうち番号らしいセルの割合
ENUMERATED_LIST_MIN_RIGHT_AVG_PLAIN = 10  # _is_plain_text_region 用の右列平均長
ENUMERATED_LIST_MIN_RIGHT_AVG_IMPLICIT = 8  # 暗黙テーブル検出用の右列平均長

CIRCLED_DIGITS = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳'

# 番号付きリストのマーカー: (1) / （1） / 1) / 1． / I / (a) 等
ENUM_MARKER_RE = re.compile(r'^[\(\（]?\s*(?:\d+|[IVXivx]+|[A-Za-z])\s*[\)\）]?[\.．]?$')
# フォールバック: 1文字の英数・ハイフン（例: '-', 'a', '1'）
SINGLE_CHAR_MARKER_RE = re.compile(r'^[A-Za-z0-9\-]$')


def is_path_like(text: str) -> bool:
    """パス・URL・XML/タグらしい文字列かを判定する

    条件: ('\\' and ':') or '/' or 'http'始まり or 'xml'含有 or ('<' and '>')
    """
    if not text:
        return False
    lower = text.lower()
    return (
        ('\\' in text and ':' in text)
        or '/' in text
        or lower.startswith('http')
        or 'xml' in lower
        or ('<' in text and '>' in text)
    )


def is_enum_marker(text: str) -> bool:
    """番号付きリストのマーカーらしい文字列かを判定する

    丸数字 → NFKC正規化して ENUM_MARKER_RE → 1文字なら SINGLE_CHAR_MARKER_RE
    の順に評価する。
    """
    if not text:
        return False
    tt = text.strip()
    # 丸数字を最初にチェック（ASCIIに正規化されない）
    if any(ch in CIRCLED_DIGITS for ch in tt):
        return True
    try:
        nn = unicodedata.normalize('NFKC', tt)
    except Exception:
        nn = tt
    try:
        if ENUM_MARKER_RE.match(nn):
            return True
    except Exception:
        pass
    try:
        if len(nn.strip()) == 1 and SINGLE_CHAR_MARKER_RE.match(nn.strip()):
            return True
    except Exception:
        pass
    return False


def enum_marker_ratio(texts) -> float:
    """texts のうち番号マーカーらしいものの割合を返す"""
    if not texts:
        return 0.0
    matches = sum(1 for t in texts if is_enum_marker(t))
    return matches / len(texts)


def looks_like_enumerated_list(left_texts, right_texts, min_right_avg) -> bool:
    """2列領域が番号付きリストらしいかを判定する

    左列の番号率 >= ENUMERATED_LIST_MIN_RATIO かつ
    右列平均長 >= min_right_avg の場合に True。
    """
    if not left_texts or not right_texts or len(left_texts) < 2:
        return False
    ratio = enum_marker_ratio(left_texts)
    right_avg = sum(len(s) for s in right_texts) / len(right_texts)
    return ratio >= ENUMERATED_LIST_MIN_RATIO and right_avg >= min_right_avg
