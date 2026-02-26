"""ノイズフィルタリングモジュール

隠語候補からノイズ（数字、単一文字、一般的すぎる語など）を
除去して結果の精度を向上させる。
"""

import logging
import re

from drug_slang_poc.slang_detector import SlangCandidate

logger = logging.getLogger(__name__)

# 数字のみ・英字1-2文字のパターン
NUMERIC_PATTERN = re.compile(r"^\d+$")
SHORT_ALPHA_PATTERN = re.compile(r"^[A-Za-z]{1,2}$")

# 明らかにノイズとなる一般語
GENERAL_NOISE_WORDS: set[str] = {
    "SNS", "POP", "DM", "LINE", "PR", "RT", "FF",
    "www", "ww", "笑", "草",
}

# 年月日などの時間表現パターン
DATE_PATTERN = re.compile(
    r"^(20\d{2}|1[0-2]|[1-9]月|[0-3]?\d日|月曜|火曜|水曜|木曜|金曜|土曜|日曜)$"
)


def is_noise_candidate(word: str) -> bool:
    """単語がノイズかどうかを判定する

    以下の条件のいずれかに該当する場合、ノイズと判定:
    - 数字のみで構成
    - 英字1-2文字のみ
    - 一般的なノイズ語リストに含まれる
    - 年号・日付パターン

    Args:
        word: 判定対象の単語

    Returns:
        ノイズであればTrue
    """
    # 数字のみ
    if NUMERIC_PATTERN.match(word):
        return True
    # 英字1-2文字
    if SHORT_ALPHA_PATTERN.match(word):
        return True
    # 一般ノイズ語
    if word in GENERAL_NOISE_WORDS:
        return True
    # 日付パターン
    if DATE_PATTERN.match(word):
        return True
    return False


def filter_candidates(
    candidates: list[SlangCandidate],
    min_word_length: int = 2,
) -> list[SlangCandidate]:
    """隠語候補リストからノイズを除去する

    Args:
        candidates: フィルタ前の候補リスト
        min_word_length: 最小文字数（デフォルト2）

    Returns:
        フィルタ済みの候補リスト
    """
    filtered: list[SlangCandidate] = []
    removed_count = 0

    for candidate in candidates:
        word = candidate.word

        # 最小文字数チェック
        if len(word) < min_word_length:
            removed_count += 1
            continue

        # ノイズ判定
        if is_noise_candidate(word):
            removed_count += 1
            continue

        filtered.append(candidate)

    logger.info(
        "ノイズフィルタ: %d 件 → %d 件（%d 件除去）",
        len(candidates),
        len(filtered),
        removed_count,
    )
    return filtered
