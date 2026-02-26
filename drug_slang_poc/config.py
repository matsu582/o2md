"""設定管理モジュール

シードキーワード、API設定、分析パラメータを管理する。
"""

import os
from dataclasses import dataclass, field


@dataclass
class ApiConfig:
    """ツイート検索APIの接続設定"""

    base_url: str = "https://txap.nazuki-oto.com/search/tweets"
    username: str = ""
    password: str = ""
    default_count: int = 100
    max_pages: int = 10
    request_interval_sec: float = 1.0

    def __post_init__(self):
        """環境変数からAPI認証情報を読み込む"""
        if not self.username:
            self.username = os.environ.get("TWEET_API_USER", "")
        if not self.password:
            self.password = os.environ.get("TWEET_API_PASS", "")


@dataclass
class AnalysisConfig:
    """NLP分析のパラメータ設定"""

    # Word2Vecパラメータ
    vector_size: int = 100
    window: int = 5
    min_count: int = 3
    epochs: int = 20

    # 隠語検出パラメータ
    top_n_similar: int = 20
    candidate_threshold: float = 0.3
    min_frequency: int = 2

    # 共起分析パラメータ
    cooccurrence_window: int = 3
    pmi_threshold: float = 2.0


# 薬物取引に関するシードキーワード
# カテゴリ別に分類して管理
SEED_KEYWORDS: dict[str, list[str]] = {
    "direct_terms": [
        # 直接的な薬物名（既知の関連語）
        "大麻",
        "覚醒剤",
        "マリファナ",
        "コカイン",
        "MDMA",
        "LSD",
        "ヘロイン",
        "覚せい剤",
    ],
    "known_slang": [
        # 既知の隠語
        "手押し",
        "野菜",
        "アイス",
        "草",
        "ガンジャ",
        "シャブ",
        "クリスタル",
        "葉っぱ",
        "ハッパ",
        "ブリブリ",
        "キメる",
        "キマる",
        "パキる",
    ],
    "transaction_terms": [
        # 取引を示唆する表現
        "手渡し",
        "対面取引",
        "デリ",
        "配達",
        "在庫あり",
    ],
    "hashtags": [
        # 関連ハッシュタグ
        "#お薬もぐもぐ",
        "#お薬譲ります",
    ],
}


def get_all_seed_keywords() -> list[str]:
    """全カテゴリのシードキーワードをフラットなリストで返す"""
    result: list[str] = []
    for keywords in SEED_KEYWORDS.values():
        result.extend(keywords)
    return result


def get_bad_corpus_keywords() -> list[str]:
    """Bad Corpus（犯罪関連ツイート）収集用のキーワードを返す

    直接的な薬物名と既知の隠語を組み合わせて使用する。
    """
    bad_keywords: list[str] = []
    bad_keywords.extend(SEED_KEYWORDS.get("direct_terms", []))
    bad_keywords.extend(SEED_KEYWORDS.get("known_slang", []))
    bad_keywords.extend(SEED_KEYWORDS.get("hashtags", []))
    return bad_keywords
