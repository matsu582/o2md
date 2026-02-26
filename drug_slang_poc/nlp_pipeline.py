"""NLP前処理パイプライン

MeCabによる形態素解析とテキストクリーニングを行うモジュール。
ツイートのテキストをWord2Vec学習用のトークン列に変換する。
"""

import logging
import re
import unicodedata

import MeCab

logger = logging.getLogger(__name__)

# ストップワード（分析に不要な一般的な語）
STOP_WORDS: set[str] = {
    "の", "に", "は", "を", "た", "が", "で", "て", "と", "し",
    "れ", "さ", "ある", "いる", "も", "する", "から", "な", "こと",
    "として", "い", "や", "れる", "など", "なっ", "ない", "この",
    "ため", "その", "あっ", "よう", "また", "もの", "という",
    "あり", "まで", "られ", "なる", "へ", "か", "だ", "これ",
    "によって", "により", "おり", "より", "による", "ず", "なり",
    "られる", "において", "ば", "なかっ", "なく", "しかし",
    "について", "せ", "だっ", "その他", "できる", "それ",
    "う", "ので", "なお", "のみ", "でき", "き", "つ",
    "における", "および", "いう", "さらに", "でも", "ら", "たり",
    "RT", "rt", "https", "http", "co",
}

# ツイート特有のパターン除去用正規表現
URL_PATTERN = re.compile(r"https?://\S+")
MENTION_PATTERN = re.compile(r"@\w+")
HASHTAG_EXTRACT_PATTERN = re.compile(r"#(\S+)")
EMOJI_PATTERN = re.compile(
    "["
    "\U0001f600-\U0001f64f"
    "\U0001f300-\U0001f5ff"
    "\U0001f680-\U0001f6ff"
    "\U0001f1e0-\U0001f1ff"
    "\U00002702-\U000027b0"
    "\U000024c2-\U0001f251"
    "]+",
    flags=re.UNICODE,
)


class TextPreprocessor:
    """ツイートテキストの前処理を行うクラス

    URL除去、メンション除去、Unicode正規化、
    MeCab形態素解析を組み合わせてトークン列を生成する。
    """

    def __init__(self):
        """MeCab辞書を初期化する"""
        self._tagger = MeCab.Tagger()
        # 初回解析でヘッダ情報を破棄
        self._tagger.parse("")

    def clean_text(self, text: str) -> str:
        """ツイートテキストをクリーニングする

        Args:
            text: 元のツイートテキスト

        Returns:
            クリーニング済みテキスト
        """
        # URL除去
        cleaned = URL_PATTERN.sub("", text)
        # メンション除去
        cleaned = MENTION_PATTERN.sub("", cleaned)
        # 絵文字除去
        cleaned = EMOJI_PATTERN.sub("", cleaned)
        # Unicode正規化（全角→半角統一など）
        cleaned = unicodedata.normalize("NFKC", cleaned)
        # 連続空白を単一スペースに
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned

    def extract_hashtags(self, text: str) -> list[str]:
        """ツイートからハッシュタグを抽出する

        Args:
            text: 元のツイートテキスト

        Returns:
            ハッシュタグのリスト（#記号なし）
        """
        return HASHTAG_EXTRACT_PATTERN.findall(text)

    def tokenize(self, text: str) -> list[str]:
        """テキストをMeCabで形態素解析しトークン列を返す

        名詞・動詞・形容詞のみを抽出し、ストップワードを除去する。

        Args:
            text: クリーニング済みテキスト

        Returns:
            トークンのリスト
        """
        tokens: list[str] = []
        node = self._tagger.parseToNode(text)

        while node:
            surface = node.surface
            features = node.feature.split(",")
            pos = features[0] if features else ""

            # 名詞・動詞・形容詞のみ抽出
            if pos in ("名詞", "動詞", "形容詞"):
                # 1文字の平仮名・カタカナはスキップ
                if len(surface) == 1 and re.match(r"[\u3040-\u30ff]", surface):
                    node = node.next
                    continue
                # ストップワード除去
                if surface.lower() not in STOP_WORDS:
                    tokens.append(surface)

            node = node.next

        return tokens

    def process_tweet(self, text: str) -> list[str]:
        """ツイート1件を前処理してトークン列を返す

        クリーニング→形態素解析→フィルタリングを一括で行う。

        Args:
            text: 元のツイートテキスト

        Returns:
            処理済みトークンのリスト
        """
        cleaned = self.clean_text(text)
        tokens = self.tokenize(cleaned)
        # ハッシュタグも独立したトークンとして追加
        hashtags = self.extract_hashtags(text)
        tokens.extend(hashtags)
        return tokens

    def process_tweets_batch(self, texts: list[str]) -> list[list[str]]:
        """複数ツイートを一括処理する

        Args:
            texts: ツイートテキストのリスト

        Returns:
            各ツイートのトークン列のリスト
        """
        result: list[list[str]] = []
        for i, text in enumerate(texts):
            tokens = self.process_tweet(text)
            if tokens:
                result.append(tokens)
            if (i + 1) % 500 == 0:
                logger.info("前処理進捗: %d/%d 件完了", i + 1, len(texts))
        logger.info("前処理完了: %d 件 → %d 件（空を除外）", len(texts), len(result))
        return result
