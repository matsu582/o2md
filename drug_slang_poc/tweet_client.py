"""ツイート検索APIクライアント

既存の検索APIを利用してツイートを収集するモジュール。
Bad Corpus（薬物関連）とGood Corpus（一般）の両方を収集する。
"""

import logging
import time
import urllib.parse
from dataclasses import dataclass

import requests

from drug_slang_poc.config import ApiConfig

logger = logging.getLogger(__name__)


@dataclass
class Tweet:
    """ツイートデータを保持する構造体"""

    tweet_id: str
    text: str
    author_id: str
    created_at: str
    lang: str = "ja"


class TweetSearchClient:
    """ツイート検索APIのクライアント

    既存の検索エンドポイントを利用し、キーワードベースで
    ツイートを収集する。ページネーション対応。
    """

    def __init__(self, config: ApiConfig):
        self._config = config
        self._session = requests.Session()
        self._session.auth = (config.username, config.password)

    def search(
        self,
        query: str,
        count: int | None = None,
        max_pages: int | None = None,
    ) -> list[Tweet]:
        """キーワードでツイートを検索する

        Args:
            query: 検索クエリ（URL未エンコード）
            count: 1ページあたりの取得件数
            max_pages: 最大ページ数

        Returns:
            取得したツイートのリスト
        """
        if count is None:
            count = self._config.default_count
        if max_pages is None:
            max_pages = self._config.max_pages

        tweets: list[Tweet] = []
        encoded_query = urllib.parse.quote(query, safe="")
        url = (
            f"{self._config.base_url}"
            f"?q={encoded_query}&count={count}"
        )

        for page in range(max_pages):
            logger.info("ページ %d を取得中: %s", page + 1, url[:120])
            try:
                resp = self._session.get(url, timeout=30)
                resp.raise_for_status()
            except requests.RequestException as exc:
                logger.warning("API呼び出し失敗 (ページ %d): %s", page + 1, exc)
                break

            data = resp.json()
            statuses = data.get("statuses", [])
            if not statuses:
                logger.info("取得結果が空のため終了")
                break

            for status in statuses:
                tweet_data = status.get("data", {})
                if tweet_data.get("lang") != "ja":
                    continue
                tweet = Tweet(
                    tweet_id=tweet_data.get("id", ""),
                    text=tweet_data.get("text", ""),
                    author_id=tweet_data.get("author_id", ""),
                    created_at=tweet_data.get("created_at", ""),
                    lang=tweet_data.get("lang", "ja"),
                )
                if tweet.text:
                    tweets.append(tweet)

            # 次ページのURL取得
            metadata = data.get("search_metadata", {})
            next_results = metadata.get("next_results", "")
            if not next_results:
                logger.info("次ページなし、取得終了")
                break

            url = f"{self._config.base_url}{next_results}"
            time.sleep(self._config.request_interval_sec)

        logger.info("合計 %d 件のツイートを取得", len(tweets))
        return tweets

    def collect_by_keywords(
        self,
        keywords: list[str],
        count_per_keyword: int = 100,
        max_pages: int = 5,
    ) -> list[Tweet]:
        """複数キーワードでOR検索してツイートを収集する

        APIのクエリ仕様に合わせて、キーワードをOR結合して検索する。
        重複排除も行う。

        Args:
            keywords: 検索キーワードのリスト
            count_per_keyword: キーワードあたりの取得件数
            max_pages: キーワードあたりの最大ページ数

        Returns:
            重複排除済みのツイートリスト
        """
        seen_ids: set[str] = set()
        all_tweets: list[Tweet] = []

        # APIの制限を考慮し、キーワードを5個ずつのグループに分割
        batch_size = 5
        for i in range(0, len(keywords), batch_size):
            batch = keywords[i : i + batch_size]
            # ダブルクォートで完全一致、ORで結合
            parts = [f'"{kw}"' for kw in batch]
            query = " OR ".join(parts)

            logger.info("検索バッチ: %s", query[:80])
            tweets = self.search(
                query=query,
                count=count_per_keyword,
                max_pages=max_pages,
            )

            for tweet in tweets:
                if tweet.tweet_id not in seen_ids:
                    seen_ids.add(tweet.tweet_id)
                    all_tweets.append(tweet)

            time.sleep(self._config.request_interval_sec)

        logger.info("全バッチ合計: %d 件（重複排除済み）", len(all_tweets))
        return all_tweets

    def collect_general_tweets(
        self,
        general_keywords: list[str] | None = None,
        count: int = 100,
        max_pages: int = 5,
    ) -> list[Tweet]:
        """Good Corpus用の一般ツイートを収集する

        薬物とは無関係な日常的なトピックのツイートを収集する。
        2コーパス差分分析のベースライン（Good Corpus）として使用する。

        Args:
            general_keywords: 一般的なキーワード（指定しない場合はデフォルト使用）
            count: 取得件数
            max_pages: 最大ページ数

        Returns:
            一般ツイートのリスト
        """
        if general_keywords is None:
            general_keywords = [
                "天気",
                "ランチ",
                "仕事",
                "映画",
                "音楽",
                "電車",
                "カフェ",
                "散歩",
                "買い物",
                "料理",
            ]
        return self.collect_by_keywords(
            keywords=general_keywords,
            count_per_keyword=count,
            max_pages=max_pages,
        )
