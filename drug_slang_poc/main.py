"""薬物取引隠語検出PoC メインスクリプト

シード→収集→分析→候補出力の一連のパイプラインを実行する。

使用方法:
    python -m drug_slang_poc.main [オプション]

環境変数:
    TWEET_API_USER: API認証ユーザー名
    TWEET_API_PASS: API認証パスワード
"""

import argparse
import json
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

from drug_slang_poc.config import (
    AnalysisConfig,
    ApiConfig,
    get_all_seed_keywords,
    get_bad_corpus_keywords,
)
from drug_slang_poc.nlp_pipeline import TextPreprocessor
from drug_slang_poc.noise_filter import filter_candidates
from drug_slang_poc.slang_detector import SlangCandidate, SlangDetectionEngine
from drug_slang_poc.tweet_client import TweetSearchClient

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    """ロギングを設定する"""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def collect_tweets(
    client: TweetSearchClient,
    bad_keywords: list[str],
    count_per_keyword: int,
    max_pages: int,
) -> tuple[list[str], list[str]]:
    """Bad/Good両コーパスのツイートを収集する

    Args:
        client: ツイート検索クライアント
        bad_keywords: 薬物関連キーワード
        count_per_keyword: キーワードあたりの取得件数
        max_pages: 最大ページ数

    Returns:
        (bad_texts, good_texts) のタプル
    """
    logger.info("=== Bad Corpus（薬物関連ツイート）収集開始 ===")
    bad_tweets = client.collect_by_keywords(
        keywords=bad_keywords,
        count_per_keyword=count_per_keyword,
        max_pages=max_pages,
    )
    bad_texts = [t.text for t in bad_tweets]
    logger.info("Bad Corpus: %d 件収集", len(bad_texts))

    logger.info("=== Good Corpus（一般ツイート）収集開始 ===")
    good_tweets = client.collect_general_tweets(
        count=count_per_keyword,
        max_pages=max_pages,
    )
    good_texts = [t.text for t in good_tweets]
    logger.info("Good Corpus: %d 件収集", len(good_texts))

    return bad_texts, good_texts


def run_analysis(
    bad_texts: list[str],
    good_texts: list[str],
    seed_keywords: list[str],
    analysis_config: AnalysisConfig,
) -> list[SlangCandidate]:
    """NLP分析パイプラインを実行する

    Args:
        bad_texts: 薬物関連ツイートテキスト
        good_texts: 一般ツイートテキスト
        seed_keywords: シードキーワード
        analysis_config: 分析パラメータ

    Returns:
        隠語候補リスト
    """
    preprocessor = TextPreprocessor()

    logger.info("=== NLP前処理開始 ===")
    bad_corpus = preprocessor.process_tweets_batch(bad_texts)
    good_corpus = preprocessor.process_tweets_batch(good_texts)

    if len(bad_corpus) < 10:
        logger.warning(
            "Bad Corpusの有効文書数が少なすぎます(%d件)。"
            "結果の信頼性が低い可能性があります。",
            len(bad_corpus),
        )
    if len(good_corpus) < 10:
        logger.warning(
            "Good Corpusの有効文書数が少なすぎます(%d件)。"
            "結果の信頼性が低い可能性があります。",
            len(good_corpus),
        )

    logger.info("=== 隠語検出エンジン実行 ===")
    engine = SlangDetectionEngine(config=analysis_config)
    candidates = engine.detect(
        bad_corpus=bad_corpus,
        good_corpus=good_corpus,
        seed_keywords=seed_keywords,
    )
    return candidates


def save_results(
    candidates: list[SlangCandidate],
    output_path: Path,
) -> None:
    """検出結果をJSONファイルに保存する

    Args:
        candidates: 隠語候補リスト
        output_path: 出力ファイルパス
    """
    timestamp = datetime.now(timezone.utc).isoformat()
    result = {
        "generated_at": timestamp,
        "total_candidates": len(candidates),
        "candidates": [
            {
                "rank": i + 1,
                "word": c.word,
                "score": round(c.score, 4),
                "detection_method": c.detection_method,
                "related_seed": c.related_seed,
            }
            for i, c in enumerate(candidates)
        ],
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    logger.info("結果を保存しました: %s", output_path)


def print_top_candidates(
    candidates: list[SlangCandidate],
    top_n: int = 30,
) -> None:
    """上位N件の候補をコンソールに表示する

    Args:
        candidates: 隠語候補リスト
        top_n: 表示件数
    """
    print("\n" + "=" * 70)
    print("検出された隠語候補（上位 {} 件）".format(min(top_n, len(candidates))))
    print("=" * 70)
    print(f"{'順位':>4}  {'スコア':>7}  {'検出手法':<30}  {'関連シード':<10}  {'候補語'}")
    print("-" * 70)

    display_count = min(top_n, len(candidates))
    for i in range(display_count):
        c = candidates[i]
        print(
            f"{i + 1:>4}  "
            f"{c.score:>7.4f}  "
            f"{c.detection_method:<30}  "
            f"{c.related_seed:<10}  "
            f"{c.word}"
        )

    print("=" * 70)
    print(f"合計候補数: {len(candidates)}")


def build_arg_parser() -> argparse.ArgumentParser:
    """コマンドライン引数パーサーを構築する"""
    parser = argparse.ArgumentParser(
        description="薬物取引隠語検出PoC - ツイート分析パイプライン",
    )
    parser.add_argument(
        "--count",
        type=int,
        default=100,
        help="キーワードあたりの取得ツイート数（デフォルト: 100）",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=5,
        help="キーワードあたりの最大ページ数（デフォルト: 5）",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="output/slang_candidates.json",
        help="結果出力先ファイルパス",
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=30,
        help="表示する上位候補数（デフォルト: 30）",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="詳細ログを出力",
    )
    parser.add_argument(
        "--api-user",
        type=str,
        default="",
        help="API認証ユーザー名（環境変数TWEET_API_USERでも指定可）",
    )
    parser.add_argument(
        "--api-pass",
        type=str,
        default="",
        help="API認証パスワード（環境変数TWEET_API_PASSでも指定可）",
    )
    return parser


def main() -> None:
    """メイン処理"""
    parser = build_arg_parser()
    args = parser.parse_args()

    setup_logging(verbose=args.verbose)
    logger.info("薬物取引隠語検出PoC を開始します")

    # API設定
    api_config = ApiConfig()
    if args.api_user:
        api_config.username = args.api_user
    if args.api_pass:
        api_config.password = args.api_pass

    if not api_config.username or not api_config.password:
        logger.error(
            "API認証情報が未設定です。"
            "--api-user/--api-pass または "
            "環境変数 TWEET_API_USER/TWEET_API_PASS を設定してください。"
        )
        sys.exit(1)

    # 分析設定
    analysis_config = AnalysisConfig()

    # ツイート収集
    client = TweetSearchClient(api_config)
    bad_keywords = get_bad_corpus_keywords()
    bad_texts, good_texts = collect_tweets(
        client=client,
        bad_keywords=bad_keywords,
        count_per_keyword=args.count,
        max_pages=args.max_pages,
    )

    if not bad_texts:
        logger.error("Bad Corpusのツイートを取得できませんでした。終了します。")
        sys.exit(1)
    if not good_texts:
        logger.error("Good Corpusのツイートを取得できませんでした。終了します。")
        sys.exit(1)

    # 分析実行
    seed_keywords = get_all_seed_keywords()
    candidates = run_analysis(
        bad_texts=bad_texts,
        good_texts=good_texts,
        seed_keywords=seed_keywords,
        analysis_config=analysis_config,
    )

    # ノイズフィルタリング
    candidates = filter_candidates(candidates)

    # 結果表示・保存
    print_top_candidates(candidates, top_n=args.top_n)

    output_path = Path(args.output)
    save_results(candidates, output_path)

    logger.info("処理完了")


if __name__ == "__main__":
    main()
