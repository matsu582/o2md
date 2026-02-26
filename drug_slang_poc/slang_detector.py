"""隠語検出エンジン

2コーパス差分分析と共起ネットワーク分析を組み合わせて
新たな隠語候補を検出するモジュール。

主要手法:
  A. 2コーパス差分分析（Hada方式）
     - Bad Corpus（薬物関連ツイート）とGood Corpus（一般ツイート）で
       それぞれWord2Vecモデルを学習し、同一単語の類似語リストの
       差分から隠語候補を抽出する。
  B. 共起ネットワーク分析
     - PMI（自己相互情報量）を用いて、既知キーワードと密に共起する
       未知語を発見する。
"""

import logging
import math
from collections import Counter
from dataclasses import dataclass, field

from gensim.models import Word2Vec

from drug_slang_poc.config import AnalysisConfig

logger = logging.getLogger(__name__)


@dataclass
class SlangCandidate:
    """隠語候補の情報を保持する構造体"""

    word: str
    score: float
    detection_method: str
    related_seed: str = ""
    context_examples: list[str] = field(default_factory=list)

    def __repr__(self) -> str:
        return (
            f"SlangCandidate("
            f"word='{self.word}', "
            f"score={self.score:.3f}, "
            f"method='{self.detection_method}', "
            f"seed='{self.related_seed}')"
        )


class TwoCorpusDifferentialAnalyzer:
    """2コーパス差分分析による隠語検出

    Hada et al. (ICAART 2023) の手法を参考に、
    Bad Corpus（犯罪関連）とGood Corpus（一般）の
    Word2Vecモデルの差分から隠語候補を抽出する。
    """

    def __init__(self, config: AnalysisConfig):
        self._config = config
        self._bad_model: Word2Vec | None = None
        self._good_model: Word2Vec | None = None

    def train_models(
        self,
        bad_corpus: list[list[str]],
        good_corpus: list[list[str]],
    ) -> None:
        """Bad/Good両コーパスでWord2Vecモデルを学習する

        Args:
            bad_corpus: 薬物関連ツイートのトークン列リスト
            good_corpus: 一般ツイートのトークン列リスト
        """
        logger.info(
            "Bad Corpusモデル学習開始 (%d 文書)", len(bad_corpus)
        )
        self._bad_model = Word2Vec(
            sentences=bad_corpus,
            vector_size=self._config.vector_size,
            window=self._config.window,
            min_count=self._config.min_count,
            epochs=self._config.epochs,
            workers=4,
            seed=42,
        )
        logger.info(
            "Bad Corpusモデル学習完了: 語彙数=%d",
            len(self._bad_model.wv),
        )

        logger.info(
            "Good Corpusモデル学習開始 (%d 文書)", len(good_corpus)
        )
        self._good_model = Word2Vec(
            sentences=good_corpus,
            vector_size=self._config.vector_size,
            window=self._config.window,
            min_count=self._config.min_count,
            epochs=self._config.epochs,
            workers=4,
            seed=42,
        )
        logger.info(
            "Good Corpusモデル学習完了: 語彙数=%d",
            len(self._good_model.wv),
        )

    def find_differential_words(
        self,
        seed_keywords: list[str],
    ) -> list[SlangCandidate]:
        """シードキーワードに対するBad/Good差分語を抽出する

        各シードキーワードについて:
        1. Bad Corpusモデルの類似語上位N語を取得
        2. Good Corpusモデルの類似語上位N語を取得
        3. Badのみに出現する語を「隠語候補」としてスコア付与

        Args:
            seed_keywords: シードキーワードのリスト

        Returns:
            隠語候補のリスト（スコア順）
        """
        if self._bad_model is None or self._good_model is None:
            logger.error("モデル未学習。train_models()を先に実行してください。")
            return []

        candidates_map: dict[str, SlangCandidate] = {}
        top_n = self._config.top_n_similar

        for seed in seed_keywords:
            # Bad Corpusでの類似語
            bad_similar = self._get_similar_words(
                self._bad_model, seed, top_n
            )
            # Good Corpusでの類似語
            good_similar = self._get_similar_words(
                self._good_model, seed, top_n
            )

            if not bad_similar:
                continue

            good_words = {w for w, _ in good_similar}

            for word, bad_score in bad_similar:
                # Goodにも出現する場合はスキップ（一般的な語と判定）
                if word in good_words:
                    continue
                # シードキーワード自体はスキップ
                if word in seed_keywords:
                    continue

                # 差分スコア = Badでの類似度
                if word in candidates_map:
                    # 複数のシードから検出された場合はスコアを加算
                    existing = candidates_map[word]
                    existing.score = max(existing.score, bad_score)
                    existing.related_seed += f", {seed}"
                else:
                    candidates_map[word] = SlangCandidate(
                        word=word,
                        score=bad_score,
                        detection_method="2corpus_differential",
                        related_seed=seed,
                    )

        # スコア降順でソート
        candidates = sorted(
            candidates_map.values(),
            key=lambda c: c.score,
            reverse=True,
        )
        logger.info(
            "2コーパス差分分析: %d 件の隠語候補を検出", len(candidates)
        )
        return candidates

    def _get_similar_words(
        self,
        model: Word2Vec,
        word: str,
        top_n: int,
    ) -> list[tuple[str, float]]:
        """Word2Vecモデルから類似語を取得する

        Args:
            model: 学習済みWord2Vecモデル
            word: 対象語
            top_n: 取得する類似語数

        Returns:
            (語, 類似度) のリスト
        """
        if word not in model.wv:
            return []
        try:
            return model.wv.most_similar(word, topn=top_n)
        except KeyError:
            return []


class CooccurrenceAnalyzer:
    """共起ネットワーク分析による隠語検出

    PMI（自己相互情報量）を用いて、既知の薬物関連キーワードと
    密に共起する未知語を隠語候補として検出する。
    """

    def __init__(self, config: AnalysisConfig):
        self._config = config

    def analyze(
        self,
        tokenized_tweets: list[list[str]],
        seed_keywords: list[str],
    ) -> list[SlangCandidate]:
        """共起分析で隠語候補を検出する

        Args:
            tokenized_tweets: トークン化済みツイートのリスト
            seed_keywords: シードキーワード

        Returns:
            隠語候補のリスト（PMIスコア順）
        """
        # 単語頻度と共起頻度を集計
        word_freq: Counter[str] = Counter()
        pair_freq: Counter[tuple[str, str]] = Counter()
        total_windows = 0
        window = self._config.cooccurrence_window

        for tokens in tokenized_tweets:
            word_freq.update(tokens)
            # ウィンドウ内の共起ペアを集計
            for i, token_a in enumerate(tokens):
                end = min(i + window + 1, len(tokens))
                for j in range(i + 1, end):
                    token_b = tokens[j]
                    if token_a != token_b:
                        pair = tuple(sorted([token_a, token_b]))
                        pair_freq[pair] += 1
                        total_windows += 1

        if total_windows == 0:
            logger.warning("共起ペアが見つかりませんでした")
            return []

        total_words = sum(word_freq.values())
        seed_set = set(seed_keywords)
        candidates_map: dict[str, SlangCandidate] = {}

        # 各シードキーワードとの共起を分析
        for pair, co_count in pair_freq.items():
            word_a, word_b = pair

            # ペアの片方がシード、もう片方が未知語であるもの
            seed_word = None
            unknown_word = None
            if word_a in seed_set and word_b not in seed_set:
                seed_word, unknown_word = word_a, word_b
            elif word_b in seed_set and word_a not in seed_set:
                seed_word, unknown_word = word_b, word_a
            else:
                continue

            # PMI計算
            freq_a = word_freq[seed_word]
            freq_b = word_freq[unknown_word]
            if freq_b < self._config.min_frequency:
                continue

            prob_pair = co_count / total_windows
            prob_a = freq_a / total_words
            prob_b = freq_b / total_words

            if prob_a == 0 or prob_b == 0:
                continue

            pmi = math.log2(prob_pair / (prob_a * prob_b))

            if pmi < self._config.pmi_threshold:
                continue

            if unknown_word in candidates_map:
                existing = candidates_map[unknown_word]
                if pmi > existing.score:
                    existing.score = pmi
                    existing.related_seed = seed_word
            else:
                candidates_map[unknown_word] = SlangCandidate(
                    word=unknown_word,
                    score=pmi,
                    detection_method="cooccurrence_pmi",
                    related_seed=seed_word,
                )

        candidates = sorted(
            candidates_map.values(),
            key=lambda c: c.score,
            reverse=True,
        )
        logger.info(
            "共起分析: %d 件の隠語候補を検出", len(candidates)
        )
        return candidates


class SlangDetectionEngine:
    """隠語検出の統合エンジン

    2コーパス差分分析と共起分析を統合し、
    アンサンブルスコアリングで最終候補を決定する。
    """

    def __init__(self, config: AnalysisConfig | None = None):
        if config is None:
            config = AnalysisConfig()
        self._config = config
        self._differential = TwoCorpusDifferentialAnalyzer(config)
        self._cooccurrence = CooccurrenceAnalyzer(config)

    def detect(
        self,
        bad_corpus: list[list[str]],
        good_corpus: list[list[str]],
        seed_keywords: list[str],
    ) -> list[SlangCandidate]:
        """隠語検出のメインパイプライン

        1. 2コーパス差分分析
        2. 共起ネットワーク分析
        3. アンサンブルスコアリング

        Args:
            bad_corpus: 薬物関連ツイートのトークン列
            good_corpus: 一般ツイートのトークン列
            seed_keywords: シードキーワード

        Returns:
            最終的な隠語候補リスト（アンサンブルスコア順）
        """
        logger.info("=== 隠語検出開始 ===")

        # 手法A: 2コーパス差分分析
        logger.info("--- 手法A: 2コーパス差分分析 ---")
        self._differential.train_models(bad_corpus, good_corpus)
        diff_candidates = self._differential.find_differential_words(
            seed_keywords
        )

        # 手法B: 共起ネットワーク分析（Bad Corpusのみで実施）
        logger.info("--- 手法B: 共起ネットワーク分析 ---")
        cooc_candidates = self._cooccurrence.analyze(
            bad_corpus, seed_keywords
        )

        # アンサンブル: 両手法の結果を統合
        ensemble = self._ensemble_candidates(
            diff_candidates, cooc_candidates
        )

        logger.info("=== 隠語検出完了: %d 件の候補 ===", len(ensemble))
        return ensemble

    def _ensemble_candidates(
        self,
        diff_candidates: list[SlangCandidate],
        cooc_candidates: list[SlangCandidate],
    ) -> list[SlangCandidate]:
        """2つの手法の結果をアンサンブルする

        同一単語が両手法で検出された場合はスコアを加重平均し、
        片方のみの場合は元スコアの半分とする。

        Args:
            diff_candidates: 差分分析の候補
            cooc_candidates: 共起分析の候補

        Returns:
            統合された候補リスト
        """
        score_map: dict[str, dict] = {}

        # 差分分析結果を登録（正規化スコア0-1）
        if diff_candidates:
            max_diff = max(c.score for c in diff_candidates)
            for c in diff_candidates:
                norm_score = c.score / max_diff if max_diff > 0 else 0
                score_map[c.word] = {
                    "diff_score": norm_score,
                    "cooc_score": 0.0,
                    "related_seed": c.related_seed,
                    "methods": [c.detection_method],
                }

        # 共起分析結果を統合（正規化スコア0-1）
        if cooc_candidates:
            max_cooc = max(c.score for c in cooc_candidates)
            for c in cooc_candidates:
                norm_score = c.score / max_cooc if max_cooc > 0 else 0
                if c.word in score_map:
                    score_map[c.word]["cooc_score"] = norm_score
                    score_map[c.word]["methods"].append(c.detection_method)
                else:
                    score_map[c.word] = {
                        "diff_score": 0.0,
                        "cooc_score": norm_score,
                        "related_seed": c.related_seed,
                        "methods": [c.detection_method],
                    }

        # アンサンブルスコア計算
        # 両手法で検出 → ボーナス
        ensemble_results: list[SlangCandidate] = []
        for word, data in score_map.items():
            method_count = len(data["methods"])
            base_score = (
                0.6 * data["diff_score"] + 0.4 * data["cooc_score"]
            )
            # 両手法で検出された場合はボーナス（1.3倍）
            if method_count >= 2:
                final_score = base_score * 1.3
            else:
                final_score = base_score

            method_str = " + ".join(data["methods"])
            ensemble_results.append(
                SlangCandidate(
                    word=word,
                    score=final_score,
                    detection_method=method_str,
                    related_seed=data["related_seed"],
                )
            )

        ensemble_results.sort(key=lambda c: c.score, reverse=True)
        return ensemble_results
