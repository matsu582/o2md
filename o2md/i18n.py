"""
多言語対応モジュール

gettextを使用してユーザー向けメッセージの多言語化を実現する。
デフォルト言語は日本語(ja)。--langオプションまたはLANG環境変数で切替可能。
"""

import gettext
import os
from pathlib import Path

# 翻訳オブジェクトを保持するグローバル変数
_translation = None
_current_lang = None

# localeディレクトリのパス
_LOCALE_DIR = Path(__file__).parent / "locale"


def _parse_lang_env(env_lang: str) -> str:
    """LANG環境変数から言語コードを抽出する

    例: 'ja_JP.UTF-8' → 'ja', 'en_US.UTF-8' → 'en', 'C' → 'ja'
    """
    if not env_lang or env_lang in ("C", "POSIX"):
        return "ja"
    # 'ja_JP.UTF-8' → 'ja_JP' → 'ja'
    lang_part = env_lang.split(".")[0]  # エンコーディング除去
    lang_code = lang_part.split("_")[0]  # 国コード除去
    if lang_code == "en":
        return "en"
    return "ja"


def setup_i18n(lang: str | None = None) -> None:
    """多言語設定を初期化する

    Args:
        lang: 言語コード ('ja', 'en' など)。
              Noneの場合、LANG環境変数から判定。
              未指定時は'ja'をデフォルトとする。
    """
    global _translation, _current_lang

    if not lang:
        lang = _parse_lang_env(os.environ.get("LANG", ""))

    _current_lang = lang

    try:
        _translation = gettext.translation(
            "o2md",
            localedir=str(_LOCALE_DIR),
            languages=[lang],
        )
    except FileNotFoundError:
        # 翻訳ファイルが存在しない場合はNullTranslationsを使用
        _translation = gettext.NullTranslations()

    _translation.install()


def get_current_lang() -> str:
    """現在設定されている言語コードを返す"""
    return _current_lang or "ja"


def _(message: str) -> str:
    """翻訳関数

    メッセージを現在の言語に翻訳する。
    setup_i18nが未呼び出しの場合は原文をそのまま返す。
    """
    if _translation is None:
        return message
    return _translation.gettext(message)
