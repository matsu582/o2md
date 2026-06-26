"""
多言語対応モジュール

gettextを使用してユーザー向けメッセージの多言語化を実現する。
デフォルト言語は日本語(ja)。--langオプションまたはO2MD_LANG環境変数で切替可能。
"""

import gettext
import os
from pathlib import Path

# 翻訳オブジェクトを保持するグローバル変数
_translation = None
_current_lang = None

# localeディレクトリのパス
_LOCALE_DIR = Path(__file__).parent / "locale"


def setup_i18n(lang: str | None = None) -> None:
    """多言語設定を初期化する

    Args:
        lang: 言語コード ('ja', 'en' など)。
              Noneの場合、O2MD_LANG環境変数→LANGの順で判定。
              未指定時は'ja'をデフォルトとする。
    """
    global _translation, _current_lang

    if lang is None:
        lang = os.environ.get("O2MD_LANG", "").strip()

    if not lang:
        env_lang = os.environ.get("LANG", "")
        if env_lang.startswith("en"):
            lang = "en"
        else:
            lang = "ja"

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
