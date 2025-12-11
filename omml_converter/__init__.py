# -*- coding: utf-8 -*-
"""
数式変換モジュール

Office Math Markup Language (OMML) を LaTeX 形式に変換するためのモジュール。
markitdown (MIT License) の実装を参考にしています。
"""

from .omml import oMath2Latex, OMML_NS
from .pre_process import pre_process_docx

__all__ = ['oMath2Latex', 'OMML_NS', 'pre_process_docx']
