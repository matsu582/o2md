#!/usr/bin/env python3
"""
MPXJ JAR 自動ダウンロード・管理モジュール

MS Project ファイル変換に必要な MPXJ 13.5.1 と依存 JAR を、
Maven Central から o2md パッケージのインストール済みフォルダ内
(o2md/libs/) にダウンロードして使用する。

Gradle/Maven コマンド実行に依存しない Pure Python 実装。
"""

import os
import urllib.request
import urllib.error
from pathlib import Path


# MPXJ 13.5.1 が必要とする JAR と Maven Central URL
# （Gradle が実際に解決した依存バージョン・正しい groupId 座標を使用）
_BASE = "https://repo1.maven.org/maven2"
MPXJ_DEPENDENCIES = {
    "mpxj-13.5.1.jar": f"{_BASE}/net/sf/mpxj/mpxj/13.5.1/mpxj-13.5.1.jar",
    "rtfparserkit-1.16.0.jar": f"{_BASE}/com/github/joniles/rtfparserkit/1.16.0/rtfparserkit-1.16.0.jar",
    "poi-5.3.0.jar": f"{_BASE}/org/apache/poi/poi/5.3.0/poi-5.3.0.jar",
    "commons-io-2.16.1.jar": f"{_BASE}/commons-io/commons-io/2.16.1/commons-io-2.16.1.jar",
    "commons-codec-1.17.0.jar": f"{_BASE}/commons-codec/commons-codec/1.17.0/commons-codec-1.17.0.jar",
    "commons-collections4-4.4.jar": f"{_BASE}/org/apache/commons/commons-collections4/4.4/commons-collections4-4.4.jar",
    "commons-lang3-3.10.jar": f"{_BASE}/org/apache/commons/commons-lang3/3.10/commons-lang3-3.10.jar",
    "commons-logging-1.2.jar": f"{_BASE}/commons-logging/commons-logging/1.2/commons-logging-1.2.jar",
    "commons-math3-3.6.1.jar": f"{_BASE}/org/apache/commons/commons-math3/3.6.1/commons-math3-3.6.1.jar",
    "log4j-api-2.23.1.jar": f"{_BASE}/org/apache/logging/log4j/log4j-api/2.23.1/log4j-api-2.23.1.jar",
    "log4j-core-2.23.1.jar": f"{_BASE}/org/apache/logging/log4j/log4j-core/2.23.1/log4j-core-2.23.1.jar",
    "jsoup-1.15.3.jar": f"{_BASE}/org/jsoup/jsoup/1.15.3/jsoup-1.15.3.jar",
    "jackcess-4.0.1.jar": f"{_BASE}/com/healthmarketscience/jackcess/jackcess/4.0.1/jackcess-4.0.1.jar",
    "sqlite-jdbc-3.42.0.0.jar": f"{_BASE}/org/xerial/sqlite-jdbc/3.42.0.0/sqlite-jdbc-3.42.0.0.jar",
    "SparseBitSet-1.3.jar": f"{_BASE}/com/zaxxer/SparseBitSet/1.3/SparseBitSet-1.3.jar",
    "jakarta.activation-2.0.1.jar": f"{_BASE}/com/sun/activation/jakarta.activation/2.0.1/jakarta.activation-2.0.1.jar",
    "jakarta.xml.bind-api-3.0.1.jar": f"{_BASE}/jakarta/xml/bind/jakarta.xml.bind-api/3.0.1/jakarta.xml.bind-api-3.0.1.jar",
    "jaxb-core-3.0.2.jar": f"{_BASE}/com/sun/xml/bind/jaxb-core/3.0.2/jaxb-core-3.0.2.jar",
    "jaxb-runtime-3.0.2.jar": f"{_BASE}/org/glassfish/jaxb/jaxb-runtime/3.0.2/jaxb-runtime-3.0.2.jar",
    "txw2-3.0.2.jar": f"{_BASE}/org/glassfish/jaxb/txw2/3.0.2/txw2-3.0.2.jar",
    "istack-commons-runtime-4.0.1.jar": f"{_BASE}/com/sun/istack/istack-commons-runtime/4.0.1/istack-commons-runtime-4.0.1.jar",
    "jgoodies-common-1.8.1.jar": f"{_BASE}/com/jgoodies/jgoodies-common/1.8.1/jgoodies-common-1.8.1.jar",
    "jgoodies-binding-2.13.0.jar": f"{_BASE}/com/jgoodies/jgoodies-binding/2.13.0/jgoodies-binding-2.13.0.jar",
}


def get_jar_cache_dir() -> Path:
    """JAR キャッシュディレクトリ（o2md パッケージ内 libs/）を取得

    pip でインストールされたパッケージフォルダ内に libs/ を作成する。
    """
    cache_dir = Path(__file__).parent / "libs"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir


def ensure_mpxj_jars(verbose: bool = False) -> str:
    """必要な MPXJ JAR をダウンロードし、クラスパスを返す

    o2md パッケージフォルダ内 libs/ に JAR をダウンロード・キャッシュする。
    2 回目以降は既存 JAR を再利用する。

    Args:
        verbose: 詳細ログを出力するか

    Returns:
        クラスパス文字列 (OSに応じた区切り文字で連結)

    Raises:
        RuntimeError: 必須 JAR のダウンロード失敗時
    """
    cache_dir = get_jar_cache_dir()

    # 不足している JAR を抽出
    missing = [
        (name, url)
        for name, url in MPXJ_DEPENDENCIES.items()
        if not (cache_dir / name).exists()
    ]

    # ダウンロード実行
    if missing:
        print(f"MS Project 変換用ライブラリを取得中... ({len(missing)} 個)")
        print(f"  保存先: {cache_dir}")
        failed = []
        for name, url in missing:
            jar_path = cache_dir / name
            try:
                if verbose:
                    print(f"  ダウンロード中: {name}")
                urllib.request.urlretrieve(url, str(jar_path))
            except Exception as e:
                failed.append(name)
                print(f"  警告: {name} の取得に失敗しました ({e})")
        if failed:
            raise RuntimeError(
                f"必要な JAR の取得に失敗しました: {', '.join(failed)}\n"
                f"ネットワーク接続を確認して再度実行してください。"
            )

    # クラスパスを構築
    jar_files = sorted(cache_dir.glob("*.jar"))
    if not jar_files:
        raise RuntimeError(
            f"JAR が見つかりません: {cache_dir}\n"
            f"ダウンロードに失敗した可能性があります。"
        )

    return os.pathsep.join(str(jar) for jar in jar_files)
