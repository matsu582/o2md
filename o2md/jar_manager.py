#!/usr/bin/env python3
"""
MPXJ JAR 自動ダウンロード・管理モジュール

MS Project ファイル変換に必要な MPXJ 13.5.1 と依存 JAR を、
Maven Central から o2md パッケージのインストール済みフォルダ内
(o2md/libs/) にダウンロードして使用する。

Gradle/Maven コマンド実行に依存しない Pure Python 実装。
"""

import os
import hashlib
import stat
import time
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

# Maven Centralの.sha256がない成果物は.sha1照合後に算出した値を使用する。
MPXJ_SHA256 = {
    "mpxj-13.5.1.jar": "fe4af5b26359128ba9d9696e5830c7004285fdc8811a94f7699eb0b737a4ff94",
    "rtfparserkit-1.16.0.jar": "aa13625c4a9fc0234caf5e13a6cc3f44ce0d34c5541bb4b1d8ecaf52cb031237",
    "poi-5.3.0.jar": "d514ebff22327762d38f551b6d1d78bb764770afd8d37546387ca41790323fef",
    "commons-io-2.16.1.jar": "f41f7baacd716896447ace9758621f62c1c6b0a91d89acee488da26fc477c84f",
    "commons-codec-1.17.0.jar": "f700de80ac270d0344fdea7468201d8b9c805e5c648331c3619f2ee067ccfc59",
    "commons-collections4-4.4.jar": "1df8b9430b5c8ed143d7815e403e33ef5371b2400aadbe9bda0883762e0846d1",
    "commons-lang3-3.10.jar": "28968ae55fff465494083aeba856f8824c34902329882bf61e77246a91e25aa9",
    "commons-logging-1.2.jar": "daddea1ea0be0f56978ab3006b8ac92834afeefbd9b7e4e6316fca57df0fa636",
    "commons-math3-3.6.1.jar": "1e56d7b058d28b65abd256b8458e3885b674c1d588fa43cd7d1cbb9c7ef2b308",
    "log4j-api-2.23.1.jar": "92ec1fd36ab3bc09de6198d2d7c0914685c0f7127ea931acc32fd2ecdd82ea89",
    "log4j-core-2.23.1.jar": "7079368005fc34f56248f57f8a8a53361c3a53e9007d556dbc66fc669df081b5",
    "jsoup-1.15.3.jar": "e20a5e78b1372f2a4e620832db4442d5077e5cbde280b24c666a3770844999bc",
    "jackcess-4.0.1.jar": "8d2eecb226c6f2ece3d44a96d688e2c03656af557b585804a90713fb67fbd95e",
    "sqlite-jdbc-3.42.0.0.jar": "53174d76087bb73cc29db9c02766fb921fd7fc652f7952f3609e0018e3dd5ded",
    "SparseBitSet-1.3.jar": "f76b85adb0c00721ae267b7cfde4da7f71d3121cc2160c9fc00c0c89f8c53c8a",
    "jakarta.activation-2.0.1.jar": "b9e24b7dd6e07495562ea96531be3130c96dba4d78e1dfd88adbbdebf4332871",
    "jakarta.xml.bind-api-3.0.1.jar": "b8fb4bee3ff5b5c1ef77144d8411316018d7bbd41fcf1ede0646f7978546b867",
    "jaxb-core-3.0.2.jar": "9beb3f846ac998d619b8aab9ac498d30a8f3a01632ed8791e4aee8d0ba02fac9",
    "jaxb-runtime-3.0.2.jar": "7b3a3784b9c6e343a8d38fc108602ba810177f03f2d8d1e7258b1cef7fd9d4c7",
    "txw2-3.0.2.jar": "b4bcf94fb0a759456e2521724513baec94b78e93127544af162e3cff08d93343",
    "istack-commons-runtime-4.0.1.jar": "9f91115f449384886f572bd62c8812ee1004273d4b5c85cac65179ad4c16990f",
    "jgoodies-common-1.8.1.jar": "ddca10c16e1dc7a1b399c14580f0aae23014851e57d224cb96c260e6d649d2ad",
    "jgoodies-binding-2.13.0.jar": "83c4db194424416f1698f2a290f8b09592738dfc5cf5ceb1cc3a5154d2ada256",
}

_MAX_RETRIES = 3


def _download(url: str) -> bytes:
    """Maven Centralからデータを取得し、429を待機して再試行する。"""
    for attempt in range(_MAX_RETRIES):
        try:
            with urllib.request.urlopen(url, timeout=60) as response:
                return response.read()
        except urllib.error.HTTPError as error:
            if error.code != 429 or attempt == _MAX_RETRIES - 1:
                raise
            wait = int(error.headers.get("Retry-After", "1"))
            time.sleep(max(wait, 1) * (attempt + 1))
    raise RuntimeError(f"取得に失敗しました: {url}")


def _verify_jar(path: Path, expected_sha256: str) -> bool:
    """JARのSHA-256を検証する。"""
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    return digest == expected_sha256


def _is_safe_cache_dir(path: Path) -> bool:
    """キャッシュ先が他ユーザから書き換えられない実ディレクトリか判定する。"""
    if path.is_symlink() or not path.is_dir():
        return False
    info = path.stat()
    if info.st_mode & (stat.S_IWGRP | stat.S_IWOTH):
        return False
    if hasattr(os, "geteuid") and info.st_uid != os.geteuid():
        return False
    return True


def get_jar_cache_dir() -> Path:
    """JAR キャッシュディレクトリを取得

    既定は o2md パッケージフォルダ内 libs/ で、所有者のみ書き込み可として作成する。
    パッケージフォルダが他ユーザ書き込み可などで安全に使えない場合は、
    ユーザ専用のキャッシュ（$XDG_CACHE_HOME もしくは ~/.cache/o2md/libs）へ退避する。
    """
    package_dir = Path(__file__).parent / "libs"
    try:
        package_dir.mkdir(parents=True, exist_ok=True, mode=0o755)
        if _is_safe_cache_dir(package_dir):
            return package_dir
    except OSError:
        pass

    base = os.environ.get("XDG_CACHE_HOME") or str(Path.home() / ".cache")
    user_dir = Path(base) / "o2md" / "libs"
    user_dir.mkdir(parents=True, exist_ok=True, mode=0o700)
    if not _is_safe_cache_dir(user_dir):
        raise RuntimeError(f"JARキャッシュディレクトリを安全に用意できません: {user_dir}")
    return user_dir


def ensure_mpxj_jars(verbose: bool = False) -> str:
    """必要な MPXJ JAR をダウンロードし、クラスパスを返す

    o2md パッケージフォルダ内 libs/ に JAR をダウンロード・キャッシュする。
    2 回目以降はSHA-256検証済みの既存JARを再利用する。
    CLASSPATH環境変数が指定されている場合は、この関数を呼ばず、
    利用者が用意したJARを優先する。

    Args:
        verbose: 詳細ログを出力するか

    Returns:
        クラスパス文字列 (OSに応じた区切り文字で連結)

    Raises:
        RuntimeError: 必須 JAR のダウンロード失敗時
    """
    cache_dir = get_jar_cache_dir()

    for name, url in MPXJ_DEPENDENCIES.items():
        expected = MPXJ_SHA256[name]
        jar_path = cache_dir / name
        if jar_path.exists() and _verify_jar(jar_path, expected):
            continue
        if jar_path.exists():
            print(f"警告: {name} のハッシュが不一致のため再取得します")
            jar_path.unlink()
        try:
            if verbose:
                print(f"ダウンロード中: {name}")
            jar_path.write_bytes(_download(url))
        except Exception as error:
            if jar_path.exists():
                jar_path.unlink()
            raise RuntimeError(
                f"{name} の取得に失敗しました: {error}"
            ) from error
        if not _verify_jar(jar_path, expected):
            jar_path.unlink()
            raise RuntimeError(f"{name} のSHA-256検証に失敗しました")

    jar_files = [
        cache_dir / name
        for name in MPXJ_DEPENDENCIES
        if (cache_dir / name).exists()
        and _verify_jar(cache_dir / name, MPXJ_SHA256[name])
    ]
    if len(jar_files) != len(MPXJ_DEPENDENCIES):
        raise RuntimeError(
            f"必要なJARが揃っていないか検証に失敗しました: {cache_dir}"
        )

    return os.pathsep.join(str(jar) for jar in jar_files)
