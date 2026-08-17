"""MPXJ JAR管理の回帰テスト。"""

import hashlib

import pytest

from o2md import jar_manager


def test_jar_hash_mismatch_is_redownloaded_and_unexpected_jar_excluded(
    tmp_path, monkeypatch
):
    content = b"valid-jar"
    digest = hashlib.sha256(content).hexdigest()
    dependencies = {"required.jar": "https://example.test/required.jar"}
    monkeypatch.setattr(jar_manager, "MPXJ_DEPENDENCIES", dependencies)
    monkeypatch.setattr(jar_manager, "MPXJ_SHA256", {"required.jar": digest})
    monkeypatch.setattr(jar_manager, "get_jar_cache_dir", lambda: tmp_path)
    (tmp_path / "required.jar").write_bytes(b"old-jar")
    (tmp_path / "unexpected.jar").write_bytes(b"unexpected")
    monkeypatch.setattr(jar_manager, "_download", lambda _url: content)

    classpath = jar_manager.ensure_mpxj_jars()

    assert classpath == str(tmp_path / "required.jar")
    assert (tmp_path / "required.jar").read_bytes() == content
    assert "unexpected.jar" not in classpath


def test_downloaded_jar_hash_mismatch_is_removed(tmp_path, monkeypatch):
    dependencies = {"required.jar": "https://example.test/required.jar"}
    monkeypatch.setattr(jar_manager, "MPXJ_DEPENDENCIES", dependencies)
    monkeypatch.setattr(
        jar_manager,
        "MPXJ_SHA256",
        {"required.jar": hashlib.sha256(b"valid-jar").hexdigest()},
    )
    monkeypatch.setattr(jar_manager, "get_jar_cache_dir", lambda: tmp_path)
    monkeypatch.setattr(jar_manager, "_download", lambda _url: b"invalid-jar")

    with pytest.raises(RuntimeError, match="SHA-256検証"):
        jar_manager.ensure_mpxj_jars()

    assert not (tmp_path / "required.jar").exists()
