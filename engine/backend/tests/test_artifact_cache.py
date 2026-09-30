"""产物缓存单元测试：键敏感性 / upsert / LRU。全程无 Word。"""

import time

import pytest

from backend.api.artifact_cache import ArtifactCache, _assets_digest


@pytest.fixture
def env(tmp_path):
    """一套最小输入：md + 模板 docx + 档案 json + photo 资产 + 缓存目录"""
    md = tmp_path / "input" / "论文.md"
    md.parent.mkdir()
    md.write_text("# 标题\n正文", encoding="utf-8")
    (md.parent / "photo").mkdir()
    (md.parent / "photo" / "1.png").write_bytes(b"PNG0")
    tpl = tmp_path / "模板.docx"
    tpl.write_bytes(b"DOCX")
    tpl.with_suffix(".json").write_text('{"name":"t"}', encoding="utf-8")
    cache = ArtifactCache(tmp_path / "_cache", max_entries=3,
                          max_bytes=10 * 1024 * 1024)
    return md, tpl, cache


def test_key_sensitivity(env):
    """任一输入变化必须换键；同输入必须同键。"""
    md, tpl, cache = env
    base = cache.key("full", md, tpl, {"a": 1})

    assert cache.key("full", md, tpl, {"a": 1}) == base          # 稳定
    assert cache.key("fast", md, tpl, {"a": 1}) != base          # 模式
    assert cache.key("full", md, tpl, {"a": 2}) != base          # config

    md.write_text("# 标题\n改了", encoding="utf-8")
    changed_md = cache.key("full", md, tpl, {"a": 1})
    assert changed_md != base                                    # md 内容

    tpl.write_bytes(b"DOCX2")
    changed_tpl = cache.key("full", md, tpl, {"a": 1})
    assert changed_tpl != changed_md                             # 模板 docx

    tpl.with_suffix(".json").write_text('{"name":"t2"}', encoding="utf-8")
    assert cache.key("full", md, tpl, {"a": 1}) != changed_tpl   # 档案 json


def test_key_assets(env):
    """photo/formula 资产变化必须换键（换图后不能命中旧文档）。"""
    md, tpl, cache = env
    base = cache.key("full", md, tpl, None)
    photo = md.parent / "photo" / "1.png"
    photo.write_bytes(b"PNG00")            # 大小变
    assert cache.key("full", md, tpl, None) != base
    assert "photo/1.png" in _assets_digest(md.parent)


def test_put_get_upsert(env, tmp_path):
    """docx 先入、pdf 后补，pages/warnings 保留。"""
    md, tpl, cache = env
    key = cache.key("full", md, tpl, None)
    assert cache.get(key) is None

    docx = tmp_path / "a.docx"
    docx.write_bytes(b"D1")
    cache.put(key, title="论文A", docx=docx, pages=18, warnings=["w1"])

    hit = cache.get(key)
    assert hit is not None
    assert hit["title"] == "论文A" and hit["pages"] == 18
    assert hit["warnings"] == ["w1"]
    assert hit["docx"] is not None and hit["docx"].read_bytes() == b"D1"
    assert hit["pdf"] is None

    pdf = tmp_path / "a.pdf"
    pdf.write_bytes(b"P1")
    cache.put(key, title="论文A", pdf=pdf)   # upsert：不带 pages/warnings

    hit = cache.get(key)
    assert hit["pdf"] is not None and hit["pdf"].read_bytes() == b"P1"
    assert hit["pages"] == 18 and hit["warnings"] == ["w1"]     # 未被清掉


def test_lru_prune(env, tmp_path):
    """超出条目上限时按 last_used 淘汰最旧。"""
    md, tpl, cache = env                     # max_entries=3
    docx = tmp_path / "a.docx"
    docx.write_bytes(b"D")
    keys = []
    for i in range(4):
        k = cache.key("full", md, tpl, {"i": i})
        cache.put(k, title=f"t{i}", docx=docx, pages=1)
        keys.append(k)
        time.sleep(0.01)                     # 让 last_used 可排序
    assert cache.get(keys[0]) is None        # 最旧的被淘汰
    assert all(cache.get(k) is not None for k in keys[1:])
