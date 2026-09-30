"""Preview cache routing and same-key request coalescing; no Word required."""

import asyncio
import json
from types import SimpleNamespace

from backend.api import main, routes
from backend.api.artifact_cache import ArtifactCache
from backend.api.word_preview_worker import WordPreviewWorker
from backend.thesis_builder import converter


def test_desktop_template_pin_requires_explicit_change_confirmation(tmp_path, monkeypatch):
    project = tmp_path / "data"
    input_path = project / "input" / "1" / "paper.md"
    first = project / "template" / "first.docx"
    second = project / "template" / "managed" / "new" / "second.docx"
    input_path.parent.mkdir(parents=True)
    first.parent.mkdir(parents=True)
    second.parent.mkdir(parents=True)
    input_path.write_text("# paper", encoding="utf-8")
    first.write_bytes(b"FIRST")
    second.write_bytes(b"SECOND")
    monkeypatch.setattr(routes, "PROJECT_ROOT", project)
    monkeypatch.setattr(routes, "INPUT_DIR", project / "input")
    monkeypatch.setattr(routes, "OUTPUT_DIR", project / "output")
    monkeypatch.setattr(routes, "TEMPLATE_PINNING_REQUIRED", True)

    routes._enforce_template_pin(input_path, first, False)
    pin = project / "output" / "workspaces" / "1" / ".template-pin.json"
    assert json.loads(pin.read_text(encoding="utf-8"))["path"] == "template/first.docx"

    try:
        routes._enforce_template_pin(input_path, second, False)
    except Exception as exc:
        assert exc.status_code == 409
        assert exc.detail["code"] == "template_change_confirmation_required"
    else:
        raise AssertionError("template change was not blocked")

    routes._enforce_template_pin(input_path, second, True)
    assert json.loads(pin.read_text(encoding="utf-8"))["path"].endswith("second.docx")


def test_desktop_lists_only_controller_managed_templates(tmp_path, monkeypatch):
    project = tmp_path / "data"
    template_root = project / "template"
    managed = template_root / "managed" / "template-id"
    managed.mkdir(parents=True)
    bundled = template_root / "bundled.docx"
    cloud_docx = managed / "a.docx"
    bundled.write_bytes(b"BUNDLED")
    cloud_docx.write_bytes(b"CLOUD")
    cloud_docx.with_suffix(".json").write_text('{"name":"cloud"}', encoding="utf-8")
    (template_root / "managed" / "index.json").write_text(json.dumps({
        "templates": [{
            "id": "template-id",
            "school": "测试大学",
            "version": "2026.1",
            "path": "template/managed/template-id/a.docx",
        }]
    }), encoding="utf-8")
    monkeypatch.setattr(routes, "PROJECT_ROOT", project)
    monkeypatch.setattr(routes, "TEMPLATE_DIR", template_root)
    monkeypatch.setattr(routes, "TEMPLATE_PINNING_REQUIRED", True)

    result = asyncio.run(routes.api_list_templates())
    assert len(result["templates"]) == 1
    assert result["templates"][0]["managed"] is True
    assert result["templates"][0]["has_profile"] is True


def test_preview_hit_uses_content_address_and_coalesces(tmp_path, monkeypatch):
    md = tmp_path / "paper.md"
    md.write_text("# title", encoding="utf-8")
    template = tmp_path / "template.docx"
    template.write_bytes(b"DOCX")
    pdf = tmp_path / "preview.pdf"
    pdf.write_bytes(b"%PDF-test")
    cache = ArtifactCache(tmp_path / "cache")
    key = cache.key("fast", md, template, None)
    cache.put(key, title="paper", pdf=pdf, pages=7, warnings=["cached"])

    async def ensure_md(_path):
        return md, []

    monkeypatch.setattr(routes, "ARTIFACT_CACHE", cache)
    monkeypatch.setattr(routes, "_resolve_input", lambda _path: md)
    monkeypatch.setattr(routes, "_resolve_template", lambda _path: template)
    monkeypatch.setattr(routes, "_ensure_md_input", ensure_md)

    response = asyncio.run(routes.api_preview(routes.PreviewRequest(
        input_md="ignored.md", template_path="ignored.docx")))
    assert response.pdf_url == f"/api/cache/{key}/pdf"
    assert response.total_pages == 7
    assert response.warnings == ["cached"]

    # First lookup misses before the lock; the second lookup, after acquiring
    # the Word slot, finds the completed artifact. convert() must not run.
    original_get = cache.get
    calls = 0

    def delayed_hit(cache_key):
        nonlocal calls
        calls += 1
        return None if calls == 1 else original_get(cache_key)

    monkeypatch.setattr(cache, "get", delayed_hit)
    monkeypatch.setattr(routes, "convert", lambda **_kwargs: (_ for _ in ()).throw(
        AssertionError("same-key preview must not call convert twice")))
    response = asyncio.run(routes.api_preview(routes.PreviewRequest(
        input_md="ignored.md", template_path="ignored.docx")))
    assert calls == 2
    assert response.pdf_url == f"/api/cache/{key}/pdf"


def test_rewritten_preview_uses_separate_accurate_cache_key(tmp_path, monkeypatch):
    md = tmp_path / "paper.md"
    md.write_text("# rewritten", encoding="utf-8")
    template = tmp_path / "template.docx"
    template.write_bytes(b"DOCX")
    pdf = tmp_path / "accurate.pdf"
    pdf.write_bytes(b"%PDF-accurate")
    cache = ArtifactCache(tmp_path / "cache")
    accurate_key = cache.key("accurate", md, template, None)
    cache.put(accurate_key, title="paper", pdf=pdf, pages=9)

    async def ensure_md(_path):
        return md, []

    monkeypatch.setattr(routes, "ARTIFACT_CACHE", cache)
    monkeypatch.setattr(routes, "_resolve_input", lambda _path: md)
    monkeypatch.setattr(routes, "_resolve_template", lambda _path: template)
    monkeypatch.setattr(routes, "_ensure_md_input", ensure_md)

    response = asyncio.run(routes.api_preview(routes.PreviewRequest(
        input_md="ignored.md", template_path="ignored.docx", accurate_pagination=True)))
    assert response.pdf_url == f"/api/cache/{accurate_key}/pdf"
    assert response.total_pages == 9


def test_rewritten_preview_runs_full_pagination_before_pdf(tmp_path, monkeypatch):
    md = tmp_path / "paper.md"
    md.write_text("# rewritten", encoding="utf-8")
    template = tmp_path / "template.docx"
    template.write_bytes(b"DOCX")
    docx = tmp_path / "paper.docx"
    docx.write_bytes(b"DOCX")
    pdf = tmp_path / "paper.pdf"
    cache = ArtifactCache(tmp_path / "cache")
    calls = []

    async def ensure_md(_path):
        return md, []

    def restore(*_args):
        calls.append("source")
        return {"output_path": str(docx), "title": "paper", "pages": 5,
                "warnings": [], "section_parities": ["any", "odd"]}

    def finalize(source):
        calls.append("full-pagination")
        return {"success": True, "output_path": source["output_path"],
                "preview_url": "/api/preview/paper", "pages": 7, "warnings": []}

    def to_pdf(*_args, **kwargs):
        calls.append(("pdf", kwargs["update_fields"], kwargs["persist_preview_fields"]))
        pdf.write_bytes(b"%PDF-accurate")
        return pdf

    monkeypatch.setattr(routes, "ARTIFACT_CACHE", cache)
    monkeypatch.setattr(routes, "PREVIEW_DIR", tmp_path / "previews")
    monkeypatch.setattr(routes, "_resolve_input", lambda _path: md)
    monkeypatch.setattr(routes, "_resolve_template", lambda _path: template)
    monkeypatch.setattr(routes, "_ensure_md_input", ensure_md)
    monkeypatch.setattr(routes, "_restore_or_build_source", restore)
    monkeypatch.setattr(routes, "_finalize_source", finalize)
    monkeypatch.setattr(routes, "_convert_to_pdf", to_pdf)
    monkeypatch.setattr(routes, "_count_pdf_pages", lambda _path: 7)

    response = asyncio.run(routes.api_preview(routes.PreviewRequest(
        input_md="ignored.md", template_path="ignored.docx", accurate_pagination=True)))

    assert response.total_pages == 7
    assert calls == ["source", "full-pagination", ("pdf", False, False)]


def test_cached_pdf_endpoint_rejects_bad_key():
    try:
        asyncio.run(routes.api_get_cached_pdf("not-a-sha"))
    except routes.HTTPException as exc:
        assert exc.status_code == 404
    else:
        raise AssertionError("bad cache key must be rejected")


def test_source_docx_is_reused_before_final_pagination(tmp_path, monkeypatch):
    """Preview and export share only the immutable pre-pagination DOCX."""
    md = tmp_path / "paper.md"
    md.write_text("# title", encoding="utf-8")
    template = tmp_path / "template.docx"
    template.write_bytes(b"DOCX")
    cache = ArtifactCache(tmp_path / "cache")
    work = tmp_path / "work"
    work.mkdir()
    source_key = cache.key("source", md, template, None)
    calls = 0

    def fake_convert(**_kwargs):
        nonlocal calls
        calls += 1
        docx = work / "paper.docx"
        docx.write_bytes(b"raw-docx")
        return {
            "output_path": str(docx), "pages": 3, "warnings": ["w"],
            "section_parities": ["odd", "even"],
        }

    monkeypatch.setattr(routes, "ARTIFACT_CACHE", cache)
    monkeypatch.setattr(routes, "convert", fake_convert)

    first = routes._restore_or_build_source(
        source_key, md, template, None, work)
    (work / "paper.docx").unlink()  # prove the second call comes from cache
    second = routes._restore_or_build_source(
        source_key, md, template, None, work)

    assert calls == 1
    assert first["section_parities"] == ["odd", "even"]
    assert second["section_parities"] == ["odd", "even"]
    assert (work / "paper.docx").read_bytes() == b"raw-docx"


def test_word_preview_uses_screen_pdf_switch(tmp_path, monkeypatch):
    """Only the interactive preview reaches the persistent screen-PDF worker."""
    docx = tmp_path / "paper.docx"
    docx.write_bytes(b"DOCX")
    pdf = tmp_path / "paper.pdf"
    calls = []

    class FakeWorker:
        def export(self, input_docx, output_pdf):
            calls.append(("preview", input_docx, output_pdf))
            output_pdf.write_bytes(b"PDF")

        def export_final_pdf(self, input_docx, output_pdf):
            calls.append(("final", input_docx, output_pdf))
            output_pdf.write_bytes(b"PDF")

    def fake_run(command, **_kwargs):
        calls.append(command)
        pdf.write_bytes(b"PDF")
        return SimpleNamespace(returncode=0, stdout=b"", stderr=b"")

    monkeypatch.setattr(routes.subprocess, "run", fake_run)
    monkeypatch.setattr(routes, "WORD_PREVIEW_WORKER", FakeWorker())
    assert routes._convert_via_word(docx, pdf, update_fields=True, preview=True)
    assert calls[-1] == ("preview", docx, pdf)

    assert routes._convert_via_word(docx, pdf, preview=False)
    assert calls[-1] == ("final", docx, pdf)


def test_preview_can_persist_fields_only_through_worker(tmp_path, monkeypatch):
    docx = tmp_path / "paper.docx"
    docx.write_bytes(b"DOCX")
    pdf = tmp_path / "paper.pdf"
    calls = []

    class FakeWorker:
        def export(self, input_docx, output_pdf, *, persist_fields=False):
            calls.append((input_docx, output_pdf, persist_fields))
            output_pdf.write_bytes(b"PDF")

    monkeypatch.setattr(routes, "WORD_PREVIEW_WORKER", FakeWorker())
    details = {}
    assert routes._convert_via_word(
        docx, pdf, update_fields=True, preview=True,
        persist_preview_fields=True, conversion_details=details)
    assert calls == [(docx, pdf, True)]
    assert details == {"preview_fields_persisted": True}


def test_convert_reuses_only_matching_persisted_preview_docx(tmp_path, monkeypatch):
    md = tmp_path / "paper.md"
    md.write_text("# title", encoding="utf-8")
    template = tmp_path / "template.docx"
    template.write_bytes(b"DOCX")
    source_docx = tmp_path / "source.docx"
    source_docx.write_bytes(b"raw")
    preview_docx = tmp_path / "preview.docx"
    preview_docx.write_bytes(b"fields-updated")
    cache = ArtifactCache(tmp_path / "cache")
    source_key = cache.key("source", md, template, None)
    fast_key = cache.key("fast", md, template, None)
    cache.put(fast_key, title="paper", docx=preview_docx, extra={
        "preview_fields_persisted": True,
        "source_key": source_key,
        "section_parities": ["any", "odd"],
    })
    captured = {}

    async def ensure_md(_path):
        return md, []

    def restore(*_args):
        return {"output_path": str(source_docx), "title": "paper", "pages": 2,
                "warnings": [], "section_parities": ["any", "odd"]}

    def finalize(source):
        captured.update(source)
        return {"success": True, "output_path": source["output_path"],
                "preview_url": "/api/preview/paper", "pages": 2, "warnings": []}

    monkeypatch.setattr(routes, "ARTIFACT_CACHE", cache)
    monkeypatch.setattr(routes, "_resolve_input", lambda _path: md)
    monkeypatch.setattr(routes, "_resolve_template", lambda _path: template)
    monkeypatch.setattr(routes, "_ensure_md_input", ensure_md)
    monkeypatch.setattr(routes, "_restore_or_build_source", restore)
    monkeypatch.setattr(routes, "_finalize_source", finalize)
    response = asyncio.run(routes.api_convert(routes.ConvertRequest(input_md="ignored.md")))

    assert response.success
    assert captured["fields_already_updated"] is True
    assert source_docx.read_bytes() == b"fields-updated"


def test_preview_worker_failure_falls_back_to_one_shot_word(tmp_path, monkeypatch):
    docx = tmp_path / "paper.docx"
    docx.write_bytes(b"DOCX")
    pdf = tmp_path / "paper.pdf"
    calls = []

    class BrokenWorker:
        def export(self, *_args):
            raise RuntimeError("simulated worker failure")

    def fake_run(command, **_kwargs):
        calls.append(command)
        pdf.write_bytes(b"PDF")
        return SimpleNamespace(returncode=0, stdout=b"", stderr=b"")

    monkeypatch.setattr(routes, "WORD_PREVIEW_WORKER", BrokenWorker())
    monkeypatch.setattr(routes.subprocess, "run", fake_run)
    assert routes._convert_via_word(docx, pdf, update_fields=True, preview=True)
    assert "-Preview" in calls[-1]


def test_pdf_conversion_never_reuses_stale_same_title_file(tmp_path, monkeypatch):
    docx = tmp_path / "paper.docx"
    docx.write_bytes(b"DOCX")
    pdf = tmp_path / "paper.pdf"
    pdf.write_bytes(b"stale")

    def fake_word(_docx, expected_pdf, **_kwargs):
        assert not expected_pdf.exists()
        expected_pdf.write_bytes(b"fresh")
        return True

    monkeypatch.setattr(routes, "_convert_via_word", fake_word)
    assert routes._convert_to_pdf(docx, "paper", tmp_path) == pdf
    assert pdf.read_bytes() == b"fresh"


def test_desktop_word_requirement_never_falls_back_to_libreoffice(tmp_path, monkeypatch):
    docx = tmp_path / "paper.docx"
    docx.write_bytes(b"docx")
    attempted = []

    monkeypatch.setattr(routes, "REQUIRE_WORD", True)
    monkeypatch.setattr(routes, "_convert_via_word", lambda *_args, **_kwargs: False)
    monkeypatch.setattr(routes, "_find_libreoffice", lambda: attempted.append(True))

    assert routes._convert_to_pdf(docx, "paper", tmp_path) is None
    assert attempted == []


def test_uploaded_papers_get_separate_output_workspaces(tmp_path, monkeypatch):
    input_root = tmp_path / "input"
    output_root = tmp_path / "output"
    first = input_root / "1" / "paper.md"
    second = input_root / "2" / "paper.md"
    first.parent.mkdir(parents=True)
    second.parent.mkdir(parents=True)
    first.write_text("first", encoding="utf-8")
    second.write_text("second", encoding="utf-8")

    monkeypatch.setattr(routes, "INPUT_DIR", input_root)
    monkeypatch.setattr(routes, "OUTPUT_DIR", output_root)

    assert routes._workspace_output_dir(first) == output_root / "workspaces" / "1"
    assert routes._workspace_output_dir(second) == output_root / "workspaces" / "2"
    assert routes._workspace_output_dir(first) != routes._workspace_output_dir(second)


def test_persistent_worker_formats_word_pass_measurement(monkeypatch, tmp_path):
    worker = WordPreviewWorker()
    request = {}
    monkeypatch.setattr(worker, "_ensure_started", lambda: None)
    monkeypatch.setattr(worker, "_send", lambda payload: request.update(payload) or {
        "ok": True, "sections": {"2": 7, "1": 1}, "pages": 9,
    })

    result = worker.pass_docx(tmp_path / "paper.docx", update_fields=True,
                              update_page_numbers=False, measure=True)
    assert request == {
        "op": "pass", "input": str(tmp_path / "paper.docx"),
        "update_fields": True, "update_page_numbers": False, "measure": True,
    }
    assert result == "1:1\n2:7\nPAGES:9\nOK"


def test_finalization_skips_only_reused_preview_field_pass(monkeypatch, tmp_path):
    calls = []

    def fake_word_pass(_path, **kwargs):
        calls.append(kwargs)
        return "1:1\nPAGES:1\nOK"

    monkeypatch.setattr(converter, "_word_pass", fake_word_pass)
    assert converter.finalize_docx(
        tmp_path / "paper.docx", ["any"], [], fields_already_updated=True) == 1
    assert calls == [{"update_fields": False, "measure": True}]


def test_persistent_worker_warm_is_idempotent_entrypoint(monkeypatch):
    worker = WordPreviewWorker()
    calls = []
    monkeypatch.setattr(worker, "_ensure_started", lambda: calls.append("start"))
    worker.warm()
    worker.warm()
    assert calls == ["start", "start"]


def test_windows_startup_warms_worker_in_background(monkeypatch):
    started = []

    class FakeThread:
        def __init__(self, **kwargs):
            started.append(kwargs)

        def start(self):
            started[-1]["started"] = True

    monkeypatch.setattr(main.os, "name", "nt")
    monkeypatch.setattr(main.threading, "Thread", FakeThread)
    main.start_word_preview_worker()
    assert started == [{
        "target": main.warm_word_preview_worker,
        "name": "word-preview-warmup", "daemon": True, "started": True,
    }]


def test_lifespan_warms_and_closes_worker(monkeypatch):
    calls = []
    monkeypatch.setattr(main, "start_word_preview_worker", lambda: calls.append("start"))
    monkeypatch.setattr(main, "close_word_preview_worker", lambda: calls.append("close"))

    async def exercise_lifespan():
        async with main.lifespan(main.app):
            assert calls == ["start"]

    asyncio.run(exercise_lifespan())
    assert calls == ["start", "close"]


def test_converter_word_pass_uses_persistent_worker(tmp_path, monkeypatch):
    calls = []

    class FakeWorker:
        def pass_docx(self, path, **kwargs):
            calls.append((path, kwargs))
            return "1:1\nPAGES:1\nOK"

    from backend.api import word_preview_worker
    monkeypatch.setattr(word_preview_worker, "WORD_PREVIEW_WORKER", FakeWorker())
    path = tmp_path / "paper.docx"
    assert converter._word_pass(path, update_fields=True, measure=True) == "1:1\nPAGES:1\nOK"
    assert calls == [(path, {
        "update_fields": True, "update_page_numbers": False, "measure": True,
    })]
