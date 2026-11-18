"""Opt-in local-equation integration on synthetic documents. Runs hidden Office sessions.

Usage: python tests/run-local-integration.py --work-dir <absolute output folder>
The folder must be new. It retains source/output documents and a JSON report for visual review.
"""
from __future__ import annotations
import argparse
import json
import subprocess
import sys
import time
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import mathtype_local as local
import mcp_server


def selector(row):
    return {k: row[k] for k in ("part", "index", "native_sha256")}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work-dir", required=True)
    args = parser.parse_args()
    folder = Path(args.work_dir).resolve()
    folder.mkdir(parents=True, exist_ok=False)
    baseline_word = mcp_server._process_pids("WINWORD.EXE")
    fixture = folder / "draft.docx"
    subprocess.run([mcp_server._powershell_executable(), "-NoLogo", "-NoProfile", "-NonInteractive",
                    "-ExecutionPolicy", "Bypass", "-File", str(ROOT / "tests/create-live-fixture.ps1"),
                    "-OutputPath", str(fixture)], check=True, timeout=90)
    manifest = ROOT / "tests/fixtures/live-manifest.json"
    rendered = folder / "rendered.docx"
    result = mcp_server._invoke_bridge("render", {"input_path": str(fixture), "output_path": str(rendered), "manifest_path": str(manifest)})
    (folder / "render.json").write_text(json.dumps(result, indent=2), encoding="utf-8")
    if not result.get("ok"):
        raise RuntimeError(result)
    # Put a soft-break explanation and a synthetic comment in the same physical paragraph as a numbered equation.
    source = folder / "混排 source.docx"
    p = local.Package(rendered)
    doc = p.docs[local.MAIN]
    obj = p.objects[0]["_node"]
    para = obj.parentNode.parentNode
    mark = doc.createElementNS(local.W, "w:commentRangeStart")
    mark.setAttributeNS(local.W, "w:id", "0")
    para.insertBefore(mark, para.firstChild)
    for name in ("br", "t"):
        run = doc.createElementNS(local.W, "w:r")
        child = doc.createElementNS(local.W, "w:" + name)
        if name == "t":
            child.appendChild(doc.createTextNode("说明，x"))
        run.appendChild(child)
        para.appendChild(run)
    run = doc.createElementNS(local.W, "w:r")
    props = doc.createElementNS(local.W, "w:rPr")
    superscript = doc.createElementNS(local.W, "w:vertAlign")
    superscript.setAttributeNS(local.W, "w:val", "superscript")
    props.appendChild(superscript)
    run.appendChild(props)
    text = doc.createElementNS(local.W, "w:t")
    text.appendChild(doc.createTextNode("2"))
    run.appendChild(text)
    para.appendChild(run)
    run = doc.createElementNS(local.W, "w:r")
    text = doc.createElementNS(local.W, "w:t")
    text.setAttribute("xml:space", "preserve")
    text.appendChild(doc.createTextNode(" stays ordinary Word text; preserve this paragraph."))
    run.appendChild(text)
    para.appendChild(run)
    end = doc.createElementNS(local.W, "w:commentRangeEnd")
    end.setAttributeNS(local.W, "w:id", "0")
    para.appendChild(end)
    run = doc.createElementNS(local.W, "w:r")
    reference = doc.createElementNS(local.W, "w:commentReference")
    reference.setAttributeNS(local.W, "w:id", "0")
    run.appendChild(reference)
    para.appendChild(run)
    comments = (f'<w:comments xmlns:w="{local.W}"><w:comment w:id="0" w:author="Synthetic test">'
                '<w:p><w:r><w:t>Preserve this synthetic comment.</w:t></w:r></w:p></w:comment></w:comments>')
    rels = local.xml(p.data["word/_rels/document.xml.rels"])
    rel = rels.createElementNS(local.REL, "Relationship")
    for k, v in {"Id": "localComment", "Type": local.R + "/comments", "Target": "comments.xml"}.items():
        rel.setAttribute(k, v)
    rels.documentElement.appendChild(rel)
    types = local.xml(p.data["[Content_Types].xml"])
    entry = types.createElementNS(types.documentElement.namespaceURI, "Override")
    entry.setAttribute("PartName", "/word/comments.xml")
    entry.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
    types.documentElement.appendChild(entry)
    updates = {local.MAIN: doc.toxml(encoding="utf-8"), "word/comments.xml": comments.encode("utf-8"),
               "word/_rels/document.xml.rels": rels.toxml(encoding="utf-8"), "[Content_Types].xml": types.toxml(encoding="utf-8")}
    with zipfile.ZipFile(source, "w") as z:
        for name, data in {**p.data, **updates}.items():
            z.writestr(name, data)
    original = local.Package(source)
    plan_path = folder / "size-plan.json"
    plan = {"schema_version": 1, "source_sha256": original.digest, "operations": [
        {"target": selector(row), "full_size_pt": size}
        for row, size in zip(original.objects[:4], (10, 10.5, 11, 12))]}
    plan_path.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    output = folder / "formatted.docx"
    result = local.format_equations(source, output, plan_path)
    if not result.get("ok"):
        raise RuntimeError(result)
    saved = local.Package(output)
    assert original.digest == local.sha(source.read_bytes())
    assert len(saved.objects) == 6
    for old, new in zip(original.objects[4:], saved.objects[4:]):
        assert old["native_sha256"] == new["native_sha256"]
    validation = mcp_server._invoke_bridge("validate", {"document_path": str(output), "manifest_path": str(manifest)})
    if not validation.get("ok"):
        raise RuntimeError(validation)
    # Repeat at the same sizes: no duplicate equations, paragraphs, numbers, bookmarks or references.
    plan["source_sha256"] = saved.digest
    for op, row in zip(plan["operations"], saved.objects):
        op["target"] = selector(row)
    plan_path = folder / "repeat-plan.json"
    plan_path.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    repeated = folder / "repeated.docx"
    repeat = local.format_equations(output, repeated, plan_path)
    if not repeat.get("ok"):
        raise RuntimeError(repeat)
    # Graft a real, taller fraction into the first mixed paragraph; retain the old native number and reference.
    plan = {"schema_version": 1, "source_sha256": original.digest, "donor_sha256": saved.digest,
            "operations": [{"target": selector(original.objects[0]), "donor": selector(saved.objects[2])}]}
    plan_path = folder / "replace-plan.json"
    plan_path.write_text(json.dumps(plan, indent=2), encoding="utf-8")
    replacement = local.replace_equations(source, folder / "replaced.docx", plan_path, output)
    # Word.Quit returns before the process finishes shutting down; allow a bounded exit grace period.
    deadline = time.monotonic() + 8
    while mcp_server._process_pids("WINWORD.EXE") != baseline_word:
        if time.monotonic() >= deadline:
            raise RuntimeError("Office process baseline was not restored.")
        time.sleep(0.1)
    report = {"ok": True, "format": result, "validation": validation, "repeat": repeat, "replacement": replacement,
              "environment": "Local installed Word/MathType; synthetic documents only"}
    (folder / "report.json").write_text(json.dumps(report, indent=2, ensure_ascii=True), encoding="utf-8")
    print(json.dumps({"ok": True, "report": str(folder / "report.json")}))


if __name__ == "__main__":
    main()
