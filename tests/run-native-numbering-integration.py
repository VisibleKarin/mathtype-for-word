"""Opt-in MathType native-number/chapter/section test on synthetic DOCX files.

Usage: python tests/run-native-numbering-integration.py --output-dir <new absolute folder>
Requires installed Word/MathType. Uses the bridge watchdog and leaves the source/global defaults untouched.
"""
from __future__ import annotations
import argparse
import json
import sys
import time
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import mathtype_batch
import mathtype_local
import mcp_server


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    mathtype_batch.refuse_concurrent_job()
    baseline = mcp_server._process_pids("WINWORD.EXE")
    bridge = mcp_server.SCRIPT_PATH
    try:
        mcp_server.SCRIPT_PATH = ROOT / "tests" / "run-native-numbering-integration.ps1"
        result = mcp_server._invoke_bridge("render", {"output_path": str(args.output_dir.resolve())}, timeout=240)
    finally:
        mcp_server.SCRIPT_PATH = bridge
    if not result.get("ok"):
        print(json.dumps(result, ensure_ascii=True))
        return 1
    packages = {}
    for sample in result["samples"]:
        package = mathtype_local.Package(sample["path"])
        tree = ET.fromstring(package.data[mathtype_local.MAIN])
        if tree.findall(".//{" + mathtype_local.W + "}numPr"):
            raise AssertionError("The sample contains ordinary Word list numbering.")
        packages[sample["name"]] = package
    before = [row["native_sha256"] for row in packages["按章和节编号"].objects]
    after = [row["native_sha256"] for row in packages["调整公式顺序后"].objects]
    if after != [before[1], before[0], *before[2:]]:
        raise AssertionError("Reordering changed native equation content.")
    deadline = time.monotonic() + 8
    while mcp_server._process_pids("WINWORD.EXE") != baseline:
        if time.monotonic() >= deadline:
            raise AssertionError("Word process baseline was not restored.")
        time.sleep(0.1)
    result["native_content_permutation_verified"] = True
    result["ordinary_word_list_numbering"] = False
    result["word_process_baseline_restored"] = True
    report = args.output_dir.resolve() / "mathtype-for-word-numbering-report.json"
    report.write_text(json.dumps(result, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps({"ok": True, "report": str(report), "samples": len(result["samples"])}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
