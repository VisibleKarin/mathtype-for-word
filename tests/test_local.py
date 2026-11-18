from __future__ import annotations

import json
import struct
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
import mathtype_local as local
import mcp_server


def compound(payload: bytes) -> bytes:
    """Small synthetic CFB mini-stream, independent of any private manuscript/vendor file."""
    native = struct.pack("<I", 28) + bytes(24) + b"\x05" + payload + bytes(6)
    sectors = (len(native) + 63) // 64
    if sectors > 8:
        raise ValueError("Fixture too large")
    header = bytearray(512)
    header[:8] = bytes.fromhex("d0cf11e0a1b11ae1")
    struct.pack_into("<HHHH", header, 24, 0x3e, 3, 0xfffe, 9)
    struct.pack_into("<H", header, 32, 6)
    for offset, value in ((44, 1), (48, 1), (56, 4096), (60, 3), (64, 1), (68, 0xfffffffe)):
        struct.pack_into("<I", header, offset, value)
    struct.pack_into("<109I", header, 76, 0, *([0xffffffff] * 108))
    fat = struct.pack("<128I", 0xfffffffd, 0xfffffffe, 0xfffffffe, 0xfffffffe, *([0xffffffff] * 124))

    def entry(name, kind, start, size):
        value = bytearray(128)
        encoded = (name + "\0").encode("utf-16le")
        value[:len(encoded)] = encoded
        struct.pack_into("<H", value, 64, len(encoded))
        value[66], value[67] = kind, 1
        struct.pack_into("<III", value, 68, 0xffffffff, 0xffffffff, 1 if kind == 5 else 0xffffffff)
        struct.pack_into("<IQ", value, 116, start, size)
        return value
    directory = entry("Root Entry", 5, 2, sectors * 64) + entry("Equation Native", 2, 0, len(native)) + bytes(256)
    mini_fat = struct.pack("<128I", *list(range(1, sectors)), 0xfffffffe, *([0xffffffff] * (128 - sectors)))
    return header + fat + directory + native.ljust(512, b"\0") + mini_fat


def document(path: Path, token=b"original", prefix="w", shared=False):
    p = prefix
    def obj(i, word):
        return (f'<{p}:r><{p}:rPr><{p}:i/><{p}:position {p}:val="-{i * 10}"/></{p}:rPr>'
                f'<{p}:object {p}:dxaOrig="100" {p}:dyaOrig="100">'
                f'<v:shape id="shape{i}" type="#type75" alt="preserved description" style="width:{word}pt;height:12pt">'
                f'<v:imagedata r:id="image{i}"/></v:shape>'
                f'<o:OLEObject ProgID="Equation.DSMT4" ShapeID="shape{i}" ObjectID="identity{i}" r:id="ole{i}"/>'
                f'</{p}:object></{p}:r>')
    raw = (f'<{p}:document xmlns:{p}="{local.W}" xmlns:r="{local.R}" xmlns:v="{local.V}" xmlns:o="{local.O}">'
           f'<{p}:body><{p}:p><{p}:pPr><{p}:pStyle {p}:val="Body"/></{p}:pPr>'
           f'<{p}:commentRangeStart {p}:id="7"/>{obj(1, 30)}'
           f'<{p}:bookmarkStart {p}:id="2" {p}:name="ZEqnNum123"/>'
           f'<{p}:fldSimple {p}:instr=" SEQ MTEqn "><{p}:r><{p}:t>5</{p}:t></{p}:r></{p}:fldSimple>'
           f'<{p}:bookmarkEnd {p}:id="2"/><{p}:r><{p}:br/><{p}:t>说明，x</{p}:t></{p}:r>'
           f'<{p}:r><{p}:rPr><{p}:vertAlign {p}:val="superscript"/></{p}:rPr><{p}:t>2</{p}:t></{p}:r>'
           f'<{p}:r><{p}:t> remains plain Word text.</{p}:t></{p}:r>'
           f'<{p}:commentRangeEnd {p}:id="7"/></{p}:p><{p}:p>{obj(2, 25)}</{p}:p>'
           f'<{p}:p><{p}:r><{p}:t>Intentional blank follows.</{p}:t></{p}:r></{p}:p><{p}:p/>'
           f'</{p}:body></{p}:document>')
    rels = f'<Relationships xmlns="{local.REL}">'
    for i in (1, 2):
        rels += f'<Relationship Id="ole{i}" Type="{local.R}/oleObject" Target="embeddings/object{i}.bin"/>'
        rels += f'<Relationship Id="image{i}" Type="{local.R}/image" Target="media/preview{i}.wmf"/>'
    if shared:
        rels += f'<Relationship Id="ordinaryImage" Type="{local.R}/image" Target="media/preview1.wmf"/>'
    rels += '</Relationships>'
    with zipfile.ZipFile(path, "w") as z:
        z.writestr(local.MAIN, raw)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/styles.xml", f'<w:styles xmlns:w="{local.W}"/>')
        z.writestr("word/comments.xml", f'<w:comments xmlns:w="{local.W}"><w:comment w:id="7"><w:p><w:r><w:t>Keep this comment</w:t></w:r></w:p></w:comment></w:comments>')
        z.writestr("docProps/core.xml", b"<metadata>synthetic</metadata>")
        for i in (1, 2):
            z.writestr(f"word/embeddings/object{i}.bin", compound(token + bytes([i])))
            z.writestr(f"word/media/preview{i}.wmf", token + b"preview" + bytes([i]))


class LocalEquationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source, self.donor, self.output = [self.folder / n for n in ("source.docx", "donor.docx", "result.docx")]
        document(self.source)
        document(self.donor, b"new", prefix="q")
        self.src = local.Package(self.source)
        self.dnr = local.Package(self.donor)
        self.plan = self.folder / "plan.json"

    def selector(self, row):
        return {k: row[k] for k in ("part", "index", "native_sha256")}

    def write_plan(self, operation):
        plan = {"schema_version": 1, "source_sha256": self.src.digest, "donor_sha256": self.dnr.digest,
                "operations": [operation]}
        self.plan.write_text(json.dumps(plan), encoding="utf-8")
        return plan

    def replacement(self):
        return {"target": self.selector(self.src.objects[0]), "donor": self.selector(self.dnr.objects[1])}

    def test_inventory_follows_relationships_and_native_data(self):
        report = local.inspect_document(self.source)
        self.assertEqual(len(report["equations"]), 2)
        self.assertEqual(report["source_sha256"], self.src.digest)
        self.assertEqual(report["equations"][0]["native_sha256"], local.sha(b"\x05original\x01" + bytes(6)))

    def test_local_replacement_preserves_every_other_part_and_mixed_paragraph(self):
        self.write_plan(self.replacement())
        before = self.source.read_bytes()
        report = local.replace_equations(self.source, self.output, self.plan, self.donor)
        self.assertTrue(report["validation"]["outside_target_xml_preserved"])
        saved = local.Package(self.output)
        allowed = {local.MAIN, "word/embeddings/object1.bin", "word/media/preview1.wmf"}
        self.assertEqual(set(report["changed_parts"]), allowed)
        for part in self.src.data.keys() - allowed:
            self.assertEqual(self.src.data[part], saved.data[part])
        self.assertEqual(saved.objects[0]["native_sha256"], self.dnr.objects[1]["native_sha256"])
        self.assertEqual(saved.objects[0]["object_id"], "identity1")
        self.assertIn(b'val="-20"', saved.data[local.MAIN])
        self.assertEqual(before, self.source.read_bytes())
        # Namespace inheritance differs between donor/target; resulting XML must be valid.
        ET.fromstring(saved.data[local.MAIN])

    def test_shared_preview_is_refused(self):
        document(self.source, shared=True)
        self.src = local.Package(self.source)
        self.write_plan(self.replacement())
        with self.assertRaisesRegex(ValueError, "Shared"):
            local.replace_equations(self.source, self.output, self.plan, self.donor)
        self.assertFalse(self.output.exists())

    def test_stale_target_or_source_never_publishes(self):
        op = self.replacement()
        op["target"]["native_sha256"] = "0" * 64
        self.write_plan(op)
        with self.assertRaisesRegex(ValueError, "Stale"):
            local.replace_equations(self.source, self.output, self.plan, self.donor)
        plan = self.write_plan(self.replacement())
        plan["source_sha256"] = "0" * 64
        self.plan.write_text(json.dumps(plan))
        with self.assertRaisesRegex(ValueError, "Source SHA"):
            local.replace_equations(self.source, self.output, self.plan, self.donor)
        self.assertFalse(self.output.exists())

    def test_duplicate_targets_and_overwrites_are_refused(self):
        op = self.replacement()
        plan = self.write_plan(op)
        plan["operations"].append(op)
        self.plan.write_text(json.dumps(plan))
        with self.assertRaisesRegex(ValueError, "Duplicate"):
            local.replace_equations(self.source, self.output, self.plan, self.donor)
        self.write_plan(op)
        self.output.write_bytes(b"user content")
        with self.assertRaisesRegex(ValueError, "already exists"):
            local.replace_equations(self.source, self.output, self.plan, self.donor)
        self.assertEqual(self.output.read_bytes(), b"user content")
        with self.assertRaisesRegex(ValueError, "distinct"):
            local.replace_equations(self.source, self.source, self.plan, self.donor, True)

    def fake_bridge(self, changed_math=None, size="10.5", extra=""):
        def run(action, args):
            self.assertEqual(action, "format-native")
            job = json.loads(Path(args["manifest_path"]).read_text())
            work = Path(job["work_dir"])
            document(Path(job["output_path"]), b"formatted")
            ml = '<math xmlns="http://www.w3.org/1998/Math/MathML"><msup><mi>x</mi><mn>2</mn></msup></math>'
            for op in job["operations"]:
                index = op["index"]
                (work / f"{index}-before.mml").write_text(ml, encoding="utf-8")
                (work / f"{index}-saved.mml").write_text(changed_math or ml, encoding="utf-8")
                (work / f"{index}-before.prefs").write_text("[Sizes]\nFull=12 pt\nScript=58 %\n", encoding="utf-8")
                (work / f"{index}-saved.prefs").write_text(f"[Sizes]\nFull={size} pt\nScript=58 %\n" + extra, encoding="utf-8")
            return {"ok": True}
        return run

    def test_format_requires_saved_size_and_unchanged_structure_preferences(self):
        self.write_plan({"target": self.selector(self.src.objects[0]), "full_size_pt": 10.5})
        report = local.format_equations(self.source, self.output, self.plan, bridge=self.fake_bridge())
        self.assertEqual(report["evidence"][0]["saved_full_size_pt"], 10.5)
        self.assertTrue(report["evidence"][0]["mathml_tree_equal"])
        self.assertEqual(local.Package(self.output).objects[1]["native_sha256"], self.src.objects[1]["native_sha256"])

    def test_append_and_same_text_different_structure_fail_transaction(self):
        self.write_plan({"target": self.selector(self.src.objects[0]), "full_size_pt": 10.5})
        for body in ('<msub><mi>x</mi><mn>2</mn></msub>',
                     '<msup><mi>x</mi><mn>2</mn></msup><msup><mi>x</mi><mn>2</mn></msup>'):
            with self.subTest(body=body), self.assertRaisesRegex(ValueError, "structure"):
                local.format_equations(self.source, self.output, self.plan,
                                       bridge=self.fake_bridge('<math xmlns="http://www.w3.org/1998/Math/MathML">' + body + '</math>'))
            self.assertFalse(self.output.exists())
        for bridge in (self.fake_bridge(size="12"), self.fake_bridge(extra="User1=50 %\n")):
            with self.assertRaisesRegex(ValueError, "readback"):
                local.format_equations(self.source, self.output, self.plan, bridge=bridge)
            self.assertFalse(self.output.exists())

    def test_invalid_size_and_office_failure_do_not_publish(self):
        for size in (True, float("nan"), -1, 0, 128, "10.5"):
            self.write_plan({"target": self.selector(self.src.objects[0]), "full_size_pt": size})
            with self.assertRaisesRegex(ValueError, "full_size_pt"):
                local.format_equations(self.source, self.output, self.plan, bridge=self.fake_bridge())
        self.write_plan({"target": self.selector(self.src.objects[0]), "full_size_pt": 10.5})
        result = local.format_equations(self.source, self.output, self.plan, bridge=lambda *args: {"ok": False, "error": "SDK unavailable"})
        self.assertEqual(result["error"], "SDK unavailable")
        self.assertFalse(self.output.exists())

    def test_malformed_cfb_cycles_and_truncation_fail(self):
        valid = compound(b"test")
        self.assertEqual(local.mtef(valid), b"\x05test" + bytes(6))
        cycle = bytearray(valid)
        struct.pack_into("<I", cycle, 512 + 4, 1)  # directory loops to itself
        for value in (b"x", valid[:900], bytes(cycle)):
            with self.assertRaises((ValueError, struct.error)):
                local.stream(value)

    def test_mathml_namespaces_annotations_and_matrix_dimensions(self):
        a = '<math xmlns="http://www.w3.org/1998/Math/MathML"><mi mathvariant="italic">x</mi></math>'
        b = '<m:math xmlns:m="http://www.w3.org/1998/Math/MathML"><m:mi mathvariant="italic" mathsize="10.5pt">x</m:mi></m:math>'
        self.assertEqual(local.math_tree(a), local.math_tree(b))
        self.assertNotEqual(local.math_tree(a), local.math_tree(a.replace("italic", "bold")))
        a = '<math><mtable><mtr><mtd><mi>x</mi></mtd><mtd><mi>y</mi></mtd></mtr></mtable></math>'
        b = '<math><mtable><mtr><mtd><mi>x</mi></mtd></mtr><mtr><mtd><mi>y</mi></mtd></mtr></mtable></math>'
        self.assertNotEqual(local.math_tree(a), local.math_tree(b))
        with self.assertRaisesRegex(ValueError, "DTD"):
            local.math_tree('<!DOCTYPE math [<!ENTITY x "bad">]><math>&x;</math>')

    def test_mcp_offline_inspect_does_not_install_or_launch_office(self):
        with patch.object(mcp_server, "_ensure_powershell", side_effect=AssertionError("offline inspection installed PowerShell")):
            result = mcp_server._call_tool("inspect_mathtype_equations", {"document_path": str(self.source)})
        data = json.loads(result["content"][0]["text"])
        self.assertTrue(data["ok"])
        self.assertEqual(len(data["equations"]), 2)


if __name__ == "__main__":
    unittest.main()
