"""Targeted DOCX equation edits. Standard library only; never resave the source in Word.

The Office worker operates on a disposable carrier. Only the selected OLE,
matching preview, geometry and baseline are grafted back into the source package.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
import os
import posixpath
import re
import struct
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom as DOM
from xml.etree import ElementTree as ET
from urllib.parse import unquote

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O = "urn:schemas-microsoft-com:office:office"
V = "urn:schemas-microsoft-com:vml"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN = "word/document.xml"


def sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def xml(data: bytes):
    if re.search(br"<!\s*(?:DOCTYPE|ENTITY)\b", data, re.I):
        raise ValueError("DTD/entity declarations are not supported in Office XML or MathML.")
    return DOM.parseString(data)


def elements(node, namespace: str, local: str):
    return list(node.getElementsByTagNameNS(namespace, local))


def stream(data: bytes, wanted: str = "Equation Native") -> bytes:
    """Read a bounded CFB v3/v4 stream, including mini streams. Reject cycles/truncation."""
    if len(data) < 512 or data[:8] != bytes.fromhex("d0cf11e0a1b11ae1"):
        raise ValueError("Not an OLE compound file.")
    u16 = lambda at: struct.unpack_from("<H", data, at)[0]
    u32 = lambda at: struct.unpack_from("<I", data, at)[0]
    version = u16(26)
    shift = u16(30)
    if u16(28) != 0xFFFE or (version, shift) not in ((3, 9), (4, 12)) or u16(32) != 6:
        raise ValueError("Unsupported compound-file header.")
    size, mini_size = 1 << shift, 64
    limit = len(data) // size - 1
    if limit < 1:
        raise ValueError("Truncated compound file.")

    def sector(index):
        if not 0 <= index < limit:
            raise ValueError("Compound-file sector outside file.")
        return data[(index + 1) * size:(index + 2) * size]

    def chain(start, table):
        seen = set()
        while start != 0xFFFFFFFE:
            if start in seen or not 0 <= start < len(table) or len(seen) > limit * (size // mini_size):
                raise ValueError("Cyclic or invalid compound-file chain.")
            seen.add(start)
            yield start
            start = table[start]

    fat_count = u32(44)
    if fat_count > limit or u32(72) > limit:
        raise ValueError("Invalid compound-file allocation count.")
    difat = list(struct.unpack_from("<109I", data, 76))
    nxt, seen = u32(68), set()
    for _ in range(u32(72)):
        if nxt in seen:
            raise ValueError("Cyclic DIFAT.")
        seen.add(nxt)
        values = struct.unpack("<" + "I" * (size // 4), sector(nxt))
        difat.extend(values[:-1])
        nxt = values[-1]
    fat = []
    for index in difat[:fat_count]:
        fat.extend(struct.unpack("<" + "I" * (size // 4), sector(index)))

    def big(start):
        return b"".join(sector(i) for i in chain(start, fat))

    directory = big(u32(48))
    entries = []
    for at in range(0, len(directory), 128):
        entry = directory[at:at + 128]
        n = struct.unpack_from("<H", entry, 64)[0]
        kind = entry[66]
        if not kind:
            continue
        if n < 2 or n > 64 or n % 2:
            raise ValueError("Invalid compound-file directory name.")
        name = entry[:n - 2].decode("utf-16le")
        start = struct.unpack_from("<I", entry, 116)[0]
        length = struct.unpack_from("<Q" if version == 4 else "<I", entry, 120)[0]
        if length > len(data):
            raise ValueError("Compound-file stream exceeds file size.")
        entries.append((name, kind, start, length))
    matches = [e for e in entries if e[0] == wanted and e[1] == 2]
    if len(matches) != 1:
        raise ValueError("Expected one Equation Native stream.")
    _, _, start, length = matches[0]
    if length == 0:
        return b""
    if length >= u32(56):
        result = big(start)
    else:
        roots = [e for e in entries if e[1] == 5]
        if len(roots) != 1:
            raise ValueError("Expected one compound-file root.")
        mini = big(roots[0][2])[:roots[0][3]]
        raw = big(u32(60))
        mini_fat = struct.unpack("<" + "I" * (len(raw) // 4), raw)
        pieces = []
        for i in chain(start, mini_fat):
            if (i + 1) * mini_size > len(mini):
                raise ValueError("Mini stream outside root stream.")
            pieces.append(mini[i * mini_size:(i + 1) * mini_size])
        result = b"".join(pieces)
    if len(result) < length:
        raise ValueError("Truncated compound-file stream.")
    return result[:length]


def mtef(data: bytes) -> bytes:
    native = stream(data)
    if len(native) < 34 or struct.unpack_from("<I", native)[0] != 28 or native[28] != 5:
        raise ValueError("Only Equation Native with a 28-byte header and MTEF 5 is supported.")
    return native[28:]


def math_tree(value: str):
    """Compare structure, tokens, roles and matrix shape; ignore only size/presentation metadata.

    Intentionally conservative: mrow/msub/msup/mfrac and mathvariant remain significant.
    A formatting-only round trip that changes any of them must fail, even if its text matches.
    """
    doc = xml(value.encode("utf-8"))
    root = doc.documentElement
    if root.localName != "math":
        raise ValueError("Expected a MathML math root.")
    ignored_attributes = {"mathsize", "fontsize", "fontfamily", "fontname", "color", "mathcolor"}

    def walk(node):
        tag = node.localName
        if tag in ("annotation", "annotation-xml"):
            return None
        attrs = tuple(sorted((a.namespaceURI or "", a.localName or a.name, a.value)
                             for a in node.attributes.values()
                             if a.namespaceURI != "http://www.w3.org/2000/xmlns/"
                             and (a.localName or a.name) not in ignored_attributes))
        children = []
        for child in node.childNodes:
            if child.nodeType == child.ELEMENT_NODE:
                item = walk(child)
                if item is not None:
                    children.append(item)
            elif child.nodeType in (child.TEXT_NODE, child.CDATA_SECTION_NODE):
                # Token whitespace is significant (especially mtext); indentation is not.
                if tag in ("mi", "mn", "mo", "mtext", "ms") or child.data.strip():
                    children.append(child.data.replace("\r\n", "\n"))
        if tag == "semantics" and len(children) == 1 and not attrs:
            return children[0]
        return (node.namespaceURI or "", tag, attrs, tuple(children))

    return walk(root)


class Package:
    def __init__(self, path):
        self.path = Path(path).resolve()
        if self.path.suffix.lower() != ".docx":
            raise ValueError("Local equation operations require a .docx file.")
        self.digest = sha(self.path.read_bytes())
        with zipfile.ZipFile(self.path) as z:
            self.infos = z.infolist()
            if len({i.filename for i in self.infos}) != len(self.infos):
                raise ValueError("Duplicate ZIP entries are not supported.")
            self.data = {i.filename: z.read(i) for i in self.infos}
        if any(n.startswith("_xmlsignatures/") for n in self.data):
            raise ValueError("Editing a digitally signed package requires a separate signing workflow.")
        self.docs = {}
        self.objects = []
        self.owners = {}
        for rel_path, raw in self.data.items():
            if not rel_path.endswith(".rels"):
                continue
            folder, name = posixpath.split(rel_path)
            base = posixpath.dirname(folder)
            owner = posixpath.join(base, name[:-5])
            owner_tree = ET.fromstring(self.data[owner]) if owner in self.data else None
            for rel in ET.fromstring(raw):
                if rel.get("TargetMode") == "External":
                    continue
                target = unquote(rel.get("Target", ""))
                resolved = posixpath.normpath(posixpath.join(base, target)) if not target.startswith("/") else target[1:]
                uses = sum(v == rel.get("Id") for n in owner_tree.iter() for k, v in n.attrib.items()
                           if k.startswith("{" + R + "}")) if owner_tree is not None else 1
                self.owners.setdefault(resolved, []).append((owner, rel.get("Id"), uses))
        for part in sorted(self.data):
            if not part.startswith("word/") or not part.endswith(".xml") or "/_rels/" in part:
                continue
            if b"Equation.DSMT4" not in self.data[part]:
                continue
            doc = xml(self.data[part])
            self.docs[part] = doc
            rel_path = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
            rels = ET.fromstring(self.data[rel_path])
            mapping = {e.attrib["Id"]: e.attrib for e in rels}

            def related(rid, kind):
                rel = mapping.get(rid, {})
                target = unquote(rel.get("Target", ""))
                if rel.get("TargetMode") == "External" or not rel.get("Type", "").endswith("/" + kind):
                    raise ValueError("Expected an internal " + kind + " relationship.")
                name = posixpath.normpath(posixpath.join(posixpath.dirname(part), target)) if not target.startswith("/") else target[1:]
                if name not in self.data or name.startswith("../"):
                    raise ValueError("Missing or invalid equation relationship: " + name)
                return name

            index = 0
            for ole in elements(doc, O, "OLEObject"):
                if ole.getAttribute("ProgID") != "Equation.DSMT4":
                    continue
                index += 1
                obj = ole.parentNode
                while obj is not None and not (obj.namespaceURI == W and obj.localName == "object"):
                    obj = obj.parentNode
                if obj is None:
                    raise ValueError("Unsupported MathType object container.")
                previews = elements(obj, V, "imagedata")
                embedding = related(ole.getAttributeNS(R, "id"), "oleObject")
                preview = related(previews[0].getAttributeNS(R, "id"), "image") if len(previews) == 1 else None
                payload = mtef(self.data[embedding])
                self.objects.append({"part": part, "index": index, "native_sha256": sha(payload),
                                     "embedding": embedding, "preview": preview,
                                     "object_id": ole.getAttribute("ObjectID"), "_node": obj})

    def rows(self):
        return [{k: v for k, v in row.items() if not k.startswith("_")} for row in self.objects]

    def target(self, selector):
        part = selector.get("part", MAIN)
        index = selector.get("index")
        if type(index) is not int or index < 1:
            raise ValueError("A target needs a positive integer index.")
        matches = [r for r in self.objects if r["part"] == part and r["index"] == index]
        if len(matches) != 1:
            raise ValueError(f"Equation target not found: {part} #{index}.")
        row = matches[0]
        if selector.get("native_sha256") != row["native_sha256"]:
            raise ValueError(f"Stale or missing native_sha256 for {part} #{index}; inspect the source again.")
        return row


def inspect_document(document_path):
    p = Package(document_path)
    return {"ok": True, "document_path": str(p.path), "source_sha256": p.digest,
            "equations": p.rows(), "scope": "DOCX package; no Office session or internal Full-size readback"}


def read_plan(path, source):
    plan = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    if not isinstance(plan, dict) or type(plan.get("schema_version")) is not int or plan["schema_version"] != 1:
        raise ValueError("Local edit plan needs schema_version: 1.")
    if plan.get("source_sha256") != source.digest:
        raise ValueError("Source SHA-256 changed; inspect and review the edit plan again.")
    operations = plan.get("operations")
    if not isinstance(operations, list) or not operations:
        raise ValueError("Local edit plan needs a nonempty operations array.")
    seen = set()
    for op in operations:
        if not isinstance(op, dict) or not isinstance(op.get("target"), dict):
            raise ValueError("Each operation needs a target object.")
        row = source.target(op["target"])
        key = (row["part"], row["index"])
        if key in seen:
            raise ValueError("Duplicate local edit target.")
        seen.add(key)
    return plan


def _position(run):
    props = [n for n in run.childNodes if n.nodeType == n.ELEMENT_NODE and n.namespaceURI == W and n.localName == "rPr"]
    return elements(props[0], W, "position") if props else []


def _import(doc, node):
    # minidom preserves namespace prefixes but importNode does not copy declarations inherited from ancestors.
    bindings = {}
    ancestor = node
    while ancestor is not None and ancestor.nodeType == ancestor.ELEMENT_NODE:
        for attr in ancestor.attributes.values():
            if attr.name == "xmlns" or attr.name.startswith("xmlns:"):
                bindings.setdefault(attr.name, attr.value)
        ancestor = ancestor.parentNode
    clone = doc.importNode(node, True)
    for name, value in bindings.items():
        clone.setAttributeNS("http://www.w3.org/2000/xmlns/", name, value)
    return clone


def _graft(source, target, donor, replacement, changes):
    old, new = target["_node"], replacement["_node"]
    for package, row in ((source, target), (donor, replacement)):
        for name in (row["embedding"], row["preview"]):
            if name is None:
                raise ValueError("A local edit requires one VML preview per equation.")
            # A shared OLE/preview must not change another equation, image, or alternate-content fallback.
            owners = package.owners.get(name, [])
            if len(owners) != 1 or owners[0][2] != 1:
                raise ValueError("Shared equation embedding/preview cannot be edited independently.")
        if row["preview"].rsplit(".", 1)[-1].lower() != target["preview"].rsplit(".", 1)[-1].lower():
            raise ValueError("Preview formats must match; content-type changes are not implicit.")
    run, donor_run = old.parentNode, new.parentNode
    if run.namespaceURI != W or run.localName != "r" or donor_run.namespaceURI != W or donor_run.localName != "r":
        raise ValueError("Local edits require inline MathType objects in Word runs.")
    if any(n.nodeType == n.ELEMENT_NODE and not (n.namespaceURI == W and n.localName in ("rPr", "object")) for n in run.childNodes):
        raise ValueError("The target object shares a run with text/breaks; isolate that run before editing.")
    if len(elements(run, W, "object")) != 1:
        raise ValueError("The target run must contain exactly one object.")
    old_shapes, new_shapes = elements(old, V, "shape"), elements(new, V, "shape")
    if len(old_shapes) != 1 or len(new_shapes) != 1:
        raise ValueError("Local edits support a single VML equation shape.")
    # Keep target relationships, ObjectID, ShapeID, alt text and numbering. Refresh the complete geometry.
    shape = old_shapes[0]
    fresh = _import(source.docs[target["part"]], new_shapes[0])
    for attr in ("id", "type", "alt", "title"):
        if shape.hasAttribute(attr):
            fresh.setAttribute(attr, shape.getAttribute(attr))
    image = elements(fresh, V, "imagedata")[0]
    old_image = elements(old, V, "imagedata")[0]
    image.setAttributeNS(R, old_image.getAttributeNodeNS(R, "id").name, old_image.getAttributeNS(R, "id"))
    image.setAttributeNS("http://www.w3.org/2000/xmlns/", "xmlns:" + old_image.getAttributeNodeNS(R, "id").prefix, R)
    shape.parentNode.replaceChild(fresh, shape)
    for attr in ("dxaOrig", "dyaOrig"):
        old.removeAttributeNS(W, attr) if old.hasAttributeNS(W, attr) else None
        if new.hasAttributeNS(W, attr):
            old.setAttributeNS(W, old.prefix + ":" + attr, new.getAttributeNS(W, attr))
    for pos in _position(run):
        pos.parentNode.removeChild(pos)
    positions = _position(donor_run)
    if len(positions) > 1:
        raise ValueError("Ambiguous donor baseline.")
    if positions:
        props = [n for n in run.childNodes if n.nodeType == n.ELEMENT_NODE and n.namespaceURI == W and n.localName == "rPr"]
        if not props:
            props = [source.docs[target["part"]].createElementNS(W, run.prefix + ":rPr")]
            run.insertBefore(props[0], run.firstChild)
        props[0].appendChild(_import(source.docs[target["part"]], positions[0]))
    for field in ("embedding", "preview"):
        name = target[field]
        if name in changes:
            raise ValueError("Two operations would overwrite the same package part.")
        changes[name] = donor.data[replacement[field]]


def _outside_targets(raw, indices):
    root = ET.fromstring(raw)
    parents = {child: node for node in root.iter() for child in node}
    index = 0
    for ole in list(root.iter("{" + O + "}OLEObject")):
        if ole.get("ProgID") != "Equation.DSMT4":
            continue
        index += 1
        if index not in indices:
            continue
        obj = ole
        while obj.tag != "{" + W + "}object":
            obj = parents[obj]
        run = parents[obj]
        marker = ET.Element("local-edit-target", {"index": str(index)})
        marker.tail = obj.tail
        run.insert(list(run).index(obj), marker)
        run.remove(obj)
        props = run.find("{" + W + "}rPr")
        if props is not None:
            for pos in list(props.findall("{" + W + "}position")):
                props.remove(pos)
            if not len(props) and not props.attrib and not (props.text or "").strip():
                run.remove(props)

    def tree(node):
        return node.tag, sorted(node.attrib.items()), node.text, node.tail, [tree(c) for c in node]
    return tree(root)


def _output_path(source, output_path, overwrite):
    output = Path(output_path).resolve()
    if output == source.path or output.suffix.lower() != ".docx":
        raise ValueError("Output must be a distinct .docx file.")
    if output.exists() and (os.path.samefile(source.path, output) or not overwrite):
        raise ValueError("Output already exists or aliases the source; use a distinct path or overwrite.")
    return output


def _publish(source, output_path, pairs, overwrite=False, evidence=None):
    output = _output_path(source, output_path, overwrite)
    changes = {}
    changed_xml = set()
    for target, donor, replacement in pairs:
        _graft(source, target, donor, replacement, changes)
        changed_xml.add(target["part"])
    for part in changed_xml:
        changes[part] = source.docs[part].toxml(encoding="utf-8")
    output.parent.mkdir(parents=True, exist_ok=True)
    temp = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix=".docx", delete=False) as f:
            temp = Path(f.name)
        with zipfile.ZipFile(temp, "w") as z:
            for info in source.infos:
                z.writestr(copy.copy(info), changes.get(info.filename, source.data[info.filename]))
        check = Package(temp)
        if set(check.data) != set(source.data) or len(check.objects) != len(source.objects):
            raise ValueError("Saved package lost parts or equations.")
        actual_changes = [n for n in source.data if source.data[n] != check.data[n]]
        if not set(actual_changes) <= set(changes):
            raise ValueError("Saved package changed an unselected part.")
        for part in changed_xml:
            indices = {t["index"] for t, _, _ in pairs if t["part"] == part}
            if _outside_targets(source.data[part], indices) != _outside_targets(check.data[part], indices):
                raise ValueError("Saved XML changed text, fields, anchors or layout outside selected objects/baselines.")
        for target, donor, replacement in pairs:
            row = next(r for r in check.objects if (r["part"], r["index"]) == (target["part"], target["index"]))
            if row["native_sha256"] != replacement["native_sha256"]:
                raise ValueError("Saved native content differs from the verified donor.")
        if sha(source.path.read_bytes()) != source.digest:
            raise ValueError("Source changed during processing; output was not published.")
        if overwrite:
            os.replace(temp, output)
        else:
            # Windows rename refuses an existing destination, including one created during processing.
            if os.name == "nt":
                os.rename(temp, output)
            else:
                os.link(temp, output)
                temp.unlink()
        temp = None
        return {"ok": True, "output_path": str(output), "source_sha256": source.digest,
                "output_sha256": check.digest, "changed_parts": sorted(actual_changes),
                "preserved_parts": len(source.data) - len(actual_changes),
                "equations_changed": len(pairs), "evidence": evidence or [],
                "validation": {"ok": True, "unselected_parts_byte_identical": True,
                               "saved_native_matches_donor": True, "outside_target_xml_preserved": True},
                "limitations": ["Local package validation; visual layout still requires review.",
                                "Number/reference fields, styles, comments and paragraph layout are preserved; fields are not refreshed."]}
    finally:
        if temp is not None:
            temp.unlink(missing_ok=True)


def replace_equations(input_path, output_path, manifest_path, donor_path, overwrite=False):
    source, donor = Package(input_path), Package(donor_path)
    if Path(output_path).resolve() == donor.path:
        raise ValueError("Output must differ from the donor document.")
    plan = read_plan(manifest_path, source)
    if plan.get("donor_sha256") != donor.digest:
        raise ValueError("Donor SHA-256 changed; inspect and review the edit plan again.")
    pairs = []
    for op in plan["operations"]:
        if set(op) != {"target", "donor"}:
            raise ValueError("Replacement operations require exactly target and donor selectors.")
        pairs.append((source.target(op["target"]), donor, donor.target(op["donor"])))
    return _publish(source, output_path, pairs, overwrite)


def preferences(value):
    section, result = "", {}
    for line in value.splitlines():
        line = line.strip()
        if line.startswith("[") and line.endswith("]"):
            section = line[1:-1]
        elif "=" in line and not line.startswith(";"):
            key, val = line.split("=", 1)
            result[section, key.strip()] = val.strip()
    return result


def format_equations(input_path, output_path, manifest_path, overwrite=False, bridge=None):
    source = Package(input_path)
    _output_path(source, output_path, overwrite)
    plan = read_plan(manifest_path, source)
    jobs = []
    for op in plan["operations"]:
        if set(op) != {"target", "full_size_pt"}:
            raise ValueError("Formatting operations require exactly target and full_size_pt.")
        size = op["full_size_pt"]
        if isinstance(size, bool) or not isinstance(size, (float, int)) or not math.isfinite(size) or not 1 <= size <= 127:
            raise ValueError("full_size_pt must be a finite number between 1 and 127.")
        target = source.target(op["target"])
        if target["part"] != MAIN:
            raise ValueError("Internal-size formatting currently supports main-document inline equations only.")
        jobs.append((target, size))
    # A private, short-lived carrier prevents InsertFile/Word SaveAs from importing styles or renumbering the source.
    with tempfile.TemporaryDirectory(prefix="mathtype-local-") as folder:
        work = Path(folder)
        carrier = work / "carrier.docx"
        job = {"input_path": str(source.path), "output_path": str(carrier), "work_dir": str(work), "operations": []}
        for target, size in jobs:
            path = work / f"{target['index']}.mtef"
            path.write_bytes(mtef(source.data[target["embedding"]]))
            job["operations"].append({"index": target["index"], "full_size_pt": size, "mtef_path": str(path)})
        job_path = work / "job.json"
        job_path.write_text(json.dumps(job, ensure_ascii=True), encoding="utf-8")
        if bridge is None:
            import mcp_server
            import mathtype_batch
            mathtype_batch.refuse_concurrent_job()
            bridge = mcp_server._invoke_bridge
        result = bridge("format-native", {"manifest_path": str(job_path)})
        if not result.get("ok"):
            return result
        donor = Package(carrier)
        evidence = []
        pairs = []
        for target, size in jobs:
            index = target["index"]
            before = (work / f"{index}-before.mml").read_text(encoding="utf-8-sig")
            saved = (work / f"{index}-saved.mml").read_text(encoding="utf-8-sig")
            if math_tree(before) != math_tree(saved):
                raise ValueError(f"Equation #{index} changed mathematical structure (possible appended content).")
            old_prefs = preferences((work / f"{index}-before.prefs").read_text(encoding="utf-8-sig"))
            new_prefs = preferences((work / f"{index}-saved.prefs").read_text(encoding="utf-8-sig"))
            actual = new_prefs.pop(("Sizes", "Full"), "")
            old_prefs.pop(("Sizes", "Full"), None)
            match = re.fullmatch(r"\s*([0-9.]+)\s*pt\s*", actual)
            if not match or abs(float(match[1]) - size) > 0.001 or old_prefs != new_prefs:
                raise ValueError(f"Equation #{index} failed internal Full-size/style readback.")
            replacement = next(r for r in donor.objects if r["part"] == MAIN and r["index"] == index)
            pairs.append((target, donor, replacement))
            evidence.append({"index": index, "target_full_size_pt": size, "saved_full_size_pt": float(match[1]),
                             "mathml_tree_equal": True, "other_preferences_preserved": True,
                             "saved_native_sha256": replacement["native_sha256"]})
        return _publish(source, output_path, pairs, overwrite, evidence)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("inspect", "format", "replace"))
    parser.add_argument("--input", required=True)
    parser.add_argument("--output")
    parser.add_argument("--manifest")
    parser.add_argument("--donor")
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()
    try:
        if args.action == "inspect":
            result = inspect_document(args.input)
        elif not args.output or not args.manifest:
            raise ValueError("format/replace require --output and --manifest.")
        elif args.action == "format":
            result = format_equations(args.input, args.output, args.manifest, args.overwrite)
        elif not args.donor:
            raise ValueError("replace requires --donor.")
        else:
            result = replace_equations(args.input, args.output, args.manifest, args.donor, args.overwrite)
    except Exception as exc:
        result = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
    print(json.dumps(result, ensure_ascii=True))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
