"""Opt-in regression for reviewed OMML marking, native chapter/section batching, and actual body size."""
import argparse
import json
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
import mcp_server as server
import mathtype_local as local


def require(result):
    if not result.get('ok'):
        raise RuntimeError(json.dumps(result))
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output-dir', required=True, type=Path)
    folder = parser.parse_args().output_dir.resolve()
    folder.mkdir(parents=True, exist_ok=False)
    source = folder / 'omml-source.docx'
    subprocess.run([server._powershell_executable(), '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                    '-File', str(ROOT / 'tests/create-live-fixture.ps1'), '-OutputPath', str(source), '-Omml'], check=True, timeout=90)
    source_hash = local.sha(source.read_bytes())
    scan_path = folder / 'scan.json'
    require(server._invoke_bridge('scan', {'document_path':str(source), 'candidates_path':str(scan_path)}))
    scan = json.loads(scan_path.read_text(encoding='utf-8-sig'))
    assert len(scan['word_equation_objects']) == 6
    marks = {'schema_version':1,'source_sha256':scan['source_sha256'], 'marks':[
        {'omml_index':i,'marker':'{{MATH:e'+str(i)+'}}'} for i in range(1,7)]}
    marks_path = folder / 'marks.json'; marks_path.write_text(json.dumps(marks),encoding='utf-8')
    marked = folder / 'marked.docx'
    require(server._invoke_bridge('mark', {'document_path':str(source),'output_path':str(marked),'marks_path':str(marks_path)}))
    manifest = {'schema_version':1,'numbering':'chapter_section','equations':[
        {'id':'e'+str(i),'marker':'{{MATH:e'+str(i)+'}}','tex':'x_'+str(i)+'^2','layout':'display','numbered':True,
         'break_before':{3:'next_section',4:'next_chapter',6:'next_section'}.get(i,'none')}
        for i in range(1,7)], 'references':[{'marker':'{{EQREF:e1}}','target':'e1'},{'marker':'{{EQREF:e6}}','target':'e6'}]}
    manifest_path = folder / 'manifest.json';manifest_path.write_text(json.dumps(manifest),encoding='utf-8')
    output = folder / 'rendered.docx'
    rendered = require(server._render_word({'input_path':str(marked),'output_path':str(output),
        'manifest_path':str(manifest_path),'batch_size':2}))
    checked = require(server._invoke_bridge('validate',{'document_path':str(output),'manifest_path':str(manifest_path)}))
    expected = ['(1.1.1)','(1.1.2)','(1.2.1)','(2.1.1)','(2.1.2)','(2.2.1)']
    assert checked['equation_labels'] == expected, checked
    updated = folder / 'updated.docx'
    require(server._invoke_bridge('update',{'input_path':str(output),'output_path':str(updated)}))
    package = local.Package(updated)
    row = package.objects[2]
    font_plan = {'schema_version':1,'source_sha256':package.digest,'operations':[
        {'target':{k:row[k] for k in ('part','index','native_sha256')},'full_size_pt':'match_context'}]}
    plan_path=folder / 'size-plan.json';plan_path.write_text(json.dumps(font_plan),encoding='utf-8')
    formatted=require(local.format_equations(updated,folder/'body-size.docx',plan_path))
    assert formatted['evidence'][0]['saved_full_size_pt'] == 10.5, formatted
    assert formatted['evidence'][0]['body_size_source']['full_size_pt'] == 10.5
    profiles = {}
    for mode, boundaries, labels in (
        ('chapter', {4:'next_chapter'}, ['(1-1)','(1-2)','(1-3)','(2-1)','(2-2)','(2-3)']),
        ('section', {3:'next_section',6:'next_section'}, ['(1.1)','(1.2)','(2.1)','(2.2)','(2.3)','(3.1)']),
    ):
        alternative = dict(manifest, numbering=mode, equations=[dict(e, break_before=boundaries.get(i,'none'))
                           for i,e in enumerate(manifest['equations'],1)])
        alternative_path = folder / (mode+'.json')
        alternative_path.write_text(json.dumps(alternative),encoding='utf-8')
        result=require(server._render_word({'input_path':str(marked),'output_path':str(folder/(mode+'.docx')),
            'manifest_path':str(alternative_path),'batch_size':0}))
        assert result['validation']['equation_labels'] == labels, result
        profiles[mode]=result['validation']

    # Selecting one OMML must not require converting every other equation.
    partial_marks_path=folder/'partial-marks.json'
    partial_marks_path.write_text(json.dumps(dict(marks,marks=marks['marks'][:1])),encoding='utf-8')
    partial_marked=folder/'partial-marked.docx'
    require(server._invoke_bridge('mark',{'document_path':str(source),'output_path':str(partial_marked),
                                         'marks_path':str(partial_marks_path)}))
    partial_manifest={'schema_version':1,'expected_remaining_omml':5,
                      'equations':[dict(manifest['equations'][0],layout='inline',numbered=False)]}
    partial_path=folder/'partial.json';partial_path.write_text(json.dumps(partial_manifest),encoding='utf-8')
    partial_output=folder/'partial.docx'
    partial=require(server._render_word({'input_path':str(partial_marked),'output_path':str(partial_output),
        'manifest_path':str(partial_path),'partial_manifest':True,'batch_size':0}))
    assert partial['validation']['counts']['word_builtin_omath'] == 5, partial
    def omml(path):
        with zipfile.ZipFile(path) as package:
            root=ET.fromstring(package.read('word/document.xml'))
        return [ET.tostring(e) for e in root.iter('{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath')]
    assert omml(source)[1:] == omml(partial_output), 'Unselected OMML changed'
    assert source_hash == local.sha(source.read_bytes())
    report={'ok':True,'render':rendered,'validation':checked,'format':formatted,'profiles':profiles,
            'partial_conversion':partial,'unselected_omml_equal':True,'source_unchanged':True}
    (folder/'report.json').write_text(json.dumps(report,ensure_ascii=True,indent=2),encoding='utf-8')
    print(json.dumps({'ok':True,'report':str(folder/'report.json')}))


if __name__ == '__main__':
    main()
