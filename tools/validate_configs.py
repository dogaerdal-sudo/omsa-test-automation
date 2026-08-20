from __future__ import annotations
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import yaml

ROOT=Path(__file__).resolve().parents[1]
NS={'a':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def xlsx_sheets(path: Path):
    with zipfile.ZipFile(path) as z:
        root=ET.fromstring(z.read('xl/workbook.xml'))
    return [s.attrib['name'] for s in root.findall('a:sheets/a:sheet',NS)]

def book_sheets(path: Path):
    if path.suffix.lower() in {'.xlsx','.xlsm'}:
        return xlsx_sheets(path)
    if path.suffix.lower()=='.xlsb':
        from pyxlsb import open_workbook
        with open_workbook(path) as wb:
            return list(wb.sheets)
    return []

def check_sheet(name, sheets):
    return name in sheets

errors=[]; warnings=[]; counts={'loader':0,'view':0,'control':0}
for ctype,folder in [('loader','loaders'),('view','views'),('control','controls')]:
    for cfg_path in sorted((ROOT/'config'/folder).glob('*.yaml')):
        d=yaml.safe_load(cfg_path.read_text(encoding='utf-8'))
        counts[ctype]+=1
        template=ROOT/d['template'].get('folder',f'templates_local/{folder}')/d['template']['suggested_filename']
        if not template.exists():
            errors.append(f'{d["component_id"]}: missing template {template}')
            continue
        try: sheets=book_sheets(template)
        except Exception as e:
            errors.append(f'{d["component_id"]}: cannot inspect template: {e}')
            continue
        for slot in d.get('inputs',[]):
            s=slot.get('destination',{}).get('sheet')
            if s and not check_sheet(s,sheets): errors.append(f'{d["component_id"]}: input {slot["id"]} target sheet not found: {s}')
        rb=d.get('result_behavior',{})
        rs=rb.get('sheet') or rb.get('check_sheet')
        if rs and rs not in sheets: warnings.append(f'{d["component_id"]}: result sheet not found: {rs}')
        for a in d.get('post_import_actions',[]):
            for key in ['source','destination']:
                obj=a.get(key,{})
                s=obj.get('sheet')
                if s and s not in sheets: errors.append(f'{d["component_id"]}: action {a.get("action")} {key} sheet not found: {s}')
            s=a.get('sheet')
            if s and s not in sheets: errors.append(f'{d["component_id"]}: action {a.get("action")} sheet not found: {s}')

print('Component counts:',counts)
if warnings:
    print('\nWARNINGS:')
    for x in warnings: print(' -',x)
if errors:
    print('\nERRORS:')
    for x in errors: print(' -',x)
    sys.exit(1)
print('\nAll configured template/sheet references validated.')
