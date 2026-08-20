from __future__ import annotations

import os
import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from .utils import project_root, safe_name, col_to_num, num_to_col

XL_UP = -4162
XL_TO_LEFT = -4159
XL_PASTE_VALUES = -4163
XL_PASTE_FORMATS = -4122
XL_WHOLE = 1
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_DONE = 0


class ExcelUnavailable(RuntimeError):
    pass


def _win32():
    if os.name != 'nt':
        raise ExcelUnavailable('Generating and recalculating Excel workbooks requires Windows + Microsoft Excel Desktop.')
    try:
        import win32com.client as win32  # type: ignore
        return win32
    except Exception as exc:
        raise ExcelUnavailable('pywin32 is not available. Run SETUP_AND_RUN.bat and ensure Microsoft Excel Desktop is installed.') from exc


def _norm(value: Any) -> str:
    return ' '.join(str(value or '').strip().casefold().replace('_', ' ').replace('-', ' ').split())


def _find_header(ws, header: str, header_row: int = 1, aliases: list[str] | None = None) -> int:
    wanted = {_norm(header), *[_norm(x) for x in (aliases or [])]}
    last_col = int(ws.Cells(header_row, ws.Columns.Count).End(XL_TO_LEFT).Column)
    for col in range(1, last_col + 1):
        if _norm(ws.Cells(header_row, col).Value) in wanted:
            return col
    raise ValueError(f'Header {header!r} not found in sheet {ws.Name!r}')


def _last_row(ws, col: int, minimum: int = 1) -> int:
    row = int(ws.Cells(ws.Rows.Count, col).End(XL_UP).Row)
    return max(row, minimum)


def _last_col(ws, row: int = 1) -> int:
    return int(ws.Cells(row, ws.Columns.Count).End(XL_TO_LEFT).Column)


def _source_start_col(ws, slot: dict) -> int:
    src = slot.get('source', {})
    if src.get('start_header'):
        return _find_header(ws, src['start_header'], int(src.get('header_row', 1)), src.get('header_aliases'))
    return col_to_num(src.get('start_column', 'A'))


def _source_end_col(ws, slot: dict, header_row: int) -> int:
    src = slot.get('source', {})
    if src.get('end_header'):
        return _find_header(ws, src['end_header'], header_row, src.get('end_header_aliases'))
    if src.get('end_column'):
        return col_to_num(src['end_column'])
    return _last_col(ws, header_row)


def _parse_protected(spec: list[str] | None) -> set[int]:
    out: set[int] = set()
    for item in spec or []:
        item = str(item).upper().strip()
        if ':' in item:
            a, b = item.split(':', 1)
            out.update(range(col_to_num(a), col_to_num(b) + 1))
        else:
            out.add(col_to_num(item))
    return out




def _formula_columns_in_range(ws, start_row: int, start_col: int, end_col: int, scan_rows: int = 3) -> set[int]:
    """Return destination columns that already contain template formulas near the import anchor.

    These columns are treated as protected by default so data imports cannot overwrite the
    workbook's business logic. A component may explicitly opt out with
    import_rules.protect_existing_formula_columns: false.
    """
    protected: set[int] = set()
    last_scan_row = start_row + max(0, scan_rows - 1)
    for col in range(start_col, end_col + 1):
        for row in range(start_row, last_scan_row + 1):
            try:
                if bool(ws.Cells(row, col).HasFormula):
                    protected.add(col)
                    break
            except Exception:
                continue
    return protected


def _wait_for_calculation(excel, timeout: int = 240) -> None:
    start = time.time()
    while True:
        try:
            if int(excel.CalculationState) == XL_DONE:
                return
        except Exception:
            return
        if time.time() - start > timeout:
            raise TimeoutError('Excel calculation did not finish within the configured timeout.')
        time.sleep(0.2)


def _fill_formula_columns(ws, data_start_row: int, last_row: int, explicit_columns: list[str] | None = None) -> None:
    if last_row < data_start_row:
        return
    if explicit_columns:
        candidates = [col_to_num(c) for c in explicit_columns]
    else:
        used_last = max(_last_col(ws, max(1, data_start_row - 1)), _last_col(ws, data_start_row))
        candidates = list(range(1, used_last + 1))
    for col in candidates:
        try:
            seed = ws.Cells(data_start_row, col)
            if bool(seed.HasFormula):
                seed_rng = ws.Range(seed, seed)
                dest = ws.Range(ws.Cells(data_start_row, col), ws.Cells(last_row, col))
                seed_rng.AutoFill(Destination=dest)
        except Exception:
            continue


def _clear_range(ws, start_row: int, start_col: int, end_col: int) -> None:
    ws.Range(ws.Cells(start_row, start_col), ws.Cells(ws.Rows.Count, end_col)).ClearContents()


def _replace_null(rng) -> None:
    try:
        rng.Replace(What='NULL', Replacement='', LookAt=XL_WHOLE)
    except Exception:
        pass


def _import_contiguous(excel, src_ws, dst_ws, slot: dict) -> dict:
    src = slot.get('source', {})
    rules = slot.get('import_rules', {})
    header_row = int(src.get('header_row', 1))
    src_col = _source_start_col(src_ws, slot)
    src_end_col = _source_end_col(src_ws, slot, header_row)
    src_row = int(src.get('start_row', header_row + (1 if rules.get('skip_source_header', True) else 0)))
    last_row = _last_row(src_ws, src_col, header_row)
    if last_row < src_row or src_end_col < src_col:
        raise ValueError(f'No importable rows found for {slot.get("label", slot.get("id"))}')

    dst = slot['destination']
    dst_row = int(dst.get('start_row', 2))
    dst_col = col_to_num(dst.get('start_column', 'A'))
    width = src_end_col - src_col + 1
    end_dst_col = dst_col + width - 1
    protected = _parse_protected(rules.get('protected_destination_columns'))
    if rules.get('protect_existing_formula_columns', True):
        protected |= _formula_columns_in_range(dst_ws, dst_row, dst_col, end_dst_col)

    if rules.get('clear_existing_data', True):
        if not protected:
            _clear_range(dst_ws, dst_row, dst_col, end_dst_col)
        else:
            for c in range(dst_col, end_dst_col + 1):
                if c not in protected:
                    dst_ws.Range(dst_ws.Cells(dst_row, c), dst_ws.Cells(dst_ws.Rows.Count, c)).ClearContents()

    src_rng = src_ws.Range(src_ws.Cells(src_row, src_col), src_ws.Cells(last_row, src_end_col))
    dst_anchor = dst_ws.Cells(dst_row, dst_col)

    if protected:
        # Paste one destination column at a time to avoid protected formula columns.
        for offset in range(width):
            dc = dst_col + offset
            if dc in protected:
                continue
            sc = src_col + offset
            one = src_ws.Range(src_ws.Cells(src_row, sc), src_ws.Cells(last_row, sc))
            one.Copy()
            dst_ws.Cells(dst_row, dc).PasteSpecial(Paste=XL_PASTE_VALUES)
            if rules.get('paste_formats', True):
                dst_ws.Cells(dst_row, dc).PasteSpecial(Paste=XL_PASTE_FORMATS)
            excel.CutCopyMode = False
    else:
        src_rng.Copy()
        if rules.get('paste_values', True):
            dst_anchor.PasteSpecial(Paste=XL_PASTE_VALUES)
        if rules.get('paste_formats', True):
            dst_anchor.PasteSpecial(Paste=XL_PASTE_FORMATS)
        excel.CutCopyMode = False

    dst_last_row = dst_row + (last_row - src_row)
    pasted = dst_ws.Range(dst_ws.Cells(dst_row, dst_col), dst_ws.Cells(dst_last_row, end_dst_col))
    if rules.get('replace_literal_NULL_with_blank', False):
        _replace_null(pasted)

    if rules.get('fill_helper_formulas_left', True) and dst_col > 1:
        _fill_formula_columns(dst_ws, dst_row, dst_last_row, [num_to_col(c) for c in range(1, dst_col)])
    if rules.get('fill_formula_columns'):
        _fill_formula_columns(dst_ws, dst_row, dst_last_row, rules.get('fill_formula_columns'))
    # Also extend any formula columns that already exist in the template row.
    if rules.get('auto_fill_existing_formula_columns', True):
        _fill_formula_columns(dst_ws, dst_row, dst_last_row)

    return {'last_row': dst_last_row, 'row_count': dst_last_row - dst_row + 1, 'start_row': dst_row, 'start_col': dst_col, 'end_col': end_dst_col}


def _import_header_match(excel, src_ws, dst_ws, slot: dict) -> dict:
    src = slot.get('source', {})
    rules = slot.get('import_rules', {})
    header_row = int(src.get('header_row', 1))
    src_start_row = int(src.get('start_row', 2))
    dst_row = int(slot['destination'].get('start_row', 2))
    dst_header_row = int(slot['destination'].get('header_row', 1))
    src_last_col = _last_col(src_ws, header_row)
    dst_last_col = _last_col(dst_ws, dst_header_row)
    protected = _parse_protected(rules.get('protected_destination_columns'))
    if rules.get('protect_existing_formula_columns', True):
        protected |= _formula_columns_in_range(dst_ws, dst_row, 1, dst_last_col)

    dest_headers: dict[str, int] = {}
    for c in range(1, dst_last_col + 1):
        h = _norm(dst_ws.Cells(dst_header_row, c).Value)
        if h:
            dest_headers[h] = c

    matched: list[tuple[int, int]] = []
    for sc in range(1, src_last_col + 1):
        h = _norm(src_ws.Cells(header_row, sc).Value)
        dc = dest_headers.get(h)
        if dc and dc not in protected:
            matched.append((sc, dc))
    if not matched:
        raise ValueError(f'No matching headers found for {slot.get("label", slot.get("id"))}')

    last_row = max(_last_row(src_ws, sc, header_row) for sc, _ in matched)
    if last_row < src_start_row:
        raise ValueError(f'No importable rows found for {slot.get("label", slot.get("id"))}')

    for sc, dc in matched:
        if rules.get('clear_existing_data', True):
            dst_ws.Range(dst_ws.Cells(dst_row, dc), dst_ws.Cells(dst_ws.Rows.Count, dc)).ClearContents()
        rng = src_ws.Range(src_ws.Cells(src_start_row, sc), src_ws.Cells(last_row, sc))
        rng.Copy()
        dst_ws.Cells(dst_row, dc).PasteSpecial(Paste=XL_PASTE_VALUES)
        if rules.get('paste_formats', True):
            dst_ws.Cells(dst_row, dc).PasteSpecial(Paste=XL_PASTE_FORMATS)
        excel.CutCopyMode = False
        if rules.get('replace_literal_NULL_with_blank', False):
            _replace_null(dst_ws.Range(dst_ws.Cells(dst_row, dc), dst_ws.Cells(dst_row + last_row - src_start_row, dc)))

    dst_last_row = dst_row + last_row - src_start_row
    _fill_formula_columns(dst_ws, dst_row, dst_last_row)
    return {'last_row': dst_last_row, 'row_count': dst_last_row - dst_row + 1, 'start_row': dst_row}


def _cell_values(ws, col: int, start_row: int, last_row: int) -> list[Any]:
    if last_row < start_row:
        return []
    raw = ws.Range(ws.Cells(start_row, col), ws.Cells(last_row, col)).Value
    if last_row == start_row:
        return [raw]
    out=[]
    for row in raw:
        out.append(row[0] if isinstance(row, tuple) else row)
    return out


def _write_matrix(ws, start_row: int, start_col: int, rows: list[list[Any]], clear: bool = True, width: int | None = None) -> int:
    width = width or (len(rows[0]) if rows else 1)
    if clear:
        ws.Range(ws.Cells(start_row, start_col), ws.Cells(ws.Rows.Count, start_col + width - 1)).ClearContents()
    if not rows:
        return start_row - 1
    normalized = tuple(tuple(r) for r in rows)
    ws.Range(ws.Cells(start_row, start_col), ws.Cells(start_row + len(rows) - 1, start_col + width - 1)).Value = normalized
    return start_row + len(rows) - 1


def _header_col_with_aliases(ws, header: str, action: dict) -> int:
    aliases_map = action.get('header_aliases', {})
    return _find_header(ws, header, int(action.get('source', {}).get('header_row', 1)), aliases_map.get(header, []))


def _filter_mask(ws, action: dict, start_row: int, last_row: int) -> list[bool]:
    f = action.get('filter')
    if not f:
        return [True] * max(0, last_row - start_row + 1)
    if f.get('column_header'):
        col = _find_header(ws, f['column_header'], int(action.get('source', {}).get('header_row', 1)), f.get('aliases'))
    else:
        col = col_to_num(f.get('column_letter', 'A'))
    vals = _cell_values(ws, col, start_row, last_row)
    allowed = {_norm(x) for x in f.get('allowed_values', [])}
    out=[]
    for v in vals:
        if f.get('nonblank'):
            ok = v not in (None, '') and _norm(v) not in {'null'}
        elif allowed:
            ok = _norm(v) in allowed
        else:
            ok = True
        out.append(ok)
    return out


def _run_action(wb, action: dict, reporting_date: str | None, slot_context: dict[str, dict]) -> dict | None:
    typ = action.get('action')

    if typ == 'copy_column_values':
        src = wb.Worksheets(action['source']['sheet'])
        dst = wb.Worksheets(action['destination']['sheet'])
        sr = int(action['source'].get('start_row', 2))
        if action['source'].get('column_header'):
            sc = _find_header(src, action['source']['column_header'], int(action['source'].get('header_row', 1)), action['source'].get('aliases'))
        else:
            sc = col_to_num(action['source'].get('column_letter', 'A'))
        lr = _last_row(src, sc, sr - 1)
        vals = _cell_values(src, sc, sr, lr)
        mask = _filter_mask(src, action, sr, lr)
        vals = [v for v, keep in zip(vals, mask) if keep]
        if action.get('remove_blank', False):
            vals = [v for v in vals if v not in (None, '') and _norm(v) != 'null']
        if action.get('unique'):
            seen=set(); tmp=[]
            for v in vals:
                k=str(v)
                if k not in seen:
                    seen.add(k); tmp.append(v)
            vals=tmp
        if action.get('sort'):
            vals=sorted(vals, key=lambda x: str(x))
        anchor = dst.Range(action['destination'].get('start_cell', 'A4'))
        rows=[[v] for v in vals]
        last = _write_matrix(dst, anchor.Row, anchor.Column, rows, action.get('clear_destination_before_copy', True), 1)
        _fill_formula_columns(dst, anchor.Row, last)
        return {'destination_last_row': last, 'row_count': len(rows)}

    if typ == 'copy_unique_values':
        src = wb.Worksheets(action['source']['sheet'])
        dst = wb.Worksheets(action['destination']['sheet'])
        sr = int(action['source'].get('start_row', 2)); sc = col_to_num(action['source'].get('column_letter', 'A'))
        lr = _last_row(src, sc, sr - 1)
        vals = _cell_values(src, sc, sr, lr)
        if action.get('remove_blank', True):
            vals=[v for v in vals if v not in (None,'') and _norm(v) != 'null']
        seen=set(); uniq=[]
        for v in vals:
            k=str(v)
            if k not in seen:
                seen.add(k); uniq.append(v)
        if action.get('sort', False): uniq=sorted(uniq,key=lambda x:str(x))
        a=dst.Range(action['destination']['start_cell'])
        last=_write_matrix(dst,a.Row,a.Column,[[v] for v in uniq],action.get('clear_destination_before_copy',True),1)
        _fill_formula_columns(dst,a.Row,last)
        return {'destination_last_row':last,'row_count':len(uniq)}

    if typ == 'copy_unique_values_filtered':
        src=wb.Worksheets(action['source']['sheet']); dst=wb.Worksheets(action['destination']['sheet'])
        sr=int(action['source'].get('start_row',2)); hr=int(action['source'].get('header_row',1))
        sc=_find_header(src,action['source']['column_header'],hr)
        distinct_col=_find_header(src,action['distinct_by_header'],hr) if action.get('distinct_by_header') else sc
        lr=max(_last_row(src,sc,sr-1),_last_row(src,distinct_col,sr-1))
        vals=_cell_values(src,sc,sr,lr); distinct=_cell_values(src,distinct_col,sr,lr); mask=_filter_mask(src,action,sr,lr)
        seen=set(); out=[]
        for v,d,k in zip(vals,distinct,mask):
            if not k: continue
            key=str(d)
            if key in seen: continue
            seen.add(key); out.append(v)
        a=dst.Range(action['destination']['start_cell'])
        last=_write_matrix(dst,a.Row,a.Column,[[v] for v in out],action.get('clear_destination_before_copy',True),1)
        _fill_formula_columns(dst,a.Row,last)
        return {'destination_last_row':last,'row_count':len(out)}

    if typ == 'copy_columns_filtered_unique':
        src=wb.Worksheets(action['source']['sheet']); dst=wb.Worksheets(action['destination']['sheet'])
        sr=int(action['source'].get('start_row',2)); hr=int(action['source'].get('header_row',1))
        cols=[_header_col_with_aliases(src,h,action) for h in action['source']['column_headers']]
        filter_col=_header_col_with_aliases(src,action['filter']['column_header'],action)
        lr=max([_last_row(src,c,sr-1) for c in cols+[filter_col]])
        matrices=[_cell_values(src,c,sr,lr) for c in cols]
        fvals=_cell_values(src,filter_col,sr,lr)
        rows=[]; seen=set()
        for idx,fv in enumerate(fvals):
            if action['filter'].get('nonblank') and (fv in (None,'') or _norm(fv)=='null'):
                continue
            row=[m[idx] for m in matrices]
            key=tuple(str(x) for x in row)
            if action.get('remove_duplicates',True) and key in seen: continue
            seen.add(key); rows.append(row)
        a=dst.Range(action['destination']['start_cell'])
        last=_write_matrix(dst,a.Row,a.Column,rows,action.get('clear_destination_before_copy',True),len(cols))
        _fill_formula_columns(dst,a.Row,last)
        return {'destination_last_row':last,'row_count':len(rows)}

    if typ == 'copy_range_values':
        src=wb.Worksheets(action['source']['sheet']); dst=wb.Worksheets(action['destination']['sheet'])
        sr=int(action['source'].get('start_row',2)); sc=col_to_num(action['source']['start_column']); ec=col_to_num(action['source']['end_column'])
        lr=max(_last_row(src,c,sr-1) for c in range(sc,ec+1))
        if lr < sr: return {'row_count':0}
        vals=src.Range(src.Cells(sr,sc),src.Cells(lr,ec)).Value
        a=dst.Range(action['destination']['start_cell'])
        width=ec-sc+1
        if action.get('clear_destination_before_copy',True):
            _clear_range(dst,a.Row,a.Column,a.Column+width-1)
        dst.Range(dst.Cells(a.Row,a.Column),dst.Cells(a.Row+lr-sr,a.Column+width-1)).Value=vals
        last=a.Row+lr-sr
        _fill_formula_columns(dst,a.Row,last)
        return {'destination_last_row':last,'row_count':last-a.Row+1}

    if typ == 'remove_blank_values':
        ws=wb.Worksheets(action['sheet']); col=col_to_num(action.get('column','A')); sr=int(action.get('start_row',2))
        lr=_last_row(ws,col,sr-1); vals=_cell_values(ws,col,sr,lr)
        vals=[v for v in vals if v not in (None,'') and _norm(v)!='null']
        last=_write_matrix(ws,sr,col,[[v] for v in vals],True,1)
        _fill_formula_columns(ws,sr,last)
        return {'destination_last_row':last,'row_count':len(vals)}

    if typ == 'set_column_to_reporting_date':
        if not reporting_date:
            return None
        ws=wb.Worksheets(action['sheet']); col=col_to_num(action['column']); sr=int(action.get('start_row',2))
        anchor=action.get('anchor_input'); ctx=slot_context.get(anchor,{}) if anchor else {}
        lr=int(ctx.get('last_row') or _last_row(ws,col,sr-1))
        if lr>=sr:
            ws.Range(ws.Cells(sr,col),ws.Cells(lr,col)).Value=reporting_date
        return {'destination_last_row':lr,'row_count':max(0,lr-sr+1)}

    return None


def _summarize_results(wb, behavior: dict) -> dict:
    sheet_name = behavior.get('sheet') or behavior.get('check_sheet') or 'CHECK'
    try:
        ws = wb.Worksheets(sheet_name)
    except Exception:
        return {'true_count':0,'false_count':0,'error_count':0,'non_boolean_count':0}

    start_row=int(behavior.get('data_start_row',behavior.get('first_data_row',4)))
    key_col=col_to_num(behavior.get('key_column','A'))
    last_row=_last_row(ws,key_col,start_row-1)
    if last_row < start_row:
        return {'true_count':0,'false_count':0,'error_count':0,'non_boolean_count':0}

    result_cols=behavior.get('result_columns')
    if result_cols:
        cols=[col_to_num(c) for c in result_cols]
    else:
        header_row=int(behavior.get('header_row',max(1,start_row-1)))
        last_col=max(_last_col(ws,header_row),_last_col(ws,start_row))
        cols=list(range(key_col+1,last_col+1))
    if not cols:
        return {'true_count':0,'false_count':0,'error_count':0,'non_boolean_count':0}

    true_count=false_count=error_count=non_boolean_count=0
    formula_only=behavior.get('scan_formula_cells_only',True)
    for col in cols:
        rng=ws.Range(ws.Cells(start_row,col),ws.Cells(last_row,col))
        values=rng.Value
        formulas=rng.Formula
        if last_row==start_row:
            values=((values,),) if not isinstance(values,tuple) else (values if isinstance(values[0] if values else None,tuple) else (values,))
            formulas=((formulas,),) if not isinstance(formulas,tuple) else (formulas if isinstance(formulas[0] if formulas else None,tuple) else (formulas,))
        for i,row in enumerate(values or []):
            value=row[0] if isinstance(row,tuple) else row
            frow=formulas[i] if isinstance(formulas,tuple) and i<len(formulas) else None
            formula=frow[0] if isinstance(frow,tuple) else frow
            if formula_only and not (isinstance(formula,str) and formula.startswith('=')):
                continue
            if value is True or (isinstance(value,str) and value.strip().upper()=='TRUE'):
                true_count+=1
            elif value is False or (isinstance(value,str) and value.strip().upper()=='FALSE'):
                false_count+=1
            elif isinstance(value,str) and value.startswith('#'):
                error_count+=1
            elif value not in (None,''):
                non_boolean_count+=1
    return {'true_count':true_count,'false_count':false_count,'error_count':error_count,'non_boolean_count':non_boolean_count}


def generate(component: dict[str, Any], template_path: Path | None, inputs: dict[str, Path], reporting_date: str | None = None) -> dict:
    if template_path is None or not Path(template_path).exists():
        raise FileNotFoundError('Active master template is missing for this component.')
    if not component.get('allow_generate', True):
        raise ValueError('Generation is disabled for this component until the remaining mapping is confirmed.')

    win32=_win32()
    run_id=f"{component['component_id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"
    run_dir=project_root()/'storage'/'runs'/run_id
    run_dir.mkdir(parents=True,exist_ok=True)
    output_name=f"{safe_name(component['component_id'])}_{reporting_date or datetime.now().strftime('%Y-%m-%d')}_{run_id[-6:]}{Path(template_path).suffix}"
    output_path=run_dir/output_name
    shutil.copy2(template_path,output_path)

    saved_inputs={}
    for slot in component.get('inputs',[]):
        sid=slot['id']
        if sid not in inputs:
            if slot.get('required',True):
                raise ValueError(f"Missing required input: {slot.get('label',sid)}")
            continue
        src=Path(inputs[sid]); dest=run_dir/'inputs'/safe_name(src.name); dest.parent.mkdir(parents=True,exist_ok=True); shutil.copy2(src,dest); saved_inputs[sid]=dest

    excel=win32.DispatchEx('Excel.Application')
    excel.Visible=False; excel.DisplayAlerts=False; excel.ScreenUpdating=False; excel.EnableEvents=False
    slot_context={}
    try:
        try: excel.AutomationSecurity=3
        except Exception: pass
        excel.Calculation=XL_CALC_MANUAL
        wb=excel.Workbooks.Open(str(output_path.resolve()))
        try:
            for slot in component.get('inputs',[]):
                sid=slot['id']
                if sid not in saved_inputs: continue
                src_wb=excel.Workbooks.Open(str(saved_inputs[sid].resolve()))
                try:
                    src_ws=src_wb.Worksheets(int(slot.get('source',{}).get('sheet_index',1)))
                    dst_ws=wb.Worksheets(slot['destination']['sheet'])
                    if slot.get('import_mode','contiguous')=='header_match':
                        ctx=_import_header_match(excel,src_ws,dst_ws,slot)
                    else:
                        ctx=_import_contiguous(excel,src_ws,dst_ws,slot)
                    slot_context[sid]=ctx
                finally:
                    src_wb.Close(SaveChanges=False)

            excel.Calculation=XL_CALC_AUTOMATIC
            excel.CalculateFullRebuild(); _wait_for_calculation(excel)

            for action in component.get('post_import_actions',[]):
                _run_action(wb,action,reporting_date,slot_context)

            # Ensure result formulas extend to the active key range.
            behavior=component.get('result_behavior',{})
            try:
                result_ws=wb.Worksheets(behavior.get('sheet') or behavior.get('check_sheet') or 'CHECK')
                sr=int(behavior.get('data_start_row',behavior.get('first_data_row',4)))
                kc=col_to_num(behavior.get('key_column','A'))
                lr=_last_row(result_ws,kc,sr-1)
                _fill_formula_columns(result_ws,sr,lr)
            except Exception:
                pass

            excel.CalculateFullRebuild(); _wait_for_calculation(excel)
            summary=_summarize_results(wb,behavior)
            wb.Save()
        finally:
            wb.Close(SaveChanges=True)
    finally:
        try: excel.EnableEvents=True
        except Exception: pass
        try: excel.DisplayAlerts=True; excel.ScreenUpdating=True
        except Exception: pass
        excel.Quit()

    status='ERROR' if summary['error_count'] else ('FAILED' if summary['false_count'] else 'PASSED')
    return {'run_id':run_id,'output_path':str(output_path),'status':status,**summary,'input_files':{k:str(v) for k,v in saved_inputs.items()}}
