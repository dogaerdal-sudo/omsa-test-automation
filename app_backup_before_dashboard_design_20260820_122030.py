from __future__ import annotations

import base64
import html
import os
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st

from src.config_store import load_components, load_component
from src.template_store import active_template, template_info, register_template
from src.excel_engine import generate, ExcelUnavailable
from src.history import save_run, list_runs
from src.utils import safe_name

st.set_page_config(page_title='OMSA Test Automation', page_icon='📊', layout='wide', initial_sidebar_state='collapsed')

CSS = r'''
<style>
:root {
  --omsa-navy:#13354d;
  --omsa-navy-mid:#20586a;
  --omsa-teal:#3d949d;
  --omsa-header:#c9ddd8;
  --omsa-page:#e7e7e7;
  --omsa-line:#d2d6d8;
  --omsa-text:#101820;
  --omsa-muted:#69757e;
  --omsa-blue:#3978c9;
  --omsa-blue-hover:#2f69b4;
  --omsa-red:#e00022;
}

/* Hide Streamlit chrome and remove default margins that were clipping the custom shell. */
header[data-testid="stHeader"] {display:none !important;}
footer {display:none !important;}
[data-testid="stSidebar"] {display:none !important;}
html, body, [data-testid="stAppViewContainer"], .stApp {margin:0 !important; padding:0 !important; overflow-x:hidden !important;}
.stApp {background:var(--omsa-page) !important; color:var(--omsa-text) !important;}
[data-testid="stMain"], [data-testid="stMainBlockContainer"], .block-container {
  width:100% !important; max-width:none !important; margin:0 !important; padding:0 0 2.5rem 0 !important;
}
div[data-testid="stVerticalBlock"] {gap:.65rem;}

/* ===== OMSA TOP HEADER ===== */
.omsa-topbar {
  box-sizing:border-box;
  position:relative;
  width:100%;
  height:70px;
  background:#fff;
  border-bottom:1px solid #d8dadd;
  display:grid;
  grid-template-columns:230px minmax(0,1fr) 55px;
  align-items:center;
  padding:0 14px 0 8px;
}
.omsa-brand {display:flex; align-items:center; min-width:0; height:100%;}
.omsa-brand img {display:block; width:166px; height:auto; object-fit:contain;}
.omsa-nav {
  min-width:0;
  display:flex;
  justify-content:center;
  align-items:center;
  gap:2px;
  white-space:nowrap;
}
.omsa-nav a {
  color:#111 !important;
  text-decoration:none !important;
  font-size:11.5px;
  font-weight:400;
  line-height:1;
  padding:13px 14px;
  border-radius:8px;
}
.omsa-nav a:hover {background:#f2f2f2;}
.omsa-nav a.active {background:#efeff0; color:#3e969d !important;}
.omsa-dot {
  justify-self:end;
  width:38px;
  height:38px;
  border-radius:50%;
  background:var(--omsa-red);
}

/* ===== OMSA SECTION BAR ===== */
.omsa-sectionbar {
  box-sizing:border-box;
  width:100%;
  height:64px;
  display:flex;
  align-items:center;
  padding:0 4px;
  background:linear-gradient(90deg,#13354d 0%,#1f5365 48%,#3d949d 100%);
  color:#fff;
  font-size:16px;
  font-weight:700;
  border-bottom:1px solid rgba(0,0,0,.06);
}

/* ===== MAIN WORK AREA ===== */
.omsa-body {box-sizing:border-box; width:100%; padding:20px 2px 34px 2px;}
.omsa-table-wrap {
  width:min(1455px, calc(100vw - 4px));
  max-width:1455px;
  background:#fff;
  border:1px solid #d6dadc;
  border-radius:4px;
  overflow:hidden;
  box-shadow:none;
}
table.omsa-table {width:100%; border-collapse:collapse; table-layout:fixed; font-size:12px;}
.omsa-table th {
  background:var(--omsa-header);
  color:#24323a;
  text-align:left;
  font-weight:500;
  padding:6px 7px;
  border-right:1px solid #b9cbc7;
  white-space:nowrap;
  overflow:hidden;
  text-overflow:ellipsis;
  font-size:11px;
  height:26px;
}
.omsa-table td {
  background:#fff;
  color:#171d22;
  padding:5px 7px;
  height:26px;
  border-top:1px solid #e1e4e5;
  border-right:1px solid #e1e4e5;
  vertical-align:middle;
  overflow:hidden;
  text-overflow:ellipsis;
  white-space:nowrap;
}
.omsa-table tr:hover td {background:#f7f8f8;}
.omsa-table td.center,.omsa-table th.center{text-align:center;}
.omsa-table th:nth-child(1){width:23%;}
.omsa-table th:nth-child(2){width:42%;}
.omsa-table th:nth-child(3){width:11%;}
.omsa-table th:nth-child(4){width:9%;}
.omsa-table th:nth-child(5){width:15%;}

.run-config-btn {
  display:inline-block;
  min-width:128px;
  box-sizing:border-box;
  background:var(--omsa-blue);
  color:#fff !important;
  text-decoration:none !important;
  padding:4px 10px;
  border-radius:4px;
  font-size:10.5px;
  font-weight:500;
  text-align:center;
}
.run-config-btn:hover {background:var(--omsa-blue-hover);}
.launch-btn {
  display:inline-block;
  min-width:128px;
  box-sizing:border-box;
  background:var(--omsa-blue);
  color:#fff !important;
  text-decoration:none !important;
  padding:4px 10px;
  border-radius:4px;
  font-size:10.5px;
  font-weight:500;
  text-align:center;
}
.launch-btn:hover {background:var(--omsa-blue-hover);}
.status {font-style:italic;font-weight:600;}
.status.ready {color:#222;}
.status.review {color:#9a6400;}
.status.blocked {color:#b42318;}

/* Detail page */
.breadcrumb {font-size:12px; margin:0 0 10px 4px;}
.breadcrumb a {color:#2f6da6 !important; text-decoration:none !important;}
.info-strip {
  box-sizing:border-box;
  width:min(1455px, calc(100vw - 4px));
  max-width:1455px;
  background:#fff;
  border:1px solid #d2d6d8;
  border-left:3px solid #3d949d;
  border-radius:3px;
  padding:9px 12px;
  margin:0 0 12px 0;
  display:flex;
  gap:28px;
  align-items:center;
  flex-wrap:wrap;
}
.info-strip .label {color:#737e86;font-size:10px;text-transform:uppercase;letter-spacing:.35px;}
.info-strip .value {font-size:12px;font-weight:600;margin-top:2px;}
.upload-head {
  box-sizing:border-box;
  width:min(1455px, calc(100vw - 4px));
  max-width:1455px;
  background:var(--omsa-header);
  padding:6px 8px;
  font-size:11.5px;
  font-weight:600;
  border:1px solid #bdcfcb;
  border-bottom:0;
}
.mapping,.rule-note {font-size:11px;color:#5d6870;margin:3px 0 0 4px;}
.small-tag {display:inline-block;padding:2px 7px;border-radius:9px;background:#eef2f4;color:#52606a;font-size:9px;margin-left:7px;}
.small-tag.warn {background:#fff3d6;color:#8a5a00;}
.small-tag.ok {background:#e8f5ed;color:#216e47;}

/* Dashboard: compact, left-aligned, OMSA-like */
.kpi-row {
  display:grid;
  grid-template-columns:repeat(4,240px);
  gap:12px;
  margin:2px 0 10px 0;
}
.kpi {background:#fff;border:1px solid #d5d9db;padding:16px 14px;min-height:82px;box-sizing:border-box;}
.kpi .v{font-size:24px;font-weight:700;color:#23313a;}
.kpi .l{font-size:11px;color:#68747d;margin-top:6px;}
.card-lite {
  box-sizing:border-box;
  width:min(1455px, calc(100vw - 4px));
  max-width:1455px;
  background:#fff;
  border:1px solid #d5d9db;
  padding:13px 14px;
  border-radius:3px;
}

/* Native Streamlit inputs */
.stButton>button,.stDownloadButton>button {
  background:var(--omsa-blue) !important;
  color:#fff !important;
  border:0 !important;
  border-radius:4px !important;
  font-size:12px !important;
}
.stButton>button:hover,.stDownloadButton>button:hover {background:var(--omsa-blue-hover) !important;}
[data-testid="stFileUploader"] {
  box-sizing:border-box;
  width:min(1455px, calc(100vw - 4px));
  max-width:1455px;
  background:#fff;
  border:1px solid #d5dbde;
  padding:9px;
  border-radius:0 0 3px 3px;
}
[data-testid="stFileUploaderDropzone"] {background:#fafbfb;border:1px dashed #aeb8bd;}
[data-testid="stDateInput"] {max-width:220px;}
[data-testid="stDataFrame"] {background:#fff;border:1px solid #d7dbdd;}
.stAlert {border-radius:3px;}

@media (max-width:1250px) {
  .omsa-topbar {grid-template-columns:190px minmax(0,1fr) 48px; padding-left:6px;}
  .omsa-brand img {width:150px;}
  .omsa-nav a {font-size:10.5px; padding:12px 9px;}
  .omsa-dot {width:34px;height:34px;}
  .kpi-row {grid-template-columns:repeat(4,minmax(150px,1fr));}
}
</style>
'''
st.markdown(CSS, unsafe_allow_html=True)

st.markdown(r'''
<style>

/* ===== OMSA UI OVERRIDE ===== */

.stApp {
    background: #e6e6e6 !important;
    color: #17212a;
}

.block-container {
    max-width: 100% !important;
    padding: 0 0 3rem 0 !important;
}

/* White OMSA-style header */
.omsa-topbar {
    height: 80px !important;
    background: #ffffff !important;
    border-bottom: 1px solid #d7dadd !important;
    display: flex !important;
    align-items: center !important;
    padding: 0 22px !important;
}

/* ARMUNDIA branding */
.omsa-brand {
    width: 260px !important;
}

.omsa-brand .main {
    color: #df1748 !important;
    font-size: 20px !important;
    font-weight: 750 !important;
}

.omsa-brand .sub {
    color: #6a6a6a !important;
    font-size: 10px !important;
}

/* Horizontal navigation like OMSA */
.omsa-nav a {
    color: #111111 !important;
    font-size: 12px !important;
    padding: 13px 14px !important;
    border-radius: 8px !important;
}

.omsa-nav a:hover {
    background: #f1f2f3 !important;
}

.omsa-nav a.active {
    background: #eeeeef !important;
    color: #3f9ca3 !important;
}

/* Navy → teal OMSA section header */
.omsa-sectionbar {
    height: 60px !important;
    background: linear-gradient(
        90deg,
        #14324a 0%,
        #1c5363 45%,
        #409aa3 100%
    ) !important;
    color: white !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    padding: 0 18px !important;
}

/* Main working area */
.omsa-body {
    padding: 18px !important;
}

/* OMSA-like compact tables */
.omsa-table-wrap {
    background: white !important;
    border: 1px solid #d7dbdd !important;
    border-radius: 3px !important;
    box-shadow: none !important;
}

table.omsa-table {
    width: 100% !important;
    border-collapse: collapse !important;
    font-size: 12px !important;
}

.omsa-table th {
    background: #c8ddd8 !important;
    color: #22323a !important;
    font-size: 11px !important;
    font-weight: 500 !important;
    padding: 6px 7px !important;
    border-right: 1px solid #b9cbc7 !important;
}

.omsa-table td {
    background: white !important;
    padding: 6px 7px !important;
    height: 25px !important;
    border-top: 1px solid #e0e3e5 !important;
    border-right: 1px solid #e0e3e5 !important;
}

.omsa-table tr:hover td {
    background: #f5f7f7 !important;
}

/* OMSA blue action buttons */
.launch-btn {
    background: #3777c8 !important;
    color: white !important;
    padding: 4px 12px !important;
    border-radius: 4px !important;
    font-size: 11px !important;
    min-width: 150px !important;
    text-align: center !important;
}

.launch-btn:hover {
    background: #2d68b5 !important;
}

/* Component detail */
.info-strip {
    background: white !important;
    border: 1px solid #d2d6d8 !important;
    border-left: 3px solid #409aa3 !important;
    border-radius: 0 !important;
}

.upload-head {
    background: #c8ddd8 !important;
    border: 1px solid #bdcfcb !important;
    border-bottom: 0 !important;
    padding: 6px 9px !important;
}

/* Streamlit buttons */
.stButton > button,
.stDownloadButton > button {
    background: #3777c8 !important;
    color: white !important;
    border: 0 !important;
    border-radius: 4px !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: #2d68b5 !important;
}

/* Upload boxes */
[data-testid="stFileUploader"] {
    background: white !important;
    border: 1px solid #d5dbde !important;
    border-radius: 0 0 3px 3px !important;
}

[data-testid="stFileUploaderDropzone"] {
    background: #fafbfb !important;
    border: 1px dashed #aeb8bd !important;
}

</style>
''', unsafe_allow_html=True)

NAV = [
    ('dashboard','Dashboard'),
    ('loaders','Loaders'),
    ('views','Views'),
    ('controls','Controls'),
    ('relations','Relations'),
    ('templates','Templates'),
    ('history','Run History'),
]


def qp_get(name: str, default: str='') -> str:
    try:
        value = st.query_params.get(name, default)
        if isinstance(value, list): return value[0] if value else default
        return str(value)
    except Exception:
        q = st.experimental_get_query_params()
        v = q.get(name, [default])
        return v[0] if isinstance(v, list) else str(v)


def nav_shell(active: str) -> None:
    links=[]
    for key,label in NAV:
        cls='active' if key==active else ''
        links.append(f'<a class="{cls}" href="?page={key}" target="_self">{html.escape(label)}</a>')
    logo_path = Path(__file__).resolve().parent / 'assets' / 'armundia_group_logo.png'
    logo_b64 = base64.b64encode(logo_path.read_bytes()).decode('ascii') if logo_path.exists() else ''
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" alt="Armundia Group">' if logo_b64 else '<b>ARMUNDIA GROUP</b>'
    st.markdown(
        '<div class="omsa-topbar">'
        f'<div class="omsa-brand">{logo_html}</div>'
        f'<div class="omsa-nav">{"".join(links)}</div>'
        '<div class="omsa-dot"></div>'
        '</div>',
        unsafe_allow_html=True,
    )

def section_bar(title: str) -> None:
    st.markdown(f'<div class="omsa-sectionbar">{html.escape(title)}</div>', unsafe_allow_html=True)


def temp_uploaded(uploaded) -> Path:
    suffix=Path(uploaded.name).suffix
    fd,name=tempfile.mkstemp(suffix=suffix)
    os.close(fd)
    p=Path(name); p.write_bytes(uploaded.getbuffer()); return p


def status_label(comp: dict) -> tuple[str,str]:
    status=comp.get('implementation_status','ready')
    if not comp.get('allow_generate',True): return ('Mapping required','blocked')
    if status in {'review','needs_confirmation'}: return ('Review','review')
    return ('OK','ready')


def module_table(component_type: str, page_key: str) -> None:
    comps=load_components(component_type)
    rows=[]
    for c in comps:
        info=template_info(c['component_id'])
        template=info['filename'] if info else 'Missing template'
        status,cls=status_label(c)
        href=f'?page={page_key}&component={html.escape(c["component_id"])}'
        rows.append(
            '<tr>'
            f'<td><b>{html.escape(c["display_name"])}</b><div style="font-size:10px;color:#77818a">{html.escape(c["component_id"])}</div></td>'
            f'<td>{html.escape(template)}</td>'
            f'<td class="center">{len(c.get("inputs",[]))}</td>'
            f'<td class="center"><span class="status {cls}">{html.escape(status)}</span></td>'
            f'<td class="center"><a class="Run Configure Control-btn" target="_self" href="{href}">Run Configure Control</a></td>'
            '</tr>'
        )
    st.markdown(
        '<div class="omsa-body"><div class="omsa-table-wrap"><table class="omsa-table">'
        '<thead><tr><th>COMPONENT</th><th>ACTIVE TEMPLATE</th><th class="center">REQUIRED FILES</th><th class="center">STATUS</th><th class="center">ACTION</th></tr></thead>'
        f'<tbody>{"".join(rows)}</tbody></table></div></div>',
        unsafe_allow_html=True,
    )


def component_detail(component_id: str, page_key: str) -> None:
    comp=load_component(component_id)
    info=template_info(component_id)
    st.markdown('<div class="omsa-body">', unsafe_allow_html=True)
    st.markdown(f'<div class="breadcrumb"><a target="_self" href="?page={page_key}">← Back to {page_key.title()}</a></div>', unsafe_allow_html=True)
    if info:
        st.markdown(
            '<div class="info-strip">'
            f'<div><div class="label">Component</div><div class="value">{html.escape(comp["display_name"])}</div></div>'
            f'<div><div class="label">Active Template</div><div class="value">{html.escape(info["filename"])}</div></div>'
            f'<div><div class="label">Template Version</div><div class="value">{html.escape(str(info["active_version"]))}</div></div>'
            f'<div><div class="label">Required Files</div><div class="value">{len(comp.get("inputs",[]))}</div></div>'
            '</div>', unsafe_allow_html=True)
    else:
        st.error('The fixed master template is missing. Restore it under templates_local or register a replacement from Templates.')

    c1,c2=st.columns([5,1.2])
    with c2:
        reporting_date=st.date_input('Reference Date', value=date.today(), key=f'date_{component_id}')
    with c1:
        st.markdown('### Required Files')
        st.caption('Upload the run-specific files below. The master Excel template is fixed and is not uploaded on each run.')

    uploads={}
    all_required=True
    for idx,slot in enumerate(comp.get('inputs',[]),start=1):
        field_status=slot.get('status','confirmed')
        tag='<span class="small-tag ok">confirmed</span>' if field_status=='confirmed' else '<span class="small-tag warn">needs confirmation</span>'
        st.markdown(f'<div class="upload-head">{idx}. {html.escape(slot.get("label",slot["id"]))}{tag}</div>', unsafe_allow_html=True)
        up=st.file_uploader(
            slot.get('expected_file',slot.get('label',slot['id'])),
            type=slot.get('accepted_formats',['csv','xlsx','xls','xlsb']),
            key=f'upload_{component_id}_{slot["id"]}',
            label_visibility='collapsed',
        )
        d=slot.get('destination',{}); s=slot.get('source',{}); rules=slot.get('import_rules',{})
        source_desc=s.get('start_header') or s.get('start_column','A')
        rule_bits=[]
        if rules.get('replace_literal_NULL_with_blank'): rule_bits.append('literal NULL → blank')
        if rules.get('fill_helper_formulas_left'): rule_bits.append('preserve/extend helper formulas')
        if slot.get('import_mode')=='header_match': rule_bits.append('header-matched import')
        mapping=f"Source starts {source_desc} → {d.get('sheet')}!{d.get('start_column','A')}{d.get('start_row',2)}"
        st.markdown(f'<div class="mapping">{html.escape(mapping)}</div>', unsafe_allow_html=True)
        if rule_bits:
            st.markdown(f'<div class="rule-note">Rules: {html.escape(" · ".join(rule_bits))}</div>', unsafe_allow_html=True)
        if up: uploads[slot['id']]=up
        elif slot.get('required',True): all_required=False
        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

    if comp.get('procedure_summary'):
        with st.expander('Automation details / workbook rules'):
            for x in comp['procedure_summary']:
                st.write('•',x)
            if comp.get('notes'):
                for x in comp['notes']: st.warning(x)
            st.caption('Excel formulas remain the source of truth. The application automates file population and Excel recalculation only.')

    if comp.get('implementation_status') in {'review','needs_confirmation'}:
        st.warning('This component contains one or more mappings flagged for review. The current configuration is visible above and can be corrected as we validate it against a manual run.')
    if not comp.get('allow_generate',True):
        st.error('Generate is intentionally disabled for this template until the unresolved mapping is confirmed.')

    ready=bool(info) and all_required and comp.get('allow_generate',True)
    if st.button('Generate Excel', disabled=not ready, type='primary', use_container_width=False, key=f'generate_{component_id}'):
        local_inputs={sid:temp_uploaded(up) for sid,up in uploads.items()}
        try:
            with st.status('Generating workbook...', expanded=True) as status:
                st.write('Copying fixed master template...')
                st.write('Populating uploaded files into configured sheets...')
                st.write('Preserving and extending existing Excel formulas...')
                st.write('Running Microsoft Excel full recalculation...')
                result=generate(comp,active_template(component_id),local_inputs,reporting_date.isoformat())
                status.update(label='Generation completed',state='complete')
            save_run({**result,'component_id':component_id,'component_type':comp['component_type'],'display_name':comp['display_name'],
                      'reporting_date':reporting_date.isoformat(),'template':info['filename'] if info else None})
            st.success(f"Completed — {result['status']}")
            a,b,c,d=st.columns(4)
            a.metric('TRUE',result['true_count']); b.metric('FALSE',result['false_count']); c.metric('Formula errors',result['error_count']); d.metric('Other formula results',result['non_boolean_count'])
            output=Path(result['output_path'])
            st.download_button('Download Completed Excel',data=output.read_bytes(),file_name=output.name,mime='application/octet-stream',use_container_width=False)
        except ExcelUnavailable as exc:
            st.error(str(exc))
        except Exception as exc:
            st.error(str(exc))
            with st.expander('Technical details'):
                st.exception(exc)
    st.markdown('</div>', unsafe_allow_html=True)


def dashboard() -> None:
    loaders=load_components('loader'); views=load_components('view'); controls=load_components('control'); runs=list_runs(200)
    ready=sum(1 for c in loaders+views+controls if c.get('allow_generate',True))
    st.markdown('<div class="omsa-body">', unsafe_allow_html=True)
    st.markdown(
        '<div class="kpi-row">'
        f'<div class="kpi"><div class="v">{len(loaders)}</div><div class="l">Loader Templates</div></div>'
        f'<div class="kpi"><div class="v">{len(views)}</div><div class="l">View Templates</div></div>'
        f'<div class="kpi"><div class="v">{len(controls)}</div><div class="l">Control Templates</div></div>'
        f'<div class="kpi"><div class="v">{ready}</div><div class="l">Generation-enabled</div></div>'
        '</div>', unsafe_allow_html=True)
    st.markdown('<div class="card-lite"><b>Operating model</b><br><br>Fixed formula-based Excel template + named run-specific upload fields → automated population → native Excel recalculation → completed workbook.</div>',unsafe_allow_html=True)
    if runs:
        st.markdown('### Recent Runs')
        display=[]
        for r in runs[:25]:
            display.append({k:r.get(k) for k in ['generated_at','component_id','reporting_date','status','true_count','false_count','error_count']})
        st.dataframe(display,use_container_width=True,hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)


def relations_page() -> None:
    comps=load_components()
    rows=[]
    for c in comps:
        for slot in c.get('inputs',[]):
            src=slot.get('source_component') or slot.get('label')
            d=slot.get('destination',{})
            rows.append({'Source / Upload':src,'Target Component':c['component_id'],'Type':c['component_type'].title(),'Upload Field':slot.get('label'),
                         'Destination':f"{d.get('sheet')}!{d.get('start_column','A')}{d.get('start_row',2)}",'Status':slot.get('status','confirmed')})
    st.markdown('<div class="omsa-body">',unsafe_allow_html=True)
    st.caption('Relations documents dependencies only. Files are still uploaded manually under the target component for each run.')
    st.dataframe(rows,use_container_width=True,hide_index=True)
    st.markdown('</div>',unsafe_allow_html=True)


def templates_page() -> None:
    comps=load_components()
    st.markdown('<div class="omsa-body">',unsafe_allow_html=True)
    labels={f"[{c['component_type'].upper()}] {c['display_name']}":c for c in comps}
    selected=st.selectbox('Component',list(labels))
    comp=labels[selected]; info=template_info(comp['component_id'])
    if info: st.success(f"Active fixed template: {info['filename']} · {info['active_version']}")
    st.caption('Use this only when a newer formula-based master template must replace the current version. This is separate from run-specific input uploads.')
    version=st.text_input('New template version',value=date.today().isoformat())
    up=st.file_uploader('New master template',type=['xlsx','xlsm','xlsb','xls'],key='new_master_template')
    if st.button('Register as Active Template',disabled=up is None):
        p=temp_uploaded(up); register_template(comp['component_id'],p,up.name,version); st.success('Template registered.'); st.rerun()
    st.markdown('</div>',unsafe_allow_html=True)


def history_page() -> None:
    rows=list_runs(500)
    st.markdown('<div class="omsa-body">',unsafe_allow_html=True)
    if not rows: st.info('No generated runs yet.')
    else: st.dataframe(rows,use_container_width=True,hide_index=True)
    st.markdown('</div>',unsafe_allow_html=True)


page=qp_get('page','dashboard').lower()
if page not in {x[0] for x in NAV}: page='dashboard'
nav_shell(page)
section_titles={'dashboard':'Dashboard','loaders':'Loaders','views':'Views','controls':'Controls','relations':'Relations','templates':'Template Management','history':'Run History'}
section_bar(section_titles[page])
component=qp_get('component','')

if page=='dashboard':
    dashboard()
elif page=='loaders':
    component_detail(component,'loaders') if component else module_table('loader','loaders')
elif page=='views':
    component_detail(component,'views') if component else module_table('view','views')
elif page=='controls':
    component_detail(component,'controls') if component else module_table('control','controls')
elif page=='relations':
    relations_page()
elif page=='templates':
    templates_page()
elif page=='history':
    history_page()
