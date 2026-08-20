from __future__ import annotations

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
  --navy:#18364b;
  --teal:#3f99a2;
  --teal-soft:#c9ded9;
  --page:#e9eaec;
  --line:#d3d9dc;
  --text:#18222b;
  --muted:#66727d;
  --blue:#3478cf;
  --blue-hover:#2868b8;
  --green:#2c7a52;
  --amber:#a96a00;
  --red:#b42318;
}
header[data-testid="stHeader"] {height:0; visibility:hidden;}
footer {visibility:hidden;}
[data-testid="stSidebar"] {display:none;}
.stApp {background: var(--page); color:var(--text);}
.block-container {max-width:100%; padding:0 26px 3rem 26px;}

/* top white header */
.omsa-topbar {
  height: 86px; background:#fff; border-bottom:1px solid #d9dde0;
  margin-left:-26px; margin-right:-26px; width:calc(100% + 52px);
  display:flex; align-items:center; padding:0 28px; gap:34px;
}
.omsa-brand {width:245px; line-height:1.0;}
.omsa-brand .main {font-size:25px; font-weight:800; letter-spacing:.2px; color:#e4003b;}
.omsa-brand .sub {font-size:12px; font-weight:600; color:#555; margin-top:5px; letter-spacing:.4px;}
.omsa-nav {display:flex; gap:4px; align-items:center; margin-left:auto; margin-right:72px;}
.omsa-nav a {color:#111; text-decoration:none; font-size:14px; padding:14px 16px; border-radius:10px;}
.omsa-nav a:hover {background:#f1f2f3;}
.omsa-nav a.active {background:#f0f1f2; color:#3a9d9d;}
.omsa-dot {width:39px;height:39px;background:#df001f;border-radius:50%;position:absolute;right:30px;top:24px;}

.omsa-sectionbar {
  height:62px; display:flex; align-items:center; padding:0 28px;
  margin-left:-26px; margin-right:-26px; width:calc(100% + 52px);
  background:linear-gradient(90deg,var(--navy) 0%, #235d6a 45%, var(--teal) 100%);
  color:#fff; font-size:18px; font-weight:700; border-bottom:1px solid rgba(0,0,0,.08);
}
.omsa-body {padding:18px 0 34px 0;}

/* remove Streamlit default vertical gaps near custom shell */
div[data-testid="stVerticalBlock"] {gap:.65rem;}

.omsa-table-wrap {background:#fff; border-radius:4px; overflow:hidden; max-width:1500px; border:1px solid #d8dddf;}
table.omsa-table {width:100%; border-collapse:collapse; font-size:13px;}
.omsa-table th {background:var(--teal-soft); color:#22323a; text-align:left; font-weight:600; padding:8px 10px; border-right:1px solid #bfcfcb; white-space:nowrap;}
.omsa-table td {background:#fff; padding:8px 10px; border-top:1px solid #e1e5e7; border-right:1px solid #e1e5e7; vertical-align:middle;}
.omsa-table tr:hover td {background:#f8fafb;}
.omsa-table td.center,.omsa-table th.center{text-align:center;}
.launch-btn {display:inline-block;background:var(--blue);color:#fff!important;text-decoration:none!important;padding:5px 20px;border-radius:4px;font-size:12px;min-width:74px;text-align:center;}
.launch-btn:hover {background:var(--blue-hover);}
.status {font-style:italic;font-weight:600;}
.status.ready {color:#212121;}
.status.review {color:var(--amber);}
.status.blocked {color:var(--red);}

.breadcrumb {font-size:13px;margin:0 0 10px 0;}
.breadcrumb a {color:#2e6ea7;text-decoration:none;}
.info-strip {background:#fff;border:1px solid #d5dbde;border-left:4px solid var(--teal);padding:11px 14px;margin-bottom:12px;display:flex;gap:28px;align-items:center;}
.info-strip .label {color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.4px;}
.info-strip .value {font-size:13px;font-weight:600;margin-top:2px;}
.upload-head {background:var(--teal-soft);padding:8px 12px;font-size:13px;font-weight:650;border:1px solid #c0d2ce;border-bottom:0;}
.mapping {font-size:12px;color:#59656e;margin-top:4px;}
.rule-note {font-size:12px;color:#59656e;margin-top:2px;}
.small-tag {display:inline-block;padding:2px 7px;border-radius:10px;background:#eef2f4;color:#52606a;font-size:10px;margin-left:7px;}
.small-tag.warn {background:#fff3d6;color:#8a5a00;}
.small-tag.ok {background:#e8f5ed;color:#216e47;}

.card-lite {background:#fff;border:1px solid #d5dbde;padding:14px 16px;border-radius:4px;}
.kpi-row {display:grid;grid-template-columns:repeat(4,minmax(140px,1fr));gap:12px;max-width:1100px;margin-bottom:14px;}
.kpi {background:#fff;border:1px solid #d5dbde;padding:14px 16px;}
.kpi .v{font-size:26px;font-weight:700;color:#1d2b34}.kpi .l{font-size:12px;color:#68747d;margin-top:2px;}

/* Streamlit controls */
.stButton>button {background:var(--blue);color:#fff;border:0;border-radius:4px;font-size:13px;padding:.48rem .95rem;}
.stButton>button:hover {background:var(--blue-hover);color:#fff;border:0;}
.stDownloadButton>button {background:var(--blue);color:#fff;border:0;border-radius:4px;}
[data-testid="stFileUploader"] {background:#fff;border:1px solid #d5dbde;padding:12px;border-radius:0 0 4px 4px;}
[data-testid="stFileUploaderDropzone"] {background:#fafbfb;border:1px dashed #aeb8bd;}
[data-testid="stDateInput"] {max-width:220px;}
.stAlert {border-radius:4px;}
</style>
'''
st.markdown(CSS, unsafe_allow_html=True)

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
    st.markdown(
        '<div class="omsa-topbar">'
        '<div class="omsa-brand"><div class="main">OMSA</div><div class="sub">TEST AUTOMATION</div></div>'
        f'<div class="omsa-nav">{"".join(links)}</div><div class="omsa-dot"></div></div>',
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
            f'<td class="center"><a class="launch-btn" target="_self" href="{href}">Launch</a></td>'
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
