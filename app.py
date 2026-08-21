from __future__ import annotations

import html
import hmac
import base64
import os
import tempfile
from datetime import date
from functools import lru_cache
from pathlib import Path

import streamlit as st
import pandas as pd

from src.config_store import load_components, load_component
from src.template_store import active_template, template_info, register_template, list_template_versions, activate_template_version
from src.excel_engine import generate, ExcelUnavailable
from src.date_detector import detect_reference_date, validate_reference_dates
from src.history import save_run, list_runs
from src.utils import safe_name

st.set_page_config(page_title='OMSA Test Automation', page_icon='📊', layout='wide', initial_sidebar_state='collapsed')

CSS = '<style>\n:root {--navy:#12344a;--navy-2:#1d5967;--teal:#3e969f;--teal-soft:#c9ddd8;--page:#e7e8e8;--text:#16252e;--muted:#667681;--blue:#3477c8;--blue-hover:#2866b2;--amber:#996300;--red:#b42318;}\nheader[data-testid="stHeader"] {height:0;visibility:hidden;} footer {visibility:hidden;} [data-testid="stSidebar"] {display:none;}\n.stApp {background:var(--page);color:var(--text);font-size:15px;} .block-container {max-width:100%;padding:0 28px 3.2rem 28px;}\n.omsa-topbar {height:92px;background:#fff;border-bottom:1px solid #d9ddde;width:calc(100% + 56px);margin-left:-28px;display:flex;align-items:center;padding:0 30px;box-sizing:border-box;position:relative;}\n.omsa-brand {position:absolute;left:30px;top:50%;transform:translateY(-50%);width:210px;display:flex;align-items:center;} .omsa-brand img {width:175px;max-height:58px;object-fit:contain;object-position:left center;display:block;} .omsa-brand-fallback .main {font-size:22px;font-weight:750;color:#df1748;} .omsa-brand-fallback .sub {font-size:11px;color:#6a6a6a;margin-top:4px;letter-spacing:.5px;}\n.omsa-nav {position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);display:flex;gap:24px;align-items:center;justify-content:center;} .omsa-nav a {color:#111;text-decoration:none;font-size:15px;font-weight:500;padding:14px 16px;border-radius:8px;white-space:nowrap;} .omsa-nav a:hover {background:#f0f2f2;} .omsa-nav a.active {background:#eeeeef;color:#2f8d96;font-weight:650;}\n.omsa-sectionbar {height:66px;width:calc(100% + 56px);margin-left:-28px;display:flex;align-items:center;padding:0 30px;box-sizing:border-box;background:linear-gradient(90deg,var(--navy) 0%,var(--navy-2) 43%,var(--teal) 100%);color:#fff;font-size:19px;font-weight:700;border-bottom:1px solid rgba(0,0,0,.08);} .omsa-body {padding-top:20px;} div[data-testid="stVerticalBlock"] {gap:.85rem;}\n.omsa-table-wrap {background:#fff;border-radius:4px;overflow:auto;width:100%;border:1px solid #d7dbdd;box-shadow:none;} table.omsa-table {width:100%;border-collapse:collapse;font-size:14px;min-width:940px;} .omsa-table th {background:var(--teal-soft);color:#21343b;text-align:left;font-weight:650;padding:9px 10px;border-right:1px solid #b7cbc6;white-space:nowrap;font-size:13px;} .omsa-table td {background:#fff;padding:9px 10px;border-top:1px solid #e0e4e5;border-right:1px solid #e0e4e5;vertical-align:middle;min-height:38px;} .omsa-table tr:hover td {background:#f5f8f7;} .omsa-table td.center,.omsa-table th.center{text-align:center;}\n.launch-btn {display:inline-block;background:var(--blue);color:#fff!important;text-decoration:none!important;padding:7px 14px;border-radius:4px;font-size:13px;min-width:170px;text-align:center;font-weight:600;} .launch-btn:hover {background:var(--blue-hover);} .status {font-style:italic;font-weight:650;} .status.ready {color:#202629;} .status.review {color:var(--amber);} .status.blocked {color:var(--red);}\n.breadcrumb {font-size:14px;margin:18px 0 12px 0;} .breadcrumb a {color:#2e6ea7;text-decoration:none;font-weight:550;} .info-strip {background:#fff;border:1px solid #d2d7d9;border-left:4px solid var(--teal);padding:14px 16px;margin-bottom:16px;display:flex;gap:34px;align-items:center;width:100%;box-sizing:border-box;} .info-strip .label {color:var(--muted);font-size:12px;text-transform:uppercase;letter-spacing:.45px;} .info-strip .value {font-size:15px;font-weight:650;margin-top:3px;} .upload-head {background:var(--teal-soft);padding:10px 12px;font-size:14px;font-weight:700;border:1px solid #bdcfcb;border-bottom:0;width:100%;box-sizing:border-box;} .mapping,.rule-note {font-size:13px;color:#53636c;margin:5px 0 0 2px;} .small-tag {display:inline-block;padding:3px 8px;border-radius:12px;background:#eef2f4;color:#52606a;font-size:11px;margin-left:8px;} .small-tag.warn {background:#fff3d6;color:#8a5a00;} .small-tag.ok {background:#e8f5ed;color:#216e47;}\n.dashboard-intro {display:flex;justify-content:space-between;align-items:flex-end;gap:22px;margin:22px 0 18px 0;} .dashboard-intro h1 {font-size:27px;line-height:1.2;margin:0;color:#17313f;} .dashboard-intro p {font-size:15px;color:#667681;margin:7px 0 0 0;} .dashboard-badge {background:#fff;border:1px solid #d4dadc;border-radius:4px;padding:10px 14px;font-size:13px;color:#596a74;white-space:nowrap;}\n.kpi-row {display:grid;grid-template-columns:repeat(3,minmax(220px,1fr));gap:14px;width:100%;margin-bottom:18px;} .kpi-link {text-decoration:none!important;color:inherit!important;} .kpi {background:#fff;border:1px solid #d2d8da;padding:18px 18px 16px;border-radius:4px;min-height:112px;box-sizing:border-box;transition:.15s ease;} .kpi-link:hover .kpi {transform:translateY(-1px);border-color:#9dbec0;box-shadow:0 3px 10px rgba(26,61,72,.08);} .kpi.active {border-top:4px solid var(--teal);padding-top:15px;background:#fbfdfd;} .kpi .v {font-size:31px;font-weight:750;color:#18323f;line-height:1;} .kpi .l {font-size:15px;font-weight:650;color:#344a55;margin-top:9px;} .kpi .hint {font-size:12px;color:#7a888f;margin-top:6px;line-height:1.35;} .card-lite {background:#fff;border:1px solid #d5dbde;padding:16px 18px;border-radius:4px;width:100%;box-sizing:border-box;font-size:14px;} .dashboard-section-title {font-size:19px;font-weight:750;color:#18323f;margin:24px 0 5px 0;} .dashboard-section-sub {font-size:14px;color:#6b7981;margin-bottom:14px;}\n.network-grid {display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;width:100%;} .network-card {background:#fff;border:1px solid #d3d9db;border-radius:5px;padding:15px 16px;min-height:170px;box-sizing:border-box;} .network-card:hover {border-color:#a9c7c7;box-shadow:0 2px 8px rgba(26,61,72,.06);} .network-title-row {display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:13px;} .network-title {font-size:16px;font-weight:750;color:#173542;text-decoration:none!important;} .network-title:hover {color:#2d7d86;} .network-template {font-size:12px;color:#6d7b82;margin-top:3px;max-width:520px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;} .network-count {background:#edf5f3;border:1px solid #c8ddd8;border-radius:12px;padding:4px 9px;font-size:11px;color:#35656a;white-space:nowrap;} .flow-line {display:grid;grid-template-columns:minmax(0,1fr) 34px minmax(150px,.55fr);gap:9px;align-items:center;} .input-node-wrap {display:flex;flex-wrap:wrap;gap:7px;align-items:center;} .input-node {background:#f6f8f8;border:1px solid #d5dcde;border-radius:3px;padding:6px 8px;font-size:12px;color:#354b56;line-height:1.2;} .input-node strong {display:block;font-size:11px;color:#22717a;font-weight:650;margin-bottom:2px;} .flow-arrow {font-size:25px;color:#80a6a8;text-align:center;font-weight:300;} .template-node {background:linear-gradient(135deg,#17384b,#347d84);color:#fff;border-radius:4px;padding:12px 10px;text-align:center;font-size:12px;line-height:1.3;min-height:58px;display:flex;align-items:center;justify-content:center;} .network-footer {margin-top:12px;padding-top:10px;border-top:1px solid #edf0f1;font-size:12px;color:#748187;display:flex;justify-content:space-between;gap:10px;}\n.stButton>button {background:var(--blue);color:#fff;border:0;border-radius:4px;font-size:14px;padding:.58rem 1.05rem;font-weight:600;} .stButton>button:hover {background:var(--blue-hover);color:#fff;border:0;} .stDownloadButton>button {background:var(--blue);color:#fff;border:0;border-radius:4px;font-size:14px;} [data-testid="stFileUploader"] {background:#fff;border:1px solid #d5dbde;padding:11px;border-radius:0 0 4px 4px;width:100%;box-sizing:border-box;} [data-testid="stFileUploaderDropzone"] {background:#fafbfb;border:1px dashed #aeb8bd;min-height:88px;} [data-testid="stDateInput"] {max-width:240px;} [data-testid="stDataFrame"] {background:#fff;border:1px solid #d7dbdd;} .stSelectbox,.stTextInput,.stDateInput {font-size:14px;} .stAlert {border-radius:4px;}\n@media (max-width:1100px) {.omsa-brand {width:220px;min-width:220px;} .omsa-brand img {width:150px;} .omsa-nav {gap:4px;} .omsa-nav a {font-size:13px;padding:12px 9px;} .kpi-row {grid-template-columns:repeat(2,1fr);} .network-grid {grid-template-columns:1fr;}}\n</style>'
st.markdown(CSS, unsafe_allow_html=True)


def require_login() -> None:
    if st.session_state.get("omsa_authenticated", False):
        return

    logo_path = Path(__file__).resolve().parent / "assets" / "armundia_group_logo.png"
    bg_path = Path(__file__).resolve().parent / "assets" / "omsa_login_background.png"

    logo_html = ""
    if logo_path.exists():
        encoded = base64.b64encode(logo_path.read_bytes()).decode("ascii")
        logo_html = f'<img src="data:image/png;base64,{encoded}" alt="Armundia Group">'

    bg_css = ""
    if bg_path.exists():
        bg_encoded = base64.b64encode(bg_path.read_bytes()).decode("ascii")
        bg_css = f"url('data:image/png;base64,{bg_encoded}')"

    st.markdown(
        f'''
        <style>
        [data-testid="stHeader"],
        [data-testid="stToolbar"],
        #MainMenu,
        footer {{
            display: none !important;
        }}

        .stApp {{
            background:
                linear-gradient(rgba(12,49,60,.20), rgba(7,33,43,.35)),
                {bg_css} center center / cover no-repeat fixed !important;
        }}

        .block-container {{
            max-width: 100% !important;
            padding: 0 !important;
        }}

        .omsa-login-header {{
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            height: 82px;
            background: #ffffff;
            border-bottom: 1px solid #d9dede;
            z-index: 100;
            display: flex;
            align-items: center;
            padding: 0 30px;
            box-sizing: border-box;
        }}

        .omsa-login-header img {{
            width: 145px;
            max-height: 57px;
            object-fit: contain;
        }}

        .omsa-login-title-wrap {{
            position: fixed;
            top: 31%;
            left: 50%;
            transform: translate(-50%, -50%);
            width: min(620px, 84vw);
            text-align: center;
            z-index: 20;
            color: #ffffff;
            text-shadow: 0 2px 9px rgba(0,0,0,.25);
        }}

        .omsa-login-title {{
            font-size: 34px;
            line-height: 1.15;
            font-weight: 650;
            margin-bottom: 8px;
        }}

        .omsa-login-subtitle {{
            font-size: 15px;
            font-weight: 400;
            opacity: .92;
        }}

        [data-testid="stForm"] {{
            position: fixed !important;
            top: 43% !important;
            left: 50% !important;
            transform: translateX(-50%) !important;
            width: min(430px, 82vw) !important;
            background: transparent !important;
            border: none !important;
            padding: 0 !important;
            box-shadow: none !important;
            z-index: 30 !important;
        }}

        [data-testid="stTextInput"] label {{
            color: #ffffff !important;
            font-size: 14px !important;
            font-weight: 550 !important;
            text-shadow: 0 1px 4px rgba(0,0,0,.28);
        }}

        [data-testid="stTextInput"] input {{
            height: 52px !important;
            background: rgba(15,44,54,.72) !important;
            color: #ffffff !important;
            border: 1px solid rgba(255,255,255,.72) !important;
            border-radius: 5px !important;
            font-size: 16px !important;
            padding-left: 16px !important;
            backdrop-filter: blur(5px);
        }}

        [data-testid="stTextInput"] input::placeholder {{
            color: rgba(255,255,255,.72) !important;
        }}

        [data-testid="stFormSubmitButton"] button {{
            height: 50px !important;
            width: 100% !important;
            margin-top: 10px !important;
            background: #2d8790 !important;
            color: #ffffff !important;
            border: 1px solid rgba(255,255,255,.34) !important;
            border-radius: 5px !important;
            font-size: 16px !important;
            font-weight: 600 !important;
            box-shadow: 0 2px 8px rgba(0,0,0,.14) !important;
        }}

        [data-testid="stFormSubmitButton"] button:hover {{
            background: #236f78 !important;
            border-color: #ffffff !important;
        }}

        [data-testid="stAlert"] {{
            position: fixed !important;
            top: calc(43% + 145px) !important;
            left: 50% !important;
            transform: translateX(-50%) !important;
            width: min(430px, 82vw) !important;
            z-index: 40 !important;
            border-radius: 4px !important;
        }}
        </style>

        <div class="omsa-login-header">
            {logo_html}
        </div>

        <div class="omsa-login-title-wrap">
            <div class="omsa-login-title">OMSA Test Automation</div>
            <div class="omsa-login-subtitle">Secure access to Loader, View and Control automation</div>
        </div>
        ''',
        unsafe_allow_html=True,
    )

    with st.form("omsa_login_form"):
        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter application password",
        )
        submitted = st.form_submit_button(
            "Sign In",
            use_container_width=True,
        )

    if submitted:
        expected = st.secrets.get("APP_PASSWORD", "")
        if not expected:
            st.error("Application password has not been configured.")
        elif hmac.compare_digest(password, expected):
            st.session_state["omsa_authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")

    st.stop()


require_login()


NAV = [
    ('dashboard','Dashboard'),
    ('loaders','Loaders'),
    ('views','Views'),
    ('controls','Controls'),
    ('relations','Relations'),
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



@lru_cache(maxsize=1)
def cached_logo_data_uri() -> str | None:
    logo_path = Path(__file__).resolve().parent / 'assets' / 'armundia_group_logo.png'
    if not logo_path.exists():
        return None
    encoded = base64.b64encode(logo_path.read_bytes()).decode('ascii')
    return f'data:image/png;base64,{encoded}'

def nav_shell(active: str) -> None:
    links=[]
    for key,label in NAV:
        cls='active' if key==active else ''
        links.append(f'<a class="{cls}" href="?page={key}" target="_self">{html.escape(label)}</a>')

    logo_uri = cached_logo_data_uri()
    if logo_uri:
        brand = f'<div class="omsa-brand"><img src="{logo_uri}" alt="Armundia Group"></div>'
    else:
        brand = '<div class="omsa-brand omsa-brand-fallback"><div><div class="main">ARMUNDIA GROUP</div><div class="sub">OMSA TEST AUTOMATION</div></div></div>'

    st.markdown(
        '<div class="omsa-topbar">'
        f'{brand}'
        f'<div class="omsa-nav">{"".join(links)}</div></div>',
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
            f'<td class="center"><a class="launch-btn" target="_self" href="{href}">Run Configure Control</a></td>'
            '</tr>'
        )
    st.markdown(
        '<div class="omsa-body"><div class="omsa-table-wrap"><table class="omsa-table">'
        '<thead><tr><th>COMPONENT</th><th>ACTIVE TEMPLATE</th><th class="center">REQUIRED FILES</th><th class="center">STATUS</th><th class="center">RUN / CONFIGURE CONTROL</th></tr></thead>'
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
        st.error('The active master template is missing for this component.')

    # Template management lives inside each component rather than in a separate module.
    with st.expander('Template Management · Change / Version Master Template', expanded=False):
        if info and active_template(component_id):
            current_path=Path(active_template(component_id))
            tc1,tc2=st.columns([2.2,1])
            with tc1:
                st.markdown(f'**Current active template:** `{info["filename"]}`  \n**Version:** `{info["active_version"]}`')
                st.caption('Formula/business-logic changes can normally be delivered by uploading a new Excel template version. Existing versions are retained.')
            with tc2:
                try:
                    st.download_button('Download Active Template', data=current_path.read_bytes(), file_name=current_path.name,
                                       mime='application/octet-stream', key=f'download_template_{component_id}', use_container_width=True)
                except Exception:
                    pass

        new_template=st.file_uploader(
            'Upload new formula-based master template',
            type=['xlsx','xlsm','xlsb','xls'],
            key=f'new_master_template_{component_id}',
            help='This changes the master template, not the run-specific source files below.'
        )
        version=st.text_input('Template version', value=date.today().isoformat(), key=f'template_version_{component_id}')
        if new_template is not None:
            current_suffix=Path(info['filename']).suffix.lower() if info else ''
            new_suffix=Path(new_template.name).suffix.lower()
            if current_suffix and new_suffix != current_suffix:
                st.warning(f'The new template extension ({new_suffix}) differs from the active template ({current_suffix}). Confirm that this is intentional.')
        if st.button('Upload and Set as Active Template', disabled=new_template is None, key=f'activate_new_template_{component_id}'):
            p=temp_uploaded(new_template)
            register_template(component_id,p,new_template.name,version)
            st.success('New template version uploaded and set as active. Previous versions were retained.')
            st.rerun()

        versions=list_template_versions(component_id)
        if versions:
            st.markdown('**Template version history**')
            st.dataframe([
                {'Version':v['version'],'File':v['filename'],'Active':'Yes' if v['active'] else 'No','Modified':v['modified_at']}
                for v in versions
            ], use_container_width=True, hide_index=True)
            option_rows=[v for v in versions if not v['active']]
            if option_rows:
                option_map={f"{v['version']} · {v['filename']}":v for v in option_rows}
                selected_version=st.selectbox('Restore a previous template version', ['—']+list(option_map), key=f'restore_version_{component_id}')
                if st.button('Set Selected Version Active', disabled=selected_version=='—', key=f'restore_template_{component_id}'):
                    chosen=option_map[selected_version]
                    activate_template_version(component_id,chosen['path'],chosen['version'],chosen['filename'])
                    st.success('Selected template version is now active.')
                    st.rerun()

    # This placeholder is physically above the upload fields, but it is populated after
    # the files have been inspected so Reference Date always comes from the run files.
    date_placeholder=st.empty()

    st.markdown('### Required Files')
    st.caption('Upload each named run-specific source file. These files populate a copy of the active Excel template; they do not replace the template itself.')

    uploads={}
    local_inputs={}
    date_detections={}
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
        if rules.get('protect_existing_formula_columns',True): rule_bits.append('protect existing template formula columns')
        if slot.get('import_mode')=='header_match': rule_bits.append('header-matched import')
        mapping=f"Source starts {source_desc} → {d.get('sheet')}!{d.get('start_column','A')}{d.get('start_row',2)}"
        st.markdown(f'<div class="mapping">{html.escape(mapping)}</div>', unsafe_allow_html=True)
        if rule_bits:
            st.markdown(f'<div class="rule-note">Rules: {html.escape(" · ".join(rule_bits))}</div>', unsafe_allow_html=True)
        if up:
            uploads[slot['id']]=up
            tmp=temp_uploaded(up)
            local_inputs[slot['id']]=tmp
            date_detections[slot['id']]=detect_reference_date(tmp,slot)
        elif slot.get('required',True):
            all_required=False
        st.markdown('<div style="height:8px"></div>', unsafe_allow_html=True)

    date_summary=validate_reference_dates(comp,date_detections)
    reference_date=date_summary.get('reference_date')
    previous_date=date_summary.get('previous_date')
    with date_placeholder.container():
        st.markdown('### Reference Date Validation')
        if not date_detections:
            st.info('Reference Date: —  ·  Upload run files and the date will be detected automatically from configured date columns or filenames.')
        else:
            dc1,dc2,dc3=st.columns([1.2,1.2,3])
            dc1.metric('Reference Date (T)', pd.Timestamp(reference_date).strftime('%d/%m/%Y') if reference_date else '—')
            dc2.metric('Previous Date (T-1)', pd.Timestamp(previous_date).strftime('%d/%m/%Y') if previous_date else '—')
            dc3.metric('Date Check', 'Valid' if date_summary['is_valid'] else 'Mismatch / Review')
            detail_rows=[]
            slots={slot['id']:slot for slot in comp.get('inputs',[])}
            for sid,result in date_detections.items():
                slot=slots[sid]
                detail_rows.append({
                    'File':slot.get('label',sid),
                    'Role':slot.get('role','T'),
                    'Detected Date':result.get('display_date','—'),
                    'Detected From':result.get('source','—'),
                })
            with st.expander('Show detected dates by uploaded file', expanded=False):
                st.dataframe(detail_rows,use_container_width=True,hide_index=True)
            for warning in date_summary.get('warnings',[]):
                st.warning(warning)
            for error in date_summary.get('errors',[]):
                st.error(error)

    if comp.get('procedure_summary'):
        with st.expander('Automation details / workbook rules'):
            for x in comp['procedure_summary']:
                st.write('•',x)
            if comp.get('notes'):
                for x in comp['notes']: st.warning(x)
            st.caption('Excel formulas remain the source of truth. The application automates controlled copy/paste, formula preservation/fill-down and native Excel recalculation only.')

    if comp.get('implementation_status') in {'review','needs_confirmation'}:
        st.warning('This component contains one or more mappings flagged for review. The current configuration is visible above and can be corrected as we validate it against a manual run.')
    if not comp.get('allow_generate',True):
        st.error('Generate is intentionally disabled for this template until the unresolved mapping is confirmed.')

    ready=bool(info) and all_required and comp.get('allow_generate',True) and date_summary['is_valid']
    if st.button('Generate Excel', disabled=not ready, type='primary', use_container_width=False, key=f'generate_{component_id}'):
        try:
            with st.status('Generating workbook...', expanded=True) as status:
                st.write('Creating a run-specific copy of the active master template...')
                st.write('Populating uploaded files into configured sheets/ranges...')
                st.write('Protecting existing template formula columns...')
                st.write('Extending existing helper/check formulas where configured...')
                st.write('Running Microsoft Excel full recalculation...')
                result=generate(comp,active_template(component_id),local_inputs,reference_date)
                status.update(label='Generation completed',state='complete')
            save_run({**result,'component_id':component_id,'component_type':comp['component_type'],'display_name':comp['display_name'],
                      'reporting_date':reference_date,'previous_reporting_date':previous_date,'template':info['filename'] if info else None,
                      'template_version':info['active_version'] if info else None})
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

def dashboard_network(group: str) -> None:
    group = group if group in {'loader','view','control'} else 'loader'
    page_key = {'loader':'loaders','view':'views','control':'controls'}[group]
    comps = load_components(group)
    group_label = {'loader':'Loaders','view':'Views','control':'Controls'}[group]
    st.markdown(f'<div class="dashboard-section-title">{group_label} — Template & Input Network</div>', unsafe_allow_html=True)
    st.markdown('<div class="dashboard-section-sub">Each input node is a separate upload field. The arrow shows which fixed Excel template those files populate. Click a component name to open its run screen.</div>', unsafe_allow_html=True)
    cards=[]
    for comp in comps:
        info=template_info(comp['component_id'])
        template=info['filename'] if info else comp.get('template',{}).get('suggested_filename','Template missing')
        input_nodes=[]
        for slot in comp.get('inputs',[]):
            source_component=slot.get('source_component')
            label=slot.get('label',slot.get('id','Input'))
            if source_component:
                node=f'<span class="input-node"><strong>{html.escape(str(source_component))}</strong>{html.escape(str(label))}</span>'
            else:
                node=f'<span class="input-node">{html.escape(str(label))}</span>'
            input_nodes.append(node)
        if not input_nodes:
            input_nodes=['<span class="input-node">No configured inputs</span>']
        status,_=status_label(comp)
        href=f'?page={page_key}&component={html.escape(comp["component_id"])}'
        cards.append('<div class="network-card"><div class="network-title-row">' + f'<div><a class="network-title" target="_self" href="{href}">{html.escape(comp["display_name"])}</a><div class="network-template">{html.escape(template)}</div></div>' + f'<div class="network-count">{len(comp.get("inputs",[]))} upload fields</div></div>' + '<div class="flow-line">' + f'<div class="input-node-wrap">{"".join(input_nodes)}</div><div class="flow-arrow">→</div><div class="template-node">{html.escape(comp["display_name"])}<br>Fixed Excel Template</div></div>' + f'<div class="network-footer"><span>Status: {html.escape(status)}</span><span>Excel formulas remain source of truth</span></div></div>')
    st.markdown(f'<div class="network-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def dashboard() -> None:
    loaders=load_components('loader'); views=load_components('view'); controls=load_components('control'); runs=list_runs(20)
    explore=qp_get('explore','loader').lower()
    if explore not in {'loader','view','control'}:
        explore='loader'
    st.markdown('<div class="dashboard-intro"><div><h1>OMSA Test Automation</h1><p>Select a template group to explore the components and the files required to populate each workbook.</p></div><div class="dashboard-badge">Fixed templates · Run-specific uploads · Native Excel logic</div></div>', unsafe_allow_html=True)
    def card(key: str, value: int, label: str, hint: str) -> str:
        active=' active' if explore==key else ''
        return f'<a class="kpi-link" target="_self" href="?page=dashboard&explore={key}"><div class="kpi{active}"><div class="v">{value}</div><div class="l">{label}</div><div class="hint">{hint}</div></div></a>'
    st.markdown('<div class="kpi-row">' + card('loader',len(loaders),'Loader Templates','Click to see loaders and their required upload files') + card('view',len(views),'View Templates','Click to see views and their Loader/View inputs') + card('control',len(controls),'Control Templates','Click to see controls and their required inputs') + '</div>', unsafe_allow_html=True)
    st.markdown('<div class="card-lite"><b>Operating model</b><br><span style="display:inline-block;margin-top:7px">Choose component → upload every named run-specific source file → populate a copy of the fixed Excel template → extend existing helper/check formulas where configured → recalculate in Microsoft Excel → download the completed workbook.</span></div>', unsafe_allow_html=True)
    dashboard_network(explore)
    if runs:
        st.markdown('<div class="dashboard-section-title">Recent Runs</div>',unsafe_allow_html=True)
        display=[]
        for r in runs[:20]:
            display.append({k:r.get(k) for k in ['generated_at','component_id','reporting_date','status','true_count','false_count','error_count']})
        st.dataframe(display,use_container_width=True,hide_index=True)



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
