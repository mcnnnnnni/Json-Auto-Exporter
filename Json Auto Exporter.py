import streamlit as st
import pandas as pd
import json
import datetime
import re
import io


# 页面更宽

# Set page config: left-aligned, wide, English
st.set_page_config(page_title="JSON Auto Exporter", layout="wide")

# Custom CSS for left alignment and modern look
st.markdown("""
    <style>
    .main .block-container {
        max-width: 1100px;
        margin-left: 0 !important;
        margin-right: auto !important;
        padding-left: 32px;
        padding-right: 32px;
    }
    .stApp {
        text-align: left !important;
        background: #fff;
    }
    .stButton > button {
        background: #fff;
        color: #1a237e;
        border: 1.5px solid #1a237e;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5em 1.5em;
        transition: background 0.2s, color 0.2s;
    }
    .stButton > button:hover {
        background: #1a237e;
        color: #fff;
    }
    .stDownloadButton > button {
        background: #fff;
        color: #1a237e;
        border: 1.5px solid #1a237e;
        border-radius: 8px;
        font-weight: 600;
        padding: 0.5em 1.5em;
        margin-top: 0.5em;
        margin-bottom: 0.5em;
        transition: background 0.2s, color 0.2s;
    }
    .stDownloadButton > button:hover {
        background: #1a237e;
        color: #fff;
    }
    .stTable, .stDataFrame {
        background: #fff;
        border-radius: 10px;
        box-shadow: 0 2px 12px #e0e7ef;
        padding: 8px 16px;
    }
    .stTabs [data-baseweb="tab-list"] {
        justify-content: flex-start;
    }
    .card-section {
        background: #fff;
        border-radius: 0;
        box-shadow: none;
        padding: 0 0 18px 0;
        margin-bottom: 2px;
        border: none;
        max-width: 900px;
        margin-left: 0;
        margin-right: auto;
    }
    .step-title {
        font-size: 1.25em;
        font-weight: 700;
        color: #1a237e;
        margin-bottom: 2px;
        letter-spacing: 0.5px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .step-desc {
        color: #333;
        font-size: 0.98em;
        margin-bottom: 2px;
    }
    .step-num {
        display: inline-block;
        background: #e3e8f0;
        color: #1a237e;
        border-radius: 50%; 
        width: 1.6em;
        height: 1.6em;
        text-align: center;
        line-height: 1.6em;
        font-size: 0.95em;
        font-weight: bold;
        margin-right: 8px;
    }
    .small-font { font-size: 0.95em; }
    </style>
""", unsafe_allow_html=True)

st.title("📝 JSON Auto Exporter")

# ====== 上传区美化 ======

# Upload section (English, left-aligned)
st.markdown("""
<div class='card-section'>
    <div class='step-title'><span class='step-num'>1</span>📤 Upload Excel File</div>
    <div class='step-desc small-font'>
        Please upload your Excel file (<b>.xlsx</b>, <b>.xls</b>) and click <b>Start</b> to process. You will get categorized JSON preview and downloads below.
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Excel File", type=["xlsx", "xls"], key="uploaded_file")

# Start按钮右下角
import streamlit.components.v1 as components
if uploaded_file:
    # 重新渲染文件名和Start按钮在一行
    file_display_col, start_btn_col = st.columns([8,1], gap="small")
    with file_display_col:
        # 复用st的文件名展示
        st.markdown(f"<div style='display:flex;align-items:center;gap:8px;'><span style='font-size:0.98em;'>📄 {uploaded_file.name} <span style='color:#888;font-size:0.92em;'>({round(uploaded_file.size/1024/1024,1)}MB)</span></span></div>", unsafe_allow_html=True)
    with start_btn_col:
        start_clicked = st.button("🚀 Start", key="start_btn")
else:
    start_clicked = False

st.markdown("</div>", unsafe_allow_html=True)


start_processing = False
if uploaded_file:
    if start_clicked:
        st.session_state["start_processing"] = True
    start_processing = st.session_state.get("start_processing", False)
else:
    st.session_state["start_processing"] = False

if not uploaded_file or not st.session_state.get("start_processing", False):
    st.stop()



def get_info_fields(df, info_fields):
    info = {}
    for i in range(min(30, len(df))):
        key = str(df.iloc[i,0]).strip()
        val = str(df.iloc[i,1]).strip() if df.shape[1]>1 else ""
        if key in info_fields:
            info[key] = val
    return info

def get_generated_on(xls=None, uploaded_file=None):
    # 统一返回当前时刻，格式为m/d/Y h:M:S AM/PM（兼容Windows，去掉-）
    now = datetime.datetime.now()
    # %I 是12小时制，去除前导零
    hour = str(int(now.strftime("%I")))
    formatted = now.strftime(f"%m/%d/%Y {hour}:%M:%S %p")
    # 去除月日的前导零
    if formatted.startswith("0"): formatted = formatted[1:]
    formatted = formatted.replace("/0", "/")
    return formatted

def extract_table(sheet_name, xls, uploaded_file, info, generated_on, table_name, filename, extra_fields=None):
    try:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=str)
        header_row_idx = None
        # 只找 row nbr
        for i in range(len(raw)):
            cell_val = str(raw.iloc[i,0]).strip().lower()
            if cell_val == "row nbr":
                header_row_idx = i
                break
        if header_row_idx is None:
            return None, f"Suspected sheet but no matching table found (missing 'row nbr'): {sheet_name}"
        header_row = raw.iloc[header_row_idx]
        valid_cols = []
        for idx, col in enumerate(header_row):
            col_str = str(col).strip()
            if col_str == "" or col_str == "nan":
                break
            valid_cols.append(idx)
            if col_str == "Notes":
                break
        df = raw.iloc[header_row_idx+1:, valid_cols]
        df.columns = header_row[valid_cols]
        df = df.dropna(how='all').fillna("").astype(str)

        # 检查主字段缺失并收集warning
        # 提取Version字段（所有类型都支持）
        version_val = "1"
        for i in range(min(30, len(raw))):
            for j in range(raw.shape[1]-1):
                cell = str(raw.iloc[i, j]).strip().lower()
                if "version" in cell:
                    version_val = str(raw.iloc[i, j+1]).strip()
        if table_name in ["CFM", "Power"]:
            main_fields = ["Document Number", "Document Revision", "Part Number", "Role", "Subrole"]
            json_main = {
                "Document Number": info.get("Item Number", ""),
                "Document Revision": info.get("Part Revision", ""),
                "Part Number": extra_fields["Part Number"] if extra_fields and "Part Number" in extra_fields else "",
                "Role": info.get("Role", ""),
                "Subrole": info.get("Subrole", ""),
            }
        else:
            main_fields = ["Document Number", "Document Revision", "Child Part Number", "Parent Part Number", "Device", "Role", "Subrole"]
            # 提取Subrole、Child Part Number
            subrole_header = ""
            for i in range(min(30, len(raw))):
                for j in range(raw.shape[1]-1):
                    cell = str(raw.iloc[i, j]).strip().lower()
                    if "subrole" in cell:
                        subrole_header = str(raw.iloc[i, j+1]).strip()
            child_part_number = ""
            try:
                bom_df = st.session_state.get("bom_report_level1")
                device_kw = str(extra_fields["Device"]).strip() if extra_fields and "Device" in extra_fields else ""
                subrole_kw = subrole_header
                if bom_df is not None and device_kw:
                    for idx, row in bom_df.iterrows():
                        desc = str(row.get("Part Description", "")).lower()
                        if device_kw.lower() in desc and (not subrole_kw or subrole_kw.lower() in desc):
                            child_part_number = str(row.get("Part Number", ""))
                            break
            except Exception:
                child_part_number = extra_fields["Child Part Number"] if extra_fields and "Child Part Number" in extra_fields else ""
            json_main = {
                "Document Number": info.get("Item Number", ""),
                "Document Revision": info.get("Part Revision", ""),
                "Child Part Number": child_part_number,
                "Parent Part Number": extra_fields["Parent Part Number"] if extra_fields and "Parent Part Number" in extra_fields else "",
                "Device": extra_fields["Device"] if extra_fields and "Device" in extra_fields else "",
                "Role": info.get("Role", ""),
                "Subrole": subrole_header,
            }

        # 收集warning到session_state['json_warnings']
        if 'json_warnings' not in st.session_state:
            st.session_state['json_warnings'] = []
        for k in main_fields:
            v = str(json_main.get(k, "")).strip()
            if v == "" or v.lower() == "nan":
                if not (table_name in ["CFM", "Power"] and k == "Part Number"):
                    warn_msg = f"<b>Sheet:</b> <span style='color:#0072C6'>{sheet_name}</span> &nbsp; <b>Field:</b> <span style='color:#C80000'>{k}</span> &nbsp; <b>Status:</b> <span style='color:#C80000'>Missing</span>"
                    st.session_state['json_warnings'].append(warn_msg)

        # 生成最终json_dict，顺序与模板一致
        json_dict = {
            **json_main,
            "TableName": table_name,
            "Version": version_val,
            "Generated On": generated_on,
            "Rows": df.to_dict(orient="records")
        }
        return json_dict, None
    except Exception as e:
        return None, str(e)

def extract_device_child_parent(sheet_name, xls, uploaded_file, info):
    # Device: B2
    device_val = ""
    try:
        raw = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, dtype=str)
        device_val = str(raw.iloc[1,1]).strip()
    except Exception:
        device_val = ""
    parent_part_number = info.get("Item Number", "")
    child_part_number = ""
    if device_val and "Bom Report" in xls.sheet_names:
        bom_df = pd.read_excel(uploaded_file, sheet_name="Bom Report", header=0, dtype=str)
        bom_df.columns = [str(c).strip() for c in bom_df.columns]
        if "BOM Level" in bom_df.columns and "Part Number" in bom_df.columns and "Part Description" in bom_df.columns:
            mask = (bom_df["BOM Level"].astype(str).str.strip() == "1")
            filtered = bom_df[mask]
            for _, row in filtered.iterrows():
                desc = str(row["Part Description"]).strip()
                if device_val and device_val in desc:
                    child_part_number = str(row["Part Number"]).strip()
                    break
    return device_val, parent_part_number, child_part_number

def get_sheet_by_keyword(xls, keyword):
    # 返回第一个包含keyword（不区分大小写）的sheet名
    for name in xls.sheet_names:
        if keyword.lower() in name.lower():
            return name
    return None

def download_all_button(json_files, key=None):
    import zipfile
    buf = io.BytesIO()
    for_download_name = "All_JSON_Exports.zip"
    # 尝试从session_state中获取Item Number
    item_number = None
    part_info = st.session_state.get("part_properties_info")
    if part_info is not None:
        try:
            # part_info是DataFrame，列名应为'Name'和'Value'
            item_row = part_info[part_info['Name'] == 'Item Number']
            if not item_row.empty:
                val = item_row.iloc[0]['Value']
                if val and val != "N/A":
                    item_number = str(val).strip()
        except Exception:
            pass
    if item_number:
        for_download_name = f"{item_number}.zip"
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, content in json_files.items():
            zf.writestr(fname, json.dumps(content, ensure_ascii=False, indent=4))
    buf.seek(0)
    st.download_button(
        label="📦 Downloadd all JSON (ZIP)",
        data=buf,
        file_name=for_download_name,
        mime="application/zip",
        key=key or "download_all_zip_button"
    )

if uploaded_file:
    # 用文件名+文件大小做缓存key，保证同一文件不重复处理
    file_id = f"{uploaded_file.name}_{uploaded_file.size}"
    # 每次新文件上传时清空旧warning
    if st.session_state.get("last_file_id") != file_id:
        st.session_state['json_warnings'] = []
    if (
        "json_files" not in st.session_state or
        "preview_tabs" not in st.session_state or
        st.session_state.get("last_file_id") != file_id
    ):
        try:
            xls = pd.ExcelFile(uploaded_file)
            # 统计所有sheet匹配数
            keywords = [
                ("Storage", "Storage Mapping"),
                ("PCIe", "PCIe Slot Mapping"),
                ("Memory", "Memory Mapping"),
                ("Power", "Power"),
                ("CFM", "CFM")
            ]
            matched_sheets = []
            for sheet_name in xls.sheet_names:
                for kw, display_name in keywords:
                    if display_name == "Power":
                        # 只有表名等于Power（忽略大小写）才匹配
                        if sheet_name.strip().lower() == "power":
                            matched_sheets.append((display_name, sheet_name))
                    else:
                        if kw.lower() in sheet_name.lower():
                            matched_sheets.append((display_name, sheet_name))
            # 进度条总步数：Part Properties + BOM Report + 所有matched_sheets
            total_steps = 2 + len(matched_sheets)
            processed_steps = 0
            avg_time_per_sheet = 1.5  # 估算每步1.5秒
            est_total = int(total_steps * avg_time_per_sheet)
            progress_bar = st.progress(0, text="Preparing to process...")

            # Step 1: Part Properties
            if "Part Properties" not in xls.sheet_names:
                st.error("No 'Part Properties' sheet found in this Excel file.")
                st.session_state["json_files"] = {}
                st.session_state["preview_tabs"] = []
                st.session_state["last_file_id"] = file_id
            else:
                df = pd.read_excel(uploaded_file, sheet_name="Part Properties", header=None)
                info_fields = [
                    "Item Number",
                    "Part Class Path",
                    "Part Description",
                    "Part Revision",
                    "Business Group",
                    "Role",
                    "Subrole",
                    "Generation",
                    "Processor Type"
                ]
                info = get_info_fields(df, info_fields)
                # 构建左右两列的表格，左为字段，右为值
                df_display = pd.DataFrame({"Name": info_fields, "Value": [info.get(k, "N/A") for k in info_fields]})
                st.session_state["part_properties_info"] = df_display.copy()
                # 只赋值，不展示，展示逻辑统一放在信息区
                processed_steps += 1
                progress_bar.progress(processed_steps/total_steps, text=f"Step 1/Part Properties info displayed. Estimated {est_total-processed_steps*int(avg_time_per_sheet)}s left.")

                # Step 2: BOM Report
                if "Bom Report" in xls.sheet_names:
                    bom_fields = [
                        "BOM/Substitute BOM?",
                        "BOM Level",
                        "Part Number",
                        "Part Revision",
                        "Part Description",
                        "Part Classification",
                        "MSF IDs",
                        "Substitutes",
                        "BOM Quantity"
                    ]
                    try:
                        # 先不指定header，读取原始数据
                        raw_bom = pd.read_excel(uploaded_file, sheet_name="Bom Report", header=None, dtype=str)
                        header_row_idx = None
                        for i in range(len(raw_bom)):
                            first_cell = str(raw_bom.iloc[i,0]).strip()
                            if first_cell == "BOM/Substitute BOM?":
                                header_row_idx = i
                                break
                        if header_row_idx is not None:
                            bom_df = pd.read_excel(
                                uploaded_file,
                                sheet_name="Bom Report",
                                header=header_row_idx,
                                dtype=str
                            )
                            bom_df.columns = [str(c).strip() for c in bom_df.columns]
                            # 只保留BOM Level=1的行
                            if "BOM Level" in bom_df.columns:
                                filtered_bom = bom_df[bom_df["BOM Level"].astype(str).str.strip() == "1"]
                            else:
                                filtered_bom = bom_df.iloc[0:0]  # 空表
                            # 只保留需要的列
                            display_bom = filtered_bom[[c for c in bom_fields if c in filtered_bom.columns]].fillna("")
                            st.session_state["bom_report_level1"] = display_bom.copy()
                            # 只赋值，不展示，展示逻辑统一放在信息区
                        else:
                            st.warning("BOM Report header (BOM/Substitute BOM?) not found.")
                    except Exception as e:
                        st.warning(f"Failed to read BOM Report: {e}")
                else:
                    st.warning("No 'Bom Report' sheet found in this Excel file.")
                processed_steps += 1
                progress_bar.progress(processed_steps/total_steps, text=f"Step 2/BOM Report displayed. Estimated {est_total-processed_steps*int(avg_time_per_sheet)}s left.")

                # Step 3: 遍历所有sheet生成JSON
                generated_on = get_generated_on()
                json_files = {}
                preview_tabs = []
                start_time = datetime.datetime.now()
                for idx, (display_name, sheet) in enumerate(matched_sheets):
                    current_msg = f"Generating {display_name} ({sheet}) JSON..."
                    elapsed = (datetime.datetime.now() - start_time).total_seconds()
                    est_left = max(0, est_total - int(elapsed))
                    progress_bar.progress((processed_steps+idx)/total_steps, text=f"{current_msg} Estimated {est_left}s left.")
                    if display_name in ["Storage Mapping", "PCIe Slot Mapping", "Memory Mapping"]:
                        device, parent, child = extract_device_child_parent(sheet, xls, uploaded_file, info)
                        extra = {"Device": device, "Parent Part Number": parent, "Child Part Number": child}
                        table_name = display_name if display_name != "Memory Mapping" else "Memory Mappping"
                    else:
                        extra = {"Part Number": ""}
                        table_name = display_name
                    content, err = extract_table(
                        sheet, xls, uploaded_file, info, generated_on,
                        table_name=table_name,
                        filename=None,  # filename参数不再用于命名
                        extra_fields=extra
                    )
                    # 统一文件名格式：CFM/Power不带subrole，其他带subrole，且保证文件名唯一
                    if content:
                        doc_num = content.get("Document Number", "")
                        doc_rev = content.get("Document Revision", "")
                        table_name_val = content.get("TableName", "")
                        if table_name_val in ["CFM", "Power"]:
                            base_fname = f"{doc_num}_Rev{doc_rev}_{table_name_val}.json"
                        else:
                            subrole_val = content.get("Subrole", "")
                            base_fname = f"{doc_num}_Rev{doc_rev}_{table_name_val}_{subrole_val}.json"
                        # 保证文件名唯一，如有重名自动加sheet名或序号
                        fname = base_fname
                        if fname in json_files:
                            # 若sheet名已在文件名中则不再重复加
                            sheet_tag = str(sheet).replace(' ', '_').replace('/', '_')
                            alt_fname = f"{base_fname.rsplit('.json',1)[0]}_{sheet_tag}.json"
                            idx2 = 2
                            while alt_fname in json_files:
                                alt_fname = f"{base_fname.rsplit('.json',1)[0]}_{sheet_tag}{idx2}.json"
                                idx2 += 1
                            fname = alt_fname
                        json_files[fname] = content
                        preview_tabs.append((display_name, sheet, content, fname))
                    else:
                        st.warning(f"{sheet}: {err}")
                processed_steps += len(matched_sheets)
                # 进度条100%
                progress_bar.progress(1.0, text="All JSON files generated!")
                progress_bar.empty()
                st.session_state["json_files"] = json_files
                st.session_state["preview_tabs"] = preview_tabs
                st.session_state["last_file_id"] = file_id
        except Exception as e:
            st.error(f"Error reading Excel: {e}")
            st.session_state["json_files"] = {}
            st.session_state["preview_tabs"] = []
            st.session_state["last_file_id"] = file_id


    # 始终展示Part Properties Info和BOM Report（避免下载后消失）


    # ====== 信息区 ======

    st.markdown("""
    <div class='card-section'>
        <div class='step-title'><span class='step-num'>2</span>Part Properties & BOM Report</div>
    """, unsafe_allow_html=True)
    part_info = st.session_state.get("part_properties_info")
    if part_info is not None:
        with st.expander("Part Properties Info", expanded=True):
            st.table(part_info)
    bom_report = st.session_state.get("bom_report_level1")
    if bom_report is not None:
        with st.expander("BOM Report (BOM Level=1)", expanded=False):
            st.dataframe(bom_report, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)


    # 展示所有JSON预览和下载（无论是否刚刚生成还是缓存）
    json_files = st.session_state.get("json_files", {})
    preview_tabs = st.session_state.get("preview_tabs", [])
    json_warnings = st.session_state.get("json_warnings", [])

    if preview_tabs:
        st.markdown("""
        <div class='card-section'>
            <div class='step-title'><span class='step-num'>3</span>JSON Preview & Download</div>
        """, unsafe_allow_html=True)
        st.success("Processing Done. See below.")
        if json_warnings:
            # 美化 warning 展示为 HTML 列表，保留原有 HTML 样式
            warning_html = "<ul style='margin-left:1em;'>" + "".join([f"<li style='margin-bottom:4px'>{w}</li>" for w in json_warnings]) + "</ul>"
            st.warning("Warnings:", icon="⚠️")
            st.markdown(warning_html, unsafe_allow_html=True)
        category_map = {
            "Power": "Power",
            "CFM": "CFM",
            "Memory Mapping": "Memory Mapping",
            "PCIe Slot Mapping": "PCIe Mapping",
            "Storage Mapping": "Storage Mapping"
        }
        categories = ["Power", "CFM", "Memory Mapping", "PCIe Mapping", "Storage Mapping"]
        cat_tabs = st.tabs(categories)
        for idx, cat in enumerate(categories):
            with cat_tabs[idx]:
                found = False
                for i, (display_name, sheet, content, fname) in enumerate(preview_tabs):
                    if category_map.get(display_name) == cat:
                        found = True
                        st.markdown(f"**Sheet:** {sheet}")
                        st.json(content, expanded=False)
                        st.download_button(
                            label=f"Download {fname}",
                            data=json.dumps(content, ensure_ascii=False, indent=4),
                            file_name=fname,
                            mime="application/json",
                            key=f"download_{file_id}_{fname}_{display_name}_{sheet}_{i}"
                        )
                if not found:
                    st.info(f"No JSON file in this category.")
        st.markdown("\n")
        download_all_button(json_files, key=f"download_all_zip_button_{file_id}")
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        json_warnings = st.session_state.get("json_warnings", [])
        if json_warnings:
            warning_html = "<ul style='margin-left:1em;'>" + "".join([f"<li style='margin-bottom:4px'>{w}</li>" for w in json_warnings]) + "</ul>"
            st.warning("Warnings:", icon="⚠️")
            st.markdown(warning_html, unsafe_allow_html=True)
        st.info("No JSON export file generated. Please check the Excel content or sheet names.")