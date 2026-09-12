import streamlit as st
import pandas as pd
import numpy as np
import zipfile
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="蝙蝠偵測分析工具",
    page_icon="🦇",
    layout="wide",
)

# ── Custom CSS ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Space+Mono:wght@400;700&display=swap');

html, body, [class*="css"] { font-family: 'Noto Sans TC', sans-serif; }

.stApp { background: #0f1117; color: #e8eaf0; }

.hero {
    background: linear-gradient(135deg, #1a1f35 0%, #0d1b2a 50%, #1a1f35 100%);
    border: 1px solid #2a3550;
    border-radius: 16px;
    padding: 2.5rem 3rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: radial-gradient(circle at 70% 30%, rgba(99,179,237,0.06) 0%, transparent 60%);
    pointer-events: none;
}
.hero h1 {
    font-family: 'Space Mono', monospace;
    font-size: 2rem;
    color: #63b3ed;
    margin: 0 0 0.5rem 0;
    letter-spacing: -0.5px;
}
.hero p { color: #8896b3; font-size: 0.95rem; margin: 0; }

.section-header {
    font-family: 'Space Mono', monospace;
    font-size: 0.85rem;
    color: #63b3ed;
    text-transform: uppercase;
    letter-spacing: 2px;
    border-bottom: 1px solid #2a3550;
    padding-bottom: 0.6rem;
    margin: 2rem 0 1rem 0;
}

.file-card {
    background: #161b2e;
    border: 1px solid #2a3550;
    border-radius: 12px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}
.file-card-title {
    font-family: 'Space Mono', monospace;
    font-size: 0.9rem;
    color: #63b3ed;
    margin-bottom: 0.75rem;
}
.file-card-error {
    border-color: #c53030;
    background: #1a0d0d;
}
.file-card-error .file-card-title { color: #fc8181; }

.stDataFrame { border-radius: 10px; overflow: hidden; }

.stDownloadButton > button {
    background: linear-gradient(135deg, #2b6cb0, #2c5282) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 2rem !important;
    font-family: 'Space Mono', monospace !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.5px !important;
    transition: opacity 0.2s !important;
}
.stDownloadButton > button:hover { opacity: 0.85 !important; }

section[data-testid="stSidebar"] {
    background: #0d1117 !important;
    border-right: 1px solid #1e2433;
}
</style>
""", unsafe_allow_html=True)


# ── Core analysis functions ────────────────────────────────────────────────────

def compute_concurrent(df_all, source_zx, source_cx, threshold=2.0, add_pair_id=False):
    times_zx = source_zx['video_time_sec'].values
    times_cx = source_cx['video_time_sec'].values

    pairs = []
    for i, ta in enumerate(times_zx):
        for j, tb in enumerate(times_cx):
            if abs(float(ta) - float(tb)) < threshold:
                pairs.append((i, j))

    c_zx = {p[0] for p in pairs}
    c_cx = {p[1] for p in pairs}

    df_all = df_all.copy()
    df_all['同時出現'] = ''
    zx_idx = df_all[df_all['角度'] == '正下'].index.tolist()
    cx_idx = df_all[df_all['角度'] == '側向'].index.tolist()

    for pos, idx in enumerate(zx_idx):
        if pos in c_zx:
            df_all.at[idx, '同時出現'] = 'true'
    for pos, idx in enumerate(cx_idx):
        if pos in c_cx:
            df_all.at[idx, '同時出現'] = 'true'

    if add_pair_id:
        df_all['配對組'] = pd.NA
        zx_to_group, cx_to_group = {}, {}
        for group_id, (zx_pos, cx_pos) in enumerate(pairs):
            zx_to_group.setdefault(zx_pos, group_id)
            cx_to_group.setdefault(cx_pos, group_id)
        for pos, idx in enumerate(zx_idx):
            if pos in zx_to_group:
                df_all.at[idx, '配對組'] = zx_to_group[pos]
        for pos, idx in enumerate(cx_idx):
            if pos in cx_to_group:
                df_all.at[idx, '配對組'] = cx_to_group[pos]

    return df_all, pairs


def run_analysis(uploaded_file, sheet_zx, sheet_cx, concurrent_threshold):
    """Returns (df1, df2, pairs2, pivot, bat_conc, bird_conc, conc_species_total, err)."""
    xl = pd.read_excel(uploaded_file, sheet_name=None)

    if sheet_zx not in xl or sheet_cx not in xl:
        return None, None, None, None, None, None, None, \
            f"找不到工作表：請確認工作表名稱（{sheet_zx} / {sheet_cx}）"

    df_zx = xl[sheet_zx].copy()
    df_cx = xl[sheet_cx].copy()

    required_cols = {'video_time_sec', '角度'}
    for label, df in [(sheet_zx, df_zx), (sheet_cx, df_cx)]:
        missing = required_cols - set(df.columns)
        if missing:
            return None, None, None, None, None, None, None, \
                f"工作表「{label}」缺少必要欄位：{missing}"

    for col in ['物種', 'bat_count', 'CP', '高度']:
        for df in [df_zx, df_cx]:
            if col not in df.columns:
                df[col] = np.nan

    df_zx['bat_count'] = pd.to_numeric(df_zx['bat_count'], errors='coerce').fillna(1).astype(int)
    df_cx['bat_count'] = pd.to_numeric(df_cx['bat_count'], errors='coerce').fillna(1).astype(int)
    df_zx['CP'] = pd.to_numeric(df_zx['CP'], errors='coerce')
    df_cx['CP'] = pd.to_numeric(df_cx['CP'], errors='coerce')

    # Sheet 1
    df1 = (pd.concat([df_zx, df_cx], ignore_index=True)
             .sort_values('video_time_sec').reset_index(drop=True))
    df1, _ = compute_concurrent(
        df1, df_zx.reset_index(drop=True), df_cx.reset_index(drop=True), concurrent_threshold)

    # Sheet 2
    df2_zx = df_zx[df_zx['CP'] == 0].reset_index(drop=True)
    df2_cx = df_cx[df_cx['CP'] == 0].reset_index(drop=True)
    df2 = (pd.concat([df2_zx, df2_cx], ignore_index=True)
             .sort_values('video_time_sec').reset_index(drop=True))
    df2, pairs2 = compute_concurrent(df2, df2_zx, df2_cx, concurrent_threshold, add_pair_id=True)

    if '同時出現' in df2.columns:
        pos = df2.columns.tolist().index('同時出現')
        df2.insert(pos + 1, '人工驗證', '')

    # Pivot
    if len(df2) > 0 and df2['物種'].notna().any() and df2['高度'].notna().any():
        pivot = (df2.dropna(subset=['物種', '高度'])
                    .groupby(['角度', '物種', '高度'])['bat_count'].sum()
                    .reset_index()
                    .rename(columns={'bat_count': '數量(bat_count加總)'})
                    .sort_values(['角度', '物種', '高度']))
    else:
        pivot = pd.DataFrame(columns=['角度', '物種', '高度', '數量(bat_count加總)'])

    # Concurrent pair stats
    conc_species_total = {}
    bat_conc = bird_conc = 0

    if len(df2_zx) > 0 and len(df2_cx) > 0:
        candidates = []
        for i, rz in df2_zx.iterrows():
            for j, rc in df2_cx.iterrows():
                if abs(float(rz['video_time_sec']) - float(rc['video_time_sec'])) < concurrent_threshold:
                    candidates.append((
                        max(int(rz['bat_count']), int(rc['bat_count'])),
                        i, j, rz['物種'], rc['物種'],
                    ))
        candidates.sort(reverse=True)

        used_zx, used_cx = set(), set()
        for pair_max, i, j, sp_zx, sp_cx in candidates:
            if i not in used_zx and j not in used_cx:
                used_zx.add(i); used_cx.add(j)
                if pd.isna(sp_zx) and pd.isna(sp_cx):
                    continue
                if sp_zx == sp_cx:
                    k = sp_zx if pd.notna(sp_zx) else '未知'
                    conc_species_total[k] = conc_species_total.get(k, 0) + pair_max
                else:
                    for sp in [sp_zx, sp_cx]:
                        if pd.notna(sp):
                            conc_species_total[sp] = conc_species_total.get(sp, 0) + pair_max

        bat_conc  = int(conc_species_total.get('蝙蝠', 0))
        bird_conc = int(conc_species_total.get('鳥',   0))

    return df1, df2, pairs2, pivot, bat_conc, bird_conc, conc_species_total, None


def build_excel(df1, df2, pairs2, pivot, bat_conc, bird_conc, conc_species_total):
    from openpyxl.worksheet.datavalidation import DataValidation

    wb   = Workbook()
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),  bottom=Side(style='thin'))
    H_FONT  = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    H_FILL  = PatternFill('solid', start_color='2E75B6')
    H_FILL2 = PatternFill('solid', start_color='70AD47')
    D_FONT  = Font(name='Arial', size=10)
    PAIR_FILLS = [PatternFill('solid', start_color='FFFFFF'),
                  PatternFill('solid', start_color='D9F0D3')]

    def auto_width(ws):
        for col in ws.columns:
            mx = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 42)

    def write_df(ws, df):
        for ci, col in enumerate(df.columns, 1):
            c = ws.cell(1, ci, col)
            c.font = H_FONT; c.fill = H_FILL
            c.alignment = Alignment(horizontal='center'); c.border = thin
        for ri, row in enumerate(df.itertuples(index=False), 2):
            for ci, val in enumerate(row, 1):
                c = ws.cell(ri, ci, val if pd.notna(val) else '')
                c.font = D_FONT; c.border = thin
        auto_width(ws)

    # Sheet 1
    ws1 = wb.active
    ws1.title = '全部資料合併'
    write_df(ws1, df1)

    # Sheet 2
    ws2 = wb.create_sheet('CP=0資料合併')
    group_col   = '配對組'
    output_cols = [c for c in df2.columns if c != group_col]

    for ci, col in enumerate(output_cols, 1):
        c = ws2.cell(1, ci, col)
        c.font = H_FONT; c.fill = H_FILL
        c.alignment = Alignment(horizontal='center'); c.border = thin

    if len(df2) == 0:
        ws2.cell(2, 1, '（無 CP=0 資料）').font = D_FONT
    else:
        group_color_map = {}
        color_counter   = 0
        for ri, row_data in enumerate(df2.itertuples(index=False), 2):
            row_dict = dict(zip(df2.columns, row_data))
            gid      = row_dict.get(group_col, pd.NA)
            if pd.notna(gid):
                gid_int = int(gid)
                if gid_int not in group_color_map:
                    group_color_map[gid_int] = color_counter % 2
                    color_counter += 1
                fill = PAIR_FILLS[group_color_map[gid_int]]
            else:
                fill = None
            for ci, col in enumerate(output_cols, 1):
                val = row_dict[col]
                c   = ws2.cell(ri, ci, val if pd.notna(val) else '')
                c.font = D_FONT; c.border = thin
                if fill:
                    c.fill = fill

        if '人工驗證' in output_cols:
            manval_col_idx = output_cols.index('人工驗證') + 1
            last_row  = len(df2) + 1
            sqref_str = (f"{get_column_letter(manval_col_idx)}2:"
                         f"{get_column_letter(manval_col_idx)}{last_row}")
            dv = DataValidation(type="list", formula1='"true,false"',
                                allow_blank=True, showDropDown=False,
                                showErrorMessage=True,
                                errorTitle='輸入錯誤', error='請選擇 true 或 false')
            dv.add(sqref_str)
            ws2.add_data_validation(dv)

    auto_width(ws2)

    # Sheet 3
    ws3 = wb.create_sheet('總成果表')
    ws3.merge_cells('A1:D1')
    ws3['A1'] = '總成果表（CP=0）'
    ws3['A1'].font      = Font(bold=True, name='Arial', size=14, color='1F3864')
    ws3['A1'].alignment = Alignment(horizontal='center')

    row = 3
    ws3.cell(row, 1, '各角度 / 物種 / 高度 數量統計（bat_count 加總）').font = Font(bold=True, name='Arial', size=11)
    row += 1
    for ci, h in enumerate(['角度', '物種', '高度', '數量(bat_count加總)'], 1):
        c = ws3.cell(row, ci, h)
        c.font = H_FONT; c.fill = H_FILL
        c.alignment = Alignment(horizontal='center'); c.border = thin
    row += 1

    if len(pivot) == 0:
        ws3.cell(row, 1, '（無 CP=0 且含物種/高度資料）').font = D_FONT
        row += 1; total = 0
    else:
        for _, r in pivot.iterrows():
            for ci, val in enumerate([r['角度'], r['物種'], int(r['高度']), int(r['數量(bat_count加總)'])], 1):
                c = ws3.cell(row, ci, val)
                c.font = D_FONT; c.border = thin; c.alignment = Alignment(horizontal='center')
            row += 1
        total = int(pivot['數量(bat_count加總)'].sum())

    for ci, val in enumerate(['合計', '', '', total], 1):
        c = ws3.cell(row, ci, val)
        c.font = Font(bold=True, name='Arial'); c.border = thin
        c.fill = PatternFill('solid', start_color='BDD7EE')
        c.alignment = Alignment(horizontal='center')
    row += 2

    ws3.cell(row, 1, '同時出現（正下 ↔ 側向 video_time_sec 相差 < 2 秒）各配對取 bat_count 最大值後加總').font = Font(bold=True, name='Arial', size=11)
    row += 1
    for ci, h in enumerate(['物種', '數量（配對最大值加總）'], 1):
        c = ws3.cell(row, ci, h)
        c.font = H_FONT; c.fill = H_FILL2
        c.alignment = Alignment(horizontal='center'); c.border = thin
    row += 1

    grand_total = 0
    if conc_species_total:
        for sp, cnt in conc_species_total.items():
            for ci, val in enumerate([sp, cnt], 1):
                c = ws3.cell(row, ci, val)
                c.font = D_FONT; c.border = thin; c.alignment = Alignment(horizontal='center')
            grand_total += cnt; row += 1
    else:
        ws3.cell(row, 1, '（無同時出現配對）').font = D_FONT; row += 1

    for ci, val in enumerate(['合計', grand_total], 1):
        c = ws3.cell(row, ci, val)
        c.font = Font(bold=True, name='Arial'); c.border = thin
        c.fill = PatternFill('solid', start_color='C6EFCE')
        c.alignment = Alignment(horizontal='center')

    auto_width(ws3)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def build_zip(results: list) -> bytes:
    """Pack multiple xlsx results into a single zip file.
    results: list of (original_filename, xlsx_bytes)
    """
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for orig_name, xlsx_bytes in results:
            stem = orig_name.replace('.xlsx', '')
            zf.writestr(f"{stem}_分析結果.xlsx", xlsx_bytes)
    return buf.getvalue()


# ── UI ─────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="hero">
    <h1>🦇 蝙蝠偵測分析工具</h1>
    <p>上傳一或多個熱像儀偵測結果 Excel，自動合併正下／側向資料、標記同時出現事件、輸出三頁分析報表。</p>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ 分析設定")
    st.markdown("---")
    sheet_zx  = st.text_input("正下 工作表名稱", value="正下")
    sheet_cx  = st.text_input("側向 工作表名稱", value="側向")
    threshold = st.slider(
        "同時出現閾值（秒）",
        min_value=0.5, max_value=10.0, value=2.0, step=0.5,
        help="正下與側向 video_time_sec 相差小於此值，即標記為同時出現",
    )
    st.markdown("---")
    st.markdown("**必要欄位**")
    st.markdown("- `角度`\n- `video_time_sec`")
    st.markdown("**選填欄位**（空白視為無偵測）")
    st.markdown("- `物種`\n- `bat_count`\n- `CP`\n- `高度`")

st.markdown('<div class="section-header">上傳檔案</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "選擇一或多個 Excel 檔案（.xlsx）",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploaded_files:
    with st.spinner(f"分析中（共 {len(uploaded_files)} 個檔案）..."):
        all_results = []   # (filename, xlsx_bytes) for zip
        all_summaries = [] # (filename, metrics_dict, err)

        for uf in uploaded_files:
            df1, df2, pairs2, pivot, bat_conc, bird_conc, conc_species_total, err = run_analysis(
                uf, sheet_zx, sheet_cx, threshold
            )
            if err:
                all_summaries.append((uf.name, None, err))
            else:
                xlsx_bytes = build_excel(df1, df2, pairs2, pivot, bat_conc, bird_conc, conc_species_total)
                all_results.append((uf.name, xlsx_bytes))
                all_summaries.append((uf.name, {
                    'df1': df1, 'df2': df2, 'pivot': pivot,
                    'bat_conc': bat_conc, 'bird_conc': bird_conc,
                    'conc_species_total': conc_species_total,
                }, None))

    # ── Download buttons ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">下載結果</div>', unsafe_allow_html=True)

    if len(all_results) == 1:
        # Single file → direct download
        fname, xlsx_bytes = all_results[0]
        st.download_button(
            label=f"⬇️  下載分析結果：{fname.replace('.xlsx','')}",
            data=xlsx_bytes,
            file_name=f"{fname.replace('.xlsx','')}_分析結果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif len(all_results) > 1:
        # Multiple files → zip download + individual downloads
        zip_bytes = build_zip(all_results)
        st.download_button(
            label=f"⬇️  一次下載全部（{len(all_results)} 個檔案，ZIP）",
            data=zip_bytes,
            file_name="蝙蝠分析結果_全部.zip",
            mime="application/zip",
        )
        st.markdown("或個別下載：")
        cols = st.columns(min(len(all_results), 3))
        for idx, (fname, xlsx_bytes) in enumerate(all_results):
            with cols[idx % 3]:
                st.download_button(
                    label=f"⬇️ {fname.replace('.xlsx','')}",
                    data=xlsx_bytes,
                    file_name=f"{fname.replace('.xlsx','')}_分析結果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{idx}",
                )

    # ── Per-file results ──────────────────────────────────────────────────────
    st.markdown('<div class="section-header">各檔案分析結果</div>', unsafe_allow_html=True)

    for fname, metrics, err in all_summaries:
        is_error = err is not None
        card_cls = "file-card file-card-error" if is_error else "file-card"
        st.markdown(f'<div class="{card_cls}"><div class="file-card-title">📄 {fname}</div></div>',
                    unsafe_allow_html=True)

        if is_error:
            st.error(f"❌ {err}")
            continue

        df1  = metrics['df1']
        df2  = metrics['df2']
        pivot = metrics['pivot']
        bat_conc  = metrics['bat_conc']
        bird_conc = metrics['bird_conc']
        conc_species_total = metrics['conc_species_total']

        # Metrics row
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: st.metric("全部資料（筆）", f"{len(df1):,}")
        with c2: st.metric("CP=0 資料（筆）", f"{len(df2):,}")
        with c3:
            simul = int((df2['同時出現'] == 'true').sum()) if len(df2) > 0 else 0
            st.metric("同時出現（筆）", simul)
        with c4: st.metric("同時出現 蝙蝠", f"{bat_conc} 隻")
        with c5: st.metric("同時出現 鳥",   f"{bird_conc} 隻")

        # Tabs per file
        tab1, tab2, tab3 = st.tabs([
            "📋 工作表1：全部資料",
            "📋 工作表2：CP=0",
            "📊 工作表3：總成果表",
        ])

        with tab1:
            simul1 = int((df1['同時出現'] == 'true').sum())
            st.markdown(f"共 **{len(df1):,}** 筆，同時出現標記 **{simul1}** 筆")
            st.dataframe(df1, use_container_width=True, height=350)

        with tab2:
            if len(df2) == 0:
                st.info("此檔案中沒有 CP=0 的資料。")
            else:
                bat_sum = int(df2['bat_count'].sum())
                simul2  = int((df2['同時出現'] == 'true').sum())
                st.markdown(f"共 **{len(df2):,}** 筆，bat_count 總和 **{bat_sum}**，同時出現標記 **{simul2}** 筆")
                display_df2 = df2.drop(columns=['配對組'], errors='ignore')
                st.dataframe(display_df2, use_container_width=True, height=350)

        with tab3:
            col_a, col_b = st.columns([2, 1])
            with col_a:
                st.markdown("#### 各角度／物種／高度數量（bat_count加總）")
                if len(pivot) == 0:
                    st.info("無 CP=0 且含物種／高度的資料可統計。")
                else:
                    st.dataframe(pivot, use_container_width=True)
                    st.markdown(f"**合計：{int(pivot['數量(bat_count加總)'].sum())} 隻**")
            with col_b:
                st.markdown("#### 同時出現（配對最大值加總）")
                if conc_species_total:
                    conc_df = pd.DataFrame({
                        '物種': list(conc_species_total.keys()) + ['合計'],
                        '數量': list(conc_species_total.values()) + [sum(conc_species_total.values())],
                    })
                    st.dataframe(conc_df, use_container_width=True, hide_index=True)
                else:
                    st.info("無同時出現配對。")

        st.markdown("---")

else:
    st.info("👆 請在上方上傳一或多個 Excel 檔案以開始分析")
    st.markdown("""
    **使用流程：**
    1. 在左側設定工作表名稱（預設：正下 / 側向）與同時出現閾值
    2. 上傳一或多個 .xlsx 檔案
    3. 預覽各檔案的分析結果
    4. 下載個別或全部打包（ZIP）的 Excel 報表
    """)
