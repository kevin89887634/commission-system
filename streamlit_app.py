"""
💰 佣金管理系统 v3.0
简化版：单页面完成所有操作，实时汇总，颜色区分
"""
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ==================== 配置 ====================
st.set_page_config(page_title="佣金管理系统", page_icon="💰", layout="wide")

# 工具函数
def normalize_policy(policy_num):
    if policy_num is None:
        return ""
    s = str(policy_num).strip()
    s = re.sub(r'^(LS|NL|L)', '', s, flags=re.IGNORECASE)
    if s.endswith('00') and len(s) > 2:
        s = s[:-2]
    return s

def safe_float(value, default=0.0):
    try:
        if value is None or pd.isna(value):
            return default
        return float(value)
    except:
        return default

def is_valid_policy(policy):
    if policy is None or pd.isna(policy):
        return False
    s = str(policy).strip()
    if not s or s.lower() in ['nan', 'none', 'policy', 'policy #']:
        return False
    if not any(c.isdigit() for c in s):
        return False
    return True

def parse_nlg_file(uploaded_file):
    for header_row in [5, 4, 6, 3, 1, 0]:
        try:
            df = pd.read_excel(uploaded_file, header=header_row, engine='openpyxl')
            uploaded_file.seek(0)
            cols_lower = [str(c).lower() for c in df.columns]
            if any('policy' in c for c in cols_lower) and len(df) > 0:
                policy_col = next((c for c in df.columns if 'policy' in str(c).lower()), None)
                first_val = str(df[policy_col].iloc[0]) if len(df) > 0 else ''
                if is_valid_policy(first_val):
                    return df, None
        except:
            uploaded_file.seek(0)
            continue
    return None, "Unable to parse file"

# 解析zhubiao文件
def parse_zhubiao(uploaded_file):
    """解析zhubiao文件，返回 (policy, insured) 到分佣信息的映射"""
    try:
        df = pd.read_excel(uploaded_file, header=0, engine='openpyxl')
        uploaded_file.seek(0)

        # zhubiao列结构: A=Policy, B=Insured, H=Recruiter, I=Rate, J=Split, L=CFT, M=CFT Rate, N=CFT Split
        result_by_policy = {}      # 按保单号匹配
        result_by_insured = {}     # 按客户名匹配（备用）

        for idx, row in df.iterrows():
            # 获取保单号 (第一列A)
            policy_raw = row.iloc[0] if len(row) > 0 else None
            if not is_valid_policy(policy_raw):
                continue
            policy = normalize_policy(policy_raw)

            # 获取客户名 (第二列B)
            insured = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''

            # 获取分佣信息 (H=7, I=8, J=9, L=11, M=12, N=13)
            recruiter = str(row.iloc[7]).strip() if len(row) > 7 and pd.notna(row.iloc[7]) else ''
            rate1 = safe_float(row.iloc[8], 55) if len(row) > 8 else 55
            split1 = safe_float(row.iloc[9], 100) if len(row) > 9 else 100
            cft = str(row.iloc[11]).strip() if len(row) > 11 and pd.notna(row.iloc[11]) else ''
            rate2 = safe_float(row.iloc[12], 55) if len(row) > 12 else 55
            split2 = safe_float(row.iloc[13], 0) if len(row) > 13 else 0

            # 处理百分比格式 (如果是小数则转为百分比)
            if rate1 < 1: rate1 = rate1 * 100
            if split1 < 1 and split1 > 0: split1 = split1 * 100
            if rate2 < 1: rate2 = rate2 * 100
            if split2 < 1 and split2 > 0: split2 = split2 * 100

            info = {
                'Recruiter': recruiter,
                'Rate1': int(rate1),
                'Split1': int(split1),
                'CFT': cft if cft and cft != '-' and cft.lower() != 'nan' else '',
                'Rate2': int(rate2),
                'Split2': int(split2),
                'Insured': insured
            }

            # 保存到两个映射
            result_by_policy[policy] = info
            if insured:
                # 标准化客户名（去空格、转小写）用于匹配
                insured_key = insured.lower().replace(' ', '')
                result_by_insured[insured_key] = info

        return {'by_policy': result_by_policy, 'by_insured': result_by_insured}, None
    except Exception as e:
        return None, str(e)

# Session State
if 'data' not in st.session_state:
    st.session_state.data = None
if 'zhubiao_map' not in st.session_state:
    st.session_state.zhubiao_map = None

# ==================== 页面标题 ====================
st.title("💰 佣金管理系统 v3.0")

# ==================== 上传区域 ====================
with st.expander("📤 Upload NLG Report", expanded=st.session_state.data is None):
    uploaded_file = st.file_uploader("Select Excel file", type=['xlsx', 'xls'])

    if uploaded_file and st.button("📥 Import", type="primary"):
        df, error = parse_nlg_file(uploaded_file)
        if error:
            st.error(f"❌ {error}")
        else:
            # 标准化列名
            col_map = {}
            for col in df.columns:
                col_lower = str(col).lower().strip()
                if 'policy' in col_lower:
                    col_map[col] = 'Policy'
                elif 'insured' in col_lower or 'annuitant' in col_lower:
                    col_map[col] = 'Insured'
                elif col_lower == 'agent':
                    col_map[col] = 'Agent'
                elif 'recruiter' in col_lower:
                    col_map[col] = 'Recruiter'
                elif 'modal' in col_lower:
                    col_map[col] = 'Modal'
                elif 'aap' in col_lower:
                    col_map[col] = 'AAP'
                elif 'product' in col_lower:
                    col_map[col] = 'Product'
            df = df.rename(columns=col_map)

            # 处理数据
            df = df[df['Policy'].apply(is_valid_policy)]
            df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
            df['Modal'] = df['Modal'].apply(safe_float) if 'Modal' in df.columns else 0
            df['AAP'] = df['AAP'].apply(safe_float) if 'AAP' in df.columns else 0
            df = df[(df['AAP'] > 0) | (df['Modal'] > 0)].reset_index(drop=True)

            if len(df) == 0:
                st.error("❌ No valid data")
            else:
                # 构建工作数据
                rows = []
                for _, row in df.iterrows():
                    modal = safe_float(row.get('Modal', 0))
                    aap = safe_float(row.get('AAP', 0))
                    premium = modal if (modal > 0 and aap > 0 and aap/modal > 6) else (aap if aap > 0 else modal)
                    product = str(row.get('Product', '')).lower()
                    comm_rate = 0.67 if 'term' in product else 0.80
                    agent = str(row.get('Agent', '')) if pd.notna(row.get('Agent')) else ''
                    recruiter = str(row.get('Recruiter', '')) if pd.notna(row.get('Recruiter')) else ''

                    rows.append({
                        'Policy': row['Policy_Norm'],
                        'Insured': str(row.get('Insured', ''))[:30] if pd.notna(row.get('Insured')) else '',
                        'Premium': premium,
                        'CommRate': int(comm_rate * 100),  # 百分比格式 (80 or 67)
                        'Agent': agent,
                        'Person1': '',     # 待从zhubiao匹配
                        'Rate1': 55,       # 默认值
                        'Split1': 100,     # 默认值
                        'Person2': '',
                        'Rate2': 55,
                        'Split2': 0,
                        'MatchStatus': '❓ 未匹配'  # 匹配状态
                    })

                st.session_state.data = pd.DataFrame(rows)
                st.success(f"✅ Import successful! {len(rows)} records")
                st.rerun()

# ==================== zhubiao匹配区域 ====================
if st.session_state.data is not None:
    with st.expander("📎 Upload zhubiao (Auto Match by Policy + Insured)", expanded=True):
        zhubiao_file = st.file_uploader("Select zhubiao Excel file", type=['xlsx', 'xls'], key="zhubiao")

        if zhubiao_file and st.button("🔄 Match from zhubiao", type="primary"):
            zhubiao_maps, error = parse_zhubiao(zhubiao_file)
            if error:
                st.error(f"❌ Parse error: {error}")
            else:
                st.session_state.zhubiao_map = zhubiao_maps
                by_policy = zhubiao_maps['by_policy']
                by_insured = zhubiao_maps['by_insured']

                # 匹配并更新数据
                matched_by_policy = 0
                matched_by_insured = 0
                unmatched = 0

                for idx, row in st.session_state.data.iterrows():
                    policy = row['Policy']
                    insured = row['Insured']
                    insured_key = insured.lower().replace(' ', '') if insured else ''

                    info = None
                    match_type = ''

                    # 优先按保单号匹配
                    if policy in by_policy:
                        info = by_policy[policy]
                        match_type = '✅ Policy匹配'
                        matched_by_policy += 1
                    # 其次按客户名匹配
                    elif insured_key and insured_key in by_insured:
                        info = by_insured[insured_key]
                        match_type = '🔄 Insured匹配'
                        matched_by_insured += 1
                    else:
                        unmatched += 1
                        st.session_state.data.loc[idx, 'MatchStatus'] = '❌ 未找到匹配'
                        continue

                    # 更新数据
                    if info:
                        if info['Recruiter']:
                            st.session_state.data.loc[idx, 'Person1'] = info['Recruiter']
                        st.session_state.data.loc[idx, 'Rate1'] = info['Rate1']
                        st.session_state.data.loc[idx, 'Split1'] = info['Split1']
                        st.session_state.data.loc[idx, 'Person2'] = info['CFT']
                        st.session_state.data.loc[idx, 'Rate2'] = info['Rate2']
                        st.session_state.data.loc[idx, 'Split2'] = info['Split2']
                        st.session_state.data.loc[idx, 'MatchStatus'] = match_type

                # 显示匹配结果
                st.success(f"✅ 匹配完成!")
                st.write(f"- 按保单号匹配: {matched_by_policy} 条")
                st.write(f"- 按客户名匹配: {matched_by_insured} 条")
                if unmatched > 0:
                    st.warning(f"- ❌ 未匹配: {unmatched} 条 (需手动填写)")
                st.rerun()

# ==================== 主数据表 ====================
if st.session_state.data is not None:
    df = st.session_state.data.copy()

    # 计算每行的佣金 (Rate和Split都是百分比，需要除以100)
    df['Comm1'] = df['Premium'] * (df['Rate1']/100) * (df['Split1']/100)
    df['Comm2'] = df['Premium'] * (df['Rate2']/100) * (df['Split2']/100)
    df['TotalSplit'] = df['Split1'] + df['Split2']

    # ==================== 汇总区域 ====================
    st.markdown("---")
    col1, col2, col3 = st.columns(3)

    # 按人汇总
    summary_data = []
    all_persons = set(df['Person1'].dropna().unique()) | set(df['Person2'].dropna().unique())
    all_persons = [p for p in all_persons if p and str(p).strip()]

    for person in sorted(all_persons):
        comm1 = df[df['Person1'] == person]['Comm1'].sum()
        comm2 = df[df['Person2'] == person]['Comm2'].sum()
        count1 = len(df[df['Person1'] == person])
        count2 = len(df[df['Person2'] == person])
        summary_data.append({
            'Person': person,
            'Commission': comm1 + comm2,
            'Count': count1 + count2
        })

    if summary_data:
        summary_df = pd.DataFrame(summary_data).sort_values('Commission', ascending=False)
        total_comm = summary_df['Commission'].sum()

        with col1:
            st.metric("📊 Total Commission", f"${total_comm:,.2f}")
        with col2:
            st.metric("👥 Recruiters", len(summary_data))
        with col3:
            st.metric("📋 Records", len(df))

        # 汇总表
        st.markdown("### 📈 Commission by Recruiter")
        for _, row in summary_df.iterrows():
            pct = row['Commission'] / total_comm * 100 if total_comm > 0 else 0
            st.write(f"**{row['Person']}**: ${row['Commission']:,.2f} ({pct:.1f}%) - {row['Count']} policies")

    st.markdown("---")

    # ==================== 批量设置 ====================
    st.markdown("### 🔧 Batch Settings")
    bcol1, bcol2, bcol3, bcol4, bcol5, bcol6, bcol7 = st.columns([2, 1, 1, 2, 1, 1, 2])

    with bcol1:
        batch_p1 = st.text_input("Recruiter", key="bp1")
    with bcol2:
        batch_r1 = st.number_input("佣金比例%", 0, 100, 55, 1, key="br1")
    with bcol3:
        batch_s1 = st.number_input("分佣比例%", 0, 100, 100, 10, key="bs1")
    with bcol4:
        batch_p2 = st.text_input("CFT", key="bp2")
    with bcol5:
        batch_r2 = st.number_input("CFT比例%", 0, 100, 55, 1, key="br2")
    with bcol6:
        batch_s2 = st.number_input("CFT分佣%", 0, 100, 0, 10, key="bs2")

    with bcol7:
        total_split = batch_s1 + batch_s2
        if total_split == 100:
            st.success(f"✓ 分佣={total_split}%")
            can_apply = True
        else:
            st.error(f"✗ 分佣={total_split}%≠100%")
            can_apply = False

        if st.button("📝 Apply", disabled=not can_apply, type="primary"):
            mask = st.session_state.data['_selected'] == True
            if mask.sum() > 0:
                if batch_p1:
                    st.session_state.data.loc[mask, 'Person1'] = batch_p1
                st.session_state.data.loc[mask, 'Rate1'] = batch_r1
                st.session_state.data.loc[mask, 'Split1'] = batch_s1
                st.session_state.data.loc[mask, 'Person2'] = batch_p2
                st.session_state.data.loc[mask, 'Rate2'] = batch_r2
                st.session_state.data.loc[mask, 'Split2'] = batch_s2
                st.session_state.data['_selected'] = False
                st.rerun()
            else:
                st.warning("Please select rows first")

    # 快捷按钮
    qcol1, qcol2, qcol3, qcol4 = st.columns(4)
    with qcol1:
        if st.button("☑️ Select All"):
            st.session_state.data['_selected'] = True
            st.rerun()
    with qcol2:
        if st.button("⬜ Deselect All"):
            st.session_state.data['_selected'] = False
            st.rerun()
    with qcol3:
        if '_selected' in st.session_state.data.columns:
            st.info(f"Selected: {st.session_state.data['_selected'].sum()}")

    st.markdown("---")

    # ==================== 数据表 ====================
    st.markdown("### 📋 Detail Data")

    # 添加选择列
    if '_selected' not in st.session_state.data.columns:
        st.session_state.data['_selected'] = False

    # 显示匹配状态统计
    if 'MatchStatus' in st.session_state.data.columns:
        status_counts = st.session_state.data['MatchStatus'].value_counts()
        scol1, scol2, scol3 = st.columns(3)
        with scol1:
            matched = len(st.session_state.data[st.session_state.data['MatchStatus'].str.contains('✅|🔄', na=False)])
            st.metric("✅ 已匹配", matched)
        with scol2:
            unmatched = len(st.session_state.data[st.session_state.data['MatchStatus'].str.contains('❌|❓', na=False)])
            st.metric("❌ 未匹配", unmatched)
        with scol3:
            st.metric("📋 总计", len(st.session_state.data))

    # 显示表格
    display_cols = ['_selected', 'MatchStatus', 'Policy', 'Insured', 'CommRate', 'Premium', 'Person1', 'Rate1', 'Split1', 'Person2', 'Rate2', 'Split2']
    # 确保所有列都存在
    for col in display_cols:
        if col not in st.session_state.data.columns:
            if col == 'MatchStatus':
                st.session_state.data[col] = '❓ 未匹配'
            else:
                st.session_state.data[col] = ''

    display_df = st.session_state.data[display_cols].copy()

    # 计算佣金用于显示 (百分比需要除以100)
    display_df['Comm1'] = st.session_state.data['Premium'] * (st.session_state.data['Rate1']/100) * (st.session_state.data['Split1']/100)
    display_df['Comm2'] = st.session_state.data['Premium'] * (st.session_state.data['Rate2']/100) * (st.session_state.data['Split2']/100)

    edited = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            '_selected': st.column_config.CheckboxColumn('✓', default=False, width='small'),
            'MatchStatus': st.column_config.TextColumn('Status', disabled=True, width='small'),
            'Policy': st.column_config.TextColumn('Policy', disabled=True, width='small'),
            'Insured': st.column_config.TextColumn('Insured', disabled=True, width='medium'),
            'CommRate': st.column_config.NumberColumn('Comm Rate %', disabled=True, format='%.0f%%', width='small'),
            'Premium': st.column_config.NumberColumn('Gross Comm Earned', disabled=True, format='$%.2f', width='small'),
            'Person1': st.column_config.TextColumn('Recruiter', width='medium'),
            'Rate1': st.column_config.NumberColumn('Recruiter佣金比例', format='%.0f%%', width='small'),
            'Split1': st.column_config.NumberColumn('Recruiter分佣比例', format='%.0f%%', width='small'),
            'Comm1': st.column_config.NumberColumn('Recruiter佣金', disabled=True, format='$%.2f', width='small'),
            'Person2': st.column_config.TextColumn('CFT', width='medium'),
            'Rate2': st.column_config.NumberColumn('CFT比例', format='%.0f%%', width='small'),
            'Split2': st.column_config.NumberColumn('CFT分佣比例', format='%.0f%%', width='small'),
            'Comm2': st.column_config.NumberColumn('CFT佣金', disabled=True, format='$%.2f', width='small'),
        },
        column_order=['_selected', 'MatchStatus', 'Policy', 'Insured', 'CommRate', 'Premium', 'Person1', 'Rate1', 'Split1', 'Comm1', 'Person2', 'Rate2', 'Split2', 'Comm2'],
    )

    # 更新数据
    st.session_state.data['_selected'] = edited['_selected']
    st.session_state.data['Person1'] = edited['Person1']
    st.session_state.data['Rate1'] = edited['Rate1']
    st.session_state.data['Split1'] = edited['Split1']
    st.session_state.data['Person2'] = edited['Person2']
    st.session_state.data['Rate2'] = edited['Rate2']
    st.session_state.data['Split2'] = edited['Split2']

    # ==================== 校验和导出 ====================
    st.markdown("---")

    # 校验
    errors = []
    for idx, row in st.session_state.data.iterrows():
        split_sum = row['Split1'] + row['Split2']
        if split_sum != 100:
            errors.append(f"{row['Policy']}: 分佣={split_sum}%")

    if errors:
        st.error(f"❌ {len(errors)} split errors: " + ", ".join(errors[:5]))
    else:
        st.success("✅ All records validated")

        # 导出
        col1, col2 = st.columns(2)
        with col1:
            # 导出明细
            output = BytesIO()
            export_df = st.session_state.data.copy()
            export_df['Comm1'] = export_df['Premium'] * (export_df['Rate1']/100) * (export_df['Split1']/100)
            export_df['Comm2'] = export_df['Premium'] * (export_df['Rate2']/100) * (export_df['Split2']/100)
            # 重命名列名匹配zhubiao格式
            export_df = export_df.rename(columns={
                'MatchStatus': 'Match Status',
                'CommRate': 'Comm Rate %',
                'Premium': 'Gross Comm Earned',
                'Person1': 'Recruiter',
                'Rate1': 'Recruiter佣金比例',
                'Split1': 'Recruiter分佣比例',
                'Comm1': 'Recruiter佣金',
                'Person2': 'CFT',
                'Rate2': 'CFT比例',
                'Split2': 'CFT分佣比例',
                'Comm2': 'CFT佣金'
            })
            export_cols = ['Policy', 'Insured', 'Match Status', 'Comm Rate %', 'Gross Comm Earned', 'Recruiter', 'Recruiter佣金比例', 'Recruiter分佣比例', 'Recruiter佣金', 'CFT', 'CFT比例', 'CFT分佣比例', 'CFT佣金']
            export_df = export_df[[c for c in export_cols if c in export_df.columns]]
            export_df.to_excel(output, index=False, engine='openpyxl')
            st.download_button("📥 Download Detail", output.getvalue(), f"commission_detail_{datetime.now().strftime('%Y%m%d')}.xlsx")

        with col2:
            # 导出汇总
            if summary_data:
                output2 = BytesIO()
                pd.DataFrame(summary_data).to_excel(output2, index=False, engine='openpyxl')
                st.download_button("📥 Download Summary", output2.getvalue(), f"commission_summary_{datetime.now().strftime('%Y%m%d')}.xlsx")
