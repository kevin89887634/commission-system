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
    return None, "无法解析文件"

# Session State
if 'data' not in st.session_state:
    st.session_state.data = None

# ==================== 页面标题 ====================
st.title("💰 佣金管理系统 v3.0")

# ==================== 上传区域 ====================
with st.expander("📤 上传NLG报表", expanded=st.session_state.data is None):
    uploaded_file = st.file_uploader("选择Excel文件", type=['xlsx', 'xls'])

    if uploaded_file and st.button("📥 导入", type="primary"):
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
                st.error("❌ 没有有效数据")
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

                    rows.append({
                        'Policy': row['Policy_Norm'],
                        'Insured': str(row.get('Insured', ''))[:20] if pd.notna(row.get('Insured')) else '',
                        'Premium': premium,
                        'CommRate': comm_rate,
                        'Agent': agent,
                        'Person1': agent,  # 默认用Agent
                        'Rate1': 0.55,
                        'Split1': 1.0,
                        'Person2': '',
                        'Rate2': 0.55,
                        'Split2': 0.0,
                    })

                st.session_state.data = pd.DataFrame(rows)
                st.success(f"✅ 导入成功！{len(rows)} 条记录")
                st.rerun()

# ==================== 主数据表 ====================
if st.session_state.data is not None:
    df = st.session_state.data.copy()

    # 计算每行的佣金
    df['Comm1'] = df['Premium'] * df['Rate1'] * df['Split1']
    df['Comm2'] = df['Premium'] * df['Rate2'] * df['Split2']
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
            st.metric("📊 总佣金", f"${total_comm:,.2f}")
        with col2:
            st.metric("👥 分佣人数", len(summary_data))
        with col3:
            st.metric("📋 记录数", len(df))

        # 汇总表
        st.markdown("### 📈 分人汇总")
        for _, row in summary_df.iterrows():
            pct = row['Commission'] / total_comm * 100 if total_comm > 0 else 0
            st.write(f"**{row['Person']}**: ${row['Commission']:,.2f} ({pct:.1f}%) - {row['Count']}单")

    st.markdown("---")

    # ==================== 批量设置 ====================
    st.markdown("### 🔧 批量设置")
    bcol1, bcol2, bcol3, bcol4, bcol5, bcol6, bcol7 = st.columns([2, 1, 1, 2, 1, 1, 2])

    with bcol1:
        batch_p1 = st.text_input("分佣人1", key="bp1")
    with bcol2:
        batch_r1 = st.number_input("比例1", 0.0, 1.0, 0.55, 0.01, key="br1")
    with bcol3:
        batch_s1 = st.number_input("分成1", 0.0, 1.0, 1.0, 0.1, key="bs1")
    with bcol4:
        batch_p2 = st.text_input("分佣人2", key="bp2")
    with bcol5:
        batch_r2 = st.number_input("比例2", 0.0, 1.0, 0.55, 0.01, key="br2")
    with bcol6:
        batch_s2 = st.number_input("分成2", 0.0, 1.0, 0.0, 0.1, key="bs2")

    with bcol7:
        total_split = batch_s1 + batch_s2
        if abs(total_split - 1.0) < 0.01:
            st.success(f"✓ 分成={total_split:.1f}")
            can_apply = True
        else:
            st.error(f"✗ 分成={total_split:.1f}≠1")
            can_apply = False

        if st.button("📝 应用到选中", disabled=not can_apply, type="primary"):
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
                st.warning("请先选择行")

    # 快捷按钮
    qcol1, qcol2, qcol3, qcol4 = st.columns(4)
    with qcol1:
        if st.button("☑️ 全选"):
            st.session_state.data['_selected'] = True
            st.rerun()
    with qcol2:
        if st.button("⬜ 取消全选"):
            st.session_state.data['_selected'] = False
            st.rerun()
    with qcol3:
        if '_selected' in st.session_state.data.columns:
            st.info(f"已选: {st.session_state.data['_selected'].sum()} 条")

    st.markdown("---")

    # ==================== 数据表 ====================
    st.markdown("### 📋 明细数据")

    # 添加选择列
    if '_selected' not in st.session_state.data.columns:
        st.session_state.data['_selected'] = False

    # 显示表格
    display_df = st.session_state.data[['_selected', 'Policy', 'Insured', 'Premium', 'Person1', 'Rate1', 'Split1', 'Person2', 'Rate2', 'Split2']].copy()

    # 计算佣金用于显示
    display_df['Comm1'] = st.session_state.data['Premium'] * st.session_state.data['Rate1'] * st.session_state.data['Split1']
    display_df['Comm2'] = st.session_state.data['Premium'] * st.session_state.data['Rate2'] * st.session_state.data['Split2']

    edited = st.data_editor(
        display_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            '_selected': st.column_config.CheckboxColumn('✓', default=False, width='small'),
            'Policy': st.column_config.TextColumn('保单号', disabled=True, width='small'),
            'Insured': st.column_config.TextColumn('被保人', disabled=True, width='small'),
            'Premium': st.column_config.NumberColumn('保费', disabled=True, format='$%.0f', width='small'),
            'Person1': st.column_config.TextColumn('分佣人1', width='medium'),
            'Rate1': st.column_config.NumberColumn('比例1', format='%.2f', width='small'),
            'Split1': st.column_config.NumberColumn('分成1', format='%.1f', width='small'),
            'Comm1': st.column_config.NumberColumn('佣金1', disabled=True, format='$%.2f', width='small'),
            'Person2': st.column_config.TextColumn('分佣人2', width='medium'),
            'Rate2': st.column_config.NumberColumn('比例2', format='%.2f', width='small'),
            'Split2': st.column_config.NumberColumn('分成2', format='%.1f', width='small'),
            'Comm2': st.column_config.NumberColumn('佣金2', disabled=True, format='$%.2f', width='small'),
        },
        column_order=['_selected', 'Policy', 'Insured', 'Premium', 'Person1', 'Rate1', 'Split1', 'Comm1', 'Person2', 'Rate2', 'Split2', 'Comm2'],
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
        if abs(split_sum - 1.0) > 0.01:
            errors.append(f"{row['Policy']}: 分成={split_sum:.1f}")

    if errors:
        st.error(f"❌ {len(errors)} 条分成比例错误: " + ", ".join(errors[:5]))
    else:
        st.success("✅ 所有记录校验通过")

        # 导出
        col1, col2 = st.columns(2)
        with col1:
            # 导出明细
            output = BytesIO()
            export_df = st.session_state.data.copy()
            export_df['Comm1'] = export_df['Premium'] * export_df['Rate1'] * export_df['Split1']
            export_df['Comm2'] = export_df['Premium'] * export_df['Rate2'] * export_df['Split2']
            export_df = export_df[['Policy', 'Insured', 'Premium', 'Person1', 'Rate1', 'Split1', 'Comm1', 'Person2', 'Rate2', 'Split2', 'Comm2']]
            export_df.to_excel(output, index=False, engine='openpyxl')
            st.download_button("📥 下载明细", output.getvalue(), f"commission_detail_{datetime.now().strftime('%Y%m%d')}.xlsx")

        with col2:
            # 导出汇总
            if summary_data:
                output2 = BytesIO()
                pd.DataFrame(summary_data).to_excel(output2, index=False, engine='openpyxl')
                st.download_button("📥 下载汇总", output2.getvalue(), f"commission_summary_{datetime.now().strftime('%Y%m%d')}.xlsx")
