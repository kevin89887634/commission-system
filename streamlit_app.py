"""
💰 佣金管理系统 v2.8
- 上传zhubiao自动匹配分单
- 表格内选择+批量编辑
- 数据保存/加载/删除
"""
import streamlit as st
import pandas as pd
import re
import json
from io import BytesIO
from datetime import datetime

# ==================== 工具函数 ====================
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
    if not (s.upper().startswith(('LS', 'NL', 'L')) or s[0].isdigit()):
        return False
    return True

def parse_nlg_file(uploaded_file):
    for header_row in [5, 4, 6, 3, 1, 0]:
        try:
            df = pd.read_excel(uploaded_file, header=header_row, engine='openpyxl')
            uploaded_file.seek(0)
            cols_lower = [str(c).lower() for c in df.columns]
            has_policy = any('policy' in c for c in cols_lower)
            if has_policy and len(df) > 0:
                policy_col = next((c for c in df.columns if 'policy' in str(c).lower()), None)
                first_val = str(df[policy_col].iloc[0]) if len(df) > 0 else ''
                if is_valid_policy(first_val):
                    return df, header_row, None
        except:
            uploaded_file.seek(0)
            continue
    return None, None, "无法找到有效的表头行"

# ==================== 页面配置 ====================
st.set_page_config(page_title="佣金管理系统", page_icon="💰", layout="wide")

# Session State 初始化
if 'df_raw' not in st.session_state:
    st.session_state.df_raw = None
if 'df_splits' not in st.session_state:
    st.session_state.df_splits = None
if 'df_results' not in st.session_state:
    st.session_state.df_results = None
if 'saved_datasets' not in st.session_state:
    st.session_state.saved_datasets = {}  # {name: {'raw': df, 'splits': df, 'time': str}}
if 'current_dataset' not in st.session_state:
    st.session_state.current_dataset = None

# ==================== 侧边栏 ====================
with st.sidebar:
    st.title("💰 佣金管理系统")
    st.caption("v2.8")
    st.markdown("---")

    step = st.radio("操作步骤", [
        "1️⃣ 上传数据",
        "2️⃣ 编辑分单",
        "3️⃣ 计算佣金",
        "4️⃣ 对账核验"
    ])

    st.markdown("---")

    # 当前数据状态
    if st.session_state.df_raw is not None:
        st.success(f"✅ 当前: {len(st.session_state.df_raw)} 条")
        if st.session_state.current_dataset:
            st.caption(f"📁 {st.session_state.current_dataset}")

    # 已保存的数据
    if st.session_state.saved_datasets:
        st.markdown("### 📂 已保存数据")
        for name, data in list(st.session_state.saved_datasets.items()):
            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button(f"📄 {name}", key=f"load_{name}", use_container_width=True):
                    st.session_state.df_raw = data['raw'].copy()
                    st.session_state.df_splits = data['splits'].copy()
                    st.session_state.current_dataset = name
                    st.session_state.df_results = None
                    st.rerun()
            with col2:
                if st.button("🗑️", key=f"del_{name}"):
                    del st.session_state.saved_datasets[name]
                    if st.session_state.current_dataset == name:
                        st.session_state.current_dataset = None
                    st.rerun()

# ==================== 第一步：上传数据 ====================
if step == "1️⃣ 上传数据":
    st.header("1️⃣ 上传数据")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📄 NLG New Business Report")
        uploaded_file = st.file_uploader("必填：NLG报表", type=['xlsx', 'xls'], key="nlg")
    with col2:
        st.markdown("#### 📋 分单模板 (可选)")
        template_file = st.file_uploader("可选：已有分单表(zhubiao)", type=['xlsx', 'xls'], key="template")

    if uploaded_file and st.button("📥 导入数据", type="primary"):
        with st.spinner("导入中..."):
            try:
                # 解析NLG文件
                df, header_row, error = parse_nlg_file(uploaded_file)
                if error:
                    st.error(f"❌ {error}")
                    st.stop()

                # 标准化列名
                col_map = {}
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if 'policy' in col_lower:
                        col_map[col] = 'Policy'
                    elif 'insured' in col_lower or 'annuitant' in col_lower:
                        col_map[col] = 'Insured'
                    elif col_lower == 'agent':
                        col_map[col] = 'Recruiter'
                    elif 'modal' in col_lower:
                        col_map[col] = 'Modal'
                    elif 'aap' in col_lower:
                        col_map[col] = 'AAP'
                    elif 'product' in col_lower:
                        col_map[col] = 'Product'
                df = df.rename(columns=col_map)

                # 过滤和处理
                df = df[df['Policy'].apply(is_valid_policy)]
                df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
                df['Modal'] = df['Modal'].apply(safe_float) if 'Modal' in df.columns else 0
                df['AAP'] = df['AAP'].apply(safe_float) if 'AAP' in df.columns else 0
                df = df[(df['AAP'] > 0) | (df['Modal'] > 0)].reset_index(drop=True)

                if len(df) == 0:
                    st.error("❌ 没有有效数据")
                    st.stop()

                st.session_state.df_raw = df

                # 解析分单模板 (如果有)
                template_map = {}
                if template_file:
                    try:
                        # 尝试读取分单模板
                        df_tpl = pd.read_excel(template_file, header=0, engine='openpyxl')
                        template_file.seek(0)

                        # 查找关键列
                        policy_col = next((c for c in df_tpl.columns if 'policy' in str(c).lower()), None)

                        # 查找分佣人1相关列 (H,I,J 或 CFT相关 或 经纪人相关)
                        person1_col = None
                        rate1_col = None
                        split1_col = None
                        person2_col = None
                        rate2_col = None
                        split2_col = None

                        cols = list(df_tpl.columns)
                        for i, col in enumerate(cols):
                            col_str = str(col).lower()
                            # 找CFT或第一个经纪人列
                            if 'cft' in col_str or col_str == '经纪人':
                                person1_col = col
                                # 后面两列可能是比例和分佣
                                if i + 1 < len(cols):
                                    rate1_col = cols[i + 1]
                                if i + 2 < len(cols):
                                    split1_col = cols[i + 2]
                            # 找第二个分佣人
                            if i > 0 and person1_col and col_str in ['经纪人', '分佣人2', 'agent2']:
                                person2_col = col
                                if i + 1 < len(cols):
                                    rate2_col = cols[i + 1]
                                if i + 2 < len(cols):
                                    split2_col = cols[i + 2]

                        # 如果没找到，尝试按位置（H=7, I=8, J=9, L=11, M=12, N=13）
                        if not person1_col and len(cols) > 9:
                            person1_col = cols[7] if len(cols) > 7 else None  # H列
                            rate1_col = cols[8] if len(cols) > 8 else None    # I列
                            split1_col = cols[9] if len(cols) > 9 else None   # J列
                            person2_col = cols[11] if len(cols) > 11 else None # L列
                            rate2_col = cols[12] if len(cols) > 12 else None   # M列
                            split2_col = cols[13] if len(cols) > 13 else None  # N列

                        if policy_col and person1_col:
                            for _, row in df_tpl.iterrows():
                                policy_val = str(row.get(policy_col, ''))
                                policy_norm = normalize_policy(policy_val)
                                if policy_norm:
                                    template_map[policy_norm] = {
                                        'Person1': str(row.get(person1_col, '')) if pd.notna(row.get(person1_col, '')) else '',
                                        'Rate1': safe_float(row.get(rate1_col, 0.55)),
                                        'Split1': safe_float(row.get(split1_col, 1.0)),
                                        'Person2': str(row.get(person2_col, '')) if pd.notna(row.get(person2_col, '')) else '',
                                        'Rate2': safe_float(row.get(rate2_col, 0.55)),
                                        'Split2': safe_float(row.get(split2_col, 0)),
                                    }
                            st.success(f"✅ 模板匹配: {len(template_map)} 条分单规则")
                    except Exception as e:
                        st.warning(f"⚠️ 模板解析失败: {e}，将使用默认分单")

                # 生成分单表
                splits_data = []
                matched_count = 0
                for _, row in df.iterrows():
                    modal = safe_float(row.get('Modal', 0))
                    aap = safe_float(row.get('AAP', 0))
                    if modal > 0 and aap > 0 and aap / modal > 6:
                        pay_type = '月缴'
                        premium = modal
                    else:
                        pay_type = '年缴'
                        premium = aap if aap > 0 else modal

                    product = str(row.get('Product', '')).lower()
                    comm_rate = 0.67 if 'term' in product else 0.80
                    recruiter = str(row.get('Recruiter', '')) if pd.notna(row.get('Recruiter', '')) else ''

                    policy_norm = row['Policy_Norm']

                    # 检查是否有模板匹配
                    if policy_norm in template_map:
                        tpl = template_map[policy_norm]
                        matched_count += 1
                        splits_data.append({
                            '选择': False,
                            'Policy': policy_norm,
                            'Insured': str(row.get('Insured', '')) if pd.notna(row.get('Insured', '')) else '',
                            'AAP': aap,
                            'Modal': modal,
                            'PayType': pay_type,
                            'Premium': premium,
                            'CommRate': comm_rate,
                            'Person1': tpl['Person1'],
                            'Rate1': tpl['Rate1'],
                            'Split1': tpl['Split1'],
                            'Person2': tpl['Person2'],
                            'Rate2': tpl['Rate2'],
                            'Split2': tpl['Split2'],
                        })
                    else:
                        splits_data.append({
                            '选择': False,
                            'Policy': policy_norm,
                            'Insured': str(row.get('Insured', '')) if pd.notna(row.get('Insured', '')) else '',
                            'AAP': aap,
                            'Modal': modal,
                            'PayType': pay_type,
                            'Premium': premium,
                            'CommRate': comm_rate,
                            'Person1': recruiter,
                            'Rate1': 0.55,
                            'Split1': 1.0,
                            'Person2': '',
                            'Rate2': 0.55,
                            'Split2': 0.0,
                        })

                st.session_state.df_splits = pd.DataFrame(splits_data)
                st.session_state.df_results = None
                st.session_state.current_dataset = None

                if matched_count > 0:
                    st.success(f"✅ 导入成功！{len(df)} 条记录，其中 {matched_count} 条已自动匹配分单")
                else:
                    st.success(f"✅ 导入成功！{len(df)} 条记录")

            except Exception as e:
                st.error(f"❌ 导入失败: {e}")

    # 保存当前数据
    if st.session_state.df_raw is not None:
        st.markdown("---")
        st.markdown("### 💾 保存数据")
        col1, col2 = st.columns([3, 1])
        with col1:
            save_name = st.text_input("数据名称", value=f"数据_{datetime.now().strftime('%m%d_%H%M')}")
        with col2:
            st.write("")  # 占位
            st.write("")
            if st.button("💾 保存", type="primary"):
                st.session_state.saved_datasets[save_name] = {
                    'raw': st.session_state.df_raw.copy(),
                    'splits': st.session_state.df_splits.copy(),
                    'time': datetime.now().strftime('%Y-%m-%d %H:%M')
                }
                st.session_state.current_dataset = save_name
                st.success(f"✅ 已保存: {save_name}")

        # 数据预览
        st.markdown("### 📊 数据预览")
        preview_cols = [c for c in ['Policy', 'Insured', 'Recruiter', 'Product', 'Modal', 'AAP']
                       if c in st.session_state.df_raw.columns]
        st.dataframe(st.session_state.df_raw[preview_cols], use_container_width=True)

# ==================== 第二步：编辑分单 ====================
elif step == "2️⃣ 编辑分单":
    st.header("2️⃣ 编辑分单")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先上传数据")
        st.stop()

    # 批量编辑工具栏
    st.markdown("### 🔧 批量编辑工具")
    st.caption("先在表格中勾选要修改的行，然后设置分佣信息并点击应用")

    tool_cols = st.columns([1, 1, 1, 1, 1, 1, 2])
    with tool_cols[0]:
        batch_person1 = st.text_input("分佣人1", key="bp1", placeholder="姓名")
    with tool_cols[1]:
        batch_rate1 = st.number_input("比例1", value=0.55, min_value=0.0, max_value=1.0, step=0.01, key="br1")
    with tool_cols[2]:
        batch_split1 = st.number_input("分成1", value=1.0, min_value=0.0, max_value=1.0, step=0.1, key="bs1")
    with tool_cols[3]:
        batch_person2 = st.text_input("分佣人2", key="bp2", placeholder="可选")
    with tool_cols[4]:
        batch_rate2 = st.number_input("比例2", value=0.55, min_value=0.0, max_value=1.0, step=0.01, key="br2")
    with tool_cols[5]:
        batch_split2 = st.number_input("分成2", value=0.0, min_value=0.0, max_value=1.0, step=0.1, key="bs2")

    # 验证和应用按钮
    total_split = batch_split1 + batch_split2
    with tool_cols[6]:
        if abs(total_split - 1.0) > 0.001:
            st.error(f"分成={total_split:.1f}≠1")
            apply_disabled = True
        else:
            st.success(f"分成={total_split:.1f}✓")
            apply_disabled = False

        if st.button("📝 应用到选中行", disabled=apply_disabled, type="primary", use_container_width=True):
            df = st.session_state.df_splits.copy()
            count = 0
            for idx, row in df.iterrows():
                if row.get('选择', False):
                    if batch_person1:
                        df.at[idx, 'Person1'] = batch_person1
                    df.at[idx, 'Rate1'] = batch_rate1
                    df.at[idx, 'Split1'] = batch_split1
                    df.at[idx, 'Person2'] = batch_person2
                    df.at[idx, 'Rate2'] = batch_rate2
                    df.at[idx, 'Split2'] = batch_split2
                    df.at[idx, '选择'] = False  # 取消选择
                    count += 1
            if count > 0:
                st.session_state.df_splits = df
                st.success(f"✅ 已修改 {count} 条")
                st.rerun()
            else:
                st.warning("⚠️ 请先勾选要修改的行")

    # 快捷操作
    qk_cols = st.columns(4)
    with qk_cols[0]:
        if st.button("☑️ 全选"):
            st.session_state.df_splits['选择'] = True
            st.rerun()
    with qk_cols[1]:
        if st.button("⬜ 取消全选"):
            st.session_state.df_splits['选择'] = False
            st.rerun()
    with qk_cols[2]:
        selected_count = st.session_state.df_splits['选择'].sum()
        st.info(f"已选: {selected_count} 条")

    st.markdown("---")

    # 可编辑表格
    edited_df = st.data_editor(
        st.session_state.df_splits,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            '选择': st.column_config.CheckboxColumn('✓', default=False),
            'Policy': st.column_config.TextColumn('保单号', disabled=True, width="small"),
            'Insured': st.column_config.TextColumn('被保人', disabled=True, width="small"),
            'AAP': st.column_config.NumberColumn('AAP', disabled=True, format="$%.0f", width="small"),
            'Modal': st.column_config.NumberColumn('Modal', disabled=True, format="$%.0f", width="small"),
            'PayType': st.column_config.TextColumn('类型', disabled=True, width="small"),
            'Premium': st.column_config.NumberColumn('保费', disabled=True, format="$%.0f", width="small"),
            'CommRate': st.column_config.NumberColumn('佣金率', format="%.2f", width="small"),
            'Person1': st.column_config.TextColumn('分佣人1', width="medium"),
            'Rate1': st.column_config.NumberColumn('比例1', format="%.2f", width="small"),
            'Split1': st.column_config.NumberColumn('分成1', format="%.1f", width="small"),
            'Person2': st.column_config.TextColumn('分佣人2', width="medium"),
            'Rate2': st.column_config.NumberColumn('比例2', format="%.2f", width="small"),
            'Split2': st.column_config.NumberColumn('分成2', format="%.1f", width="small"),
        },
        column_order=['选择', 'Policy', 'Insured', 'Premium', 'PayType', 'Person1', 'Rate1', 'Split1', 'Person2', 'Rate2', 'Split2'],
        hide_index=True,
    )

    # 实时更新选择状态
    st.session_state.df_splits = edited_df

    st.markdown("---")

    # 验证并保存
    errors = []
    for idx, row in edited_df.iterrows():
        split_sum = safe_float(row.get('Split1', 0)) + safe_float(row.get('Split2', 0))
        if abs(split_sum - 1.0) > 0.001:
            errors.append(f"保单 {row['Policy']}: 分成={split_sum:.1f}")

    if errors:
        st.error(f"❌ {len(errors)} 条记录分成比例错误（应为1.0）:")
        st.warning("、".join(errors[:5]) + ("..." if len(errors) > 5 else ""))
    else:
        st.success("✅ 所有记录验证通过")

    if st.button("💾 保存修改", type="primary", disabled=len(errors) > 0):
        st.session_state.df_splits = edited_df
        st.success("✅ 已保存")

# ==================== 第三步：计算佣金 ====================
elif step == "3️⃣ 计算佣金":
    st.header("3️⃣ 计算佣金")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先完成前面的步骤")
        st.stop()

    # 先验证
    errors = []
    for idx, row in st.session_state.df_splits.iterrows():
        split_sum = safe_float(row.get('Split1', 0)) + safe_float(row.get('Split2', 0))
        if abs(split_sum - 1.0) > 0.001:
            errors.append(row['Policy'])

    if errors:
        st.error(f"❌ 有 {len(errors)} 条记录分成比例错误，请先修正")
        st.stop()

    if st.button("🧮 开始计算", type="primary"):
        results = []
        df = st.session_state.df_splits

        for _, row in df.iterrows():
            policy = row['Policy']
            premium = safe_float(row['Premium'])
            comm_rate = safe_float(row.get('CommRate', 0.80))

            gross = premium * comm_rate
            override = premium * 0.48
            total_comm = premium * (comm_rate + 0.48)

            for i in [1, 2]:
                person = str(row.get(f'Person{i}', '')).strip()
                rate = safe_float(row.get(f'Rate{i}', 0))
                split = safe_float(row.get(f'Split{i}', 0))

                if person and split > 0:
                    person_comm = premium * rate * split
                    results.append({
                        'Policy': policy,
                        'Insured': row.get('Insured', ''),
                        'Premium': premium,
                        'GrossComm': gross,
                        'Override': override,
                        'TotalComm': total_comm,
                        'Person': person,
                        'Rate': rate,
                        'Split': split,
                        'PersonComm': person_comm,
                    })

        if results:
            st.session_state.df_results = pd.DataFrame(results)
            st.success(f"✅ 计算完成！{len(results)} 条记录")

    if st.session_state.df_results is not None:
        st.markdown("### 📊 计算结果")
        st.dataframe(st.session_state.df_results, use_container_width=True)

        st.markdown("### 📈 分人汇总")
        summary = st.session_state.df_results.groupby('Person').agg({
            'PersonComm': 'sum',
            'Policy': 'count'
        }).rename(columns={'Policy': '单数', 'PersonComm': '佣金总额'})
        summary['佣金总额'] = summary['佣金总额'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(summary, use_container_width=True)

        st.markdown("### 📥 导出")
        output = BytesIO()
        st.session_state.df_results.to_excel(output, index=False, engine='openpyxl')
        st.download_button(
            "📥 下载Excel",
            data=output.getvalue(),
            file_name=f"commission_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ==================== 第四步：对账核验 ====================
elif step == "4️⃣ 对账核验":
    st.header("4️⃣ 对账核验")

    if st.session_state.df_results is None:
        st.warning("⚠️ 请先完成佣金计算")
        st.stop()

    col1, col2 = st.columns(2)
    with col1:
        override_file = st.file_uploader("Override by Policy", type=['xlsx', 'xls'], key='override')
    with col2:
        gross_file = st.file_uploader("Payable Gross Commission", type=['xlsx', 'xls'], key='gross')

    if st.button("🔍 开始对账", type="primary"):
        results = st.session_state.df_results.copy()

        if override_file:
            try:
                df_ov = pd.read_excel(override_file, header=1, engine='openpyxl')
                policy_col = next((c for c in df_ov.columns if 'policy' in str(c).lower()), None)
                amount_col = next((c for c in df_ov.columns if 'amount' in str(c).lower() or 'total' in str(c).lower()), None)
                if policy_col and amount_col:
                    df_ov['Policy_Norm'] = df_ov[policy_col].apply(lambda x: normalize_policy(str(x)))
                    override_map = dict(zip(df_ov['Policy_Norm'], df_ov[amount_col].apply(safe_float)))
                    results['Override_Actual'] = results['Policy'].map(override_map)
                    st.success(f"✅ Override: {len(override_map)} 条")
            except Exception as e:
                st.error(f"❌ Override解析失败: {e}")

        if gross_file:
            try:
                df_gr = pd.read_excel(gross_file, header=4, engine='openpyxl')
                policy_col = next((c for c in df_gr.columns if 'policy' in str(c).lower()), None)
                gross_col = next((c for c in df_gr.columns if 'gross' in str(c).lower()), None)
                if policy_col and gross_col:
                    df_gr['Policy_Norm'] = df_gr[policy_col].apply(lambda x: normalize_policy(str(x)))
                    gross_map = dict(zip(df_gr['Policy_Norm'], df_gr[gross_col].apply(safe_float)))
                    results['Gross_Actual'] = results['Policy'].map(gross_map)
                    st.success(f"✅ Gross: {len(gross_map)} 条")
            except Exception as e:
                st.error(f"❌ Gross解析失败: {e}")

        st.markdown("### 📊 对账结果")
        st.dataframe(results, use_container_width=True)
