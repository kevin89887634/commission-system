"""
💰 佣金管理系统 v2.2
修复：公式改为 Premium × PersonRate × SplitRatio
"""
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ==================== 工具函数 ====================
def normalize_policy(policy_num):
    """标准化保单号：移除LS/NL/L前缀和00后缀"""
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

def format_currency(amount):
    if amount is None or pd.isna(amount):
        return "$0.00"
    return f"${amount:,.2f}"

def is_valid_policy(policy):
    """检查是否为有效保单号"""
    if policy is None or pd.isna(policy):
        return False
    s = str(policy).strip().lower()
    invalid = ['policy', 'nan', 'none', '', '* for', 'exported', 'for ul']
    for p in invalid:
        if p in s:
            return False
    if not any(c.isdigit() for c in s):
        return False
    return True

# ==================== 页面配置 ====================
st.set_page_config(page_title="佣金管理系统", page_icon="💰", layout="wide")

# Session State
if 'df_raw' not in st.session_state:
    st.session_state.df_raw = None
if 'df_splits' not in st.session_state:
    st.session_state.df_splits = None
if 'df_results' not in st.session_state:
    st.session_state.df_results = None

# 侧边栏
with st.sidebar:
    st.title("💰 佣金管理系统")
    st.caption("v2.2")
    st.markdown("---")
    step = st.radio("操作步骤", [
        "1️⃣ 上传数据",
        "2️⃣ 编辑分单",
        "3️⃣ 计算佣金",
        "4️⃣ 对账核验"
    ])
    st.markdown("---")
    if st.session_state.df_raw is not None:
        st.success(f"✅ 已导入 {len(st.session_state.df_raw)} 条")

# ==================== 第一步：上传数据 ====================
if step == "1️⃣ 上传数据":
    st.header("1️⃣ 上传数据")

    file = st.file_uploader("上传 NLG New Business Report", type=['xlsx', 'xls'])

    if file and st.button("📥 导入数据", type="primary"):
        with st.spinner("导入中..."):
            try:
                # 第一步：读取原始数据，找到表头行
                df_raw = pd.read_excel(file, header=None)

                # 动态查找包含 "Policy" 的表头行
                header_row = None
                for idx in range(min(10, len(df_raw))):
                    row_values = [str(v).lower() if pd.notna(v) else '' for v in df_raw.iloc[idx]]
                    if any('policy' in v for v in row_values):
                        header_row = idx
                        break

                if header_row is None:
                    st.error("❌ 找不到表头行（包含'Policy'的行）")
                    st.stop()

                # 第二步：用正确的表头行重新读取
                df = pd.read_excel(file, header=header_row)

                # 标准化列名（处理"Policy #"等带特殊字符的列名）
                col_mapping = {}
                for col in df.columns:
                    col_lower = str(col).lower().strip()
                    if 'policy' in col_lower:
                        col_mapping[col] = 'Policy'
                    elif 'insured' in col_lower or 'annuitant' in col_lower:
                        col_mapping[col] = 'Insured'
                    elif col_lower == 'agent':
                        col_mapping[col] = 'Recruiter'
                    elif 'modal' in col_lower:
                        col_mapping[col] = 'Modal'
                    elif 'aap' in col_lower:
                        col_mapping[col] = 'AAP'
                    elif 'product' in col_lower:
                        col_mapping[col] = 'Product'
                    elif 'status' in col_lower:
                        col_mapping[col] = 'Status'

                df = df.rename(columns=col_mapping)

                # 确保必要的列存在
                required_cols = ['Policy', 'Modal', 'AAP']
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    st.error(f"❌ 缺少必要列: {missing}")
                    st.error(f"当前列: {list(df.columns)}")
                    st.stop()

                # 清洗：过滤无效保单
                df = df[df['Policy'].apply(is_valid_policy)]
                df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
                df['Modal'] = df['Modal'].apply(safe_float)
                df['AAP'] = df['AAP'].apply(safe_float)

                # 过滤有效保费记录
                df = df[(df['AAP'] > 0) | (df['Modal'] > 0)].reset_index(drop=True)

                st.session_state.df_raw = df

                # 生成分单表
                splits_data = []
                for _, row in df.iterrows():
                    modal = safe_float(row['Modal'])
                    aap = safe_float(row['AAP'])
                    # 判断缴费类型
                    if modal > 0 and aap / modal > 6:
                        pay_type = '月缴'
                        premium = modal  # 月缴保费
                    else:
                        pay_type = '年缴'
                        premium = aap  # 年缴保费

                    # 判断佣金比例
                    product = str(row.get('Product', '')).lower()
                    if 'term' in product:
                        comm_rate = 0.67
                    else:
                        comm_rate = 0.80

                        # 获取Recruiter（可能是Agent列）
                    recruiter = ''
                    if 'Recruiter' in row.index:
                        recruiter = str(row['Recruiter']) if pd.notna(row['Recruiter']) else ''

                    splits_data.append({
                        'Policy': row['Policy_Norm'],
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
                st.success(f"✅ 导入成功！{len(df)} 条有效记录")

            except Exception as e:
                st.error(f"❌ 导入失败: {e}")

    if st.session_state.df_raw is not None:
        st.markdown("### 📊 数据预览")
        # 只显示存在的列
        preview_cols = ['Policy', 'Insured', 'Recruiter', 'Product', 'Modal', 'AAP']
        available_cols = [c for c in preview_cols if c in st.session_state.df_raw.columns]
        st.dataframe(
            st.session_state.df_raw[available_cols],
            use_container_width=True
        )

# ==================== 第二步：编辑分单 ====================
elif step == "2️⃣ 编辑分单":
    st.header("2️⃣ 批量编辑分单配置")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先在第1步上传数据")
    else:
        st.info("""
        💡 **公式说明**：
        - **个人佣金 = Premium × Rate × Split**
        - Premium = 月缴保费(Modal) 或 年缴保费(AAP)
        - Rate = 个人佣金比例 (如0.55=55%)
        - Split = 分佣比例 (Split1 + Split2 必须 = 100%)
        """)

        # 编辑表格
        edited_df = st.data_editor(
            st.session_state.df_splits,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                'Policy': st.column_config.TextColumn('保单号', disabled=True, width="small"),
                'Insured': st.column_config.TextColumn('被保人', disabled=True, width="medium"),
                'AAP': st.column_config.NumberColumn('AAP', disabled=True, format="$%.0f", width="small"),
                'Modal': st.column_config.NumberColumn('Modal', disabled=True, format="$%.2f", width="small"),
                'PayType': st.column_config.TextColumn('类型', disabled=True, width="small"),
                'Premium': st.column_config.NumberColumn('Premium', disabled=True, format="$%.2f", width="small"),
                'CommRate': st.column_config.NumberColumn('佣金率', disabled=True, format="%.0f%%", width="small"),
                'Person1': st.column_config.TextColumn('人员1', width="medium"),
                'Rate1': st.column_config.NumberColumn('比例1', min_value=0, max_value=1, step=0.05, format="%.0f%%", width="small"),
                'Split1': st.column_config.NumberColumn('分佣1', min_value=0, max_value=1, step=0.1, format="%.0f%%", width="small"),
                'Person2': st.column_config.TextColumn('人员2', width="medium"),
                'Rate2': st.column_config.NumberColumn('比例2', min_value=0, max_value=1, step=0.05, format="%.0f%%", width="small"),
                'Split2': st.column_config.NumberColumn('分佣2', min_value=0, max_value=1, step=0.1, format="%.0f%%", width="small"),
            },
            hide_index=True
        )

        # 验证
        st.markdown("### ✅ 验证")
        errors = []
        for idx, row in edited_df.iterrows():
            s1 = safe_float(row['Split1'])
            s2 = safe_float(row['Split2'])
            total = s1 + s2
            if abs(total - 1.0) > 0.001 and total > 0:
                errors.append(f"❌ {row['Policy']}: Split总和={total*100:.0f}% (应为100%)")

        if errors:
            for err in errors[:10]:
                st.error(err)
        else:
            st.success("✅ 所有分佣比例正确")

        if st.button("💾 保存配置", type="primary"):
            if errors:
                st.error("❌ 请先修正错误")
            else:
                st.session_state.df_splits = edited_df
                st.session_state.df_results = None
                st.success("✅ 已保存！请前往第3步计算")

# ==================== 第三步：计算佣金 ====================
elif step == "3️⃣ 计算佣金":
    st.header("3️⃣ 计算佣金")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先完成第1、2步")
    else:
        st.markdown("""
        **计算公式**：
        - Gross Comm = Premium × CommRate (80%或67%)
        - Override = Premium × 48%
        - 个人佣金 = Premium × PersonRate × SplitRatio
        - 平台剩余 = Gross + Override - 已分配佣金
        """)

        if st.button("🔄 开始计算", type="primary"):
            results = []
            df = st.session_state.df_splits

            for _, row in df.iterrows():
                policy = row['Policy']
                insured = row['Insured']
                aap = safe_float(row['AAP'])
                premium = safe_float(row['Premium'])
                comm_rate = safe_float(row['CommRate'])
                pay_type = row['PayType']

                # 计算总佣金
                gross_comm = premium * comm_rate
                override_comm = premium * 0.48
                total_comm = gross_comm + override_comm

                # 计算每人
                distributed = 0
                for i in [1, 2]:
                    person = str(row.get(f'Person{i}', '')).strip()
                    rate = safe_float(row.get(f'Rate{i}', 0))
                    split = safe_float(row.get(f'Split{i}', 0))

                    if person and split > 0:
                        # 公式: Premium × Rate × Split
                        person_comm = premium * rate * split
                        distributed += person_comm

                        results.append({
                            'Policy': policy,
                            'Insured': insured,
                            'AAP': aap,
                            'Premium': premium,
                            'PayType': pay_type,
                            'CommRate': comm_rate,
                            'GrossComm': gross_comm,
                            'Override': override_comm,
                            'TotalComm': total_comm,
                            'Person': person,
                            'Rate': rate,
                            'Split': split,
                            'PersonComm': person_comm,
                        })

                # 平台剩余
                platform = total_comm - distributed
                if platform > 0.01:
                    results.append({
                        'Policy': policy,
                        'Insured': insured,
                        'AAP': aap,
                        'Premium': premium,
                        'PayType': pay_type,
                        'CommRate': comm_rate,
                        'GrossComm': gross_comm,
                        'Override': override_comm,
                        'TotalComm': total_comm,
                        'Person': '【平台】',
                        'Rate': 0,
                        'Split': 0,
                        'PersonComm': platform,
                    })

            st.session_state.df_results = pd.DataFrame(results)
            st.success("✅ 计算完成！")

        if st.session_state.df_results is not None:
            df_r = st.session_state.df_results

            # 汇总
            st.markdown("### 📊 汇总")
            unique_policies = df_r.drop_duplicates('Policy')
            total_premium = unique_policies['Premium'].sum()
            total_gross = unique_policies['GrossComm'].sum()
            total_override = unique_policies['Override'].sum()
            total_comm = df_r['PersonComm'].sum()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("保单数", len(unique_policies))
            c2.metric("总Premium", format_currency(total_premium))
            c3.metric("总Gross", format_currency(total_gross))
            c4.metric("总Override", format_currency(total_override))

            # 按人员
            st.markdown("### 👥 按人员汇总")
            person_sum = df_r.groupby('Person')['PersonComm'].sum().reset_index()
            person_sum.columns = ['人员', '佣金']
            person_sum = person_sum.sort_values('佣金', ascending=False)
            st.dataframe(person_sum.style.format({'佣金': '${:,.2f}'}), use_container_width=True)

            # 明细
            st.markdown("### 📋 明细")
            display_cols = ['Policy', 'Insured', 'Premium', 'PayType', 'GrossComm', 'Override', 'Person', 'Rate', 'Split', 'PersonComm']
            st.dataframe(
                df_r[display_cols].style.format({
                    'Premium': '${:,.2f}',
                    'GrossComm': '${:,.2f}',
                    'Override': '${:,.2f}',
                    'Rate': '{:.0%}',
                    'Split': '{:.0%}',
                    'PersonComm': '${:,.2f}',
                }),
                use_container_width=True
            )

            # 导出
            st.markdown("### 📥 导出")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                person_sum.to_excel(writer, sheet_name='人员汇总', index=False)
                df_r.to_excel(writer, sheet_name='佣金明细', index=False)
                st.session_state.df_splits.to_excel(writer, sheet_name='分单配置', index=False)
            output.seek(0)
            st.download_button("📥 下载Excel", data=output,
                             file_name=f"佣金报表_{datetime.now().strftime('%Y%m%d')}.xlsx")

# ==================== 第四步：对账 ====================
elif step == "4️⃣ 对账核验":
    st.header("4️⃣ 对账核验")

    if st.session_state.df_results is None:
        st.warning("⚠️ 请先完成第3步")
    else:
        st.info("上传对账单进行比对")

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 🏢 Gross Commission")
            gross_file = st.file_uploader("NLG Payable/Pending Gross", type=['xlsx'], key='gross')
        with col2:
            st.markdown("#### 📋 Override")
            override_file = st.file_uploader("Override by Policy", type=['xlsx'], key='override')

        if st.button("🔍 开始对账", type="primary"):
            df_r = st.session_state.df_results

            # 按保单汇总计算结果
            calc = df_r.groupby('Policy').agg({
                'GrossComm': 'first',
                'Override': 'first',
            }).reset_index()

            # 读取对账单
            actual_gross = {}
            actual_override = {}

            if gross_file:
                try:
                    df_g = pd.read_excel(gross_file, skiprows=4)
                    for _, row in df_g.iterrows():
                        p = normalize_policy(row.iloc[2])  # Policy # 在第3列
                        if p:
                            amt = safe_float(row.iloc[6])  # Gross Com. Paid 在第7列
                            actual_gross[p] = actual_gross.get(p, 0) + amt
                except Exception as e:
                    st.error(f"Gross文件格式错误: {e}")

            if override_file:
                try:
                    df_o = pd.read_excel(override_file, skiprows=1)
                    for _, row in df_o.iterrows():
                        p = normalize_policy(row.iloc[2])  # Policy# 在第3列
                        if p:
                            amt = safe_float(row.iloc[5])  # Total Amount 在第6列
                            actual_override[p] = actual_override.get(p, 0) + amt
                except Exception as e:
                    st.error(f"Override文件格式错误: {e}")

            # 对账
            reconcile = []
            for _, row in calc.iterrows():
                policy = row['Policy']
                calc_gross = row['GrossComm']
                calc_override = row['Override']

                act_gross = actual_gross.get(policy, 0)
                act_override = actual_override.get(policy, 0)

                gross_diff = act_gross - calc_gross
                override_diff = act_override - calc_override

                gross_ok = '✅' if abs(gross_diff) < 1 else ('⚠️' if act_gross == 0 else '❌')
                override_ok = '✅' if abs(override_diff) < 1 else ('⚠️' if act_override == 0 else '❌')

                reconcile.append({
                    '保单号': policy,
                    '计算Gross': calc_gross,
                    '实际Gross': act_gross,
                    'Gross差额': gross_diff,
                    'Gross状态': gross_ok,
                    '计算Override': calc_override,
                    '实际Override': act_override,
                    'Override差额': override_diff,
                    'Override状态': override_ok,
                })

            df_rec = pd.DataFrame(reconcile)

            # 统计
            st.markdown("### 📊 对账结果")
            gross_match = (df_rec['Gross状态'] == '✅').sum()
            override_match = (df_rec['Override状态'] == '✅').sum()
            total = len(df_rec)

            c1, c2, c3 = st.columns(3)
            c1.metric("总保单", total)
            c2.metric("Gross匹配", f"{gross_match}/{total}")
            c3.metric("Override匹配", f"{override_match}/{total}")

            # 差异
            df_diff = df_rec[(df_rec['Gross状态'] == '❌') | (df_rec['Override状态'] == '❌')]
            if len(df_diff) > 0:
                st.markdown("### ❌ 差异记录")
                st.dataframe(df_diff.style.format({
                    '计算Gross': '${:,.2f}', '实际Gross': '${:,.2f}', 'Gross差额': '${:,.2f}',
                    '计算Override': '${:,.2f}', '实际Override': '${:,.2f}', 'Override差额': '${:,.2f}',
                }), use_container_width=True)

            # 完整表
            st.markdown("### 📋 完整对账表")
            st.dataframe(df_rec.style.format({
                '计算Gross': '${:,.2f}', '实际Gross': '${:,.2f}', 'Gross差额': '${:,.2f}',
                '计算Override': '${:,.2f}', '实际Override': '${:,.2f}', 'Override差额': '${:,.2f}',
            }), use_container_width=True)
