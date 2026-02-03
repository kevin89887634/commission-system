"""
💰 佣金管理系统 v2.1
流程: 上传数据 → 批量编辑分单 → 计算佣金 → 对账核验
"""
import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime

# ==================== 工具函数 ====================
def normalize_policy(policy_num):
    """标准化保单号"""
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

# ==================== 数据导入 ====================
def import_new_business(file):
    """导入并清洗 New Business Report"""
    df = pd.read_excel(file, skiprows=4)
    cols = ['Policy', 'Insured', 'Recruiter', 'Status', 'Delivery',
            'Action', 'SubmitDate', 'Modal', 'Product', 'Sent',
            'Owner', 'SubmitMethod', 'CaseManager', 'AAP',
            'AgentNum', 'Agency', 'CompanyCode', 'Bookmark']
    df.columns = cols[:len(df.columns)]

    # 清洗
    df = df[df['Policy'].apply(is_valid_policy)]
    df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
    df['Modal'] = df['Modal'].apply(safe_float)
    df['AAP'] = df['AAP'].apply(safe_float)
    df = df[df['AAP'] > 0]
    df = df.reset_index(drop=True)
    return df

# ==================== 页面配置 ====================
st.set_page_config(page_title="佣金管理系统", page_icon="💰", layout="wide")

# Session State 初始化
if 'df_raw' not in st.session_state:
    st.session_state.df_raw = None
if 'df_splits' not in st.session_state:
    st.session_state.df_splits = None
if 'df_results' not in st.session_state:
    st.session_state.df_results = None

# 侧边栏
with st.sidebar:
    st.title("💰 佣金管理系统")
    st.caption("v2.1")
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
    if st.session_state.df_splits is not None:
        st.info(f"📋 已配置 {len(st.session_state.df_splits)} 条分单")

# ==================== 第一步：上传数据 ====================
if step == "1️⃣ 上传数据":
    st.header("1️⃣ 上传数据")

    file = st.file_uploader("上传 NLG New Business Report", type=['xlsx', 'xls'])

    if file and st.button("📥 导入数据", type="primary"):
        with st.spinner("导入中..."):
            try:
                df = import_new_business(file)
                st.session_state.df_raw = df

                # 生成初始分单表（每个保单一行，默认Recruiter 55%分佣100%）
                splits_data = []
                for _, row in df.iterrows():
                    splits_data.append({
                        'Policy': row['Policy_Norm'],
                        'Insured': row.get('Insured', ''),
                        'AAP': row['AAP'],
                        'Product': row.get('Product', ''),
                        'Person1': row.get('Recruiter', ''),
                        'Rate1': 0.55,
                        'Split1': 1.0,
                        'Person2': '',
                        'Rate2': 0.55,
                        'Split2': 0.0,
                        'Person3': '',
                        'Rate3': 0.55,
                        'Split3': 0.0,
                    })
                st.session_state.df_splits = pd.DataFrame(splits_data)
                st.session_state.df_results = None  # 清空计算结果

                st.success(f"✅ 导入成功！{len(df)} 条有效记录")
            except Exception as e:
                st.error(f"❌ 导入失败: {e}")

    # 预览
    if st.session_state.df_raw is not None:
        st.markdown("### 📊 数据预览")
        st.dataframe(
            st.session_state.df_raw[['Policy', 'Insured', 'Recruiter', 'Product', 'Modal', 'AAP']],
            use_container_width=True
        )

# ==================== 第二步：编辑分单 ====================
elif step == "2️⃣ 编辑分单":
    st.header("2️⃣ 批量编辑分单配置")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先在第1步上传数据")
    else:
        st.info("""
        💡 **使用说明**：
        - 每行最多3人分单（Person1/2/3）
        - Rate = 个人佣金比例（如0.55表示55%）
        - Split = 分佣比例（如0.5表示50%）
        - **Split1 + Split2 + Split3 必须 = 1（100%）**
        - 不需要的人员留空，Split填0
        """)

        # 可编辑表格
        edited_df = st.data_editor(
            st.session_state.df_splits,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                'Policy': st.column_config.TextColumn('保单号', disabled=True, width="small"),
                'Insured': st.column_config.TextColumn('被保人', disabled=True, width="medium"),
                'AAP': st.column_config.NumberColumn('AAP', disabled=True, format="$%.0f", width="small"),
                'Product': st.column_config.TextColumn('产品', disabled=True, width="small"),
                'Person1': st.column_config.TextColumn('人员1', width="medium"),
                'Rate1': st.column_config.NumberColumn('比例1', min_value=0, max_value=1, step=0.05, format="%.0f%%", width="small"),
                'Split1': st.column_config.NumberColumn('分佣1', min_value=0, max_value=1, step=0.1, format="%.0f%%", width="small"),
                'Person2': st.column_config.TextColumn('人员2', width="medium"),
                'Rate2': st.column_config.NumberColumn('比例2', min_value=0, max_value=1, step=0.05, format="%.0f%%", width="small"),
                'Split2': st.column_config.NumberColumn('分佣2', min_value=0, max_value=1, step=0.1, format="%.0f%%", width="small"),
                'Person3': st.column_config.TextColumn('人员3', width="medium"),
                'Rate3': st.column_config.NumberColumn('比例3', min_value=0, max_value=1, step=0.05, format="%.0f%%", width="small"),
                'Split3': st.column_config.NumberColumn('分佣3', min_value=0, max_value=1, step=0.1, format="%.0f%%", width="small"),
            },
            hide_index=True
        )

        # 验证分佣比例
        st.markdown("### ✅ 验证分佣比例")
        errors = []
        for idx, row in edited_df.iterrows():
            total = safe_float(row['Split1']) + safe_float(row['Split2']) + safe_float(row['Split3'])
            if abs(total - 1.0) > 0.001:
                errors.append(f"❌ {row['Policy']}: 分佣总和 = {total*100:.0f}% (应为100%)")

        if errors:
            for err in errors[:10]:  # 最多显示10条
                st.error(err)
            if len(errors) > 10:
                st.error(f"... 还有 {len(errors)-10} 条错误")
        else:
            st.success("✅ 所有保单分佣比例正确 (100%)")

        # 保存按钮
        col1, col2 = st.columns([1, 3])
        with col1:
            if st.button("💾 保存分单配置", type="primary"):
                if errors:
                    st.error("❌ 请先修正分佣比例错误")
                else:
                    st.session_state.df_splits = edited_df
                    st.session_state.df_results = None  # 清空旧结果
                    st.success("✅ 配置已保存！请前往第3步计算佣金")

# ==================== 第三步：计算佣金 ====================
elif step == "3️⃣ 计算佣金":
    st.header("3️⃣ 计算佣金")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先完成第1、2步")
    else:
        if st.button("🔄 开始计算", type="primary"):
            with st.spinner("计算中..."):
                results = []
                df_splits = st.session_state.df_splits

                for _, row in df_splits.iterrows():
                    policy = row['Policy']
                    aap = safe_float(row['AAP'])
                    insured = row['Insured']
                    product = row['Product']

                    # 判断佣金比例 (80%/67%/2%)
                    if aap > 10000:  # 大额保单可能是2%
                        comm_rate = 0.02
                    elif 'term' in str(product).lower():
                        comm_rate = 0.67
                    else:
                        comm_rate = 0.80

                    # 总佣金 = AAP × (comm_rate + 48%)
                    total_gross = aap * comm_rate
                    total_override = aap * 0.48
                    total_comm = total_gross + total_override

                    # 计算每人佣金
                    total_distributed = 0
                    for i in [1, 2, 3]:
                        person = row.get(f'Person{i}', '')
                        rate = safe_float(row.get(f'Rate{i}', 0))
                        split = safe_float(row.get(f'Split{i}', 0))

                        if person and split > 0:
                            person_comm = total_comm * rate * split
                            total_distributed += person_comm

                            results.append({
                                'Policy': policy,
                                'Insured': insured,
                                'AAP': aap,
                                'CommRate': comm_rate,
                                'TotalComm': total_comm,
                                'Person': person,
                                'PersonRate': rate,
                                'SplitRatio': split,
                                'PersonComm': person_comm
                            })

                    # 平台剩余
                    platform = total_comm - total_distributed
                    if platform > 0.01:
                        results.append({
                            'Policy': policy,
                            'Insured': insured,
                            'AAP': aap,
                            'CommRate': comm_rate,
                            'TotalComm': total_comm,
                            'Person': '【平台】',
                            'PersonRate': 0,
                            'SplitRatio': 0,
                            'PersonComm': platform
                        })

                st.session_state.df_results = pd.DataFrame(results)
                st.success("✅ 计算完成！")

        # 显示结果
        if st.session_state.df_results is not None:
            df_results = st.session_state.df_results

            # 汇总统计
            st.markdown("### 📊 汇总统计")
            total_aap = df_results.drop_duplicates('Policy')['AAP'].sum()
            total_comm = df_results['PersonComm'].sum()

            col1, col2, col3 = st.columns(3)
            col1.metric("保单数", len(df_results['Policy'].unique()))
            col2.metric("总AAP", format_currency(total_aap))
            col3.metric("总佣金", format_currency(total_comm))

            # 按人员汇总
            st.markdown("### 👥 按人员汇总")
            person_summary = df_results.groupby('Person').agg({
                'Policy': 'count',
                'PersonComm': 'sum'
            }).reset_index()
            person_summary.columns = ['人员', '保单数', '总佣金']
            person_summary = person_summary.sort_values('总佣金', ascending=False)

            st.dataframe(
                person_summary.style.format({'总佣金': '${:,.2f}'}),
                use_container_width=True
            )

            # 明细
            st.markdown("### 📋 佣金明细")
            st.dataframe(
                df_results.style.format({
                    'AAP': '${:,.0f}',
                    'CommRate': '{:.0%}',
                    'TotalComm': '${:,.2f}',
                    'PersonRate': '{:.0%}',
                    'SplitRatio': '{:.0%}',
                    'PersonComm': '${:,.2f}'
                }),
                use_container_width=True
            )

            # 导出
            st.markdown("### 📥 导出报表")
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                person_summary.to_excel(writer, sheet_name='人员汇总', index=False)
                df_results.to_excel(writer, sheet_name='佣金明细', index=False)
                st.session_state.df_splits.to_excel(writer, sheet_name='分单配置', index=False)
            output.seek(0)

            st.download_button(
                "📥 下载Excel报表",
                data=output,
                file_name=f"佣金报表_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ==================== 第四步：对账核验 ====================
elif step == "4️⃣ 对账核验":
    st.header("4️⃣ 对账核验")

    if st.session_state.df_results is None:
        st.warning("⚠️ 请先完成第3步佣金计算")
    else:
        st.info("💡 上传保险公司或平台的对账单，与计算结果进行比对")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### 🏢 保险公司对账单")
            ins_file = st.file_uploader("Gross Commission", type=['xlsx', 'xls'], key='ins')

        with col2:
            st.markdown("#### 📋 平台Override对账单")
            plat_file = st.file_uploader("Override明细", type=['xlsx', 'xls'], key='plat')

        if st.button("🔍 开始对账", type="primary"):
            df_results = st.session_state.df_results

            # 按保单汇总计算结果
            calc_by_policy = df_results.groupby('Policy').agg({
                'AAP': 'first',
                'TotalComm': 'first',
                'PersonComm': 'sum'
            }).reset_index()

            reconcile_data = []

            # 读取对账单
            ins_data = {}
            plat_data = {}

            if ins_file:
                try:
                    df_ins = pd.read_excel(ins_file, skiprows=3)
                    df_ins['Policy_Norm'] = df_ins.iloc[:, 0].apply(normalize_policy)
                    for _, row in df_ins.iterrows():
                        p = row['Policy_Norm']
                        if p:
                            ins_data[p] = safe_float(row.iloc[6]) if len(row) > 6 else 0
                except:
                    st.error("保险公司对账单格式错误")

            if plat_file:
                try:
                    df_plat = pd.read_excel(plat_file, skiprows=2)
                    df_plat['Policy_Norm'] = df_plat.iloc[:, 2].apply(normalize_policy)
                    for _, row in df_plat.iterrows():
                        p = row['Policy_Norm']
                        if p:
                            plat_data[p] = safe_float(row.iloc[5]) if len(row) > 5 else 0
                except:
                    st.error("平台对账单格式错误")

            # 对账
            for _, row in calc_by_policy.iterrows():
                policy = row['Policy']
                calc_comm = row['TotalComm']

                actual_ins = ins_data.get(policy, 0)
                actual_plat = plat_data.get(policy, 0)
                actual_total = actual_ins + actual_plat

                diff = actual_total - calc_comm
                status = '✅' if abs(diff) < 1 else '❌'

                reconcile_data.append({
                    '保单号': policy,
                    '计算佣金': calc_comm,
                    '保险公司': actual_ins,
                    '平台Override': actual_plat,
                    '实际合计': actual_total,
                    '差额': diff,
                    '状态': status
                })

            df_reconcile = pd.DataFrame(reconcile_data)

            # 统计
            st.markdown("### 📊 对账结果")
            ok_count = (df_reconcile['状态'] == '✅').sum()
            total_count = len(df_reconcile)

            col1, col2, col3 = st.columns(3)
            col1.metric("总保单", total_count)
            col2.metric("匹配", ok_count)
            col3.metric("差异", total_count - ok_count)

            # 差异记录
            df_diff = df_reconcile[df_reconcile['状态'] == '❌']
            if len(df_diff) > 0:
                st.markdown("### ❌ 差异记录")
                st.dataframe(
                    df_diff.style.format({
                        '计算佣金': '${:,.2f}',
                        '保险公司': '${:,.2f}',
                        '平台Override': '${:,.2f}',
                        '实际合计': '${:,.2f}',
                        '差额': '${:,.2f}'
                    }),
                    use_container_width=True
                )
            else:
                st.success("✅ 全部匹配！")

            # 完整对账表
            st.markdown("### 📋 完整对账表")
            st.dataframe(
                df_reconcile.style.format({
                    '计算佣金': '${:,.2f}',
                    '保险公司': '${:,.2f}',
                    '平台Override': '${:,.2f}',
                    '实际合计': '${:,.2f}',
                    '差额': '${:,.2f}'
                }),
                use_container_width=True
            )
