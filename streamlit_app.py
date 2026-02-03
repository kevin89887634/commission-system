"""
💰 佣金管理系统 v2.0
- 数据清洗：过滤无效行
- 批量分单配置
- 对账功能
- 月度佣金计算
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
    """安全转换为浮点数"""
    try:
        if value is None or pd.isna(value):
            return default
        return float(value)
    except (ValueError, TypeError):
        return default

def format_currency(amount):
    """格式化货币"""
    if amount is None or pd.isna(amount):
        return "$0.00"
    return f"${amount:,.2f}"

def format_percent(rate):
    """格式化百分比"""
    if rate is None or pd.isna(rate):
        return "0%"
    return f"{rate*100:.0f}%"

def is_valid_policy(policy):
    """检查是否为有效保单号"""
    if policy is None or pd.isna(policy):
        return False
    s = str(policy).strip()
    # 排除无效值
    invalid_patterns = [
        'policy #', 'policy#', 'nan', 'none', '',
        '* for ul', 'exported on', 'exported by',
        'for ul life'
    ]
    s_lower = s.lower()
    for pattern in invalid_patterns:
        if pattern in s_lower:
            return False
    # 必须包含数字
    if not any(c.isdigit() for c in s):
        return False
    return True

def is_valid_recruiter(recruiter):
    """检查是否为有效经纪人名称"""
    if recruiter is None or pd.isna(recruiter):
        return False
    s = str(recruiter).strip()
    invalid_patterns = ['agent', 'none', 'nan', '', 'exported', '*']
    s_lower = s.lower()
    for pattern in invalid_patterns:
        if s_lower == pattern or s_lower.startswith(pattern):
            return False
    return True

# ==================== 数据导入 ====================
def import_new_business(file):
    """导入 New Business Report，自动清洗数据"""
    df = pd.read_excel(file, skiprows=4)
    expected_cols = [
        'Policy', 'Insured', 'Recruiter', 'Status', 'Delivery',
        'Action', 'SubmitDate', 'Modal', 'Product', 'Sent',
        'Owner', 'SubmitMethod', 'CaseManager', 'AAP',
        'AgentNum', 'Agency', 'CompanyCode', 'Bookmark'
    ]
    df.columns = expected_cols[:len(df.columns)]

    # 数据清洗：过滤无效行
    df = df[df['Policy'].apply(is_valid_policy)]
    df = df[df['Recruiter'].apply(is_valid_recruiter)]

    # 标准化Policy号
    df['Policy_Norm'] = df['Policy'].apply(normalize_policy)

    # 转换数值列
    for col in ['Modal', 'AAP']:
        if col in df.columns:
            df[col] = df[col].apply(safe_float)

    # 过滤 AAP = 0 的行
    df = df[df['AAP'] > 0]

    # 重置索引
    df = df.reset_index(drop=True)

    return df

def import_zhubiao(file):
    """导入分单配置表 (zhubiao格式)"""
    df = pd.read_excel(file, sheet_name=0)
    # 标准化列名
    col_mapping = {
        'Policy #': 'Policy',
        '被保人': 'Insured',
        'Process Date': 'ProcessDate',
        'Premium Amt': 'Premium',
        'Comm Rate %': 'CommRate',
        'Gross Comm Earned': 'GrossComm',
        'Payment Date': 'PaymentDate',
        'Recruiter': 'Recruiter',
        'Recruiter佣金比例': 'RecruiterRate',
        'Recruiter分佣比例': 'RecruiterSplit',
        'Recruiter佣金': 'RecruiterComm',
        'CFT': 'CFT',
        'CFT比例': 'CFTRate',
        'CFT分佣比例': 'CFTSplit',
        'CFT佣金': 'CFTComm'
    }
    df = df.rename(columns=col_mapping)
    df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
    return df

def import_override(file):
    """导入 Override by Policy"""
    df = pd.read_excel(file, skiprows=2)
    if len(df.columns) >= 6:
        df = df.iloc[:, [2, 5]]
        df.columns = ['Policy', 'Override']
    df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
    df['Override'] = df['Override'].apply(safe_float)
    df_sum = df.groupby('Policy_Norm')['Override'].sum().reset_index()
    return df_sum

# ==================== 分单配置管理 ====================
def get_default_split_config(df):
    """从数据生成默认分单配置"""
    config = {}
    for _, row in df.iterrows():
        policy = row['Policy_Norm']
        recruiter = row.get('Recruiter', 'Unknown')
        if policy not in config:
            config[policy] = {
                'insured': row.get('Insured', ''),
                'aap': safe_float(row.get('AAP', 0)),
                'modal': safe_float(row.get('Modal', 0)),
                'product': row.get('Product', ''),
                'splits': [
                    {'name': recruiter, 'rate': 0.55, 'split': 1.0, 'role': 'Recruiter'}
                ]
            }
    return config

def validate_split_config(config):
    """验证分单配置是否正确（总比例=100%）"""
    errors = []
    for policy, data in config.items():
        total_split = sum(s['split'] for s in data['splits'])
        if abs(total_split - 1.0) > 0.001:
            errors.append(f"保单 {policy}: 分佣比例总和为 {total_split*100:.0f}%，应为 100%")
    return errors

# ==================== 佣金计算 ====================
def calculate_commissions(df, split_config, override_data=None):
    """
    计算所有佣金
    公式: 个人佣金 = Gross Comm × 个人佣金比例 × 分佣比例
    Gross Comm = Premium × Comm Rate (80%/67%/2%)
    """
    results = []

    # 合并Override数据
    override_map = {}
    if override_data is not None:
        override_map = dict(zip(override_data['Policy_Norm'], override_data['Override']))

    for _, row in df.iterrows():
        policy = row['Policy_Norm']
        aap = safe_float(row.get('AAP', 0))
        modal = safe_float(row.get('Modal', 0))
        insured = row.get('Insured', '')
        product = row.get('Product', '')

        # 判断佣金比例
        if '2' in str(row.get('Product', '')).lower() or modal > 0 and aap/modal < 2:
            comm_rate = 0.02  # 特殊产品 2%
        elif 'term' in str(row.get('Product', '')).lower():
            comm_rate = 0.67  # Term产品 67%
        else:
            comm_rate = 0.80  # 标准产品 80%

        # 判断缴费类型
        if modal > 0:
            ratio = aap / modal
            pay_type = '月缴' if ratio > 6 else '年缴'
        else:
            pay_type = '未知'

        # 获取Override
        override = override_map.get(policy, 0)

        # Gross Commission (基于AAP计算总佣金潜力)
        total_gross = aap * comm_rate
        total_override = aap * 0.48
        total_comm = total_gross + total_override  # = AAP × 128%

        # 获取分单配置
        if policy in split_config:
            splits = split_config[policy]['splits']
        else:
            recruiter = row.get('Recruiter', 'Unknown')
            splits = [{'name': recruiter, 'rate': 0.55, 'split': 1.0, 'role': 'Recruiter'}]

        # 计算每人佣金
        total_distributed = 0
        for split in splits:
            name = split['name']
            rate = safe_float(split.get('rate', 0.55))
            split_ratio = safe_float(split.get('split', 1.0))
            role = split.get('role', 'Recruiter')

            # 个人佣金 = 总佣金 × 个人比例 × 分佣比例
            person_comm = total_comm * rate * split_ratio
            total_distributed += person_comm

            results.append({
                'Policy': policy,
                'Policy_Orig': row.get('Policy', ''),
                'Insured': insured,
                'Product': product,
                'AAP': aap,
                'Modal': modal,
                'PayType': pay_type,
                'CommRate': comm_rate,
                'TotalGross': total_gross,
                'TotalOverride': total_override,
                'TotalComm': total_comm,
                'Override_Received': override,
                'Person': name,
                'Role': role,
                'PersonRate': rate,
                'SplitRatio': split_ratio,
                'PersonComm': person_comm
            })

        # 平台剩余
        platform_comm = total_comm - total_distributed
        if platform_comm > 0.01:
            results.append({
                'Policy': policy,
                'Policy_Orig': row.get('Policy', ''),
                'Insured': insured,
                'Product': product,
                'AAP': aap,
                'Modal': modal,
                'PayType': pay_type,
                'CommRate': comm_rate,
                'TotalGross': total_gross,
                'TotalOverride': total_override,
                'TotalComm': total_comm,
                'Override_Received': override,
                'Person': '【平台】',
                'Role': 'Platform',
                'PersonRate': 0,
                'SplitRatio': 0,
                'PersonComm': platform_comm
            })

    return pd.DataFrame(results)

def generate_person_summary(df_results):
    """生成按人员汇总"""
    summary = df_results.groupby('Person').agg({
        'Policy': 'count',
        'AAP': 'sum',
        'PersonComm': 'sum'
    }).reset_index()
    summary.columns = ['人员', '保单数', '总AAP', '总佣金']
    summary = summary.sort_values('总佣金', ascending=False)
    return summary

# ==================== 对账功能 ====================
def reconcile_statements(df_results, insurance_statement=None, platform_statement=None):
    """
    对账功能：比对保险公司和平台对账单
    """
    reconciliation = []

    # 按保单汇总计算结果
    calc_summary = df_results.groupby('Policy').agg({
        'TotalGross': 'first',
        'TotalOverride': 'first',
        'TotalComm': 'first'
    }).reset_index()

    for _, row in calc_summary.iterrows():
        policy = row['Policy']
        calc_gross = row['TotalGross']
        calc_override = row['TotalOverride']

        # 匹配保险公司对账单
        ins_gross = 0
        if insurance_statement is not None:
            match = insurance_statement[insurance_statement['Policy_Norm'] == policy]
            if len(match) > 0:
                ins_gross = safe_float(match.iloc[0].get('Commission', 0))

        # 匹配平台对账单
        plat_override = 0
        if platform_statement is not None:
            match = platform_statement[platform_statement['Policy_Norm'] == policy]
            if len(match) > 0:
                plat_override = safe_float(match.iloc[0].get('Override', 0))

        # 比对
        gross_diff = ins_gross - calc_gross
        override_diff = plat_override - calc_override

        gross_status = '✅' if abs(gross_diff) < 0.01 else '❌'
        override_status = '✅' if abs(override_diff) < 0.01 else '❌'

        reconciliation.append({
            'Policy': policy,
            'Calc_Gross': calc_gross,
            'Actual_Gross': ins_gross,
            'Gross_Diff': gross_diff,
            'Gross_Status': gross_status,
            'Calc_Override': calc_override,
            'Actual_Override': plat_override,
            'Override_Diff': override_diff,
            'Override_Status': override_status
        })

    return pd.DataFrame(reconciliation)

# ==================== 报表导出 ====================
def export_to_excel(data_dict):
    """导出Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in data_dict.items():
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    output.seek(0)
    return output

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="佣金管理系统 v2.0",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化 session state
if 'df_business' not in st.session_state:
    st.session_state.df_business = None
if 'df_override' not in st.session_state:
    st.session_state.df_override = None
if 'split_config' not in st.session_state:
    st.session_state.split_config = {}
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = None

# 侧边栏
with st.sidebar:
    st.title("💰 佣金管理系统")
    st.caption("v2.0 专业版")
    st.markdown("---")
    page = st.radio(
        "功能菜单",
        ["📁 数据导入", "👥 分单配置", "💵 佣金计算", "📊 对账核验", "📈 报表导出"],
        label_visibility="collapsed"
    )
    st.markdown("---")

    # 显示数据状态
    if st.session_state.df_business is not None:
        st.success(f"✅ 已导入 {len(st.session_state.df_business)} 条保单")
    if st.session_state.split_config:
        st.info(f"📋 已配置 {len(st.session_state.split_config)} 条分单")

# 主内容
st.title(page)

# ==================== 数据导入 ====================
if page == "📁 数据导入":
    st.markdown("### 上传数据文件")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📋 New Business Report")
        new_biz_file = st.file_uploader("从NLG下载的新业务报表", type=['xlsx', 'xls'], key='new_biz')

    with col2:
        st.markdown("#### 💰 Override by Policy (可选)")
        override_file = st.file_uploader("Override佣金明细", type=['xlsx', 'xls'], key='override')

    st.markdown("---")
    st.markdown("#### 📑 分单配置表 (可选)")
    zhubiao_file = st.file_uploader("上传zhubiao格式的分单配置", type=['xlsx', 'xls'], key='zhubiao')

    if st.button("🔄 导入数据", type="primary"):
        if new_biz_file:
            with st.spinner("正在导入并清洗数据..."):
                try:
                    st.session_state.df_business = import_new_business(new_biz_file)
                    st.success(f"✅ New Business 导入成功: {len(st.session_state.df_business)} 条有效记录")

                    # 生成默认分单配置
                    st.session_state.split_config = get_default_split_config(st.session_state.df_business)

                    if override_file:
                        st.session_state.df_override = import_override(override_file)
                        st.success("✅ Override 导入成功")

                    if zhubiao_file:
                        df_zhubiao = import_zhubiao(zhubiao_file)
                        # 从zhubiao更新分单配置
                        for _, row in df_zhubiao.iterrows():
                            policy = row['Policy_Norm']
                            if policy in st.session_state.split_config:
                                splits = []
                                # Recruiter
                                if pd.notna(row.get('Recruiter')) and row.get('Recruiter') != '-':
                                    splits.append({
                                        'name': row['Recruiter'],
                                        'rate': safe_float(row.get('RecruiterRate', 0.55)),
                                        'split': safe_float(row.get('RecruiterSplit', 1.0)),
                                        'role': 'Recruiter'
                                    })
                                # CFT
                                if pd.notna(row.get('CFT')) and row.get('CFT') != '-':
                                    splits.append({
                                        'name': row['CFT'],
                                        'rate': safe_float(row.get('CFTRate', 0.55)),
                                        'split': safe_float(row.get('CFTSplit', 0)),
                                        'role': 'CFT'
                                    })
                                if splits:
                                    st.session_state.split_config[policy]['splits'] = splits
                        st.success("✅ 分单配置已从zhubiao导入")

                except Exception as e:
                    st.error(f"❌ 导入失败: {str(e)}")
        else:
            st.warning("⚠️ 请上传 New Business Report")

    # 数据预览
    if st.session_state.df_business is not None:
        st.markdown("---")
        st.markdown("### 📊 数据预览 (已清洗)")
        display_cols = ['Policy', 'Insured', 'Recruiter', 'Product', 'Modal', 'AAP']
        available_cols = [c for c in display_cols if c in st.session_state.df_business.columns]
        st.dataframe(st.session_state.df_business[available_cols], use_container_width=True)

# ==================== 分单配置 ====================
elif page == "👥 分单配置":
    st.markdown("### 批量分单配置")
    st.info("💡 设置每个保单的佣金分配比例，确保总分佣比例 = 100%")

    if st.session_state.df_business is None:
        st.warning("⚠️ 请先在「数据导入」页面上传数据")
    else:
        # 显示所有保单的分单配置
        config = st.session_state.split_config

        # 验证配置
        errors = validate_split_config(config)
        if errors:
            st.error("❌ 配置错误:")
            for err in errors:
                st.write(f"  • {err}")

        # 批量编辑表格
        st.markdown("#### 📋 分单配置表")

        # 构建编辑数据
        edit_data = []
        for policy, data in config.items():
            for i, split in enumerate(data['splits']):
                edit_data.append({
                    '保单号': policy,
                    '被保人': data.get('insured', ''),
                    'AAP': data.get('aap', 0),
                    '序号': i + 1,
                    '人员': split['name'],
                    '角色': split['role'],
                    '佣金比例': split['rate'],
                    '分佣比例': split['split']
                })

        df_edit = pd.DataFrame(edit_data)

        # 使用data_editor进行批量编辑
        edited_df = st.data_editor(
            df_edit,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                '保单号': st.column_config.TextColumn('保单号', disabled=True),
                '被保人': st.column_config.TextColumn('被保人', disabled=True),
                'AAP': st.column_config.NumberColumn('AAP', format="$%.2f", disabled=True),
                '序号': st.column_config.NumberColumn('序号', disabled=True),
                '人员': st.column_config.TextColumn('人员'),
                '角色': st.column_config.SelectboxColumn('角色', options=['Recruiter', 'CFT', 'Other']),
                '佣金比例': st.column_config.NumberColumn('佣金比例', min_value=0, max_value=1, step=0.05, format="%.0f%%"),
                '分佣比例': st.column_config.NumberColumn('分佣比例', min_value=0, max_value=1, step=0.1, format="%.0f%%")
            }
        )

        if st.button("💾 保存配置", type="primary"):
            # 从编辑后的数据更新配置
            new_config = {}
            for _, row in edited_df.iterrows():
                policy = row['保单号']
                if policy not in new_config:
                    new_config[policy] = {
                        'insured': row['被保人'],
                        'aap': row['AAP'],
                        'splits': []
                    }
                new_config[policy]['splits'].append({
                    'name': row['人员'],
                    'role': row['角色'],
                    'rate': row['佣金比例'],
                    'split': row['分佣比例']
                })

            # 验证
            errors = validate_split_config(new_config)
            if errors:
                st.error("❌ 保存失败，请修正以下错误:")
                for err in errors:
                    st.write(f"  • {err}")
            else:
                st.session_state.split_config = new_config
                st.success("✅ 配置已保存")

        # 显示汇总
        st.markdown("---")
        st.markdown("#### 📊 分佣比例汇总")
        summary_data = []
        for policy, data in config.items():
            total_split = sum(s['split'] for s in data['splits'])
            status = '✅' if abs(total_split - 1.0) < 0.001 else '❌'
            summary_data.append({
                '保单号': policy,
                '参与人数': len(data['splits']),
                '总分佣比例': f"{total_split*100:.0f}%",
                '状态': status
            })
        st.dataframe(pd.DataFrame(summary_data), use_container_width=True)

# ==================== 佣金计算 ====================
elif page == "💵 佣金计算":
    st.markdown("### 佣金计算")

    if st.session_state.df_business is None:
        st.warning("⚠️ 请先在「数据导入」页面上传数据")
    else:
        # 验证分单配置
        errors = validate_split_config(st.session_state.split_config)
        if errors:
            st.warning("⚠️ 分单配置有误，请先在「分单配置」页面修正")
            for err in errors:
                st.write(f"  • {err}")

        if st.button("🔄 计算佣金", type="primary"):
            with st.spinner("计算中..."):
                results = calculate_commissions(
                    st.session_state.df_business,
                    st.session_state.split_config,
                    st.session_state.df_override
                )
                st.session_state.calc_results = results
                st.success("✅ 计算完成")

        if st.session_state.calc_results is not None:
            results = st.session_state.calc_results

            # 汇总统计
            st.markdown("---")
            st.markdown("### 📊 汇总统计")

            total_aap = results.drop_duplicates('Policy')['AAP'].sum()
            total_comm = results['PersonComm'].sum()

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("保单总数", len(results['Policy'].unique()))
            with col2:
                st.metric("总AAP", format_currency(total_aap))
            with col3:
                st.metric("总佣金 (128%)", format_currency(total_aap * 1.28))
            with col4:
                st.metric("已分配佣金", format_currency(total_comm))

            # 按人员汇总
            st.markdown("---")
            st.markdown("### 👥 按人员汇总")
            person_summary = generate_person_summary(results)
            st.dataframe(
                person_summary.style.format({
                    '总AAP': '${:,.2f}',
                    '总佣金': '${:,.2f}'
                }),
                use_container_width=True
            )

            # 保单明细
            st.markdown("---")
            st.markdown("### 📋 保单佣金明细")
            display_cols = ['Policy', 'Insured', 'AAP', 'PayType', 'Person', 'Role',
                          'PersonRate', 'SplitRatio', 'PersonComm']
            st.dataframe(
                results[display_cols].style.format({
                    'AAP': '${:,.2f}',
                    'PersonRate': '{:.0%}',
                    'SplitRatio': '{:.0%}',
                    'PersonComm': '${:,.2f}'
                }),
                use_container_width=True
            )

# ==================== 对账核验 ====================
elif page == "📊 对账核验":
    st.markdown("### 对账核验")
    st.info("💡 上传保险公司和平台对账单，自动比对差异")

    if st.session_state.calc_results is None:
        st.warning("⚠️ 请先在「佣金计算」页面完成计算")
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### 🏢 保险公司对账单")
            ins_file = st.file_uploader("Gross Commission对账单", type=['xlsx', 'xls'], key='ins')

        with col2:
            st.markdown("#### 📋 平台Override对账单")
            plat_file = st.file_uploader("Override对账单", type=['xlsx', 'xls'], key='plat')

        if st.button("🔍 开始对账", type="primary"):
            insurance_df = None
            platform_df = None

            if ins_file:
                try:
                    insurance_df = pd.read_excel(ins_file, skiprows=3)
                    insurance_df['Policy_Norm'] = insurance_df.iloc[:, 0].apply(normalize_policy)
                    insurance_df['Commission'] = insurance_df.iloc[:, 6].apply(safe_float) if len(insurance_df.columns) > 6 else 0
                except:
                    st.error("保险公司对账单格式错误")

            if plat_file:
                try:
                    platform_df = import_override(plat_file)
                except:
                    st.error("平台对账单格式错误")

            reconciliation = reconcile_statements(
                st.session_state.calc_results,
                insurance_df,
                platform_df
            )

            st.markdown("---")
            st.markdown("### 📊 对账结果")

            # 统计
            if len(reconciliation) > 0:
                gross_ok = (reconciliation['Gross_Status'] == '✅').sum()
                override_ok = (reconciliation['Override_Status'] == '✅').sum()
                total = len(reconciliation)

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("总保单数", total)
                with col2:
                    st.metric("Gross匹配", f"{gross_ok}/{total}")
                with col3:
                    st.metric("Override匹配", f"{override_ok}/{total}")

                # 显示不匹配的记录
                st.markdown("#### ❌ 差异记录")
                diff_records = reconciliation[
                    (reconciliation['Gross_Status'] == '❌') |
                    (reconciliation['Override_Status'] == '❌')
                ]
                if len(diff_records) > 0:
                    st.dataframe(
                        diff_records.style.format({
                            'Calc_Gross': '${:,.2f}',
                            'Actual_Gross': '${:,.2f}',
                            'Gross_Diff': '${:,.2f}',
                            'Calc_Override': '${:,.2f}',
                            'Actual_Override': '${:,.2f}',
                            'Override_Diff': '${:,.2f}'
                        }),
                        use_container_width=True
                    )
                else:
                    st.success("✅ 全部匹配，无差异！")

                # 完整对账表
                st.markdown("#### 📋 完整对账表")
                st.dataframe(reconciliation, use_container_width=True)

# ==================== 报表导出 ====================
elif page == "📈 报表导出":
    st.markdown("### 报表导出")

    if st.session_state.calc_results is None:
        st.warning("⚠️ 请先在「佣金计算」页面完成计算")
    else:
        results = st.session_state.calc_results
        person_summary = generate_person_summary(results)

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### 📊 完整报表")
            excel_data = export_to_excel({
                '人员汇总': person_summary,
                '保单明细': results,
                '分单配置': pd.DataFrame([
                    {'保单号': p, '人员': s['name'], '比例': s['rate'], '分佣': s['split']}
                    for p, d in st.session_state.split_config.items()
                    for s in d['splits']
                ])
            })
            st.download_button(
                label="📥 下载完整报表",
                data=excel_data,
                file_name=f"佣金报表_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            st.markdown("#### 👤 个人报表")
            persons = [p for p in person_summary['人员'].tolist() if p != '【平台】']
            selected = st.selectbox("选择人员", persons)

            if selected:
                person_data = results[results['Person'] == selected]
                person_total = person_data['PersonComm'].sum()

                st.metric(f"{selected} 总佣金", format_currency(person_total))

                excel_data = export_to_excel({
                    '个人明细': person_data
                })
                st.download_button(
                    label="📥 下载个人报表",
                    data=excel_data,
                    file_name=f"佣金报表_{selected}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
