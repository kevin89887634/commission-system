"""
💰 佣金管理系统 - Streamlit Cloud 版本
单文件版本，适合云端部署
"""
import streamlit as st
import pandas as pd
import re
from io import BytesIO

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
    """安全转换为浮点数"""
    try:
        if value is None:
            return default
        return float(value)
    except (ValueError, TypeError):
        return default

def format_currency(amount):
    """格式化货币"""
    if amount is None:
        return "$0.00"
    return f"${amount:,.2f}"

# ==================== 数据导入 ====================
def import_new_business(file):
    """导入 New Business Report"""
    df = pd.read_excel(file, skiprows=4)
    expected_cols = [
        'Policy', 'Insured', 'Recruiter', 'Status', 'Delivery',
        'Action', 'SubmitDate', 'Modal', 'Product', 'Sent',
        'Owner', 'SubmitMethod', 'CaseManager', 'AAP',
        'AgentNum', 'Agency', 'CompanyCode', 'Bookmark'
    ]
    df.columns = expected_cols[:len(df.columns)]
    df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
    for col in ['Modal', 'AAP']:
        if col in df.columns:
            df[col] = df[col].apply(safe_float)
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

def get_merged_data(df_business, df_override=None):
    """合并数据"""
    df = df_business.copy()

    if df_override is not None:
        override_map = dict(zip(df_override['Policy_Norm'], df_override['Override']))
        df['Override'] = df['Policy_Norm'].map(override_map).fillna(0)
    else:
        df['Override'] = 0

    df['Expected_Override'] = df['AAP'] * 0.48

    def get_pay_type(row):
        if row['Modal'] > 0:
            return '月缴' if row['AAP'] / row['Modal'] > 6 else '年缴'
        return '未知'
    df['PayType'] = df.apply(get_pay_type, axis=1)

    def get_paid_months(row):
        if row['AAP'] > 0 and row['PayType'] == '月缴':
            return round(row['Override'] / (row['AAP'] * 0.04))
        return 1 if row['Override'] > 0 else 0
    df['PaidMonths'] = df.apply(get_paid_months, axis=1)

    return df

# ==================== 佣金计算 ====================
# 经纪人佣金比例配置
RECRUITER_RATES = {
    "Cindy Li": 0.60,
    "David Wang": 0.55,
    "Thomas Chen": 0.50,
    "唐慧": 0.02,
}
DEFAULT_RATE = 0.50

def get_recruiter_rate(name):
    """获取经纪人佣金比例"""
    return RECRUITER_RATES.get(name, DEFAULT_RATE)

def calculate_all(df, split_config=None):
    """计算所有佣金"""
    if split_config is None:
        split_config = {}

    details = []

    for _, row in df.iterrows():
        policy = row.get('Policy_Norm', '')
        tp = safe_float(row.get('AAP', 0))
        modal = safe_float(row.get('Modal', 0))
        recruiter = row.get('Recruiter', '')
        override_received = safe_float(row.get('Override', 0))
        paid_months = int(row.get('PaidMonths', 0))
        pay_type = row.get('PayType', '年缴')

        total_potential = tp * 1.28

        if pay_type == '月缴':
            monthly_commission = (tp / 12) * 1.28
            received_total = monthly_commission * paid_months
            remaining_months = max(0, 12 - paid_months)
            expected_remaining = monthly_commission * remaining_months
        else:
            received_total = override_received + (modal * 0.8)
            remaining_months = 0
            expected_remaining = 0

        splits = split_config.get(policy, [{'name': recruiter, 'ratio': 1.0}])

        for split in splits:
            name = split['name']
            split_ratio = split['ratio']
            rate = get_recruiter_rate(name)

            person_received = received_total * rate * split_ratio
            person_expected = expected_remaining * rate * split_ratio

            details.append({
                'Policy': policy,
                'TP': tp,
                'PayType': pay_type,
                'TotalPotential': total_potential,
                'Recruiter': name,
                'Rate': rate,
                'SplitRatio': split_ratio,
                'Received': person_received,
                'Expected': person_expected,
                'Total': person_received + person_expected
            })

    df_details = pd.DataFrame(details)

    if len(df_details) > 0:
        df_by_recruiter = df_details.groupby('Recruiter').agg({
            'Policy': 'count',
            'TP': 'sum',
            'Received': 'sum',
            'Expected': 'sum',
            'Total': 'sum'
        }).reset_index()
        df_by_recruiter.columns = ['Recruiter', 'PolicyCount', 'TotalTP',
                                   'ReceivedComm', 'ExpectedComm', 'TotalComm']
    else:
        df_by_recruiter = pd.DataFrame()

    summary = {
        'total_policies': len(df),
        'total_tp': df['AAP'].sum(),
        'total_potential': df['AAP'].sum() * 1.28,
        'total_received': df_details['Received'].sum() if len(df_details) > 0 else 0,
        'total_expected': df_details['Expected'].sum() if len(df_details) > 0 else 0
    }

    return {
        'summary': summary,
        'details': df_details,
        'by_recruiter': df_by_recruiter
    }

# ==================== 报表导出 ====================
def export_to_excel(data_dict):
    """导出Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in data_dict.items():
            if df is not None and len(df) > 0:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)
    return output

# ==================== 页面配置 ====================
st.set_page_config(
    page_title="佣金管理系统",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 初始化 session state
if 'df_business' not in st.session_state:
    st.session_state.df_business = None
if 'df_override' not in st.session_state:
    st.session_state.df_override = None
if 'df_merged' not in st.session_state:
    st.session_state.df_merged = None
if 'calc_results' not in st.session_state:
    st.session_state.calc_results = None
if 'split_config' not in st.session_state:
    st.session_state.split_config = {}

# 侧边栏
with st.sidebar:
    st.title("💰 佣金管理系统")
    st.markdown("---")
    page = st.radio(
        "功能菜单",
        ["📁 数据导入", "💵 佣金计算", "👥 分单配置", "✅ 收款核对", "📈 预期收益", "📊 报表导出"],
        label_visibility="collapsed"
    )
    st.markdown("---")
    st.caption("v1.0 Cloud | NLG佣金系统")

# 主内容
st.title(page)

# ==================== 数据导入 ====================
if page == "📁 数据导入":
    st.markdown("### 上传 NLG 数据文件")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📋 必需文件")
        new_biz_file = st.file_uploader("New Business Report", type=['xlsx', 'xls'])

    with col2:
        st.markdown("#### 💰 可选文件")
        override_file = st.file_uploader("Override by Policy", type=['xlsx', 'xls'])

    if st.button("🔄 导入数据", type="primary"):
        if new_biz_file:
            with st.spinner("正在导入..."):
                try:
                    st.session_state.df_business = import_new_business(new_biz_file)
                    st.success(f"✅ New Business 导入成功: {len(st.session_state.df_business)} 条记录")

                    if override_file:
                        st.session_state.df_override = import_override(override_file)
                        st.success("✅ Override 导入成功")

                    st.session_state.df_merged = get_merged_data(
                        st.session_state.df_business,
                        st.session_state.df_override
                    )
                except Exception as e:
                    st.error(f"❌ 导入失败: {str(e)}")
        else:
            st.warning("⚠️ 请上传 New Business Report")

    if st.session_state.df_merged is not None:
        st.markdown("---")
        st.markdown("### 数据预览")
        display_cols = ['Policy', 'Recruiter', 'Product', 'Modal', 'AAP', 'PayType', 'Override', 'PaidMonths']
        available_cols = [c for c in display_cols if c in st.session_state.df_merged.columns]
        st.dataframe(st.session_state.df_merged[available_cols].head(20), use_container_width=True)
        st.caption(f"共 {len(st.session_state.df_merged)} 条记录")

# ==================== 佣金计算 ====================
elif page == "💵 佣金计算":
    if st.session_state.df_merged is None:
        st.warning("⚠️ 请先在「数据导入」页面上传数据")
    else:
        if st.button("🔄 开始计算", type="primary"):
            with st.spinner("计算中..."):
                results = calculate_all(st.session_state.df_merged, st.session_state.split_config)
                st.session_state.calc_results = results
                st.success("✅ 计算完成")

        if st.session_state.calc_results:
            results = st.session_state.calc_results

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("保单总数", results['summary']['total_policies'])
            with col2:
                st.metric("总TP", format_currency(results['summary']['total_tp']))
            with col3:
                st.metric("已收佣金", format_currency(results['summary']['total_received']))
            with col4:
                st.metric("预期佣金", format_currency(results['summary']['total_expected']))

            st.markdown("---")
            st.markdown("### 按经纪人汇总")
            if len(results['by_recruiter']) > 0:
                st.dataframe(results['by_recruiter'], use_container_width=True)

            st.markdown("### 保单明细")
            if len(results['details']) > 0:
                st.dataframe(results['details'], use_container_width=True)

# ==================== 分单配置 ====================
elif page == "👥 分单配置":
    st.markdown("### 保单分单比例配置")
    st.info("💡 设置每个保单的佣金分配比例，支持多人分单")

    if st.session_state.df_merged is None:
        st.warning("⚠️ 请先在「数据导入」页面上传数据")
    else:
        policies = st.session_state.df_merged['Policy_Norm'].unique().tolist()
        selected_policy = st.selectbox("选择保单", policies)

        if selected_policy:
            policy_info = st.session_state.df_merged[
                st.session_state.df_merged['Policy_Norm'] == selected_policy
            ].iloc[0]

            col1, col2, col3 = st.columns(3)
            with col1:
                st.write(f"**TP:** {format_currency(policy_info['AAP'])}")
            with col2:
                st.write(f"**原Recruiter:** {policy_info['Recruiter']}")
            with col3:
                st.write(f"**类型:** {policy_info['PayType']}")

            st.markdown("---")
            st.markdown("#### 分单人员配置")

            current_splits = st.session_state.split_config.get(
                selected_policy,
                [{'name': policy_info['Recruiter'], 'ratio': 1.0}]
            )

            num_splits = st.number_input("分单人数", min_value=1, max_value=5, value=len(current_splits))

            new_splits = []
            for i in range(int(num_splits)):
                col1, col2 = st.columns([2, 1])
                with col1:
                    default_name = current_splits[i]['name'] if i < len(current_splits) else ""
                    name = st.text_input(f"经纪人 {i+1}", value=default_name, key=f"name_{i}")
                with col2:
                    default_ratio = current_splits[i]['ratio'] if i < len(current_splits) else 0.0
                    ratio = st.number_input(f"比例 {i+1}", min_value=0.0, max_value=1.0,
                                           value=float(default_ratio), step=0.1, key=f"ratio_{i}")
                if name:
                    new_splits.append({'name': name, 'ratio': ratio})

            total_ratio = sum(s['ratio'] for s in new_splits)
            if abs(total_ratio - 1.0) > 0.01:
                st.warning(f"⚠️ 分单比例总和为 {total_ratio:.0%}，应为 100%")

            if st.button("💾 保存配置"):
                st.session_state.split_config[selected_policy] = new_splits
                st.success("✅ 配置已保存")

# ==================== 收款核对 ====================
elif page == "✅ 收款核对":
    st.markdown("### 收款金额核对")

    if st.session_state.calc_results is None:
        st.warning("⚠️ 请先在「佣金计算」页面完成计算")
    else:
        results = st.session_state.calc_results

        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 计算结果")
            st.metric("计算总额", format_currency(results['summary']['total_received']))
        with col2:
            st.markdown("#### 实际收款")
            actual_amount = st.number_input("输入实际收款金额", min_value=0.0, step=100.0)

        if st.button("🔍 核对"):
            calculated = results['summary']['total_received']
            difference = actual_amount - calculated

            if abs(difference) < 0.01:
                st.success(f"✅ 金额匹配！")
            else:
                st.error(f"❌ 金额不匹配！差额: {format_currency(difference)}")

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("计算金额", format_currency(calculated))
            with col2:
                st.metric("实际金额", format_currency(actual_amount))
            with col3:
                st.metric("差额", format_currency(difference))

# ==================== 预期收益 ====================
elif page == "📈 预期收益":
    st.markdown("### 未来预期收益")

    if st.session_state.calc_results is None:
        st.warning("⚠️ 请先在「佣金计算」页面完成计算")
    else:
        results = st.session_state.calc_results

        total_expected = results['summary']['total_expected']
        total_potential = results['summary']['total_potential']
        total_received = results['summary']['total_received']

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("已收佣金", format_currency(total_received))
        with col2:
            st.metric("预期佣金", format_currency(total_expected))
        with col3:
            st.metric("总佣金潜力", format_currency(total_potential))

        if len(results['details']) > 0:
            st.markdown("### 按保单预期")
            monthly_policies = results['details'][results['details']['Expected'] > 0]
            if len(monthly_policies) > 0:
                st.dataframe(monthly_policies[['Policy', 'Recruiter', 'Expected']], use_container_width=True)

# ==================== 报表导出 ====================
elif page == "📊 报表导出":
    st.markdown("### 导出报表")

    if st.session_state.calc_results is None:
        st.warning("⚠️ 请先在「佣金计算」页面完成计算")
    else:
        results = st.session_state.calc_results

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("#### 完整报表")
            summary_df = pd.DataFrame([{
                '项目': '保单总数', '数值': results['summary']['total_policies']
            }, {
                '项目': '总TP', '数值': results['summary']['total_tp']
            }, {
                '项目': '已收佣金', '数值': results['summary']['total_received']
            }, {
                '项目': '预期佣金', '数值': results['summary']['total_expected']
            }])

            excel_data = export_to_excel({
                '汇总': summary_df,
                '按经纪人': results['by_recruiter'],
                '明细': results['details']
            })

            st.download_button(
                label="📥 下载完整报表",
                data=excel_data,
                file_name="佣金报表_完整.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col2:
            st.markdown("#### 个人报表")
            if len(results['by_recruiter']) > 0:
                recruiters = results['by_recruiter']['Recruiter'].tolist()
                selected = st.selectbox("选择经纪人", recruiters)

                person_data = results['details'][results['details']['Recruiter'] == selected]
                if len(person_data) > 0:
                    excel_data = export_to_excel({
                        '个人明细': person_data
                    })
                    st.download_button(
                        label="📥 下载个人报表",
                        data=excel_data,
                        file_name=f"佣金报表_{selected}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
