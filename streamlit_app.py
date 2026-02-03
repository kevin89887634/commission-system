"""
💰 佣金管理系统 v2.5
最终修复版：直接解析Excel，不依赖pandas的header检测
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
    s = str(policy).strip()
    if not s or s.lower() in ['nan', 'none', 'policy', 'policy #']:
        return False
    if not any(c.isdigit() for c in s):
        return False
    # 必须以LS/NL/L开头或者纯数字
    if not (s.upper().startswith(('LS', 'NL', 'L')) or s[0].isdigit()):
        return False
    return True

def parse_nlg_file(uploaded_file):
    """
    解析NLG文件，返回DataFrame
    尝试多种方式读取直到成功
    """
    # 方法1: 尝试不同的header行
    for header_row in [5, 4, 6, 3, 1, 0]:
        try:
            df = pd.read_excel(uploaded_file, header=header_row, engine='openpyxl')
            uploaded_file.seek(0)  # 重置文件指针

            # 检查是否找到了Policy列
            cols_lower = [str(c).lower() for c in df.columns]
            has_policy = any('policy' in c for c in cols_lower)

            if has_policy and len(df) > 0:
                # 找到Policy列的实际名称
                policy_col = None
                for c in df.columns:
                    if 'policy' in str(c).lower():
                        policy_col = c
                        break

                # 检查第一行数据是否是有效的保单号
                first_val = str(df[policy_col].iloc[0]) if len(df) > 0 else ''
                if is_valid_policy(first_val):
                    return df, header_row, None
        except Exception as e:
            uploaded_file.seek(0)
            continue

    # 方法2: 读取原始数据，手动查找header
    try:
        uploaded_file.seek(0)
        df_raw = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
        uploaded_file.seek(0)

        # 遍历前15行找包含Policy的行
        for idx in range(min(15, len(df_raw))):
            row_str = ' '.join([str(v).lower() for v in df_raw.iloc[idx] if pd.notna(v)])
            if 'policy' in row_str and ('insured' in row_str or 'agent' in row_str or 'modal' in row_str):
                # 找到表头行
                df = pd.read_excel(uploaded_file, header=idx, engine='openpyxl')
                uploaded_file.seek(0)
                return df, idx, None
    except Exception as e:
        return None, None, str(e)

    return None, None, "无法找到有效的表头行"

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
    st.caption("v2.5 - 最终修复版")
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

    uploaded_file = st.file_uploader("上传 NLG New Business Report", type=['xlsx', 'xls'])

    if uploaded_file and st.button("📥 导入数据", type="primary"):
        with st.spinner("导入中..."):
            try:
                # 解析文件
                df, header_row, error = parse_nlg_file(uploaded_file)

                if error:
                    st.error(f"❌ 解析失败: {error}")
                    st.stop()

                if df is None or len(df) == 0:
                    st.error("❌ 未能读取到数据")
                    st.stop()

                st.info(f"📋 检测到表头在第 {header_row + 1} 行，共 {len(df)} 行数据")

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
                    elif 'status' in col_lower:
                        col_map[col] = 'Status'

                df = df.rename(columns=col_map)

                # 显示找到的列
                st.info(f"📊 识别的列: {list(col_map.values())}")

                # 检查必要列
                if 'Policy' not in df.columns:
                    st.error(f"❌ 找不到Policy列。当前列: {list(df.columns)}")
                    st.stop()

                # 过滤有效保单
                df['_valid'] = df['Policy'].apply(is_valid_policy)
                valid_count_before = df['_valid'].sum()
                df = df[df['_valid']].drop(columns=['_valid'])

                st.info(f"📋 有效保单: {valid_count_before} 条")

                if len(df) == 0:
                    st.error("❌ 过滤后没有有效数据")
                    # 显示原始数据前5行帮助调试
                    st.write("原始数据前5行:")
                    uploaded_file.seek(0)
                    df_debug = pd.read_excel(uploaded_file, header=header_row, engine='openpyxl')
                    st.dataframe(df_debug.head())
                    st.stop()

                # 处理数值列
                df['Policy_Norm'] = df['Policy'].apply(normalize_policy)
                df['Modal'] = df['Modal'].apply(safe_float) if 'Modal' in df.columns else 0
                df['AAP'] = df['AAP'].apply(safe_float) if 'AAP' in df.columns else 0

                # 过滤有保费的记录
                df = df[(df['AAP'] > 0) | (df['Modal'] > 0)].reset_index(drop=True)

                if len(df) == 0:
                    st.error("❌ 没有找到有保费的记录（AAP或Modal > 0）")
                    st.stop()

                st.session_state.df_raw = df

                # 生成分单表
                splits_data = []
                for _, row in df.iterrows():
                    modal = safe_float(row.get('Modal', 0))
                    aap = safe_float(row.get('AAP', 0))

                    # 判断缴费类型
                    if modal > 0 and aap > 0 and aap / modal > 6:
                        pay_type = '月缴'
                        premium = modal
                    else:
                        pay_type = '年缴'
                        premium = aap if aap > 0 else modal

                    # 判断佣金比例
                    product = str(row.get('Product', '')).lower()
                    comm_rate = 0.67 if 'term' in product else 0.80

                    # 获取Recruiter
                    recruiter = str(row.get('Recruiter', '')) if pd.notna(row.get('Recruiter', '')) else ''

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
                import traceback
                st.code(traceback.format_exc())

    # 数据预览
    if st.session_state.df_raw is not None:
        st.markdown("### 📊 数据预览")
        preview_cols = [c for c in ['Policy', 'Insured', 'Recruiter', 'Product', 'Modal', 'AAP']
                       if c in st.session_state.df_raw.columns]
        st.dataframe(st.session_state.df_raw[preview_cols], use_container_width=True)

# ==================== 第二步：编辑分单 ====================
elif step == "2️⃣ 编辑分单":
    st.header("2️⃣ 编辑分单")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先上传并导入数据")
        st.stop()

    st.markdown("### 📝 编辑分佣信息")
    st.caption("可以修改分佣人员和比例，Split1 + Split2 应该等于 1.0")

    # 编辑表格
    edited_df = st.data_editor(
        st.session_state.df_splits,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            'Policy': st.column_config.TextColumn('保单号', disabled=True),
            'Insured': st.column_config.TextColumn('被保人', disabled=True),
            'AAP': st.column_config.NumberColumn('AAP', disabled=True, format="$%.2f"),
            'Modal': st.column_config.NumberColumn('Modal', disabled=True, format="$%.2f"),
            'PayType': st.column_config.TextColumn('缴费类型', disabled=True),
            'Premium': st.column_config.NumberColumn('计算保费', disabled=True, format="$%.2f"),
            'CommRate': st.column_config.NumberColumn('佣金率', format="%.2f"),
            'Person1': st.column_config.TextColumn('分佣人1'),
            'Rate1': st.column_config.NumberColumn('比例1', format="%.2f"),
            'Split1': st.column_config.NumberColumn('分成1', format="%.2f"),
            'Person2': st.column_config.TextColumn('分佣人2'),
            'Rate2': st.column_config.NumberColumn('比例2', format="%.2f"),
            'Split2': st.column_config.NumberColumn('分成2', format="%.2f"),
        }
    )

    if st.button("💾 保存修改", type="primary"):
        st.session_state.df_splits = edited_df
        st.success("✅ 已保存")

# ==================== 第三步：计算佣金 ====================
elif step == "3️⃣ 计算佣金":
    st.header("3️⃣ 计算佣金")

    if st.session_state.df_splits is None:
        st.warning("⚠️ 请先完成前面的步骤")
        st.stop()

    if st.button("🧮 开始计算", type="primary"):
        results = []
        df = st.session_state.df_splits

        for _, row in df.iterrows():
            policy = row['Policy']
            premium = safe_float(row['Premium'])
            comm_rate = safe_float(row.get('CommRate', 0.80))

            # 计算总佣金
            gross = premium * comm_rate
            override = premium * 0.48
            total_comm = premium * (comm_rate + 0.48)

            # 分佣计算: 个人佣金 = Premium × Rate × Split
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
        else:
            st.error("❌ 没有可计算的记录")

    # 显示结果
    if st.session_state.df_results is not None:
        st.markdown("### 📊 计算结果")
        st.dataframe(st.session_state.df_results, use_container_width=True)

        # 汇总
        st.markdown("### 📈 分人汇总")
        summary = st.session_state.df_results.groupby('Person').agg({
            'PersonComm': 'sum',
            'Policy': 'count'
        }).rename(columns={'Policy': 'Count', 'PersonComm': 'TotalComm'})
        summary['TotalComm'] = summary['TotalComm'].apply(lambda x: f"${x:,.2f}")
        st.dataframe(summary, use_container_width=True)

        # 导出
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

    st.markdown("### 📤 上传对账文件")

    col1, col2 = st.columns(2)

    with col1:
        override_file = st.file_uploader("Override by Policy", type=['xlsx', 'xls'], key='override')

    with col2:
        gross_file = st.file_uploader("Payable Gross Commission", type=['xlsx', 'xls'], key='gross')

    if st.button("🔍 开始对账", type="primary"):
        results = st.session_state.df_results.copy()

        # 处理Override文件
        if override_file:
            try:
                df_ov = pd.read_excel(override_file, header=1, engine='openpyxl')
                # 找到Policy和Amount列
                policy_col = None
                amount_col = None
                for col in df_ov.columns:
                    col_lower = str(col).lower()
                    if 'policy' in col_lower:
                        policy_col = col
                    if 'amount' in col_lower or 'total' in col_lower:
                        amount_col = col

                if policy_col and amount_col:
                    df_ov['Policy_Norm'] = df_ov[policy_col].apply(lambda x: normalize_policy(str(x)))
                    override_map = dict(zip(df_ov['Policy_Norm'], df_ov[amount_col].apply(safe_float)))
                    results['Override_Actual'] = results['Policy'].map(override_map)
                    st.success(f"✅ Override文件: {len(override_map)} 条")
            except Exception as e:
                st.error(f"❌ Override解析失败: {e}")

        # 处理Gross文件
        if gross_file:
            try:
                df_gr = pd.read_excel(gross_file, header=4, engine='openpyxl')
                policy_col = None
                gross_col = None
                for col in df_gr.columns:
                    col_lower = str(col).lower()
                    if 'policy' in col_lower:
                        policy_col = col
                    if 'gross' in col_lower or 'commission' in col_lower:
                        gross_col = col

                if policy_col and gross_col:
                    df_gr['Policy_Norm'] = df_gr[policy_col].apply(lambda x: normalize_policy(str(x)))
                    gross_map = dict(zip(df_gr['Policy_Norm'], df_gr[gross_col].apply(safe_float)))
                    results['Gross_Actual'] = results['Policy'].map(gross_map)
                    st.success(f"✅ Gross文件: {len(gross_map)} 条")
            except Exception as e:
                st.error(f"❌ Gross解析失败: {e}")

        # 显示对账结果
        st.markdown("### 📊 对账结果")
        st.dataframe(results, use_container_width=True)
