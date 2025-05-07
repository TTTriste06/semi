import os
import streamlit as st
import pandas as pd
import requests
import base64
from datetime import datetime, timedelta
from openpyxl.utils import get_column_letter

GITHUB_TOKEN = st.secrets["GITHUB_TOKEN"]  # 在 Streamlit Cloud 用 secrets
REPO_NAME = "TTTriste06/semi"
BRANCH = "main"

CONFIG = {
    "output_file": f"运营数据订单-在制-库存汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    "selected_month": None,
    "pivot_config": {
        "赛卓-未交订单.xlsx": {
            "index": ["晶圆品名", "规格", "品名"],
            "columns": "预交货日",
            "values": ["订单数量", "未交订单数量"],
            "aggfunc": "sum",
            "date_format": "%Y-%m"
        },
        "赛卓-成品在制.xlsx": {
            "index": ["工作中心", "封装形式", "晶圆型号", "产品规格", "产品品名"],
            "columns": "预计完工日期",
            "values": ["未交"],
            "aggfunc": "sum",
            "date_format": "%Y-%m"
        },
        "赛卓-CP在制.xlsx": {
            "index": ["晶圆型号", "产品品名"],
            "columns": "预计完工日期",
            "values": ["未交"],
            "aggfunc": "sum",
            "date_format": "%Y-%m"
        },
        "赛卓-成品库存.xlsx": {
            "index": ["WAFER品名", "规格", "品名"],
            "columns": "仓库名称",
            "values": ["数量"],
            "aggfunc": "sum"
        },
        "赛卓-晶圆库存.xlsx": {
            "index": ["WAFER品名", "规格"],
            "columns": "仓库名称",
            "values": ["数量"],
            "aggfunc": "sum"
        }
    }
}

def upload_to_github(file, path_in_repo, commit_message):
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"
    file_content = file.read()
    encoded_content = base64.b64encode(file_content).decode('utf-8')

    response = requests.get(api_url, headers={
        "Authorization": f"token {GITHUB_TOKEN}"
    })
    if response.status_code == 200:
        sha = response.json()['sha']
    else:
        sha = None

    payload = {
        "message": commit_message,
        "content": encoded_content,
        "branch": BRANCH
    }
    if sha:
        payload["sha"] = sha

    response = requests.put(api_url, json=payload, headers={
        "Authorization": f"token {GITHUB_TOKEN}"
    })

    if response.status_code in [200, 201]:
        st.success(f"{path_in_repo} 上传成功！")
    else:
        st.error(f"上传失败: {response.json()}")

def download_mapping_from_github(path_in_repo):
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"
    response = requests.get(api_url, headers={
        "Authorization": f"token {GITHUB_TOKEN}"
    })
    if response.status_code == 200:
        content = base64.b64decode(response.json()['content'])
        df = pd.read_excel(pd.io.common.BytesIO(content))
        return df
    else:
        st.warning("GitHub 上找不到 mapping_file.xlsx，用默认表或请先上传")
        return pd.DataFrame()  # 或者你可以这里 return 一个默认 MAPPING_TABLE
        
def process_date_column(df, date_col, date_format):
    if pd.api.types.is_numeric_dtype(df[date_col]):
        df[date_col] = df[date_col].apply(lambda x: datetime(1899, 12, 30) + timedelta(days=float(x)))
    else:
        df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    df[f'{date_col}_年月'] = df[date_col].dt.strftime(date_format)
    return df

def apply_mapping_and_merge(df, mapping_df):
    mapping_df = mapping_df.dropna()
    
    df = df.merge(
        mapping_df,
        how='left',
        left_on=['晶圆品名', '规格', '品名'],
        right_on=['旧晶圆品名', '旧规格', '旧品名']
    )
    
    df['晶圆品名'] = df['新晶圆品名'].combine_first(df['晶圆品名'])
    df['规格'] = df['新规格'].combine_first(df['规格'])
    df['品名'] = df['新品名'].combine_first(df['品名'])
    
    df.drop(columns=['旧晶圆品名', '旧规格', '旧品名', '新晶圆品名', '新规格', '新品名'], inplace=True)
    
    group_cols = [col for col in df.columns if col not in df.select_dtypes(include='number').columns]
    agg_cols = df.select_dtypes(include='number').columns.tolist()
    df_merged = df.groupby(group_cols, as_index=False)[agg_cols].sum()
    
    return df_merged

def create_pivot(df, config, filename):
    if 'date_format' in config:
        config = config.copy()
        config['columns'] = f"{config['columns']}_年月"
    pivoted = pd.pivot_table(df, index=config['index'], columns=config['columns'], values=config['values'],
                             aggfunc=config['aggfunc'], fill_value=0)
    pivoted.columns = [f"{col[0]}_{col[1]}" if isinstance(col, tuple) else col for col in pivoted.columns]
    pivoted = pivoted.reset_index()
    
    if filename == "赛卓-未交订单.xlsx":
        if '规格' in pivoted.columns and '品名' in pivoted.columns and '晶圆品名' in pivoted.columns:
            pivoted = apply_mapping_and_merge(pivoted, MAPPING_TABLE)
    
    if CONFIG['selected_month'] and filename == "赛卓-未交订单.xlsx":
        history_cols = [col for col in pivoted.columns if '_' in col and col.split('_')[-1][:4].isdigit() and col.split('_')[-1] < CONFIG['selected_month']]
        history_order_cols = [col for col in history_cols if '订单数量' in col and '未交订单数量' not in col]
        history_pending_cols = [col for col in history_cols if '未交订单数量' in col]
        if history_order_cols:
            pivoted['历史订单数量'] = pivoted[history_order_cols].sum(axis=1)
        if history_pending_cols:
            pivoted['历史未交订单数量'] = pivoted[history_pending_cols].sum(axis=1)
        pivoted.drop(columns=history_cols, inplace=True)
        fixed_cols = [col for col in pivoted.columns if col not in ['历史订单数量', '历史未交订单数量']]
        if '历史订单数量' in pivoted.columns:
            fixed_cols.insert(len(config['index']), '历史订单数量')
        if '历史未交订单数量' in pivoted.columns:
            fixed_cols.insert(len(config['index']) + 1, '历史未交订单数量')
        pivoted = pivoted[fixed_cols]
    
    return pivoted


def adjust_column_width(writer, sheet_name, df):
    worksheet = writer.sheets[sheet_name]
    for idx, col in enumerate(df.columns, 1):
        max_len = df[col].astype(str).str.len().max()
        header_len = len(str(col))
        width = max(max_len, header_len) * 1.2 + 5
        worksheet.column_dimensions[get_column_letter(idx)].width = min(width, 50)

def main():
    st.set_page_config(
        page_title='我是标题',
        page_icon=' ',
        layout='wide'
    )

    with st.sidebar:
        st.title("欢迎来到我的应用")
        st.markdown('---')
        st.markdown('这是它的特性：\n- feature 1\n- feature 2\n- feature 3')

    global MAPPING_TABLE
    MAPPING_TABLE = download_mapping_from_github("mapping_file.xlsx")
    
    st.title('Excel 数据处理与汇总工具')
    selected_month = st.text_input('请输入截至月份（如 2025-03，可选）')
    CONFIG['selected_month'] = selected_month if selected_month else None

    # 普通 5个文件的上传
    uploaded_files = st.file_uploader('上传 Excel 文件（5个文件）', type=['xlsx'], accept_multiple_files=True)

    pred_file = st.file_uploader('上传预测文件', type=['xlsx'], key='pred_file')
    if pred_file and st.button("保存预测文件到 GitHub"):
        upload_to_github(pred_file, "pred_file.xlsx", "上传预测文件")
    
    safety_file = st.file_uploader('上传安全库存文件', type=['xlsx'], key='safety_file')
    if safety_file and st.button("保存安全库存文件到 GitHub"):
        upload_to_github(safety_file, "safety_file.xlsx", "上传安全库存文件")
    
    mapping_file = st.file_uploader('上传新旧料号文件', type=['xlsx'], key='mapping_file')
    if mapping_file and st.button("保存新旧料号文件到 GitHub"):
        upload_to_github(mapping_file, "mapping_file.xlsx", "上传新旧料号文件")
        
    if st.button('提交并生成报告') and uploaded_files:
        with pd.ExcelWriter(CONFIG['output_file'], engine='openpyxl') as writer:
            for f in uploaded_files:
                filename = f.name
                if filename not in CONFIG['pivot_config']:
                    st.warning(f"跳过未配置的文件: {filename}")
                    continue
                df = pd.read_excel(f)
                config = CONFIG['pivot_config'][filename]
                if 'date_format' in config and config['columns'] in df.columns:
                    df = process_date_column(df, config['columns'], config['date_format'])
                pivoted = create_pivot(df, config, filename)
                sheet_name = filename[:30].rstrip('.xlsx')
                pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                adjust_column_width(writer, sheet_name, pivoted)

        with open(CONFIG['output_file'], 'rb') as f:
            st.download_button('下载汇总报告', f, CONFIG['output_file'])

if __name__ == '__main__':
    main()
