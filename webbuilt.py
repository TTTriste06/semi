import os
import streamlit as st
import pandas as pd
import requests
import base64
from io import BytesIO
from datetime import datetime, timedelta
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment


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

def preprocess_mapping_file(df):
    # 只取前6列
    df = df.iloc[:, :6]
    # 重命名列
    df.columns = ['旧规格', '旧品名', '旧晶圆品名', '新规格', '新品名', '新晶圆品名']
    return df

def download_mapping_from_github(path_in_repo):
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{path_in_repo}"
    response = requests.get(api_url, headers={
        "Authorization": f"token {GITHUB_TOKEN}"
    })
    if response.status_code == 200:
        content = base64.b64decode(response.json()['content'])
        df = pd.read_excel(pd.io.common.BytesIO(content))
        if df.shape[1] >= 6:
            df = preprocess_mapping_file(df)
        else:
            st.warning(f"mapping_file.xlsx 列数不足：只发现 {df.shape[1]} 列，需要至少 6 列")
            df = pd.DataFrame(columns=['旧规格', '旧品名', '旧晶圆品名', '新规格', '新品名', '新晶圆品名'])
        return df
    else:
        st.warning("GitHub 上找不到 mapping_file.xlsx，用默认表或请先上传")
        return pd.DataFrame(columns=['旧规格', '旧品名', '旧晶圆品名', '新规格', '新品名', '新晶圆品名'])

def download_excel_from_github(url, token=None):
    headers = {"Authorization": f"token {token}"} if token else {}
    response = requests.get(url, headers=headers)
    content_type = response.headers.get('Content-Type', '')

    # 检查文件是不是 Excel
    if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' not in content_type:
        raise ValueError("下载的不是 Excel 文件，请检查 GitHub 链接或权限")

    return pd.read_excel(BytesIO(response.content))

import pandas as pd
import requests
from io import BytesIO

def download_backup_file(file_name):
    api_url = f"https://api.github.com/repos/{REPO_NAME}/contents/{file_name}"
    headers = {"Authorization": f"token {GITHUB_TOKEN}"}
    response = requests.get(api_url, headers=headers)

    if response.status_code != 200:
        st.warning(f"⚠️ 无法下载 {file_name}，GitHub 返回码 {response.status_code}")
        return pd.DataFrame()  # 返回空 DataFrame，保证主程序不崩溃

    content = response.json().get('content')
    if not content:
        st.warning(f"⚠️ {file_name} 文件内容为空或解析失败")
        return pd.DataFrame()

    file_bytes = BytesIO(base64.b64decode(content))

    try:
        df = pd.read_excel(file_bytes)
    except Exception as e:
        st.warning(f"⚠️ {file_name} 解析 Excel 失败：{e}，将创建空 sheet。")
        return pd.DataFrame()

    return df
        
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

def create_pivot(df, config, filename, mapping_df=None):
    if 'date_format' in config:
        config = config.copy()
        config['columns'] = f"{config['columns']}_年月"
    pivoted = pd.pivot_table(df, index=config['index'], columns=config['columns'], values=config['values'],
                             aggfunc=config['aggfunc'], fill_value=0)
    pivoted.columns = [f"{col[0]}_{col[1]}" if isinstance(col, tuple) else col for col in pivoted.columns]
    pivoted = pivoted.reset_index()

    if mapping_df is not None and filename == "赛卓-未交订单.xlsx":
        pivoted = apply_mapping_and_merge(pivoted, mapping_df)

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

    st.title('Excel 数据处理与汇总工具')
    selected_month = st.text_input('请输入截至月份（如 2025-03，可选）')
    CONFIG['selected_month'] = selected_month if selected_month else None

    uploaded_files = st.file_uploader('上传 Excel 文件（5个文件）', type=['xlsx'], accept_multiple_files=True)
    pred_file = st.file_uploader('上传预测文件', type=['xlsx'], key='pred_file')
    safety_file = st.file_uploader('上传安全库存文件', type=['xlsx'], key='safety_file')
    mapping_file = st.file_uploader('上传新旧料号文件', type=['xlsx'], key='mapping_file')

    # 加载 mapping_file DataFrame
    mapping_df = None
    if mapping_file:
        mapping_df = pd.read_excel(mapping_file)
        mapping_df = preprocess_mapping_file(mapping_df)

    if pred_file and st.button("保存预测文件到 GitHub"):
        upload_to_github(pred_file, "pred_file.xlsx", "上传预测文件")
    if safety_file and st.button("保存安全库存文件到 GitHub"):
        upload_to_github(safety_file, "safety_file.xlsx", "上传安全库存文件")
    if mapping_file and st.button("保存新旧料号文件到 GitHub"):
        upload_to_github(mapping_file, "mapping_file.xlsx", "上传新旧料号文件")

    if st.button('提交并生成报告') and uploaded_files:
        with pd.ExcelWriter(CONFIG['output_file'], engine='openpyxl') as writer:
            # 用于存储未交订单的前三列数据
            unfulfilled_orders_summary = pd.DataFrame()
    
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
    
                # 保存未交订单的前三列（去重）
                if filename == "赛卓-未交订单.xlsx":
                    cols_to_copy = [col for col in pivoted.columns if col in ["晶圆品名", "规格", "品名"]]
                    unfulfilled_orders_summary = pivoted[cols_to_copy].drop_duplicates()
    
            # 写入安全库存 sheet
            if safety_file:
                df_safety = pd.read_excel(safety_file)
            else:
                df_safety = download_backup_file("safety_file.xlsx")
            df_safety.to_excel(writer, sheet_name='赛卓-安全库存', index=False)
            adjust_column_width(writer, '赛卓-安全库存', df_safety)
            
            # 写入预测文件 sheet
            if pred_file:
                df_pred = pd.read_excel(pred_file)
            else:
                df_pred = download_backup_file("pred_file.xlsx")
            df_pred.to_excel(writer, sheet_name='赛卓-预测', index=False)
            adjust_column_width(writer, '赛卓-预测', df_pred)
            
            # 写入新旧料号文件 sheet
            if mapping_file:
                df_mapping = pd.read_excel(mapping_file)
            else:
                df_mapping = download_backup_file("mapping_file.xlsx")
            df_mapping.to_excel(writer, sheet_name='赛卓-新旧料号', index=False)
            adjust_column_width(writer, '赛卓-新旧料号', df_mapping)

    
            # 写入汇总 sheet
            if not unfulfilled_orders_summary.empty:
                unfulfilled_orders_summary.to_excel(writer, sheet_name='汇总', index=False, startrow=1)
                adjust_column_width(writer, '汇总', unfulfilled_orders_summary)

                worksheet = writer.book['汇总']
                
                # 合并 D1:E1（第4,5列的第一行）并写入 "安全库存"
                worksheet.merge_cells('D1:E1')
                worksheet['D1'] = '安全库存'
                worksheet['D1'].alignment = Alignment(horizontal='center', vertical='center')
            
                # 设置 D2, E2 的标题
                worksheet['D2'] = 'InvWaf（片）'
                worksheet['D2'].alignment = Alignment(horizontal='center', vertical='center')
                worksheet['E2'] = 'InvPart'
                worksheet['E2'].alignment = Alignment(horizontal='center', vertical='center')
            
                # 自动调整列宽
                for idx, col in enumerate(worksheet.columns, 1):
                    col_letter = get_column_letter(idx)
                    max_length = 0
                    for cell in col:
                        try:
                            if cell.value:
                                cell_len = len(str(cell.value))
                                # 中文字符按1.5倍宽度处理
                                cell_len = sum(2 if ord(char) > 127 else 1 for char in str(cell.value))
                                max_length = max(max_length, cell_len)
                        except:
                            pass
                    worksheet.column_dimensions[col_letter].width = max_length + 5  # 留余量

                                

    
        # 下载按钮
        with open(CONFIG['output_file'], 'rb') as f:
            st.download_button('下载汇总报告', f, CONFIG['output_file'])


if __name__ == '__main__':
    main()
