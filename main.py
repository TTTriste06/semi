import streamlit as st
import pandas as pd
from openpyxl import load_workbook

from config import CONFIG, OUTPUT_FILE, PIVOT_CONFIG, FULL_MAPPING_COLUMNS
from github_utils import upload_to_github, download_backup_file
from preprocessing import preprocess_mapping_file
from pivot_processor import create_pivot
from excel_utils import adjust_column_width, auto_adjust_column_width_by_worksheet, add_black_border
from merge_sections import (
    merge_safety_inventory,
    merge_unfulfilled_orders,
    merge_prediction_data,
    mark_unmatched_rows
)
from ui import setup_sidebar, get_user_inputs


def main():
    st.set_page_config(page_title='数据汇总自动化工具', layout='wide')
    setup_sidebar()

    uploaded_files, pred_file, safety_file, mapping_file = get_user_inputs()

    if pred_file:
        upload_to_github(pred_file, "pred_file.xlsx", "上传预测文件")
    if safety_file:
        upload_to_github(safety_file, "safety_file.xlsx", "上传安全库存文件")
    if mapping_file:
        upload_to_github(mapping_file, "mapping_file.xlsx", "上传新旧料号文件")

    if st.button('🚀 提交并生成报告') and uploaded_files:
        mapping_df = pd.read_excel(mapping_file) if mapping_file else download_backup_file("mapping_file.xlsx")
        mapping_df = preprocess_mapping_file(mapping_df)

        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            summary_df = pd.DataFrame()
            pending_df = None

            for f in uploaded_files:
                filename = f.name
                if filename not in PIVOT_CONFIG:
                    st.warning(f"⚠️ 未配置文件：{filename}")
                    continue

                df = pd.read_excel(f)
                pivoted = create_pivot(df, PIVOT_CONFIG[filename], filename, mapping_df)
                sheet_name = filename.replace('.xlsx', '')[:30]
                pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                adjust_column_width(writer, sheet_name, pivoted)

                if filename == "赛卓-未交订单.xlsx":
                    summary_df = pivoted[['晶圆品名', '规格', '品名']].drop_duplicates()
                    pending_df = pivoted.copy()

            # 汇总 sheet 初步写入
            summary_df.to_excel(writer, sheet_name='汇总', index=False, startrow=1)
            summary_sheet = writer.sheets['汇总']

            # 合并安全库存
            df_safety = pd.read_excel(safety_file) if safety_file else download_backup_file("safety_file.xlsx")
            merged_summary_df, df_safety = merge_safety_inventory(summary_df, df_safety, summary_sheet)

            # 合并未交订单
            if pending_df is not None:
                start_col = summary_df.shape[1] + 2 + 1
                _ = merge_unfulfilled_orders(summary_sheet, pending_df, start_col)

            # 合并预测文件
            df_pred = pd.read_excel(pred_file) if pred_file else download_backup_file("pred_file.xlsx")
            df_pred = merge_prediction_data(summary_sheet, df_pred, summary_df)

            # 标红未匹配预测
            pred_ws = writer.book['赛卓-预测']
            mark_unmatched_rows(pred_ws, df_pred, start_row=3)

            # 汇总样式调整
            auto_adjust_column_width_by_worksheet(summary_sheet)
            add_black_border(summary_sheet, 2, summary_sheet.max_column)

        # 下载按钮
        with open(OUTPUT_FILE, 'rb') as f:
            st.download_button('📥 下载汇总报告', f, OUTPUT_FILE)


if __name__ == '__main__':
    main()
