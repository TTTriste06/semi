import streamlit as st
import pandas as pd
from openpyxl import load_workbook

from config import (
    GITHUB_TOKEN_KEY, REPO_NAME, BRANCH,
    CONFIG, OUTPUT_FILE, PIVOT_CONFIG,
    FULL_MAPPING_COLUMNS, COLUMN_MAPPING
)
from github_utils import upload_to_github, download_excel_from_url, download_excel_from_repo
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
        mapping_df = pd.read_excel(mapping_file) if mapping_file else download_excel_from_repo("mapping_file.xlsx")

        with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
            summary_df = pd.DataFrame()
            pending_df = None

            for f in uploaded_files:
                filename = f.name
            
                if filename not in PIVOT_CONFIG:
                    st.warning(f"跳过未配置的文件: {filename}")
                    continue
            
                df = pd.read_excel(f)
            
                # 统一新旧料号映射（若该文件定义了列名映射）
                if filename in COLUMN_MAPPING:
                    mapping = COLUMN_MAPPING[filename]
                    spec_col = mapping["规格"]
                    prod_col = mapping["品名"]
                    wafer_col = mapping["晶圆品名"]
            
                    if all(col in df.columns for col in [spec_col, prod_col, wafer_col]):
                        try:
                            df = apply_full_mapping(df, mapping_df, spec_col, prod_col, wafer_col)
                        except Exception as e:
                            st.warning(f"⚠️ 文件 {filename} 替换失败: {e}")
                    else:
                        st.warning(f"⚠️ 文件 {filename} 缺少字段: {spec_col}, {prod_col}, {wafer_col}")
                else:
                    st.info(f"📂 文件 {filename} 未定义映射字段，跳过 apply_full_mapping")

                # 透视
                pivoted = create_pivot(df, PIVOT_CONFIG[filename], filename, mapping_df)
                sheet_name = filename.replace('.xlsx', '')[:30]
                pivoted.to_excel(writer, sheet_name=sheet_name, index=False)
                adjust_column_width(writer, sheet_name, pivoted)

                # 提取汇总sheet的信息
                if filename == "赛卓-未交订单.xlsx":
                    summary_df = pivoted[['晶圆品名', '规格', '品名']].drop_duplicates()
                    pending_df = pivoted.copy()

            # 汇总 sheet 初步写入
            summary_df.to_excel(writer, sheet_name='汇总', index=False, startrow=1)
            summary_sheet = writer.sheets['汇总']

            # 加入安全库存
            df_safety = pd.read_excel(safety_file) if safety_file else download_excel_from_repo("safety_file.xlsx")
            merged_summary_df, df_safety = merge_safety_inventory(summary_df, df_safety, summary_sheet)

            # 加入未交订单
            if pending_df is not None:
                start_col = summary_df.shape[1] + 2 + 1
                _ = merge_unfulfilled_orders(summary_sheet, pending_df, start_col)

            # 加入预测
            df_pred = pd.read_excel(pred_file) if pred_file else download_excel_from_repo("pred_file.xlsx")
            df_pred = merge_prediction_data(summary_sheet, df_pred, summary_df)

            # 样式调整
            auto_adjust_column_width_by_worksheet(summary_sheet)
            add_black_border(summary_sheet, 2, summary_sheet.max_column)

        # 下载按钮
        with open(OUTPUT_FILE, 'rb') as f:
            st.download_button('📥 下载汇总报告', f, OUTPUT_FILE)


if __name__ == '__main__':
    main()
