import streamlit as st
from config import CONFIG


def setup_sidebar():
    with st.sidebar:
        st.title("欢迎来到我的应用")
        st.markdown('---')
        st.markdown('### 功能简介：')
        st.markdown('- 生成汇总报告')


def get_user_inputs():
    st.title('Excel 数据处理与汇总工具')
    selected_month = st.text_input('📅 请输入截至月份（如 2025-03，可选）')
    CONFIG['selected_month'] = selected_month if selected_month else None

    uploaded_files = st.file_uploader('📂 上传 5 个核心 Excel 文件（订单/库存/在制）', type=['xlsx'], accept_multiple_files=True)
    pred_file = st.file_uploader('📈 上传预测文件', type=['xlsx'], key='pred_file')
    safety_file = st.file_uploader('🔐 上传安全库存文件', type=['xlsx'], key='safety_file')
    mapping_file = st.file_uploader('🔁 上传新旧料号对照表', type=['xlsx'], key='mapping_file')

    return uploaded_files, pred_file, safety_file, mapping_file

