import pandas as pd
from preprocessing import apply_mapping_and_merge
from excel_utils import process_date_column
from config import CONFIG


def create_pivot(df, config, filename, mapping_df=None):
    """
    根据配置创建透视表（支持多 index / columns / values），可选进行新旧料号映射。
    """
    df_copy = df.copy()
    
    # 添加 "_年月" 列（如果有日期格式要求）
    if 'date_format' in config and config.get('columns') in df_copy.columns:
        df_copy = process_date_column(df_copy, config['columns'], config['date_format'])
        config = config.copy()  # 不修改原始配置
        config['columns'] = f"{config['columns']}_年月"

    # 创建透视表
    pivoted = pd.pivot_table(
        df_copy,
        index=config['index'],
        columns=config['columns'],
        values=config['values'],
        aggfunc=config['aggfunc'],
        fill_value=0
    )

    # 扁平化多层列名（例如：('订单数量', '2025-03') → '订单数量_2025-03'）
    pivoted.columns = [
        f"{col[0]}_{col[1]}" if isinstance(col, tuple) else str(col)
        for col in pivoted.columns
    ]
    pivoted = pivoted.reset_index()

    # 对未交订单文件做额外的料号映射
    if mapping_df is not None and filename == "赛卓-未交订单.xlsx":
        pivoted = apply_mapping_and_merge(pivoted, mapping_df)

    # 对未交订单文件进行历史订单汇总
    if CONFIG['selected_month'] and filename == "赛卓-未交订单.xlsx":
        pivoted = add_historical_order_columns(pivoted, config)

    return pivoted


def add_historical_order_columns(pivoted_df, config):
    """
    对透视表添加 '历史订单数量' 与 '历史未交订单数量' 列，并删除原始旧列。
    """
    # 提取历史月份列（例如 "_2024-12" < selected_month）
    history_cols = [
        col for col in pivoted_df.columns
        if '_' in col and col.split('_')[-1][:4].isdigit() and col.split('_')[-1] < CONFIG['selected_month']
    ]

    history_order_cols = [
        col for col in history_cols if '订单数量' in col and '未交订单数量' not in col
    ]
    history_pending_cols = [
        col for col in history_cols if '未交订单数量' in col
    ]

    if history_order_cols:
        pivoted_df['历史订单数量'] = pivoted_df[history_order_cols].sum(axis=1)
    if history_pending_cols:
        pivoted_df['历史未交订单数量'] = pivoted_df[history_pending_cols].sum(axis=1)

    pivoted_df.drop(columns=history_cols, inplace=True)

    # 控制输出顺序：index列 + 历史订单数量 + 历史未交订单数量 + 其他
    fixed_cols = [col for col in pivoted_df.columns if col not in ['历史订单数量', '历史未交订单数量']]
    if '历史订单数量' in pivoted_df.columns:
        fixed_cols.insert(len(config['index']), '历史订单数量')
    if '历史未交订单数量' in pivoted_df.columns:
        fixed_cols.insert(len(config['index']) + 1, '历史未交订单数量')

    pivoted_df = pivoted_df[fixed_cols]
    return pivoted_df
