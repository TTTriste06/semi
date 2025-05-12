import pandas as pd
from config import MAPPING_COLUMNS

def apply_full_mapping(df, mapping_df, spec_col, prod_col, wafer_col):
    """
    应用完整的新旧料号替换逻辑。
    
    参数说明：
    - df: 需要进行替换的主 DataFrame
    - mapping_df: 原始上传的新旧料号表（不管列名先做重命名）
    - spec_col, prod_col, wafer_col: 分别指定 df 中的 "规格"、"品名"、"晶圆品名" 的列名
    """
    # 标准列名
    full_cols = ['旧规格', '旧品名', '旧晶圆品名', '新规格', '新品名', '新晶圆品名', '封装厂', 'PC', '半成品']
    
    # 保证列数充足再赋名
    if mapping_df.shape[1] >= len(full_cols):
        mapping_df = mapping_df.iloc[:, :len(full_cols)]
        mapping_df.columns = full_cols
    else:
        raise ValueError(f"新旧料号表列数不足（当前 {mapping_df.shape[1]} 列），应为至少 9 列")

    # 删除空值
    mapping_df = mapping_df.dropna(subset=['旧规格', '旧品名', '旧晶圆品名'])

    # 进行 merge 替换
    df = df.merge(
        mapping_df,
        how='left',
        left_on=[wafer_col, spec_col, prod_col],
        right_on=['旧晶圆品名', '旧规格', '旧品名']
    )

    # 替换为新字段（优先用新）
    df[wafer_col] = df['新晶圆品名'].combine_first(df[wafer_col])
    df[spec_col] = df['新规格'].combine_first(df[spec_col])
    df[prod_col] = df['新品名'].combine_first(df[prod_col])

    # 打印出被替换的记录
    if show_changes:
        changed = df[
            (df[spec_col] != df['_原规格']) |
            (df[prod_col] != df['_原品名']) |
            (df[wafer_col] != df['_原晶圆'])
        ][[wafer_col, spec_col, prod_col, '_原晶圆', '_原规格', '_原品名']]

        if not changed.empty:
            st.write("✅ 以下记录发生了新旧料号替换：")
            st.dataframe(changed)
        else:
            st.info("ℹ️ 没有任何行发生新旧料号替换")

    # 删除 merge 带来的辅助列
    df.drop(columns=['旧规格', '旧品名', '旧晶圆品名', '新规格', '新品名', '新晶圆品名'], errors='ignore', inplace=True)

    # 聚合所有数值列
    group_cols = [col for col in df.columns if col not in df.select_dtypes(include='number').columns]
    agg_cols = df.select_dtypes(include='number').columns.tolist()
    df_merged = df.groupby(group_cols, as_index=False)[agg_cols].sum()

    return df_merged
