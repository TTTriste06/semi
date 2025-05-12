import pandas as pd
from config import MAPPING_COLUMNS


def preprocess_mapping_file(df):
    """
    清洗上传的新旧料号表，保留前6列并统一列名。
    """
    df = df.iloc[:, :6]
    df.columns = MAPPING_COLUMNS
    return df


def apply_mapping_and_merge(df, mapping_df):
    """
    将原始 DataFrame 中的 ['晶圆品名', '规格', '品名']
    替换为新料号表中的 ['新晶圆品名', '新规格', '新品名']。
    """
    mapping_df = mapping_df.dropna()

    # 合并新旧料号表
    df = df.merge(
        mapping_df,
        how='left',
        left_on=['晶圆品名', '规格', '品名'],
        right_on=['旧晶圆品名', '旧规格', '旧品名']
    )

    # 替换字段：新值优先
    df['晶圆品名'] = df['新晶圆品名'].combine_first(df['晶圆品名'])
    df['规格'] = df['新规格'].combine_first(df['规格'])
    df['品名'] = df['新品名'].combine_first(df['品名'])

    # 删除辅助列
    df.drop(columns=['旧晶圆品名', '旧规格', '旧品名', '新晶圆品名', '新规格', '新品名'], inplace=True)

    # 对替换后的结果重新 groupby 汇总（只保留非数值列为分组键）
    group_cols = [col for col in df.columns if col not in df.select_dtypes(include='number').columns]
    agg_cols = df.select_dtypes(include='number').columns.tolist()
    df_merged = df.groupby(group_cols, as_index=False)[agg_cols].sum()

    return df_merged
