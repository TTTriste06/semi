from datetime import datetime

# GitHub 配置
GITHUB_TOKEN_KEY = "GITHUB_TOKEN"  # 用于 st.secrets 获取
REPO_NAME = "TTTriste06/semi"
BRANCH = "main"

# 选择月份
CONFIG = {
    "selected_month": None
}

# 默认输出文件名
OUTPUT_FILE = f"运营数据订单-在制-库存汇总报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# Excel 文件透视表配置
PIVOT_CONFIG = {
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

# 新旧料号列名（全字段）
FULL_MAPPING_COLUMNS = [
    '旧规格', '旧品名', '旧晶圆品名',
    '新规格', '新品名', '新晶圆品名',
    '封装厂', 'PC', '半成品'
]

# 精简版本用于 apply_mapping_and_merge()
MAPPING_COLUMNS = [
    '旧规格', '旧品名', '旧晶圆品名',
    '新规格', '新品名', '新晶圆品名'
]

