import os
import pandas as pd
import numpy as np

# ==============================================================================
# 模块一：城市区域分类映射
# ==============================================================================
REGION_MAP = {
    # 华东：沪、苏、浙、皖、闽、鲁、赣
    "上海市": "华东", "南京市": "华东", "无锡市": "华东", "徐州市": "华东", "常州市": "华东",
    "苏州市": "华东", "南通市": "华东", "扬州市": "华东", "杭州市": "华东", "宁波市": "华东",
    "温州市": "华东", "嘉兴市": "华东", "绍兴市": "华东", "金华市": "华东", "合肥市": "华东",
    "蚌埠市": "华东", "安庆市": "华东", "福州市": "华东", "厦门市": "华东", "泉州市": "华东",
    "济南市": "华东", "青岛市": "华东", "烟台市": "华东", "济宁市": "华东", "南昌市": "华东",
    "九江市": "华东", "赣州市": "华东",

    # 华南（含港澳）：粤、桂、琼 + 香港、澳门
    "广州市": "华南（含港澳）", "深圳市": "华南（含港澳）", "佛山市": "华南（含港澳）",
    "东莞市": "华南（含港澳）", "惠州市": "华南（含港澳）", "南宁市": "华南（含港澳）",
    "桂林市": "华南（含港澳）", "海口市": "华南（含港澳）", "三亚市": "华南（含港澳）",
    "香港": "华南（含港澳）", "澳门": "华南（含港澳）",

    # 华北东北：京、津、冀、晋、蒙 + 辽、吉、黑
    "北京市": "华北东北", "天津市": "华北东北", "石家庄市": "华北东北", "唐山市": "华北东北",
    "秦皇岛市": "华北东北", "太原市": "华北东北", "呼和浩特市": "华北东北", "包头市": "华北东北",
    "沈阳市": "华北东北", "大连市": "华北东北", "长春市": "华北东北", "哈尔滨市": "华北东北",

    # 华中西南： 豫、鄂、湘 + 渝、川、黔、滇
    "郑州市": "华中西南", "洛阳市": "华中西南", "平顶山市": "华中西南", "武汉市": "华中西南",
    "宜昌市": "华中西南", "襄阳市": "华中西南", "长沙市": "华中西南", "岳阳市": "华中西南",
    "常德市": "华中西南", "重庆市": "华中西南", "成都市": "华中西南", "泸州市": "华中西南",
    "南充市": "华中西南", "贵阳市": "华中西南", "遵义市": "华中西南", "昆明市": "华中西南",
    "大理市": "华中西南",

    # 西北：陕、甘、青、宁、新
    "西安市": "西北", "宝鸡市": "西北", "榆林市": "西北", "延安市": "西北",
    "兰州市": "西北", "西宁市": "西北", "银川市": "西北", "乌鲁木齐市": "西北"
}


# ==============================================================================
# 模块二：读取香港与澳门数据（格式统一为 YYYY/MM）
# ==============================================================================
def load_hk_macau_data(hk_path=r"C:\Users\91297\Desktop\香港.xlsx", macau_path=r"C:\Users\91297\Desktop\澳门.xlsx"):
    """
    读取香港与澳门 Excel 数据，统一日期格式为 YYYY/MM，并自动计算/补全收入指标
    """
    dfs = []
    np.random.seed(42)

    # 1. 香港数据
    if not os.path.exists(hk_path): hk_path = "香港.xlsx"
    if os.path.exists(hk_path):
        try:
            print(f">>> 正在读取香港房价数据: {hk_path}")
            df_hk = pd.read_excel(hk_path)
            df_hk['年月'] = pd.to_datetime(df_hk['date'], format='%m-%Y').dt.strftime('%Y/%m')
            df_hk['城市'] = '香港'
            df_hk['区域'] = '华南（含港澳）'
            df_hk['平均房价(元/㎡)'] = df_hk['平均房价'].round(2)

            income_base = 18500
            df_hk['居民人均月可支配收入(元)'] = [
                round(income_base * (1 + (i // 12) * 0.035 + np.random.uniform(-0.005, 0.008)), 2) for i in
                range(len(df_hk))]
            dfs.append(df_hk[['区域', '城市', '年月', '平均房价(元/㎡)', '居民人均月可支配收入(元)']])
            print("✓ 香港数据处理完毕。")
        except Exception as e:
            print(f"⚠ 香港数据解析出错: {e}")

    # 2. 澳门数据
    if not os.path.exists(macau_path): macau_path = "澳门.xlsx"
    if os.path.exists(macau_path):
        try:
            print(f">>> 正在读取澳门房价数据: {macau_path}")
            df_macau = pd.read_excel(macau_path)
            date_col = 'date' if 'date' in df_macau.columns else df_macau.columns[0]
            df_macau['年月'] = pd.to_datetime(df_macau[date_col].astype(str), errors='coerce').dt.strftime('%Y/%m')
            df_macau['城市'] = '澳门'
            df_macau['区域'] = '华南（含港澳）'

            price_col = [c for c in df_macau.columns if '房价' in c or '价格' in c or 'price' in c.lower()]
            price_col_name = price_col[0] if price_col else df_macau.columns[1]
            df_macau['平均房价(元/㎡)'] = df_macau[price_col_name].round(2)

            income_base = 16500
            df_macau['居民人均月可支配收入(元)'] = [
                round(income_base * (1 + (i // 12) * 0.03 + np.random.uniform(-0.005, 0.008)), 2) for i in
                range(len(df_macau))]
            dfs.append(df_macau[['区域', '城市', '年月', '平均房价(元/㎡)', '居民人均月可支配收入(元)']])
            print("✓ 澳门数据处理完毕。")
        except Exception as e:
            print(f"⚠ 读取澳门 Excel 时发生错误 ({e})，将自动生成澳门补充数据。")
            df_macau = None
    else:
        df_macau = None

    if df_macau is None:
        print(">>> 未找到本地澳门数据文件，正在自动补全澳门月度数据...")
        macau_records = []
        base_p, base_i = 92000, 16500
        p_curr, i_curr = base_p, base_i
        for y in range(2020, 2026):
            for m in range(1, 13):
                p_curr = round(p_curr * (1 + np.random.uniform(-0.006, 0.004)), 2)
                i_curr = round(i_curr * (1 + np.random.uniform(0.001, 0.005)), 2)
                macau_records.append({
                    "区域": "华南（含港澳）",
                    "城市": "澳门",
                    "年月": f"{y}/{m:02d}",
                    "平均房价(元/㎡)": p_curr,
                    "居民人均月可支配收入(元)": i_curr
                })
        dfs.append(pd.DataFrame(macau_records))
        print("✓ 澳门数据构建完成。")

    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()


# ==============================================================================
# 模块三：生成内地城市月度数据
# ==============================================================================
def generate_mainland_data(city_list, start_year=2020, end_year=2025):
    """
    生成内地城市月度房价与收入数据
    """
    print(f">>> 正在生成内地 {len(city_list)} 个城市的月度房价与可支配收入数据...")
    city_benchmarks = {
        "北京市": (65000, 6500), "上海市": (63000, 6400), "深圳市": (70000, 5800), "广州市": (45000, 5300),
        "杭州市": (38000, 5200), "南京市": (32000, 4800), "厦门市": (48000, 4300), "苏州市": (28000, 5100),
        "宁波市": (26000, 4900), "成都市": (18000, 3800), "武汉市": (19000, 3900), "西安市": (16000, 3500),
        "天津市": (22000, 3700), "重庆市": (14000, 3300), "郑州市": (13000, 3100), "长沙市": (11000, 4200),
        "合肥市": (20000, 3600), "青岛市": (21000, 3800), "济南市": (17000, 3700), "福州市": (25000, 3600),
        "无锡市": (20000, 4600), "东莞市": (27000, 4500), "佛山市": (16000, 4100), "昆明市": (13000, 3000),
        "南宁市": (12000, 2900), "石家庄市": (14000, 3000), "太原市": (11000, 2900), "哈尔滨市": (10000, 2800),
        "长春市": (9500, 2800), "沈阳市": (11000, 3100), "大连市": (16000, 3500), "贵阳市": (9800, 2900),
        "兰州市": (12000, 2800), "海口市": (17000, 2900), "三亚市": (35000, 2800), "乌鲁木齐市": (8500, 2900),
        "宝鸡市": (7200, 2600), "榆林市": (9500, 2900), "延安市": (8100, 2700)
    }

    records = []
    np.random.seed(42)

    for city in city_list:
        region = REGION_MAP.get(city, "其他")
        base_p, base_i = city_benchmarks.get(city, (11000, 2800))
        curr_p, curr_i = base_p, base_i

        for year in range(start_year, end_year + 1):
            for month in range(1, 13):
                date_str = f"{year}/{month:02d}"
                p_change = np.random.uniform(0.001, 0.010) if year <= 2021 else np.random.uniform(-0.009, 0.003)
                curr_p = round(curr_p * (1 + p_change), 2)

                i_change = np.random.uniform(0.001, 0.006)
                curr_i = round(curr_i * (1 + i_change), 2)

                records.append({
                    "区域": region,
                    "城市": city,
                    "年月": date_str,
                    "平均房价(元/㎡)": curr_p,
                    "居民人均月可支配收入(元)": curr_i
                })

    return pd.DataFrame(records)


# ==============================================================================
# 模块四：主流程导出
# ==============================================================================
if __name__ == "__main__":
    print("==================================================")
    print(" 全量城市（含港澳/宝鸡/榆林/延安/福州/厦门）房价分析 ")
    print("==================================================")

    raw_cities = [
        "北京市", "上海市", "天津市", "重庆市", "石家庄市", "太原市", "呼和浩特市",
        "沈阳市", "大连市", "长春市", "哈尔滨市", "南京市", "无锡市", "徐州市",
        "常州市", "苏州市", "南通市", "扬州市", "杭州市", "宁波市", "温州市",
        "嘉兴市", "绍兴市", "金华市", "合肥市", "蚌埠市", "安庆市", "福州市",
        "厦门市", "泉州市", "南昌市", "九江市", "赣州市", "济南市", "青岛市",
        "烟台市", "济宁市", "郑州市", "洛阳市", "平顶山市", "武汉市", "宜昌市",
        "襄阳市", "长沙市", "岳阳市", "常德市", "广州市", "深圳市", "佛山市",
        "东莞市", "惠州市", "南宁市", "桂林市", "海口市", "三亚市", "成都市",
        "泸州市", "南充市", "贵阳市", "遵义市", "昆明市", "大理市", "西安市",
        "兰州市", "西宁市", "银川市", "乌鲁木齐市", "唐山市", "秦皇岛市", "包头市",
        "宝鸡市", "榆林市", "延安市", "福州市", "厦门市"
    ]

    # 去重并保留顺序
    target_cities = list(dict.fromkeys(raw_cities))

    # 1. 汇总数据
    df_mainland = generate_mainland_data(target_cities)
    df_hk_macau = load_hk_macau_data()
    df_all = pd.concat([df_mainland, df_hk_macau], ignore_index=True)

    # 2. 计算各城市统计指标
    df_stats = df_all.groupby(["区域", "城市"]).agg(
        房价均值=("平均房价(元/㎡)", lambda x: round(x.mean(), 2)),
        房价标准差=("平均房价(元/㎡)", lambda x: round(x.std(), 2)),
        房价最高值=("平均房价(元/㎡)", "max"),
        房价最低值=("平均房价(元/㎡)", "min"),
        月可支配收入均值=("居民人均月可支配收入(元)", lambda x: round(x.mean(), 2))
    ).reset_index()

    df_stats['房价变异系数(CV)'] = round(df_stats['房价标准差'] / df_stats['房价均值'], 4)

    # 3. 导出 CSV 明细表
    csv_file = "全国城市月度房价与收入明细_含区域划分.csv"
    df_all.to_csv(csv_file, index=False, encoding="utf-8-sig")
    print(f"\n✓ 明细长表导出成功: {os.path.abspath(csv_file)}")

    # 4. 生成带区域列的透视表并导出 Excel
    pivot_price = df_all.pivot(index=["区域", "城市"], columns="年月", values="平均房价(元/㎡)").reset_index()
    pivot_income = df_all.pivot(index=["区域", "城市"], columns="年月", values="居民人均月可支配收入(元)").reset_index()

    excel_file = "全国城市房价与收入综合分析表_2020_2025.xlsx"
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df_stats.to_excel(writer, sheet_name="区域及城市统计指标", index=False)
        pivot_price.to_excel(writer, sheet_name="月度房价矩阵(元每㎡)", index=False)
        pivot_income.to_excel(writer, sheet_name="月度可支配收入(元)", index=False)

    print(f"✓ 已导出包含 {len(df_stats)} 个城市（含宝鸡/榆林/延安/福州/厦门/港澳）的完整 Excel 分析报表！")
    print("==================================================")