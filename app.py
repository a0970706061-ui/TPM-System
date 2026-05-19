import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from collections import defaultdict
from datetime import datetime
import re

st.set_page_config(layout="wide")
st.title("TPM 保養排程系統")

production_file = st.file_uploader("上傳生產排程 Excel", type=["xlsx"])
maintenance_file = st.file_uploader("上傳保養紀錄 Excel", type=["xlsx"])

selected_year = st.number_input("選擇年份", min_value=2024, max_value=2035, value=2026)
selected_month = st.number_input("選擇月份", min_value=1, max_value=12, value=6)
start_day = st.number_input("起始日期", min_value=1, max_value=31, value=1)

exclude_precision = st.checkbox("排除精密加工區", value=True)
max_per_day = st.number_input("每日最多保養台數", min_value=1, max_value=10, value=2)

def norm_machine(x):

    if x is None:
        return None

    s = str(x).strip().upper()

    s = s.replace(" ", "")
    s = s.replace("－", "-")
    s = s.replace("_", "-")

    s = s.split("\n")[0]

    # 抓 英文字母 + 數字
    match = re.search(r"([A-Z]+)[-]?0*(\d+)", s)

    if match:

        prefix = match.group(1)
        num = int(match.group(2))

        return f"{prefix}-{num:03d}"

    return s
    
def is_machine_code(x):
    if x is None:
        return False

    s = str(x).strip().upper()
    s = s.replace(" ", "")
    s = s.split("\n")[0]

    return bool(re.match(r"^[A-Z]+-?\d+$", s))

def is_precision_machine(machine):
    machine = str(machine)
    return machine.startswith(("HC-", "HW-", "HMA-", "HLR-"))

if st.button("生成TPM排程"):

    if production_file is None:
        st.error("請上傳生產排程")
        st.stop()

    if maintenance_file is None:
        st.error("請上傳保養紀錄")
        st.stop()

    st.write("讀取生產排程中...")

    wb = load_workbook(production_file, data_only=True)

    # 找日期列
    target_ws = None
    target_date_row = None

    for sheet in wb.worksheets:
        for r in range(1, 30):
            date_count = 0

            for c in range(1, sheet.max_column + 1):
                v = sheet.cell(r, c).value
                if isinstance(v, datetime):
                    date_count += 1

            if date_count >= 10:
                target_ws = sheet
                target_date_row = r
                break

        if target_ws is not None:
            break

    if target_ws is None:
        st.error("找不到日期列")
        st.stop()

    ws = target_ws

    st.success(f"已抓到工作表：{ws.title}")
    st.success(f"日期列：第 {target_date_row} 列")

    # 抓指定年月日期欄
    date_cols = {}

    for col in range(1, ws.max_column + 1):
        v = ws.cell(target_date_row, col).value

        if isinstance(v, datetime):
            if v.year == selected_year and v.month == selected_month and v.day >= start_day:
                date_cols[col] = {
                    "日期": f"{v.month}/{v.day}",
                    "日期排序": v
                }

    if not date_cols:
        st.error(f"找不到 {selected_year} 年 {selected_month} 月 {start_day} 號之後的日期")
        st.stop()

    st.success(f"抓到 {len(date_cols)} 個日期欄")

    # 找機台與料號列
    st.write("搜尋機台與料號列...")

    machine_part_rows = defaultdict(list)
    machine_area = {}

    current_machine = None
    current_area = ""

    for row in range(target_date_row + 1, ws.max_row + 1):

        machine_cell = ws.cell(row, 1).value
        area_cell = ws.cell(row, 3).value
        label_cell = ws.cell(row, 4).value

        if is_machine_code(machine_cell):
            current_machine = norm_machine(machine_cell)
            current_area = str(area_cell).strip() if area_cell is not None else ""
            machine_area[current_machine] = current_area

        # 重點：只抓 D欄 = 料號 的列
        if current_machine and str(label_cell).strip() == "料號":
            machine_part_rows[current_machine].append(row)

    st.success(f"抓到 {len(machine_part_rows)} 台機台")

    if len(machine_part_rows) == 0:
        st.error("沒有抓到料號列，請確認 D 欄是否有『料號』")
        st.stop()

    # 判斷空機
    st.write("開始判斷空機...")

    idle_rows = []
    progress = st.progress(0)

    machines = list(machine_part_rows.keys())

    for idx, machine in enumerate(machines):

        rows = machine_part_rows[machine]
        area = machine_area.get(machine, "")

        for col, info in date_cols.items():

            has_schedule = False

            for r in rows:
                val = ws.cell(r, col).value

                # 日期格只要有任何內容，例如重工，也算有排程
                if val is not None and str(val).strip() != "":
                    has_schedule = True
                    break

            if not has_schedule:
                idle_rows.append({
                    "日期": info["日期"],
                    "日期排序": info["日期排序"],
                    "機台": machine,
                    "廠區": area
                })

        progress.progress((idx + 1) / len(machines))

    idle_df = pd.DataFrame(idle_rows)

    if idle_df.empty:
        st.warning("沒有空機")
        st.stop()

    st.success(f"抓到 {len(idle_df)} 筆空機")

    # 排除精密加工區
    if exclude_precision:
        idle_df = idle_df[
            (idle_df["廠區"] != "精密加工區") &
            (~idle_df["機台"].apply(is_precision_machine))
        ]

    if idle_df.empty:
        st.warning("排除精密加工區後沒有可排機台")
        st.stop()

    # 讀取所有機台保養紀錄
    st.write("讀取所有機台上次保養時間...")

    try:
        maint_df = pd.read_excel(maintenance_file, sheet_name="所有機台")
    except:
        st.error("找不到『所有機台』工作表")
        st.stop()

    machine_col = maint_df.columns[0]
    time_col = maint_df.columns[1]

    maintenance_map = {}

    for _, row in maint_df.iterrows():

        machine = norm_machine(row[machine_col])
        last_time = row[time_col]

        if machine is None:
            continue

        if pd.isna(last_time) or str(last_time).strip() in ["", "無紀錄", "無", "N/A"]:
            maintenance_map[machine] = "無紀錄"
        else:
            maintenance_map[machine] = str(last_time)

    idle_df["上次保養時間"] = idle_df["機台"].map(maintenance_map).fillna("無紀錄")

    idle_df["是否無保養紀錄"] = idle_df["上次保養時間"].apply(
        lambda x: "是" if str(x).strip() == "無紀錄" else "否"
    )

    # 讀總加工時間
    total_time_map = {}

    try:
        time_df = pd.read_excel(maintenance_file, sheet_name="總加工時間")

        time_machine_col = time_df.columns[0]
        time_sec_col = time_df.columns[1]

        for _, row in time_df.iterrows():

            machine = norm_machine(row[time_machine_col])

            if machine is None:
                continue

            try:
                total_time_map[machine] = float(row[time_sec_col])
            except:
                total_time_map[machine] = 0

    except:
        st.warning("找不到『總加工時間』工作表，總加工時間以 0 計算")

    idle_df["總加工秒數"] = idle_df["機台"].map(total_time_map).fillna(0)

    # 高加工時間：前20%
    if idle_df["總加工秒數"].max() > 0:
        threshold = idle_df["總加工秒數"].quantile(0.8)
    else:
        threshold = 0

    idle_df["是否高加工時間"] = idle_df["總加工秒數"].apply(
        lambda x: "是" if x >= threshold and x > 0 else "否"
    )

    # 優先排序
    def priority_score(row):
        if row["是否無保養紀錄"] == "是" and row["是否高加工時間"] == "是":
            return 1
        elif row["是否無保養紀錄"] == "是":
            return 2
        elif row["是否高加工時間"] == "是":
            return 3
        else:
            return 4

    def priority_reason(row):
        if row["是否無保養紀錄"] == "是" and row["是否高加工時間"] == "是":
            return "無保養紀錄 + 高加工時間"
        elif row["是否無保養紀錄"] == "是":
            return "無保養紀錄"
        elif row["是否高加工時間"] == "是":
            return "高加工時間"
        else:
            return "一般空機"

    idle_df["排序"] = idle_df.apply(priority_score, axis=1)
    idle_df["優先原因"] = idle_df.apply(priority_reason, axis=1)

    idle_df = idle_df.sort_values(
        ["日期排序", "排序", "總加工秒數"],
        ascending=[True, True, False]
    )

    # 生成保養排程：每天最多 max_per_day，不重複，不跨日補位
    scheduled = set()
    result_rows = []

    for date in idle_df["日期"].drop_duplicates():

        day_df = idle_df[idle_df["日期"] == date]
        count = 0

        for _, row in day_df.iterrows():

            machine = row["機台"]

            if machine in scheduled:
                continue

            result_rows.append({
                "日期": row["日期"],
                "順序": count + 1,
                "機台": machine,
                "優先原因": row["優先原因"],
                "上次保養時間": row["上次保養時間"],
                "總加工秒數": int(row["總加工秒數"]),
                "完成確認": "□"
            })

            scheduled.add(machine)
            count += 1

            if count >= max_per_day:
                break

    result_df = pd.DataFrame(result_rows)

    st.success("TPM排程完成")

    st.write("### TPM保養排程")
    st.dataframe(result_df)

    st.write("### 空機優先清單")
    show_df = idle_df.drop(columns=["日期排序", "排序", "廠區"])
    st.dataframe(show_df)