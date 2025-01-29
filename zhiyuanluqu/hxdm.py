import streamlit as st
import pandas as pd
from io import BytesIO
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# 网页标识
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        right: 10px;
        bottom: 10px;
        color: #888;
        font-family: Arial;
    }
    </style>
    <div class="footer">Developed by Jsyy</div>
    """,
    unsafe_allow_html=True
)

def process_time_selections(df):
    """处理时间段拆分逻辑"""
    try:
        time_pattern = r"\d{4}-\d{2}-\d{2} \d{2}:\d{2} \d{4}-\d{2}-\d{2} \d{2}:\d{2}"
        expanded_rows = []
        for _, row in df.iterrows():
            time_entries = re.findall(time_pattern, str(row["所选时间"]))
            for entry in time_entries:
                new_row = row.copy()
                new_row["被录取时间段"] = entry
                expanded_rows.append(new_row)
        return pd.DataFrame(expanded_rows).reset_index(drop=True)
    except KeyError:
        st.error("输入文件缺少必要列'所选时间'")
        return pd.DataFrame()

def load_existing_records(uploaded_file):
    """加载历史录取文件"""
    try:
        df = pd.read_excel(uploaded_file)
        # 自动识别时间段列
        time_col = next((col for col in df.columns if any(
            re.match(r".*\d{4}-\d{2}-\d{2}.*", str(x)) for x in df[col])), None)
        if time_col:
            df["被录取时间段"] = df[time_col].apply(
                lambda x: re.findall(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2} \d{4}-\d{2}-\d{2} \d{2}:\d{2}", str(x))[0])
        # 检查必要列是否存在
        required_columns = ["姓名", "学号", "联系方式", "被录取时间段"]
        if not all(col in df.columns for col in required_columns):
            st.error("历史文件缺少必要列，请确保包含：姓名、学号、联系方式、时间段")
            return pd.DataFrame()
        return df
    except Exception as e:
        st.error(f"历史文件解析失败: {str(e)}")
        return pd.DataFrame()

def auto_adjust_column_width(ws):
    """自动调整Excel列宽"""
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

def main():
    st.title("志愿者录取系统")
    
    # ================= 核心功能 =================
    # 1. 上传报名信息表
    uploaded_file = st.file_uploader("上传报名信息表", type=["xlsx"])
    if not uploaded_file:
        return

    # 处理原始数据
    try:
        df = pd.read_excel(uploaded_file, dtype={"学号": str, "联系方式": str})
        df["报名时间"] = range(1, len(df) + 1)  # 生成唯一报名序号
    except Exception as e:
        st.error(f"报名表读取失败: {str(e)}")
        return

    # 2. 判断是否第一次录取
    is_first_time = st.radio("是否是第一次录取？", ("是", "否"), index=0)
    existing_df = pd.DataFrame()
    
    if is_first_time == "否":
        existing_file = st.file_uploader("上传历史录取文件", type=["xlsx"])
        if existing_file:
            existing_df = load_existing_records(existing_file)
            if existing_df.empty:
                return
        else:
            st.warning("请上传历史录取文件！")
            return
    
    # 3. 增量输出选项
    need_incremental = st.checkbox("仅输出新增录取人员（不包含历史记录）") if is_first_time == "否" else False
    
    # 4. 黑名单管理
    blacklist = st.multiselect("选择黑名单学生", 
                             options=df[["姓名", "学号", "联系方式"]].to_dict("records"),
                             format_func=lambda x: f"{x['姓名']} ({x['学号']})")
    
    # 5. 退出时间段管理
    processed_df = process_time_selections(df)
    if processed_df.empty:
        return
    exit_options = processed_df[["姓名", "学号", "联系方式", "被录取时间段"]].to_dict("records")
    exit_students = st.multiselect("选择退出学生（需选择具体时间段）",
                                 options=exit_options,
                                 format_func=lambda x: f"{x['姓名']} - {x['被录取时间段']}")
    
    # 6. 设置录取人数
    required = st.number_input("每个时间段需要人数", min_value=1, value=5)
    
    if st.button("开始录取"):
        # ================= 数据处理 =================
        # 过滤黑名单
        blacklist_ids = {f"{x['姓名']}|{x['学号']}|{x['联系方式']}" for x in blacklist}
        processed_df["student_id"] = processed_df["姓名"] + "|" + processed_df["学号"] + "|" + processed_df["联系方式"]
        filtered_df = processed_df[~processed_df["student_id"].isin(blacklist_ids)]
        
        # 过滤退出时间段
        exit_ids = {f"{x['姓名']}|{x['学号']}|{x['联系方式']}|{x['被录取时间段']}" for x in exit_students}
        filtered_df["exit_id"] = filtered_df["student_id"] + "|" + filtered_df["被录取时间段"]
        filtered_df = filtered_df[~filtered_df["exit_id"].isin(exit_ids)]
        
        # ================= 核心算法 =================
        # 生成唯一标识符
        filtered_df["unique_id"] = filtered_df["student_id"] + "|" + filtered_df["被录取时间段"]
        
        # 排除历史记录
        if not existing_df.empty:
            existing_df["unique_id"] = existing_df["姓名"] + "|" + existing_df["学号"].astype(str) + "|" + existing_df["联系方式"].astype(str) + "|" + existing_df["被录取时间段"]
            filtered_df = filtered_df[~filtered_df["unique_id"].isin(existing_df["unique_id"])]
        
        # 按时间段计算剩余名额
        final_results = []
        for time_slot, group in filtered_df.groupby("被录取时间段"):
            # 计算历史已用名额
            existing_count = len(existing_df[existing_df["被录取时间段"] == time_slot]) if not existing_df.empty else 0
            remaining = max(required - existing_count, 0)
            
            # 按报名顺序录取
            new_records = group.sort_values("报名时间").head(remaining)
            final_results.append(new_records)
        
        final_df = pd.concat(final_results)
        
        # ================= 输出控制 =================
        if not need_incremental and not existing_df.empty:
            final_df = pd.concat([
                existing_df[["姓名", "学号", "性别", "联系方式", "被录取时间段", "报名时间"]],
                final_df[["姓名", "学号", "性别", "联系方式", "被录取时间段", "报名时间"]]
            ], axis=0)
        
        # 重命名列并调整顺序
        final_df = final_df[["姓名", "学号", "性别", "联系方式", "被录取时间段", "报名时间"]]
        
        # ================= 生成Excel =================
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False)
            workbook = writer.book
            worksheet = workbook.active
            auto_adjust_column_width(worksheet)
        
        st.download_button("下载录取结果", 
                         data=output.getvalue(),
                         file_name="录取结果.xlsx",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        st.success(f"共录取 {len(final_df)} 人")

if __name__ == "__main__":
    main()



