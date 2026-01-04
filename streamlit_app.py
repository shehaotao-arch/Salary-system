import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill, Protection
from openpyxl.utils import get_column_letter
import io
from datetime import datetime

st.set_page_config(page_title="计件工资结算系统", layout="wide")

# 辅助函数：校验数字
def is_valid_number(value):
    try:
        if pd.isna(value) or str(value).strip() == "": return False
        float(value)
        return True
    except: return False

def safe_sheet_name(name):
    """处理工作表名称，移除非法字符"""
    illegal_chars = ['/', '\\', '?', '*', '[', ']', ':', "'"]
    safe_name = name
    for char in illegal_chars:
        safe_name = safe_name.replace(char, '-')
    # 截断到31个字符（Excel工作表名称最大长度）
    return safe_name[:31]

def set_excel_style_vba(ws, emp_name, target_month, data_rows, price_dict, subsist_val):
    """根据VBA代码重构的样式设置函数"""
    
    # 样式定义
    error_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黄色 - 错误值
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")  # 表头灰色
    theme_color = "BFBFBF"  # 边框颜色
    thin_border = Border(
        left=Side(style='thin', color=theme_color),
        right=Side(style='thin', color=theme_color),
        top=Side(style='thin', color=theme_color),
        bottom=Side(style='thin', color=theme_color)
    )
    
    # ===== 1. 标题与表头 =====
    # 插入两行空行
    ws.insert_rows(1, amount=2)
    
    # 大标题
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = f"{target_month.replace('-', '年')}月{emp_name}工资明细表"
    title_cell.font = Font(name='黑体', size=16, bold=True)
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 35
    
    # 员工姓名
    ws["A2"].value = f"员工：{emp_name}"
    ws["A2"].font = Font(name='微软雅黑', size=12, bold=True)
    
    # 表头
    headers = ["日期", "产品名称", "数量", "工价", "金额"]
    for i, header in enumerate(headers):
        ws.cell(row=3, column=i+1, value=header)
        ws.cell(row=3, column=i+7, value=header)
    
    # ===== 2. 写入数据明细 =====
    details = []
    for _, row in data_rows.iterrows():
        # 数据行结构: 第2列:日期, 第3列:产品名称, 第4列:数量, 价格从字典获取
        product_name = str(row.iloc[3]).strip() if len(row) > 3 else ""
        date_val = row.iloc[2] if len(row) > 2 else ""
        qty_val = row.iloc[4] if len(row) > 4 else 0
        price_val = price_dict.get(product_name, 0)
        details.append({
            'date': date_val,
            'product': product_name,
            'qty': qty_val,
            'price': price_val
        })
    
    # 分栏逻辑
    data_count = len(details)
    left_count = (data_count + 1) // 2
    
    # 写入左栏数据（A-E列）
    for i in range(min(left_count, data_count)):
        row_idx = 4 + i
        detail = details[i]
        
        # 日期
        date_cell = ws.cell(row=row_idx, column=1, value=detail['date'])
        if hasattr(detail['date'], 'strftime'):
            date_cell.number_format = 'm/d'  # 日期格式
        
        # 产品名称
        ws.cell(row=row_idx, column=2, value=detail['product'])
        
        # 数量
        qty_cell = ws.cell(row=row_idx, column=3, value=detail['qty'])
        
        # 工价
        price_cell = ws.cell(row=row_idx, column=4)
        
        # 检查工价是否存在，如果不存在则设为0并标黄
        if is_valid_number(detail['price']) and float(detail['price']) != 0:
            price_cell.value = float(detail['price'])
        else:
            price_cell.value = 0
            price_cell.fill = error_fill
        
        # 金额公式（=C列*D列）
        amount_cell = ws.cell(row=row_idx, column=5)
        amount_cell.value = f"=C{row_idx}*D{row_idx}"
        
        # 检查数量是否为0或无效，如果是则标黄
        if not is_valid_number(detail['qty']) or float(detail['qty']) == 0:
            qty_cell.fill = error_fill
            amount_cell.fill = error_fill  # 金额列也标黄
    
    # 写入右栏数据（G-K列）
    for i in range(left_count, data_count):
        row_idx = 4 + (i - left_count)
        detail = details[i]
        
        # 日期
        date_cell = ws.cell(row=row_idx, column=7, value=detail['date'])
        if hasattr(detail['date'], 'strftime'):
            date_cell.number_format = 'm/d'  # 日期格式
        
        # 产品名称
        ws.cell(row=row_idx, column=8, value=detail['product'])
        
        # 数量
        qty_cell = ws.cell(row=row_idx, column=9, value=detail['qty'])
        
        # 工价
        price_cell = ws.cell(row=row_idx, column=10)
        
        # 检查工价是否存在，如果不存在则设为0并标黄
        if is_valid_number(detail['price']) and float(detail['price']) != 0:
            price_cell.value = float(detail['price'])
        else:
            price_cell.value = 0
            price_cell.fill = error_fill
        
        # 金额公式（=I列*J列）
        amount_cell = ws.cell(row=row_idx, column=11)
        amount_cell.value = f"=I{row_idx}*J{row_idx}"
        
        # 检查数量是否为0或无效，如果是则标黄
        if not is_valid_number(detail['qty']) or float(detail['qty']) == 0:
            qty_cell.fill = error_fill
            amount_cell.fill = error_fill  # 金额列也标黄
    
    # ===== 3. 确定汇总行 =====
    last_row_A = ws.max_row
    last_row_G = ws.max_row
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=1).value is not None:
            last_row_A = row
            break
    for row in range(ws.max_row, 0, -1):
        if ws.cell(row=row, column=7).value is not None:
            last_row_G = row
            break
    
    sum_row = max(last_row_A, last_row_G) + 1
    
    # ===== 4. 汇总信息 =====
    # 生活费
    ws.cell(row=sum_row, column=8, value="生活费：")
    
    subsist_cell = ws.cell(row=sum_row, column=9)
    if is_valid_number(subsist_val):
        subsist_cell.value = float(subsist_val)
    else:
        subsist_cell.value = 0
        subsist_cell.fill = error_fill  # 生活费为0或无效时标黄
    
    # 总计标签
    ws.cell(row=sum_row, column=10, value="总计：")
    
    # 总计公式 - 使用SUMIF忽略错误值
    total_cell = ws.cell(row=sum_row, column=11)
    
    # 构建E列和K列的范围
    e_end = sum_row - 1
    k_end = sum_row - 1
    
    # 公式：=(SUMIF(E4:E{end},">0")+SUMIF(K4:K{end},">0"))*0.97-I{sum_row}
    formula = f"=(SUMIF(E4:E{e_end},\">0\")+SUMIF(K4:K{k_end},\">0\"))*0.97-I{sum_row}"
    total_cell.value = formula
    total_cell.number_format = '0'  # 整数格式
    
    # ===== 5. 列宽设置 =====
    # 完全按照VBA代码设置列宽
    column_widths = {
        'A': 7.25,   # 日期左栏
        'B': 18.0,   # 产品名称左栏
        'C': 6.25,   # 数量左栏
        'D': 5.75,   # 工价左栏
        'E': 7.18,   # 金额左栏
        'F': 2.0,    # 空栏
        'G': 7.25,   # 日期右栏
        'H': 18.0,   # 产品名称右栏
        'I': 6.25,   # 数量右栏
        'J': 5.75,   # 工价右栏
        'K': 7.18    # 金额右栏
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    # ===== 6. 行高设置 =====
    # 第一行：35
    ws.row_dimensions[1].height = 35
    # 第二行：25
    ws.row_dimensions[2].height = 25
    # 第3行到sum_row行：21
    for row in range(3, sum_row + 1):
        ws.row_dimensions[row].height = 21
    
    # ===== 7. 边框与对齐 =====
    # 应用边框（A3:K{sum_row}）
    for row in range(3, sum_row + 1):
        for col in range(1, 12):  # A-K列
            cell = ws.cell(row=row, column=col)
            cell.border = thin_border
            cell.font = Font(name='微软雅黑', size=12)
            cell.alignment = Alignment(vertical='center')
            
            # 水平对齐方式
            if col in [1, 7]:  # A列和G列（日期）居中
                cell.alignment = Alignment(vertical='center', horizontal='center')
            elif col in [2, 8]:  # B列和H列（产品名称）左对齐
                cell.alignment = Alignment(vertical='center', horizontal='left')
            elif col in [3, 4, 5, 9, 10, 11]:  # 数量、工价、金额列右对齐
                cell.alignment = Alignment(vertical='center', horizontal='right')
    
    # ===== 8. 表头样式 =====
    for col in range(1, 6):  # A-E列表头
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = Font(name='微软雅黑', size=12, bold=True)
    
    for col in range(7, 12):  # G-K列表头
        cell = ws.cell(row=3, column=col)
        cell.fill = header_fill
        cell.font = Font(name='微软雅黑', size=12, bold=True)
    
    # ===== 9. 打印设置 =====
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.horizontalCentered = True
    
    # 页边距（接近VBA的默认值）
    ws.page_margins.left = 0.7
    ws.page_margins.right = 0.7
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3
    
    # ===== 10. 检查并标黄所有错误金额 =====
    # 检查左栏金额（E列）
    for row in range(4, sum_row):
        amount_cell = ws.cell(row=row, column=5)
        # 尝试计算金额，如果公式结果可能为0或错误，则标黄
        try:
            # 获取数量和工价
            qty_cell = ws.cell(row=row, column=3)
            price_cell = ws.cell(row=row, column=4)
            
            qty_val = qty_cell.value
            price_val = price_cell.value
            
            # 检查是否应该标黄
            should_fill = False
            if qty_cell.fill.start_color.index == error_fill.start_color.index:
                should_fill = True
            elif price_cell.fill.start_color.index == error_fill.start_color.index:
                should_fill = True
            elif not is_valid_number(qty_val) or not is_valid_number(price_val):
                should_fill = True
            elif float(qty_val) == 0 or float(price_val) == 0:
                should_fill = True
            
            if should_fill:
                amount_cell.fill = error_fill
        except:
            amount_cell.fill = error_fill
    
    # 检查右栏金额（K列）
    for row in range(4, sum_row):
        amount_cell = ws.cell(row=row, column=11)
        # 尝试计算金额，如果公式结果可能为0或错误，则标黄
        try:
            # 获取数量和工价
            qty_cell = ws.cell(row=row, column=9)
            price_cell = ws.cell(row=row, column=10)
            
            qty_val = qty_cell.value
            price_val = price_cell.value
            
            # 检查是否应该标黄
            should_fill = False
            if qty_cell.fill.start_color.index == error_fill.start_color.index:
                should_fill = True
            elif price_cell.fill.start_color.index == error_fill.start_color.index:
                should_fill = True
            elif not is_valid_number(qty_val) or not is_valid_number(price_val):
                should_fill = True
            elif float(qty_val) == 0 or float(price_val) == 0:
                should_fill = True
            
            if should_fill:
                amount_cell.fill = error_fill
        except:
            amount_cell.fill = error_fill

# --- Streamlit 界面 ---
st.title(" 🚀浩德陶瓷工资导出系统")

with st.sidebar:
    target_month = st.text_input("请输入年月 (YYYY-MM)", "2025-10")
    uploaded_file = st.file_uploader("请上传 Excel资料表（包含工价表、生产表、生活费、员工）", type=["xlsx", "xlsm"])

if st.button("开始生成"):
    if uploaded_file:
        try:
            # 读取所有sheet
            sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=object)
            
            # 获取各表
            df_s = sheets["生产表"]
            df_e = sheets["员工"]
            df_p = sheets["工价表"]
            df_b = sheets["生活费"]
            
            # 显示数据预览
            st.subheader("数据预览")
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("员工表 (前10行):")
                st.dataframe(df_e.head(10))
                st.write(f"员工总数: {len(df_e)}")
            
            with col2:
                st.write("生产表结构:")
                st.write(df_s.head())
            
            # 数据清洗
            # 1. 员工表
            employee_col = df_e.columns[0]
            df_e[employee_col] = df_e[employee_col].astype(str).str.strip()
            
            # 2. 生产表
            name_col = df_s.columns[1]  # B列
            date_col = df_s.columns[2]  # C列
            
            # 填充姓名并清理
            df_s[name_col] = df_s[name_col].ffill().astype(str).str.strip()
            
            # 3. 工价表字典
            price_dict = {}
            if len(df_p.columns) >= 2:
                for _, row in df_p.iterrows():
                    key = str(row.iloc[0]).strip()
                    if key and key.lower() != 'nan':
                        price_dict[key] = row.iloc[1]
            
            # 4. 生活费字典
            subsist_dict = {}
            if len(df_b.columns) >= 2:
                for _, row in df_b.iterrows():
                    key = str(row.iloc[0]).strip()
                    if key and key.lower() != 'nan':
                        subsist_dict[key] = row.iloc[1]
            
            # 创建新工作簿
            new_wb = Workbook()
            # 删除默认sheet
            if 'Sheet' in new_wb.sheetnames:
                del new_wb['Sheet']
            
            # 处理目标月份格式
            target_month_formatted = target_month.replace('.', '-')
            
            # 为每个员工创建sheet
            count = 0
            employee_names = []
            
            for _, row in df_e.iterrows():
                emp_name = str(row.iloc[0]).strip()
                if not emp_name or emp_name.lower() == 'nan':
                    continue
                
                employee_names.append(emp_name)
                
                # 筛选该员工的生产记录
                mask = (
                    (df_s[name_col] == emp_name) &
                    (df_s[date_col].astype(str).str.contains(target_month_formatted))
                )
                emp_data = df_s[mask]
                
                # 获取生活费
                subsist_val = 0
                for key in subsist_dict:
                    if emp_name in key or key in emp_name:
                        subsist_val = subsist_dict[key]
                        break
                
                # 创建sheet
                safe_name = safe_sheet_name(emp_name)
                ws = new_wb.create_sheet(title=safe_name)
                
                # 设置样式
                set_excel_style_vba(ws, emp_name, target_month_formatted, emp_data, price_dict, subsist_val)
                count += 1
                
                st.write(f"✓ 已生成: {emp_name} - 记录: {len(emp_data)}条")
            
            # 保存到BytesIO
            output = io.BytesIO()
            new_wb.save(output)
            output.seek(0)
            
            st.success(f"✅ 成功生成 {count} 位员工工资表！")
            
            # 下载按钮
            st.download_button(
                label="📥 下载 Excel 文件",
                data=output.getvalue(),
                file_name=f"{target_month_formatted}_全员工资表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"❌ 运行出错: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
    else:
        st.warning("请先上传Excel文件！")