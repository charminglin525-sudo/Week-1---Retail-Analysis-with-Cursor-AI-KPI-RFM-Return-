import pandas as pd
import numpy as np
from datetime import datetime

# 讀取資料
print("正在讀取 online_retail_II.xlsx...")
df = pd.read_excel('online_retail_II.xlsx')

print(f"原始資料行數: {len(df)}")
print("\n資料列名:")
print(df.columns.tolist())

# 標準化列名（去除前後空格，統一大小寫處理）
df.columns = df.columns.str.strip()

# 查找相關列名（不區分大小寫）
def find_column(df, keywords):
    """查找包含關鍵詞的列名"""
    keywords_lower = [k.lower() for k in keywords]
    for col in df.columns:
        col_lower = col.lower()
        for keyword in keywords_lower:
            if keyword in col_lower:
                return col
    return None

quantity_col = find_column(df, ['Quantity', 'quantity', '數量'])
customerid_col = find_column(df, ['Customer ID', 'CustomerID', 'Customer', 'customer'])
description_col = find_column(df, ['Description', 'description', '描述'])
price_col = find_column(df, ['Price', 'price', '價格', '單價'])
date_col = find_column(df, ['InvoiceDate', 'Invoice Date', 'Date', 'date', '日期', '交易日期'])

print("\n=== 識別的列名 ===")
print(f"Quantity 列: {quantity_col}")
print(f"CustomerID 列: {customerid_col}")
print(f"Description 列: {description_col}")
print(f"Price 列: {price_col}")
print(f"Date 列: {date_col}")

# 檢查必要的列是否存在
if not quantity_col:
    raise ValueError("找不到 Quantity 列，請檢查資料文件")
if not customerid_col:
    raise ValueError("找不到 CustomerID 列，請檢查資料文件")
if not description_col:
    raise ValueError("找不到 Description 列，請檢查資料文件")
if not price_col:
    raise ValueError("找不到 Price 列，請檢查資料文件")
if not date_col:
    raise ValueError("找不到日期列，請檢查資料文件")

# 步驟1: 計算 Total = Quantity * Price
print("\n=== 步驟1: 計算 Total = Quantity * Price ===")
df['Total'] = df[quantity_col] * df[price_col]
print(f"已創建 Total 列")

# 步驟2: 篩除 CustomerID 和 Description 缺失的資料
print("\n=== 步驟2: 篩除 CustomerID 和 Description 缺失的資料 ===")
missing_mask = (df[customerid_col].isna()) | (df[description_col].isna())
abnormal_order = df[missing_mask].copy()
abnormal_count = len(abnormal_order)
print(f"CustomerID 或 Description 缺失的資料行數: {abnormal_count}")

# 從原始資料中移除異常訂單資料
df = df[~missing_mask].copy()

# 步驟3: 篩除 Quantity < 0 的資料
print("\n=== 步驟3: 篩除 Quantity < 0 的資料 ===")
return_mask = df[quantity_col] < 0
return_order = df[return_mask].copy()
return_count = len(return_order)
print(f"Quantity < 0 的資料行數: {return_count}")

# 從資料中移除退貨訂單資料
df = df[~return_mask].copy()

# 步驟4: 剩餘資料為正常訂單
print("\n=== 步驟4: 正常訂單 ===")
normal_order = df.copy()
normal_count = len(normal_order)
print(f"正常訂單資料行數: {normal_count}")

# 步驟5: 創建匯總統計
print("\n=== 步驟5: 創建匯總統計 ===")

# 計算 Gross Order = Normal Order + Return Order
gross_order_count = normal_count + return_count
gross_order_total = (normal_order['Total'].sum() if normal_count > 0 else 0) + (return_order['Total'].sum() if return_count > 0 else 0)

summary_data = {
    '訂單類型': ['Abnormal Order', 'Normal Order', 'Return Order'],
    'Count': [abnormal_count, normal_count, return_count],
    'Total': [
        abnormal_order['Total'].sum() if abnormal_count > 0 else 0,
        normal_order['Total'].sum() if normal_count > 0 else 0,
        return_order['Total'].sum() if return_count > 0 else 0
    ]
}
summary_df = pd.DataFrame(summary_data)

# 添加 Gross Order
gross_order_data = pd.DataFrame({
    '訂單類型': ['Gross Order'],
    'Count': [gross_order_count],
    'Total': [gross_order_total]
})

# 添加總計行
total_row = pd.DataFrame({
    '訂單類型': ['總計'],
    'Count': [summary_df['Count'].sum()],
    'Total': [summary_df['Total'].sum()]
})

# 合併所有資料
summary_df = pd.concat([summary_df, gross_order_data, total_row], ignore_index=True)

print("\n匯總表:")
print(summary_df)
print(f"\nGross Order = Normal Order + Return Order: Count={gross_order_count}, Total={gross_order_total:.2f}")

# 步驟6: 計算月度 KPI (Gross Revenue, Return, Revenue, Gross Orders, Return Orders, Normal Orders, Customer, Growth Rates)
print("\n=== 步驟6: 計算月度 KPI (Gross Revenue, Return, Revenue, Gross Orders, Return Orders, Normal Orders, Customer, Growth Rates) ===")
print("備註: Gross Revenue = -Return + Revenue, Gross Orders = -Return Orders + Normal Orders")

# 處理日期列
normal_order[date_col] = pd.to_datetime(normal_order[date_col], errors='coerce')
normal_order = normal_order.dropna(subset=[date_col])

# 處理 Return Order 的日期列
return_order[date_col] = pd.to_datetime(return_order[date_col], errors='coerce')
return_order = return_order.dropna(subset=[date_col])

# 添加年月列
normal_order['YearMonth'] = normal_order[date_col].dt.to_period('M')
return_order['YearMonth'] = return_order[date_col].dt.to_period('M')

# 按月分組計算 Revenue, Orders, Customer (從 Normal Order)
normal_monthly = normal_order.groupby('YearMonth').agg({
    'Total': ['sum', 'count'],  # Revenue 和 Orders
    customerid_col: 'nunique'   # Customers
}).reset_index()
normal_monthly.columns = ['YearMonth', 'Revenue', 'Normal_Orders', 'Customer']
normal_monthly['YearMonth'] = normal_monthly['YearMonth'].astype(str)

# 按月分組計算 Return Order KPI
if return_count > 0:
    return_monthly = return_order.groupby('YearMonth').agg({
        'Total': ['sum', 'count'],
        customerid_col: 'nunique'
    }).reset_index()
    return_monthly.columns = ['YearMonth', 'Return', 'Return_Orders', 'Return_Customers']
    return_monthly['YearMonth'] = return_monthly['YearMonth'].astype(str)
else:
    return_monthly = pd.DataFrame(columns=['YearMonth', 'Return', 'Return_Orders', 'Return_Customers'])

# 合併 Normal 和 Return 資料
monthly_kpis = normal_monthly.merge(return_monthly, on='YearMonth', how='left').fillna(0)

# 根據備註計算 Gross KPI: Gross Revenue = -Return + Revenue, Gross Orders = -Return Orders + Normal Orders
# 注意：Return 和 Return_Orders 在 Return Order 中通常是負數（因為 Quantity < 0），所以 -Return 會變成正數
monthly_kpis['Gross_Revenue'] = -monthly_kpis['Return'] + monthly_kpis['Revenue']
monthly_kpis['Gross_Orders'] = -monthly_kpis['Return_Orders'] + monthly_kpis['Normal_Orders']

# 注意：Customer 使用 Normal Order 的 Customer
monthly_kpis['Customer'] = monthly_kpis['Customer']  # 保持 Normal Order 的 Customer

# 排序並計算增長率
monthly_kpis = monthly_kpis.sort_values('YearMonth')

# Revenue, Orders, Customer Growth Rates (基於 Normal Order 的資料)
monthly_kpis['Revenue_Growth'] = monthly_kpis['Revenue'].pct_change() * 100
monthly_kpis['Orders_Growth'] = monthly_kpis['Normal_Orders'].pct_change() * 100
monthly_kpis['Customer_Growth'] = monthly_kpis['Customer'].pct_change() * 100

# Gross Growth Rates
monthly_kpis['Gross_Revenue_Growth'] = monthly_kpis['Gross_Revenue'].pct_change() * 100
monthly_kpis['Gross_Orders_Growth'] = monthly_kpis['Gross_Orders'].pct_change() * 100

# 填充第一行的增長率為 0
growth_cols = ['Revenue_Growth', 'Orders_Growth', 'Customer_Growth',
               'Gross_Revenue_Growth', 'Gross_Orders_Growth']
for col in growth_cols:
    monthly_kpis[col] = monthly_kpis[col].fillna(0).round(2)

# 格式化數值
monthly_kpis['Gross_Revenue'] = monthly_kpis['Gross_Revenue'].round(2)
monthly_kpis['Revenue'] = monthly_kpis['Revenue'].round(2)
monthly_kpis['Return'] = monthly_kpis['Return'].round(2)

# 選擇需要的列（按順序：Gross Revenue, Return, Revenue, Gross Orders, Return Orders, Normal Orders, Customer, Growth Rates）
monthly_kpis_df = monthly_kpis[['YearMonth', 
                                 'Gross_Revenue', 'Return', 'Revenue',
                                 'Gross_Orders', 'Return_Orders', 'Normal_Orders',
                                 'Customer',
                                 'Gross_Revenue_Growth', 'Gross_Orders_Growth',
                                 'Revenue_Growth', 'Orders_Growth', 'Customer_Growth']].copy()

# 重命名 Orders_Growth 為 Normal_Orders_Growth 以保持一致性
monthly_kpis_df = monthly_kpis_df.rename(columns={'Orders_Growth': 'Normal_Orders_Growth'})

print("\n月度 KPI 預覽:")
print(monthly_kpis_df.head(10))

# 步驟7: 計算 AOV, ARPU, Growth (只使用 Revenue, Orders, Customer，不使用 Gross)
print("\n=== 步驟7: 計算 AOV, ARPU, Growth (使用 Revenue, Normal_Orders, Customer，不使用 Gross) ===")

# AOV = Average Order Value = Revenue / Orders
# ARPU = Average Revenue Per User = Revenue / Customers
aov_arpu_df = monthly_kpis_df[['YearMonth', 'Revenue', 'Normal_Orders', 'Customer']].copy()

# 避免除零錯誤
aov_arpu_df['AOV'] = (aov_arpu_df['Revenue'] / aov_arpu_df['Normal_Orders'].replace(0, np.nan)).round(2)
aov_arpu_df['ARPU'] = (aov_arpu_df['Revenue'] / aov_arpu_df['Customer'].replace(0, np.nan)).round(2)

# 計算增長率
aov_arpu_df['AOV_Growth'] = aov_arpu_df['AOV'].pct_change() * 100
aov_arpu_df['ARPU_Growth'] = aov_arpu_df['ARPU'].pct_change() * 100
aov_arpu_df['AOV_Growth'] = aov_arpu_df['AOV_Growth'].fillna(0).round(2)
aov_arpu_df['ARPU_Growth'] = aov_arpu_df['ARPU_Growth'].fillna(0).round(2)

# 選擇需要的列
aov_arpu_df = aov_arpu_df[['YearMonth', 'AOV', 'ARPU', 'AOV_Growth', 'ARPU_Growth']].copy()

print("\nAOV & ARPU KPI 預覽 (基於 Revenue, Orders, Customer):")
print(aov_arpu_df.head(10))

# 步驟8: 計算產品 KPI
print("\n=== 步驟8: 計算產品 KPI ===")

# 查找產品相關列
stockcode_col = find_column(normal_order, ['StockCode', 'Stock Code', 'SKU', 'sku', '產品代碼'])
country_col = find_column(normal_order, ['Country', 'country', '國家', '國家'])

print(f"StockCode 列: {stockcode_col}")
print(f"Country 列: {country_col}")

product_kpis = []

# 1. Top SKUs (按 Revenue 排序)
if stockcode_col:
    top_skus = normal_order.groupby(stockcode_col).agg({
        'Total': 'sum',
        quantity_col: 'sum'
    }).reset_index()
    top_skus.columns = ['SKU', 'Revenue', 'Quantity']
    top_skus = top_skus.sort_values('Revenue', ascending=False).head(20)
    top_skus['Rank'] = range(1, len(top_skus) + 1)
    top_skus = top_skus[['Rank', 'SKU', 'Revenue', 'Quantity']]
    
    print(f"\nTop 20 SKUs (按 Revenue):")
    print(top_skus.head(10))
    
    product_kpis.append(('Top SKUs', top_skus))

# 2. SKU Diversity (每月不同的 SKU 數量)
if stockcode_col:
    monthly_sku_diversity = normal_order.groupby('YearMonth')[stockcode_col].nunique().reset_index()
    monthly_sku_diversity.columns = ['YearMonth', 'SKU_Count']
    monthly_sku_diversity['YearMonth'] = monthly_sku_diversity['YearMonth'].astype(str)
    
    print(f"\n月度 SKU Diversity:")
    print(monthly_sku_diversity.head(10))
    
    product_kpis.append(('SKU Diversity', monthly_sku_diversity))

# 3. Sales by Country
if country_col:
    sales_by_country = normal_order.groupby(country_col).agg({
        'Total': ['sum', 'count'],
        customerid_col: 'nunique'
    }).reset_index()
    sales_by_country.columns = ['Country', 'Revenue', 'Orders', 'Customers']
    sales_by_country = sales_by_country.sort_values('Revenue', ascending=False)
    sales_by_country['Revenue'] = sales_by_country['Revenue'].round(2)
    
    print(f"\nSales by Country (Top 20):")
    print(sales_by_country.head(20))
    
    product_kpis.append(('Sales by Country', sales_by_country))

# 步驟9: 計算 RFM 分析
print("\n=== 步驟9: 計算 RFM 分析 ===")

# 查找 Invoice 列
invoice_col = find_column(normal_order, ['InvoiceNo', 'Invoice No', 'Invoice', 'InvoiceNumber', '發票號碼', '發票'])

if not invoice_col:
    print("警告: 找不到 Invoice 列，跳過 RFM 分析")
    rfm_df = pd.DataFrame()
else:
    print(f"Invoice 列: {invoice_col}")
    
    # 確保日期列已處理
    if date_col not in normal_order.columns or not pd.api.types.is_datetime64_any_dtype(normal_order[date_col]):
        normal_order[date_col] = pd.to_datetime(normal_order[date_col], errors='coerce')
    
    # 移除日期為空的記錄
    rfm_data = normal_order[[customerid_col, invoice_col, date_col, 'Total']].copy()
    rfm_data = rfm_data.dropna(subset=[customerid_col, date_col])
    
    # 計算報告最後一天（數據中的最大日期）
    last_report_date = rfm_data[date_col].max()
    print(f"報告最後日期: {last_report_date.date()}")
    
    # 按客戶計算 RFM
    rfm_calc = rfm_data.groupby(customerid_col).agg({
        date_col: 'max',              # 最後購買日
        invoice_col: 'nunique',       # Invoice 數量（唯一值）
        'Total': 'sum'                # 總金額
    }).reset_index()
    
    rfm_calc.columns = ['CustomerID', 'LastPurchaseDate', 'Frequency', 'Monetary']
    
    # 計算 Recency（距離最後報告日的天數）
    rfm_calc['Recency'] = (last_report_date - rfm_calc['LastPurchaseDate']).dt.days
    
    # 使用 qcut 給每個維度打分（1-5）
    # Frequency: 越高越好（保持原樣）
    # Monetary: 越高越好（保持原樣）
    # Recency: 越低越好（需要反轉）
    
    try:
        # Frequency 分數（1-5，越高越好）
        rfm_calc['F_Score'] = pd.qcut(rfm_calc['Frequency'], q=5, labels=[1, 2, 3, 4, 5], duplicates='drop').astype(int)
    except:
        # 如果無法分5組，使用 rank
        rfm_calc['F_Score'] = pd.qcut(rfm_calc['Frequency'].rank(method='first'), q=5, labels=[1, 2, 3, 4, 5], duplicates='drop').astype(int)
    
    try:
        # Monetary 分數（1-5，越高越好）
        rfm_calc['M_Score'] = pd.qcut(rfm_calc['Monetary'], q=5, labels=[1, 2, 3, 4, 5], duplicates='drop').astype(int)
    except:
        rfm_calc['M_Score'] = pd.qcut(rfm_calc['Monetary'].rank(method='first'), q=5, labels=[1, 2, 3, 4, 5], duplicates='drop').astype(int)
    
    try:
        # Recency 分數（1-5，但需要反轉：天數越少分數越高）
        # qcut 按天數從小到大分組，labels=[1,2,3,4,5]（天數少的在前得到1）
        r_rank = pd.qcut(rfm_calc['Recency'], q=5, labels=[1, 2, 3, 4, 5], duplicates='drop').astype(int)
        # 反轉：天數越少分數越高（6 - rank：天數少的1變成5，天數多的5變成1）
        rfm_calc['R_Score'] = 6 - r_rank
    except:
        # 如果無法分5組，使用 rank
        r_rank = pd.qcut(rfm_calc['Recency'].rank(method='first'), q=5, labels=[1, 2, 3, 4, 5], duplicates='drop').astype(int)
        rfm_calc['R_Score'] = 6 - r_rank
    
    # 計算總分
    rfm_calc['Total_Score'] = rfm_calc['R_Score'] + rfm_calc['F_Score'] + rfm_calc['M_Score']
    
    # 根據總分分類
    def categorize_rfm(score):
        if 13 <= score <= 15:
            return 'Champions'
        elif 10 <= score <= 12:
            return 'Loyal'
        elif 7 <= score <= 9:
            return 'Potential Loyalist'
        elif 4 <= score <= 6:
            return 'At Risk'
        elif 1 <= score <= 3:
            return 'Lost'
        else:
            return 'Unknown'
    
    rfm_calc['Category'] = rfm_calc['Total_Score'].apply(categorize_rfm)
    
    # 選擇需要的列並排序
    rfm_df = rfm_calc[['CustomerID', 'Recency', 'Frequency', 'Monetary', 
                       'R_Score', 'F_Score', 'M_Score', 'Total_Score', 'Category']].copy()
    rfm_df = rfm_df.sort_values('Total_Score', ascending=False)
    
    # 格式化數值
    rfm_df['Monetary'] = rfm_df['Monetary'].round(2)
    
    print(f"\nRFM 分析結果:")
    print(f"客戶數量: {len(rfm_df)}")
    print(f"\n各類別客戶數量:")
    print(rfm_df['Category'].value_counts().sort_index())
    print(f"\nRFM 分析預覽:")
    print(rfm_df.head(10))

# 將所有資料寫入 彙總表.xlsx
print("\n=== 寫入 彙總表.xlsx ===")

with pd.ExcelWriter('彙總表.xlsx', engine='openpyxl') as writer:
    # 寫入 Abnormal Order
    if abnormal_count > 0:
        abnormal_order.to_excel(writer, sheet_name='Abnormal Order', index=False)
        print(f"已寫入 Abnormal Order 工作表 ({abnormal_count} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='Abnormal Order', index=False)
        print("已寫入空的 Abnormal Order 工作表")
    
    # 寫入 Return Order
    if return_count > 0:
        return_order.to_excel(writer, sheet_name='return order', index=False)
        print(f"已寫入 return order 工作表 ({return_count} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='return order', index=False)
        print("已寫入空的 return order 工作表")
    
    # 寫入 Normal Order
    if normal_count > 0:
        normal_order.to_excel(writer, sheet_name='Normal Order', index=False)
        print(f"已寫入 Normal Order 工作表 ({normal_count} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='Normal Order', index=False)
        print("已寫入空的 Normal Order 工作表")
    
    # 寫入匯總表
    # 先寫入步驟5的匯總統計
    summary_df.to_excel(writer, sheet_name='彙總', index=False, startrow=0)
    
    # 在下面添加月度 KPI（留空行）
    start_row = len(summary_df) + 3
    monthly_kpis_df.to_excel(writer, sheet_name='彙總', index=False, startrow=start_row)
    
    # 在下面添加 AOV & ARPU KPI
    start_row = start_row + len(monthly_kpis_df) + 3
    aov_arpu_df.to_excel(writer, sheet_name='彙總', index=False, startrow=start_row)
    
    # 在下面添加產品 KPI
    current_row = start_row + len(aov_arpu_df) + 3
    worksheet = writer.sheets['彙總']
    for kpi_name, kpi_df in product_kpis:
        if len(kpi_df) > 0:
            # 添加標題
            worksheet.cell(row=current_row, column=1, value=kpi_name)
            current_row += 1
            # 寫入資料（包含列名）
            for c_idx, col_name in enumerate(kpi_df.columns, start=1):
                worksheet.cell(row=current_row, column=c_idx, value=col_name)
            current_row += 1
            # 寫入資料行
            for row_idx, row_data in enumerate(kpi_df.values):
                for c_idx, value in enumerate(row_data, start=1):
                    worksheet.cell(row=current_row + row_idx, column=c_idx, value=value)
            current_row += len(kpi_df) + 3
    
    # 設置列寬
    worksheet = writer.sheets['彙總']
    for column in worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # 寫入 RFM 分析
    if 'rfm_df' in locals() and len(rfm_df) > 0:
        rfm_df.to_excel(writer, sheet_name='RFM', index=False)
        print(f"已寫入 RFM 工作表 ({len(rfm_df)} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='RFM', index=False)
        print("已寫入空的 RFM 工作表")
    
    # 在其他工作表也設置列寬
    for sheet_name in writer.sheets:
        if sheet_name != '彙總':
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"\n已保存所有資料到 彙總表.xlsx")
print(f"\n工作表包含:")
print(f"  - Abnormal Order 工作表")
print(f"  - return order 工作表")
print(f"  - Normal Order 工作表")
print(f"  - 彙總工作表:")
print(f"    * 訂單類型匯總統計 (Count 和 Total, 包含 Gross Order = Normal Order + Return Order)")
print(f"    * 月度 KPI (Gross Revenue, Return, Revenue, Gross Orders, Return Orders, Normal Orders, Customer, Growth Rates)")
print(f"      備註: Gross Revenue = -Return + Revenue, Gross Orders = -Return Orders + Normal Orders")
print(f"    * AOV & ARPU KPI (基於 Revenue, Normal Orders, Customer: AOV, ARPU, Growth)")
print(f"    * 產品 KPI (Top SKUs, SKU Diversity, Sales by Country)")
print(f"  - RFM 工作表 (Recency, Frequency, Monetary 分析)")

print("\n=== 完成 ===")
print(f"原始資料行數: {len(pd.read_excel('online_retail_II.xlsx'))}")
print(f"異常訂單: {abnormal_count} 行")
print(f"退貨訂單: {return_count} 行")
print(f"正常訂單: {normal_count} 行")
print(f"總計: {abnormal_count + return_count + normal_count} 行")

