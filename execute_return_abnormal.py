import pandas as pd
import numpy as np

# 1.1 從彙總表載入 Normal Order/Abnormal Order/Return Order tabs
print("=== 步驟1.1: 從彙總表載入資料 ===")

try:
    normal_order = pd.read_excel('彙總表.xlsx', sheet_name='Normal Order')
    abnormal_order = pd.read_excel('彙總表.xlsx', sheet_name='Abnormal Order')
    return_order = pd.read_excel('彙總表.xlsx', sheet_name='return order')
    
    print(f"Normal Order: {len(normal_order)} 行")
    print(f"Abnormal Order: {len(abnormal_order)} 行")
    print(f"Return Order: {len(return_order)} 行")
except Exception as e:
    print(f"錯誤: 無法讀取 彙總表.xlsx，請先運行 execute_prompt.py")
    print(f"錯誤詳情: {e}")
    raise

# 標準化列名
normal_order.columns = normal_order.columns.str.strip()
abnormal_order.columns = abnormal_order.columns.str.strip()
return_order.columns = return_order.columns.str.strip()

# 查找相關列名
def find_column(df, keywords):
    """查找包含關鍵詞的列名"""
    keywords_lower = [k.lower() for k in keywords]
    for col in df.columns:
        col_lower = col.lower()
        for keyword in keywords_lower:
            if keyword in col_lower:
                return col
    return None

stockcode_col = find_column(normal_order, ['StockCode', 'Stock Code', 'SKU', 'sku', '產品代碼'])
country_col = find_column(normal_order, ['Country', 'country', '國家', '國家'])
customerid_col = find_column(normal_order, ['Customer ID', 'CustomerID', 'Customer', 'customer'])
quantity_col = find_column(normal_order, ['Quantity', 'quantity', '數量'])
description_col = find_column(normal_order, ['Description', 'description', '描述'])

print(f"\n識別的列名:")
print(f"StockCode: {stockcode_col}")
print(f"Country: {country_col}")
print(f"CustomerID: {customerid_col}")
print(f"Quantity: {quantity_col}")
print(f"Description: {description_col}")

# 2.1 在 Return Order 數據中計算（按 StockCode）
print("\n=== 步驟2.1: 計算 Return Order 分析（按 StockCode） ===")

if len(return_order) > 0 and stockcode_col:
    # 計算每個 StockCode 的總退貨金額和退貨數量
    return_by_stockcode = return_order.groupby(stockcode_col).agg({
        'Total': 'sum',
        quantity_col: 'count'  # Return_Count 是記錄數
    }).reset_index()
    return_by_stockcode.columns = ['StockCode', 'Return_Amount', 'Return_Count']
    
    print(f"Return Order 分析完成: {len(return_by_stockcode)} 個 StockCode")
else:
    return_by_stockcode = pd.DataFrame(columns=['StockCode', 'Return_Amount', 'Return_Count'])
    print("Return Order 數據為空或缺少 StockCode 列")

# 2.2 合併 Normal Order 計算產品退貨率（按 StockCode）
print("\n=== 步驟2.2: 計算產品退貨率（按 StockCode） ===")

if len(return_by_stockcode) > 0 and len(normal_order) > 0 and stockcode_col:
    # 按 StockCode 計算 Normal 數據
    normal_by_stockcode = normal_order.groupby(stockcode_col).agg({
        'Total': 'sum',
        quantity_col: 'count'
    }).reset_index()
    normal_by_stockcode.columns = ['StockCode', 'Revenue', 'Normal_Count']
    
    # 合併計算退貨率
    product_analysis = return_by_stockcode.merge(normal_by_stockcode, on='StockCode', how='outer').fillna(0)
    
    # 計算 Return_Rate: -Return_Amount / (-Return_Amount + Revenue)
    # 注意：Return_Amount 在 Return Order 中是負數，所以 -Return_Amount 是正數
    product_analysis['Return_Rate'] = ((-product_analysis['Return_Amount']) / 
                                       ((-product_analysis['Return_Amount']) + product_analysis['Revenue']).replace(0, np.nan)).fillna(0)
    
    # 計算 Return_Frequency: Return_Count / (Return_Count + Normal_Count)
    total_count = product_analysis['Return_Count'] + product_analysis['Normal_Count']
    product_analysis['Return_Frequency'] = (product_analysis['Return_Count'] / 
                                           total_count.replace(0, np.nan)).fillna(0)
    
    # 根據不同情況給標籤
    # 注意：outlier 條件優先（Return_Count <= 5）
    def categorize_product_return(row):
        if row['Return_Count'] <= 5:
            return '100% return items(outlier)'
        elif row['Return_Rate'] > 0.7 and row['Return_Count'] > 5:
            return 'High-return items'
        elif 0.3 < row['Return_Rate'] <= 0.7:
            return 'Medium-return items'
        elif row['Return_Rate'] <= 0.3:
            return 'Low-return items'
        else:
            return 'Unknown'
    
    product_analysis['Category'] = product_analysis.apply(categorize_product_return, axis=1)
    
    # 格式化
    product_analysis['Return_Rate'] = product_analysis['Return_Rate'].round(4)
    product_analysis['Return_Frequency'] = product_analysis['Return_Frequency'].round(4)
    product_analysis['Return_Amount'] = product_analysis['Return_Amount'].round(2)
    product_analysis['Revenue'] = product_analysis['Revenue'].round(2)
    product_analysis = product_analysis.sort_values('Return_Rate', ascending=False)
    
    print(f"產品退貨率分析完成: {len(product_analysis)} 個產品")
    print(f"分類統計:")
    print(product_analysis['Category'].value_counts())
else:
    product_analysis = pd.DataFrame()
    print("無法計算產品退貨率（數據為空或缺少必要列）")

# 3.1 在 Return Order 數據中計算（按 CustomerID）
print("\n=== 步驟3.1: 計算 Return Order 分析（按 CustomerID） ===")

if len(return_order) > 0 and customerid_col:
    # 計算每個 CustomerID 的總退貨金額和退貨數量
    return_by_customer = return_order.groupby(customerid_col).agg({
        'Total': 'sum',
        quantity_col: 'count'  # Return_Count 是記錄數
    }).reset_index()
    return_by_customer.columns = [customerid_col, 'Return_Amount', 'Return_Count']
    
    print(f"Return Order 分析完成: {len(return_by_customer)} 個客戶")
else:
    return_by_customer = pd.DataFrame()
    print("Return Order 數據為空或缺少 CustomerID 列")

# 3.2 合併 Normal Order 計算客戶退貨率（按 CustomerID）
print("\n=== 步驟3.2: 計算客戶退貨率（按 CustomerID） ===")

if len(return_by_customer) > 0 and len(normal_order) > 0 and customerid_col:
    # 按客戶計算 Normal 數據
    normal_by_customer = normal_order.groupby(customerid_col).agg({
        'Total': 'sum',
        quantity_col: 'count'
    }).reset_index()
    normal_by_customer.columns = [customerid_col, 'Revenue', 'Normal_Count']
    
    # 合併計算退貨率
    customer_analysis = return_by_customer.merge(normal_by_customer, on=customerid_col, how='outer').fillna(0)
    
    # 計算 Return_Rate: -Return_Amount / (-Return_Amount + Revenue)
    # 注意：Return_Amount 在 Return Order 中是負數，所以 -Return_Amount 是正數
    customer_analysis['Return_Rate'] = ((-customer_analysis['Return_Amount']) / 
                                       ((-customer_analysis['Return_Amount']) + customer_analysis['Revenue']).replace(0, np.nan)).fillna(0)
    
    # 計算 Return_Frequency: Return_Count / (Return_Count + Normal_Count)
    total_count = customer_analysis['Return_Count'] + customer_analysis['Normal_Count']
    customer_analysis['Return_Frequency'] = (customer_analysis['Return_Count'] / 
                                           total_count.replace(0, np.nan)).fillna(0)
    
    # 根據不同情況給標籤
    # 注意：outlier 條件優先（Return_Count <= 5）
    def categorize_customer_return(row):
        if row['Return_Count'] <= 5:
            return '100% return customer(outlier)'
        elif row['Return_Rate'] > 0.7 and row['Return_Count'] > 5:
            return 'High-return customer'
        elif 0.3 < row['Return_Rate'] <= 0.7:
            return 'Medium-return customer'
        elif row['Return_Rate'] <= 0.3:
            return 'Low-return customer'
        else:
            return 'Unknown'
    
    customer_analysis['Category'] = customer_analysis.apply(categorize_customer_return, axis=1)
    
    # 格式化
    customer_analysis['Return_Rate'] = customer_analysis['Return_Rate'].round(4)
    customer_analysis['Return_Frequency'] = customer_analysis['Return_Frequency'].round(4)
    customer_analysis['Return_Amount'] = customer_analysis['Return_Amount'].round(2)
    customer_analysis['Revenue'] = customer_analysis['Revenue'].round(2)
    customer_analysis = customer_analysis.sort_values('Return_Rate', ascending=False)
    
    print(f"客戶退貨率分析完成: {len(customer_analysis)} 個客戶")
    print(f"分類統計:")
    print(customer_analysis['Category'].value_counts())
else:
    customer_analysis = pd.DataFrame()
    print("無法計算客戶退貨率（數據為空或缺少必要列）")

# 4.1 計算按國家的退貨率
print("\n=== 步驟4.1: 計算國家退貨率 ===")

if country_col and len(return_order) > 0 and len(normal_order) > 0:
    # 按國家計算 Return 數據
    return_by_country = return_order.groupby(country_col).agg({
        'Total': 'sum',
        quantity_col: 'count'
    }).reset_index()
    return_by_country.columns = ['Country', 'Return_Amount', 'Return_Count']
    
    # 按國家計算 Normal 數據
    normal_by_country = normal_order.groupby(country_col).agg({
        'Total': 'sum',
        quantity_col: 'count'
    }).reset_index()
    normal_by_country.columns = ['Country', 'Revenue', 'Normal_Count']
    
    # 合併計算退貨率
    country_analysis = return_by_country.merge(normal_by_country, on='Country', how='outer').fillna(0)
    
    # 計算 Return_Rate: -Return_Amount / (-Return_Amount + Revenue)
    country_analysis['Return_Rate'] = ((-country_analysis['Return_Amount']) / 
                                      ((-country_analysis['Return_Amount']) + country_analysis['Revenue']).replace(0, np.nan)).fillna(0)
    
    # 計算 Return_Frequency: Return_Count / (Return_Count + Normal_Count)
    total_count = country_analysis['Return_Count'] + country_analysis['Normal_Count']
    country_analysis['Return_Frequency'] = (country_analysis['Return_Count'] / 
                                          total_count.replace(0, np.nan)).fillna(0)
    
    # 格式化
    country_analysis['Return_Rate'] = country_analysis['Return_Rate'].round(4)
    country_analysis['Return_Frequency'] = country_analysis['Return_Frequency'].round(4)
    country_analysis['Return_Amount'] = country_analysis['Return_Amount'].round(2)
    country_analysis['Revenue'] = country_analysis['Revenue'].round(2)
    country_analysis = country_analysis.sort_values('Return_Rate', ascending=False)
    
    print(f"國家退貨率分析完成: {len(country_analysis)} 個國家")
else:
    country_analysis = pd.DataFrame(columns=['Country', 'Return_Amount', 'Return_Count', 
                                           'Revenue', 'Normal_Count', 'Return_Rate', 'Return_Frequency'])
    print("無法計算國家退貨率（數據為空或缺少必要列）")

# 5.1 在 Abnormal Order 數據中計算
print("\n=== 步驟5.1: 計算 Abnormal Order 分析 ===")

if len(abnormal_order) > 0:
    # 計算缺失 CustomerID 和缺失 Description 的數量
    missing_customerid = abnormal_order[customerid_col].isna().sum() if customerid_col else 0
    missing_description = abnormal_order[description_col].isna().sum() if description_col else 0
    total_abnormal = len(abnormal_order)
    
    # 計算比例
    missing_customerid_pct = (missing_customerid / total_abnormal * 100) if total_abnormal > 0 else 0
    missing_description_pct = (missing_description / total_abnormal * 100) if total_abnormal > 0 else 0
    
    print(f"缺失 CustomerID: {missing_customerid} ({missing_customerid_pct:.2f}%)")
    print(f"缺失 Description: {missing_description} ({missing_description_pct:.2f}%)")
    
    # 按 StockCode 和 Country 計算缺失計數和比例
    if stockcode_col and country_col:
        abnormal_by_product = abnormal_order.groupby([stockcode_col, country_col]).agg({
            customerid_col: lambda x: x.isna().sum() if customerid_col else 0,
            description_col: lambda x: x.isna().sum() if description_col else 0,
        }).reset_index()
        
        # 重命名列
        if customerid_col and description_col:
            abnormal_by_product.columns = [stockcode_col, country_col, 'Missing_CustomerID_Count', 'Missing_Description_Count']
            abnormal_by_product['Missing_Total_Count'] = (abnormal_by_product['Missing_CustomerID_Count'] + 
                                                          abnormal_by_product['Missing_Description_Count'])
        elif customerid_col:
            abnormal_by_product.columns = [stockcode_col, country_col, 'Missing_CustomerID_Count']
            abnormal_by_product['Missing_Total_Count'] = abnormal_by_product['Missing_CustomerID_Count']
        elif description_col:
            abnormal_by_product.columns = [stockcode_col, country_col, 'Missing_Description_Count']
            abnormal_by_product['Missing_Total_Count'] = abnormal_by_product['Missing_Description_Count']
        else:
            abnormal_by_product = pd.DataFrame()
        
        if len(abnormal_by_product) > 0:
            # 計算比例
            total_by_product = abnormal_order.groupby([stockcode_col, country_col]).size().reset_index(name='Total_Count')
            abnormal_by_product = abnormal_by_product.merge(total_by_product, on=[stockcode_col, country_col], how='left')
            
            if 'Missing_CustomerID_Count' in abnormal_by_product.columns:
                abnormal_by_product['Missing_CustomerID_Proportion'] = (
                    abnormal_by_product['Missing_CustomerID_Count'] / 
                    abnormal_by_product['Total_Count'].replace(0, np.nan) * 100
                ).fillna(0).round(2)
            
            if 'Missing_Description_Count' in abnormal_by_product.columns:
                abnormal_by_product['Missing_Description_Proportion'] = (
                    abnormal_by_product['Missing_Description_Count'] / 
                    abnormal_by_product['Total_Count'].replace(0, np.nan) * 100
                ).fillna(0).round(2)
            
            abnormal_by_product['Missing_Total_Proportion'] = (
                abnormal_by_product['Missing_Total_Count'] / 
                abnormal_by_product['Total_Count'].replace(0, np.nan) * 100
            ).fillna(0).round(2)
            
            abnormal_by_product = abnormal_by_product.sort_values('Missing_Total_Count', ascending=False)
        
        print(f"產品異常分析完成: {len(abnormal_by_product)} 個產品-國家組合")
    else:
        abnormal_by_product = pd.DataFrame()
        print("無法計算產品異常分析（缺少 StockCode 或 Country 列）")
else:
    abnormal_by_product = pd.DataFrame()
    print("Abnormal Order 數據為空")

# 6.1 生成見解
print("\n=== 步驟6.1: 生成見解 ===")

insights_list = []

# List "High-return items"
if len(product_analysis) > 0 and 'Category' in product_analysis.columns:
    high_return_items = product_analysis[product_analysis['Category'] == 'High-return items']
    if len(high_return_items) > 0:
        for rank, (idx, row) in enumerate(high_return_items.iterrows(), start=1):
            insights_list.append({
                'Type': 'High-return items',
                'Rank': rank,
                'Detail': row['StockCode'],
                'Return_Rate': row['Return_Rate'],
                'Return_Count': row['Return_Count'],
                'Return_Amount': row['Return_Amount']
            })
        print(f"High-return items: {len(high_return_items)} 個")

# List "High-return customer"
if len(customer_analysis) > 0 and 'Category' in customer_analysis.columns:
    high_return_customers = customer_analysis[customer_analysis['Category'] == 'High-return customer']
    if len(high_return_customers) > 0:
        for rank, (idx, row) in enumerate(high_return_customers.iterrows(), start=1):
            customer_id = row[customerid_col] if customerid_col else row.index[0]
            insights_list.append({
                'Type': 'High-return customer',
                'Rank': rank,
                'Detail': customer_id,
                'Return_Rate': row['Return_Rate'],
                'Return_Count': row['Return_Count'],
                'Return_Amount': row['Return_Amount']
            })
        print(f"High-return customer: {len(high_return_customers)} 個")

# List top 10 Country based on the Return_Rate (higher the top)
if len(country_analysis) > 0 and 'Return_Rate' in country_analysis.columns:
    top_countries = country_analysis.nlargest(10, 'Return_Rate')
    if len(top_countries) > 0:
        for rank, (idx, row) in enumerate(top_countries.iterrows(), start=1):
            insights_list.append({
                'Type': 'Top 10 Country by Return_Rate',
                'Rank': rank,
                'Detail': row['Country'],
                'Return_Rate': row['Return_Rate'],
                'Return_Count': row['Return_Count']
            })
        print(f"Top 10 國家（Return_Rate）: {len(top_countries)} 個")

# List top 10 Products based on the missing count (higher the top)
if len(abnormal_by_product) > 0 and 'Missing_Total_Count' in abnormal_by_product.columns:
    top_products_missing = abnormal_by_product.nlargest(10, 'Missing_Total_Count')
    if len(top_products_missing) > 0:
        for rank, (idx, row) in enumerate(top_products_missing.iterrows(), start=1):
            insights_list.append({
                'Type': 'Top 10 Product by Missing Count',
                'Rank': rank,
                'Detail': row[stockcode_col] if stockcode_col else row.index[0],
                'Missing_Total_Count': row['Missing_Total_Count'],
                'Missing_Total_Proportion': row.get('Missing_Total_Proportion', 0)
            })
        print(f"Top 10 產品（缺失計數）: {len(top_products_missing)} 個")

insights_df = pd.DataFrame(insights_list)

# 1.2 寫入 Excel
print("\n=== 步驟1.2: 寫入 Return and Abnormal.xlsx ===")

with pd.ExcelWriter('Return and Abnormal.xlsx', engine='openpyxl') as writer:
    # Return analysis product (步驟2.2的結果 - 產品分析)
    if len(product_analysis) > 0:
        product_analysis.to_excel(writer, sheet_name='Return analysis product', index=False)
        print(f"已寫入 Return analysis product 工作表 ({len(product_analysis)} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='Return analysis product', index=False)
        print("已寫入空的 Return analysis product 工作表")
    
    # Return analysis customer (步驟3.2的結果 - 客戶分析)
    if len(customer_analysis) > 0:
        customer_analysis.to_excel(writer, sheet_name='Return analysis customer', index=False)
        print(f"已寫入 Return analysis customer 工作表 ({len(customer_analysis)} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='Return analysis customer', index=False)
        print("已寫入空的 Return analysis customer 工作表")
    
    # Return analysis country (步驟4.1的結果)
    if len(country_analysis) > 0:
        country_analysis.to_excel(writer, sheet_name='Return analysis country', index=False)
        print(f"已寫入 Return analysis country 工作表 ({len(country_analysis)} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='Return analysis country', index=False)
        print("已寫入空的 Return analysis country 工作表")
    
    # Abnormal analysis product (步驟5.1的結果)
    if len(abnormal_by_product) > 0:
        abnormal_by_product.to_excel(writer, sheet_name='Abnormal analysis product', index=False)
        print(f"已寫入 Abnormal analysis product 工作表 ({len(abnormal_by_product)} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='Abnormal analysis product', index=False)
        print("已寫入空的 Abnormal analysis product 工作表")
    
    # Insights (步驟6.1的結果)
    if len(insights_df) > 0:
        insights_df.to_excel(writer, sheet_name='insights', index=False)
        print(f"已寫入 insights 工作表 ({len(insights_df)} 行)")
    else:
        pd.DataFrame().to_excel(writer, sheet_name='insights', index=False)
        print("已寫入空的 insights 工作表")
    
    # 設置列寬
    for sheet_name in writer.sheets:
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
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

print(f"\n已保存所有資料到 Return and Abnormal.xlsx")
print(f"\n工作表包含:")
print(f"  - Return analysis product (產品退貨率分析 - 步驟2.2，包含分類標籤)")
print(f"  - Return analysis customer (客戶退貨率分析 - 步驟3.2，包含分類標籤)")
print(f"  - Return analysis country (國家退貨率分析 - 步驟4.1)")
print(f"  - Abnormal analysis product (產品異常分析 - 步驟5.1)")
print(f"  - insights (見解：High-return items、High-return customer、Top 10 國家、Top 10 產品缺失 - 步驟6.1)")

print("\n=== 完成 ===")
