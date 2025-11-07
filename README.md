# Online Retail 數據分析項目

## 項目簡介

本項目是一個完整的在線零售數據分析系統，用於處理和分析在線零售交易數據。主要功能包括：

1. **數據清理與分類**：將原始數據分類為正常訂單、退貨訂單和異常訂單
2. **KPI 計算**：計算月度關鍵績效指標（收入、訂單、客戶等）
3. **RFM 分析**：基於 Recency、Frequency、Monetary 進行客戶價值分析
4. **退貨分析**：分析產品、客戶和國家的退貨率
5. **異常訂單分析**：識別和統計數據質量問題

## 項目結構

```
Online Retail/
├── execute_prompt.py              # 主分析腳本（數據清理 + KPI + RFM）
├── execute_return_abnormal.py     # 退貨和異常訂單分析腳本
├── online_retail_II.xlsx          # 原始數據文件（輸入）
├── 彙總表.xlsx                    # 主要輸出文件（數據清理 + KPI + RFM）
├── Return and Abnormal.xlsx       # 退貨和異常分析輸出文件
├── prompt - Return,Abnormal.txt   # 退貨分析需求說明
├── prompt - KPI, RFM.txt          # KPI 和 RFM 需求說明
├── prompt - visualization.txt     # 可視化需求說明
└── README.md                      # 本文件
```

## 功能說明

### 1. execute_prompt.py - 主分析腳本

執行完整的數據清理、KPI 計算和 RFM 分析。

#### 主要功能：

1. **數據清理**：
   - 計算 Total = Quantity × Price
   - 篩除缺失 CustomerID 或 Description 的數據（Abnormal Order）
   - 篩除 Quantity < 0 的數據（Return Order）
   - 保存正常訂單（Normal Order）

2. **匯總統計**：
   - 計算各類訂單的 Count 和 Total
   - 計算 Gross Order = Normal Order + Return Order

3. **月度 KPI 計算**：
   - Gross Revenue = -Return + Revenue
   - Gross Orders = -Return Orders + Normal Orders
   - Revenue, Return, Return Orders, Normal Orders
   - Customer（從 Normal Order）
   - 各項增長率（Revenue_Growth, Orders_Growth, Customer_Growth 等）

4. **AOV & ARPU 計算**：
   - AOV (Average Order Value) = Revenue / Normal Orders
   - ARPU (Average Revenue Per User) = Revenue / Customer
   - 增長率計算

5. **產品 KPI**：
   - Top SKUs（按 Revenue 排序，Top 20）
   - SKU Diversity（每月不同的 SKU 數量）
   - Sales by Country（按國家統計）

6. **RFM 分析**：
   - Recency：距離最後報告日的天數
   - Frequency：唯一 Invoice 數量
   - Monetary：總消費金額
   - 客戶分類：Champions, Loyal, Potential Loyalist, At Risk, Lost

#### 輸出文件：

- `彙總表.xlsx`：包含以下工作表
  - `Abnormal Order`：異常訂單數據
  - `return order`：退貨訂單數據
  - `Normal Order`：正常訂單數據
  - `彙總`：包含所有 KPI 數據的匯總表
  - `RFM`：RFM 客戶價值分析結果

### 2. execute_return_abnormal.py - 退貨和異常分析腳本

執行退貨率和異常訂單的詳細分析。

#### 主要功能：

1. **產品退貨率分析**（按 StockCode）：
   - 計算 Return_Rate 和 Return_Frequency
   - 分類標籤：High-return items, Medium-return items, Low-return items, 100% return items(outlier)

2. **客戶退貨率分析**（按 CustomerID）：
   - 計算 Return_Rate 和 Return_Frequency
   - 分類標籤：High-return customer, Medium-return customer, Low-return customer, 100% return customer(outlier)

3. **國家退貨率分析**（按 Country）：
   - 計算 Return_Rate 和 Return_Frequency

4. **異常訂單分析**：
   - 統計缺失 CustomerID 和缺失 Description 的數量和比例
   - 按 StockCode 和 Country 分組統計

5. **見解生成**：
   - 列出所有 High-return items
   - 列出所有 High-return customer
   - Top 10 國家（按 Return_Rate）
   - Top 10 產品（按缺失計數）

#### 輸出文件：

- `Return and Abnormal.xlsx`：包含以下工作表
  - `Return analysis product`：產品退貨率分析
  - `Return analysis customer`：客戶退貨率分析
  - `Return analysis country`：國家退貨率分析
  - `Abnormal analysis product`：異常訂單分析
  - `insights`：見解列表

## 使用方法

### 環境要求

- Python 3.7+
- 必需的 Python 包：
  - pandas
  - numpy
  - openpyxl

### 安裝依賴

```bash
pip install pandas numpy openpyxl
```

### 執行步驟

#### 步驟 1：執行主分析腳本

```bash
python execute_prompt.py
```

這將：
- 讀取 `online_retail_II.xlsx`
- 執行數據清理和 KPI 計算
- 執行 RFM 分析
- 生成 `彙總表.xlsx`

#### 步驟 2：執行退貨和異常分析腳本

```bash
python execute_return_abnormal.py
```

這將：
- 從 `彙總表.xlsx` 讀取數據
- 執行退貨率和異常訂單分析
- 生成 `Return and Abnormal.xlsx`

**注意**：步驟 2 需要在步驟 1 之後執行，因為它依賴於 `彙總表.xlsx` 的輸出。

## 關鍵計算公式

### 退貨率計算

- **Return_Rate** = -Return_Amount / (-Return_Amount + Revenue)
  - 表示退貨金額佔總交易金額的比例（0-1）
- **Return_Frequency** = Return_Count / (Return_Count + Normal_Count)
  - 表示退貨訂單數佔總訂單數的比例（0-1）

### 月度 KPI

- **Gross Revenue** = -Return + Revenue
  - 由於 Return 為負數，實際是 Revenue + |Return|
- **Gross Orders** = -Return Orders + Normal Orders
  - 實際是 Normal Orders - Return Orders

### AOV & ARPU

- **AOV** = Revenue / Normal Orders
- **ARPU** = Revenue / Customer

### RFM 評分

- **Recency**：距離最後報告日的天數（越少越好，評分反轉）
- **Frequency**：唯一 Invoice 數量（越多越好）
- **Monetary**：總消費金額（越多越好）
- **Total_Score** = R_Score + F_Score + M_Score（3-15分）

### RFM 客戶分類

- **Champions**：13-15分（高價值活躍客戶）
- **Loyal**：10-12分（忠誠客戶）
- **Potential Loyalist**：7-9分（潛在忠誠客戶）
- **At Risk**：4-6分（風險客戶）
- **Lost**：1-3分（流失客戶）

## 分類標籤說明

### 產品退貨率分類

- **High-return items**：Return_Rate > 0.7 且 Return_Count > 5
- **Medium-return items**：0.3 < Return_Rate <= 0.7
- **Low-return items**：Return_Rate <= 0.3
- **100% return items(outlier)**：Return_Count <= 5

### 客戶退貨率分類

- **High-return customer**：Return_Rate > 0.7 且 Return_Count > 5
- **Medium-return customer**：0.3 < Return_Rate <= 0.7
- **Low-return customer**：Return_Rate <= 0.3
- **100% return customer(outlier)**：Return_Count <= 5

## 數據流程

```
原始數據 (online_retail_II.xlsx)
  ↓
execute_prompt.py
  ↓
數據清理與分類
  ├─ Abnormal Order（異常訂單）
  ├─ Return Order（退貨訂單）
  └─ Normal Order（正常訂單）
  ↓
KPI 計算
  ├─ 月度 KPI（收入、訂單、客戶、增長率）
  ├─ AOV & ARPU
  └─ 產品 KPI（Top SKUs, SKU Diversity, Sales by Country）
  ↓
RFM 分析
  ↓
彙總表.xlsx
  ↓
execute_return_abnormal.py
  ↓
退貨和異常分析
  ├─ 產品退貨率分析
  ├─ 客戶退貨率分析
  ├─ 國家退貨率分析
  ├─ 異常訂單分析
  └─ 見解生成
  ↓
Return and Abnormal.xlsx
```

## 特性

1. **自動列名識別**：支持不同大小寫、空格變化的列名
2. **數據質量檢查**：自動識別和分類異常數據
3. **全面的 KPI 計算**：包含收入、訂單、客戶等多個維度的指標
4. **RFM 客戶價值分析**：幫助識別不同價值的客戶群體
5. **退貨率分析**：詳細分析產品、客戶和國家的退貨情況
6. **異常訂單分析**：識別數據質量問題

## 注意事項

1. **執行順序**：必須先執行 `execute_prompt.py`，再執行 `execute_return_abnormal.py`
2. **輸入文件**：確保 `online_retail_II.xlsx` 文件存在且格式正確
3. **列名要求**：數據文件應包含以下列（支持不同命名）：
   - Quantity（數量）
   - Price（價格）
   - CustomerID（客戶ID）
   - Description（描述）
   - InvoiceNo（發票號）
   - InvoiceDate（發票日期）
   - StockCode（產品代碼）
   - Country（國家）
4. **數據格式**：
   - Quantity < 0 的記錄會被識別為退貨訂單
   - 缺失 CustomerID 或 Description 的記錄會被識別為異常訂單

## 輸出文件說明

### 彙總表.xlsx

- **Abnormal Order**：異常訂單數據（缺失關鍵字段）
- **return order**：退貨訂單數據（Quantity < 0）
- **Normal Order**：正常訂單數據
- **彙總**：包含所有 KPI 數據的匯總表
  - 訂單類型匯總統計
  - 月度 KPI（Gross Revenue, Return, Revenue, Gross Orders, Return Orders, Normal Orders, Customer, Growth Rates）
  - AOV & ARPU KPI
  - 產品 KPI（Top SKUs, SKU Diversity, Sales by Country）
- **RFM**：RFM 客戶價值分析結果

### Return and Abnormal.xlsx

- **Return analysis product**：產品退貨率分析（包含分類標籤）
- **Return analysis customer**：客戶退貨率分析（包含分類標籤）
- **Return analysis country**：國家退貨率分析
- **Abnormal analysis product**：異常訂單分析
- **insights**：見解列表（High-return items/customer, Top 10 國家/產品）

## 技術細節

### 使用的 Python 庫

- **pandas**：數據處理和分析
- **numpy**：數值計算
- **openpyxl**：Excel 文件讀寫

### 關鍵技術

1. **自動列名識別**：使用模糊匹配識別列名
2. **數據分組與聚合**：使用 pandas groupby 進行統計分析
3. **分位數分組**：使用 pd.qcut 進行 RFM 評分
4. **Excel 多工作表操作**：使用 openpyxl 引擎進行多工作表寫入

## 項目歷史

- **2024-11-02**：初始版本，實現數據清理和分類
- **2024-11-04**：添加月度 KPI、AOV & ARPU、產品 KPI 計算
- **2024-11-06**：添加 RFM 分析、退貨和異常訂單分析

## 維護者

本項目由數據分析團隊維護。

## 許可證

本項目僅供內部使用。

---

**最後更新**：2024-11-06

