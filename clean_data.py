import pandas as pd
import numpy as np

# 1. โหลดข้อมูล
df_trans = pd.read_csv('transactions_data.csv')
df_cust = pd.read_csv('customers_data.csv')

print(f"Original Transactions Shape: {df_trans.shape}")

# 2. แปลงรูปแบบวันที่ (Data Formatting)
df_trans['Transaction_Date'] = pd.to_datetime(df_trans['Transaction_Date'])
df_cust['Signup_Date'] = pd.to_datetime(df_cust['Signup_Date'])

# 3. จัดการ Missing Values (เติม Median ใน Units_Sold และคำนวณ Total_Amount ใหม่)
median_units = df_trans['Units_Sold'].median()
df_trans['Units_Sold'] = df_trans['Units_Sold'].fillna(median_units)
df_trans['Total_Amount'] = df_trans['Units_Sold'] * df_trans['Price_Per_Unit']

# 4. จัดการ Outliers ด้วยเกณฑ์ IQR สำหรับ Units_Sold
Q1 = df_trans['Units_Sold'].quantile(0.25)
Q3 = df_trans['Units_Sold'].quantile(0.75)
IQR = Q3 - Q1
lower_bound = Q1 - 1.5 * IQR
upper_bound = Q3 + 1.5 * IQR

# กรองเฉพาะข้อมูลที่อยู่ในช่วงปกติ
df_clean = df_trans[(df_trans['Units_Sold'] >= lower_bound) & (df_trans['Units_Sold'] <= upper_bound)].copy()

# คำนวณ Total_Amount อีกรอบให้แม่นยำหลังกรอง Outlier
df_clean['Total_Amount'] = df_clean['Units_Sold'] * df_clean['Price_Per_Unit']

print(f"Cleaned Transactions Shape: {df_clean.shape}")

# 5. บันทึกไฟล์ข้อมูลที่ทำความสะอาดแล้ว
df_clean.to_csv('cleaned_transactions.csv', index=False)
print("Data Cleaning completed and saved successfully!")