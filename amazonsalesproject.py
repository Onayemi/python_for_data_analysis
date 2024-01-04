# -*- coding: utf-8 -*-
"""
Created on Thu Jan  4 13:34:48 2024

@author: REMLEX
"""

import pandas as pd

# Load the sales data from the excel file into a pandas DataFrame

# sales_data = pd.read_excel('C:/Users/REMLEX/Desktop/Python course/Data/Amazon Sales Project/sales_data.xlsx')
sales_data = pd.read_excel('sales_data.xlsx')

# =============================================================================
# Exploring the data
# =============================================================================

#Get the summary of sales data
sales_data.info()

sales_data.describe()

# Looking at column
print(sales_data.columns)

# Having a look at the first few rows of data
print(sales_data.head())

# Check the data types of the columns
print(sales_data.dtypes)

# =============================================================================
# Cleaning the data
# =============================================================================

# Check for missing values in our sales data
print(sales_data.isnull())

# Summary of missing value in our sales data
print(sales_data.isnull().sum())

# drop any row that has any missing/nan values
sales_data_dropped = sales_data.dropna()

# drop rows with missing amounts based on the amount column
sales_data_cleaned = sales_data.dropna(subset = ['Amount'])

# Check for missing values in our sales data cleaned
print(sales_data_cleaned.isnull().sum())

# =============================================================================
# Slicing and Filtering Data
# =============================================================================

# Select a subset of our data based on the category Column
category_data = sales_data[sales_data['Category'] == 'Top']
print(category_data)

# Select a subset of our data where the Amount > 1000
high_amount_data = sales_data[sales_data['Amount'] > 1000]
print(high_amount_data)

# Select a subset of data based on multiple conditions
filtered_data = sales_data[(sales_data['Category'] == 'Top') & (sales_data['Qty'] == 3)]


# =============================================================================
# Aggregating Data
# =============================================================================

# Total sales by category
category_totals = sales_data.groupby('Category')['Amount'].sum()
category_totals = sales_data.groupby('Category', as_index=False)['Amount'].sum()
# Sorted from largest to smallest
category_totals = category_totals.sort_values('Amount', ascending=False)
# double the resulr and click on format change from %.6g to %.2f
# %.6g is exponentials while %.2f is float 2 decimal place

# calculate the average Amount by Category and Fulfilment
fulfilment_averages = sales_data.groupby(['Category','Fulfilment'])['Amount'].mean()
fulfilment_averages = sales_data.groupby(['Category','Fulfilment'], as_index=False)['Amount'].mean()
# Sorted from largest to smallest
fulfilment_averages = fulfilment_averages.sort_values('Amount', ascending=False)

# calculate the average Amount by Category and Status
status_averages = sales_data.groupby(['Category','Status'])['Amount'].mean()
status_averages = sales_data.groupby(['Category','Status'], as_index=False)['Amount'].mean()
# Sorted from largest to smallest
status_averages = fulfilment_averages.sort_values('Amount', ascending=False)


# calculate totals sales by shipment and Fulfilment
total_sales_shipandfulfil = sales_data.groupby(['Courier Status','Fulfilment'], as_index=False)['Amount'].sum()
# Sorted from largest to smallest
total_sales_shipandfulfil = total_sales_shipandfulfil.sort_values('Amount', ascending=False)

# =============================================================================
# Renaming Columns and Exporting Data
# =============================================================================

# Rename
total_sales_shipandfulfil.rename(columns={'Courier Status':'Shipment'}, inplace=True)

# Exporting Data
high_amount_data.to_excel('amount_one_thousand.xlsx', index=False)

fulfilment_averages.to_excel('average_amount_by_category_and_fulfilment.xlsx', index=False)

status_averages.to_excel('average_sales_by_category_and_status.xlsx', index=False)
total_sales_shipandfulfil.to_excel('total_sales_by_ship_and_fulfil.xlsx', index=False)


