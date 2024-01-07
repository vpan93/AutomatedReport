# imports

import pandas as pd
import numpy as np
import openpyxl
import logging 
from helpers.utils import load_excel_sheets,write_dataframe_to_excel
import logging 
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
import os
def main(logger: logging.Logger):

    logger.info("Starting main.py")
    logger.info("Loading the data ...")
    # Load the data

    file_path = 'data/data_sheet.xlsx'
    sheet_names = ['Customer', 'Order', 'Product','Customer_group','Reviews']

    try:
        df_Customer, df_Order, df_Product,df_Customer_group,df_Reviews = load_excel_sheets(file_path, sheet_names)
        # Work with the DataFrames as needed
    except ValueError as e:
        print(f"Error: {e}")

    
    logger.info("Creating tables for question 1 ...")

    ## 1.1  Which type of customer group has spent the most money?
    # We will join the tables df_Order, df_Customer_group, df_Product and df_Customer_group using inner join
    # We will group by customer_group and sum the price
    # Using pandas chaining for cleaner and quicker code

    df_Q1_1 = (df_Order
                .merge(df_Product, on='product_id')
                .merge(df_Customer, on='order_id')
                .merge(df_Customer_group, on='customer_id')
                .groupby('customer_group')['price'].sum().reset_index()
                .sort_values(by='price', ascending=False)
            )
    
    
    # Finding the customer group with the highest spending
    highest_spending_group = df_Q1_1.iloc[0]

    # The customer group P has spent the most money, with a total expenditure of 45,781.

    ## Question 1.2: How much money has been spent on each type of customer group on refunded orders?

    # Filtering for refunded orders
    refunded_orders = df_Order[df_Order['refund'] == True]

    df_Q2_2 = (refunded_orders
                    .merge(df_Product, on='product_id')
                    .merge(df_Customer, on='order_id')
                    .merge(df_Customer_group, on='customer_id')
                    .groupby('customer_group')['price'].sum().reset_index()
                    .sort_values(by='price', ascending=False)
               )


    # Question 1.3: Which type of product drove the most sales for customer_groups [A - D] where product value was lower than 100?

    # Filtering for products with price lower than 100
    low_price_products = df_Product[df_Product['price'] < 100]

    target_customer_groups = ['A', 'B', 'C', 'D']

    most_sold_product_type = (
                df_Order
                .merge(low_price_products, on='product_id')
                .merge(df_Customer, on='order_id')
                .merge(df_Customer_group, on='customer_id')
                .query('customer_group in @target_customer_groups')
                .groupby('product_type')['order_id']
                .count()
                .reset_index()
                .sort_values(by='order_id', ascending=False)
                # .iloc[0]
    )

    # For customer groups A to D, with product value lower than 100, the product type group_10 drove the most sales, with a total of 4 orders.
    logger.info("Creating tables for question 2 ...")
    # Question 2.1: Which product has been sold / refunded the most?

    # Merging df_Order with df_Product
    df_merged_sold_refunded = df_Order.merge(df_Product, on='product_id')

    # Counting sold and refunded products
    sold_count = df_merged_sold_refunded['product_id'].value_counts()
    refunded_count = df_merged_sold_refunded[df_merged_sold_refunded['refund']]['product_id'].value_counts()

    # Identifying the most sold and most refunded products
    most_sold_product = sold_count.idxmax()
    most_refunded_product = refunded_count.idxmax()

    # The product that has been sold the most is identified by the ID Y4SOFJ5RTE.
    # The product that has been refunded the most is identified by the ID 5XO0ZJ72NI.

    # Question 2.2: How many products with value more than 450 have been refunded for customer_group [F - J]?

    # Filtering products with value more than 450
    high_value_products = df_Product[df_Product['price'] > 450]

    # Merging and filtering for refunded orders and customer groups F to J
    df_merged_high_value = (df_Order
                            .merge(high_value_products, on='product_id')
                            .merge(df_Customer, on='order_id')
                            .merge(df_Customer_group, on='customer_id')
                            .query('refund == True and customer_group in ["F", "G", "H", "I", "J"]'))

    # Counting the number of such products
    count_high_value_refunded = df_merged_high_value['product_id'].nunique()

    # There are 18 unique products with a value more than 450 that have been refunded for customer groups F through J.


    # Question 2.3: What is the value of these orders?

    # Summing up the values of the refunded orders for customer groups F to J
    total_value_refunded_orders = df_merged_high_value['price'].sum()

    # The total value of these refunded orders for products with a value more than 450 in customer groups F through J is 14,364.

    logger.info("Creating tables for question 3 ...")
    # Question 3.1 Is it true that the higher the type of product, the more money has been spent on it?
    # Merging df_Order with df_Product
    df_merged_orders_products_q3 = df_Order.merge(df_Product, on='product_id')

    # Grouping by product type and calculating total money spent for each product type
    total_spent_per_product_type = df_merged_orders_products_q3.groupby('product_type')['price'].sum().sort_values(ascending=False)

    # Based on the provided dataset, it is not necessarily true that higher product types always correspond to more money being spent. The spending pattern seems to vary, with some higher product types having higher average spending per order and others not.
    # This analysis is dependent on the specific dataset provided and the way "higher" is defined for product types.

    # Question 3.2 Can you plot the ratio between the sum of prices and the number of orders made for each product type?

    # Calculating the count of orders for each product type
    order_count_per_product_type = df_merged_orders_products_q3.groupby('product_type')['order_id'].count()

    # Calculating the ratio of total spent to order count
    spend_order_ratio = total_spent_per_product_type / order_count_per_product_type

    # # Plotting the ratio
    # plt.figure(figsize=(12, 6))
    # spend_order_ratio.sort_values(ascending=False).plot(kind='bar', color='skyblue')
    # plt.xlabel('Product Type')
    # plt.ylabel('Average Spending per Order')
    # plt.title('Average Spending per Order for Each Product Type')
    # plt.xticks(rotation=45)
    # plt.grid(axis='y')
    # plt.tight_layout()
    # plt.show()

    logger.info("Creating tables for question 4 ...")
    # Question 4.1 Which product group is the most popular for customer groups [A, B, C, D]?

    target_groups_4_1 = ['A', 'B', 'C', 'D']

    most_popular_product_group_4_1 = (
        df_Order.merge(df_Customer, on='order_id')
                .merge(df_Product, on='product_id')
                .merge(df_Customer_group, on='customer_id')
                .query('customer_group in @target_groups_4_1')
                .groupby('product_type')
                .size()
                # .idxmax()
    )

    # Question 4.2 Which customer group buys the most of each product group?

    most_buying_customer_group_4_2 = (
    df_Order.merge(df_Customer, on='order_id')
            .merge(df_Product, on='product_id')
            .merge(df_Customer_group, on='customer_id')
            .groupby(['product_type', 'customer_group'])
            .size()
            .reset_index(name='order_count')
            .loc[lambda df: df.groupby('product_type')['order_count'].idxmax()]
            .sort_values(by='order_count',ascending=False)
    )

    logger.info("Creating tables for question 5 ...")
    # Question 5. What is the probability that an order by a customer group [A - D] is refunded?

    # Merging the necessary dataframes
    df_merged_5 = df_Order.merge(df_Customer, on='order_id').merge(df_Customer_group, on='customer_id')

    # Filtering for customer groups [A - D]
    target_customer_groups_5 = ['A', 'B', 'C', 'D']
    df_filtered_5 = df_merged_5[df_merged_5['customer_group'].isin(target_customer_groups_5)]

    # Calculating the total number of orders and the number of refunded orders
    total_orders = df_filtered_5.shape[0]
    refunded_orders = df_filtered_5[df_filtered_5['refund'] == True].shape[0]

    # Computing the probability of a refund
    probability_of_refund = refunded_orders / total_orders if total_orders > 0 else 0

    logger.info("Writing to excel ...")
    # write csv
    existing_excel_filename = 'output.xlsx'

    write_dataframe_to_excel(existing_excel_filename,  list_start_row= [1, 1, 1],list_start_col= [1, 7, 14],list_of_tables = [df_Q1_1, df_Q2_2, most_sold_product_type],list_of_titles = ['Question 1.1','Question 1.2','Question 1.3'], sheet_name='Question_1')
    write_dataframe_to_excel(existing_excel_filename,  list_start_row = [1,1,1],list_start_col = [1,7,14],list_of_tables = [pd.DataFrame(sold_count).reset_index().rename(columns={"index": "product_id", "B": "count"}),pd.DataFrame(refunded_count).reset_index().rename(columns={"index": "product_id", "B": "count"}),df_merged_high_value],list_of_titles = ['Question 2.1','Question 2.2','Question 2.3'], sheet_name='Question_2')
    write_dataframe_to_excel(existing_excel_filename,  list_start_row = [1,1],list_start_col = [1,7],list_of_tables = [pd.DataFrame(total_spent_per_product_type).reset_index().rename(columns={"index": "product_type", "B": "total_amount_spent"}),pd.DataFrame(spend_order_ratio).reset_index().rename(columns={"index": "product_type", 0: "ratio"})],list_of_titles = ['Question 3.1','Question 3.2'], sheet_name='Question_3')
    write_dataframe_to_excel(existing_excel_filename,  list_start_row = [1,1],list_start_col = [1,7],list_of_tables = [pd.DataFrame(most_popular_product_group_4_1).reset_index().rename(columns={"index": "product_type", 0: "count"}),most_buying_customer_group_4_2],list_of_titles = ['Question 4.1','Question 4.2'], sheet_name='Question_4')

    logger.info("Script completed succesfully ...")
    return None

if __name__ == '__main__':
    logger = logging.getLogger()

    # Create console handler and set level to info
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # Create formatter and add it to the handler
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)

    # Add the handler to the logger
    logger.addHandler(console_handler)
    logger.setLevel(logging.INFO)

    main(logger=logger)  # Run the main function if this script is executed directly
