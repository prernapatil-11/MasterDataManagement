# Import libs
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import os

from openpyxl import load_workbook

file_path = os.getcwd()
# Import file
df = pd.read_csv(file_path + '/amazon.csv')

# Replace string value and change data type
df['actual_price'] = df['actual_price'].str.replace('₹','')
df['actual_price'] = df['actual_price'].str.replace(',','').astype('float64')

df['discounted_price'] = df['discounted_price'].str.replace('₹','')
df['discounted_price'] = df['discounted_price'].str.replace(',','').astype('float64')
df['discount_percentage'] = df['discount_percentage'].str.replace('%','').astype('float64')
df['discount_percentage'] = df['discount_percentage']/100
df['rating_count'] = df['rating_count'].str.replace(',','').astype('float64')
df['rating'].value_counts()

# Look at the strange row
df[df['rating'] == '|']

df['rating'] = df['rating'].str.replace('|','4.0').astype('float64')

# Create new data frame with selected columns
df1 = df[['product_id', 'product_name', 'category', 'discounted_price', 'actual_price', 'discount_percentage', 'rating', 'rating_count','product_link','about_product']].copy()
# Split `category` column
cat_split = df1['category'].str.split('|', expand=True)
cat_split.isnull().sum()

# Rename column
cat_split = cat_split.rename(columns={0:'Main category', 1:'Sub category'})

# Add new cols to data frame and drop the old ones
df1['Main category'] = cat_split['Main category']
df1['Sub category'] = cat_split['Sub category']
df1.drop(columns ='category', inplace=True)
df1['Main category'].value_counts()

# Fix the strings in `Main category`
df1['Main category'] = df1['Main category'].str.replace('&', ' & ')
df1['Main category'] = df1['Main category'].str.replace('OfficeProducts', 'Office Products')
df1['Main category'] = df1['Main category'].str.replace('MusicalInstruments', 'Musical Instruments')
df1['Main category'] = df1['Main category'].str.replace('HomeImprovement', 'Home Improvement')
df1['Sub category'].value_counts()

# I will do the same with `Sub category`
df1['Sub category'] = df1['Sub category'].str.replace('&', ' & ')
df1['Sub category'] = df1['Sub category'].str.replace(',', ', ')
df1['Sub category'] = df1['Sub category'].str.replace('HomeAppliances', 'Home Appliances')
df1['Sub category'] = df1['Sub category'].str.replace('AirQuality', 'Air Quality')
df1['Sub category'] = df1['Sub category'].str.replace('WearableTechnology', 'Wearable Technology')
df1['Sub category'] = df1['Sub category'].str.replace('NetworkingDevices', 'Networking Devices')
df1['Sub category'] = df1['Sub category'].str.replace('OfficePaperProducts', 'Office Paper Products')
df1['Sub category'] = df1['Sub category'].str.replace('ExternalDevices', 'External Devices')
df1['Sub category'] = df1['Sub category'].str.replace('DataStorage', 'Data Storage')
df1['Sub category'] = df1['Sub category'].str.replace('HomeStorage', 'Home Storage')
df1['Sub category'] = df1['Sub category'].str.replace('HomeAudio', 'Home Audio')
df1['Sub category'] = df1['Sub category'].str.replace('GeneralPurposeBatteries', 'General Purpose Batteries')
df1['Sub category'] = df1['Sub category'].str.replace('BatteryChargers', 'Battery Chargers')
df1['Sub category'] = df1['Sub category'].str.replace('CraftMaterials', 'Craft Materials')
df1['Sub category'] = df1['Sub category'].str.replace('OfficeElectronics', 'Office Electronics')
df1['Sub category'] = df1['Sub category'].str.replace('PowerAccessories', 'Power Accessories')
df1['Sub category'] = df1['Sub category'].str.replace('CarAccessories', 'Car Accessories')
df1['Sub category'] = df1['Sub category'].str.replace('HomeMedicalSupplies', 'Home Medical Supplies')
df1['Sub category'] = df1['Sub category'].str.replace('HomeTheater', 'Home Theater')
df1.drop_duplicates()
path = file_path + "\outputAmazonSales.xlsx"
# book = load_workbook(path)
writer = pd.ExcelWriter(path, engine='openpyxl')
# df1.to_excel('outputAmazonSales.xlsx',sheet_name = "Productmasterdata")
df1.to_excel(writer,sheet_name = "Productmasterdata")

df2 = df[['user_id','user_name']]
############################
# # Convert columns of interest to list columns
df2["user_id"] = df2["user_id"].str.split(",")
df2["user_name"] = df2["user_name"].str.split(",")

df2['dic_user'] = df2[['user_id','user_name']].apply(lambda x: dict(zip(*x)),axis=1)
df2 = df2.drop(['user_id','user_name'], axis=1)
df_result = (
    df2
    .assign(dict=df2.dic_user.map(lambda d: d.items()))
    .explode("dict")
    .assign(
        userid=lambda df2: df2.dict.str.get(0),
        username=lambda df2: df2.dict.str.get(1)
    )
    .drop(columns="dict")
    .drop(columns="dic_user")
    .reset_index(drop=True)
)

# df2.to_excel('outputAmazonSales.xlsx',sheet_name = "Usermasterdata")
# df_result.to_excel('outputAmazonSales.xlsx',sheet_name = "Usermasterdata")
df_result.to_excel(writer,sheet_name = "Usermasterdata")
df.to_excel(writer,sheet_name = "rawData")
writer.close()