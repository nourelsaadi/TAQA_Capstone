#importing packages
import streamlit as st
import pandas as pd
import math
import numpy as np
from datetime import datetime
from shutil import copyfile
from openpyxl import load_workbook
from openpyxl.styles import Font, Color, Alignment, Border, Side, colors
from openpyxl.worksheet.cell_range import CellRange
import itertools
import warnings
from io import BytesIO
from pyxlsb import open_workbook as open_xlsb
import time

warnings.filterwarnings('ignore')


#Making sure the App is first loaded with a wide layout 
st.set_page_config(
     layout="wide")
     

# Setting the Home Page
st.markdown("<h1 style='text-align: center; color: #f2c70f;'>TAQA Profitability Analysis Wrangler</h1><br>", unsafe_allow_html=True) 
st.markdown('<center><img style="max-width:50%" src="https://github.com/nourelsaadi/TAQA_Capstone/blob/main/TAQA-750x500px.png?raw=true"></center>', unsafe_allow_html = True)
     #Introductory Statements
st.markdown("<center><h5> Update the Transactions table in the Excel Database you have received with the sales data you wish to analyze. <br> Next, upload the Excel Database below. <br> You will then have to download the 'Profitability Fact Table'.  </center></h5>", unsafe_allow_html=True)
     
#Setting a place to upload the dataset
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

#Conditionning the page (if data is uploaded)
if uploaded_file is not None:
    #calculation
      dbTaqa = pd.ExcelFile(uploaded_file)
      dfContinents = pd.read_excel(dbTaqa, 'tbContinents')
      dfRegions = pd.read_excel(dbTaqa, 'tbRegions')
      dfCountries = pd.read_excel(dbTaqa, 'tbCountries')
      dfCustomers = pd.read_excel(dbTaqa, 'tbCustomers')
      dfDiscounts = pd.read_excel(dbTaqa, 'tbDiscounts')
      dfProducts = pd.read_excel(dbTaqa, 'tbProducts')
      dfComposition = pd.read_excel(dbTaqa, 'tbComposition')
      dfComponents = pd.read_excel(dbTaqa, 'tbComponents')
      dfTransactions = pd.read_excel(dbTaqa, 'tbTransactions')

     ### Define functions ###
      def fnGetCountryName(iCountryID):
           #Returns the country name associated to the given country ID
           if iCountryID > 0: #to avoid 'nan'
                strRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['strCountryName'].max() 
                return strRes

      def fnGetCountryRegion(iCountryID):
           #Returns the country region name (ex. GCC for UAE) associated to the given country ID
           if iCountryID > 0:
                iRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['iRegionID'].max()
                strRes = dfRegions[(dfRegions['iRegionID']==iRes)]['strRegionName'].max()
           return strRes

      def fnGetCountryRegion(iCountryID):
           #Returns the country region ID (ex. GCC for UAE) associated to the given country ID
           if iCountryID > 0:
                iRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['iRegionID'].max()
                strRes = dfRegions[(dfRegions['iRegionID']==iRes)]['strRegionName'].max()
           return strRes

      def fnGetCountryRegionID(iCountryID):
           #Returns the country region ID (ex. GCC for UAE) associated to the given country ID
           if iCountryID > 0:
                iRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['iRegionID'].max()
           return iRes

      def fnGetCountryContinent(iCountryID):
           #Returns the country continent name (ex. Asia for UAE) associated to the given country ID
           if iCountryID > 0:
                iRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['iContinentID'].max()
                strRes = dfContinents[(dfContinents['iContinentID']==iRes)]['strContinentName'].max()
           return strRes

      def fnGetCountryAlpha2(iCountryID):
           #Returns the two-letter country code (ex. AE for UAE) associated to the given country ID
           if iCountryID > 0:
                strRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['strCountryCodeAlpha2'].max()
           return strRes

      def fnGetCountryAlpha3(iCountryID):
           #Returns the three-letter country code (ex. ARE for UAE) associated to the given country ID
           if iCountryID > 0:
                strRes = dfCountries[(dfCountries['iCountryID'] == iCountryID)]['strCountryCodeAlpha3'].max()
           return strRes

      def fnGetCustomerName(iCustomerID):
           #Returns the customer name associated to the given customer ID
           if iCustomerID > 0:
                strRes = dfCountries[(dfCustomers['iCustomerID'] == iCustomerID)]['strCustomerName'].max()
           return strRes

      def fnGetCustomerType(iCustomerID):
           #Returns the customer type associated to the given customer ID
           if iCustomerID > 0:
                strRes = dfCustomers[(dfCustomers['iCustomerID'] == iCustomerID)]['strCustomerType'].max()
           return strRes

      def fnGetCustomerCountry(iCustomerID):
           #Returns the customer country associated to the given customer ID
           if iCustomerID > 0:
                strRes = dfCustomers[(dfCustomers['iCustomerID'] == iCustomerID)]['iCountryID'].max()
           return strRes

      def fnGetDiscountRate(iDiscountID):
           #Returns the discount rate associated to the given discount ID
           if iDiscountID > 0:
                dRes = dfDiscounts[(dfDiscounts['iDiscountID'] == iDiscountID)]['dDiscountRate'].max()
           return dRes

      def fnGetProductName(iProductID):
           #Returns the product name associated to the given product ID
           if iProductID > 0:
                strRes = dfProducts[(dfProducts['iProductID'] == iProductID)]['strProductName'].max()
           return strRes

      def fnGetProductItemCount(iProductID):
           #Returns the count of items (ex. cookies) within a given product ID
           if iProductID > 0:
                iRes = dfProducts[(dfProducts['iProductID'] == iProductID)]['iItemCount'].max()
           return iRes

      def fnGetProductPriceLocalUSD(iProductID):
           #Returns the product price for local market in USD for a given product ID
           if iProductID > 0:
                dRes = dfProducts[(dfProducts['iProductID'] == iProductID)]['dProductPriceLocalUSD'].max()
           return dRes

      def fnGetProductPriceExportUSD(iProductID):
           #Returns the product price for export market in USD for a given product ID
           if iProductID > 0:
                dRes = dfProducts[(dfProducts['iProductID'] == iProductID)]['dProductPriceExportUSD'].max()
           return dRes

      def fnGetComponentName(iComponentID):
           #Returns the component name associated to the given component ID
           if iComponentID > 0:
                strRes = dfComponents[(dfComponents['iComponentID'] == iComponentID)]['strComponentName'].max()
           return strRes

      def fnGetComponentType(iComponentID):
           #Returns the component type associated to the given component ID
           if iComponentID > 0:
                strRes = dfComponents[(dfComponents['iComponentID'] == iComponentID)]['strComponentType'].max()
           return strRes

      def fnGetComponentUnitCost(iComponentID):
           #Returns the component cost per unit associated to the given component ID
           if iComponentID > 0:
                dRes = dfComponents[(dfComponents['iComponentID'] == iComponentID)]['dUnitCost'].max()
           return dRes

      def fnGetComponentUnit(iComponentID):
           #Returns the component unit associated to the given component ID
           if iComponentID > 0:
                strRes = dfComponents[(dfComponents['iComponentID'] == iComponentID)]['strUnit'].max()
           return strRes

      def fnGetGrossTheoreticalTransactionAmountUSD(row):
           #Returns theoretical transaction amount, based on product list price for customer's market
           #Chcek if customer is in Lebanon or in Export country
           if fnGetCountryAlpha2(fnGetCustomerCountry(row['iCustomerID'])) == 'LB':
                #Get theoretical transaction amount assuming local market rate 
                dGrossTheoreticalTransactionAmount = row['iQuantity'] * fnGetProductPriceLocalUSD(row['iProductID'])
           else:
                #Get theoretical transaction amount assuming export market rate
                dGrossTheoreticalTransactionAmount = row['iQuantity'] * fnGetProductPriceExportUSD(row['iProductID'])
           return round(dGrossTheoreticalTransactionAmount,2)

      def fnGetTheoreticalTransactionDiscountUSD(row):
           #Returns theoretical transaction discount amount, based on discount code
           dTheoreticalDiscountAmountUSD = row['dGrossTheoreticalTransactionAmountUSD'] * fnGetDiscountRate(row['iDiscountID'])
           return round(dTheoreticalDiscountAmountUSD,2)

      def fnGetRawMaterialCostUSD(row):
           #Returns the cost of raw materials for each product"""
           dfTmpProducts = dfComposition.merge(dfComponents, how='left', left_on='iSubComponentID', right_on='iComponentID')
           dfTmpProducts = dfTmpProducts[(dfTmpProducts['iComponentID_x'] == row['iSubComponentID'])&(dfTmpProducts['strComponentType'] == 'Raw material')]
           dfTmpProducts['dTotalCostUSD'] = dfTmpProducts['dSubComponentQuantity'] * dfTmpProducts['dUnitCost'] / 1000
    
           return round(dfTmpProducts['dTotalCostUSD'].sum(),2)
    

      def fnGetProductsRawMaterialCostUSD(iComponentID):
           #Returns the cost of raw materials associated to items in the given product ID
           dfTmpProducts = dfComposition.merge(dfComponents, how='left', left_on='iSubComponentID', right_on='iComponentID')
           dfTmpProducts = dfTmpProducts[(dfTmpProducts['iComponentID_x'] == iComponentID)&(dfTmpProducts['strComponentType'] == 'Product')]
    
           dfTmpProducts['dRawMaterialCostUSD'] = dfTmpProducts.apply(fnGetRawMaterialCostUSD, axis=1)
           dfTmpProducts['dTotalRawMaterialCostUSD'] = dfTmpProducts['dRawMaterialCostUSD'] * dfTmpProducts['dSubComponentQuantity']
     
           return round(dfTmpProducts['dTotalRawMaterialCostUSD'].sum(),2)

      def fnGetRawMaterialCostsUSD(row):
           #Returns raw material costs based on product code   
           #Multiply sum of costs of subcomponents with units per transaction
           return round(fnGetProductsRawMaterialCostUSD(row['iProductID']) * row['iQuantity'],2)

      def fnGetPackagingCostsUSD(row):
           #Returns the cost of packaging associated to product ID
           dfTmpProducts = dfComposition.merge(dfComponents, how='left', left_on='iSubComponentID', right_on='iComponentID')
           dfTmpProducts = dfTmpProducts[(dfTmpProducts['iComponentID_x'] == row['iProductID'])&(dfTmpProducts['strComponentType'] == 'Packaging')]
           dfTmpProducts['dPackagingCostUSD'] = dfTmpProducts['dSubComponentQuantity'] * dfTmpProducts['dUnitCost']
           return round(dfTmpProducts['dPackagingCostUSD'].sum() * row['iQuantity'],2)

      def fnGetDieselCostsUSD(row):
           #Returns the cost of diesel associated to product ID
           dfTmpProducts = dfComponents[(dfComponents['strComponentName'] == 'DIESEL')]
           dfTmpProducts['dDieselCostUSD'] = row['iQuantity'] * fnGetProductItemCount(row['iProductID']) * dfTmpProducts['dUnitCost']
           return round(dfTmpProducts['dDieselCostUSD'].sum(),2)

      def fnGetOtherDirectCostsUSD(row):
           #Returns the cost of direct costs (excl. diesel) associated to product ID
           #Chcek if customer is in Lebanon or in Export country
           if fnGetCountryAlpha2(fnGetCustomerCountry(row['iCustomerID'])) == 'LB':
                #Get theoretical transaction amount assuming local market rate 
                dfTmpProducts = dfComponents[(dfComponents['strComponentName'] == 'OTHER DIRECT COST LOCAL')]
           else:
                #Get theoretical transaction amount assuming export market rate
                dfTmpProducts = dfComponents[(dfComponents['strComponentName'] == 'OTHER DIRECT COST EXPORT')]
        
           dfTmpProducts['dOtherDirectCostsUSD'] = row['iQuantity'] * fnGetProductItemCount(row['iProductID']) * dfTmpProducts['dUnitCost']
           return round(dfTmpProducts['dOtherDirectCostsUSD'].sum(),2)

      def fnGetIndirectCostsUSD(row):
           #Returns the cost of indirect costs associated to product ID
           #Chcek if customer is in Lebanon or in Export country
           if fnGetCountryAlpha2(fnGetCustomerCountry(row['iCustomerID'])) == 'LB':
                #Get theoretical transaction amount assuming local market rate 
                dfTmpProducts = dfComponents[(dfComponents['strComponentName'] == 'OTHER INDIRECT COST LOCAL')]
           else:
                #Get theoretical transaction amount assuming export market rate
                dfTmpProducts = dfComponents[(dfComponents['strComponentName'] == 'OTHER INDIRECT COST EXPORT')]
        
           dfTmpProducts['dOtherIndirectCostsUSD'] = row['iQuantity'] * fnGetProductItemCount(row['iProductID']) * dfTmpProducts['dUnitCost']
           return round(dfTmpProducts['dOtherIndirectCostsUSD'].sum(),2)

      def fnGetMarketType(row):
           #Returns the market type (local or export) associated to each transaction"""
           strRes = ''
           if fnGetCountryAlpha2(fnGetCustomerCountry(row['iCustomerID'])) == 'LB':
                strRes = 'Local'
           else:
                strRes = 'Export'
           return strRes

      def fnGetCustomerCountryCode(row):
           #Returns the customer country ID associated to the given customer ID
           return fnGetCustomerCountry(row['iCustomerID'])

      def fnGetContinent(row):
           #Returns the customer continent associated to the given customer ID
           return fnGetCountryContinent(row['iCountryCode'])

      def fnGetRegion(row):
           #Returns the customer region associated to the given customer ID
           return fnGetCountryRegion(row['iCountryCode'])

      def fnGetCountry(row):
           #Returns the customer country associated to the given customer ID
           return fnGetCountryName(row['iCountryCode'])


      #Create base profitability table (fact table)
      dfProfitability = dfTransactions

      # Calculate revenue fields
      dfProfitability['dGrossTheoreticalTransactionAmountUSD'] = dfProfitability.apply(fnGetGrossTheoreticalTransactionAmountUSD, axis=1)
      dfProfitability['dTheoreticalDiscountAmountUSD'] = dfProfitability.apply(fnGetTheoreticalTransactionDiscountUSD, axis=1)
      dfProfitability['dNetTheoreticalTransactionAmountUSD'] = round(dfProfitability['dGrossTheoreticalTransactionAmountUSD'] - dfProfitability['dTheoreticalDiscountAmountUSD'],2)
      dfProfitability['dDiscountLoss'] = round(dfProfitability['dTheoreticalDiscountAmountUSD'] - dfProfitability['dDiscountAmountUSD'],2)
      dfProfitability['dSalesLoss'] = round(dfProfitability['dNetTheoreticalTransactionAmountUSD'] - dfProfitability['dTransactionAmountUSD'],2)

      # Calculate cost fields

      dfProfitability['dRawMaterialCostsUSD'] = dfProfitability.apply(fnGetRawMaterialCostsUSD, axis=1) 
      dfProfitability['dPackagingCostsUSD'] =  dfProfitability.apply(fnGetPackagingCostsUSD, axis=1) 
      dfProfitability['dDieselCostsUSD'] = dfProfitability.apply(fnGetDieselCostsUSD, axis=1)
      dfProfitability['dOtherDirectCostsUSD'] =  dfProfitability.apply(fnGetOtherDirectCostsUSD, axis=1)
      dfProfitability['dDirectCostsUSD'] = dfProfitability['dDieselCostsUSD'] + dfProfitability['dOtherDirectCostsUSD']
      dfProfitability['dIndirectCostsUSD'] = dfProfitability.apply(fnGetIndirectCostsUSD, axis=1)
      dfProfitability['dCostsUSD'] = dfProfitability['dRawMaterialCostsUSD'] + dfProfitability['dPackagingCostsUSD'] + dfProfitability['dDirectCostsUSD'] + dfProfitability['dIndirectCostsUSD']

      # Calculate net profit
      dfProfitability['dProfitUSD'] = dfProfitability['dTransactionAmountUSD'] - dfProfitability['dCostsUSD']

      dfProfitability['iCountryCode'] = dfProfitability.apply(fnGetCustomerCountryCode, axis=1)
      dfProfitability['strMarketType'] = dfProfitability.apply(fnGetMarketType, axis=1)
      dfProfitability['strContinent'] = dfProfitability.apply(fnGetContinent, axis=1)
      dfProfitability['strRegion'] = dfProfitability.apply(fnGetRegion, axis=1)
      dfProfitability['strCountry'] = dfProfitability.apply(fnGetCountry, axis=1)

      def fnGetExcelFile(df):
           output = BytesIO()
           writer = pd.ExcelWriter(output, engine='xlsxwriter')
           df.to_excel(writer, index=False, sheet_name='tbFact')
           workbook = writer.book
           worksheet = writer.sheets['tbFact']
           format1 = workbook.add_format({'num_format': '0.00'}) 
           worksheet.set_column('A:A', None, format1)  
           writer.save()
           processed_data = output.getvalue()
           return processed_data

      dfProfitability_xlsx = fnGetExcelFile(dfProfitability)
      strTime = time.strftime('%d%m%Y')
      st.download_button(label='📥Download Profitability Fact Table', data=dfProfitability_xlsx, file_name='Taqa_Profitability_Fact'+strTime+'.xlsx')

      #if st.download_button():
           #st.write('You can now upload this profitability fact table onto PowerBI!')

else: 
      # Show warning if dataset not uploaded yet st.warning("Please Upload a Dataset in the Data Upload Slot Above")
      st.warning("Please Upload your Excel File in the Data Upload Slot Above")
