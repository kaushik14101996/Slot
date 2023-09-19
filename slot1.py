import streamlit as stl
import pandas as pd

import numpy as np
pd.set_option('display.max_columns', None)
import os
import time
import warnings
import signal
warnings.filterwarnings('ignore')

# Function to run your script

def slot_1_input(file_path):

    # Read the input Excel file
    slot1 = pd.read_excel(file_path,sheet_name='Slot 1')

    # Your Python code for processing the data goes here
    # Modify this code according to your specific requirements
    slot1.drop(columns= ['Unnamed: 0','Difference'],inplace = True)
    col = slot1.iloc[:,9:]
    list_col = col.columns
    slot1 = pd.melt(slot1,id_vars = ['Srl', 'Txn Date', 'Value Date', 'Description', 'CR/DR',
           'CCY','Amount (INR)'],value_vars = list_col)
    slot1.rename(columns = {'variable' : 'Store Code', 'value' : 'Amount'}, inplace = True)
    slot1['Amount']=slot1['Amount'].astype('float')
    slot1['Amount'].fillna(0,inplace=True)
    input_slot1 = slot1.copy()
    print("Input Preprocessing Complete")
    return input_slot1


def slot_1_output(file_path):
    global a,b

    # Read the input Excel file
    slot1 = pd.read_excel(file_path,sheet_name='Slot 1')

    # Your Python code for processing the data goes here
    # Modify this code according to your specific requirements
    slot1.drop(columns= ['Unnamed: 0','Difference'],inplace = True)
    col = slot1.iloc[:,9:]
    list_col = col.columns
    slot1 = pd.melt(slot1,id_vars = ['Srl', 'Txn Date', 'Value Date', 'Description', 'CR/DR',
           'CCY','Amount (INR)'],value_vars = list_col)
    slot1.rename(columns = {'variable' : 'Store Code', 'value' : 'Amount'}, inplace = True)
    slot1['Amount']=slot1['Amount'].astype('float')
    slot1['Amount'].fillna(0,inplace=True)
    input_slot1 = slot1.copy()
    print("Input Preprocessing Complete")
    input_slot1

#     #===============================================================

    input_slot1 = slot1[slot1['Amount'] != 0]

#     #===============================================================

    filtered_data_1 = input_slot1[input_slot1['Store Code'] != 67110022]

    output_slot1_1 = pd.DataFrame({
        'Document Date': filtered_data_1['Value Date'],
        'PstKy': 11,
        'Customer/Vendor': filtered_data_1['Store Code'],
        'Ttype': filtered_data_1['Value Date'].astype(str) + '-' + filtered_data_1['Srl'].astype(str),
        'Amount': filtered_data_1['Amount']
    })

#     #==================================================================

    filtered_data_2 = input_slot1[(input_slot1['Store Code'] == 67110022) & (input_slot1['Amount'] < 0)]

    grouped_data_2 = filtered_data_2.groupby(['Value Date', 'Srl']).agg({'Amount': 'sum'}).reset_index()

    output_slot1_2 = pd.DataFrame({
        'Document Date': grouped_data_2['Value Date'],
        'PstKy': 40,
        'Ttype': grouped_data_2['Value Date'].astype(str) + '-' + grouped_data_2['Srl'].astype(str),
        'GL Account': '67110022',
        'Amount': grouped_data_2['Amount']
    })

#     #===============================================================

    filtered_data_3 = input_slot1[(input_slot1['Store Code'] == 67110022) & (input_slot1['Amount'] > 0)]

    grouped_data_3 = filtered_data_3.groupby(['Value Date', 'Srl']).agg({'Amount': 'sum'}).reset_index()

    output_slot1_3 = pd.DataFrame({
        'Document Date': grouped_data_3['Value Date'],
        'PstKy': 50,
        'Ttype': grouped_data_3['Value Date'].astype(str) + '-' + grouped_data_3['Srl'].astype(str),
        'GL Account': '67110022',
        'Amount': grouped_data_3['Amount']
    })

#     #=======================================================================
    filtered_df_4 = input_slot1[~input_slot1['Store Code'].isin(['67110022'])]  # Exclude '67110022'
    filtered_df_4['Amount (INR)'] = pd.to_numeric(filtered_df_4['Amount (INR)'])
    grouped_data_4 = filtered_df_4.groupby(['Value Date', 'Srl','Amount (INR)']).agg({
        'Amount': 'sum'
    }).reset_index()

    output_slot1_4 = pd.DataFrame({
        'Document Date': grouped_data_4['Value Date'],
        'PstKy': 40,
        'Ttype': grouped_data_4['Value Date'].astype(str) + '-' + grouped_data_4['Srl'].astype(str),
        'GL Account': '10021419',
        'Amount': grouped_data_4['Amount (INR)']
    })
    

#     #==================================================================

    final_output = pd.concat([output_slot1_1, output_slot1_2, output_slot1_3,output_slot1_4])
    final_output['Amount']=round(final_output['Amount'],2)

#     #=================================================================

    final_output['Number'] = final_output['Ttype'].rank(method='dense').astype(int)

#     #====================================================================

    final_output = final_output[['Number','Document Date', 'Ttype', 'Amount', 'PstKy', 'GL Account', 'Customer/Vendor']]
    final_output = final_output.sort_values(by=['Ttype', 'Customer/Vendor'], ascending=[True, False])
    final_output=final_output[final_output['Amount']!=0]
    final_output['Amount']=round(final_output['Amount'],2)

#     #===================================================================

    def format_date(date_string):
        date_parts = date_string.split('-')
        return ''.join(date_parts)

    final_output['Document Date'] = final_output['Document Date'].astype('string')
    final_output['Document Date'] = final_output['Document Date'].apply(format_date)

#     #==========================================================


    final_output['Posting Date']=pd.Timestamp.now().strftime('%Y%m%d').replace('-', '')
    final_output['Period'] = pd.Timestamp.now().month
    final_output['Type']='DZ'
    final_output['Company Code']=1380
    final_output['Currency']='INR'
    final_output['Reference']='Mi home Card /Cash settlement'
    final_output['Doc#Header Tex']='Mi home Card /Cash settlement'
    final_output['India original invoice number']=' '
    final_output['India original invoice date']=pd.Timestamp.now().strftime('%Y%m%d').replace('-', '')
    final_output['SGL Ind']=' '
    final_output['Asset']=' '
    final_output['Amount in LC']=' '
    final_output['Cost Center']=' '
    final_output['Profit Center']=' '
    final_output['Order']=' '
    final_output['Payt Terms']=' '
    final_output['Bline Date']=' '
    final_output['Tax code']=' '
    final_output['Assignment']='Mi home Card /Cash settlement'
    final_output['Text']='Mi home Card /Cash settlement'
    final_output['Reason code']=' '
    final_output['Reason code 1']=' '
    final_output['Reference Key 1']=' '
    final_output['Reference Key 2']=' '
    final_output['Reference Key 3']=' '
    final_output['Trading Partner']=' '
    final_output['Number1']=' '
    final_output['Unit']=' '
    final_output['SKU']=' '
    final_output['Customer']=' '
    final_output['product']=' '
    final_output['Industry']=' '
    final_output['Xiaomi Bank Account']=' '
    final_output['Bank trading serial number']=' '

#     #=================================================================

    final_output.loc[final_output['GL Account'] == '67110022', 'Profit Center'] = 'P0WWZ1'
    final_output.loc[final_output['GL Account'] == '10021419', 'Customer/Vendor'] = ''
    final_output.loc[final_output['PstKy'] == 40, 'Reason code'] = '103'
    final_output = final_output[final_output['Customer/Vendor'] != '67110022']

#     #===================================================================

    final_output.insert(1,'Item',1)

    for i in range(len(final_output['Item'])):
        final_output['Item'] = np.where(final_output.iloc[::,0] == final_output.iloc[::,0].shift(1), final_output.iloc[::,1].shift(1) + 1, 1)

    final_output['Item']=final_output.Item.astype('int')

#     #==================================================================

    extract = final_output[['Number', 'Item','Document Date', 'Posting Date', 'Period', 'Type', 'Company Code', 'Currency',
               'Reference', 'Doc#Header Tex', 'India original invoice number', 'India original invoice date',
               'PstKy', 'Customer/Vendor', 'SGL Ind', 'Asset', 'Ttype', 'GL Account', 'Amount', 'Amount in LC',
               'Cost Center', 'Profit Center', 'Order', 'Payt Terms', 'Bline Date', 'Tax code', 'Assignment',
               'Text', 'Reason code', 'Reason code 1', 'Reference Key 1', 'Reference Key 2', 'Reference Key 3',
               'Trading Partner', 'Number1', 'Unit', 'SKU', 'Customer', 'product', 'Industry',
               'Xiaomi Bank Account', 'Bank trading serial number']]

#     #====================================================================

    for i in extract.columns.tolist():

        if (extract[i].dtype=='object'):
            extract[i].fillna('',inplace=True)
        elif (extract[i].dtype=='string'):
            extract[i].fillna('',inplace=True)    
            extract[i].replace(np.nan,'',inplace=True)
        elif (extract[i].dtype=='int64'):
            extract[i].fillna(0,inplace=True)
            extract[i].replace(np.nan,0,inplace=True) 
        elif (extract[i].dtype=='float64'):
            extract[i].fillna(0,inplace=True)
            extract[i].replace(np.nan,0,inplace=True) 

#      #======================================================================

    d=extract[(extract['PstKy'] == 50) | (extract['PstKy'] == 11)]
    a=round(d['Amount'].sum(),2)

    c=extract[extract['PstKy'] == 40]
    b=round(c['Amount'].sum(),2)

    if (a==b):
        print("Validation steps Match")
    else:
        print("Validation steps not matched") 

#     #======================================================================
    IP = slot1.copy()
    IP[['Srl', 'Store Code']] = IP[['Srl', 'Store Code']].apply(pd.to_numeric)
    extract[['Document Date', 'Posting Date', 'Period', 'India original invoice date', 'PstKy','Customer/Vendor', 'GL Account']]=extract[['Document Date', 'Posting Date', 'Period', 'India original invoice date', 'PstKy','Customer/Vendor', 'GL Account']].apply(pd.to_numeric)
#     #==================================================

    return  extract
#     if(a==b):
#         extract_dict = {'Upload_Template': extract, 'IP': IP}

#         st = pd.Timestamp.now().strftime('%Y-%m-%d')

#         filename = f'C:\\Users\\Rohit Kaushik\\Desktop\\BANK_SLOT1_UPLOAD_NEW_TEMPLATE1_final_{st}.xlsx'


#         with pd.ExcelWriter(filename) as writer:
#             for sheet_name, df in extract_dict.items():
#                 df.to_excel(writer, sheet_name=sheet_name, index=False)
#     else:
#         print("Error in the input file")   

#========================================================================================================
from io import BytesIO

from openpyxl import Workbook

import pandas as pd

# from pyxlsb import open_workbook as open_xlb
from io import BytesIO

def download_excel(Dataframe1, Dataframe2):
    op = BytesIO()
    wr = pd.ExcelWriter(op, engine='xlsxwriter')
    
    # First Sheet (output_slot_1)
    Dataframe1.to_excel(wr, index=False, sheet_name='output_slot_1')
    workbook = wr.book
    worksheet1 = wr.sheets['output_slot_1']
    fm1 = workbook.add_format({'num_format': '0.00'})
    worksheet1.set_column("A:A", None, fm1)
    
    # Second Sheet (input_slot_1)
    Dataframe2.to_excel(wr, index=False, sheet_name='input_slot_1')
    worksheet2 = wr.sheets['input_slot_1']
    fm2 = workbook.add_format({'bold': True})
    worksheet2.set_column("A:A", None, fm2)
    
    wr.close()
    data_ = op.getvalue()
    return data_

# Streamlit web application
def main():
    stl.title("Please Upload Input For Slot 1")
    
    # Upload Excel file
    uploaded_file = stl.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file is not None:
        # Read the uploaded file
        df = pd.read_excel(uploaded_file, sheet_name='Slot 1')

        # Display the uploaded data
        stl.subheader("Source Data")
        stl.dataframe(df)
        fla_click = 0
        # Run the processing function
        if stl.button("Run Processing And Generate Output"):
            fla_click = 1

            stl.subheader("Input Data After preprocess")
            # Call your function
            output_df = slot_1_output(uploaded_file)
            temp_output_df = output_df.copy(deep=True)
            input_df = slot_1_input(uploaded_file)  # Call your input_slot_1 function

            # Display the processed data
            
            stl.subheader("Output Data")
            stl.dataframe(output_df)
            if a == b:
                stl.write("Sum of amount for Pstky 11 and 40 =  {}".format(a))
                stl.write("Sum of amount for Pstky 50 =  {}".format(b))
                stl.subheader("Validation Step's Match")
            else:
                stl.write("Sum of amount for Pstky 11 and 40 = {}".format(a))
                stl.write("Sum of amount for Pstky 50 = {}".format(b))
                stl.subheader("Validation Step's Not Matched")  


            # Download the processed data
            # stl.subheader("Download Output Data")
            if fla_click == 1:
                st = pd.Timestamp.now().strftime('%Y-%m-%d')
                output_filename = f"Slot_1_{st}.xlsx"
                excel_data = download_excel(output_df, input_df)  # Pass both DataFrames
                stl.download_button(label="Download Output into Excel", data=excel_data, file_name=f"Slot_1_{st}.xlsx")

if __name__ == "__main__":
    main()


# In[ ]:




