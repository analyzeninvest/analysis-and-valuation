#!/usr/bin/env python

def search_for_excel(ticker):
    import os, re
    path = "/home/arnashree/analyzeninvest-projects/NSE_Financial_Database/excel_path/"
    for files in os.listdir(path):
        xlsx_name = ticker + ".xlsx"
        if re.match(xlsx_name, files):
            return(path + xlsx_name)
        
def make_comparison_excel(comparison_name, list_of_stocks, dict_of_attribute):
    import pandas as pd
    from datetime import date
    from openpyxl import Workbook
    path = "/home/arnashree/analyzeninvest-projects/analysis-and-valuation/comparison_path/"
    excel_name = path + comparison_name + ".xlsx"
    wb = Workbook()
    wb.save(excel_name)
    today = date.today()
    current_year = today.year
    array_of_stocks_with_attributes = fetch_attributes_from_excel(list_of_stocks, dict_of_attribute)
    start_index = 1
    writer = pd.ExcelWriter(excel_name, engine = 'openpyxl')
    for stock in list_of_stocks:
        print("################")
        print("Starting for " + stock)
        print("################")
        index_array = []
        for i in range(0,19):
            index_array.append(current_year - i)
        index = {"Year": index_array}
        dict_stock_attribute = index
        for items in array_of_stocks_with_attributes:
            for key in items:
                if key == stock:
                    dict_stock_attribute.update(items[key])
                    start_index += 1 
        #print(dict_stock_attribute)
        df_stock = pd.DataFrame(data = dict_stock_attribute)
        df_stock.set_index("Year")
        df_stock_T = df_stock.T
        df_stock_T.to_excel(writer, sheet_name = stock)
        writer.save()
        #print(df_stock_T)

def fetch_attributes_from_excel(list_of_stocks, dict_of_attribute):
    """
    This will make the industry comparison excel from the existing excel sheets.
    """
    from openpyxl import load_workbook
    stock_sheet_attribute_details = []
    for stock in list_of_stocks:
        print("################")
        print("Fetching for " + stock)
        print("################")
        stock_xls_path = search_for_excel(stock)
        if not stock_xls_path:
            sys.error("Stock neither can be found nor downloaded. Exiting !!!")
        wb = load_workbook(stock_xls_path)
        for sheet_name in dict_of_attribute:
            # need to open the xls by sheet_name name.
            for sheet in wb:
                if sheet.title == sheet_name:
                    # choosing BZ for the last col from row 2
                    #cell_headers = sheet['A2':'BZ2']
                    #print(cell_headers)
                    #for cell in cell_headers:
                    for attribute in dict_of_attribute[sheet_name]:
                        #print(cell) download all lines from this
                        # attribute.  will have to go for the
                        # formatted xlsx ? no need as this will work
                        # for rawdata.
                        for col in sheet.iter_cols(min_row=2,max_col=65,max_row=2):
                            #print(col)
                            for cell in col:
                                #print(cell.value)
                                if cell.value == attribute:
                                    column = cell.column_letter
                                    #print(column)
                                    value = []
                                    key = sheet[column + str(2)].value
                                    for index in range(3, 22):
                                        value.append(sheet[column + str(index)].value)
                                    key_value_pair = {key:value}
                                    stock_sheet_attribute_details.append({stock: key_value_pair})
                                    #print(key_value_pair)
    return(stock_sheet_attribute_details)
                        
def add_attribute():
    pass

def add_stocks():
    pass

def main():
    """
    For now this is only for testing the codes.
    """
    print("Running ... ")
    stock_ticker_list = ['BDL', 'ASTRAMICRO']
    dict_of_attribute = {"Standalone_Balance_Sheet":[
        "Total Non-Current Liabilities",
        "Current Investments Unquoted Book Value",
        "Equity Share Capital"],
                         "Standalone_Profit_and_Loss":[
                             "Total Revenue",
                             "Depreciation And Amortisation Expenses"
                         ]}
    #print(fetch_attributes_from_excel(stock_ticker_list, dict_of_attribute))
    make_comparison_excel("Defence",stock_ticker_list, dict_of_attribute)

if __name__ == '__main__':
    main()
