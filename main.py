#!/usr/bin/env python

def search_for_excel(ticker):
    import os, re
    path = "/home/arnashree/analyzeninvest-projects/NSE_Financial_Database/excel_path/"
    for files in os.listdir(path):
        xlsx_name = ticker + ".xlsx"
        if re.match(xlsx_name, files):
            return(path + xlsx_name)

def make_industry_comparison_excel(list_of_stocks, dict_of_attribute):
    pass

def fetch_attributes_from_excel(list_of_stocks, dict_of_attribute):
    """
    This will make the industry comparison excel from the existing excel sheets.
    """
    from openpyxl import load_workbook
    for stock in list_of_stocks:
        stock_xls_path = search_for_excel(stock)
        if not stock_xls_path:
            sys.error("Stock neither can be found nor downloaded. Exiting !!!")
        wb = load_workbook(stock_xls_path)
        for key in dict_of_attribute:
            # need to open the xls by key name.
            for sheet in wb:
                if sheet.title == key:
                    # choosing BZ for the last col from row 2
                    #cell_headers = sheet['A2':'BZ2']
                    #print(cell_headers)
                    #for cell in cell_headers:
                    for attribute in dict_of_attribute[key]:
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
                                        values.append(sheet[column + str(index)].value)
                                    key_value_pair = {key:value}
                                    print(key_value_pair)
                        
                        
def add_attribute():
    pass


def add_stocks():
    pass


def main():
    """
    For now this is only for testing the codes.
    """
    stock_ticker_list = ['BDL']
    dict_of_attribute = {"Standalone_Balance_Sheet":[
        "Total Non-Current Liabilities",
        "Current Investments Unquoted Book Value",
        "Equity Share Capital"]}
    make_industry_comparison_excel(stock_ticker_list, dict_of_attribute)


if __name__ == '__main__':
    main()
