from ast import ExceptHandler
from decimal import Decimal
import os
from pathlib import Path
import sys
from flask.config import T
import requests
from read_holdings_excel import get_tlv_data
import ibkr_papers
from pandas import DataFrame
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows  
from openpyxl import load_workbook  
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string


exchange_rates = {}
rate_locations = {'USD':"L63",'EUR':"L64",'JPY':'L65','SGD':'L66'}

def get_exchange_rates():
    Currencies = ['eur','usd','jpy','sgd']
    for curr in Currencies:
        url = f"https://api.frankfurter.dev/v2/rate/{curr}/ils"
        res = requests.get(url)
        data = res.json(parse_float=Decimal)
        exchange_rates[curr.upper()] =  data['rate']
    #print(exchange_rates)

    
def column_formula(ws,formula:str, col:int, start_row:int, end_row:int,number_format=None):
    """Apply a formula to a column in an Excel worksheet."""
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row=row, column=col)
        if(number_format != None):
            cell.number_format = number_format
        cell.value = formula.replace("##", str(row))

def write_row(ws,row,data):
    for col in range(1,len(data)):
        cell = ws.cell(row=row, column=col)
        cell.value = data[col -1]

def count_rows(ws):
    num_rows = 0
    cell = ""
    for row in range(2, ws.max_row):
        cell = str(ws.cell(row=row, column=2).value).strip()
        print(f"Row {row}, Column 2: {cell}")
        if(cell.isnumeric()):
            num_rows += 1
        else:
            break
    return num_rows


def add_df_columns_to_excel(ws, df,df_heading,ws_column):
    """Add columns from a DataFrame to an Excel worksheet."""
    #yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for idx, col in enumerate(df.columns):
        if col in df_heading:
            col_index = ws_column + df_heading.index(col)
            for row_idx, value in enumerate(df[col], start=2):
                ws.cell(row=row_idx, column=col_index, value=value)
                #ws.cell(row=row_idx, column=col_index).fill = yellow_fill  # Highlight the cell with yellow color



def make_performance_xl(file_path: str, output_path: str,paper_composition: str) -> None:
    """Read holdings from an Excel file, enrich them with Bizportal data, and save to a new Excel file."""

    template_path = f"{file_path.replace(sys.argv[1],"excel_files")}\\Analysis_template.xlsx"

    #holdings = DataFrame()
    # read an excel file from the output path
    workbook = load_workbook(template_path)  
    # get active sheet
    ws = workbook.active
    new_out = template_path.replace("_template.xlsx", "_MAGIC.xlsx").replace("excel_files",output_path)
    try:
        workbook.save(new_out)
    except:
        print(f"Cant write to {new_out} - make sure the excel file is not open.")
        sys.exit()

    holdings = get_tlv_data(f"{file_path}/אחזקות.xls")
    ibkr = ibkr_papers.get_ibkr_sharpe_and_performance()
    data_cell_range = "A2:AH59"
    #print(holdings.columns)
    # 'מספר נייר', 'שם הנייר', 'שווי אחזקה ב ₪', '%  מהתיק', 'פרופיל חשיפה',
    #   'תאריך פעולה אחרונה', '% שינוי יומי', 'תשואה 12 חודשים', 'רווח ב-%',
    #   'רווח ב ₪', 'שער בסיס', 'מס התיק', 'fund_name', 'management_fee',
    #   'three_year_return_percent', 'sharpe_ratio_12_months', 'bizportal_url'
    tlv_df = DataFrame(holdings, columns=['שם הנייר', 'מספר נייר', 'שווי אחזקה ב ₪','three_year_return_percent', 'תשואה 12 חודשים', 'פרופיל חשיפה',
                                      'sharpe_ratio_12_months','management_fee',"מטבע מסחר"])

    # 	ClientAccountID	Symbol	Quantity	MarkPrice	PositionValue	CostBasisPrice	CostBasisMoney	FifoPnlUnrealized
    # 	CurrencyPrimary	AssetClass	PercentOfNAV	sharpe_ratio_12_months	performance_1y	performance_3y

    ibkr_df = DataFrame(ibkr, columns=["Name", "Symbol","PositionValue","performance_3y","performance_1y","exposure","sharpe_ratio_12_months",
                                       "management_fee","CurrencyPrimary"])
    #print(ibkr_df)

    ibkr_df.rename(columns={"Name":'שם הנייר', "Symbol":'מספר נייר',"PositionValue":'שווי אחזקה ב ₪',
                            "performance_3y":'three_year_return_percent',"performance_1y":'תשואה 12 חודשים'
                            ,"exposure": 'פרופיל חשיפה',"sharpe_ratio_12_months":'sharpe_ratio_12_months',
                                       "management_fee":"management_fee","CurrencyPrimary":"מטבע מסחר"},inplace=True)
    # convert ibkr values to ILS
    for index,row in ibkr_df.iterrows():
        if row["מטבע מסחר"] in exchange_rates.keys():
            ibkr_df.loc[index,'שווי אחזקה ב ₪']  =  f"={str(row['שווי אחזקה ב ₪'])}*{rate_locations[row["מטבע מסחר"]]}"
    
    df  = pd.concat([tlv_df,ibkr_df],ignore_index=True)



    #print(ibkr_df)
    #print(df)
    cell = find_value(ws,"Amargeti house")
    amargeti_row_values = []
    max_row = len(df)+1
    amgti_row = max_row+1
    amgti_tot = f"M{amgti_row}"
    if(cell != None):
        row_num = cell.row
        amargeti_row_values = [cl.value for cl in ws[row_num]]
        for i in range(0,len(amargeti_row_values)) :
            val = amargeti_row_values[i]
            if("=" in str(val)):
                s = str(val)
                s_replaced = s.replace("38",str(amgti_row))
                amargeti_row_values[i] = s_replaced
                
        




    
    #adjust_rows(ws, df)
    clear_range(ws,data_cell_range)
    add_df_columns_to_excel(ws, df,'שם הנייר', 1)
    add_df_columns_to_excel(ws, df,'מספר נייר', 2)
    add_df_columns_to_excel(ws, df,'three_year_return_percent', 5)
    add_df_columns_to_excel(ws, df,'תשואה 12 חודשים', 6)
    add_df_columns_to_excel(ws, df,'פרופיל חשיפה', 7)
    add_df_columns_to_excel(ws, df,'sharpe_ratio_12_months', 8)
    add_df_columns_to_excel(ws, df,'management_fee', 9)
    add_df_columns_to_excel(ws, df,'מטבע מסחר', 10)
    add_df_columns_to_excel(ws, df,'שווי אחזקה ב ₪', 13)

    write_row(ws,max_row+1,amargeti_row_values)
    ws["AG60"] = f"={amgti_tot}" 
    set_formula_definitions(ws,len(df))

    compo = pd.read_excel(paper_composition)
    #  שם הקרן       ניר  מדינה  קונצרני  assets  מניות ישראל  מניות חול  אגח חול  check  נקוב $
    #print(compo)
    #print(df["מספר נייר"].values)
    for index, row in compo.iterrows():
        #print(row.values)
        #print(f"{row['ניר']} - {row['מדינה']}  ")
        num_paper = str(row['ניר']).strip()
        df.loc[df['מספר נייר'] == num_paper, 'מדינה'] =  row.loc['מדינה']
        df.loc[df['מספר נייר'] == num_paper, 'קונצרני'] =  row.loc['קונצרני']
        df.loc[df['מספר נייר'] == num_paper, 'assets'] =  row.loc['assets']
        df.loc[df['מספר נייר'] == num_paper, 'אגח חול'] =  row.loc['אגח חול']
        df.loc[df['מספר נייר'] == num_paper, 'מניות חול'] =  row.loc['מניות חול']
        df.loc[df['מספר נייר'] == num_paper, 'מניות ישראל'] =  row.loc['מניות ישראל']
        df.loc[df['מספר נייר'] == num_paper, 'נקוב $'] =  row.loc['נקוב $']
        df.loc[df['מספר נייר'] == num_paper, 'נקוב EU'] =  row.loc['נקוב EURO']

        #print(df.loc[df['מספר נייר'] == row['ניר'], 'מדינה'])
    #print(df['מדינה'])
    add_df_columns_to_excel(ws, df,'מדינה', 16)
    add_df_columns_to_excel(ws, df,'קונצרני', 17)
    add_df_columns_to_excel(ws, df,'assets', 18)
    add_df_columns_to_excel(ws, df,'אגח חול', 21)
    add_df_columns_to_excel(ws, df,'מניות חול', 20)
    add_df_columns_to_excel(ws, df,'מניות ישראל', 19)
    add_df_columns_to_excel(ws, df,'נקוב $', 23)
    add_df_columns_to_excel(ws, df,'נקוב EU', 24)


    workbook.save(new_out)


def set_formula_definitions(ws,num_rows):
    formulas = {}
    formulas["K"] = ("=E##-3*I##/100",None)
    formulas["L"] = ("=F##-I##/100",     "0.0%")
    formulas["N"] = ("=M##/$M$60",None)
    formulas["O"] = ("=M##/$M$61",None)
    formulas["V"] = ("=SUM(P##:U##)",None)
    formulas["Y"] = ("=P##*M##",None)
    formulas["Z"] = ("=Q##*M##",None)
    formulas["AA"] =("=S##*M##",None)
    formulas["AB"] =("=T##*M##" ,None)
    formulas["AC"] =("=U##*M##" ,None)
    formulas["AD"] =("=R##*M##" ,None)
    formulas["AE"] =("=W##*M##" ,None)
    formulas["AF"] =("=X##*M##" ,None)
    for col in formulas.keys():
        col_num = column_index_from_string(col)
        column_formula(ws,formulas[col][0],col_num,2,num_rows+2,formulas[col][1])
    for curr in rate_locations.keys():
        ws[rate_locations[curr]] = exchange_rates[curr]


def find_value(ws,target_value):
    found_cell = None

    # Iterate through all populated rows and columns
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == target_value:
                found_cell = cell
                break  # Break out of the inner loop
        if found_cell:
            break  # Break out of the outer loop
    return found_cell




def clear_range(ws,range):
    cell_range = ws[range]
    # Iterate and clear cell   values
    for row in cell_range:
        for cell in row:
            cell.value = None


def adjust_rows(ws, df):
    count = count_rows(ws)
    print(f"Number of rows with numeric values in column 2: {count}")
    df_row_count = len(df)
    print(f"Number of rows in the DataFrame: {df_row_count}")
    if(df_row_count < count):
        # need to delete the last tlv rows from the worksheet    
        for row in range(df_row_count+2, df_row_count - (count - df_row_count)+2, -1):
                for col in (1,2,5,6,7,8,9,13):
                    ws.cell(row=row, column=col).value = ""

                print(f"cleared row {row} from the worksheet.")
    elif(df_row_count > count):
        # need to add new rows to the worksheet
        for row in range(count+2, df_row_count+2):
            ws.insert_rows(row)
            print(f"Inserted row {row} into the worksheet.")
        
    # 5. Save the Excel file
    #workbook.save(output_file_path)

def get_exe_directory():
    """ Get the actual directory where the .exe file is running """
    if hasattr(sys, 'frozen'):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

if __name__ == "__main__":
    get_exchange_rates()
    script_dir = get_exe_directory()
    in_path = sys.argv[1]

    input_file_path = f"{script_dir}\\{in_path}"
    output_file_path = in_path
    paper_composition = f"{script_dir}\\excel_files\\הרכב ניירות.xlsx"
    make_performance_xl(input_file_path, output_file_path,paper_composition)    


