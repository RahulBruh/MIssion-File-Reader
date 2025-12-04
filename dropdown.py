import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation


def dropdown():
    BUD_PATH = "Check Request.xlsx"
    CARD_SHEET = "Credit Card Report"
    DATA_SHEET = "Data"
    cf = pd.read_excel(BUD_PATH, header=0, sheet_name=CARD_SHEET)
    df = pd.read_excel(BUD_PATH, header=0, sheet_name=DATA_SHEET)

    wb = load_workbook(BUD_PATH)
    ws = wb[CARD_SHEET]

    for idx, r in cf.iloc[12:29].iterrows():
        label = r.iloc[3]
        if label:
            for value in df.iloc[0]:
                if value == label:
                    pass
        return