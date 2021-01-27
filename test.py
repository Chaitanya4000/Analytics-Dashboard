#import win32com.client as win32
from pathlib import Path
import sys
import pandas as pd  # only used for synthetic data
import numpy as np  # only used for synthetic data
import random  # only used for synthetic data
from datetime import datetime  # only used for synthetic data

win32c = win32.constants


def pivot_table(wb: "Vitesco Data/ChangeRequests.xlsx", ws1: "Sheet1", pt_ws: "M1-M9562", ws_name: "Sheet2", pt_name: "Test", pt_rows: ["M1","M2","M3"], pt_cols: list,
                pt_filters: list, pt_fields: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    # pivot table location
    pt_loc = len(pt_filters) + 2

    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)

    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in (
    (pt_filters, win32c.xlPageField), (pt_rows, win32c.xlRowField), (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1],
                                                field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True