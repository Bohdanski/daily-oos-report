"""
Program test script to be used for updates and bug fixes.
"""

import os
import re
import glob
import zipfile
import fnmatch
import datetime
from zipfile import ZipFile
import pandas as pd
import numpy as np
from pandas import DataFrame
import openpyxl
from openpyxl import load_workbook

def create_timestamp():
    """
    Creates a timestamp in DB format.
    """
    today = datetime.date.today()
    year = today.year
    month = today.month
    day = today.day

    timestamp = f"{str(year)}-{str(month)}-{str(day)}"

    return timestamp

def find_sheet(sheet_list, sheet_name):
    """
    If a workbook matches the desired name
    store each worksheet into a list.
    """
    return [x for x in sheet_list if re.search(sheet_name.lower(), x.lower())]

def main():
    """
    Main guts of the script.
    """
    data_dir = ".\\excel\\data\\"
    archive_dir = ".\\excel\\archive\\"
    xl_list = glob.glob(data_dir + "*.xlsx")

    try:
        if not os.path.exists(archive_dir):
            os.makedirs(archive_dir)
            if not os.path.exists(data_dir):
                os.makedirs(data_dir)
            exit()
    finally:
        for xl_file in xl_list:
            sheets = pd.ExcelFile(xl_file).sheet_names
            workbook = pd.ExcelFile(xl_file)

            if fnmatch.fnmatch(xl_file.lower(), "*reason*.xlsx") == True:
                df_reason = workbook.parse(0, skiprows=2, header=None)
                df_reason.columns = ["dept",
                                     "category",
                                     "itemDesc",
                                     "itemCode",
                                     "outOfStock",
                                     "manufacIssue",
                                     "disc",
                                     "other",
                                     "newItemIssue"]
                df_reason["max"] = df_reason[[df_reason.columns[4],
                                              df_reason.columns[5],
                                              df_reason.columns[6],
                                              df_reason.columns[7],
                                              df_reason.columns[8]]].max(axis=1)
                df_reason.loc[df_reason["max"] == df_reason["outOfStock"], "primaryReason"] = "Out Of Stock"
                df_reason.loc[df_reason["max"] == df_reason["manufacIssue"], "primaryReason"] = "Manufacturer Issue"
                df_reason.loc[df_reason["max"] == df_reason["disc"], "primaryReason"] = "Discontinued"
                df_reason.loc[df_reason["max"] == df_reason["other"], "primaryReason"] = "Other"
                df_reason.loc[df_reason["max"] == df_reason["newItemIssue"], "primaryReason"] = "New Item Issue"
                df_reason.sort_values(by=["max"], ascending=False, inplace=True)
                df_reason.drop(columns=["dept",
                                        "category",
                                        "itemDesc",
                                        "outOfStock",
                                        "manufacIssue",
                                        "disc",
                                        "other",
                                        "newItemIssue",
                                        "max"], inplace=True)
            elif fnmatch.fnmatch(xl_file.lower(), "*short*.xlsx") == True:
                df_shorts = workbook.parse(0, skiprows=1, header=None)
                df_shorts.columns = ["itemDesc",
                                     "itemCode",
                                     "yesterdayOOS"]
                df_shorts.drop(columns=["itemDesc"], inplace=True)
            elif fnmatch.fnmatch(xl_file.lower(), "*detail*.xlsx") == True:
                df_detail = workbook.parse(0, skiprows=1, header=None)
                df_detail.columns = ["dept",
                                     "category",
                                     "itemDesc",
                                     "itemCode",
                                     "itemSize",
                                     "pvtLblFlag",
                                     "buyerCode",
                                     "invUnitShipped",
                                     "invCaseShipped",
                                     "storeOrdProdQty",
                                     "shortedQty",
                                     "grossSvcLvl",
                                     "netSvcLvl"]
                df_detail["buyerCode"] = df_detail["buyerCode"] * 10
                df_detail["itemDesc"] = df_detail["itemDesc"] + "   " + df_detail["itemSize"]

        df_join_1 = df_detail.merge(df_reason, how="left", on="itemCode")
        df_join_2 = df_join_1.merge(df_shorts, how="left", on="itemCode")

        df_join_2["poDueDate"] = "NO CS DATA"
        df_join_2["poApptDate"] = "NO CS DATA"
        df_join_2["inStock"] = "NO CS DATA"
        df_join_2["daysOOS"] = "NO CS DATA"
        df_join_2["InstockOrDueDate"] = "NO CS DATA"

        df_join_2.to_excel(f".\\excel\\archive\\oos-data-{create_timestamp()}.xlsx")

if __name__ == "__main__":
    main()
