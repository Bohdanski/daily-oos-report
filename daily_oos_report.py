"""
Builds the datasheet for the daily out-of-stock report.
"""

import os
import re
import sys
import glob
import zipfile
import fnmatch
import datetime
import pandas as pd
import numpy as np
import openpyxl

from zipfile import ZipFile
from pandas import DataFrame
from datetime import datetime
from openpyxl import load_workbook


def timestamp():
    """
    Creates a timestamp in DB format.
    """ 
    timestamp = datetime.today().strftime('%Y-%m-%d')
    
    return timestamp


def main():
    """
    Main guts of the script.
    """
    data_dir = ".\\excel\\data\\"
    archive_dir = ".\\excel\\archive\\"
    xl_list = glob.glob(data_dir + "*.xlsx")

    try:
        for xl_file in xl_list:
            workbook = pd.ExcelFile(xl_file)

            if fnmatch.fnmatch(xl_file.lower(), "*base*.xlsx") == True:
                print(f"Creating DataFrame for '{xl_file}'...")
                
                df_base = workbook.parse(0, skiprows=1, header=None)
                df_base.columns = ["dept", 
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
                df_base["itemCode"] = df_base["itemCode"].map('{:0>6}'.format)
                df_base["buyerCode"] = df_base["buyerCode"] * 10
                df_base["itemDesc"] = df_base["itemDesc"] + "   " + df_base["itemSize"]
                
                print(f"'{xl_file}' Successfully processed\n")     
            elif fnmatch.fnmatch(xl_file.lower(), "*short*.xlsx") == True:
                print(f"Creating DataFrame for '{xl_file}'...")
                
                df_shorts = workbook.parse(0, skiprows=1, header=None)
                df_shorts.columns = ["itemDesc", 
                                    "itemCode", 
                                    "yesterdayOOS"]
                df_shorts["itemCode"] = df_shorts["itemCode"].map('{:0>6}'.format)
                df_shorts.drop(columns=["itemDesc"], inplace=True)
                
                print(f"'{xl_file}' Successfully processed\n")        
            elif fnmatch.fnmatch(xl_file.lower(), "*reason*.xlsx") == True:
                print(f"Creating DataFrame for '{xl_file}'...")
                
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
                df_reason["itemCode"] = df_reason["itemCode"].map('{:0>6}'.format)
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
                
                print(f"'{xl_file}' Successfully processed\n")        
            elif fnmatch.fnmatch(xl_file.lower(), "*export*.xlsx") == True:
                print(f"Creating DataFrame for '{xl_file}'...")
                
                to_drop = ["14:HATFIELD NORTH", "1:BRATTLEBORO"]
                
                df_cs = workbook.parse(0, skiprows=3, skipfooter=20, header=None)
                df_cs = df_cs[~df_cs[7].isin(to_drop)]
                df_cs = df_cs.filter([0, 14, 15, 17, 34])
                df_cs.columns = ["custCode", 
                                "poDueDate", 
                                "poApptDate", 
                                "inStock", 
                                "daysOOS"]
                df_cs["itemCode"] = df_cs["custCode"].astype(str).str[9:15]
                df_cs.drop(columns=["custCode"], inplace=True)
                df_cs.drop_duplicates(inplace=True)

                print(f"'{xl_file}' Successfully processed\n")

        for data_file in os.listdir(data_dir):
            if fnmatch.fnmatch(data_file, "*.xlsx") == True:
                print(f"Deleting '{data_file}'...\n")
                os.remove(data_dir + data_file)

        df_join_1 = df_base.merge(df_reason, how="left", on="itemCode")
        df_join_2 = df_join_1.merge(df_shorts, how="left", on="itemCode")
        df_join_3 = df_join_2.merge(df_cs, how="left", on="itemCode")
        
        print("Exporting to Excel...\n")
        df_join_3.to_excel(f".\\excel\\archive\\oos-data-{timestamp()}.xlsx", index=False)

        sys.exit(0)
    except:
        try:
            df_join_1 = df_base.merge(df_reason, how="left", on="itemCode")
            df_join_2 = df_join_1.merge(df_shorts, how="left", on="itemCode")

            df_join_2["poDueDate"] = "NO CS DATA"
            df_join_2["poApptDate"] = "NO CS DATA"
            df_join_2["inStock"] = "NO CS DATA"
            df_join_2["daysOOS"] = "NO CS DATA"
            
            print("Exporting to Excel...\n")
            df_join_2.to_excel(f".\\excel\\archive\\oos-data-{timestamp()}.xlsx", index=False)
        except:
            if not os.path.exists(archive_dir):
                os.makedirs(archive_dir)
            if not os.path.exists(data_dir):
                os.makedirs(data_dir)

            sys.exit(1)


if __name__ == "__main__":
    main()
