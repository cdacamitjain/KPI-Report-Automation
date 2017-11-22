"""
# # DATE		: 03,October 2017
# # AUTHOR		: AMIT.JAIN@LNTTECHSERVICES.COM
# # DESCRIPTION	: This script provides utility methods used by other scripts.
# #
"""
import logging
import os
import win32com.client as win32

import pandas as pd

from Constants import Constants


class Utils:
    logging.basicConfig(format='%(levelname)s:  %(filename)s:%(lineno)s:    %(message)s', level=logging.DEBUG)

    @staticmethod
    def ChangeDirPath(changeTo):
        filePath = ".." + changeTo
        # logging.error(os.path.realpath(filePath))
        if os.path.isdir(filePath):
            os.chdir(filePath)
        else:
            logging.error("Directory '" + changeTo + "' is not Present..!!")
            logging.error("Please create dir '" + changeTo + "' First to processed further. Script is aborted...!!")

    @staticmethod
    def ReadCSV(file):
        filePath = os.path.abspath(file)
        # logging.debug("Read CSV: " + filePath)
        if os.path.isfile(filePath):
            return pd.read_csv(filePath, index_col=Constants.CONFIG_HEADER[0])
        else:
            logging.error(filePath + " is absent.")
            logging.error(" Script is aborted...!!")

    @staticmethod
    def AppsNameList(Appslist):
        appName_List = []
        for val in Appslist:
            appName_List.append(val.replace(" ", "_"))
        return appName_List

    @staticmethod
    def AppsNameWithTestCount(appName_List):
        uniqueAppName_List = Utils.GetUniqueAppName(appName_List)
        appTC_List = []
        for app in uniqueAppName_List:
            appTC = 0
            for appName in appName_List:
                if app == appName:
                   appTC+=1
            appTC_List.append(appTC)
        return dict((zip(uniqueAppName_List, appTC_List)))

    @staticmethod
    def GetUniqueAppName(appName_List):
        uniqueAppName_List = []
        for appName in appName_List:
            if appName not in uniqueAppName_List:
                uniqueAppName_List.append(appName)
        return uniqueAppName_List

    @staticmethod
    def MakeApplicationDictionary(df):
        Apps_List = df.index.values.tolist()
        AppNameMapped_Dic = Utils.AppsNameWithTestCount(Utils.AppsNameList(Apps_List))
        uniqueAppName_List = Utils.GetUniqueAppName(Apps_List)
        return AppNameMapped_Dic,uniqueAppName_List

    @staticmethod
    def SupportedSheetName(key, margin):
        actual_len = len(key)
        supported_len_diff = Constants.SHEET_NAME_MAXSIZE - (actual_len + margin)
        if supported_len_diff < 0:
            return key[:supported_len_diff]
        else:
            return key

    @staticmethod
    def SheetNameFormator(AppCount,AppName,MD=""):
        return "TKPI%03d_%s" % (AppCount, AppName) + MD

    @staticmethod
    def CalculateMaxRunCount(df_sorted):
        maxRun_count = 0
        for df in df_sorted:
            runcount = Utils.CalculateRunCount(df)
            if runcount > maxRun_count:
                maxRun_count = runcount
        return maxRun_count

    @staticmethod
    def CalculateRunCount(df):
        header_list = df.columns.values.tolist()
        run_list = filter(lambda x: Constants.MEASUREMENT in x, header_list)
        temp = (len(run_list) / Constants.CONFIGURATIONS_COUNT)
        return temp

    @staticmethod
    def GetBundlesList(bundleFiles_list):
        bundles_list = []
        for b in bundleFiles_list:
            bundles_list.append(b.split(".")[0])
        return bundles_list

    @staticmethod
    def GetColumnName(n):
        div = n+1
        string = ""
        while div > 0:
            module = (div - 1) % 26
            string = chr(65 + module) + string
            div = int((div - module) / 26)
        return string

    @staticmethod
    def get_cell_range(row_index, col_index, bundle_count, whole_range=False):
        if bundle_count > 2:
            update_row_index = 0
            if whole_range:
                update_row_index = 3
            range_cell = Utils.GetColumnName(col_index + 1) + str(row_index) + ":" + Utils.GetColumnName(
                col_index + bundle_count - 1) + str(row_index + update_row_index)  # P3:R3
        else:
            update_row_index = 0
            if whole_range:
                update_row_index = 3
            range_cell = Utils.GetColumnName(col_index + 1) + str(row_index) + ":" + Utils.GetColumnName(
                col_index + 1) + str(row_index + update_row_index)  # P3:R3

        return range_cell

    @staticmethod
    def Convert_XLS_to_XLSX_File(fileName, config):

        cd = os.path.dirname(os.path.abspath(fileName))
        file = os.path.join(cd, config)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        wb = None
        try:
            if os.path.isfile(file):
                wb = excel.Workbooks.Open(file)
                wb.SaveAs(file + "x", FileFormat=51)
                wb.Close()
                excel.Application.Quit()
                os.remove(file)
            else:
                print fileName + " is not Present...!!"
        except :
            if wb is not None:
                wb.Close()
            excel.Application.Quit()
