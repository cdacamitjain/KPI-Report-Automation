"""
# # DATE		    : 16th,October 2017
# # AUTHOR		    : AMIT.JAIN@LNTTECHSERVICES.COM
# # AUTHOR-MODIFIED : HASTI.SHAH@LNTTECHSERVICES.COM
# # DESCRIPTION	    : This script is used to create template file and
# # 			combine all the Configs from each available bundles into CSV File.
"""

from openpyxl import load_workbook

import pandas as pd
import os.path
import sys
import csv

from Constants import Constants
from Utils import Utils


class CreateCombineConfigs:
    __Const = None
    __FirstTimeFlag = True
    __MAX_RUN_COUNT = 0
    unique_ids =  []
    template_list = []

    def __init__(self):
        self.__Const = Constants()
        self.maxRunCount = 0
        self.finalRows = []
        self.firstTimeFlag = True
        self.count = 0
        self.unique_ids =  []
        self.__MAX_RUN_COUNT = 0
        self.template_list.append([u'Application Name', u'Test Case ID', u'Test Case Level', u'Test Case Category', u'Test Case Description'])

    def CreateCombineConfigsCSV(self):
        print "Creating CombineConfigs for each Bundle Present"
        # print "###############################################"
        df_ConfigTemp = self.ReadConfigTemplate()
        # print df_ConfigTemp.head()
        Utils.ChangeDirPath(self.__Const.BUNDLES_PATH)
        listofBundleDir = [d for d in list(reversed(os.listdir(os.getcwd()))) if os.path.isdir(d)]

        if len(listofBundleDir) != 0:
            for bundleDirName in listofBundleDir:
                if bundleDirName.startswith("Bundle"):
                    # print bundleDirName
                    listofConfigsFile = os.listdir(bundleDirName)
                    # print "listofConfigsFile-->",listofConfigsFile
                    CombineConfigsBundleFile = "..\\" + self.__Const.REFERRAL_DATA_PATH + "\\" + bundleDirName + ".csv"
                    flag = True
                    if os.path.isfile(CombineConfigsBundleFile):
                        # print "Updating already exiting.\t'" + os.path.abspath(CombineConfigsBundleFile) + "'"
                        print "Updating already exiting.\t'" + CombineConfigsBundleFile + "'"
                        flag = False
                    # print "bundleDirName", bundleDirName
                    # self.CountMaxRun(bundleDirName)


                    dfl = []
                    header_row = []
                    for i in range(self.__MAX_RUN_COUNT):
                        header_row.append("Run" + str(i + 1))
                    header_row += [u'Avg Response Time', u'Initial Condition Issue', u'Application Freeze Issue']
                    # print "len(header_row)",len(header_row)

                    for config in self.__Const.CONFIG_LIST:
                        if config in listofConfigsFile:
                            readFilePath = ".\\" + bundleDirName + "\\" + config
                            # print "readFilePath",readFilePath

                            df_read_config = pd.read_excel(readFilePath)

                            df_rows_list = df_read_config.values.tolist()[1:]
                            # print "len(df_rows_list)",len(df_rows_list),self.__MAX_RUN_COUNT,len(df_ConfigTemp.values.tolist())
                            # New Implementation
                            final_config_list = []
                            for t_row in df_ConfigTemp.values.tolist():
                                # print t_row
                                count = 0
                                for o_row in df_rows_list:
                                    if t_row[1] in o_row:
                                        row = o_row[len(Constants.EXCEL_HEADER_START_LIST):-1]
                                        missing_run = len(header_row) - len(row)

                                        if missing_run < 0:
                                            final_config_list.append(row)
                                        else:
                                            rev_index = -2
                                            for k in range(missing_run):
                                                rev_index = rev_index -1
                                                row.insert(rev_index,Constants.DEFAULT_VALUE)
                                            # print row
                                            final_config_list.append(row)
                                        break
                                    count+=1
                                if count == len(df_rows_list):
                                    default_list = []
                                    for i in range(self.__MAX_RUN_COUNT+len(Constants.EXCEL_HEADER_END_LIST)-1):
                                        default_list.append(Constants.DEFAULT_VALUE)
                                    final_config_list.append(default_list)


                            dfl.append(pd.DataFrame(final_config_list, columns=header_row))


                            #  New Implementation
                            # break

                        else:
                            final_config_list = []
                            for j in range(len(df_ConfigTemp.values.tolist())):
                                default_list = []
                                for i in range(self.__MAX_RUN_COUNT+len(Constants.EXCEL_HEADER_END_LIST)-1):
                                    default_list.append(Constants.DEFAULT_VALUE)
                                final_config_list.append(default_list)

                            dfl.append(pd.DataFrame(final_config_list, columns=header_row))

                    df_result = pd.concat([df_ConfigTemp, dfl[0], dfl[1], dfl[2], dfl[3]], axis=1)
                    df_result.to_csv(CombineConfigsBundleFile, index=False)

                    if flag:
                        print "Created new file.\t\t'" + os.path.abspath(CombineConfigsBundleFile) + "'"

                    # if bundleDirName.startswith("Bundle"):
                    #     break
                else:
                    # print "Not Bundle",bundleDirName
                    pass


        else:
            print "Bundles are absent."
            sys.exit("ERROR: " + "Bundle Directories are need to be created. Script is aborted...!!")
        print "CombineConfigs for each Bundle are created successfully inside: .'" + self.__Const.REFERRAL_DATA_PATH + "'"

    def ReadConfigTemplate(self):
        filePath = os.path.abspath(self.__Const.CONFIG_TEMPLATE_FILE)
        if os.path.isfile(filePath):
            df_configTemplate = pd.read_csv(filePath)
            return df_configTemplate
        else:
            print filePath + " is absent."
            sys.exit("ERROR: Script is aborted...!!")

    def CheckConfigTemplate(self):
        Utils.ChangeDirPath(self.__Const.REFERRAL_DATA_PATH)
        # if self.__Const.CONFIG_TEMPLATE_FILE in os.listdir(os.getcwd()):
        print "'" + os.path.abspath(self.__Const.CONFIG_TEMPLATE_FILE) + "' is Updating."
        self.CreateConfigTemplate()

    def CreateConfigTemplate(self):
        rows = []
        print "'" + self.__Const.CONFIG_TEMPLATE_FILE + "' File is Updating..."
        Utils.ChangeDirPath(self.__Const.BUNDLES_PATH)
        listofBundleDir = os.listdir(os.getcwd())

        for bundle in listofBundleDir:

            if bundle.startswith("Bundle"):
                for config in os.listdir(bundle):
                    if config in self.__Const.CONFIG_LIST:
                        readFilePath = ".\\" + bundle + "\\" + config
                        self.ReadConfigForTemplate(readFilePath)

                        # print len(rows)
        # print "self.__MAX_RUN_COUNT",self.__MAX_RUN_COUNT
        #####################################

        FinalListOfRows = []
        FinalListOfRows.append(self.template_list[0])
        for l in self.template_list[1:]:
            FinalListOfRows.append(l[0])

        self.WriteToCSV(FinalListOfRows)

    def WriteToCSV(self, rows):
        Utils.ChangeDirPath(self.__Const.REFERRAL_DATA_PATH)
        with open(self.__Const.CONFIG_TEMPLATE_FILE, 'wb') as csvfile:
            writer = csv.writer(csvfile, delimiter=',')
            for row in rows:
                writer.writerow(row[:5])

        print "'" + self.__Const.CONFIG_TEMPLATE_FILE + "' is Created Successfully."
        self.__Const.setRowCount(len(rows)+1)


    def ReadConfigForTemplate(self,filePath):
        dataframe = pd.read_excel(filePath)
        # print dataframe.columns.values.tolist()
        cal_run = len(dataframe.columns.values.tolist()) - Constants.CONSTANT_COLUMNS_COUNT_READ_EXCEL

        if self.__MAX_RUN_COUNT < cal_run:
            self.__MAX_RUN_COUNT = len(dataframe.columns.values.tolist()) - Constants.CONSTANT_COLUMNS_COUNT_READ_EXCEL

        temp_df = dataframe.iloc[:,:5]

        # print list(temp_df.columns.values)
        temp_list = [x for x in list(temp_df.iloc[:,1]) if str(x) != 'nan']

        for uid in temp_list:
            if uid not in self.unique_ids:
                self.template_list.append(temp_df.loc[temp_df[u'Test Case ID'] == uid].values.tolist())
                self.unique_ids.append(uid)
            else:
                pass

    def getMaxRunCount(self):
        return self.__MAX_RUN_COUNT
