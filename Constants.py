# -*- coding: utf-8 -*-
"""
# # DATE		: 03,October 2017
# # AUTHOR		: AMIT.JAIN@LNTTECHSERVICES.COM
# # DESCRIPTION	: This script is used to for Static Constants variables.
# #
"""


class Constants:
    # # File Name Constants
    CONFIG_LIST = ['Config1_KeyON_OFF.xlsx',
                   'Config2_KeyON_NoUDW.xlsx',
                   'Config3_KeyON_MonkeyExecution.xlsx',
                   'Config4_KeyON_HeavyUDW.xlsx']

    CONFIG_TEMPLATE_FILE = 'Configuration_Template.csv'
    WILDCARD_BUNDLE_FILE = "Bundle*.csv"
    FINAL_REPORT_FILE = 'KPI_Final_Report.xlsx'

    # # Path Constants
    REPORT_PATH = "\Report"
    SCRIPT_PATH = "\Scripts"
    BUNDLES_PATH = "\Bundles"
    REFERRAL_DATA_PATH = "\Referral_Data"

    # # Variable Constants
    SHEET_NAME = "Application_Details"  # Summary
    CONFIG_HEADER = [u'Application Name', u'Test Case ID', u'Test Case Level', u'Test Case Category',
                     u'Test Case Description', u'Min', u'Max', u'Average', u'Initial Condition Issue',
                     u'Application Freeze Issue']
    EXCEL_HEADER_START_LIST = [u'Application Name', u'Test Case ID', u'Test Case Level', u'Test Case Category', u'Test Case Description']
    EXCEL_HEADER_END_LIST = [u'Avg Response Time', u'Initial Condition Issue', u'Application Freeze Issue', u'Comments']

    CONSTANT_COLUMNS_COUNT_READ_EXCEL = len(EXCEL_HEADER_START_LIST ) + len(EXCEL_HEADER_END_LIST)

    DEFAULT_ROW = ["NA", "NA", "NA"]
    DEFAULT_VALUE = "NA"
    MD_TAB_HEADER = [u'Test Case No.', u'Test Category', u'Test Case ID', u'Test Type', u'Configurations']
    MD_TAB_HEADER_COUNT = len(MD_TAB_HEADER)
    APP_TAB_HEADER = [u'Test Case No.', u'Test Category', u'Test Case ID', u'Test Type', u'Test Case Description',
                      'Result \n PASS/FAIL', 'Comments']
    CONFIGURATIONS_LIST = ['Config1', 'Config2', 'Config3', 'Config4']
    BUNDLE_FIXED_HEADER = ['Test Case ID', 'Test Case Level', 'Test Case Category', 'Test Case Description']
    BUNDLE_FIXED_HEADER_COUNT = len(BUNDLE_FIXED_HEADER)
    AVERAGE_RESPONSE_TIME = 'Average Response Time'
    INITIAL_CONDITION_ISSUE = 'Initial Condition Issue'
    APPLICATION_FREEZE_ISSUE = 'Application Freeze Issue'
    ISSUE_HEADER_LIST = [AVERAGE_RESPONSE_TIME, INITIAL_CONDITION_ISSUE, APPLICATION_FREEZE_ISSUE]
    ISSUE_HEADER_COUNT = len(ISSUE_HEADER_LIST)
    TEST_CASE_DESCRIPTION_LIST = ["Response time Improvement/No Change - Config 1",
                                  "Response time Improvement/No Change - Config 2",
                                  "Response time Improvement/No Change - Config 3",
                                  "Response time Improvement/No Change - Config 4",
                                  "Response time not impacted by display for Available Configurations",
                                  "Application not Freezing/Responding",
                                  "First response time measurement taken comparable to consecutive ones"]
    TC_DESCRIPTION_WIDTH = 70
    REPORT_DESCRIPTION_WIDTH = 40
    CELL_WIDTH = 20

    PERFORMANCE_COMPARED_TO_ALL_RELEASES = "Performance compared to all releases"
    PERFORMANCE_COMPARED_BETWEEN_ALL_CONFIGURATION = "Performance compared between all configuration"
    FIRST_MEASUREMENT_GREATER_THAN_2X_LONGER_FROM_CONSECUTIVE_ONES = "First measurement >2x longer from consecutive ones"
    APPLICATION_FREEZING_NOT_RESPONDING = "Application Freezing/Not Responding"
    PERFORMANCE_COMPARED_TO_LAST_RELEASE = "Performance compared to last release"
    IMPROVEMENT_PERCENTAGE_COMPARED_TO_CURRENT_RELEASE = "Improvement percentage compared to current release"

    PERFORMANCE_HEADER_LIST = [PERFORMANCE_COMPARED_TO_ALL_RELEASES, PERFORMANCE_COMPARED_BETWEEN_ALL_CONFIGURATION ,
                               FIRST_MEASUREMENT_GREATER_THAN_2X_LONGER_FROM_CONSECUTIVE_ONES,
                               APPLICATION_FREEZING_NOT_RESPONDING,IMPROVEMENT_PERCENTAGE_COMPARED_TO_CURRENT_RELEASE]

    PERFORMANCE_HEADER_COUNT = len(PERFORMANCE_HEADER_LIST)

    BUNDLES_PERFORMANCE_HEADER = [PERFORMANCE_COMPARED_TO_ALL_RELEASES, PERFORMANCE_COMPARED_TO_LAST_RELEASE]
    BUNDLES_PERFORMANCE_TITLE = "Bundles Performance"



    TEST_REPORT_DESCRIPTION_TITLE = "GHMI APPLICATION PERFORMANCE TEST PERFORMED WITH FOLLOWING HARDWARE AND SOFTWARE BUILD"
    SUMMARY_MD_TAB_TITLE =  "System Level Performance Compared to all Releases (Bundle to Bundle Comparison)"
    TEST_REPORT_DESCRIPTION_INDEX_LIST =\
        ["Validation Starting Date",
            "Validation End Date",
            "Tractor Application Software Version",
            "CNHi Firmware Version",
            "Accenture CAN manager Version",
            "Core Software Bundle Version",
            "Tractor Application Software Date of Release",
            "Previous Software Version",
            "Hardware Serial Number",
            "Hardware Version",
            "SIMULATOR Type",
            "SIMULATOR SOFTWARE Firmware Version"]



    CONFIGURATIONS_DETAIL_TITLE = "Configuration Details"
    CONFIGURATIONS_DESCRIPTION_HEADER = "Configuration Description"
    CONFIGURATIONS_NAME_HEADER = "Configuration Name"
    CONFIGURATIONS1_DESCRIPTION = u"KPI Test with Key On/Off functionality after booting up device for Every test case"
    CONFIGURATIONS2_DESCRIPTION = u"KPI Test with Key On only once for all test cases after booting up device (No UDW's on Runscreen and LHA)"
    CONFIGURATIONS3_DESCRIPTION = u"KPI test with Key On only once for all test cases and Stress test done with Monkey test before test case execution"
    CONFIGURATIONS4_DESCRIPTION = u"KPI Test with Key On only once for all test cases after booting up device (Heavy CPU UDW's on Runscreen and LHA)"
    CONFIGURATIONS_DESCRIPTION_LIST = [CONFIGURATIONS1_DESCRIPTION, CONFIGURATIONS2_DESCRIPTION,
                                       CONFIGURATIONS3_DESCRIPTION, CONFIGURATIONS4_DESCRIPTION]

    CURRENT_RELEASE_HEADER = "CURRENT RELEASE - RESULT"
    PREVIOUS_RELEASE_HEADER = "PREVIOUS RELEASE - RESULT"

    SUMMARY_MD_TAB_HEADER = [CONFIGURATIONS_NAME_HEADER,CURRENT_RELEASE_HEADER,PREVIOUS_RELEASE_HEADER]

    HYPERLINK_FORMULA_STRING_FORMAT = '=HYPERLINK("#{}","{}")'
    PERFORMANCE_ROW_INDEX_REFERENCE_TABLE = 100
    NA_ROW_INDEX_REFERENCE = 1000

    STATUS_IMPROVED = "Improved Test"
    STATUS_DEGRADED = "Degraded Test"
    STATUS_SIMILAR = "Similar Test"
    STATUS_BETTER = "Better Test"
    STATUS_WORSE = "Worse Test"
    STATUS_NOT_EXECUTE = "Not Compared Test"
    STATUS_NEW_EXECUTE = "New Test"

    STATUS_TEST_CASE_COMPARED_FOR_AVERAGE_RESULTS = "Test Cases Compared for Average Results"
    STATUS_TEST_CASE_NOT_COMPARED_FOR_AVERAGE_RESULTS = "Test Cases Not Compared"

    AVERAGE_RESULTS_HEADER = [STATUS_TEST_CASE_COMPARED_FOR_AVERAGE_RESULTS,STATUS_TEST_CASE_NOT_COMPARED_FOR_AVERAGE_RESULTS]

    PERFORMANCE_STATUS_HEADER = [STATUS_IMPROVED, STATUS_DEGRADED, STATUS_SIMILAR,STATUS_NOT_EXECUTE,STATUS_NEW_EXECUTE]
    PERFORMANCE_STATUS_HEADER_DICT = {STATUS_IMPROVED:"Improved",STATUS_DEGRADED:"Degraded",STATUS_SIMILAR : "Similar",
                                      STATUS_NOT_EXECUTE : "NA",STATUS_NEW_EXECUTE : "New"}

    PERFORMANCE_STATUS_RESULT_HEADER = [STATUS_BETTER,STATUS_WORSE,STATUS_SIMILAR,STATUS_NOT_EXECUTE]
    PERFORMANCE_STATUS_RESULT_HEADER_DICT = {STATUS_SIMILAR: "Similar", STATUS_BETTER: "Better",
                                             STATUS_WORSE: "Worse", STATUS_NOT_EXECUTE: "NA"}



    TEST_LEVEL_SYSTEM = "System Test"
    TEST_LEVEL_GRANULAR = "Granular Test"
    TEST_LEVEL_USER = "User Test"
    TEST_LEVEL_HEADER_LIST = [TEST_LEVEL_SYSTEM, TEST_LEVEL_GRANULAR, TEST_LEVEL_USER]

    TOTAL_TEST_CASES = "Total Test Cases"

    TEST_LEVEL = "Test Case ID"
    PASS = "PASS"
    FAIL = "FAIL"
    OPEN = "OPEN"
    TOTAL = "Total"
    MODULE_NAME = "MODULE NAME"
    TEST_TYPE = "TEST TYPE"
    EXEC_TIME_SEC = "EXEC TIME SEC"
    MANUAL_TIME_SEC = "MANUAL TIME SEC"
    NUMBER_OF_TESTS = "NUMBER OF TESTS"

    TEST_LEVEL_INDEX_LIST = [TEST_LEVEL, PASS, FAIL, OPEN, TOTAL]

    MODULE_DETAILS_TAB_HEADER = [MODULE_NAME, TEST_TYPE, EXEC_TIME_SEC, MANUAL_TIME_SEC, NUMBER_OF_TESTS,
                                 CURRENT_RELEASE_HEADER]
    SUMMARY_TAB_HEADER = [MODULE_NAME,CURRENT_RELEASE_HEADER,PREVIOUS_RELEASE_HEADER]
    NUMBER_OF_TESTS = "NUMBER OF TESTCASE"
    SUMMARY_TAB_SUB_HEADER = [NUMBER_OF_TESTS ,PASS,FAIL,OPEN]

    MODULE_DETAILS_TAB_NAME = "Module_Details"
    MODULE_DETAILS_MD_TAB_NAME = "Module_Details_MD"
    SUMMARY_BUNDLE_MD_TAB_NAME = "Summary_Bundle_MD"
    SUMMARY_CONFIG_MD_TAB_NAME = "Summary_Config_MD"
    SUMMARY_TAB_NAME = "Summary"

    LOGGING_FORMAT = '%(levelname)s:  %(filename)s:%(lineno)s:    %(message)s'
    SHEET_NAME_MAXSIZE = 31
    CONFIGURATIONS_COUNT = len(CONFIGURATIONS_LIST)
    AVERAGE_KEY = "AVERAGE"
    COUNT_KEY = "COUNT"
    SUM_KEY = "SUM"
    MEASUREMENT = "Run"
    MEASUREMENT_STRING_FORMAT = "Measurement %d"
    BUNDLES_AVERAGE_RESPONSE_TIME = "Bundles Average Response Time"
    TEST_CASE_DESCRIPTION_COUNT = len(TEST_CASE_DESCRIPTION_LIST)

    APPLICATION_NAME_HEADER = "Application Name"
    CONFIGURATIONS_HEADER = "Configuration"
    IMPROVEMENT_HEADER = "Bundle to Bundle Comparison Performance Result for ({})"

    BUNDLE_TO_BUNDLE_COMPARISON_HEADER = "Bundle to Bundle Comparison Performance Result for ({})"
    CONFIG_TO_CONFIG_COMPARISON_HEADER = "Config to Config Comparison Performance Result for ({})"

    COLOR_GREY = "#C0C0C0"
    COLOR_LIGHT_ORANGE = "#F8CBAD"
    COLOR_LIGHT_BLUE = "#BDD7EE"
    COLOR_NAVY_BLUE = "#0066CC"
    COLOR_LIGHT_GREEN = "#C4D79B"
    COLOR_GOLD = "#F0E68C"


    __TESTCASE_COUNT = 120
    __MAX_RUN_COUNT = 0

    # # Set Counting total rows
    def setRowCount(self, row_count):
        self.__TESTCASE_COUNT = row_count

    # # Get Counting total rows
    def getRowCount(self):
        return self.__TESTCASE_COUNT

    # # Set Maximum Run Count
    def setMaxRunCount(self, max_run_count):
        self.__MAX_RUN_COUNT = max_run_count

    # # Get Maximum Run Count
    def getMaxRunCount(self):
        return self.__MAX_RUN_COUNT
