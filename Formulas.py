"""
# # DATE		: 06,October 2017
# # AUTHOR		: AMIT.JAIN@LNTTECHSERVICES.COM
# # DESCRIPTION	: This script provides Excel Formulas methods.
# #
"""
from Constants import Constants
from Utils import Utils


class Formulas:

    @staticmethod
    def Formula_PerformanceComparedToAllReleases(row_index, col_index, bundle_count, config_count, MD=True):
        # bundle_count = 4
        # row_index = 3
        # col_index = 14
        # config_count = 4

        Constant_range_cell = Utils.get_cell_range(row_index, col_index, bundle_count, MD)
        formula_list = []

        row = row_index
        for j in range(config_count):
            row_index = row + j
            col_name = Utils.GetColumnName(col_index)
            current_cell = col_name + str(row_index)  # O3
            range_cell = Utils.get_cell_range(row_index, col_index, bundle_count)
            # print "current_cell",current_cell
            # print "range_cell",range_cell
            str_ = "="
            extra_bracket = 0
            if MD:
                str_ += 'IF(COUNT({0})=0,"NEW",'.format(Constant_range_cell)
                extra_bracket = 1
            next_cell = Utils.GetColumnName(col_index + 1) + str(row_index)
            str_ += 'IF({0}="NA","NA",IF({1}="NA","NA",'.format(current_cell,next_cell)
            # print "row_index",row_index

            if bundle_count <= 2:
                temp_str = Formulas.condition(row_index, col_index, bundle_count)
                for i in range(bundle_count + extra_bracket):
                    temp_str += ")"
                str_ += temp_str
            else:
                next_cell = ""
                for i in range(2, bundle_count):
                    next_cell = Utils.GetColumnName(col_index + i) + str(row_index)
                    if i > 2:
                        str_ += ","
                    str_ += 'IF({}="NA",'.format(next_cell)
                    str_ += Formulas.condition(row_index, col_index, i)

                    temp_str = "," + Formulas.condition(row_index, col_index, bundle_count)
                # print temp_str
                for i in range(bundle_count + extra_bracket):
                    temp_str += ")"
                str_ += temp_str
            formula_list.append(str_)

        return formula_list

    @staticmethod
    def Formula_TestCasesComparedForAverageResults(row_index,col_index,bundle_count,config_count,tc_count):
        string = '=SUM('
        for i in range(tc_count):
            cell_range = Utils.get_cell_range(row_index,col_index-1,bundle_count+1)
            string+= 'IF(COUNT({1})={0},1,0),'.format(bundle_count,cell_range)
            row_index += config_count

        return string[:-1]+")"


    @staticmethod
    def condition(row_index, col_index, bundle_count):
        range_cell = Utils.get_cell_range(row_index, col_index, bundle_count)
        current_cell = Utils.GetColumnName(col_index) + str(row_index)
        sstr = 'IF((MIN({1})-0.1*(MIN({1})))<={0},(IF({0}<=(MIN({1})+0.1*(MIN({1}))),"Similar",(IF({0}>(MIN({1})),"Degraded","Improved")))),"Improved")'
        condition_ = sstr.format(current_cell, range_cell)
        return condition_

    @staticmethod
    def Formula_PerformanceComparedBetweenAllConfiguration(colName, row_no, configCount):
        # col_id 		= "O"
        # row_no 		= 3
        # configCount 	= 4

        formula_string_list = []
        col_row_list = []

        for i in range(configCount):
            col_row_list.append(colName + str(row_no + i))

        val1 = ""
        val2 = col_row_list[0]
        val3 = col_row_list[1]
        val4 = col_row_list[2]
        val5 = col_row_list[3]
        val6 = ""

        for val in col_row_list:
            templist = col_row_list
            var1 = col_row_list[configCount - configCount] + ":" + col_row_list[configCount - 1]
            var2 = ""
            var3 = val

            for t in templist:
                if t == var3:
                    continue
                var2 = var2 + t + ","
            var2 = var2[:-1]

            val1 = var1
            val2 = val
            val6 = var2

            ss = '=IF(SUM({0})=0,"NA",IF({1}="NA","NA",IF(AND({2}="NA")*AND({3}="NA")*AND({4}="NA"),"NA",IF((MIN({5})-0.1*(MIN({5})))<={1},(IF({1}<=(MIN({5})+0.1*(MIN({5}))),"Similar",(IF({1}>MIN({5}),"Worse","Better")))),"Better"))))'
            ss = ss.format(val1,val2,val3,val4,val5,val6)

            temp = val2
            val3 = val4
            val4 = val5
            val5 = temp
            formula_string_list.append(ss)

        return formula_string_list

    @staticmethod
    def Formula_TestCaseResult(row_id, Sheet_Name):
        formula_list = []
        formula_list.append('=IF('+	Sheet_Name	+'!F'+str(row_id+0)+'="NEW","OPEN",IF('+Sheet_Name+'!F'+str(row_id+0)+'="NA","OPEN",IF('+Sheet_Name+'!F'+str(row_id+0)+'="Improved","PASS",IF('+Sheet_Name+'!F'+str(row_id+0)+'="Similar","PASS","FAIL"))))')
        formula_list.append('=IF('+	Sheet_Name	+'!F'+str(row_id+1)+'="NEW","OPEN",IF('+Sheet_Name+'!F'+str(row_id+1)+'="NA","OPEN",IF('+Sheet_Name+'!F'+str(row_id+1)+'="Improved","PASS",IF('+Sheet_Name+'!F'+str(row_id+1)+'="Similar","PASS","FAIL"))))')
        formula_list.append('=IF('+	Sheet_Name	+'!F'+str(row_id+2)+'="NEW","OPEN",IF('+Sheet_Name+'!F'+str(row_id+2)+'="NA","OPEN",IF('+Sheet_Name+'!F'+str(row_id+2)+'="Improved","PASS",IF('+Sheet_Name+'!F'+str(row_id+2)+'="Similar","PASS","FAIL"))))')
        formula_list.append('=IF('+	Sheet_Name	+'!F'+str(row_id+3)+'="NEW","OPEN",IF('+Sheet_Name+'!F'+str(row_id+3)+'="NA","OPEN",IF('+Sheet_Name+'!F'+str(row_id+3)+'="Improved","PASS",IF('+Sheet_Name+'!F'+str(row_id+3)+'="Similar","PASS","FAIL"))))')
        formula_list.append('=IF(COUNTIF('+Sheet_Name+'!G'+str(row_id+0)+':G'+str(row_id+3)+',"Worse"),"FAIL",IF(COUNTIF('+Sheet_Name+'!G'+str(row_id+0)+':G'+str(row_id+3)+',"NA")=4,"OPEN","PASS"))')
        formula_list.append('=IF(COUNTIF('+Sheet_Name+'!I'+str(row_id+0)+':I'+str(row_id+3)+',"NA")=4,"OPEN",IF(COUNTIF('+Sheet_Name+'!I'+str(row_id+0)+':I'+str(row_id+3)+',"Yes"),"FAIL","PASS"))')
        formula_list.append('=IF(COUNTIF('+Sheet_Name+'!H'+str(row_id+0)+':H'+str(row_id+3)+',"NA")=4,"OPEN",IF(COUNTIF('+Sheet_Name+'!H'+str(row_id+0)+':H'+str(row_id+3)+',"Yes"),"FAIL","PASS"))')
        return formula_list

    @staticmethod
    def Formula_PerformanceWiseCount(start_row_no,col_name, testcase_count, config_count, status):
        var = ""
        for i in range(testcase_count):
            var = var + 'COUNTIF(' + (col_name + str(start_row_no)) + ',"' + status + '"),'
            # # Remove Extra ','
            if i == (testcase_count - 1):
                var = var[:-1]
            start_row_no = start_row_no + config_count
        return "=SUM(" + var + ")"


    # col_id 			= "O"
    # start_row_no		= 3
    # testcase_count 	= 7;
    # config_count		= 4
    @staticmethod
    def Formula_BundleWiseAverage(col_id, start_row_no, testcase_count, config_count,key = "AVERAGE"):
        var = ""
        for i in range(testcase_count):
            var = var + (col_id + str(start_row_no)) + ","
            # # Remove Extra ','
            if i == (testcase_count - 1):
                var = var[:-1]
            start_row_no = start_row_no + config_count
        str_ = '=IFERROR({0}({1}),"NA")'.format(key,var)
        return str_

    @staticmethod
    def Formula_Average(row_index,col_index,current_col_index,bundle_count,config_count,test_count):
        NA_Cell = "A"+str(Constants.NA_ROW_INDEX_REFERENCE)
        string= '=IFERROR(AVERAGE('
        for i in range(test_count):
            cell_range = Utils.get_cell_range(row_index,col_index-1,bundle_count+1)
            current_cell = Utils.GetColumnName(current_col_index) + str(row_index)
            string += 'IF(COUNT({0})={1},{2},{3}),'.format(cell_range,bundle_count,current_cell,NA_Cell)
            row_index += config_count
        return string[:-1]+'),"NA")'
    @staticmethod
    def Formula_PerformanceComparedToLastRelease(row_index,col_index):
        current_cell = Utils.GetColumnName(col_index) + str(row_index)
        next_cell = Utils.GetColumnName(col_index+1) + str(row_index)
        return '=IF(COUNTIF({0}:{1},"NA"),"NA",IF({0}-0.1*{0}<{1},"Improved",IF({0}+0.1*{0}>{1},"Degraded","Similar")))'.format(current_cell,next_cell)

    @staticmethod
    def Formula_ImprovementPercentageComparedToCurrentRelease(row_index,col_index):
        current_cell = Utils.GetColumnName(col_index) + str(row_index)
        next_cell = Utils.GetColumnName(col_index+1) + str(row_index)
        return '=IFERROR(IF({0}="NA","NA",IF({1}="NA","NA",(({1}-{0})/{1}))),"NA")'.format(current_cell,next_cell)
