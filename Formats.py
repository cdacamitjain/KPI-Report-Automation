"""
# # DATE		: 06,October 2017
# # AUTHOR		: AMIT.JAIN@LNTTECHSERVICES.COM
# # DESCRIPTION	: This script is used to provide Cell Formatting.
# #
"""
class Formats:
    @staticmethod
    def Format_First_MD_Header(workbook,color):
        format = workbook.add_format()
        format.set_text_wrap()
        format.set_center_across()
        format.set_bold()
        format.set_align("center")
        format.set_align("vcenter")
        format.set_border(1)
        format.set_bg_color(color)
        return format

    @staticmethod
    def Format_Cell(workbook):
        format = workbook.add_format()
        format.set_text_wrap()
        format.set_center_across()
        format.set_align("center")
        format.set_align("vcenter")
        format.set_border(1)
        return format

    @staticmethod
    def Format_ColourStatus(workbook):
        format = workbook.add_format()
        format.set_text_wrap()
        format.set_bg_color("red")
        format.set_center_across()
        format.set_align("center")
        format.set_align("vcenter")
        format.set_border(1)
        return format

    @staticmethod
    def Format_Hyperlink(workbook):
        format = workbook.add_format()
        format.set_center_across()
        format.set_text_wrap()
        format.set_align("center")
        format.set_align("vcenter")
        format.set_underline()
        format.set_font_color("#0000FF")
        format.set_border(1)
        return format

    @staticmethod
    def Format_Percentage(workbook):
        format = workbook.add_format({'num_format': '[Green]#,##0.00%;[Red]#,##0.00%'})
        format.set_text_wrap()
        format.set_center_across()
        format.set_align("center")
        format.set_align("vcenter")
        format.set_border(1)
        return format
