from RPA.Excel.Files import Files
from RPA.Tables import Tables
import os
import logging

el = Files()
tb = Tables()

class ExcelActivities:

    def __init__(self) -> None:
        self.logger = logging.getLogger(__name__)

    def write_to_wb(self, wsname, wbname, data):
        try:
            el.open_workbook(wbname)
            el.create_worksheet(wsname)
            el.set_active_worksheet(wsname)
            el.append_rows_to_worksheet(data,wsname)
            el.save_workbook(wbname)
        except Exception as err:
            self.logger.error("An error ocurrs writing the data on excel: " + str(err))
            raise SystemError("An error ocurrs writing the data on excel: " + str(err))

    def create_file(self, name, destine):
        try:
            el.create_workbook(name)
            el.save_workbook(destine+"/"+name)
        except Exception as err:
            self.logger.error("An error ocurrs creating the excel file: " + str(err))
            raise SystemError("An error ocurrs creating the excel file:" + str(err))

    def create_table(self, agencies, amounts):
        try:
            data = tb.create_table()
            tb.add_table_column(data,"A")
            for n in agencies:
                tb.add_table_row(data)
            tb.set_table_column(data,"A", agencies)
            tb.add_table_column(data,"B")
            tb.set_table_column(data,"B", amounts)
            return data
        except Exception as err:
            self.logger.error("Unable to create table because: " + str(err))
            raise SystemError("Unable to create table because: " + str(err))


    def compare_value(self, directory, worksheet, column, value):
        try:
            el.open_workbook(directory)
            dt = el.read_worksheet_as_table(worksheet)
            tb.rename_table_columns(dt, ["UII","Bureau", "Investment Title", "Total FY2021 Spending ($M): activate to sort column ascending", "Type", "CIO Rating", "# of projects"])
            tblist = tb.group_table_by_column(dt, column)
            for li in tblist:
                if(li.__eq__(value)):
                    return True
            return False
        except Exception as err:
            self.logger.error("Compare value fails: " + str(err))
            raise SystemError("Compare value fails: " + str(err))





