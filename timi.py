import logging
from datetime import date, datetime
from openpyxl import load_workbook


class timemanager:

    def __init__(self, data, outputfile):
        # data is python object
        self.data = data
        self.outputfile = outputfile

        self.logger = logging.getLogger(__name__)

    def edit_timecard(self):
        self.logger.info("edit_timecard start")

        base_row = 8
        arrive_base_col = 6
        depart_base_col = 7

        wb = load_workbook(filename = self.outputfile)
        #always select first sheet in workbook
        ws = wb.worksheets[0]

        # self.logger.info(self.data)

        # insert arrival time
        for dat in self.data['data']:

            mycell_arrive = ws.cell(row = base_row, column= arrive_base_col)
            mycell_dept = ws.cell(row = base_row, column = depart_base_col)

            if dat.get('出社時刻') != "":
                print(dat.get('出社時刻'))
                atime = datetime.strptime(dat.get('出社時刻'), '%H:%M')
                # mycell_arrive.value = self.convert_date_to_excel_ordinal(atime)
                mycell_arrive.value = atime
                mycell_arrive.number_format = "hh:mm:ss"
                print('mycell_arrive.value:', mycell_arrive.value)
                print('atime:', atime)
                self.logger.info(dat.get('出社時刻'))

                dtime = datetime.strptime(dat.get('退社時刻'), '%H:%M')
                mycell_dept.value = self.convert_date_to_excel_ordinal(dtime)
                self.logger.info(dat.get('退社時刻'))

            else:
                mycell_arrive.value = ""
                mycell_dept.value = ""

            base_row += 1

        # write_to_outputfile
        wb.save(filename = 'saved.xlsx')

    def convert_date_to_excel_ordinal(self, targt_time):

        # Specifying offset value i.e.,
        # the date value for the date of 1900-01-00
        offset = 693594

        # Calling the toordinal() function to get
        # the excel serial date number in the form
        # of date values
        n = targt_time.toordinal()
        return (n - offset)
