import logging
from datetime import date
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
            mycell = ws.cell(row = base_row, column= arrive_base_col)
            atime = datetime.strptime(dat.get('出社時刻'), '%H:%M')
            mycell.value = self.convert_strtime_to_excel_ordinal(dtime)
            self.logger.info(dat.get('出社時刻'))

            mycell = ws.cell(row = base_row, column = depart_base_col)
            dtime = datetime.strptime(dat.get('退社時刻'), '%H:%M')
            mycell.value = self.convert_date_to_excel_ordinal(atime)
            self.logger.info(dat.get('退社時刻'))

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
