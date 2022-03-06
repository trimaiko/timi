import logging
from tkinter.tix import COLUMN
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
            mycell.value = dat.get('出社時刻')
            self.logger.info(dat.get('出社時刻'))

            mycell = ws.cell(row = base_row, column = depart_base_col)
            mycell.value = dat.get('退社時刻')
            self.logger.info(dat.get('退社時刻'))

            base_row += 1

        # write_to_outputfile
        wb.save(filename = 'saved.xlsx')
