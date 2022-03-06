import logging
from openpyxl import load_workbook



class timemanager:

    def __init__(self, data, outputfile):
        # data is python object
        self.data = data
        self.outputfile = outputfile

        self.logger = logging.getLogger(__name__)

    def edit_timecard(self):
        self.logger.info("edit_timecard start")

        rs = '5'
        cs = '9'
        
        wb = load_workbook(filename = self.outputfile)
        #always select first sheet in workbook
        ws = wb.worksheets[0]
        
        row = ws['C']
        self.logger.info(self.data)
        self.logger.info(row)
        
        for dat in self.data:
            row[rs] = dat['出社時刻']
            self.logger.info(row[rs].values())

        # write_to_outputfile
        wb.save(filename = self.outputfile)
