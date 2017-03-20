from docx import Document
import xlwings as xw

class PTS_Processor(object):
    def __init__(self, doc_path):
        f = open(doc_path, 'rb')
        self.doc= Document(f)
        self.cmd_list_path = doc_path.split('.')[0]+'.xlsx'

    def read_cmd(self):
        for n in range(len(self.doc.tables)):
            row_title = [r.text for r in doc.tables[n].row_cells(0)]

            if row_title[0].find("ID") is not -1:
                row_cmd = [r.text for r in doc.tables[n].row_cells(1)]
                if row_cmd[0] == "":
                        row_cmd = [r.text for r in doc.tables[n].row_cells(2)]
                print(row_title)
                print(row_cmd)

    def create_cmd_list(self):
        print(self.cmd_list_path)
        cmd_wb = xw.Workbook(self.cmd_list_path, app_visible=True)
        cmd_sheet = xw.Sheet('cmd_list')
        cmd_wb.save()
        cmd_wb.close()


    def write_cmd():
        pass

if __name__ == "__main__":
    file = "ESTONA_EL-ET-CAR_ID5027-88_V01_20161130-Release.docx"
    pts_or = PTS_Processor(file)
    #pts_or.read_cmd()
    pts_or.create_cmd_list()

