* *>>>* from pythonexcel import Excel
* *>>>* import os
* *>>>* excel = Excel(title=['t1', 't2'])
* *>>>* raws = {'t1': 'c1', 't2': 'c2'}
* *>>>* excel.add_row(raws)
* *>>>* excel.workbook.close()
* *>>>* print excel.xlsx_file.name
* *>>>* os.system("mv %s tmp.xls" % (excel.xlsx_file.name))