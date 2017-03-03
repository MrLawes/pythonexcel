* *>>>* from pythonexcel import Excel
* *>>>* import os
* *>>>* excel = Excel(title=['t1', 't2'])
* *>>>* raws = {'t1': 't1_1', 't2': 't2_1'}
* *>>>* excel.add_row(raws)
* *>>>* excel.append_title(title=['t3', 't4'])
* *>>>* excel.add_row({'t1': 't1_2','t2': 't2_2','t3': 't3_2','t4': 't4_2',})
* *>>>* excel.init_title(title=['new_t1', 'new_t2','new_t3','new_t4'])
* *>>>* excel.add_row({'new_t1': 'new_t1_3'})
* *>>>* excel.to_sheet('sheet2')
* *>>>* excel.init_title(title=['s2_t1', 's2_t2',])
* *>>>* excel.add_row({'s2_t1': 's2_t1_c1', 's2_t2': 's2_t2_c1'})
* *>>>* excel.to_sheet('sheet1')
* *>>>* excel.add_row({'new_t1': 'new_t1_4'})
* *>>>* excel.workbook.close()
* *>>>* print excel.xlsx_file.name
* *>>>* os.system("mv %s tmp.xls" % (excel.xlsx_file.name))