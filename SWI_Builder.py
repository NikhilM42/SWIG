import xlsxwriter
import datetime

workbook = xlsxwriter.Workbook('./output_swi/SWI_' + datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + '.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_paper(3)

filenom = input('Enter file name: ')

if (filenom == ''):
    filenom = 'SWI'

f = open('./instructions_src/'+filenom + '.txt','r')

worksheet.set_row(0, 30)
for x in range(0,18):
    worksheet.set_column(x,x, 10)

bf_1 = workbook.add_format({'bold': True, 'font_size': 18, 'font_name': 'Arial Black', 'align': 'left', 'valign': 'vcenter', 'top': 5, 'left': 5})
bf_2 = workbook.add_format({'bold': True, 'font_size': 16, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter', 'top': 5})
bf_3 = workbook.add_format({'bold': True, 'font_size': 10, 'font_name': 'Arial', 'bottom': 6, 'right': 6, 'left': 5})
bf_4 = workbook.add_format({'font_size': 10, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter', 'right': 5, 'bottom': 6})
bf_5 = workbook.add_format({'font_size': 16, 'font_name': 'Arial', 'font_color': 'blue', 'border': 5, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
bf_6 = workbook.add_format({'font_size': 16, 'font_name': 'Calibri', 'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'border': 5})
bf_7 = workbook.add_format({'bold': True, 'font_size': 12, 'font_name': 'Arial','font_color': 'red', 'border': 5, 'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
bf_8 = workbook.add_format({'bold': True, 'font_size': 16, 'font_name': 'Arial', 'border': 5, 'align': 'center', 'valign': 'vcenter'})
bf_9 = workbook.add_format({'font_size': 10, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter', 'num_format': 'd-mmm-yy', 'right': 5, 'bottom': 6})
bf_10 = workbook.add_format({'left': 5, 'right': 5, 'bottom': 5})
bf_11 = workbook.add_format({'top': 5, 'right': 5, 'bottom': 5})
bf_12 = workbook.add_format({'bold': True, 'font_size': 10, 'font_name': 'Arial', 'bottom': 5, 'right': 6, 'left': 5})
bf_13 = workbook.add_format({'font_size': 10, 'font_name': 'Arial', 'align': 'center', 'valign': 'vcenter', 'right': 5, 'bottom': 5})
bf_14 = workbook.add_format({'left': 5, 'bottom': 5})
bf_15 = workbook.add_format({'left': 5})

worksheet.merge_range(0,0,0,2, 'Company Name', bf_1)
worksheet.merge_range(0,3,0,14, 'STANDARD WORK INSTRUCTION', bf_2)
worksheet.merge_range(0,15,0,17, None, bf_11)
worksheet.write_blank(1,0,None, bf_15)
worksheet.merge_range(1,15,1,16, 'DATE ORIGINATED:', bf_3)
worksheet.write_datetime(1,17, datetime.datetime.strptime(datetime.datetime.now().strftime("%Y-%m-%d"),"%Y-%m-%d"), bf_9)
worksheet.merge_range(2,0,2,2, None, bf_14)
worksheet.merge_range(2,15,2,16, 'REVISED BY', bf_3)
worksheet.write(2,17, '', bf_4)
worksheet.merge_range(3,0,3,1,'PART NAME', bf_3)
worksheet.write(3,2, '', bf_4)
worksheet.merge_range(3,15,3,16, 'DATE REVISED:', bf_3)
worksheet.write(3,17, datetime.datetime.strptime(datetime.datetime.now().strftime("%Y-%m-%d"),"%Y-%m-%d"), bf_9)
worksheet.merge_range(4,0,4,1,'OPERATION', bf_12)
worksheet.write(4,2, '', bf_13)
worksheet.write(4,15, 'FORM #:', bf_3)
worksheet.merge_range(4,16,4,17, '2-01-072-01', bf_4)
worksheet.write_blank(5,0,None, bf_15)
worksheet.write(5,15, 'DOC #:', bf_3)
worksheet.merge_range(5,16,5,17, 'N/A', bf_4)
worksheet.write_blank(6,0,None, bf_15)
worksheet.write(6,15, 'ECR #:', bf_12)
worksheet.merge_range(6,16,6,17, '2-xx-xxx', bf_13)
worksheet.write_blank(7,0,None, bf_15)
worksheet.merge_range(7,15,7,17, None, bf_11)

worksheet.fit_to_pages(1, 0)

row = 8
col = 0
bc = 1
bl = []
step = 1
pos = 0
width = 5
h = 7

for line in f:

    if line[0] == 'r':
        width = 5
    elif line[0] == 'm':
        width = 11
    elif line[0] == 'h':
        width = 8
    elif line[0] == 'w':
        width = 17

    if line[0] == 't':
        worksheet.set_row(row, 30)
        worksheet.merge_range(row,0,row,17, line[1:], bf_8)
        row = row + 1
        step = 1
        col = 0
    else:
        for x in range(row, row+3): 
            worksheet.set_row(x, 30)
        for x in range(row+3, row+h-1): 
            worksheet.set_row(x, 60)
        for x in range(row+h-1, row+h+1): 
            worksheet.set_row(x, 30)
        worksheet.merge_range(row,col,row+2,col,step, bf_5)
        worksheet.set_column(col,col, 4)
        pos = line.find('NOTE')
        if pos>=0:
            worksheet.merge_range(row,col+1,row+2,col+width, line[1:(pos-1)], bf_6)
            worksheet.merge_range(row+3,col,row+h-2,col+width, '', bf_10)
            worksheet.merge_range(row+h-1,col,row+h,col+width, line[pos:(len(line)-1)], bf_7)
        else:
            worksheet.merge_range(row,col+1,row+2,col+width, line[1:(len(line)-1)], bf_6)
            worksheet.merge_range(row+3,col,row+h,col+width, '', bf_10)
        
        step = step + 1
        col = col + width + 1
        if col > 17:
            row = row + h + 1
            bc = bc + 1
            if bc > 4:
                bc = 1
                bl.append(row)
            col = 0

worksheet.set_h_pagebreaks(bl)
workbook.close()
