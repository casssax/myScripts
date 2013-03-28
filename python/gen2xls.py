# THIS PARM IS SET UP TO WORK WITH ALL LSC GENCOUNTS
""" Things to do:
    1. 
    """
import win32com.client as win32
win32c = win32.constants
import sys
import mmap
import itertools
import re
import traceback
from ctypes import *
##########################################################################################
##########################################################################################
#----------------------------------------------------------------
# LAURA'S FORMAT: get header row and split into columns and return start pos. of each
def get_headers(line):
    tmp_header = []
    tmp_list = line.split('          ')
    for i in tmp_list:
        tmp_header.append(i.strip())
    header_lengths = []
    for e in tmp_header:
        header_lengths.append(line.find(e))
    return tmp_header,header_lengths
#----------------------------------------------------------------
# LAURA'S FORMAT: get data rows and split into columns
def get_data(line,hd_length):
    tmp_data = []
    tmp_list = []
    if len(hd_length) == 2:
        tmp_list.append(line[hd_length[0] - 1:hd_length[1]])
        tmp_list.append(line[hd_length[1]:])
    field_count = 0
    if len(hd_length) > 2:
        while field_count < len(hd_length) - 1:
            tmp_list.append(line[hd_length[field_count] - 1:hd_length[field_count + 1] - 1])
            field_count += 1
        tmp_list.append(line[-13:])
    for i in tmp_list:  # strip white space
        tmp_data.append(i.strip())
    return tmp_data
#-----------------------------------------------------------------
def get_sheet_name(headers,sheet_names):
    if len(headers) == 2:
        test_name = headers[0].translate(None, ':\/?*[]')  #strip chars not valid in sheet names
        test_name = test_name[:28]  #trim to max legal length -2 (for duplicate sheet names)
    else:
        tmphdr = headers[:-1]
        test_name = '-'.join(tmphdr)
        test_name.translate(None, ':\/?*[]')  #strip chars not valid in sheet names
        test_name = test_name[:28]  #trim to max legal length -2 (for duplicate sheet names)
    if test_name not in sheet_names:  #test to see if name already in use
        sheet_names.append(test_name)
        return [test_name, sheet_names]
    else:  #if name already in use, append with '2' or higher
        name_suffix = 2
        flag = True
        while flag:
            test_name = test_name + str(name_suffix)
            if test_name not in sheet_names:
                sheet_names.append(test_name)
                return [test_name, sheet_names]
                flag = False
            else:
                name_suffix += 1
#---------------------------------------------------------------
# SERGEI'S FORMAT: get headers
def get_s_headers(line):
    fields = []
    pos = 0
    while pos != -1:
        pos = line.find('\t')
        if pos != -1:
            fields.append(line[0:pos].strip())
            line = line[pos + 1:]
        else:
            fields.append(line.strip())
    blank = []
    return fields,blank
#---------------------------------------------------------------
# SERGEI'S FORMAT: get data
def get_s_data(line):
    fields = []
    pos = 0
    while pos != -1:
        pos = line.find('\t')
        if pos != -1:
            fields.append(line[0:pos].strip())
            line = line[pos + 1:]
        else:
            fields.append(line.strip())
    return fields
#---------------------------------------------------------------
# LAURA'S FORMAT:
def get_l_sheets(file):
    sheets = []
    page_data = []
    headers = []
    count_on_page = 1
    for line in file:
        if line.find('\x0C') == 0:   #if newpage found, create a new sheet
            sheets.append([report_date,report_title,headers,page_data,num_entries,total_quantity])
            page_data = []
            headers = []
            count_on_page = 1
        elif count_on_page == 1:
            report_date = line.rstrip()
        elif count_on_page == 2:
            report_title = line
        elif count_on_page == 4:
            headers = get_headers(line)
        elif line[:14] == 'Total Quantity':
            total_quantity = line[14:].rstrip()
        elif line[:12] == '# of Entries':
            num_entries = line[12:].rstrip()
        else:
            if len(line) != 1 and line[:13] != 'Report Totals':
                p_data = get_data(line,headers[1])
                if p_data != headers[0]:
                    page_data.append(p_data)
        count_on_page += 1
    sheets.append([report_date,report_title,headers,page_data,num_entries,total_quantity])   
    return sheets
#-----------------------------------------------------------------------
# SERGEI'S FORMAT:
def get_s_sheets(file,report_date,report_title):
    sheets = []
    first_pass = True
    page_data = []
    headers = []
    count_on_page = 1
    for i in range(3): file.next()
    for line in file:
        if line.find('\x0C') == 0 and not first_pass:   #if newpage found, create a new sheet
            sheets.append([report_date,report_title,headers,page_data,num_entries,total_quantity])
            page_data = []
            headers = []
            count_on_page = 1
        else:  
            if count_on_page == 3:
                headers = get_s_headers(line)
            elif line[:14] == 'Total Quantity':
                total_quantity = line[15:].rstrip()
            elif line[:12] == '# of Entries':
                num_entries = line[13:].rstrip()
            else:
                if len(line) != 1 and line[:13] != 'Report Totals' and count_on_page != 4 and line[:16] != "PARAMETER 'UNIT'":
                    p_data = page_data.append(get_s_data(line))
        count_on_page += 1
        first_pass = False
    sheets.append([report_date,report_title,headers,page_data,num_entries,total_quantity]) 
    return sheets
#----------------------------------------------------------------------------
# test to see whether Sergei's or Laura's format gencount
def get_gen_type(file):
    s = mmap.mmap(file.fileno(), 0, access=mmap.ACCESS_READ)   
    check = s.find('\x0C')  #look for first newpage
    if check == -1: return check
    if check < 100:  #this is an arbitrary number that works for now. 
        file_type = 'serg'
        t = s.read(check - 1) 
        split_fields = t.split('\r\n',2)  
        report_date = split_fields[0].split(' ')[0]  
        report_title = split_fields[1]
        s.close()
        return [file_type,report_date,report_title]
    else:
         file_type = 'laur'  
         s.close()
         return [file_type]
#------------------------------------------------------------------------------
# populate excel workbook
def make_sheets(sheets,fname):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Add()
    sheet_names = []
    # excel constants
    xlLeft = -4131
    xlBottom = -4107
    xlCenter = -4108
    xlTop = -4160
    sheet_count = len(sheets) + 1
    for sheet in reversed(sheets):
        sheet_count -= 1
        ws = wb.Worksheets.Add()
        get_sheet_return = get_sheet_name(sheet[2][0],sheet_names)
        ws.Name = get_sheet_return[0]
        sheet_names = get_sheet_return[1]
        ws.Cells(6,2).Value = sheet[0]  #report_date
        ws.Cells(6,2).HorizontalAlignment = xlLeft
        header_count = 1
        for h in sheet[2][0]:   #headers 
            ws.Cells(9,header_count + 1).Value = h
            ws.Cells(9,header_count + 1).Font.Bold = True
            header_count += 1
        page_count = 10
        for p in sheet[3]:  #page_data
            data_count = 1
            for d in p:
                ws.Cells(page_count,data_count +1).Value = d
                if sheet[2][0][data_count - 1] != 'QUANTITY':
                    ws.Cells(page_count,data_count + 1).HorizontalAlignment = xlLeft
                else:
                    ws.Cells(page_count,data_count + 1).NumberFormat = "###,###,##0"  #number format on quantity column
                data_count += 1
            page_count += 1
        ws.Cells(page_count + 2,2).Value = 'Report Totals'  #Report Totals
        ws.Cells(page_count + 2,2).Font.Bold = True
        ws.Cells(page_count + 4,2).Value = 'Total Quantity'  #Report Totals
        ws.Cells(page_count + 4,2).Font.Bold = True
        ws.Cells(page_count + 4,len(sheet[2][0]) + 1).Value = sheet[5]  #total_quantity
        ws.Cells(page_count + 4,len(sheet[2][0]) + 1).Font.Bold = True
        ws.Cells(page_count + 6,2).Value = '# of Entries'  # of Entries
        ws.Cells(page_count + 6,2).Font.Bold = True
        ws.Cells(page_count + 6,len(sheet[2][0]) + 1).Value = sheet[4]  # of Entries
        ws.Cells(page_count + 6,len(sheet[2][0]) + 1).Font.Bold = True
        ws.Columns.AutoFit()
        ws.Cells(7,2).Value = sheet[1]  #report_title
        ws.Cells(7,2).WrapText = False
        ws.Cells(7,2).Font.Size = 14
        ws.Cells(7,2).Font.Bold = True
        #ws.Shapes.AddPicture("C:\\DATA_SAVE\\excel\\LSC_LOGO.bmp",0,1,0,0,50,50)
        ws.Shapes.AddPicture("C:\\DATA_SAVE\\excel\\logo3.bmp",0,1,0,0,50,50)
        ws.Cells(1,2).Value = '  List Services Corporation'
        ws.Cells(1,2).WrapText = False
        ws.Cells(1,2).Font.Size = 12
        ws.Cells(1,2).Font.Bold = True
        ws.Cells(1,2).VerticalAlignment = xlBottom
        ws.Cells(2,2).Value = '   6 Trowbridge Drive'
        ws.Cells(2,2).WrapText = False
        ws.Cells(2,2).Font.Size = 10
        ws.Cells(2,2).VerticalAlignment = xlCenter
        ws.Cells(3,2).Value = '   Bethel, CT 06801'
        ws.Cells(3,2).WrapText = False
        ws.Cells(3,2).Font.Size = 10
        ws.Cells(3,2).VerticalAlignment = xlTop
        #txb = ws.Shapes.AddTextbox(1, 50, 0, 410, 50)
        #txb.TextFrame2.TextRange.Characters.Text = "List Services Corporation\n6 Trowbridge Drive\nBethel, CT 06810"
        #txb.TextFrame2.TextRange.Characters.Font.Bold = True
        #txb.Line.Visible = False
        ws.Rows("4:4").Borders(9).LineStyle = 1
    wb.Worksheets('Sheet1').Delete()  #delete original three sheets from workbook
    wb.Worksheets('Sheet2').Delete()
    wb.Worksheets('Sheet3').Delete()
    wb.SaveAs(fname + '.xlsx')
    excel.Application.Quit()
##########################################################################################
##########################################################################################
def runexcel(args):
    sheets = []
    """Convert LSC gencount to Excel file
    """
    sawerror = False
    print "Running gen2xls"
    if len(args) == 1:
        windll.user32.MessageBoxA(None,"Error: Please drag at least one file","gen2xls",0)
        sys.exit(1)
    try: 
        for fname in args[1:]:
            file = open(fname,"r")
            get_type = get_gen_type(file) 
            if get_type == -1:
                print 'file does not appear to be a gencount'
                continue  #skip file if not a gencount (only tests for presence of pagefeed char)
            if get_type[0] == 'laur':
                sheets =  get_l_sheets(file)  #LAURA'S FORMAT:
            else:
                if get_type[0] == 'serg':
                    sheets = get_s_sheets(file,get_type[1],get_type[2])  #SERGEI'S FORMAT:
            print "Processing %s" % fname
            make_sheets(sheets,fname)
        windll.user32.MessageBoxA(None,"Finished","gen2xls",0)
    except:
        traceback.print_exc()
        print "Errors occurred, please check the above messages"
        windll.user32.MessageBoxA(None,"Error: Problems occurred, please check them and try again","gen2xls",0)
##########################################################################################
##########################################################################################

if __name__ == "__main__":
    runexcel(sys.argv)

