#
# erppivotdragdrop.py:
# Load raw EPR data, clean up header info,
# insert additional data fields and build 5 pivot tables
# Support drag and drop of multiple spreadsheets 
#
import win32com.client as win32
win32c = win32.constants
import sys
import itertools
import re
import traceback
from ctypes import *

tablecount = itertools.count(1)

def runexcel(args):
    """Open the .xls file for RDA_MSR job and convert it 
    to fixed field
    """
    sawerror = False
    print "Running RDA_MSR convert to .ff"
    if len(args) == 1:
        windll.user32.MessageBoxA(None,"Error: Please drag at least one Excel file","RDA_MSR",0)
        sys.exit(1)
    try: 
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        for fname in args[1:]:
            if not re.search(r'\.(?i)xlsx?$',fname):
                print "Error: File %s doesn't seem to be an Excel file, expecting .xls or .xlsx file" % fname
                sawerror = True
                continue
            if not re.match('[A-Za-z]:',fname):
                print "Error: RDA_MSR doesn't support command line execution"
                print "       Please drag and drop the Excel file onto the program icon"
                sawerror = True
                continue
            print "Processing %s" % fname
            try:
                wb = excel.Workbooks.Open(fname)
            except:
                print "Failed to open Excel file %s, skipping" % fname
                sawerror = True
                continue

            try:
                ws = wb.Sheets('External List Report')
            except:
                print "Failed to open Sheet 'Sheet1' in file %s, skipping" % fname
                wb.Close()
                sawerror = True
                continue
            out_file_path = re.sub('\.xlsx?','.ff',fname)
            #out_file_path = fname + '.ff'
            #print out_file_path
            out_file = open(out_file_path,"w")
            xldata = ws.UsedRange.Value
            slug = str(xldata[5][1])
            fields = [4,15,15,15,3,40,40,40,2,8]
            for r in range(5,len(xldata) - 2):  #DISCARDS FIRST 5 ROWS OF HEADER DATA
                newline = ''
                if xldata[r][1] != None:                      # DETERMINES WHICH SLUG TO USE
                    slug = str(xldata[r][1])                  #
                if xldata[r][2] != None:                      #
                    for c in range(1,len(xldata[0]) - 1):  #WRITES OUT COLUMNS 2-9 TO TEMP VAR
                        if c == 1:                       
                            newline = newline + slug     
                        else:
                            newline = newline + xldata[r][c][0:fields[c - 1]].ljust(fields[c - 1])
                    newline = newline + str("{:,}".format(int(xldata[r][len(xldata[0]) - 1]))).ljust(8)  #WRITES LAST COLUMN FORMATED FOR NUMBER TO TEMP VAR 
                    out_file.write(newline + '\n')  #WRITES TEMP VAR TO OUTPUT FILE. 
        if sawerror:
            print "Errors occurred, please check the above messages"
            windll.user32.MessageBoxA(None,"Error: Problems occurred, please check them and try again","RDA_MSR",0)
        else:
            print "Finished creating %s" % out_file_path
            windll.user32.MessageBoxA(None,"Finished","RDA_MSR",0)
    except:
        traceback.print_exc()
        print "Errors occurred, please check the above messages"
        windll.user32.MessageBoxA(None,"Error: Problems occurred, please check them and try again","RDA_MSR",0)
    excel.Application.Quit()

if __name__ == "__main__":
    runexcel(sys.argv)
