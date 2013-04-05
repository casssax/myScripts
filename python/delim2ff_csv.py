# ----------------------------------------------------------------
# This program converts a delimited text file to fixed field
# A layout needs to be provided with the lengths for each field
# layout example: 30,20,1,10,10,60,60,30,2,5,4,3,4,5,9,80
# (a negitive number in the layout will right justify the field)
#-------------------------------------------------------------------
# example: C:>python delim2ff.py filename.txt layout.txt    (default comma delimited)
# example: C:>python delim2ff.py filename.txt layout.txt -t (for tab delimited files)
# example: C:>python delim2ff.py filename.txt layout.txt -p (for pipe delimited files)
# ----------------------------------------------------------------
def parse_layout(layout): #converts layout file to list
    fields = []
    pos = 0
    while pos != -1:
        pos = layout.find(',')
        if pos != -1:
            fields.append(int(layout[0:pos]))
            layout = layout[pos + 1:]
        else:
            fields.append(int(layout))
    return fields 
#-----------------------------------------------------------------
def outname(file_name):
    prefix = file_name.find('.')
    return file_name[:prefix]
#-----------------------------------------------------------------
def fix_field_qual(file):
    total = 0
    for line in file:
        total = total + 1
        count = 0
        space = ' '
        new = ''
        pos = 0
        delm = 0
        qual= '\"'   #quote qualifier
        while delm != -1:   #not last field in record
            delm = line.find(delim_type)   #find pos of next delimeter
            if delm == 0:   #if empty field
                new = new + space.ljust(fields[count])   #fill empty field with spaces
                count = count + 1   #increment field counter
                line = line[1:]   #trims delim of empty field off of front of record
            elif delm != -1:   #if not last field in record
                if (delm - 2) > abs(fields[count]):   #truncataes if field longer than layout (for header records)
                    new = new + line[1:abs(fields[count]) + 1]   
                else:
                    if fields[count] < 1:
                        new = new + line[1:delm - 1].translate(None,"\"").rjust(abs(fields[count]))   #right justify if all digits
                    else:
                        new = new + line[1:delm - 1].translate(None,"\"").ljust(fields[count])
                count = count + 1
                line = line[delm + 1:]   #increments field number
            else:
                if pos > abs(fields[count]):   #truncataes if field longer than layout (for header records)
                    new = new + line[1:abs(fields[count]) + 1].translate(None,"\"")   
                elif fields[count] < 1:
                    new = new + (line[1:-2].translate(None,"\"").rjust(abs(fields[count])))   #right justify if all digits
                else:
                    new = new + (line[1:-2].translate(None,"\"").ljust(fields[count]))
                out_file.write(new + '\n')
    out_file.close()
    file.close()
    print 'total records processed: {}'.format(total)
#-----------------------------------------------------------------
def fix_field(file):
    total = 0
    for line in file:
        total = total + 1
        count = 0
        num_fields = len(line)
        new = ''
        pos = 0
        for field in line:
            if len(field) > abs(fields[count]):
                new = new + field[:abs(fields[count])]   #truncataes if field longer than layout (for header records)
                count = count + 1
            else:
                if fields[count] < 1:
                    new = new + field.rjust(abs(fields[count]))   #right justify if all digits
                else:
                    new = new + field.ljust(fields[count])
                count = count + 1
        out_file.write(new + '\n')
    out_file.close()
#    file.close()
    print 'total records processed: {}'.format(total)
#-----------------------------------------------------------------
#-----------------------------------------------------------------
import sys
import csv
if len(sys.argv)==1:
    print('''\
----------------------------------------------------------------
This program converts a delimited text file to fixed field
A layout needs to be provided with the lengths for each field
layout example: 30,20,1,10,10,60,60,30,2,5,4,3,4,5,9,80
(a negitive number in the layout will right justify the field)
-------------------------------------------------------------------
examples:
C:>python delim2ff.py filename.txt layout.txt    (default comma delimited)
C:>python delim2ff.py filename.txt layout.txt -t (for tab delimited files)
C:>python delim2ff.py filename.txt layout.txt -p (for pipe delimited files)
any of the above 3 commands can be followed by a ' -q' for quote qualified files''')
    sys.exit(1)
# ---------------------------------------------------------------
# import file
try:
    file_name = sys.argv[1]
except:
    sys.exit('ERROR: missing FILENAME file. \n example: C:>python delim2ff.py FILENAME LAYOUT [-t,-p] \n optional -t for tab, -p for pipe. defaults to comma delimited')
#file_name = sys.argv[1]
file_path = 'C:\\DATA_SAVE\\delim2ff\\' + file_name
file = open(file_path,"r")
# ---------------------------------------------------------------
# ---------------------------------------------------------------
# delimiter_type case statement defaults to ','
delim_dict = {'-t':'\t',
            '-T':'\t',
            '-p':'|',
            '-P':'|'}
if len(sys.argv) >= 4:
    delim_type = delim_dict[sys.argv[3]]
else:
    delim_type = ','
#-----------------------------------------------------------------
#-----------------------------------------------------------------
# flag for quote qualified text. 
quote_q = False
if len(sys.argv) == 4:
    if sys.argv[3] == '-q':
        quote_q = True
if len(sys.argv) == 5:
    if sys.argv[4] == '-q':
        quote_q = True
if len(sys.argv) == 4:
    if sys.argv[3] == '-Q':
        quote_q = True
if len(sys.argv) == 5:
    if sys.argv[4] == '-Q':
        quote_q = True
# ----------------------------------------------------------------
# import layout
try:
    layout_name = sys.argv[2]
except:
    sys.exit('ERROR: missing LAYOUT file. \n example: C:>python delim2ff.py FILENAME LAYOUT [-t,-p] \n optional -t for tab, -p for pipe. defaults to comma delimited')
#layout_name = sys.argv[2]
layout_path = 'C:\\DATA_SAVE\\delim2ff\\' + layout_name
layout = open(layout_path,"r")
for line in layout:
    fields = parse_layout(line)
layout.close()

# create output file
#out_name = outname(file_name)
out_file_path = 'C:\\DATA_SAVE\\delim2ff\\' + outname(file_name) + '.ff'
out_file = open(out_file_path,"w")
# ---------------------------------------------------------------
# process file
print('layout: ',fields)
data = csv.reader(file, delimiter=delim_type)
data = [row for row in data]
file.close()
fix_field(data)