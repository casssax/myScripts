
def find_last(field):  #returns last populated position in field
    pos = 0
    last = 0
    for e in field:
        if e != ' ':
            last = pos
            pos = pos + 1
        else:
            pos = pos + 1
    return last
    
def all_blank(field):  #returns True if empty field
    test = True
    for e in field:
        if e != ' ':
            test = False
    return test

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


import sys
# ---------------------------------------------------------------
# import layout
# imput format i.e 30,20,1,10,10,60,60,30,2,5,4,3,4,5,9,80
layout_name = sys.argv[2]
layout_path = 'C:\\DATA_SAVE\\ff2delim\\' + layout_name
layout = open(layout_path,"r")
for line in layout:
    fields = parse_layout(line)
layout.close()
# ---------------------------------------------------------------
# import file
file_name = sys.argv[1]
file_path = 'C:\\DATA_SAVE\\ff2delim\\' + file_name
file = open(file_path,"r")
# ---------------------------------------------------------------
# create output file
out_file_path = 'C:\\DATA_SAVE\\ff2delim\\output_' + file_name
out_file = open(out_file_path,"w")
# ---------------------------------------------------------------
# process file
fields_list = []
first = 0
end = len(fields)
for f in fields:  #create list of lists:[start,end] positions for each field
    fields_list.append([first,first + f])
    first = first + f  
for line in file:
    count = 0    
    for field in fields_list:
        count = count + 1
        if all_blank(line[field[0]:field[1]]):  #check for empty field
            if count == end:
                out_file.write('\n') #carrage return if last field
            else:
                out_file.write(',')  #write comma if empty field
        else:
            last = find_last(line[field[0]:field[1]]) + 1  #find last populated position in field
            if count == end:
                out_file.write(line[field[0]:field[0] + last] + '\n') #carrage return if last field
            else:
                out_file.write(line[field[0]:field[0] + last] + ',') #comma if not last field
out_file.close()
file.close()







