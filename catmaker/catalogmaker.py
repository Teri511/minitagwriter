import xlrd, xlwt, sys, os
from xlutils.copy import copy

def split_name(old_name):
    #trim off the first 8 chars and split up the remainder
    temp_name = old_name[8:]
    split = temp_name.split()
    letter_count = 0
    line_count = 0
    new_name = ""
    for word in range(0,len(split)):
        if line_count > 3:
            print "ending a name early: " + str(burger)
            break
        if letter_count + len(split[word]) > 15:
            new_name = new_name + "?" + split[word]
            letter_count = len(split[word])
            line_count +=1
        else:
            new_name = new_name + " " + split[word]
            letter_count += len(split[word])
    return new_name[1:]

############################################################

if len(sys.argv) == 2:
    print "Opening Catalog File"
    rb = xlrd.open_workbook(sys.argv[1])
    wb = xlwt.Workbook()

    print "Succesfully opened file"
    r = rb.sheet_by_index(0)
    s = wb.add_sheet("Sheet 0")
    print "Adding Entries to new Catalog"
    for row in range(0,r.nrows):
        if row == 0:
            #setup the topmost row
            #sku
            s.write(row,0,"Sku")
            #name
            s.write(row, 1, "Name")
            #desc is empty
            s.write(row, 2, "Description")
            #reg price
            s.write(row, 3, "Reg Price")
            #sale price
            s.write(row, 4, "Sale Price")
            #barcode
            s.write(row, 5, "Barcode")
        else:
            #populate the rest of the rows with necessary info
            #info from old catalog goes:
            #col 1: name
            #col 2: sku
            #col 4: barcode
            #col 6: reg price
            #col DK(base 26) or 114(base 10): sale price
            #info in the outputted catalog file goes: sku,name,desc,reg price,sale price, barcode
            #print str(row)
            name  = split_name(str(r.cell_value(row,1)))
            #sku
            #throw in a dummy sku for blank cells
            sku = 0
            if str(r.cell_value(row,2)) != '':
                sku = r.cell_value(row,2)
            s.write(row,0,sku)
            #name
            s.write(row, 1, name)
            #desc is empty
            #reg price
            if str(r.cell_value(row,6)) == "variable" or str(r.cell_value(row,6)) == '':
                s.write(row, 3, 999.99)
            else:
                #print str(r.cell_type(row,6))
                s.write(row, 3, r.cell_value(row,6))
            #sale price
            #s.write(row, 4, int(r.cell_value(row,114)))
            if str(r.cell_value(row,114)) == "variable" or str(r.cell_value(row,114)) == '':
                s.write(row, 4, 999.99)
            else:
                s.write(row, 4, r.cell_value(row,114))
            #print str(r.cell_value(0,114))
            #barcode
            s.write(row, 5, r.cell_value(row,4))

    wb.save("catalog.xlsx")
    print "Complete!"
else:
    print "Incorrect num of args"
