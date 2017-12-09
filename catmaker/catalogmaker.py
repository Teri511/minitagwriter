import xlrd, xlwt, sys, os, datetime
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
            #print "ending a name early: " + str(burger)
            break
        if letter_count + len(split[word]) > 15:
            new_name = new_name + "?" + split[word]
            letter_count = len(split[word])
            line_count +=1
        else:
            new_name = new_name + " " + split[word]
            letter_count += len(split[word])
    return new_name[1:]

def compare_catalogs():
    if os.path.isfile("old_catalog.xlsx"):
        #set the flag on whether or not we've started a price discrepancy file
        new_xl_file = False
        #create the new workbook and sheet
        new_book = xlwt.Workbook()
        new_sheet = new_book.add_sheet("Sheet 0")
        new_sheet.write(0, 0,"SKU")
        new_sheet.write(0, 1, "Small Tag Quantity")
        new_sheet.write(0, 2, "Small Sale Quantity")
        new_sheet.write(0, 4, "Large Tag Quantity")
        new_sheet.write(0, 3, "Small Clr Quantity")
        new_sheet.write(0, 5, "Large Sale Quantity")
        new_sheet.write(0, 6, "Large Clr Quantity")
        new_sheet.write(0, 7, "Parts Quantity")
        #set the row count to one
        row_count = 1
        #open the 2 workbooks
        new = xlrd.open_workbook("../catalog.xlsx")
        old = xlrd.open_workbook("old_catalog.xlsx")
        n_sheet = new.sheet_by_index(0)
        o_sheet = old.sheet_by_index(0)
        #make a dictionary of old prices
        old_prices = {}
        for ind in range(1,o_sheet.nrows):
            if o_sheet.cell_value(ind,0) != 0.0:
                old_prices[o_sheet.cell_value(ind,0)] = o_sheet.cell_value(ind,4)

        #loop through the new worksheet to find all the differences
        for ind in range(1,n_sheet.nrows):
            #pull the price
            if n_sheet.cell_value(ind,0) != 0.0:
                try:
                    price = old_prices.pop(n_sheet.cell_value(ind,0))
                    if price != n_sheet.cell_value(ind,4):
                        print "Discrepancy found at SKU " + str(n_sheet.cell_value(ind,0))
                        new_xl_file = True
                        new_sheet.write(row_count, 0,n_sheet.cell_value(ind,0))
                        row_count += 1

                except KeyError:
                    #if the sku isn't found, report that a sku has been added
                    print "SKU " + str(n_sheet.cell_value(ind,0)) + "is new"
                    new_xl_file = True
                    new_sheet.write(row_count, 0,n_sheet.cell_value(ind,0))
                    row_count += 1
        #print out what SKUs haven't been found
        if str(old_prices.keys()) != "[]":
            print "The Following SKUs have been delisted"
            print str(old_prices.keys())

        if new_xl_file:
            hour = int(datetime.datetime.now().time().hour)
            minute = int(datetime.datetime.now().time().minute)
            second = int(datetime.datetime.now().time().second)

            time = (hour*360) + (minute*60) + second
            new_book.save("../discrepancies/"+str(datetime.datetime.now().date())+"-"+str(time)+".xlsx")

            #new_book.save("../discrepancies/"+str(101)+".xlsx")

    else:
        wb.save("old_catalog.xlsx")
        print "no previous catalog found, can't compare prices..."


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
            price = 0.0
            if str(r.cell_value(row,6)) == "variable" or str(r.cell_value(row,6)) == '':
                s.write(row, 3, 999.99)
            else:
                #print str(r.cell_type(row,6))
                s.write(row, 3, r.cell_value(row,6))
            #sale price
            #s.write(row, 4, int(r.cell_value(row,114)))
            if str(r.cell_value(row,114)) == "variable" or str(r.cell_value(row,114)) == '':
                s.write(row, 4, 999.99)
                price = 999.99
            else:
                s.write(row, 4, r.cell_value(row,114))
                price = r.cell_value(row,114)
            #print str(r.cell_value(0,114))
            #barcode
            s.write(row, 5, r.cell_value(row,4))

    #check to see if the old catalog exists and move it over, while saving the catalog
    #then start the comparator
    if os.path.isfile("../catalog.xlsx"):
        src = xlrd.open_workbook("../catalog.xlsx")
        dst = copy(src)
        dst.save("old_catalog.xlsx")
        wb.save("../catalog.xlsx")
        compare_catalogs()
    else:
        #if the catalog doesn't exist, just dump out after making it
        wb.save("../catalog.xlsx")
        print "Previous Catalog Doesn't Exist, Can't Compare Prices..."

    print "Complete!"

else:
    print "Incorrect num of args"
