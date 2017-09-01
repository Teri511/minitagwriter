#xlsx tag processor
#takes in an xlsx and an int
#prints tags on rippable paper or with outlines to cut out manually based on the int
#only operates on the first sheet in the xlsx for now possibly

############################################
# Current Tag Info                         #
# tags are to be printed through the image #
# editor of choice, with the fit option    #
# chosen and margins set to 0.17           #
# dpi should be 300                        #
# small tag size in px: (449,331)          #
# small tag offsets in px: (93,258)        #
# small tag text offset: add (10,200) to xy#
# small tag char count: 15 uppr/17  lowr   #
# small tag max lines: 4
# # of small tags per col/row: (5,8)       #
# big tag info to be determined            #
############################################

import xlrd, sys
from PIL import Image, ImageFont, ImageDraw

#set up the global canvas/draw space/fonts and number of each tag generated
small_canvas = Image.new("RGB",(2430,3008),(150,150,255))
draw = ImageDraw.Draw(small_canvas)
font = ImageFont.truetype("MyriadPro-Cond.ttf",35)

def add_tag_to_canvas():

    tag = Image.open("images/small/sclear.png")
    print "lolno"

if len(sys.argv) == 2:

    #open up the excel file and get the first sheet
    #NOTE: the columns in the sheet only exist up until the farthest filled cell, same goes for rows
    #so if you only have data in cell 0,0 you can't ask for cell 1,1. However if you have data in cell 0,0 and 10,10
    #you can ask for data all the way up to 10,10 and just get nulls
    book = xlrd.open_workbook(filename=sys.argv[1])
    sheet = book.sheet_by_index(0)

    #loop across each row
    for row in range(0,sheet.nrows):
        #check to make sure our users didn't skip any rows
        if sheet.cell_value(row,0) != '':
            #check for valid data across the row
            #pull the info and pass to the tag maker routine
            print "yo"
else:
    print "incorrect number of arguments"
