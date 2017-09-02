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
# small tag max lines: 4                   #
# # of small tags per col/row: (5,8)       #
# big tag info to be determined            #
############################################

import xlrd, sys
from PIL import Image, ImageFont, ImageDraw

#set up the global canvas/draw space/fonts and number of each tag generated
small_canvas = Image.new("RGB",(2430,3008),(150,150,255))
draw = ImageDraw.Draw(small_canvas)
font = ImageFont.truetype("MyriadPro-Cond.ttf",35)

############################################
# add_tag_to_canvas                        #
# adds tag of the specified type with info #
# provided to the correct global canvas    #
#                                          #
# prices: tuple, 0 is normal price         #
#                1 is sale price           #
# text: tuple, 0 is name of item           #
#              1 is sku of item            #
#              2 is description of item    #
# quantities: tuple, 0 is # of sm.reg tags #
#                   1 is # of sm.sale tags #
#                   2 is # of sm.clr  tags #
#                   3 is # of lg.reg tags  #
#                   4 is # of lg.sale tags #
#                   5 is # of lg.clr tags  #
#                   6 is # of parts tags   #
############################################
def add_tag_to_canvas(prices,text,quantities):

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
            nprices = (sheet.cell_value(row,3),sheet.cell_value(row,4))
            ntext = (sheet.cell_value(row,1),sheet.cell_value(row,0),sheet.cell_value(row,2))
            nquantities = (sheet.cell_value(row,5),sheet.cell_value(row,6),sheet.cell_value(row,7),sheet.cell_value(row,8),sheet.cell_value(row,9),sheet.cell_value(row,10),sheet.cell_value(row,11))
            add_tag_to_canvas(nprices,ntext,nquantities)
else:
    print "incorrect number of arguments"
