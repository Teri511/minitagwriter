#xlsx tag processor
#takes in an xlsx and an int
#prints tags on rippable paper or with outlines to cut out manually based on the int
#only operates on the first sheet in the xlsx for now possibly

############################################
# Current Tag Info                         #
# tags are to be printed through the image #
# editor of choice, with the actual size   #
# option chosen and margins set to 0.17    #
# dpi should be 300                        #
#                                          #
# SMALL TAG INFO:                          #
#                                          #
# small tag size in px: (449,331)          #
# small tag offsets in px: (93,258)        #
# small tag text offset: add (10,200) to xy#
# small tag char count: 15 uppr/17  lowr   #
# small tag max lines: 4                   #
# # of small tags per col/row: (5,8)       #
#                                          #
# BIG TAG INFO:                            #
#                                          #
# big tag size in px: (1051,602)           #
# big tag offsets in px: (160,107)         #
# big tag header offset in px: (200,150)   #
# big tag header char count:
############################################

import xlrd, sys, os
from PIL import Image, ImageFont, ImageDraw

#set up the global canvas/draw space/fonts and number of each tag generated
small_canvas = Image.new("RGB",(2430,3008),(150,150,255))
big_canvas = Image.new("RGB",(2430,3108),(255,150,150))
draw = ImageDraw.Draw(small_canvas)
b_draw = ImageDraw.Draw(big_canvas)
small_font = ImageFont.truetype("MyriadPro-Cond.ttf",35)
sku_font = ImageFont.truetype("MyriadPro-Cond.ttf",35)
dollar_font = ImageFont.truetype("MyriadPro-Cond.ttf",85)
cent_font = ImageFont.truetype("MyriadPro-Cond.ttf",40)
b_dollar_font = ImageFont.truetype("MyriadPro-Cond.ttf",220)
b_cent_font = ImageFont.truetype("MyriadPro-Cond.ttf",110)
extra_font = ImageFont.truetype("MyriadPro-Cond.ttf",55)
b_extra_font = ImageFont.truetype("MyriadPro-Cond.ttf",190)
small_tag_count = 0
small_curr_row = 0
small_page_count = 0
big_tag_count = 0
big_curr_row = 0
big_page_count = 0
parts_tag_count = 0
parts_curr_row = 0
parts_page_count = 0



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
    global small_canvas
    global big_canvas
    global draw
    global b_draw
    global small_tag_count
    global small_curr_row
    global small_page_count
    global big_tag_count
    global big_curr_row
    global big_page_count

    #for each quantity, generate that many tags
    if quantities[0] > 0:
        #generate small normal tags
        print "generating small normal tags"
        tag = Image.open("images/small/sbase.png")
        name = str(text[0]).split("?")
        for count in range(0,quantities[1]):
            #check to see if we've filled a page
            if small_tag_count > 39:
                #dump the current canvas to a file and start a new one
                small_canvas.save("output/small page "+ str(small_page_count) + ".png","PNG",dpi=(300,300))
                small_tag_count = 0
                small_curr_row = 0
                small_page_count += 1
                small_canvas = Image.new("RGB",(2430,3008),(150,150,255))
                draw = ImageDraw.Draw(small_canvas)
            #split the string using evil casting magic
            dollar = str(int(prices[0]))
            #I'm so sorry
            #take the price, convert it to a string and take the last 2 chars from it, these will always be the cents
            cents = str(prices[0])[len(str(prices[0]))-2] + str(prices[0])[len(str(prices[0]))-1]
            #hacky check to get rid of the "no cent" bug
            if cents[0] == ".":
                cents=cents[1] + "0"
            px,py = draw.textsize(dollar,font=dollar_font)
            #paste in the tag
            small_canvas.paste(tag, (((small_tag_count%5)*449)+93,(small_curr_row*335)+258))
            #add the text here
            #first the name
            for line in range(0,len(name)):
                draw.text((((small_tag_count%5)*449)+103,(small_curr_row*335)+458+(30*line)),name[line],font=small_font,fill=(0,0,0,255))
            #now the dollar
            draw.text((((small_tag_count%5)*449)+458-px,(small_curr_row*335)+463),dollar,font=dollar_font,fill=(0,0,0,255))
            #now the cents
            draw.text((((small_tag_count%5)*449)+458,(small_curr_row*335)+463),cents,font=cent_font,fill=(0,0,0,255))
            #next, the sku
            draw.text((((small_tag_count%5)*449)+393,(small_curr_row*335)+548),str(text[1])[0:3]+"-"+str(text[1])[2:6],font=sku_font,fill=(0,0,0,255))


            #adjust row position
            if small_tag_count%5 == 4:
                small_curr_row += 1
            small_tag_count +=1
        small_canvas.save("output/small page "+str(small_page_count)+".png","PNG",dpi=(300,300))
    if quantities[1] > 0:
        #generate small sale tags
        print "generating small sale tags"
        tag = Image.open("images/small/ssale.png")
        name = str(text[0]).split("?")
        for count in range(0,quantities[1]):
            #check to see if we've filled a page
            if small_tag_count > 39:
                #dump the current canvas to a file and start a new one
                small_canvas.save("output/small page "+ str(small_page_count) + ".png","PNG",dpi=(300,300))
                small_tag_count = 0
                small_curr_row = 0
                small_page_count += 1
                small_canvas = Image.new("RGB",(2430,3008),(150,150,255))
                draw = ImageDraw.Draw(small_canvas)
            #split the string using evil casting magic
            dollar = str(int(prices[1]))
            #I'm so sorry
            #take the price, convert it to a string and take the last 2 chars from it, these will always be the cents
            cents = str(prices[1])[len(str(prices[1]))-2] + str(prices[1])[len(str(prices[1]))-1]
            #hacky check to get rid of the "no cent" bug
            if cents[0] == ".":
                cents=cents[1] + "0"
            px,py = draw.textsize(dollar,font=dollar_font)
            #paste in the tag
            small_canvas.paste(tag, (((small_tag_count%5)*449)+93,(small_curr_row*335)+258))
            #add the text here
            #first the name
            for line in range(0,len(name)):
                draw.text((((small_tag_count%5)*449)+103,(small_curr_row*335)+458+(30*line)),name[line],font=small_font,fill=(0,0,0,255))
            #now the dollar
            draw.text((((small_tag_count%5)*449)+458-px,(small_curr_row*335)+483),dollar,font=dollar_font,fill=(255,255,255,255))
            #now the cents
            draw.text((((small_tag_count%5)*449)+458,(small_curr_row*335)+483),cents,font=cent_font,fill=(255,255,255,255))
            #next, the sku
            draw.text((((small_tag_count%5)*449)+393,(small_curr_row*335)+548),str(text[1])[0:3]+"-"+str(text[1])[2:6],font=sku_font,fill=(255,255,255,255))
            #since its a sale tag, add the word "sale"
            draw.text((((small_tag_count%5)*449)+410,(small_curr_row*335)+438),"Sale",font=extra_font,fill=(255,255,255,255))
            #and now: the reg price

            #adjust row position
            if small_tag_count%5 == 4:
                small_curr_row += 1
            small_tag_count +=1
        small_canvas.save("output/small page "+str(small_page_count)+".png","PNG",dpi=(300,300))
    if quantities[2] > 0:
        #generate small clear tags
        print "generating small clear tags"
        tag = Image.open("images/small/sclear.png")
        for count in range(0,quantities[1]):
            small_canvas.paste(tag, (((small_tag_count%5)*449)+93,(small_curr_row*335)+258))
            #add the text here
            if small_tag_count%5 == 4:
                small_curr_row += 1
            small_tag_count +=1
            #small_canvas.show()
            #check to see if we've filled a page goes here
    if quantities[3] > 0:
        #generate large normal tags
        print "generating large normal tags"
        tag = Image.open("images/big/bbase.png")

        name = str(text[0]).split("?")
        for count in range(0,quantities[3]):
            #check to see if we've filled a page
            if big_tag_count > 9:
                #dump the current canvas to a file and start a new one
                big_canvas.save("output/big page "+ str(big_page_count) + ".png","PNG",dpi=(300,300))
                big_tag_count = 0
                big_curr_row = 0
                big_page_count += 1
                big_canvas = Image.new("RGB",(2430,3008),(150,150,255))
                b_draw = ImageDraw.Draw(big_canvas)
            #split the string using evil casting magic
            dollar = str(int(prices[0]))
            #I'm so sorry
            #take the price, convert it to a string and take the last 2 chars from it, these will always be the cents
            cents = str(prices[0])[len(str(prices[0]))-2] + str(prices[0])[len(str(prices[0]))-1]
            #hacky check to get rid of the "no cent" bug
            if cents[0] == ".":
                cents=cents[1] + "0"
            px,py = b_draw.textsize(dollar,font=b_dollar_font)
            #paste in the tag
            big_canvas.paste(tag, (((big_tag_count%2)*1051)+160,(big_curr_row*600)+107))
            #add the text here
            #first the name
            for line in range(0,len(name)):
                b_draw.text((((big_tag_count%2)*1051)+200,(big_curr_row*600)+150+(70*line)),name[line],font=dollar_font,fill=(0,0,0,255))
            #now the dollar
            b_draw.text((((big_tag_count%2)*1051)+1075-px,(big_curr_row*600)+300),dollar,font=b_dollar_font,fill=(0,0,0,255))
            #now the cents
            b_draw.text((((big_tag_count%2)*1051)+1075,(big_curr_row*600)+300),cents,font=b_cent_font,fill=(0,0,0,255))
            #next, the sku
            b_draw.text((((big_tag_count%2)*1051)+925,(big_curr_row*600)+600),str(text[1])[0:3]+"-"+str(text[1])[2:6],font=cent_font,fill=(0,0,0,255))

            #adjust row position
            if big_tag_count%2 == 1:
                big_curr_row += 1
            big_tag_count +=1
        big_canvas.save("output/big page "+ str(big_page_count) + ".png","PNG",dpi=(300,300))
    if quantities[4] > 0:
        #generate large sale tags
        print "generating large sale tags"
        tag = Image.open("images/big/bsale.png")

        name = str(text[0]).split("?")
        for count in range(0,quantities[3]):
            #check to see if we've filled a page
            if big_tag_count > 9:
                #dump the current canvas to a file and start a new one
                big_canvas.save("output/big page "+ str(big_page_count) + ".png","PNG",dpi=(300,300))
                big_tag_count = 0
                big_curr_row = 0
                big_page_count += 1
                big_canvas = Image.new("RGB",(2430,3008),(150,150,255))
                b_draw = ImageDraw.Draw(big_canvas)
            #split the string using evil casting magic
            dollar = str(int(prices[1]))
            #I'm so sorry
            #take the price, convert it to a string and take the last 2 chars from it, these will always be the cents
            cents = str(prices[1])[len(str(prices[1]))-2] + str(prices[1])[len(str(prices[1]))-1]
            #hacky check to get rid of the "no cent" bug
            if cents[0] == ".":
                cents=cents[1] + "0"
            px,py = b_draw.textsize(dollar,font=b_dollar_font)
            #paste in the tag
            big_canvas.paste(tag, (((big_tag_count%2)*1051)+160,(big_curr_row*600)+107))
            #add the text here
            #first the name
            for line in range(0,len(name)):
                b_draw.text((((big_tag_count%2)*1051)+200,(big_curr_row*600)+150+(70*line)),name[line],font=dollar_font,fill=(0,0,0,255))
            #now the dollar
            b_draw.text((((big_tag_count%2)*1051)+1075-px,(big_curr_row*600)+325),dollar,font=b_dollar_font,fill=(255,255,255,255))
            #now the cents
            b_draw.text((((big_tag_count%2)*1051)+1075,(big_curr_row*600)+325),cents,font=b_cent_font,fill=(255,255,255,255))
            #next, the sku
            b_draw.text((((big_tag_count%2)*1051)+925,(big_curr_row*600)+600),str(text[1])[0:3]+"-"+str(text[1])[2:6],font=cent_font,fill=(255,255,255,255))
            #since its a sale tag, add the word "sale"
            b_draw.text((((big_tag_count%2)*1051)+885,(big_curr_row*600)+150),"Sale",font=b_extra_font,fill=(255,255,255,255))

            #adjust row position
            if big_tag_count%2 == 1:
                big_curr_row += 1
            big_tag_count +=1
        big_canvas.save("output/big page "+ str(big_page_count) + ".png","PNG",dpi=(300,300))
    if quantities[5] > 0:
        #generate small clear tags
        print "generating large clear tags"
        tag = Image.open("images/big/bclear.png")
    if quantities[6] > 0:
        #generate parts tags
        print "generating parts tags"



###################################################
#start of Tag genning script

if len(sys.argv) == 3:
    need_big = 0
    need_small = 0
    #open up the excel file and get the first sheet
    #NOTE: the columns in the sheet only exist up until the farthest filled cell, same goes for rows
    #so if you only have data in cell 0,0 you can't ask for cell 1,1. However if you have data in cell 0,0 and 10,10
    #you can ask for data all the way up to 10,10 and just get nulls
    book = xlrd.open_workbook(filename=sys.argv[1])
    sheet = book.sheet_by_index(0)
    db_book = xlrd.open_workbook(filename=sys.argv[2])
    db_sheet = db_book.sheet_by_index(0)

    catalog = {}

    #build a dictionary of SKUs to their data
    for sku in range(1,db_sheet.nrows):
        #more error checking here
        #print "adding sku: " + str(db_sheet.cell_value(sku,0))
        catalog[db_sheet.cell_value(sku,0)] = [str(db_sheet.cell_value(sku,1)),str(db_sheet.cell_value(sku,2)),db_sheet.cell_value(sku,3)]

    #loop across each row
    for row in range(1,sheet.nrows):
        #check to make sure our users didn't skip any rows
        if sheet.cell_value(row,0) != '':
            #error checking goes here
            #find the sku in the catalog and pull its info
            #info goes: name,description,normal price
            if str(sheet.cell_value(row,0))[3]=="-":
                temp = int(str(sheet.cell_value(row,0))[0:3] + str(sheet.cell_value(row,0))[4:])
            else:
                temp = int(sheet.cell_value(row,0))
            info = catalog[temp]
            #add the info
            if int(sheet.cell_value(row,2)) > 0 or int(sheet.cell_value(row,3)) > 0 or int(sheet.cell_value(row,4)) > 0:
                need_small = 1;
            if int(sheet.cell_value(row,5)) > 0 or int(sheet.cell_value(row,6)) > 0 or int(sheet.cell_value(row,7)) > 0:
                need_big = 1;
            nprices = (info[2],sheet.cell_value(row,1))
            ntext = (info[0],int(temp),info[1])
            nquantities = (int(sheet.cell_value(row,2)),int(sheet.cell_value(row,3)),int(sheet.cell_value(row,4)),int(sheet.cell_value(row,5)),int(sheet.cell_value(row,6)),int(sheet.cell_value(row,7)),int(sheet.cell_value(row,8)))
            print "genning tags for sku: " + str(temp)
            add_tag_to_canvas(nprices,ntext,nquantities)
    print "You need to grab " + str(small_page_count+need_small) + " small tag sheets and " + str(big_page_count+need_big) + " large tag sheets"
    raw_input("Press Enter to continue...")
else:
    print "incorrect number of arguments"
