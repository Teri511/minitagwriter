# minitagwriter
A quick and dirty tag making script to replace the ancient RadioShack Tag Wiz software.
Can be used to generate small and large tags for your friendly local RadioShack shop

requires the following python libraries: Pillow/xlrd

Instructions: Run the script via Python in a terminal/via command line with 2 Excel spreadsheets as its arguments, the first being the list of SKUs to print, the second being the catalog of SKUs available

Alternatively, for Windows users, a .bat file has been included to automatically run the tag writer. Simply name the catalog of SKUs "catalog.xlsx" and drop the Excel file of SKUs you wish to print on the included "runme.bat"

The ordering of the tag information is contained in the example Excel file "demo.xlsx"

tags should for now be printed in paint with all margins set as low as they can possibly go at 100% of size
