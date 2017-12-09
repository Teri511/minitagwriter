cd C:\Users\CORP User\minitagwriter
del .\output\*.png
python taggen.py %1 catalog.xlsx
echo Press any key to continue
pause