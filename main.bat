set path=%cd%
START /WAIT %path%\ExcelToCsv.vbs input.xlsx input.csv
START /WAIT %path%\Extract.exe 
DEL %path%\input.csv
