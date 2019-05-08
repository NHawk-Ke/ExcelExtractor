set path=%cd%
ECHO %path%
START /WAIT %path%\ExcelToCsv.vbs input.xlsx temp.csv
START /WAIT %path%\processCSV.vbs 
START /WAIT %path%\CsvToExcel.vbs
DEL %path%\temp.csv
DEL %path%\result.csv