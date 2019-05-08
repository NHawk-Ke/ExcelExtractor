Dim file, WB

With CreateObject("Excel.Application")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    src_file = objFSO.GetAbsolutePathName("result.csv")
    Set WB = .Workbooks.Open(src_file)
    WB.SaveAs Replace(WB.FullName, ".csv", ".xlsx"), 51
    WB.Close False
    .Quit
End With