Attribute VB_Name = "Module1"
Sub DeleteNSPextras()


'change formula to value
    For Each cell In ActiveSheet.UsedRange.Cells
        If cell.Formula Like "*VLOOKUP*" Then
            cell.Formula = cell.Value
        End If
    Next cell
    For Each cell In ActiveSheet.UsedRange.Cells
        If cell.Formula Like "*=*" Then
            cell.Formula = cell.Value
        End If
    Next cell
'AUTO FIT
Columns("L:Q").AutoFit
Columns("E").AutoFit
Columns("B").AutoFit
Columns("J").AutoFit

'save workbook
ActiveWorkbook.Save

'Delete URL Column
Columns("T").EntireColumn.Delete

'COLOR SPECIFIC COLUNMS
Range("J1").Interior.Color = RGB(252, 32, 3)
Range("B1").Interior.Color = RGB(245, 203, 66)
Range("E1").Interior.Color = RGB(252, 232, 3)
Range("L1").Interior.Color = RGB(52, 207, 235)
Range("M1:N1").Interior.Color = RGB(252, 232, 3)
Range("O1").Interior.Color = RGB(52, 235, 195)
Range("P1").Interior.Color = RGB(252, 232, 3)
Range("Q1").Interior.Color = RGB(235, 119, 52)
Range("S1").Interior.Color = RGB(235, 97, 52)


'Delete Extra
Columns("U:AE").EntireColumn.Delete

End Sub
