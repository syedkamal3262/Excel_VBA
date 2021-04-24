Attribute VB_Name = "Module4"
Sub AdjustPNPcolunms()


'Changing the row Height
Cells.Select
    Range("A1").Activate
    Selection.RowHeight = 14.9

'Changing the 3rd-25the row Colunms
Columns("A:IX").ColumnWidth = 7
Columns("D").AutoFit
Columns("V:AE").ColumnWidth = 4
Columns("BI:BO").ColumnWidth = 2
Columns("CE:DC").ColumnWidth = 2
Columns("DN:EL").ColumnWidth = 2
Columns("FO:GB").ColumnWidth = 2
Columns("GO:HF").ColumnWidth = 2
Columns("GC:GN").ColumnWidth = 4
Columns("IA:IX").ColumnWidth = 2
Columns("DD:DH").ColumnWidth = 5
'BOLD first row
Range("A1:HS1").Font.Bold = True

'COLOR SPECIFIC CELL
Range("D1").Interior.Color = RGB(235, 116, 52)
Range("R1").Interior.Color = RGB(235, 116, 52)
Range("T1").Interior.Color = RGB(245, 203, 66)
Range("U1").Interior.Color = RGB(245, 203, 66)
Range("K1:O1").Interior.Color = RGB(252, 232, 3)
Range("AF1:AH1").Interior.Color = RGB(52, 207, 235)
Range("AK1:AN1").Interior.Color = RGB(252, 232, 3)
Range("AQ1:AT1").Interior.Color = RGB(252, 232, 3)
Range("AV1:AX1").Interior.Color = RGB(252, 232, 3)
Range("BA1:BC1").Interior.Color = RGB(252, 232, 3)
Range("BF1:BH1").Interior.Color = RGB(252, 232, 3)
Range("BP1:BR1").Interior.Color = RGB(52, 235, 195)
Range("BU1:BW1").Interior.Color = RGB(52, 235, 195)
Range("BZ1:CB1").Interior.Color = RGB(52, 235, 195)
Range("DI1:DM1").Interior.Color = RGB(189, 227, 50)
Range("EM1:EO1").Interior.Color = RGB(235, 119, 52)
Range("ER1:ET1").Interior.Color = RGB(235, 119, 52)
Range("EW1:EY1").Interior.Color = RGB(235, 119, 52)
Range("FB1:FD1").Interior.Color = RGB(50, 227, 82)
Range("FH1:FN1").Interior.Color = RGB(227, 50, 76)
Range("HG1:HK1").Interior.Color = RGB(252, 232, 3)
Range("HV1:HY1").Interior.Color = RGB(235, 52, 131)




End Sub
