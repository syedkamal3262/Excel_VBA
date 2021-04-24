Attribute VB_Name = "Module5"
Sub ReplaceBlanks()
'
' Macro1 Macro
'

'dECLARATIONS
Dim lastr As Long, lastr2 As Long, i As Long
    lastr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim sourcebook1 As String, sourcebook2 As String, sourcesheet As String
  

Columns("A:AF").ColumnWidth = 7.3

'Replace Blank Cells with "-"
Columns("I:AF").Select
     Selection.Replace What:="", Replacement:="-", LookAt:=xlPart

'Fill COLOR IN ITEM NAMES HAVING GB INFO
For Each cell In Range("I2:I" & lastr)
      If cell.Value Like "*GB*" Or cell.Value Like "*TB*" Then
            cell.Interior.ColorIndex = 35
        End If
Next

'save workbook
ActiveWorkbook.Save


End Sub


