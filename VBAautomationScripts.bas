Attribute VB_Name = "Module22"


Sub NSPValidation()
'
' NSPValidation Macro
' Edited By Kamal
'

'SAVE workbook
ActiveWorkbook.Save


'Set Excel view
ActiveWindow.Zoom = 98

    If MsgBox("  RUN " & Chr(13) & Chr(10) & "  NSP Hunter X Crawling " & Chr(13) & Chr(10) & "  Macro ", vbYesNo) = vbNo Then
        Exit Sub
    End If
    Dim lastr As Long, lastr2 As Long, i As Long
    lastr = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Dim sourcebook1 As String, sourcebook2 As String, sourcesheet As String
    
    sourcebook1 = ActiveWorkbook.Name:
    ActiveSheet.Name = "Sheet 1"
    
    'Regex Object Declaration
    Dim RE As Object, RegMC, checkingrange As Variant
    Set RE = CreateObject("VBScript.RegExp")
    RE.IgnoreCase = True
    RE.Global = True
    
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
    
    'manual formula check
    If Application.Calculation = xlCalculationManual Then
        Application.Calculation = xlCalculationAutomatic
    End If
    
    'Filters Reset
    ActiveSheet.UsedRange.AutoFilter
    ActiveSheet.UsedRange.AutoFilter

'Change Header Color
Range("A1:T1").Interior.Color = RGB(199, 197, 197)

'Changing the 3rd-25the row Height
Cells.Select
    Range("A1").Activate
    Selection.RowHeight = 14.5

'Changing the 3rd-25the row Colunms
Columns("A:H").ColumnWidth = 4
Columns("I").ColumnWidth = 40
Columns("J:Q").AutoFit
Columns("R:T").ColumnWidth = 5
Columns("B").AutoFit
Columns("C").ColumnWidth = 10
Columns("E").ColumnWidth = 7

'BOLD first row
Range("A1:Z1").Font.Bold = True
Cells.Select
    Range("A1:Z1").Activate
    With Selection
        .HorizontalAlignment = xlLeft
    End With
Range("B2").Font.Bold = True
Range("C2").Font.Bold = True
Range("B2").Interior.Color = RGB(240, 252, 3)
Range("C2").Interior.Color = RGB(240, 252, 3)

'Extra Sheets check
If Worksheets.Count > 1 Then
    MsgBox ("File has Multiple Sheets." & vbNewLine & "Click ok to continue.")
End If


nspheaders = Array("ETI_RECORD_ID", "ETI_DATE", "ETI_TimeStamp", "ETI_COUNTRY_ID", "ETI_DPG_ID", "ETI_PERIOD_ID_WEEK", "ETI_PERIOD_ID_MONTH", "ETI_RETAILER_ID", "ETI_ITEM_NAME", "ETI_PRICE", "ETI_INC_VAT", "ETI_BRAND", "ETI_Storage_Capacity", "ETI_RAM", "ETI_Color", "ETI_Screen_Size", "ETI_Manufacturer_Number", "ETI_CURRENCY", "ETI_Cellular_Connectivity", "WEBLINK", "", "", "")
'6 All Headers check
For i = 1 To 20
    If Not Cells(1, i) = nspheaders(i - 1) Then
        MsgBox ("Header " & i & " is """ & Cells(1, i) & """ instead of """ & nspheaders(i - 1) & """" & vbNewLine & "Please Fix Header")
        Cells(1, i).Select
        Exit Sub
    End If
Next i

'Excel Hash Error / formula check
Call GLOBALHASHERRORCAPTURE


'Fill CELLULAR CONNECTIVITY "1" against "32647"
For Each cell In Range("E2:E" & lastr)
      If cell.Value = "32647" Then
            cell.Offset(0, 14).Value = "1"
        End If
Next

'Trim/clean, newline character removal
Dim newlinechar As Range
Set newlinechar = Nothing
For Each cell In Range("A1:P" & lastr)
    If InStr(1, cell.Value, Chr(10), vbTextCompare) > 0 Then
        Set newlinechar = cell
        cell.Replace What:=Chr(10), Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows
'        With cell.Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .Color = 2859369
'            .TintAndShade = 0
'            .PatternTintAndShade = 0
'        End With
'        With cell.Font
'            .ThemeColor = xlThemeColorDark1
'            .TintAndShade = 0
'            .Bold = True
'        End With
    End If
    
    If cell.Column = 2 Then
    
    Else
        cell.Value = Application.WorksheetFunction.clean((Trim(cell.Value)))
    End If
Next


'Replace Blank Cells with "-"
Columns("I:T").Select
     Selection.Replace What:="", Replacement:="-", LookAt:=xlPart

'CHECK BLANK CELLS
For Each cell In Range("B2:E" & lastr)
        If cell.Value = "" Then
            MsgBox ("BLANK VALUES at " & cell.Address(0, 0))
            cell.Select
            Exit Sub
        End If
Next
For Each cell In Range("H2:T" & lastr)
        If cell.Value = "" Then
            MsgBox ("BLANK VALUES at " & cell.Address(0, 0))
            cell.Select
            Exit Sub
        End If
Next


'Blanks and newline character Check
For Each cell In Range("A1:P" & lastr)
        If cell.Value = "" And Not (cell.Column = 1 Or cell.Column = 2 Or cell.Column = 6 Or cell.Column = 7) Then
            MsgBox ("Blank Value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
        If InStr(1, cell.Value, Chr(10)) > 0 Then
            
            MsgBox ("Newline Character at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next

'ETI_INC_VAT
For Each cell In Range("K2:K" & lastr)
        If cell.Value <> "Yes" Then
            MsgBox ("Error: Incorrect ETI_INC_VAT  Value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next

'DPGID Check
For Each cell In Range("E2:E" & lastr)
        If cell.Value <> "32647" And cell.Value <> "321373" Then
            MsgBox ("Error: Incorrect DPG ID  Value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next

'ETI_COUNTRY_ID
For Each cell In Range("D2:D" & lastr)
        If cell.Value <> "28" And cell.Value <> "16" And cell.Value <> "908" And cell.Value <> "26" And cell.Value <> "13" And cell.Value <> "69" And cell.Value <> "82" And cell.Value <> "901" And cell.Value <> "17" And cell.Value <> "87" And cell.Value <> "50" And cell.Value <> "23" And cell.Value <> "29" And cell.Value <> "77" Then
        MsgBox ("Error: Incorrect ETI_COUNTRY_ID  Value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next


'ETI_CURRENCY
For Each cell In Range("R2:R" & lastr)
        If cell.Value <> "SIT" And cell.Value <> "EUR" And cell.Value <> "KZT" And cell.Value <> "DKK" And cell.Value <> "BRL" And cell.Value <> "FIM" And cell.Value <> "NOK" And cell.Value <> "SEK" And cell.Value <> "HRK" And cell.Value <> "PLN" And cell.Value <> "TRY" And cell.Value <> "NZD" Then
            MsgBox ("Error: Incorrect ETI_CURRENCY  Value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next

'ETI_ETI_Cellular_Connectivity 0/1 CHECK
For Each cell In Range("S2:S" & lastr)
        If cell.Value <> "0" And cell.Value <> "1" Then
            MsgBox ("Error: Incorrect ETI_Cellular_Connectivity  Value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next

'Date Formatting Check
For Each cell In Range("B2:B" & lastr)
        RE.Pattern = "^\d\d\d\d\-\d\d\-\d\d$"
            If (Not RE.test(cell.Value)) And cell.Value <> "" Then
               
                MsgBox ("Incorrect Date formatting at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
                cell.Select
                Exit Sub
        End If
    Next
    
'Price Formatting Check
For Each cell In Range("J2:J" & lastr)
        RE.Pattern = "\..*\."
            If RE.test(cell.Value) Then
                
                MsgBox ("Incorrect Price value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
                cell.Select
                Exit Sub
        End If
        RE.Pattern = "[^\d\.]+"
        If RE.test(cell.Value) Then
            
            MsgBox ("Incorrect Price value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
            cell.Select
            Exit Sub
        End If
    Next

For Each cell In Range("J2:J" & lastr)
    If cell.Value = "0" Or cell.Value = "0.0" Or cell.Value = "0.00" Then
        MsgBox ("Price values cannot be 0 or 0.0 at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next


'MATCH RAM WITH ROM
For Each cell In Range("M2:M" & lastr)
  If cell.Value <> "-" And Not cell.Value Like "*MB*" Then
        If cell.Value = cell.Offset(0, 1).Value Or cell.Value = Cells(cell.Row, Range("N1").Column) Then
            MsgBox ("Error: Same DATA in multiple fields, at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
            cell.Select
            End
        End If
    End If
Next

'MATCH NAME WITH MPNS
For Each cell In Range("I2:I" & lastr)
        If cell.Value = cell.Offset(0, 9).Value Or cell.Value = Cells(cell.Row, Range("Q1").Column) Then
            MsgBox ("Error: Same DATA in multiple fields, at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
            cell.Select
            End
        End If
Next

'ETI_RAM ODD VALUES CHECK
For Each cell In Range("N2:N" & lastr)
        If cell.Value = "0GB" Or cell.Value = "64GB" Or cell.Value = "128GB" Or cell.Value = "256GB" Or cell.Value = "512GB" Or cell.Value = "1TB" Or cell.Value = "250GB" Or cell.Value = "0MB" Or cell.Value = "0TB" Then
            MsgBox ("ETI_RAM ODD VALUES at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
            cell.Select
            Exit Sub
        End If
Next

'ETI_Storage Capacity ODD VALUES CHECK
For Each cell In Range("M2:M" & lastr)
        If cell.Value = "0GB" Or cell.Value = "12GB" Or cell.Value = "6GB" Or cell.Value = "3GB" Or cell.Value = "265GB" Or cell.Value = "0MB" Or cell.Value = "0TB" Then
            MsgBox ("ETI_Storage Capacity ODD VALUES at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
            cell.Select
            Exit Sub
        End If
Next



'Country ID, Retailer ID
For Each cell In Range("D2:D" & lastr)
    RE.Pattern = "[^\d]+"
    If RE.test(cell.Value) Then
        
        MsgBox ("Error: ETI_COUNTRY_ID is not a number at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next
For Each cell In Range("H2:H" & lastr)
    RE.Pattern = "[^\d]+"
    If RE.test(cell.Value) Then
        
        MsgBox ("Error: ETI_RETAILER_ID is not a number at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next

'DPG ID
For Each cell In Range("E2:E" & lastr)
    If cell.Value = "32647" Or cell.Value = "321373" Then
    Else
       
        MsgBox ("Incorrect ETI_DPG_ID at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
        cell.Select
        Exit Sub
    End If
Next

'Inc Vat Check
For Each cell In Range("K2:K" & lastr)
        If cell.Value = "Yes" Or cell.Value = "No" Then
        Else
            
            MsgBox ("Incorrect ETI_INC_VAT value at " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
    Next

Dim regEx As Object
Set regEx = CreateObject("VBScript.RegExp")
IsIgnore = True
'ITEM NAME
For Each cell In Range("I2:I" & lastr)
    With regEx
        .IgnoreCase = IsIgnore
        .Global = True
        .Pattern = "(like\s*new)|(Renewd)|(Refurb)|(Reacondicionado)|(Reconditionné)|(Refurbished)|(Prepaid)|(B\-Ware)|(Reacondicionado)|(Pack)|(Demo)|(Locked)|(Bundle)|(case)|(cables)|\b(ebook)\b|(speaker)|(headphones)"
    End With
    If regEx.test(cell.Value) Then
        
        MsgBox ("Error: ETI_Item_Name at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next
'Check Numeric Values in Color
For Each cell In Range("O2:O" & lastr)
    With regEx
        .IgnoreCase = IsIgnore
        .Pattern = "\d{1,}"
    End With
    If regEx.test(cell.Value) Then
       
        MsgBox ("Error: Numeric Values in Color at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next
'CELLULAR CONECTIVITY
For Each cell In Range("S2:S" & lastr)
    With regEx
        .IgnoreCase = IsIgnore
        .Pattern = "(-)"
    End With
    If regEx.test(cell.Value) Then
       
        MsgBox ("Error: CELLULAR CONECTIVITY at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next
'BRANDS
For Each cell In Range("L2:L" & lastr)
    With regEx
        .IgnoreCase = IsIgnore
        .Pattern = "\["
    End With
    If regEx.test(cell.Value) Then
       
        MsgBox ("Error: ETI_Item_Name at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next
'COLORS
For Each cell In Range("O2:O" & lastr)
    With regEx
        .IgnoreCase = IsIgnore
        .Pattern = "(\{)|(\})"
    End With
    If regEx.test(cell.Value) Then
       
        MsgBox ("Error: ETI_Item_Name at " & cell.Address(0, 0) & ": " & """" & cell.Value & """")
        cell.Select
        Exit Sub
    End If
Next


'Match DPG ID WITH CELLULAR CONNECTIVITY for safe side already fill it
For Each cell In Range("E2:E" & lastr)
      If cell.Value = "32647" And cell.Offset(0, 14).Value = "0" Then
            MsgBox ("Error: Invalid Cellular Connectivity for SmartPhones " & cell.Address(0, 0) & ": " & """" & cell.Value & """" & vbNewLine & "Please Fix and Rerun Macro!")
            cell.Select
            Exit Sub
        End If
Next


ActiveWorkbook.Close SaveChanges:=True
End Sub




