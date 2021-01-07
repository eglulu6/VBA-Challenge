Sub pseduo_code():

'add
'write in a table:
    'ticker Symbol
    'the diff of opening/closing per yr
        'find the opening price at the begining of yr(date min)
        'find the closing price at the end of yr (date max)
    'the % of that diff (open/closer per yr)
    'total stock volume
'loop to find?
'reset repeating values

'-----------------
'   TUTOR NOTES
'-----------------

'How to get the Tikr symbol to match min.max?
'change button to work on first click?
'get dropdown values to = ws names dynamically?

'***WorksheetFunction.Index(Range("result column"),WorksheetFunction.Match(

End Sub

Sub sumrysht():

'insert summary sheet & table
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Summary Sheet"

'insert table headers
    Range("A6") = "Year"
    Range("B6") = "Tiker Symbol"
    Range("C6") = "Total Stock Volume"
    Range("D6") = "Total Yr Change"
    Range("E6") = "Total Yr %"
    Range("A1") = "Choose YR Summary"
    Range("B1") = "ALL"
    Range("D1") = "Ticker"
    Range("E1") = "Value"
    Range("C2") = "Greatest % Increase"
    Range("C3") = "Greatest % Decrease"
    Range("C4") = "Greatest Total Volume"
    Range("A3") = "*Click 2x after filter change to calcualte values"
    
'format table
ActiveWorkbook.Sheets("Summary Sheet").Range("A1:B1").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
ActiveWorkbook.Sheets("Summary Sheet").Range("C1:E4").BorderAround LineStyle:=xlContinuous, Weight:=xlThick
ActiveWorkbook.Sheets("Summary Sheet").Range("A6:E6").Borders(xlEdgeBottom).LineStyle = XlLineStyle.xlContinuous
ActiveWorkbook.Sheets("Summary Sheet").Range("A1, C2:C4, D1:E1, A6:E6").Font.FontStyle = "Bold"
    
'variables
    Dim lstrow As Long
    Dim tkr As String
    Dim vol As Double
    Dim yr_chg As Double
    Dim prct_chg As Variant
    Dim opn As Double
    Dim cls As Double
    Dim tbl As Range
    Dim tbl_row As Integer
    
'assigned variables
    tbl_row = 7

'loop thru sheets except table sheet
For Each ws In ActiveWorkbook.Worksheets

'assigned variables
    lstrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
'all sheets except table sheet
    If ws.Name <> "Summary Sheet" Then
    
    'set first open value on sheet
    opn = ws.Cells(2, 3).Value
        
            'loop sheet contents
            For r = 7 To lstrow
            
            'if current tiker <> nex tkr
            If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
                
                'prevent overflow or Div/0 error
                If opn <> 0 Then

                    'set tkr symbol
                    tkr = ws.Cells(r, 1).Value
                    
                    'Add to ttl tkr volume
                    vol = vol + ws.Cells(r, 7).Value
                    
                    'set tkr cls & yr change values
                    cls = ws.Cells(r, 6).Value
                    yr_chg = opn - cls
                    prct_chg = yr_chg / opn
                    
                    'Print values
                    ActiveWorkbook.Sheets("Summary Sheet").Range("A" & tbl_row) = ws.Name
                    ActiveWorkbook.Sheets("Summary Sheet").Range("B" & tbl_row) = tkr
                    ActiveWorkbook.Sheets("Summary Sheet").Range("C" & tbl_row) = vol
                    ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row) = yr_chg
                    ActiveWorkbook.Sheets("Summary Sheet").Range("E" & tbl_row) = prct_chg
               
               'preventative else
               Else
                    yr_chg = 0
                    prct_chg = 0
                    
                    'Print values
                    ActiveWorkbook.Sheets("Summary Sheet").Range("A" & tbl_row) = ws.Name
                    ActiveWorkbook.Sheets("Summary Sheet").Range("B" & tbl_row) = tkr
                    ActiveWorkbook.Sheets("Summary Sheet").Range("C" & tbl_row) = vol
                    ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row) = yr_chg
                    ActiveWorkbook.Sheets("Summary Sheet").Range("E" & tbl_row) = prct_chg

                'end of preventative if
                End If
                
                'format cells
                     If yr_chg >= 0 Then
                         ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row).Interior.ColorIndex = 4
                         ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row).NumberFormat = "$0.00"
                         ActiveWorkbook.Sheets("Summary Sheet").Range("C" & tbl_row).NumberFormat = "#,##0"
                         ActiveWorkbook.Sheets("Summary Sheet").Range("E" & tbl_row).NumberFormat = "0.00%"
                    'formatting else
                     Else
                         ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row).Interior.ColorIndex = 3
                         ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row).Font.ColorIndex = 2
                         ActiveWorkbook.Sheets("Summary Sheet").Range("D" & tbl_row).NumberFormat = "$0.00"
                         ActiveWorkbook.Sheets("Summary Sheet").Range("C" & tbl_row).NumberFormat = "#,##0"
                         ActiveWorkbook.Sheets("Summary Sheet").Range("E" & tbl_row).NumberFormat = "0.00%"


                'end formatting if statment
                End If
                
                'Add one to the summary table row
                tbl_row = tbl_row + 1
                
                'Reset Open value for next tkr
                opn = ws.Cells(r + 1, 4).Value
                
                'Reset the Vol Total
                vol = 0
                
                'Reset the Tkr Symbol
                tkr = ""
            
                'add to vol
                vol = vol + ws.Cells(r, 7).Value
                
            'end if that prints values
            End If
                        
        Next r
        
    'end sheets if
    End If

Next ws

'Format other cells
ActiveWorkbook.Sheets("Summary Sheet").Range("A3:B4").Merge
ActiveWorkbook.Sheets("Summary Sheet").Range("A:A, A6:E6, D1:E1").HorizontalAlignment = xlCenter
ActiveWorkbook.Sheets("Summary Sheet").Range("A6:E6, D1:E1").Font.Bold = True
ActiveWorkbook.Sheets("Summary Sheet").Range("A3").WrapText = True
ActiveWorkbook.Sheets("Summary Sheet").Cells.EntireColumn.AutoFit

'insert validation list for YR choices
With ActiveWorkbook.Sheets("Summary Sheet").Range("B1").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Formula1:="2016, 2015, 2014, ALL"
End With
    
'add button to run filter sub filteryrbttn Macro
ActiveWorkbook.Sheets("Summary Sheet").Buttons.Delete

    Dim yr As Range
    Dim btn As Button
    Set yr = ActiveWorkbook.Sheets("Summary Sheet").Range("A2:B2")
    Set btn = ActiveWorkbook.Sheets("Summary Sheet").Buttons.Add(yr.Left, yr.Top, yr.Width, yr.Height)

With btn
    .Caption = "Get Values"
    .Name = "Get Values"
    ActiveWorkbook.Sheets("Summary Sheet").Shapes.Range(Array("Get Values")).Select
    Selection.OnAction = "btn_actions"
End With


End Sub

Sub btn_actions():

'variables
Dim yr As String
Dim lstrow As Long
Dim prct_max As Double
Dim prct_min As Double
Dim vol_max As Double
Dim prct_col As Range
Dim vol_col As Range
Dim max_mtch As Double
Dim min_mtch As Double
Dim vol_mtch As Double

'assign variables
yr = ActiveWorkbook.Sheets("Summary Sheet").Range("B1").Value
lstrow = ActiveWorkbook.Sheets("Summary Sheet").Cells(Rows.Count, 1).End(xlUp).Row
Set prct_col = ActiveWorkbook.Sheets("Summary Sheet").Range("E7:E" & lstrow).SpecialCells(xlCellTypeVisible)
Set vol_col = ActiveWorkbook.Sheets("Summary Sheet").Range("C7:C" & lstrow).SpecialCells(xlCellTypeVisible)
prct_min = Application.WorksheetFunction.Min(prct_col)
prct_max = Application.WorksheetFunction.Max(prct_col)
vol_max = Application.WorksheetFunction.Max(vol_col)
max_mtch = Application.Match(prct_max, prct_col, 0)
min_mtch = Application.Match(prct_min, prct_col, 0)
vol_mtch = Application.Match(vol_max, vol_col, 0)

'filter w/ button
If yr <> "ALL" Then
    'filter data by yr
    ActiveWorkbook.Sheets("Summary Sheet").Range("A6:E" & lstrow).AutoFilter Field:=1, Criteria1:=yr
    ActiveWorkbook.Sheets("Summary Sheet").Range("E2") = prct_max
    'ActiveWorkbook.Sheets("Summary Sheet").Range("D2") = ActiveWorkbook.Sheets("Summary Sheet").Cells(max_mtch, 2)
    ActiveWorkbook.Sheets("Summary Sheet").Range("E3") = prct_min
    'ActiveWorkbook.Sheets("Summary Sheet").Range("D3") = ActiveWorkbook.Sheets("Summary Sheet").Cells(min_mtch, 2)
    ActiveWorkbook.Sheets("Summary Sheet").Range("E4") = vol_max
    'ActiveWorkbook.Sheets("Summary Sheet").Range("D4") = ActiveWorkbook.Sheets("Summary Sheet").Cells(vol_mtch, 2)
Else
 'Show all data
    ActiveWorkbook.Sheets("Summary Sheet").AutoFilter.ShowAllData
    ActiveWorkbook.Sheets("Summary Sheet").Range("E2") = prct_max
    'ActiveWorkbook.Sheets("Summary Sheet").Range("D2") = ActiveWorkbook.Sheets("Summary Sheet").Cells(max_mtch, 2)
    ActiveWorkbook.Sheets("Summary Sheet").Range("E3") = prct_min
    'ActiveWorkbook.Sheets("Summary Sheet").Range("D3") = ActiveWorkbook.Sheets("Summary Sheet").Cells(min_mtch, 2)
    ActiveWorkbook.Sheets("Summary Sheet").Range("E4") = vol_max
    'ActiveWorkbook.Sheets("Summary Sheet").Range("D4") = ActiveWorkbook.Sheets("Summary Sheet").Cells(vol_mtch, 2)
End If

'Format results
ActiveWorkbook.Sheets("Summary Sheet").Range("E2:E3").NumberFormat = "0.00%"
ActiveWorkbook.Sheets("Summary Sheet").Range("E4").NumberFormat = "#,##0"

End Sub
