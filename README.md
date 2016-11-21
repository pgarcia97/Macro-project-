# Macro-project-
Automate process of grabbing data from ticketing system and formatting it

Sub VAMSISorganizing()
'
' VAMSISorganizing Macro
'
 
'
    'clear IM and RM worksheets
   
    Call BGNING
 
  
    Workbooks.Open Filename:= _
        "https://mxteams.massmutual.com/sites/div5/ISS/Metrics/Organizational%20Metrics%20v5.5.xlsm"
'    Windows("Org metrics tool.xlsm").Activate
    Sheets("Master").Select
   
   
    Windows("Organizational Metrics v5.5.xlsm").Activate
    Range("C2:E2") = "AOS"
    Application.Wait (Now + #12:00:01 AM#)
      Range("C4:E4") = "Vamsi Chavali Nageshwara"
   Application.Wait (Now + #12:00:02 AM#)
   Range("C6:E6") = "ESAS"
   Application.Wait (Now + #12:00:02 AM#)
     Windows("Organizational Metrics v5.5.xlsm").Activate
    ' Range("C8:E8") = "Lisa Moriarty"
  
  '
   ActiveSheet.OLEObjects(2).Select
    ActiveSheet.OLEObjects(2).Object.Value = True
    'Stop
  
 
    Windows("Organizational Metrics v5.5.xlsm").Activate
    Sheets("Queue").Select
    ActiveWorkbook.Worksheets("Queue").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Queue").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "A1:A4258"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Queue").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
    Dim OpenDate As Date
    Dim TimeToResolve As Long
    Dim TimeStillOpen As Long
    Dim intRow As Integer
    Dim intRowRem1 As Integer
    Dim intRowRem2 As Integer
    Dim intRow1 As Integer
    Dim intRow2 As Integer
   ' Dim Now As Date
    
    intRow = 2
    intRow1 = 1
    intRow2 = 2
    Do Until Range("A" & intRow).Value = ""
      
       If Left(Range("A" & intRow).Value, 1) = "I" Then
          Exit Do
       End If
       If Left(Range("A" & intRow).Value, 1) = "C" Then
          Rows(intRow & ":" & intRow).Select
         
          Selection.Delete Shift:=xlUp
       Else
         
          intRow = intRow + 1
       End If
 
    Loop
      
    intRow = 2
   
    Do Until Range("A" & intRow).Value = ""
     If Left(Range("A" & intRow).Value, 1) = "R" Then
          Exit Do
       End If
      If Range("K" & intRow).Value = "" Then
           TimeStillOpen = DateDiff("h", Range("D" & intRow).Value, Now)
           Range("U" & intRow) = TimeStillOpen
       Else
        'TimeToResolve = Range("K" & intRow).Value - Range("D" & intRow).Value
        TimeToResolve = DateDiff("h", Range("D" & intRow).Value, Range("K" & intRow).Value)
 
        Range("T" & intRow) = TimeToResolve
        End If
         intRowRem1 = intRow
        intRow = intRow + 1
       
        Loop
    Windows("Organizational Metrics v5.5.xlsm").Activate
    ActiveWindow.SmallScroll Down:=-93
   ' Rows("1:intRowRem").Select
    Rows(intRow1 & ":" & intRow - 1).Select
    Range("C1").Activate
    Selection.Copy
    Windows("VAMSIorganized version 2.xlsm").Activate
    Sheets("IM").Select
    Range("A1").Select
    ActiveSheet.Paste
   
    'add lines (boxes)
   
  
    Range("A1", ("Y" & intRow - 1)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
   
    'add colors
      
    Range("V2", ("W" & intRow - 1)).Select
         With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
      Range("X2", ("Y" & intRow - 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
   
    
    
    Windows("Organizational Metrics v5.5.xlsm").Activate
    Rows(intRow2 & ":" & intRow - 1).Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
    
        intRow = 2
       
        Do Until Range("A" & intRow).Value = ""
     If Left(Range("A" & intRow).Value, 1) = "" Then
          Exit Do
       End If
      If Range("K" & intRow).Value = "" Then
      TimeStillOpen = DateDiff("h", Range("D" & intRow).Value, Now)
           Range("U" & intRow) = TimeStillOpen
       Else
        TimeToResolve = DateDiff("h", Range("D" & intRow).Value, Range("K" & intRow).Value)
        Range("T" & intRow) = TimeToResolve
        End If
         intRowRem2 = intRow
        intRow = intRow + 1
       
        Loop
    Windows("Organizational Metrics v5.5.xlsm").Activate
    Rows(intRow1 & ":" & intRow).Select
    Range("C1").Activate
    Selection.Copy
    Windows("VAMSIorganized version 2.xlsm").Activate
    Sheets("RM").Select
    Range("A1").Select
    ActiveSheet.Paste
    Windows("Organizational Metrics v5.5.xlsm").Activate
    ActiveWindow.Close
  
   ''' add lines (boxes)
  
   
    Range("A1", ("Y" & intRow - 1)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
  
   ''' end add lines
  
     'add colors
      
    Range("V2", ("W" & intRow - 1)).Select
         With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
      Range("X2", ("Y" & intRow - 1)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
   
   
   
   
   
   
   Windows("VAMSIorganized version 2.xlsm").Activate
   Sheets("IM").Select
    Cells.Select
    Range("A4").Activate
    ActiveWorkbook.Worksheets("IM").Sort.SortFields.Clear
 
     ActiveWorkbook.Worksheets("IM").Sort.SortFields.Add Key:=Range("B2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
   
    ActiveWorkbook.Worksheets("IM").Sort.SortFields.Add Key:=Range("K2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("IM").Sort
        .SetRange Range("A1:AE1992")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
    
    Windows("VAMSIorganized version 2.xlsm").Activate
   Sheets("RM").Select
    Cells.Select
    Range("A4").Activate
    ActiveWorkbook.Worksheets("RM").Sort.SortFields.Clear
 
     ActiveWorkbook.Worksheets("RM").Sort.SortFields.Add Key:=Range("B2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
   
    ActiveWorkbook.Worksheets("RM").Sort.SortFields.Add Key:=Range("K2") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("RM").Sort
        .SetRange Range("A1:AE1992")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   
    
    'FORMATTING IM
   
    Sheets("IM").Select
    ActiveWindow.SmallScroll Down:=-30
    Cells.Select
    Cells.EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    Range("S1").Select
    Selection.Copy
    Range("T1:U1").Select
    ActiveSheet.Paste
    Range("T1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Time to Resolve (hours)"
    Columns("T:T").Select
    Selection.ColumnWidth = 5.29
    Selection.ColumnWidth = 6.43
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Time Open (hours)"
    Columns("U:U").Select
    Selection.ColumnWidth = 6.57
    Range("T1:U1").Select
    Selection.Copy
    Range("V1").Select
    ActiveSheet.Paste
    Range("V1").Select
    Application.CutCopyMode = False
   ''' just adding
   ' ActiveCell.FormulaR1C1 = "Time to Resolve"
   ' Range("V1").Select
   ' ActiveCell.FormulaR1C1 = "Total Time "
   ' Range("W1").Select
   ' ActiveCell.FormulaR1C1 = "Total # Tickets"
   ' Range("V1").Select
   
     ActiveCell.FormulaR1C1 = "Time to Resolve"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "Resolved Total Time "
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Resolved Total # Tickets"
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "Open Total Time "
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "Open Total # Tickets"
   
    
    
    'Formatting RM
   
    
   Sheets("RM").Select
    ActiveWindow.SmallScroll Down:=-30
    Cells.Select
    Cells.EntireColumn.AutoFit
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    Range("S1").Select
    Selection.Copy
    Range("T1:U1").Select
    ActiveSheet.Paste
    Range("T1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Time to Resolve (hours)"
    Columns("T:T").Select
    Selection.ColumnWidth = 5.29
    Selection.ColumnWidth = 6.43
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Time Open (hours)"
    Columns("U:U").Select
    Selection.ColumnWidth = 6.57
    Range("T1:U1").Select
    Selection.Copy
    Range("V1").Select
    ActiveSheet.Paste
    Range("V1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Time to Resolve"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "Resolved Total # Tickets "
    Range("W1").Select
    ActiveCell.FormulaR1C1 = "Resolved Total # Tickets"
    Range("X1").Select
    ActiveCell.FormulaR1C1 = "Open Total Time "
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "Open Total # Tickets"
  ' Range("V1").Select
 
  Call groupandtally
 
   Call groupandtallyIM
  
   Call Asstmisc
  
   Call reorganize
  
   '
 
 
   
   
   
  ' Sheets("Start").Select
  
  ' ActiveWorkbook.SaveAs Filename:= _
  '      "\\Mmfile\mm84854$\MyDocuments\VAMSIorganized version 2.xlsm", FileFormat:= _
  '      xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
  
  
End Sub
