Attribute VB_Name = "Module3"
Const G_COLORINDEX_OFF As Long = &HEAEAEA
Const G_COLORINDEX_DEFAULT As Long = &HFFFFFF
Const G_COLORINDEX_PRJ As Long = &H99FFFF
Const G_COLORINDEX_DESIGN As Long = &HB4D5FC
Const G_COLORINDEX_DEV As Long = &HE4CCB8
Const G_COLORINDEX_TEST As Long = &H50D092
Const G_COLORINDEX_INDUS As Long = &H6464FF
Const G_COLORINDEX_JALON As Long = &H0
Const G_HOLIDAY_NAME As String = "MKPLAN_Holidays"

' From http://forums.devx.com/showthread.php?169794-Excel-VBA-Verify-A-Named-Range-Exists
' Return False if named range does not exist
Function NameExists(TheName As String) As Boolean
On Error Resume Next
NameExists = Len(ThisWorkbook.Names(TheName).Name) <> 0
End Function



' From http://support.microsoft.com/kb/833402
' Concert column number to Excel representation of column
Function ctol2(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ctol = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ctol = ctol & Chr(iRemainder + 64)
   End If
End Function

Function ctol(iCol As Integer) As String
   Dim iAlpha As Integer
   
   iAlpha = iCol
   
   Do
     iAlpha = iAlpha - 1
     ctol = Chr((iAlpha Mod 26) + 65) & ctol
     iAlpha = (iAlpha \ 26)
   Loop While (iAlpha > 0)
   
End Function

Sub test_ctol()
    Debug.Print "1=A : " & ctol2(1)
    Debug.Print "26=Z : " & ctol2(26)
    Debug.Print "27=AA : " & ctol2(27)
    Debug.Print "157=FA : " & ctol2(157)
    Debug.Print "702=ZZ : " & ctol2(702)
    Debug.Print "703=AAA : " & ctol2(703)
    Debug.Print "833=AFA : " & ctol2(833)
End Sub


Sub run_sumActivity()

    Call sumActivity(Selection)
    
End Sub

Sub sumActivity( _
    ByRef rRangeToSum As Range)
    Dim rToSum As Range
    Dim iRow As Long
    Dim iRowNb As Long
    Dim sSum As String
    Dim itemCell As Range


    Set rSumLine = rRangeToSum.Resize(1).Offset(-1)
    
    iRowNb = rRangeToSum.Rows.Count
    iRow = rRangeToSum.Row
    
    For Each itemCell In rSumLine
    
        If (itemCell.Interior.Color <> G_COLORINDEX_OFF) Then
            Dim rs As Range
            Dim lSumColor As Long
            
            sSum = "=SUM(" & ctol(itemCell.Column) & iRow & ":" & ctol(itemCell.Column) & iRow + iRowNb - 1 & ")"
            
            Set rs = itemCell.Worksheet.Range(itemCell.Cells(2, 1), itemCell.Cells(iRowNb + 1, 1))
            
            lSumColor = G_COLORINDEX_DEFAULT
            
            For Each ii In rs
                'Debug.Print "Val col/row : " & ii.Column & " " & ii.Row
                If (ii.Interior.Color <> G_COLORINDEX_DEFAULT) Then
                   lSumColor = ii.Interior.Color
                   'Debug.Print "Color : " & ii.Interior.Color & " " & lDefaultColor
                End If
            Next ii
            
            Call applyColorRange(itemCell, lSumColor)
            
            itemCell.FormulaArray = sSum

        End If
        
    Next itemCell

End Sub


Sub run_MkPlan()
    Dim rgDate As Range
    
    Set rgDate = Selection.Resize(1, 1).Offset(0, -1)
    
    'rgDate.Interior.ColorIndex = 3
    Debug.Print "Val col/row : " & rgDate.Column & "/" & rgDate.Row
    
    Call MkPlan(Selection.Worksheet, Range("MKPLAN_Holidays"), Selection.Columns.Count - 1, Selection.Rows.Count, Selection.Column, rgDate.value)
    'Call MkPlan(Worksheets("Feuil1"), Range("MKPLAN_Holidays"), 100, 50, 3, DateSerial(2013, 11, 18))
End Sub


' This function build (layout) a planning, parameter :
' wSheet : sheet where planning will be layout
' rHolydays name of range containing Holydays
' lNbDaysCol : number of days in planning (also number of column)
' lNbTasksRow : number of row in planning (i.e. : number of task)
' lFirstDayCol : Offset of the first column of the planning (we suppose date are on first line)
' dFirstDay : first day of planning
Sub MkPlan( _
    ByRef wSheet As Worksheet, _
    ByRef rHolidays As Range, _
    ByVal lNbDaysCol As Long, _
    ByVal lNbTasksRow As Long, _
    ByVal lFirstDayCol As Long, _
    ByVal dFirstDay As Date _
    )

    Dim rng As Range
    Dim currDate As Date
    
    currDate = dFirstDay
    
    Set rng = wSheet.Range(Cells(1, lFirstDayCol), Cells(1, lFirstDayCol + lNbDaysCol))
    
    For Each it In rng
        
        currDate = DateAdd("d", 1, currDate)

        wd = offDay(currDate, Range(G_HOLIDAY_NAME))
               
        If (wd = 1 Or wd > 6) Then
            Call applyColorRange(Range(it, it.Cells(lNbTasksRow, 1)), G_COLORINDEX_OFF)
        End If
        
        it.value = currDate
        it.Orientation = 90
   Next

End Sub

Function offDay(d As Date, offDays As Range)
    offDay = Weekday(d)
    
    For Each iDay In offDays
    
      If (d = iDay.value) Then
        offDay = 8
      End If
    
    Next
End Function


Sub test_applyColorRange()
    Call applyColorRange(Selection, G_COLORINDEX_DEV)
End Sub

' Apply color in a range (or cell) except in OFF color
Private Sub applyColorRange(ByRef rToClear As Range, colNew As Long)

    For Each iCell In rToClear
        If (colNew = G_COLORINDEX_OFF) Then
            ' Reset to standard OFF color formatting ... to be sure
            With iCell.Interior
                .Color = G_COLORINDEX_OFF
                .Pattern = xlSolid
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            iCell.value = ""
        Else
            If (iCell.Interior.Color <> G_COLORINDEX_OFF) Then
                If (colNew = G_COLORINDEX_DEFAULT) Then
                    With iCell.Interior
                        .Color = colNew
                        .Pattern = xlNone
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    iCell.value = ""
                Else
                    With iCell.Interior
                        .Color = colNew
                        .Pattern = xlSolid
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    iCell.value = 1
                End If
            End If
        End If
    Next
 
End Sub

Sub colorOff()
    Call applyColorRange(Selection, G_COLORINDEX_OFF)
End Sub

Sub colorOnDesign()
    Call applyColorRange(Selection, G_COLORINDEX_DESIGN)
End Sub

Sub colorOnDev()
    Call applyColorRange(Selection, G_COLORINDEX_DEV)
End Sub

Sub colorOnTest()
    Call applyColorRange(Selection, G_COLORINDEX_TEST)
End Sub

Sub colorOnPrj()
    Call applyColorRange(Selection, G_COLORINDEX_PRJ)
End Sub

Sub colorOnClear()
    Call applyColorRange(Selection, G_COLORINDEX_DEFAULT)
End Sub

Sub colorOnIndus()
    Call applyColorRange(Selection, G_COLORINDEX_INDUS)
End Sub

Sub colorOnJalon()
    Call applyColorRange(Selection, G_COLORINDEX_JALON)
End Sub
