'#########################################################################
'# PURPOSE: Convert a given number into it's corresponding Letter Reference
'# SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'#########################################################################

Private Function Number2Letter(ByVal columnNumber As Integer)
    Dim columnLetter As String

    'Convert To Column Letter
    columnLetter = Split(Cells(1, columnNumber).Address, "$")(1)

    Number2Letter = columnLetter            'set return value
End Function

'#########################################################################
'# PURPOSE: Convert a given letter into it's corresponding Numeric Reference
'# SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'#########################################################################
Private Function Letter2Number(ByVal columnLetter As String)
    Dim columnNumber As Long

    'Convert To Column Number
    columnNumber = Range(columnLetter & 1).Column

    Letter2Number = columnNumber
End Function
'********************************************************
'* reference: http://www.rondebruin.nl/win/s9/win005.htm
'********************************************************
Private Function LastRow(sh As Worksheet)
    On Error Resume Next
    LastRow = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    On Error GoTo 0
End Function
'********************************************************
'* reference: http://www.rondebruin.nl/win/s9/win005.htm
'********************************************************
Private Function LastCol(sh As Worksheet)
    On Error Resume Next
    LastCol = sh.Cells.Find(What:="*", _
                            After:=sh.Range("A1"), _
                            Lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

'#########################################################################
'# adds a border around the selection
'#########################################################################
Private Sub AddBorder()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
