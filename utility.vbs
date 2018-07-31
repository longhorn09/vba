'#########################################################################
'# random code examples
'#########################################################################
private sub RandomStuff
    Dim regex, matches as Object
    dim basePartNo as string
    
    Set regex = CreateObject("VBScript.RegExp") 'need regular expressions to deal with part numbers
    regex.MultiLine = False
    regex.IgnoreCase = True
    regex.Pattern = "^([A-Za-z0-9]+)\-([A-Za-z0-9]{3})\-([A-Za-z0-9]+)"
    
    If (regex.test(basePartNo)) Then
        Set matches = regex.Execute(basePartNo)
        basePartNo = matches(0).submatches(0) & matches(0).submatches(1) & " " & matches(0).submatches(2)
    end if
        
    set matches = nothing
    set regex = nothing
end sub

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
'# PURPOSE: Check if the worksheet exists
'# SOURCE:  https://stackoverflow.com/questions/6688131/test-or-check-if-sheet-exists
'#########################################################################
Private Function SheetExists(ByVal shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
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
