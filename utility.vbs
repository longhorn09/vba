Option Explicit
Private Const Unix1970 As Long = 25569 'CDbl(DateSerial(1970, 1, 1))

'#########################################################################
'# random code examples
'#########################################################################
Private Sub RandomStuff()
    Dim regex, matches As Object
    Dim basePartNo As String
    
    Set regex = CreateObject("VBScript.RegExp") 'need regular expressions to deal with part numbers
    regex.MultiLine = False
    regex.IgnoreCase = True
    regex.Pattern = "^([A-Za-z0-9]+)\-([A-Za-z0-9]{3})\-([A-Za-z0-9]+)"
    
    If (regex.test(basePartNo)) Then
        Set matches = regex.Execute(basePartNo)
        basePartNo = matches(0).submatches(0) & matches(0).submatches(1) & " " & matches(0).submatches(2)
    End If
        
    Set matches = Nothing
    Set regex = Nothing
End Sub
'#########################################################################
'# File Dialog stuff
'#########################################################################
Public Sub PickFile()
    Dim wkdyMonth As String
    Dim fd As FileDialog
    Dim i As Integer
    Dim keepGoing As Boolean
    Dim filePath As String
    Dim wb As Workbook
    Dim currentSheetName As String
    Dim newStartRow As Integer
    'Const Path = "C:\Users\Desktop\"
   
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Filters.Clear
    fd.Filters.Add "Excel files", "*.xls*"
    fd.Filters.Add "CSV files", "*.csv"
    fd.Title = "Select the input file"
    'fd.InitialFileName = Path
   
   
    'Debug.Print Format(FormatDateTime(Now, vbShortDate), "mmmm")
    wkdyMonth = Application.InputBox("Enter the month", Default:=Format(FormatDateTime(Now, vbShortDate), "mmmm"), Type:=2, Title:="Enter Month")
   
    i = fd.Show
    If i <> -1 Then
        keepGoing = False
    Else
        keepGoing = True
    End If
   
    keepGoing = True
    If keepGoing Then
        Call DisableStuff
        filePath = fd.SelectedItems(1)
       
        Call DeleteExisting(ActiveSheet.Name, wkdyMonth)
        Set wb = Workbooks.Open(filePath)
        newStartRow = AppendToExisting(wb, ActiveSheet.Name)
        Call wb.Close
        Call RemakeFormulas(newStartRow, wkdyMonth)
        Call EnableStuff
        Call MsgBox(LastRow(ActiveSheet) - newStartRow & " rows in " & wkdyMonth, vbInformation)
    End If
   
    Set wb = Nothing
    Set fd = Nothing
End Sub

'#########################################################################
'# disable stuff
'#########################################################################
Private Sub DisableStuff()
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub
'#########################################################################
'# enable stuff
'#########################################################################
Private Sub EnableStuff()
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
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

'####################################################################################
'# http://www.vbforums.com/showthread.php?513727-RESOLVED-Convert-Unix-Time-to-Date
'####################################################################################
Private Function Date2Unix(ByVal vDate As Date) As Long
    Date2Unix = DateDiff("s", Unix1970, vDate)
End Function

Private Function Unix2Date(ByVal vUnixDate As Long) As Date
    Unix2Date = DateAdd("s", vUnixDate, Unix1970)
End Function
Private Function UnixTimeToDate(ByVal Timestamp As Long) As Date
    Dim intDays As Integer, intHours As Integer, intMins As Integer, intSecs As Integer
 
    intDays = Timestamp \ 86400
    intHours = (Timestamp Mod 86400) \ 3600
    intMins = (Timestamp Mod 3600) \ 60
    intSecs = Timestamp Mod 60
    
    UnixTimeToDate = DateSerial(1970, 1, intDays + 1) + TimeSerial(intHours, intMins, intSecs)
End Function
'#########################################################################
'# reference: https://www.exceldemy.com/excel-vba-paste-range-into-email-body/
'#########################################################################
Private Function RangetoHTML(rng As Range)
    Dim obj As Object
    Dim txtstr As Object
    Dim File As String
    Dim WB As Workbook
    File = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    rng.Copy
    Set WB = Workbooks.Add(1)
    With WB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With
    With WB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=File, _
         Sheet:=WB.Sheets(1).Name, _
         Source:=WB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
    Set obj = CreateObject("Scripting.FileSystemObject")
    Set txtstr = obj.GetFile(File).OpenAsTextStream(1, -2)
    RangetoHTML = txtstr.readall
    txtstr.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
    WB.Close savechanges:=False
    Kill File
    Set txtstr = Nothing
    Set obj = Nothing
    Set WB = Nothing
End Function    
        
'#########################################################################
'# https://www.extendoffice.com/documents/excel/5049-pasting-an-excel-range-into-an-email-as-a-picture.html
'#########################################################################
Private Sub createJpg(SheetName As String, xRgAddrss As String, nameFile As String)
    Dim xRgPic As Range
    Dim xShape As Shape
    
    ThisWorkbook.Activate
    Worksheets(SheetName).Activate
    Set xRgPic = ThisWorkbook.Worksheets(SheetName).Range(xRgAddrss)
    xRgPic.CopyPicture
    
    With ThisWorkbook.Worksheets(SheetName).ChartObjects.Add(xRgPic.Left, xRgPic.Top, xRgPic.Width, xRgPic.Height)
        .Activate
        For Each xShape In ActiveSheet.Shapes
            xShape.Line.Visible = msoFalse
        Next
        .Chart.Paste
        .Chart.Export Environ$("temp") & "\" & nameFile & ".jpg", "JPG"
    End With
    
    Worksheets(SheetName).ChartObjects(Worksheets(SheetName).ChartObjects.Count).Delete
    
    Set xRgPic = Nothing
End Sub        
