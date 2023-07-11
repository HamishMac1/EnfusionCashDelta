Attribute VB_Name = "Main"
Option Explicit
Sub Check_To_Run_Process()
    If Range("PreviousColumnEnd").Value <> Range("CurrentColumnEnd").Value Then
        MsgBox ("Please Check Data - Column Mis-match")
    Else
        Call Delta_Creation_Process
    End If
    
    MsgBox "Delta Process Complete"
End Sub
Sub Delta_Creation_Process()
' Delta_Creation_Process Macro
' Written By S.Gladstone 17/12/2017 @ 16:44

Dim DataRowCountCurrent As Integer
Dim DataRowCountPrevious As Integer
Dim DataRowCountDelta As Integer
Dim ChangeCounter As Integer
Dim ColumnOffsetNative As Integer
Dim ColumnOffsetBase As Integer
Dim sCalc As String
Dim dt As Date
Dim rng As Range

ActiveSheet.Calculate
Application.DisplayAlerts = False
Application.ScreenUpdating = False
sCalc = Application.Calculation
Application.Calculation = xlCalculationAutomatic

'Clear Counters
    DataRowCountCurrent = 2
    DataRowCountPrevious = 2
    DataRowCountDelta = 2
    ChangeCounter = 0
    ColumnOffsetNative = 0
    ColumnOffsetBase = 0

'Delete Delta Data Sheet If It Exists
    On Error Resume Next
        ThisWorkbook.Sheets("Delta Data").Delete
    On Error GoTo 0

'Clean Up SpecialCells(xlLastCells)
    'Call Fcns.Last_Cell_CleanUp_Call 'Removed So As Not To Delete Buttons

'Add in A New Delta Data Sheet
    ActiveWorkbook.Sheets.Add After:=Worksheets("Control")
    ActiveSheet.Name = "Delta Data"

'Clear Old Formulae Out Of The Current Data Page
    Sheets("Current Data").Select
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    Else
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    End If
    Range("A2").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    DataRowCountCurrent = ActiveCell.Row
    Range("A2:B" & DataRowCountCurrent).Select
    Selection.ClearContents
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Rows(ActiveCell.Row & ":" & WorksheetFunction.Max(ActiveCell.Row, DataRowCountCurrent)).Select
    Selection.EntireRow.Delete

'Clear Old Formulae Out Of The Previous Data Page
    Sheets("Previous Data").Select
        If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    Else
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    End If
    Range("A2").Select
    ActiveCell.SpecialCells(xlLastCell).Select
    DataRowCountPrevious = ActiveCell.Row
    Range("A2:B" & DataRowCountPrevious).Select
    Selection.ClearContents
    Application.CutCopyMode = False
    Range("C1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    Rows(ActiveCell.Row & ":" & WorksheetFunction.Max(ActiveCell.Row, DataRowCountPrevious)).Select
    Selection.EntireRow.Delete

'Enter New Formulae for Lookup Key On Current Data Page
    Sheets("Current Data").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = Range("CurrentMatchFormulae").Value

'Enter New Formulae for Lookup Key On Previous Data Page
    Sheets("Previous Data").Select
    Range("B2").Select
    ActiveCell.FormulaR1C1 = Range("CurrentMatchFormulae").Value

'Enter New Formulae for Match Type On Current Data Page
    Sheets("Current Data").Select
    Range("A2").Select
    Range("A2").Formula = "=IF(ISERROR(VLOOKUP(B2,'Previous Data'!$B$2:$B$" & DataRowCountPrevious & ",1,FALSE))=FALSE," & Chr(34) & "Match" & Chr(34) & ",IF(ISERROR(VLOOKUP(C2,'Previous Data'!$C$2:$C$" & DataRowCountPrevious & ",1,FALSE))=FALSE," & Chr(34) & "Replace" & Chr(34) & "," & Chr(34) & "New" & Chr(34) & "))"
    Range("A2:B2").Select
    DataRowCountCurrent = Range("C2").End(xlDown).Row
    Selection.AutoFill Destination:=Range("A2:B" & DataRowCountCurrent)

'Enter New Formulae for Match Type On Previous Data Page
    Sheets("Previous Data").Select
    Range("A2").Select
    Range("A2").Formula = "=IF(ISERROR(VLOOKUP(B2,'Current Data'!$B$2:$B$" & DataRowCountCurrent & ",1,FALSE))=FALSE," & Chr(34) & "Match" & Chr(34) & ",IF(ISERROR(VLOOKUP(C2,'Current Data'!$C$2:$C$" & DataRowCountCurrent & ",1,FALSE))=FALSE," & Chr(34) & "Reverse" & Chr(34) & "," & Chr(34) & "Deleted" & Chr(34) & "))"
    Range("A2:B2").Select
    DataRowCountPrevious = Range("C2").End(xlDown).Row
    Selection.AutoFill Destination:=Range("A2:B" & DataRowCountPrevious)

'Check To See If Any Rows To Copy Onto Delta Page
    If Range("UnmatchedTotal").Value = 0 Then
        MsgBox ("No Delta File As Data Is The Same")
    Else
        
        If Range("UnmatchedCurrent").Value <> 0 Then
        'Select & Copy Required Rows From Current Data Page
            Sheets("Current Data").Select
            ActiveSheet.Range("$A$1:$" & Range("CurrentColumnEnd").Value & "$" & DataRowCountCurrent).AutoFilter Field:=1, Criteria1:="=New", _
                Operator:=xlOr, Criteria2:="=Replace"
            Rows("1:1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Delta Data").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            ActiveCell.SpecialCells(xlLastCell).Select
            Cells(ActiveCell.Row + 1, 1).Select
        End If

        If Range("UnmatchedPrevious").Value <> 0 Then
        'Select & Copy Required Rows From Previous Data Page
            Sheets("Previous Data").Select
            ActiveSheet.Range("$A$1:$" & Range("PreviousColumnEnd").Value & "$" & DataRowCountPrevious).AutoFilter Field:=1, Criteria1:="=Deleted", _
                Operator:=xlOr, Criteria2:="=Reverse"
            Rows("1:1").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.Copy
            Sheets("Delta Data").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            ActiveCell.Offset(1, 0).Select
            ActiveCell.Offset(-1, 0).Select
            Application.CutCopyMode = False
            If Range("UnmatchedCurrent").Value <> 0 Then
                Selection.EntireRow.Delete
            End If
        End If

        'Reverse Sign For Altered Values On Delta Sheet
        Range("A1").Select
        ActiveCell.SpecialCells(xlLastCell).Select
        DataRowCountDelta = ActiveCell.Row
        Range("ChangedDataRows").Formula = "=COUNTIF('Delta Data'!A1:A" & DataRowCountDelta & "," & Chr(34) & "Reverse" & Chr(34) & ") + COUNTIF('Delta Data'!A1:A" & DataRowCountDelta & "," & Chr(34) & "Deleted" & Chr(34) & ")"
        If Range("ChangedDataRows").Value > 0 Then
            Range("A1").Select
            Do While ActiveCell.Value <> "Native Amount"
                ActiveCell.Offset(0, 1).Select
                ColumnOffsetNative = ColumnOffsetNative + 1
            Loop
            Range("A1").Select
            Do While ActiveCell.Value <> "Base Amount"
                ActiveCell.Offset(0, 1).Select
                ColumnOffsetBase = ColumnOffsetBase + 1
            Loop
            Range("A1").Select
            Do While ChangeCounter < Range("ChangedDataRows").Value
                If ActiveCell.Value = "Reverse" Or ActiveCell.Value = "Deleted" Then
                    ActiveCell.Offset(0, ColumnOffsetNative).Value = ActiveCell.Offset(0, ColumnOffsetNative).Value * -1
                    ActiveCell.Offset(0, ColumnOffsetBase).Value = ActiveCell.Offset(0, ColumnOffsetBase).Value * -1
                    ChangeCounter = ChangeCounter + 1
                    ActiveCell.Offset(1, 0).Select
                Else
                    ActiveCell.Offset(1, 0).Select
                End If
            Loop
        End If
    
        'Tidy Up Sheets For Review
        'Delta Data
        Range("B2").Select
        Selection.EntireColumn.Delete
        Cells.Select
        Cells.EntireColumn.AutoFit
        If Range("UnmatchedTotal").Value > 2 Then
            Range("A2:" & Range("CurrentColumnEnd").Value & "2").Select
            Selection.Copy
            Range(Selection, Selection.End(xlDown)).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        End If
        'set Report Date to iRecs Business Date (max of report date)
        Set rng = Range("AO1").Offset(1, 0)
        Set rng = Range(rng, rng.End(xlDown))
        
        dt = Application.WorksheetFunction.Max(rng)
        'rng.Value = Format(dt, "DD/MM/YYYY")
        rng.Value = dt
        Range("A1").Select
        
        'Current Data
        If Range("UnmatchedCurrent").Value <> 0 Then
            Sheets("Current Data").Select
            ActiveSheet.ShowAllData
            Range("A2:" & Range("CurrentColumnEnd").Value & "2").Select
            Selection.Copy
            Range(Selection, Selection.End(xlDown)).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Range("A1").Select
        End If
        'Previous Data
        If Range("UnmatchedPrevious").Value <> 0 Then
            Sheets("Previous Data").Select
            ActiveSheet.ShowAllData
            Range("A2:" & Range("PreviousColumnEnd").Value & "2").Select
            Selection.Copy
            Range(Selection, Selection.End(xlDown)).Select
            Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
            Range("A1").Select
        End If
    End If
    ActiveSheet.Range("A1").AutoFilter
    
'Control Sheet
    Sheets("Control").Select
    Range("A1").Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = sCalc

End Sub
Public Sub Formatting()
Dim ws As Worksheet
Dim c As Range
Dim col As Single
Dim b As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
b = Application.Calculation
Application.Calculation = xlCalculationManual

Set ws = ActiveWorkbook.Worksheets("Delta Data")
ws.Activate
Set c = ws.Range("A1")
Do Until c.Value = "" Or c.Column = 100
    If c.Value = "Entry Creation Date" Or c.Value = "Entry Modification Date" Then 'ensure date time preserevd
        col = c.Column
        Columns(col).Select
        Selection.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    Else
        If c.Value Like "*Date" Then 'ensure UK date preserevd
            col = c.Column
            Columns(col).Select
            Selection.NumberFormat = "dd/mm/yyyy"
        End If
    End If
    
    If c.Value Like "Native Amount" Or c.Value Like "Base Amount" Then 'ensure number fomat 2dps
        col = c.Column
        Columns(col).Select
        Selection.NumberFormat = "0.00"
    End If
    Set c = c.Offset(0, 1)
Loop

Application.Calculation = b
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Public Sub Export_csv()
Dim wbNew As Workbook
Dim ws As Worksheet
Dim c As Range, c2 As Range
Dim dt As Date
Dim iRow As Single
Dim StrFile As String, strDt As String, b As String, backupdir As String, irecsdir As String

Call Formatting

Application.DisplayAlerts = False
Application.ScreenUpdating = False

b = Application.Calculation
Application.Calculation = xlCalculationManual
Set ws = ActiveWorkbook.Worksheets("Delta Data")

ws.Activate
backupdir = FwdSlash(Range("_DirBAK").Value)
'irecsdir = "\\172.16.11.106\DataFiles\Enfusion\"
'irecsdir = "\\10.60.20.25\SFTP\Enfusion\"
irecsdir = FwdSlash(Range("_DirDest").Value)

Set c = Range("A1")

Do Until c.Value = "" Or c.Column = 100
    Select Case c.Value
    Case "Entry Modification Date" 'get business date for iRecs (should be an if)
        Set c2 = c.Offset(1, 0)
'        Set c2 = c.End(xlDown)
        Set c2 = Range(c.Offset(1, 0), c.End(xlDown))
    '        c.Select
        dt = Application.WorksheetFunction.Max(c2)
        dt = Format(dt, "dd/mm/yyyy")
        Set c2 = Sheets("Current Data").Range("AP2")
        dt = Application.WorksheetFunction.Max(c2, Application.WorksheetFunction.Max(c2))
'        Select Case dt
'        Case Format(Now(), "dd/mm/yyyy") ' if mod date is today then BD=Previosu BD
'            dt = Format(Application.WorksheetFunction.WorkDay_Intl(dt, -1), "dd/mm/yyyy")
'        Case Format(Application.WorksheetFunction.WorkDay_Intl(dt, 1)) = 1 Or dt = Format(Application.WorksheetFunction.WorkDay_Intl(dt, 1)) = 7 ' Weekend
'            dt = Format(Application.WorksheetFunction.WorkDay_Intl(dt, -1), "dd/mm/yyyy")
'        End Select
        
        If dt = Format(Now(), "dd/mm/yyyy") Or Application.WorksheetFunction.Weekday(dt, 1) = 1 Or Application.WorksheetFunction.Weekday(dt, 1) = 7 Then
            dt = Format(Application.WorksheetFunction.WorkDay_Intl(dt, -1), "dd/mm/yyyy")
        End If
    End Select

    Set c = c.Offset(0, 1)
Loop

strDt = Format(dt, "yyyymmdd")

ws.Copy

StrFile = "Enfusion Cash Activity ITD " & strDt
'ChDir "\\172.16.11.106\DataFiles\Enfusion\"
Set wbNew = ActiveWorkbook

wbNew.SaveAs Filename:=backupdir & StrFile & "_" & Format(Now, "hh_mm_ss"), FileFormat:=xlCSV
wbNew.SaveAs Filename:=irecsdir & StrFile, FileFormat:=xlCSV
'Call SaveAs_Defined(xlCSV, "\\172.16.11.106\DataFiles\Enfusion\", StrFile)

wbNew.Close

    '***
    '*** SFG Added 11/07/22 When Irecs Prevented Pushing To Their Site - Now Email To Operations & Save Down From There To SFTP To Be Copied Over ***
    '***
    
        'ActiveSheet.Copy
        'ActiveWorkbook.SaveAs Filename:=backupdir & StrFile & ".csv", FileFormat:=xlCSV
        'ActiveWorkbook.SendMail "operations@letterone.com", StrFile & ".csv"
        'ActiveWorkbook.Close False
        'Kill backupdir & StrFile & ".csv"
    
    '***
    'End Of Code Addition
    '***

Application.Calculation = b
Application.DisplayAlerts = True
Application.ScreenUpdating = True

Sheets("Control").Activate

MsgBox "File has been saved successfully in Datafiles and Backup in the archive"
End Sub
Public Sub Copy_Current_to_Previous()
Dim ws As Worksheet
Dim i As Integer
Dim b As String

i = MsgBox("Are you sure you want to copy contents of 'Current Data' tab to 'Previous Data' tab", vbYesNo)
If i = 7 Then Exit Sub
Set ws = ActiveSheet
Application.DisplayAlerts = False
Application.ScreenUpdating = False

b = Application.Calculation
Application.Calculation = xlCalculationManual

'Previous Data Clear
If Range("UnmatchedPrevious").Value <> 0 Then
Sheets("Previous Data").Select
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    Else
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    End If
'    ActiveSheet.ShowAllData
    Range("A2", Range("A2").SpecialCells(xlCellTypeLastCell)).Rows.Delete
End If

'Current Data
If Range("UnmatchedCurrent").Value <> 0 Then
    Sheets("Current Data").Select
    If Not ActiveSheet.AutoFilterMode Then
        ActiveSheet.Range("A1").AutoFilter
    Else
        If ActiveSheet.FilterMode Then
            ActiveSheet.ShowAllData
        End If
    End If
    Range("c2", Range("c2").SpecialCells(xlCellTypeLastCell)).Select
    Selection.Copy
    Sheets("Previous Data").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
End If
ws.Activate
Application.Calculation = b
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Function NewestFile(Directory, FileNameStart, FileSpec)
' Returns the name of the most recent file in a Directory
' That matches the FileSpec (e.g., "*.xls").
' Returns an empty string if the directory does not exist or
' it contains no matching files
    Dim Filename As String
    Dim MostRecentFile As String
    Dim MostRecentDate As Date
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    If Right(FileNameStart, 1) <> "*" Then FileNameStart = FileNameStart & "*"

    Filename = Dir(Directory & FileNameStart & FileSpec, 0)
    If Filename <> "" Then
        MostRecentFile = Filename
        MostRecentDate = FileDateTime(Directory & Filename)
        Do While Filename <> ""
            If FileDateTime(Directory & Filename) > MostRecentDate Then
                MostRecentFile = Filename
            End If
            Filename = Dir
        Loop
    End If
    NewestFile = Directory & MostRecentFile
    
End Function
Sub Import_Latest_Enfusion_Data()
Dim TW As Workbook
Set TW = ThisWorkbook

Application.DisplayAlerts = False

Workbooks.Open (NewestFile(Range("_Dir").Value, "PROD_GL Cash Transactions ITD-*", ".xls"))

Dim newdata As Worksheet
Dim ndatarng As Range
Set newdata = ActiveWorkbook.Worksheets("Sheet 1")
Set ndatarng = newdata.Range("A1").CurrentRegion

ndatarng.Copy

TW.Worksheets("Current Data").Range("C1").PasteSpecial Paste:=xlPasteValues
        
ActiveWorkbook.Close

Sheets("Current Data").Select
Rows("2:2").Select
Selection.Copy
Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
Range("A2").Select
Application.CutCopyMode = False

Application.DisplayAlerts = True

ThisWorkbook.Sheets("Control").Select

MsgBox "Import Complete"

End Sub

