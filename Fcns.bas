Attribute VB_Name = "Fcns"
Option Explicit
Public Sub Last_Cell_CleanUp_Call()
'If ws is not specified then runs for all sheets in activeworkbook
Call Last_Cell_CleanUp

End Sub
Public Sub Last_Cell_CleanUp(Optional ws As Worksheet)
'If ws is not specified then runs for all sheets in activeworkbook
'ReReferences LastCell to actual last cell in sheet
Dim i As Single
Dim Str As String
'Dim ws As Worksheet
Dim rng As Range
Dim w As Worksheet

'Dim N As Single
Str = Application.Calculation
Set rng = ActiveCell
Application.Calculation = xlCalculationManual
If Not ws Is Nothing Then
    ws.Activate
    Call DeleteUnusedFormats
    
Else
    For Each w In Worksheets
        w.Activate
        If w.ProtectContents = False Then
            Call DeleteUnusedFormats
        End If
    Next w
End If

On Error Resume Next
rng.Activate
On Error GoTo 0

Application.Calculation = Str
End Sub
Public Sub DeleteUnusedFormats()
     Dim lLastRow As Long, lLastColumn As Long
     Dim lRealLastRow As Long, lRealLastColumn As Long
     With Range("A1").SpecialCells(xlCellTypeLastCell)
         lLastRow = .Row
         lLastColumn = .Column
     End With
     On Error GoTo Crash1:
     lRealLastRow = Cells.Find("*", Range("A1"), xlFormulas, , xlByRows, xlPrevious).Row
     On Error GoTo Crash2:
     lRealLastColumn = Cells.Find("*", Range("A1"), xlFormulas, , _
               xlByColumns, xlPrevious).Column
     On Error GoTo 0
     If lRealLastRow < lLastRow Then
         Range(Cells(lRealLastRow + 1, 1), Cells(lLastRow, 1)).EntireRow.Delete
     End If
     If lRealLastColumn < lLastColumn Then
         Range(Cells(1, lRealLastColumn + 1), _
              Cells(1, lLastColumn)).EntireColumn.Delete
     End If
     ActiveSheet.UsedRange 'Resets LastCell
     ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell) = Cells(lRealLastRow, lRealLastColumn)
     ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Select
     
     Exit Sub
Crash1:
    lRealLastRow = 1
    Resume Next
Crash2:
    lRealLastColumn = 1
    Resume Next
End Sub
Public Sub SaveAs_Defined(Optional iFormat, Optional strDir As String, Optional strInitial As String)
'Saves copy of the activeworkbook via a dialogue box which requests dir and filename. inputs are fileformat id, dir and filename

'Save as with dialogue
Dim StrFile As String

If iFormat = 0 Then
    iFormat = ActiveWorkbook.FileFormat
End If

ChDir strDir

Do
    'StrFile = Application.GetSaveAsFilename(strDir & strInitial, iFormat, , "Save to '" & strDir & "'?")  ' For Dev to include filefilter eg
    '2: StrFile = Application.GetSaveAsFilename(strDir & strInitial, , , "Save to '" & strDir & "'?")
    StrFile = Application.GetSaveAsFilename(strInitial, , , "Save to '" & strDir & "'?")
    
    If StrFile = "False" Then
        StrFile = MsgBox("Nothing Saved!", vbOKOnly)
        Exit Sub
    End If
Loop Until StrFile <> "False"
'If InStr(StrFile, ".csv") > 0 Then
'    StrFile = Left(StrFile, InStr(StrFile, ".csv") - 1)
'End If
If Right(StrFile, 1) = "." Then
    StrFile = Left(StrFile, Len(StrFile) - 1)
End If
ActiveWorkbook.SaveAs Filename:=StrFile, FileFormat:=iFormat
'ActiveWorkbook.SaveAs (StrFile)

End Sub
Public Sub test()

Dim arr As Variant
Dim strDir As String, str1 As String

strDir = "c:\temp\"
arr = DirMax_x2(strDir, "PROD_GL Cash Transactions ITD-1_")

End Sub
Public Function DirMax_x2(strDir As String, Optional str1 As String) As Variant
'Find most recent File with a date in it's name
'

Dim StrFile As String, Str As String
Dim dtRec As Date, dtMax As Date
Dim iYear As Integer, iMonth As Integer, iDay As Integer, iRow As Integer

If str1 = "" Then str1 = "*"
'If str2 = "" Then str1 = "*"

StrFile = Dir(strDir & "*" & str1 & "*")

'find max date of pos file
Do While StrFile <> ""
    
    Str = Mid(StrFile, InStr(1, StrFile, "201", 1), 8) ', "yyyymmdd")
    iYear = Left(Str, 4)
    iMonth = Mid(Str, 5, 2)
    iDay = Right(Str, 2)
    dtRec = DateSerial(iYear, iMonth, iDay)
    If dtRec > dtMax Then
        dtMax = dtRec
        DirMax = StrFile
    End If
    StrFile = Dir()
Loop

End Function
Sub ListFilesinFolder()

    Dim FSO As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim FileItem As Scripting.File

    SourceFolderName = "C:\Users\Santosh"

    Set FSO = New Scripting.FileSystemObject
    Set SourceFolder = FSO.GetFolder(SourceFolderName)

    Range("A1:C1") = Array("text file", "path", "Date Last Modified")

    i = 2
    For Each FileItem In SourceFolder.Files
        Cells(i, 1) = FileItem.Name
        Cells(i, 2) = FileItem
        Cells(i, 3) = FileItem.DateLastModified
        i = i + 1
    Next FileItem

    Set FSO = Nothing

End Sub
Sub dp()
'
' dp Macro
'

'
    Selection.NumberFormat = "0.00"
End Sub


Public Function FwdSlash(Str As String)

If Str = "" Then
    FwdSlash = ""
Else
    If Right(Str, 1) <> "\" Then
        Str = Str & "\"
    End If
    FwdSlash = Str
End If



End Function
