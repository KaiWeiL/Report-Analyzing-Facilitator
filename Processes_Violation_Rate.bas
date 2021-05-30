Attribute VB_Name = "Processes_Violation_Rate"
Option Explicit

' Find empty columns on the right of data columns
Function FindEmptyColumns() As String()
    Dim EmptyColns(2) As String
    Dim ReturnArry(2) As String
    EmptyColns(0) = Range("ZZ1").End(xlToLeft).Offset(0, 1).Address
    EmptyColns(1) = Range("ZZ1").End(xlToLeft).Offset(0, 2).Address
    FindEmptyColumns = EmptyColns
End Function

Function MergeAndPlaceVendor(dataRowNumbers As Integer, RowToMerge1 As String, RowToMerge2 As String)
    Dim Left, Right As String
    Dim i As Integer
    Dim LeftColn, RightColn As String
    
    'Left is an easily used word. Therefore, it's better to prepend it with VBA.
    ' Reason of the following snippet: No regular expression library on VBA for Mac.
        If RowToMerge1 Like "$[A-Z]$[0-9]" And RowToMerge2 Like "$[A-Z]$[0-9]" Then
            LeftColn = VBA.Right(VBA.Left(RowToMerge1, 2), 1)
            RightColn = VBA.Right(VBA.Left(RowToMerge2, 2), 1)
        ElseIf RowToMerge1 Like "$[A-Z][A-Z]$[0-9]" And RowToMerge2 Like "$[A-Z][A-Z]$[0-9]" Then
            LeftColn = VBA.Right(VBA.Left(RowToMerge1, 3), 2)
            RightColn = VBA.Right(VBA.Left(RowToMerge2, 3), 2)
        End If
    '
    
    For i = 1 To dataRowNumbers
        Left = Range(LeftColn & i).Value
        Right = Range(RightColn & i).Value
        If Left <> "" Then
            Range("C" & (i + 1)).Value = Left & " " & Right
        Else: Range("C" & (i + 1)).Value = Right
        End If
    Next
End Function

Sub GenProcessViolationRate()
'
' GenProcessViolationRate Macro
' Generate a new sheet after the current active sheet. Populate it with the necessary data relating to the processes that violate the prevention policy. Sort it by violation rate.
'
    
    Dim dataRowNumbers As Integer   'Not including header
    dataRowNumbers = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row - 1
    Dim EmptyColn() As String
    Dim LeftMostColn As String   'For metadata deletion
    
    ' Generate header and count violations with CountIf
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Process_Violation_Rate"
    Range("A1").Select
    ActiveCell.Value = "Process Name"
    Range("B1").Select
    ActiveCell.Value = "Violation Count"
    Range("C1").Select
    ActiveCell.Value = "Possible Vendor"
    Range("A2").Select
    
        'Copy and paste all processe names
    Sheets("Target_Windows_Logs").Select
    Range("V2").Select
    Range(Selection, Selection.Offset(dataRowNumbers - 1)).Select
    Selection.Copy
    Sheets("Process_Violation_Rate").Select
    Range("A6666").End(xlUp).Offset(1).Select   '6666 may need to change for larger data input
    ActiveSheet.Paste
    Columns("A:A").EntireColumn.AutoFit
        'Copy and paste all potential application vendors
    Sheets("Target_Windows_Logs").Select
    Range("W2").Select
    Range(Selection, Selection.Offset(dataRowNumbers - 1)).Select
    Selection.Copy
    Sheets("Process_Violation_Rate").Select
    EmptyColn = FindEmptyColumns()
    Range(EmptyColn(0)).Select
    ActiveSheet.Paste
    Sheets("Target_Windows_Logs").Select
    Range("X2").Select
    Range(Selection, Selection.Offset(dataRowNumbers - 1)).Select
    Selection.Copy
    Sheets("Process_Violation_Rate").Select
    Range(EmptyColn(1)).Select
    ActiveSheet.Paste

    MergeAndPlaceVendor dataRowNumbers:=dataRowNumbers, RowToMerge1:=EmptyColn(0), RowToMerge2:=EmptyColn(1)
        
    
        'Remove metadata
    If EmptyColn(0) Like "$[A-Z]$[0-9]" Then
        LeftMostColn = VBA.Right(VBA.Left(EmptyColn(0), 2), 1)
    ElseIf EmptyColn(0) Like "$[A-Z][A-Z]$[0-9]" Then
        LeftMostColn = VBA.Right(VBA.Left(EmptyColn(0), 3), 2)
    End If
    
    Dim x As Integer
    For x = 1 To 2  ' 2 here could be changed for an variable for future expansion
        Columns(LeftMostColn).EntireColumn.Delete
    Next x
        '
    Columns("C").EntireColumn.AutoFit

        ' Remove duplicate rows
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$2:$C$" & dataRowNumbers).RemoveDuplicates Columns:=1, Header:=xlNo

        ' Get violation count
    Range("B2").Select
    ActiveCell.Value = _
        "=COUNTIF(Target_Windows_Logs!R2C22:R724C22,Process_Violation_Rate!RC[-1])"
    Selection.AutoFill Destination:=Range("B2:B28")
        
        ' Remove CountIf formula and sort violation count from largest to smallest
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").CurrentRegion.Select
    Range("B2").Activate
    ActiveWorkbook.Worksheets("Process_Violation_Rate").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Process_Violation_Rate").Sort.SortFields.Add2 Key:=Range("B2:B28") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Process_Violation_Rate").Sort
        .SetRange Range("A1:B28")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("B2").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("A1").Select
    
End Sub





