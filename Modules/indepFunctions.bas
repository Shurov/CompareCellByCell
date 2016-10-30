Attribute VB_Name = "indepFunctions"
Private targetWsh As Worksheet
Public veryFirstRow As Integer
Public veryLastRow As Long
Public emptyCellCol As Integer
Public sortByCols As String

Public Sub sortGroupsInit(control As IRibbonControl)
    sortGroups.Show
End Sub

Public Sub sortWithinGroups(ByVal wsh As Worksheet)
    Dim groupStartRow As Long
    Dim groupEndRow As Long
    Dim maxCol As Integer
    Dim i As Long
    Dim j As Integer
    Dim sortBy() As String
    
    maxCol = wsh.UsedRange.Columns.Count
    sortBy = Split(sortByCols, ",")
    
    For i = veryFirstRow To veryLastRow
        If wsh.Cells(i, 1) <> vbNullString And wsh.Cells(i, emptyCellCol) = vbNullString Then groupStartRow = i + 1
        Do Until wsh.Cells(i, 1) <> vbNullString And wsh.Cells(i + 1, emptyCellCol) = vbNullString
            i = i + 1
            If i > veryLastRow Then Exit Sub
        Loop
        If groupStartRow = 0 Then groupStartRow = veryFirstRow
        groupEndRow = i
        
        With wsh.sort
            .SortFields.Clear
            For j = 0 To UBound(sortBy)
                .SortFields.Add Key:=wsh.Range(sortBy(j) & groupStartRow & ":" & sortBy(j) & groupEndRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            Next j
            .SetRange wsh.Range(wsh.Cells(groupStartRow, 1), wsh.Cells(groupEndRow, maxCol))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next i
End Sub

Public Sub sortMultipleSheets(control As IRibbonControl)
    Dim sortDiagAnswer As Boolean
    
    If Not (sheetExists("Master") And sheetExists("Test")) Then
        MsgBox "Your workbook doesn't contain Master and/or Test worksheets. Please make sure you have sheets named ""Master"" and ""Test""", vbOKOnly + vbCritical
        End
    End If
    
    On Error Resume Next
    
    sortDiagAnswer = Application.Dialogs(xlDialogSort).Show
    If sortDiagAnswer Then
        Call sortOtherSheet(ActiveSheet)
        Workbook.RefreshAll
'        With Worksheets("Compare")
'            .EnableCalculation = False
'            .EnableCalculation = True
'            .Calculate
'        End With
        Application.Calculate
    End If
    
    If Err.Number = 1004 Then
        MsgBox "We couldn't do this for the selected range of cells. Select a single cell within range of data and then try again.", vbOKOnly + vbExclamation, "Microsoft Excel"
    ElseIf Err.Number = 400 Then
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub sortOtherSheet(ByVal sortedWsh As Worksheet)
    Dim srt As sort
    Dim i As Integer
    
    Call getTargetWsh(sortedWsh)

    Set srt = sortedWsh.sort

    With targetWsh.sort
        With .SortFields
            .Clear
            For i = 1 To srt.SortFields.Count
                 .Add Key:=Range(srt.SortFields(i).Key.Address), SortOn:=xlSortOnValues, Order:=xlAscending
            Next i
        End With
        .SetRange Range(srt.Rng.Address)
        .Header = srt.Header
        .MatchCase = srt.MatchCase
        .Orientation = srt.Orientation
        .SortMethod = srt.SortMethod
        .Apply
    End With
End Sub

Private Function getTargetWsh(ByVal wsh As Worksheet) As Worksheet
'    If wsh Is masterWsh Then
'        Set targetWsh = testWsh
'    Else
'        Set targetWsh = masterWsh
'    End If
    
    If wsh.Name = "Master" Then
        Set targetWsh = Worksheets("Test")
    Else
        Set targetWsh = Worksheets("Master")
    End If
End Function

Public Function sheetExists(ByVal shtName As String) As Boolean
    Dim sht As Worksheet

    On Error Resume Next
    Set sht = ActiveWorkbook.Sheets(shtName)
    On Error GoTo 0
    sheetExists = Not sht Is Nothing
End Function

Public Sub parseMQ(control As IRibbonControl)
    If Not (sheetExists("Master") And sheetExists("Test")) Then
        MsgBox "Your workbook doesn't contain Master and/or Test worksheets. Please make sure you have sheets named ""Master"" and ""Test""", vbOKOnly + vbCritical
        End
    End If
   
    Set masterWsh = ActiveWorkbook.Sheets("Master")
    Set testWsh = ActiveWorkbook.Sheets("Test")
    
    Call parseMQcontent(masterWsh)
    Call parseMQcontent(testWsh)
End Sub

Private Sub parseMQcontent(ByVal wsh As Worksheet)
    Dim i As Integer
    
    With wsh
        maxCol = .UsedRange.Columns.Count
        maxRow = .UsedRange.Rows.Count
        For i = 1 To 7
            If .Cells(i, 3) = "Message content" Then
                firstDataRow = i + 1
                Exit For
            End If
        Next i
        If firstDataRow = 0 Then
            For i = 1 To 8
                If Left(.Cells(i, 3), 5) = "BASE_" Then
                    firstDataRow = i
                    Exit For
                End If
            Next i
        End If
        
        If firstDataRow <> 0 Then
            .Range(.Cells(firstDataRow, 3), .Cells(maxRow, 3)).Copy Destination:=.Cells(firstDataRow, maxCol + 1)
            Call textToCols(wsh, True, True, False, False, False)
        Else
            MsgBox "The data doesn't look like MQ. MQ criterias are:" & vbNewLine & _
            " - ""Message content"" text in the column C header" & vbNewLine & _
            " - Column C contains ""BASE_"" text", vbOKOnly, "Error"
        End If
    End With
    
    endMacro
End Sub

Public Sub replaceZeros(control As IRibbonControl)
    'replace 0-deviation-columns with values to decrease file size
    Dim compareWsh As Worksheet
    Dim maxRow As Integer
    Dim replacedColsCount As Integer
    Dim col As Integer
     
    If Not sheetExists("Compare") Then
        MsgBox "Your workbook doesn't contain Compare worksheet of Master with Test. Please make sure you have sheet named ""Compare""", vbOKOnly + vbCritical
        End
    End If
    
    Application.ScreenUpdating = False
'    progressBar.Show (vbModeless)
'    progressBar.taskName = "Replacing columns without deviations:"
'    progressBar.ProgressBar1 = 0
    Application.Calculate

    Set compareWsh = Worksheets("Compare")
    With compareWsh
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0
        firstDataRow = .UsedRange.Find("Diff").Row + 1
        maxRow = .UsedRange.Rows.Count
        maxCol = .UsedRange.Columns.Count - 3
        If maxCol = 494 Then maxCol = 483
        For col = 1 To maxCol
'            progressBar.ProgressBar1 = (col / maxCol) * 100
'            DoEvents
            If .Cells(maxRow, col) = 0 And .Cells(maxRow, col).HasFormula Then
                .Range(.Cells(firstDataRow + 1, col), .Cells(maxRow - 2, col)).Copy
                .Range(.Cells(firstDataRow + 1, col), .Cells(maxRow - 2, col)).PasteSpecial xlPasteValues
'                replacedColsCount = replacedColsCount + 1
'                progressBar.currentStatus = "Replaced " & replacedColsCount & " undeviated columns with values"
'                DoEvents
            End If
        Next col
    End With
    
'    progressBar.Hide
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub
