Attribute VB_Name = "Sorting"
Private deviatedCols() As Integer
Private notDeviation() As Integer
Private j As Integer
Private l As Integer
Public targetWsh As Worksheet

Public Sub sort()
    If maxCol > 480 And maxCol < 485 Then                       'sort VT_TRANSACTIONS by pre-defined sort key
        processTransactionsCompare
    ElseIf maxCol > 1000 Then                                   'sort VT_RECONCIL by pre-defined sort key
        Call sortReconcils(masterWsh)
        Call sortReconcils(testWsh)
    ElseIf Trim(LCase(masterWsh.Cells(headerRow, 2))) = "message queue" Then
        processMQcompare
    Else
        If Not toSortByFullKey Then
            Call sortAnything(masterWsh)
            Call sortAnything(testWsh)
        End If
'        intelliSort
    End If
End Sub

Public Sub sortByFullKey(ByVal wsh As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Integer
    Dim i As Integer
    
    lastRow = wsh.[A1].End(xlDown).Row
    lastCol = wsh.UsedRange.Column
    With wsh.sort
        .SortFields.Clear
        For i = 1 To lastCol
            .SortFields.Add Key:=wsh.Range(wsh.Cells(headerRow, i), wsh.Cells(lastRow, i)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
        .SetRange wsh.Range(wsh.Cells(headerRow, 1), wsh.Cells(lastRow, i))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Sub matchByKey(ByVal wsh As Worksheet)
    Dim uniqKey As String

    Call getTargetWsh(wsh)
        
    If maxCol > 480 Then                        'key for VT_TRANSACTIONS
        uniqKey = "=MP" & firstDataRow & "&"";""&K" & firstDataRow & "&"";""&O" & firstDataRow & "&"";""&J" & firstDataRow & "&"";""&T" & firstDataRow & _
                "&"";""&G" & firstDataRow & "&"";""&PY" & firstDataRow & "&"";""&A" & firstDataRow & "&"";""&P" & firstDataRow & "&"";""&I" & firstDataRow & _
                "&"";""&RO" & firstDataRow & "&"";""&AP" & firstDataRow & "&"";""&JI" & firstDataRow & "&"";""&NH" & firstDataRow & "&"";""&AH" & firstDataRow
    Else                                        'key-template for everything
        uniqKey = "=A" & firstDataRow & "&"";""&B" & firstDataRow & "&"";""&C" & firstDataRow & "&"";""&D" & firstDataRow
        With wsh.Cells(headerRow, maxCol + 2)
            .AddComment
            .Comment.Visible = False
            .Comment.Text Text:="The key is formed from A:D cells and is not final. Feel free to modify"
        End With
    End If
        
    With wsh
        .Range(.Cells(firstDataRow, maxCol + 2), .Cells(firstDataRow, maxCol + 3)).NumberFormat = "General"
        .Cells(headerRow, maxCol + 2).Value = "uniqueKey"
        .Cells(firstDataRow, maxCol + 2).Formula = uniqKey
        .Cells(headerRow, maxCol + 3).Value = "Match"
        .Cells(firstDataRow, maxCol + 3).Formula = "=IF(ISERROR(MATCH(" & .Cells(firstDataRow, maxCol + 2).Address(0, 0) & "," & targetWsh.Name & "!" & _
            .Cells(firstDataRow, maxCol + 2).Address & ":" & .Cells(targetWsh.UsedRange.Rows.Count, maxCol + 2).Address & ",0)),1,0)"
        If .UsedRange.Rows.Count < 10000 Then
            .Range(.Cells(firstDataRow, maxCol + 2), .Cells(firstDataRow, maxCol + 3)).Copy Destination:=.Range(.Cells(firstDataRow + 1, maxCol + 2), .Cells(.UsedRange.Rows.Count, maxCol + 3))
        End If
    End With
End Sub

Public Function getTargetWsh(ByVal wsh As Worksheet) As Worksheet
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

Public Sub sortTransactions(ByVal wsh As Worksheet)

'    firstDataRow = 2
'    headerRow = firstDataRow -1
'    maxRow = 4650
'    Set wsh = ActiveWorkbook.Sheets(1)

    With wsh.sort
        With .SortFields
            .Clear
            .Add Key:=Range("RP" & firstDataRow & ":RP" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .Add Key:=Range("RR" & firstDataRow & ":RR" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("K" & firstDataRow & ":K" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .Add Key:=Range("O" & firstDataRow & ":O" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("J" & firstDataRow & ":J" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("T" & firstDataRow & ":T" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .Add Key:=Range("G" & firstDataRow & ":G" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("PY" & firstDataRow & ":PY" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("P" & firstDataRow & ":P" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
            .Add Key:=Range("A" & firstDataRow & ":A" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("I" & firstDataRow & ":I" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("RO" & firstDataRow & ":RO" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("AP" & firstDataRow & ":AP" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("JI" & firstDataRow & ":JI" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("NH" & firstDataRow & ":NH" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("AH" & firstDataRow & ":AH" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("AD" & firstDataRow & ":AD" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("NJ" & firstDataRow & ":NJ" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End With
        .SetRange Range("A" & headerRow & ":RR" & maxRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Sub sortReconcils(ByVal wsh As Worksheet)

'    firstDataRow = 2
'    headerRow = firstDataRow -1
'    maxRow = 4650
'    Set wsh = ActiveWorkbook.Sheets(1)

    With wsh.sort
        With .SortFields
            .Clear
            .Add Key:=Range("DF" & firstDataRow & ":DF" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("B" & firstDataRow & ":B" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("C" & firstDataRow & ":C" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("D" & firstDataRow & ":D" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("PD" & firstDataRow & ":PD" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("A" & firstDataRow & ":A" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Add Key:=Range("G" & firstDataRow & ":G" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        End With
        .SetRange Range("A" & headerRow & ":ALR" & maxRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Sub sortMQ(ByVal wsh As Worksheet)
'    With wsh.sort
'        With .SortFields
'            .Clear
'            .Add Key:=Range("NJ" & firstDataRow & ":NJ" & maxRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'        End With
'        .SetRange Range("A" & headerRow & ":RR" & maxRow)
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
End Sub
Public Sub sortAnything(ByVal wsh As Worksheet)

'    firstDataRow = 2
'    headerRow = firstDataRow -1
'    maxRow = 4650
'    Set wsh = ActiveWorkbook.Sheets(1)

    With wsh.sort
        With .SortFields
            .Clear
'            .Add Key:=Range(wsh.Cells(firstDataRow, maxCol + 3), wsh.Cells(maxRow, maxCol + 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
'            .Add Key:=Range(wsh.Cells(firstDataRow, maxCol + 2), wsh.Cells(maxRow, maxCol + 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        End With
        .SetRange Range(wsh.Cells(headerRow, 1), wsh.Cells(maxRow, maxCol + 3))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Public Sub separateExtraRows()
    'shifts extra rows in Master or Test lower, for them not to overlap
    Dim MasterExtraRowsCount As Long
    Dim TestExtraRowsCount As Long

    MasterExtraRowsCount = WsF.Sum(Range(masterWsh.Cells(firstDataRow, maxCol + 3), masterWsh.Cells(maxRowMaster, maxCol + 3)))
    TestExtraRowsCount = WsF.Sum(Range(testWsh.Cells(firstDataRow, maxCol + 3), testWsh.Cells(maxRowTest, maxCol + 3)))

    If MasterExtraRowsCount > TestExtraRowsCount And TestExtraRowsCount > 0 Then
        If MasterExtraRowsCount > maxRowMaster * 0.1 Or MasterExtraRowsCount > 50 Then Exit Sub
        masterWsh.Rows(maxRowMaster - MasterExtraRowsCount + 1).Resize(TestExtraRowsCount).Insert
        maxRowMaster = maxRowMaster + TestExtraRowsCount
    ElseIf MasterExtraRowsCount < TestExtraRowsCount And MasterExtraRowsCount > 0 Then
        If TestExtraRowsCount > maxRowTest * 0.1 Or TestExtraRowsCount > 50 Then Exit Sub
        testWsh.Rows(maxRowTest - TestExtraRowsCount + 1).Resize(MasterExtraRowsCount).Insert
        maxRowTest = maxRowTest + MasterExtraRowsCount
    End If
    
    maxMatchRow = maxRowMaster - MasterExtraRowsCount   'maxMatchRow = maxRowTest - TestExtraRowsCount
    maxRow = WsF.Max(maxRowMaster, maxRowTest)
End Sub

Public Sub intelliSort()
    Application.ScreenUpdating = False
    
     j = 1
     l = 1
     
    'show all except zeros
    compareWsh.Range(compareWsh.Cells(headerRow, 1), compareWsh.Cells(maxCol + 3, maxRow)).AutoFilter Field:=maxCol + 1, Criteria1:=">0"

    Call defineNewSortField
    
    Application.ScreenUpdating = True
End Sub

Private Sub defineNewSortField()
    Dim deviationAmountInCol() As Integer
    Dim mostlyDeviatedCol As Integer
    Dim i As Integer
    
    ReDim deviationAmountInCol(1 To maxCol)
    ReDim Preserve deviatedCols(1 To j)
    
    mostlyDeviatedCol = 1
    For i = 1 To maxCol
        deviationAmountInCol(i) = WsF.CountIf(Range(compareWsh.Cells(firstDataRow, i), compareWsh.Cells(maxRow, i)), "<>0")
        If deviationAmountInCol(i) > deviationAmountInCol(mostlyDeviatedCol) Then mostlyDeviatedCol = i
    Next i
    
    If UBound(Filter(deviatedCols, mostlyDeviatedCol)) > -1 Then
        a = UBound(Filter(deviatedCols, mostlyDeviatedCol)) > -1
    
        ReDim Preserve notDeviation(1 To l)
        notDeviation(l) = mostlyDeviatedCol
        l = l + 1
        deviatedCols(j) = WsF.Large(deviationAmountInCol, l)
        'once again
    Else
        deviatedCols(j) = mostlyDeviatedCol
    End If
        
    'apply new sorting
    Call applySorting(masterWsh)
    Call applySorting(testWsh)
    
    Calculate
    
    j = j + 1
        
    If Not noDeviationsLeft Then Call defineNewSortField
End Sub

Private Sub applySorting(ByVal wshToSort As Worksheet)
    Dim k As Integer
    
    With wshToSort.sort
        .SortFields.Clear
        'add all sort fields
        For k = 1 To UBound(deviatedCols)
            .SortFields.Add Key:=Range(wshToSort.Cells(headerRow, deviatedCols(k)), wshToSort.Cells(maxRow, deviatedCols(k))), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 'xlSortTextAsNumbers
        Next k
        .SetRange Range(wshToSort.Cells(headerRow, 1), wshToSort.Cells(maxRow, maxCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Function noDeviationsLeft() As Boolean
    If WsF.CountIf(Range(compareWsh.Cells(firstDataRow, maxCol + 1), compareWsh.Cells(maxRow, maxCol + 1)), "<>0") = 0 Then noDeviationsLeft = True
End Function
