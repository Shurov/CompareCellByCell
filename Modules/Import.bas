Attribute VB_Name = "Import"
Option Explicit

Public masterFileName As String
Private testFileName As String

Private comparedFile As Workbook
Private masterFile As Workbook
Private testFile As Workbook

Public masterWsh As Worksheet
Public testWsh As Worksheet
Public compareWsh As Worksheet

Private deletedColsCount As Integer
Public semiColDelim As Boolean
Public commaDelim As Boolean
Public tabDelim As Boolean
Public spaceDelim As Boolean
Public otherDelim As Variant
Public toSortByFullKey As Boolean
Public toCopyFormulas As Boolean
Public toSplitExtraRows As Boolean
Public toSplitMQcont As Boolean
Public c As Boolean
Public maxRow As Long
Public maxRowMaster As Long
Public maxRowTest As Long
Public maxMatchRow As Long
Public maxCol As Integer
Public firstDataRow As Integer
Public headerRow As Integer
Public WsF As WorksheetFunction

Public Sub importAndCompare(control As IRibbonControl)
    Dim maxColMaster As Integer
    Dim maxColTest As Integer
    Dim unknownFileType As Integer
    Dim wsh As Worksheet
    
    Set WsF = Application.WorksheetFunction
    
    'choose Master & Test files (Excel / csv / txt)
    masterFileName = chooseFile("Master")
    testFileName = chooseFile("Test")
       
    'file preview
    dataPreview.Show
       
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    'create new workbook for Master & Test and compared data
    Set comparedFile = Workbooks.Add
    Set compareWsh = comparedFile.Sheets(1)
    compareWsh.Name = "Compare"
    
    'detect input file format and then import data
    Dim fileExtention As String
    fileExtention = LCase(Right(masterFileName, 4))
    If InStr(masterFileName, ".xls") > 0 Or fileExtention = ".pdf" Or fileExtention = ".rpt" Then
        importExcelBasedData
    ElseIf fileExtention = ".txt" Or fileExtention = ".log" Or fileExtention = ".csv" Then
        importTxtBasedData
    Else
        unknownFileType = MsgBox("Unknown data type. Do you want to treat file as plain-text?", vbOKCancel + vbExclamation, "Error!")
        If unknownFileType = vbOK Then
            importTxtBasedData
        ElseIf unknownFileType = vbCancel Then
            comparedFile.Close False
            endMacro
        End If
    End If
    
    checkIfNoHeader
    
    'write formulas to compareWsh
    maxRowMaster = masterWsh.UsedRange.Rows.Count
    maxRowTest = testWsh.UsedRange.Rows.Count
    maxRow = WsF.Max(maxRowMaster, maxRowTest)
    maxColMaster = masterWsh.UsedRange.Columns.Count
    maxColTest = testWsh.UsedRange.Columns.Count
    maxCol = WsF.Max(maxColMaster, maxColTest)
   
    'add *ExtraLine* column to M & T sheets
    Call addColForExtraRow(masterWsh)
    Call addColForExtraRow(testWsh)
    
    deleteUnnecessaryCols
    
    Call matchByKey(masterWsh)
    Call matchByKey(testWsh)
    Calculate
   
    'auto-sorting
    sort
    
    If toSplitExtraRows Then separateExtraRows
    
    'fill compareWsh
    With compareWsh
        If maxColMaster >= maxColTest Then
            masterWsh.Range(masterWsh.Cells(headerRow, 1), masterWsh.Cells(headerRow, maxCol)).Copy Destination:=.Cells(headerRow, 1)
        Else
            testWsh.Range(testWsh.Cells(headerRow, 1), testWsh.Cells(headerRow, maxCol)).Copy Destination:=.Cells(headerRow, 1)
        End If
        .Cells(headerRow, maxCol + 1) = "Diff"
        .Cells(headerRow, maxCol + 2) = "Description"
        .Cells(headerRow, maxCol + 3) = "Comment"
        
        .Cells(firstDataRow, 1).Formula = "=If(Master!A" & firstDataRow & "=Test!A" & firstDataRow & ", 0, Master!A" & firstDataRow & " & "" | "" & Test!A" & firstDataRow & ")"
        
        .Cells(firstDataRow, 1).Copy Destination:=.Range(.Cells(firstDataRow, 2), .Cells(firstDataRow, maxCol)).SpecialCells(xlCellTypeVisible)
        
        If maxCol > 480 And maxCol < 485 Then
            .Cells(firstDataRow, maxCol + 1).Formula = "=Countif(A" & firstDataRow & ",""<>0"")+Countif(G" & firstDataRow & ":" & .Cells(firstDataRow, maxCol).Address(0, 0) & ",""<>0"")-Countblank(A" & firstDataRow & ":" & .Cells(firstDataRow, maxCol).Address(0, 0) & ")"
        ElseIf Trim(LCase(masterWsh.Cells(headerRow, 2))) = "message queue" Then
            .Cells(firstDataRow, maxCol + 1).Formula = "=Countif(B" & firstDataRow & ":E" & maxCol & ",""<>0"")-Countblank(B" & firstDataRow & ":E" & maxCol & ")"
        Else
            .Cells(firstDataRow, maxCol + 1).Formula = "=Countif(A" & firstDataRow & ":" & .Cells(firstDataRow, maxCol).Address(0, 0) & ",""<>0"")-Countblank(A" & firstDataRow & ":" & .Cells(firstDataRow, maxCol).Address(0, 0) & ")"
        End If
           
        If toCopyFormulas Then
            .Range(.Cells(firstDataRow, 1), .Cells(firstDataRow, maxCol + 1)).Copy Destination:=.Range(.Cells(firstDataRow + 1, 1), .Cells(firstDataRow + ((maxRow - headerRow) / 4), maxCol + 1))
            .Range(.Cells(firstDataRow, 1), .Cells(firstDataRow, maxCol + 1)).Copy Destination:=.Range(.Cells(firstDataRow + ((maxRow - headerRow) / 4), 1), .Cells(firstDataRow + ((maxRow - headerRow) / 2), maxCol + 1))
            .Range(.Cells(firstDataRow, 1), .Cells(firstDataRow, maxCol + 1)).Copy Destination:=.Range(.Cells(firstDataRow + ((maxRow - headerRow) / 2), 1), .Cells(firstDataRow + ((maxRow - headerRow) * 3 / 4), maxCol + 1))
            .Range(.Cells(firstDataRow, 1), .Cells(firstDataRow, maxCol + 1)).Copy Destination:=.Range(.Cells(firstDataRow + ((maxRow - headerRow) * 3 / 4), 1), .Cells(maxRow, maxCol + 1))
        End If
        
        'countif not-zeros within columns
        If maxMatchRow = 0 Then maxMatchRow = maxRow          'imperfect...
        .Cells(maxRow + 2, 1).Formula = "=Countif(A" & firstDataRow & ":A" & maxMatchRow & ",""<>0"")"
        .Cells(maxRow + 2, 1).Copy Destination:=.Range(.Cells(maxRow + 2, 2), .Cells(maxRow + 2, maxCol + 1))
        
        'highlight deviations within compare
        Call conditFormat(.Range(.Cells(firstDataRow, 1), .Cells(maxRow + 2, maxCol + 1)), False)
        Call conditFormat(.Range(.Cells(headerRow, 1), .Cells(headerRow, maxCol + 1)), True)            'highlight field header if there are deviations in the field
    End With
            
    'format worksheets (freeze panes & autofilter)
    For Each wsh In comparedFile.Sheets
        Call freezePanes(wsh)
        Call applyAutoFilter(wsh)
    Next wsh
    
    If deletedColsCount = 0 And maxRowMaster <> maxRowTest Then createPivot
    
    ''ToDo:
    'compare headers
    'add batch import and compare
    '   Save As:
    ' 1)   comparedFile.SaveAs Left(Replace(Replace(masterFileName, "[", vbNullString), "]", vbNullString), (InStrRev(masterFileName, ".", -1, vbTextCompare) - 2)) & "_.xlsx", FileFormat:=51
    ' 2)   Dim Ret
    '    Ret = Application.GetSaveAsFilename(InitialFileName:="_passed", fileFilter:="Excel Files (*.xlsx), *.xlsx", FilterIndex:=1, Title:="Save As")

    endMacro
End Sub

Private Function chooseFile(ByVal version As String) As String
    Dim FdFp As FileDialog
    Dim userChoice As Integer
    
    Set FdFp = Application.FileDialog(msoFileDialogFilePicker)

    'choose Master file (Excel / csv / txt)
    With FdFp
        .Title = "Choose " & version & " file"
        .AllowMultiSelect = False
        With .Filters
            .Clear
            .Add "All files", "*.*"
            .Add "Excel files", "*.xls*"
            .Add "CSV files", "*.csv"
            .Add "Text files (txt)", "*.txt"
        End With
        userChoice = .Show
    End With
    If userChoice <> 0 Then
        chooseFile = FdFp.SelectedItems(1)
    Else
        endMacro
    End If
End Function

Private Sub importExcelBasedData()
    Application.ScreenUpdating = False
    
    Set masterFile = Workbooks.Open(filename:=masterFileName, ReadOnly:=True)
    masterFile.Sheets(1).Copy before:=comparedFile.Sheets(1)
    masterFile.Close False
    Set masterWsh = comparedFile.Sheets(1)
    masterWsh.Name = "Master"
    
    Set testFile = Workbooks.Open(filename:=testFileName, ReadOnly:=True)
    testFile.Sheets(1).Copy before:=comparedFile.Sheets(2)
    testFile.Close False
    Set testWsh = comparedFile.Sheets(2)
    testWsh.Name = "Test"
End Sub

Private Sub importTxtBasedData()
    Application.ScreenUpdating = False
    
    Set testWsh = comparedFile.Sheets.Add
    testWsh.Name = "Test"
    Set masterWsh = comparedFile.Sheets.Add
    masterWsh.Name = "Master"
    
    'apply text format for all cells
    masterWsh.Cells.NumberFormat = "@"
    testWsh.Cells.NumberFormat = "@"

    Call doFileQuery(masterFileName, masterWsh)
    Call doFileQuery(testFileName, testWsh)
        
    checkIfNoHeader
    
    If toSortByFullKey Then
        Call sortByFullKey(masterWsh)
        Call sortByFullKey(testWsh)
        Call textToCols(masterWsh, tabDelim, semiColDelim, commaDelim, spaceDelim, otherDelim)
        Call textToCols(testWsh, tabDelim, semiColDelim, commaDelim, spaceDelim, otherDelim)
    End If

    Application.ScreenUpdating = True
End Sub

Private Sub doFileQuery(ByVal theFileName As String, ByVal wsh As Worksheet)
    With wsh.QueryTables.Add(Connection:="TEXT;" + theFileName, Destination:=wsh.[A1])
        .Name = theFileName
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = xlWindows   '437
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        If toSortByFullKey Then                                 'do not split now. Data will be sorted by full key and split afterwards
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileOtherDelimiter = "§"
            .TextFileColumnDataTypes = Array(1, 2)
        Else                                                    'split text to columns during import
            .TextFileTabDelimiter = tabDelim
            .TextFileSemicolonDelimiter = semiColDelim
            .TextFileCommaDelimiter = commaDelim
            .TextFileSpaceDelimiter = spaceDelim
            If otherDelim <> False Then .TextFileOtherDelimiter = otherDelim
            .TextFileColumnDataTypes = Array(textToColsArrayQuery)
        End If
        .Refresh BackgroundQuery:=False
    End With
End Sub

Public Sub textToCols(ByVal wsh As Worksheet, ByVal tabDelim As Boolean, ByVal semiColDelim As Boolean, ByVal commaDelim As Boolean, ByVal spaceDelim As Boolean, ByVal otherDelim As Boolean)
    wsh.UsedRange.TextToColumns _
        Destination:=wsh.[A1], _
        DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, _
        Tab:=tabDelim, _
        Semicolon:=semiColDelim, _
        comma:=commaDelim, _
        Space:=spaceDelim, _
        Other:=otherDelim, _
        FieldInfo:=splitTextToColsArray, _
        TrailingMinusNumbers:=True
End Sub

Private Function splitTextToColsArray() As Variant '(Optional ByVal upLimit As Integer = 483)
    Dim textToColsArray As Variant
    Dim i As Integer
    
    ReDim textToColsArray(1 To 483)
    For i = 1 To 483
        textToColsArray(i) = Array(i, 2)
    Next i
    
    splitTextToColsArray = textToColsArray
End Function

Private Function textToColsArrayQuery() As Variant
    Dim textToColsArray As Variant
    Dim i As Integer
    
    ReDim textToColsArray(1 To 483)
    For i = 1 To 483
        textToColsArray(i) = xlTextFormat
    Next i
    
    textToColsArrayQuery = textToColsArray
End Function

Private Sub checkIfNoHeader()
    Dim i As Integer

    If firstDataRow = 1 Then
        masterWsh.Rows("1:1").Insert Shift:=xlDown
        testWsh.Rows("1:1").Insert Shift:=xlDown
        headerRow = 1
        masterWsh.Cells(1, 1).Value = "Some header"
        testWsh.Cells(1, 1).Value = "Some header"
        For i = 2 To 484
            masterWsh.Cells(1, i) = i
            testWsh.Cells(1, i) = i
        Next i
        firstDataRow = 2
    End If
End Sub

Public Sub processTransactionsCompare()
    Call sortTransactions(masterWsh)
    Call sortTransactions(testWsh)
    
    'add main transaction identificators
    With compareWsh
        .Cells(headerRow, maxCol + 4) = "TransNo. (Master)"
        .Cells(headerRow, maxCol + 5) = "TransNo. (Test)"
        .Cells(headerRow, maxCol + 6) = "Instrument type"
        .Cells(headerRow, maxCol + 7) = "Security ID"
        .Cells(headerRow, maxCol + 8) = "BusTransCode"
        .Cells(headerRow, maxCol + 9) = "ElemTransCode"
        .Cells(headerRow, maxCol + 10) = "Trans. Status"
        .Cells(headerRow, maxCol + 11) = "Portfolio group"
        .Cells(headerRow, maxCol + 12) = "Portfolio"
        .Cells(firstDataRow, maxCol + 4).Formula = "=Master!B" & firstDataRow
        .Cells(firstDataRow, maxCol + 5).Formula = "=Test!B" & firstDataRow
        .Cells(firstDataRow, maxCol + 6).Formula = "=Master!MP" & firstDataRow
        .Cells(firstDataRow, maxCol + 7).Formula = "=Master!K" & firstDataRow
        .Cells(firstDataRow, maxCol + 8).Formula = "=Master!G" & firstDataRow
        .Cells(firstDataRow, maxCol + 9).Formula = "=Master!I" & firstDataRow
        .Cells(firstDataRow, maxCol + 10).Formula = "=Master!W" & firstDataRow
        .Cells(firstDataRow, maxCol + 11).Formula = "=Master!N" & firstDataRow
        .Cells(firstDataRow, maxCol + 12).Formula = "=Master!O" & firstDataRow
        .Range(.Cells(firstDataRow, maxCol + 4), .Cells(firstDataRow, maxCol + 11)).Copy Destination:=.Range(.Cells(firstDataRow + 1, maxCol + 4), .Cells(maxRow, maxCol + 11))
    End With
    
    
End Sub

Public Sub processMQcompare()
    'split columns with message content
    Dim i As Integer
    Dim msgCol As Integer
    
    If toSplitMQcont Then
        i = 1
        For i = 1 To maxCol
            If Trim(LCase(masterWsh.Cells(headerRow, i))) = "message content" Then
                msgCol = i
                Exit For
            End If
        Next i
    
        Call splitMQ(masterWsh, msgCol)
        Call splitMQ(testWsh, msgCol)
    End If
    Call sortMQ(masterWsh)
    Call sortMQ(testWsh)
End Sub

Private Sub splitMQ(ByVal wsh As Worksheet, ByVal msgCol As Integer)
    Dim tempWsh As Worksheet
    Dim msgContentCols As Integer
    
    Set tempWsh = comparedFile.Sheets.Add
    wsh.Columns(msgCol).Copy Destination:=tempWsh.[A1]
    Call textToCols(tempWsh, True, False, False, False, False)
    With tempWsh
        msgContentCols = .UsedRange.Columns.Count
        .Columns("A:" & Split(Cells(, msgContentCols).Address, "$")(1)).Cut
    End With
    With wsh
        .Columns(msgCol).Insert Shift:=xlToRight
        .Columns(msgCol + msgContentCols).Delete
    End With
    With Application
        .DisplayAlerts = False
        tempWsh.Delete
        .DisplayAlerts = True
    End With
End Sub

Private Sub deleteUnnecessaryCols()
    Dim rowsWithData As Long
    Dim i As Integer
    
    Application.ScreenUpdating = False
'    progressBar.Show (vbModeless)
'    progressBar.taskName = "Processing unnecessary columns:"
'    progressBar.ProgressBar1 = 0
'    deletedColsCount = 0
    
'    maxRow = 84
'    firstDataRow = 7
'    maxCol = 483
'    Set masterWsh = ActiveSheet
'    Set testWsh = ActiveSheet
'    Set compareWsh = ActiveSheet
    
    rowsWithData = maxRow - firstDataRow + 1
    For i = maxCol To 1 Step -1
        If ((WsF.CountBlank(Range(masterWsh.Cells(firstDataRow, i), masterWsh.Cells(maxRow, i))) = rowsWithData) And _
            (WsF.CountBlank(Range(testWsh.Cells(firstDataRow, i), testWsh.Cells(maxRow, i))) = rowsWithData)) Or _
           ((WsF.CountIf(Range(masterWsh.Cells(firstDataRow, i), masterWsh.Cells(maxRow, i)), 0) = rowsWithData) And _
            (WsF.CountIf(Range(testWsh.Cells(firstDataRow, i), testWsh.Cells(maxRow, i)), 0) = rowsWithData)) Then
                masterWsh.Columns(i).Hidden = True
                testWsh.Columns(i).Hidden = True
                compareWsh.Columns(i).Hidden = True
'                progressBar.ProgressBar1 = ((maxCol - i) / maxCol) * 100
'                deletedColsCount = deletedColsCount + 1
'                progressBar.currentStatus = "Hidden " & deletedColsCount & " empty columns"
'                DoEvents
        End If
    Next i
    
'    progressBar.Hide
End Sub

Private Sub addColForExtraRow(ByVal wsh As Worksheet)
    With wsh
        .Cells(headerRow, maxCol + 1) = "Extra line"
        .Cells(firstDataRow, maxCol + 1) = 0
        .Cells(firstDataRow, maxCol + 1).Copy Destination:=.Range(.Cells(firstDataRow + 1, maxCol + 1), .Cells(.UsedRange.Rows.Count, maxCol + 1))
    End With
End Sub

Private Sub freezePanes(ByVal wsh As Worksheet)
    Application.ScreenUpdating = False
    wsh.Activate
    wsh.Cells(headerRow, 1).Select
    With ActiveWindow
        .ScrollRow = headerRow
        .SplitRow = 1
        .freezePanes = True
    End With
End Sub

Private Sub applyAutoFilter(ByVal wsh As Worksheet)
    On Error Resume Next
    With wsh
        .AutoFilterMode = False
        .Range(wsh.Cells(headerRow, 1), wsh.Cells(headerRow, maxCol + 11)).AutoFilter
    End With
    On Error GoTo 0
End Sub

Private Sub createPivot()
    Dim pivotWsh As Worksheet
    Dim masterPivot As PivotTable
    Dim testPivot As PivotTable

    With comparedFile
        Set pivotWsh = .Sheets.Add(after:=.Sheets(.Sheets.Count))
        pivotWsh.Name = "Pivot"
        Set masterPivot = .PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Master!R" & headerRow & "C1:R" & maxRowMaster & "C" & maxCol, version:=xlPivotTableVersion15).CreatePivotTable(TableDestination:="Pivot!R7C2", TableName:="PivotMaster", DefaultVersion:=xlPivotTableVersion15)
        Set testPivot = .PivotCaches.Create(SourceType:=xlDatabase, SourceData:="Test!R" & headerRow & "C1:R" & maxRowTest & "C" & maxCol, version:=xlPivotTableVersion15).CreatePivotTable(TableDestination:="Pivot!R7C8", TableName:="PivotTest", DefaultVersion:=xlPivotTableVersion15)
    End With
    
    With pivotWsh
        .[F7].Formula = "=C7=I7"
        .[F7].Copy Destination:=Range(.Cells(8, 6), .Cells(24, 6))
    End With
      
'    With masterPivot.PivotFields("Portfolio Group")
'        .Orientation = xlRowField
'        .Position = 1
'    End With
'    With masterPivot.PivotFields("Security ID")
'        .Orientation = xlRowField  '.Orientation = xlHidden
'        .Position = 2
'    End With
'    masterPivot.AddDataField masterPivot.PivotFields("BFC Code"), "Count", xlCount
End Sub

Private Sub conditFormat(ByVal formatRange As Range, ByVal condForHeader As Boolean)
    'conditional formatting
    With formatRange
        If condForHeader Then
            .FormatConditions.Add Type:=xlExpression, Formula1:="=A" & maxRow + 2 & ">0"
            .FormatConditions(.FormatConditions.Count).SetFirstPriority
        Else
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
        End If
        With .FormatConditions(1)
            With .Font
                .Color = -16383844
                .TintAndShade = 0
            End With
            With .Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
            End With
            .StopIfTrue = False
        End With
    End With
End Sub

Public Sub endMacro()
    clearPublicVars

    Application.Calculate
    Application.ScreenUpdating = True
    End
End Sub

Private Sub clearPublicVars()
    'Import module
    masterFileName = vbNullString
    Set masterWsh = Nothing
    Set testWsh = Nothing
    Set compareWsh = Nothing
    semiColDelim = False
    commaDelim = False
    tabDelim = False
    spaceDelim = False
    otherDelim = Empty
    toSortByFullKey = False
    toCopyFormulas = False
    toSplitExtraRows = False
    toSplitMQcont = False
    maxRow = 0
    maxRowMaster = 0
    maxRowTest = 0
    maxMatchRow = 0
    maxCol = 0
    firstDataRow = 0
    headerRow = 0
    
    'indepFunctions module
    veryFirstRow = 0
    veryLastRow = 0
    emptyCellCol = 0
    sortByCols = vbNullString
    
    'Sorting module
    Set targetWsh = Nothing
End Sub
