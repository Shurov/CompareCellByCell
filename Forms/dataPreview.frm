VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dataPreview 
   Caption         =   "File preview"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12390
   OleObjectBlob   =   "dataPreview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dataPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private excelBased As Boolean

Private Sub UserForm_Activate()
    Dim fileExtention As String
    
    sortFullKey.Enabled = True
    sortFullKey.Value = False
    
    delimComma.Value = False
    delimSemicol.Value = False
    delimTab.Value = False
    delimSpace.Value = False
    delimOther.Value = False
    textDelim.Value = vbNullString
    
    fileExtention = LCase(Right(masterFileName, 4))
    
    'for Excel-based (non-CSV) files disable delimeter choice
    If InStr(fileExtention, "xls") > 0 Or fileExtention = ".pdf" Or fileExtention = ".rpt" Then excelBased = True
    
    If excelBased Then
        lblDelim.Enabled = False
        delimComma.Enabled = False
        delimSemicol.Enabled = False
        delimTab.Enabled = False
        delimSpace.Enabled = False
        delimOther.Enabled = False
        textDelim.Enabled = False
        textDelim.Value = vbNullString
    Else
        lblDelim.Enabled = True
        delimComma.Enabled = True
        delimSemicol.Enabled = True
        delimTab.Enabled = True
        delimSpace.Enabled = True
        delimOther.Enabled = True
        textDelim.Enabled = True
    End If
    
    'load lines for preview
    lstBoxPreview.Clear
    If Not excelBased Then
        lstBoxPreview.ListStyle = fmListStyleOption

        Open masterFileName For Input As #99
        Do Until EOF(99)
            Line Input #99, someLine
            lstBoxPreview.AddItem (Left(someLine, 1024))
            If lstBoxPreview.ListCount = 13 Then Exit Do
        Loop
        Close #99
    Else
        lstBoxPreview.ListStyle = fmListStylePlain
        lstBoxPreview.AddItem ("No preview is availible. Please look at the import file manually")
    End If
        
End Sub

Private Sub lstBoxPreview_Click()
    text1stRow.Value = lstBoxPreview.ListIndex + 1
End Sub

Private Sub delimComma_Click()
    sortFullKey.Enabled = True
End Sub

Private Sub delimSemicol_Click()
    sortFullKey.Enabled = True
End Sub

Private Sub delimTab_Click()
    sortFullKey.Enabled = False
    sortFullKey.Value = False
End Sub

Private Sub delimSpace_Click()
    sortFullKey.Enabled = True
End Sub

Private Sub delimOther_Change()
    sortFullKey.Enabled = True
    If delimOther.Value = True Then
        textDelim.Enabled = True
        textDelim.BackColor = -2147483643
    Else
        textDelim.Enabled = False
        textDelim.BackColor = -2147483633
    End If
End Sub

Private Sub cmdButGo_Click()
    If text1stRow.Value = vbNullString Then
        MsgBox "Please set first row with data", vbOKOnly + vbCritical, "Error!"
        Exit Sub
    Else
        firstDataRow = CInt(text1stRow.Value)
        headerRow = firstDataRow - 1
    End If
    
    If delimSemicol.Value = False And delimComma.Value = False And delimTab.Value = False And delimOther.Value = False And delimSpace.Value = False And excelBased = False Then
        MsgBox "Please set the delimeter", vbOKOnly + vbCritical, "Error!"
        Exit Sub
    Else
        semiColDelim = CBool(delimSemicol.Value)
        commaDelim = CBool(delimComma.Value)
        tabDelim = CBool(delimTab.Value)
        spaceDelim = CBool(delimSpace.Value)
        If delimOther.Value = True Then
            otherDelim = CBool(textDelim.Value)
        Else
            otherDelim = False
        End If
    End If
    toCopyFormulas = CBool(copyFormulasDown.Value)
    toSplitExtraRows = CBool(splitExtraRows.Value)
    toSplitMQcont = CBool(splitMQmsgCont.Value)
    toSortByFullKey = CBool(sortFullKey.Value)

    Me.Hide
End Sub

Private Sub UserForm_Terminate()
    endMacro
End Sub
