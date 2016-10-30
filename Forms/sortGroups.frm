VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} sortGroups 
   Caption         =   "Sort within groups"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3735
   OleObjectBlob   =   "sortGroups.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "sortGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub UserForm_Initialize()
    If Not (sheetExists("Master") And sheetExists("Test")) Then
        MsgBox "Your workbook doesn't contain Master and/or Test worksheets. Please make sure you have sheets named ""Master"" and ""Test""", vbOKOnly + vbCritical
        End
    End If
End Sub

Private Sub sort_Click()
    If startRow.Value = vbNullString Or endRow.Value = vbNullString Or emptyCell.Value = vbNullString Or sortBy.Value = vbNullString Then
        MsgBox "You should fill all the fields!", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    veryFirstRow = startRow.Value       'start of data range to sort
    veryLastRow = endRow.Value          'end of data range to sort
    emptyCellCol = emptyCell.Value      'column number where empty cell delimits groups
    sortByCols = sortBy.Value           'sort fields separated by comma
    
    Application.ScreenUpdating = False
    
    Call sortWithinGroups(Worksheets("Master"))
    Call sortWithinGroups(Worksheets("Test"))
    
    Unload Me
    Application.ScreenUpdating = True
End Sub
