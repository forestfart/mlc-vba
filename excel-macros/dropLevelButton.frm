VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MLCForm1 
   Caption         =   "Input Data"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3750
   OleObjectBlob   =   "dropLevelButton.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MLCForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ApplyButt_Click()
   
    'sprawdzanie bledow
    If Drop.Value = "" Then
        MsgBox "Type in long drop number", vbCritical, "Wrong"
        End
    End If
    If Level.Value = "" Then
        MsgBox "Document type not selected.", vbCritical, "Wrong level"
        End
    End If
    If SPR.Value = "Title" Then
        MsgBox "", vbCritical, "Wrong ScanPrint Report number"
        End
    End If
    Select Case Drop.Value
        Case Is = 2
            Sheet11.Name = "COLUMN 1-2"
            Sheet12.Name = "COLUMN 2-3"
            Sheet11.Cells(9, 26).Value = Level.Value
            Sheet11.Cells(2, 26).Value = Level.Value + 1
            Sheet12.Cells(9, 25).Value = Level.Value
            Sheet12.Cells(2, 25).Value = Level.Value + 1
            Sheet19.Cells(6, 3).Value = 2
            Sheet19.Cells(6, 4).Value = Level.Value
            Sheet19.Cells(6, 5).Value = SPR.Value
        Case Is = 4
            Sheet11.Name = "COLUMN 3-4"
            Sheet12.Name = "COLUMN 4-5"
            Sheet11.Cells(9, 26).Value = Level.Value
            Sheet11.Cells(2, 26).Value = Level.Value + 1
            Sheet12.Cells(9, 25).Value = Level.Value
            Sheet12.Cells(2, 25).Value = Level.Value + 1
            Sheet19.Cells(6, 3).Value = 4
            Sheet19.Cells(6, 4).Value = Level.Value
            Sheet19.Cells(6, 5).Value = SPR.Value
        Case Is = 6
            Sheet11.Name = "COLUMN 5-6"
            Sheet12.Name = "COLUMN 6-7"
            Sheet11.Cells(9, 26).Value = Level.Value
            Sheet11.Cells(2, 26).Value = Level.Value + 1
            Sheet12.Cells(9, 25).Value = Level.Value
            Sheet12.Cells(2, 25).Value = Level.Value + 1
            Sheet19.Cells(6, 3).Value = 6
            Sheet19.Cells(6, 4).Value = Level.Value
            Sheet19.Cells(6, 5).Value = SPR.Value
        Case Is = 8
            Sheet11.Name = "COLUMN 7-8"
            Sheet12.Name = "COLUMN 8-1"
            Sheet11.Cells(9, 26).Value = Level.Value
            Sheet11.Cells(2, 26).Value = Level.Value + 1
            Sheet12.Cells(9, 25).Value = Level.Value
            Sheet12.Cells(2, 25).Value = Level.Value + 1
            Sheet19.Cells(6, 3).Value = 8
            Sheet19.Cells(6, 4).Value = Level.Value
            Sheet19.Cells(6, 5).Value = SPR.Value
    End Select
    Sheet19.Shapes.Range(Array("Button 3")).Select
    Selection.Delete
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Shop Drowings locked at:"
    Range("C5:E6").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range("F8").Select
    Unload Me
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
