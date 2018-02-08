VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub transmittals_macro()
    Columns("E:CW").Select
    Range("E4").Activate
    Selection.EntireColumn.Hidden = True
End Sub
Sub Internal_macro()

End Sub
Sub shopDrawings_macro()

End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim index As Integer
Dim ZakresWierszy As Integer
Dim kolumnaDaty As Integer
    Application.ScreenUpdating = False
    'znajdz numer kolumny
    For index = 10 To ActiveSheet.UsedRange.Columns.Count
        If ActiveSheet.Cells(2, index).Value = "Date" Then
           kolumnaDaty = index
           Exit For
        End If
    Next
    'ActiveSheet.Cells(17, 100).Value = kolumnaDaty
    ZakresWierszy = ActiveSheet.UsedRange.Rows.Count
    For index = 3 To ZakresWierszy
        If ActiveSheet.Cells(index, kolumnaDaty).Value = "INTERNAL" Then
            ActiveSheet.Cells(index, kolumnaDaty + 1).Value = "---------"
        End If
    Next
End Sub
