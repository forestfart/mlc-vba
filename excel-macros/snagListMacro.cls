VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Dim keyCells As Range
Dim tmpCell As Range
Dim snagReg As Workbook
Dim index As Byte
    index = 0
    Set keyCells = Range("J1:W1")
    If Not Intersect(Target, keyCells) Is Nothing Then
        Set snagReg = Workbooks.Open("\\Mlc-dc1\mlc\DOC CTRL\REGISTERS\REG-62 = SNAG LIST.xlsx")
        Application.ScreenUpdating = False
        'Sheet1.Unprotect
        snagZakresWierszy = snagReg.Sheets(3).UsedRange.Rows.Count
        For wiersz = 3 To snagZakresWierszy
            If snagReg.Sheets(3).Cells(wiersz, 3).Value = ThisWorkbook.Sheets(1).Cells(3, 35).Value And snagReg.Sheets(3).Cells(wiersz, 4).Value = ThisWorkbook.Sheets(1).Cells(1, 20).Value Then
                'MsgBox (snagReg.Sheets(3).Cells(wiersz, 4).Value)
                ThisWorkbook.Sheets(1).Cells(20 + index, 1).Value = snagReg.Sheets(3).Cells(wiersz, 1).Value
                ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value = snagReg.Sheets(3).Cells(wiersz, 8).Value
                ThisWorkbook.Sheets(1).Cells(20 + index, 20).Value = snagReg.Sheets(3).Cells(wiersz, 6).Value
                ThisWorkbook.Sheets(1).Cells(20 + index, 16).Value = snagReg.Sheets(3).Cells(wiersz, 5).Value
                index = index + 1
                If ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value <> "" Then
                    If Len(ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value) > 80 Then
                        ThisWorkbook.Sheets(1).Cells(20 + index, 2).EntireRow.RowHeight = 75
                    ElseIf Len(ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value) > 60 Then
                        ThisWorkbook.Sheets(1).Cells(20 + index, 2).EntireRow.RowHeight = 60
                    ElseIf Len(ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value) > 40 Then
                        ThisWorkbook.Sheets(1).Cells(20 + index, 2).EntireRow.RowHeight = 45
                    ElseIf Len(ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value) > 20 Then
                        ThisWorkbook.Sheets(1).Cells(20 + index, 2).EntireRow.RowHeight = 30
                    Else
                        ThisWorkbook.Sheets(1).Cells(20 + index, 2).EntireRow.RowHeight = 15
                    End If
                End If
            End If
        Next
        While index < 20
            ThisWorkbook.Sheets(1).Cells(20 + index, 2).Value = ""
            ThisWorkbook.Sheets(1).Cells(20 + index, 20).Value = ""
            ThisWorkbook.Sheets(1).Cells(20 + index, 1).Value = ""
            ThisWorkbook.Sheets(1).Cells(20 + index, 16).Value = ""
            ThisWorkbook.Sheets(1).Cells(20 + index, 2).EntireRow.RowHeight = 15
            index = index + 1
        Wend
        snagReg.Close
        Application.ScreenUpdating = True
    End If
End Sub

