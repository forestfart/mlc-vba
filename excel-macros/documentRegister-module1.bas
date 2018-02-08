Attribute VB_Name = "Module1"
Sub Hide_Completed()
Dim index As Integer
Dim ZakresWierszy As String
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    WierszZnaleziony = 0
    'chowanie
    ZakresWierszy = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("V5:V" & ZakresWierszy)
        If cell.Value = "Completed" Then
            cell.EntireRow.Hidden = True
        End If
    Next
End Sub

Sub Unhide_Subjects()
Dim index As Integer
Dim ZakresWierszy As Integer
    ZakresWierszy = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("A8:B" & ZakresWierszy)
        If Len(cell.Value) > 2 Then
            cell.EntireRow.Hidden = False
        End If
    Next
End Sub
Sub SWMS()
Dim index As Integer
Dim index_Int As Integer
Dim Poczatek As Integer
Dim Pracownicy As Integer
Dim Znaleziony As Integer
Dim KoncowaKomorka As Integer
Dim WierszZnaleziony As String
Dim ZakresKolumn As String
Dim Zakres As String
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    WierszZnaleziony = 0
    'chowanie
    ZakresKolumn = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("E5:E" & ZakresKolumn)
        cell.EntireRow.Hidden = True
    Next
    index = 5
    For Each cell In ActiveSheet.Range("E5:E" & ZakresKolumn)
        If cell.EntireRow.Hidden = True Then
            If cell.Value = "WMS" Then
                cell.EntireRow.Hidden = False
                WierszZnaleziony = cell.Row
                'wstaw date ostatniej wersji i rev
                index_Int = WierszZnaleziony
                While ActiveSheet.Cells(index_Int, 9).Value <> ""
                    index_Int = index_Int + 1
                Wend
                ActiveSheet.Cells(WierszZnaleziony, 20).Value = ActiveSheet.Cells(index_Int - 1, 9).Value
                ActiveSheet.Cells(WierszZnaleziony, 21).Value = ActiveSheet.Cells(index_Int - 1, 10).Value
                If ActiveSheet.Cells(WierszZnaleziony, 21).Value <> "" And ActiveSheet.Cells(WierszZnaleziony, 22).Value = "Current" Then
                    If ActiveSheet.Cells(WierszZnaleziony, 24).Value <> "" Then
                        ActiveSheet.Cells(WierszZnaleziony, 23).Value = DateAdd("D", 182, ActiveSheet.Cells(WierszZnaleziony, 24).Value)
                    Else
                        ActiveSheet.Cells(WierszZnaleziony, 23).Value = DateAdd("D", 182, ActiveSheet.Cells(index_Int - 1, 10).Value)
                    End If
                Else
                    ActiveSheet.Cells(WierszZnaleziony, 23).Value = ""
                End If
                ActiveSheet.Range("S" & WierszZnaleziony & ":Y" & WierszZnaleziony).Interior.ColorIndex = 35
                If CDate(ActiveSheet.Cells(WierszZnaleziony, 23).Value) < Date And ActiveSheet.Cells(WierszZnaleziony, 22).Value = "Current" And CDate(ActiveSheet.Cells(WierszZnaleziony, 24).Value) < Date - 180 Then
                    ActiveSheet.Range("S" & WierszZnaleziony & ":Y" & WierszZnaleziony).Interior.ColorIndex = 3 '38
                End If
                If CDate(ActiveSheet.Cells(WierszZnaleziony, 23).Value) < Date + 30 And CDate(ActiveSheet.Cells(WierszZnaleziony, 23).Value) >= Date And ActiveSheet.Cells(WierszZnaleziony, 22).Value = "Current" And CDate(ActiveSheet.Cells(WierszZnaleziony, 24).Value) < Date - 150 Then
                    ActiveSheet.Range("S" & WierszZnaleziony & ":Y" & WierszZnaleziony).Interior.ColorIndex = 45 '38
                End If
                If ActiveSheet.Cells(WierszZnaleziony, 22).Value = "Completed" Then
                    ActiveSheet.Range("S" & WierszZnaleziony & ":Y" & WierszZnaleziony).Interior.ColorIndex = 15
                End If
                If ActiveSheet.Cells(WierszZnaleziony, 22).Value = "On hold" Then
                    ActiveSheet.Range("S" & WierszZnaleziony & ":Y" & WierszZnaleziony).Interior.ColorIndex = 36
                End If
                'wstaw dropdown menu
                With Range("V" & WierszZnaleziony).Validation
                    .Delete
                    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="='INFO ON CODES'!$A$35:$A$38"
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .InputTitle = ""
                    .ErrorTitle = ""
                    .InputMessage = ""
                    .ErrorMessage = ""
                    .ShowInput = True
                    .ShowError = True
                End With
                Zakres = "S" & WierszZnaleziony & ":Y" & WierszZnaleziony
                Range(Zakres).Borders(xlDiagonalDown).LineStyle = xlNone
                Range(Zakres).Borders(xlDiagonalUp).LineStyle = xlNone
                With Range(Zakres).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Range(Zakres).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Range(Zakres).Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Range(Zakres).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Range(Zakres).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                Range(Zakres).Locked = False
                'If ActiveSheet.Cells(WierszZnaleziony, 9).Value = "" Then ActiveSheet.Cells(WierszZnaleziony + 1, 6).EntireRow.Hidden = False
                'Else
                WierszZnaleziony = WierszZnaleziony + 1
                'While ActiveSheet.Cells(WierszZnaleziony, 3).Value <> "2040672" And WierszZnaleziony < 2000
                '    ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                '    WierszZnaleziony = WierszZnaleziony + 1
                'Wend
                'ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
            End If
       End If
       index = index + 1
    Next
    'Select Case Znaleziony
    '    Case Is = 0
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = "No workers found from " & Pracownicy
    '    Case Is = 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " worker found from " & Pracownicy
    '    Case Is > 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " workers found from " & Pracownicy
    'End Select
    index = 0
    'hide columns with revision and transmittal
    For Each cell In ActiveSheet.Range("I1:R1")
        cell.EntireColumn.Hidden = True
    Next
    nastepny = 0
    For Each cell In ActiveSheet.Range("V2300:V3000")
        If cell.Text = "Legend:" Then
            nastepny = 1
        End If
        If nastepny = 1 Then
            cell.EntireRow.Hidden = False
        End If
    Next
    'Unhide_Subjects
    'ActiveSheet.Protect
    ActiveWindow.LargeScroll Up:=100
    Application.ScreenUpdating = True
End Sub
Sub dim90()
    Application.ScreenUpdating = True
End Sub
Sub IdzDoDoc()
Dim index As Long
Dim Poczatek As Long
Dim Pracownicy As Long
Dim Znaleziony As Long
Dim KoncowaKomorka As Long
Dim WierszZnaleziony As String
Dim ZakresKolumn As String
    Application.ScreenUpdating = False
    WierszZnaleziony = 0
    ActiveSheet.Unprotect
    'chowanie
    'KoncowaKomorka = Poczatek - 1 + Pracownicy * 2
    'ZakresDanych = ActiveSheet.Cells(3, 4).Text & Poczatek & ":" & ActiveSheet.Cells(4, 4).Text & KoncowaKomorka
    'ZakresKolumn = ActiveSheet.Cells(3, 4).Text & Poczatek & ":" & ActiveSheet.Cells(3, 4).Text & KoncowaKomorka
    ZakresKolumn = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        cell.EntireRow.Hidden = True
    Next
    index = 5
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        If cell.EntireRow.Hidden = True Then
            If ActiveSheet.Cells(1, 6).Value = cell.Value Then
                cell.EntireRow.Hidden = False
                WierszZnaleziony = cell.Row
                'If ActiveSheet.Cells(WierszZnaleziony, 9).Value = "" Then ActiveSheet.Cells(WierszZnaleziony + 1, 6).EntireRow.Hidden = False
                'Else
                WierszZnaleziony = WierszZnaleziony + 1
                While ActiveSheet.Cells(WierszZnaleziony, 3).Value <> "2040672" And WierszZnaleziony < 10000
                    ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                    WierszZnaleziony = WierszZnaleziony + 1
                Wend
                'ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
            End If
       End If
       index = index + 1
    Next
    'Select Case Znaleziony
    '    Case Is = 0
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = "No workers found from " & Pracownicy
    '    Case Is = 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " worker found from " & Pracownicy
    '    Case Is > 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " workers found from " & Pracownicy
    'End Select
    ActiveWindow.LargeScroll Up:=100
    index = 0
    If WierszZnaleziony = 0 Then
        pokazwszystko
        MsgBox "Document not found. Click on 'Create document' button to register new document.", vbInformation, "*********########------? :("
    End If
    'unhide columns with revision and transmittal
    For Each cell In ActiveSheet.Range("I1:R1")
        cell.EntireColumn.Hidden = False
    Next
    Application.ScreenUpdating = True
End Sub
Sub pokazwszystko()
Dim WierszZnaleziony As String
Dim ZakresKolumn As String
Dim lngLast As Integer
Dim lngCounter As Integer
Dim index As Byte
Dim Kolor As Byte
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    ZakresKolumn = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("F8:F16" & ZakresKolumn)
        cell.EntireRow.Hidden = False
    Next
    'unhide columns with revision and transmittal
    For Each cell In ActiveSheet.Range("I1:R1")
        cell.EntireColumn.Hidden = False
    Next
        lngLast = Cells(Rows.Count, "I").End(xlUp).Row
    Kolor = 45 '45 to pomarañczowy
    For lngCounter = 10 To lngLast
        If Cells(lngCounter, "I").Value <> "" And Cells(lngCounter, "B").Value = "" Then
            Cells(lngCounter, "B").Interior.ColorIndex = Kolor
        Else
            Cells(lngCounter, "B").Interior.ColorIndex = 0
        End If
    Next lngCounter
    ActiveWindow.LargeScroll Up:=100
    Application.ScreenUpdating = True
End Sub
Sub WstawRevClient()
Dim index As Long
Dim Poczatek As Long
Dim Pracownicy As Long
Dim Znaleziony As Long
Dim KoncowaKomorka As Long
Dim WierszZnaleziony As String
Dim ZakresKolumn As String

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    'unhide columns with revision and transmittal
    For Each cell In ActiveSheet.Range("I1:R1")
        cell.EntireColumn.Hidden = False
    Next
    Znaleziony = 0
    'chowanie
    'KoncowaKomorka = Poczatek - 1 + Pracownicy * 2
    'ZakresDanych = ActiveSheet.Cells(3, 4).Text & Poczatek & ":" & ActiveSheet.Cells(4, 4).Text & KoncowaKomorka
    'ZakresKolumn = ActiveSheet.Cells(3, 4).Text & Poczatek & ":" & ActiveSheet.Cells(3, 4).Text & KoncowaKomorka
    ZakresKolumn = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        cell.EntireRow.Hidden = True
    Next
    index = 5
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        If cell.EntireRow.Hidden = True Then
            If ActiveSheet.Cells(1, 6).Value = cell.Value Then
                cell.EntireRow.Hidden = False
                WierszZnaleziony = cell.Row
                'If ActiveSheet.Cells(WierszZnaleziony, 9).Value = "" Then ActiveSheet.Cells(WierszZnaleziony + 1, 6).EntireRow.Hidden = False
                'Else
                While ActiveSheet.Cells(WierszZnaleziony, 9).Value <> ""
                    ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                    WierszZnaleziony = WierszZnaleziony + 1
                Wend
                
                ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                
                Rows(WierszZnaleziony & ":" & WierszZnaleziony).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                Range("I6:Q6").Select
                Selection.Copy
                'ActiveWindow.SmallScroll Down:=60
                Range("I" & WierszZnaleziony & ":Q" & WierszZnaleziony).Select
                ActiveSheet.Paste
                Range("I" & WierszZnaleziony).Select
                'Selection.Copy
            End If
       End If
       index = index + 1
    Next
    'Select Case Znaleziony
    '    Case Is = 0
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = "No workers found from " & Pracownicy
    '    Case Is = 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " worker found from " & Pracownicy
    '    Case Is > 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " workers found from " & Pracownicy
    'End Select
    index = 0
    
    'ActiveWindow.LargeScroll Up:=100
    WierszZnaleziony = 0
    Selection.Copy
    'Selection.Scroll
    Application.ScreenUpdating = True
End Sub
Sub WstawRevInternal()
Dim index As Long
Dim Poczatek As Long
Dim Pracownicy As Long
Dim Znaleziony As Long
Dim KoncowaKomorka As Long
Dim WierszZnaleziony As String
Dim ZakresKolumn As String

    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    'unhide columns with revision and transmittal
    For Each cell In ActiveSheet.Range("I1:R1")
        cell.EntireColumn.Hidden = False
    Next
    Znaleziony = 0
    'chowanie
    'KoncowaKomorka = Poczatek - 1 + Pracownicy * 2
    'ZakresDanych = ActiveSheet.Cells(3, 4).Text & Poczatek & ":" & ActiveSheet.Cells(4, 4).Text & KoncowaKomorka
    'ZakresKolumn = ActiveSheet.Cells(3, 4).Text & Poczatek & ":" & ActiveSheet.Cells(3, 4).Text & KoncowaKomorka
    ZakresKolumn = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        cell.EntireRow.Hidden = True
    Next
    index = 5
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        If cell.EntireRow.Hidden = True Then
            If ActiveSheet.Cells(1, 6).Value = cell.Value Then
                cell.EntireRow.Hidden = False
                WierszZnaleziony = cell.Row
                'If ActiveSheet.Cells(WierszZnaleziony, 9).Value = "" Then ActiveSheet.Cells(WierszZnaleziony + 1, 6).EntireRow.Hidden = False
                'Else
                While ActiveSheet.Cells(WierszZnaleziony, 9).Value <> ""
                    ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                    WierszZnaleziony = WierszZnaleziony + 1
                Wend
                ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                
                Rows(WierszZnaleziony & ":" & WierszZnaleziony).Select
                Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                Range("I7:Q7").Select
                Selection.Copy
                'ActiveWindow.SmallScroll Down:=60
                Range("I" & WierszZnaleziony & ":Q" & WierszZnaleziony).Select
                ActiveSheet.Paste
                Range("I" & WierszZnaleziony).Select
                'Selection.Copy
            End If
       End If
       index = index + 1
    Next
    'Select Case Znaleziony
    '    Case Is = 0
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = "No workers found from " & Pracownicy
    '    Case Is = 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " worker found from " & Pracownicy
    '    Case Is > 1
    '        ActiveSheet.Cells(Poczatek - 2, 2).Value = Znaleziony & " workers found from " & Pracownicy
    'End Select
    index = 0
    WierszZnaleziony = 0
    'ActiveWindow.LargeScroll Up:=100
    Selection.Copy
    Application.ScreenUpdating = True
End Sub
Sub NowyDoc_Klikniecie()
ActiveSheet.Unprotect
UserDoc.Show
End Sub
Public Function ColumnLetter(Column As Integer) As String
    If Column < 1 Then Exit Function
    ColumnLetter = ColumnLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
End Function
Sub action(wybor As Integer)
Dim kolumna As Integer
Dim kolumna_sd As Integer
Dim kolumna_koniec As Integer
Dim zakres_int As String
Dim zakres_tsm As String
Dim zakres_sd As String
    Application.ScreenUpdating = False
    kolumna = 6
    Do While ActiveSheet.Cells(2, kolumna).Value <> "Rev."
        kolumna = kolumna + 1
    Loop
    kolumna_sd = kolumna
    Do While ActiveSheet.Cells(2, kolumna_sd).Value <> "Comments"
        kolumna_sd = kolumna_sd + 1
    Loop
    kolumna_koniec = kolumna_sd + 1
    Do While ActiveSheet.Cells(2, kolumna_koniec).Value <> "Comments"
        kolumna_koniec = kolumna_koniec + 1
    Loop
    zakres_int = "H:" & ColumnLetter(kolumna - 1)
    zakres_tsm = ColumnLetter(kolumna) & ":" & ColumnLetter(kolumna_sd)
    zakres_sd = ColumnLetter(kolumna_sd + 1) & ":" & ColumnLetter(kolumna_koniec)
    kolumna_dzis = 6
    Do While ActiveSheet.Cells(2, kolumna_dzis).Value <> ""
        kolumna_dzis = kolumna_dzis + 1
    Loop
    zakres_intern_dzis = "H:" & ColumnLetter(kolumna_dzis - 2)

    Select Case wybor
        Case 1
            Columns(zakres_int).Select
            Selection.EntireColumn.Hidden = False
            Columns(zakres_tsm).Select
            Selection.EntireColumn.Hidden = True
            Columns(zakres_sd).Select
            Selection.EntireColumn.Hidden = True
        Case 2
            Columns(zakres_int).Select
            Selection.EntireColumn.Hidden = True
            Columns(zakres_tsm).Select
            Selection.EntireColumn.Hidden = False
            Columns(zakres_sd).Select
            Selection.EntireColumn.Hidden = True
        Case 3
            Columns(zakres_int).Select
            Selection.EntireColumn.Hidden = True
            Columns(zakres_tsm).Select
            Selection.EntireColumn.Hidden = True
            Columns(zakres_sd).Select
            Selection.EntireColumn.Hidden = False
        Case 4
            Columns(zakres_int).Select
            Selection.EntireColumn.Hidden = False
            Columns(zakres_tsm).Select
            Selection.EntireColumn.Hidden = True
            Columns(zakres_sd).Select
            Selection.EntireColumn.Hidden = True
            Columns(zakres_intern_dzis).Select
            Selection.EntireColumn.Hidden = True
    End Select
    ActiveSheet.Range("A1:D1").Select
    Application.ScreenUpdating = True
End Sub
Sub pokaz_internal()
    action (1)
End Sub
Sub pokaz_transmittals()
    action (2)
End Sub
Sub pokaz_sd()
    action (3)
End Sub
Sub pokaz_intern_dzis()
    action (4)
End Sub
