Attribute VB_Name = "Module1"
Function outputNumber(inputLetter As String) As Byte
    
Dim Leng As Integer
Dim i As Integer

    'inputLetter = InputBox("The Converting letter?")  ' Input the Column Letter
    Leng = Len(inputLetter)
    outputNumber = 0
    
    For i = 1 To Leng
       outputNumber = (Asc(UCase(Mid(inputLetter, i, 1))) - 64) + outputNumber * 26
    Next i
    
    'MsgBox outputNumber   'Output the corresponding number
    
End Function

Sub ImportAssessment()
Dim objExcel As New Excel.Application
Dim wb As Excel.Workbook
Dim fd As FileDialog
Set fd = Application.FileDialog(msoFileDialogFilePicker)
Dim plikWybrany As Integer
Dim plikSciezka As String
Dim assessmentColumn As Byte
Dim assessmentColumnAlpha As String
Dim importColEx As String
Dim importColNumEx As Byte
Dim tabela, ndiTabela As Byte
Dim wiersz, ndiWiersz, i As Byte
Dim heading As range
Dim iReply As Integer
Dim szukanyDefekt, szukanyNdiDefekt As String
Dim szukanaOdpowiedz As String
Dim wierszEx As Integer
Dim kolumnaEx As Byte
Dim szukany As range
Dim wierszPoczatek As Byte
Dim zmianaTekst1, zmianaTekst2, NaglowekTekst, ndiNaglowekTekst As String
    
    iReply = MsgBox(prompt:="Is it an old contract report (Contract 1)?", Buttons:=vbYesNoCancel, Title:="Which contract?")
    '==================================================== old scope macro ================================
    If iReply = vbYes Then
        Application.ScreenUpdating = False
        plikWybrany = fd.Show
        If plikWybrany <> -1 Then
            'jak nic nie wyntane to cancel
            MsgBox " you chose cancel"
        Else
            plikSciezka = fd.SelectedItems(1)
            'Declare object variables.
            Dim appXl As Excel.Application
            Dim wrkFile As Workbooks
            'Set object variables.
            Set appXl = New Excel.Application
            Set wrkFile = appXl.Workbooks
            'Open a file.
            wrkFile.Open plikSciezka, ReadOnly
            'Display Excel.
            appXl.Visible = True
            MsgBox "At this point Excel is open and displays a document." & Chr$(13) & "The following statements will close the document and then close Excel."
            
            importColEx = InputBox("Which column to import?", "Column selection")
            If importColEx <> "" Then
                importColNumEx = outputNumber(importColEx)
                'Selection.TypeText wrkFile.Application.Worksheets(1).Cells(1, 1)
                ' tutaj by trzeba bylo wstawic kod czytajacy excela i zapisujacy go w odpowiednich tabelach worda
                'Close the file.
                'assessmentColumnAlpha = InputBox("Which column contain import data? (alphabetical)")
                'podjezdzamy na poczatek
                kolumnaEx = 3
                wierszPoczatek = 1
                While wrkFile.Application.Worksheets(1).Cells(kolumnaEx, wierszPoczatek).Value <> "Panel"
                    wierszPoczatek = wierszPoczatek + 1
                Wend
                wierszPoczatek = wierszPoczatek + 1
                For tabela = 1 To ActiveDocument.Tables.Count
                    If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
                        ThisDocument.Tables.Item(tabela).Cell(1, 1).range.Select
                        Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                        'Display heading text
                        While heading.Style <> "Heading 6"
                            Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                        Wend
                        heading.Expand Unit:=wdParagraph
                        NaglowekTekst = Left(heading.text, Len(heading.text) - 1)
                        'MsgBox heading.Text
                        
                        For wiersz = 2 To ActiveDocument.Tables(tabela).Rows.Count
                            szukanyDefekt = ThisDocument.Tables.Item(tabela).Cell(wiersz, 1).range.text
                            szukanyDefekt = Left(szukanyDefekt, Len(szukanyDefekt) - 2)
                            'podjezdzamy na poczatek
                            wierszEx = wierszPoczatek
                            While wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx).Value <> NaglowekTekst And wierszEx < 500 'And wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx + 1).Value <> szukanyDefekt
                                'MsgBox wierszEx & " = " & zmianaTekst1 & " = " & naglowekTekst & " wiersz z worda: " & wiersz
                                wierszEx = wierszEx + 1
                                zmianaTekst1 = wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx).Value 'just for msgbox
                            Wend
                            While wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx + 1).Value <> szukanyDefekt And wierszEx < 500
                                wierszEx = wierszEx + 1
                            Wend
                            If wrkFile.Application.Worksheets(1).Cells(wierszEx, importColNumEx).Value = "V02 - Asbestos in column joint" Then wierszEx = wierszEx + 1
                            If wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx).Value = NaglowekTekst Then
                                    szukanaOdpowiedz = wrkFile.Application.Worksheets(1).Cells(wierszEx, importColNumEx).Value
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Italic = False
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Color = wdBlack
                                Else:
                                    szukanaOdpowiedz = "not listed"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Italic = True
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Color = wdColorGray60
                            End If
         
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Size = 6
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).range.Font.Size = 7 'properties column
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.text = szukanaOdpowiedz
                            
                            'MsgBox wierszEx & " = " & zmianaTekst1 & " = " & naglowekTekst & " wiersz z worda: " & wiersz
                            
                            'MsgBox wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx).Value & " " & heading.Text & " " & StrComp(zmianaTekst1, heading.Text), , "panel z excela in naglowek z worda"
                            
                            'MsgBox wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx + 1).Value & "....." & szukanyDefekt & " " & StrComp(wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx + 1).Value, szukanyDefekt), , "defekt z excela i defekt z worda"
                            'MsgBox Len(wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx + 1).Value) & "...." & Len(szukanyDefekt) & StrComp(wrkFile.Application.Worksheets(1).Cells(wierszEx, kolumnaEx + 1).Value, szukanyDefekt)
                        Next wiersz
                    End If
                Next tabela
            End If
        End If
        Application.DisplayAlerts = False
        'wrkFile.Close
        'Quit Excel.
        'appXl.Quit
        Application.DisplayAlerts = True
        'Close the object references.
        Set wrkFile = Nothing
        Set appXl = Nothing
    End If
    '======================================================= new scope macro ===================================
    If iReply = vbNo Then
    ' macro new scope
        'Application.ScreenUpdating = False
        plikWybrany = fd.Show
        If plikWybrany <> -1 Then
            'jak nic nie wyntane to cancel
            MsgBox " you chose cancel"
        Else
            plikSciezka = fd.SelectedItems(1)
            Dim wrdFile As Document
            Set ndiDocument = Documents.Open(plikSciezka)
            ndiDocument.Activate
            
            If ThisDocument.Tables.Count <> ndiDocument.Tables.Count Then
            '    MsgBox "The file:      " & plikSciezka & "          is having differant number of panels incorrect."
            End If
 
            For tabela = 1 To ThisDocument.Tables.Count
                    If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
                        ThisDocument.Tables.Item(tabela).Cell(1, 1).range.Select
                        Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                        'heading text
                        While heading.Style <> "Heading 6"
                            Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                        Wend
                        heading.Expand Unit:=wdParagraph
                        NaglowekTekst = Left(heading.text, Len(heading.text) - 1)
                        'MsgBox heading.text
                        ndiTabela = 1
                        'MsgBox (ndiDocument.Tables.Count)
                        While NaglowekTekst <> ndiNaglowekTekst And ndiTabela < ndiDocument.Tables.Count ' dodane 03/02/2017 " And ndiTabela < ndiDocument.Tables.Count"
                            ndiTabela = ndiTabela + 1
                            If ndiDocument.Tables.Item(ndiTabela).Cell(1, 1).range.text Like "*Code*" Then
                                ndiDocument.Tables.Item(ndiTabela).Cell(1, 1).range.Select
                                Set ndiHeading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                                'heading text
                                While ndiHeading.Style <> "Heading 6"
                                    Set ndiHeading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
                                Wend
                                ndiHeading.Expand Unit:=wdParagraph
                                ndiNaglowekTekst = Left(ndiHeading.text, Len(ndiHeading.text) - 1)
                            End If
                        Wend
                        'MsgBox naglowekTekst & "  vs.  " & ndiNaglowekTekst
                        For wiersz = 2 To ThisDocument.Tables(tabela).Rows.Count
                            szukanyDefekt = ThisDocument.Tables.Item(tabela).Cell(wiersz, 1).range.text
                            szukanyDefekt = Left(szukanyDefekt, Len(szukanyDefekt) - 2)
                            szukanyNdiDefekt = ""
                            'podjezdzamy na poczatek - tutaj zaczynamy szukac defektu w ndi report
                            'tutaj szukamy wiersza odpowiadajacego
                            ndiWiersz = 1
                            While szukanyDefekt <> szukanyNdiDefekt And ndiWiersz <= ndiDocument.Tables(ndiTabela).Rows.Count
                                ndiWiersz = ndiWiersz + 1
                                szukanyNdiDefekt = ndiDocument.Tables.Item(ndiTabela).Cell(ndiWiersz, 1).range.text
                                szukanyNdiDefekt = Left(szukanyNdiDefekt, Len(szukanyNdiDefekt) - 2)
                                If szukanyDefekt = szukanyNdiDefekt Then
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.text = ndiDocument.Tables.Item(ndiTabela).Cell(ndiWiersz, 8).range.text 'przypisanie wlasciwe
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Italic = False
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Color = wdBlack
                                End If
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Size = 6
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).range.Font.Size = 7 'properties column
                                If ndiWiersz = ndiDocument.Tables(ndiTabela).Rows.Count And szukanyDefekt <> szukanyNdiDefekt Then
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.text = "not listed"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Italic = True
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 3).range.Font.Color = wdColorGray60
                                End If
                                
                            Wend
                            'szukanyNdiDefekt = ""
                        Next wiersz
                    End If
            Next tabela
            
            'MsgBox ThisDocument.Tables.Count & " vs. " & ndiDocument.Tables.Count
            
            ndiDocument.Close (True)
        End If
    End If
    'zmiana V02 -> C11/C12
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "V02 - Asbestos"
        .Replacement.text = "Concrete repair to column panel edge"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Application.ScreenUpdating = True
End Sub

Sub nazwijTabele()
Dim tabela As Byte
Dim wiersz As Byte
Dim heading As range
    'Application.ScreenUpdating = False
    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
            ThisDocument.Tables.Item(tabela).Cell(1, 1).range.Select
            Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            'Display heading text
            While heading.Style <> "Heading 6"
                Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
            Wend
            heading.Expand Unit:=wdParagraph
            MsgBox heading.text
        End If
    Next tabela
    'Application.ScreenUpdating = True
End Sub

Sub CountConditions()
Dim tabela As Byte
Dim wiersz As Byte
Dim newEntries As Integer
Dim notRepaired As Integer
Dim repaired As Integer
    Application.ScreenUpdating = False
    newEntries = 0
    notRepaired = 0
    repaired = 0
    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
            For wiersz = 2 To ActiveDocument.Tables(tabela).Rows.Count
                If ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text Like "*New entry*" Then
                    newEntries = newEntries + 1
                End If
                If ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text Like "*Repaired*" Then
                    repaired = repaired + 1
                End If
                If ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text Like "*Not repaired*" Then
                    notRepaired = notRepaired + 1
                End If
            Next wiersz
        End If
    Next tabela
    Application.ScreenUpdating = True
    'insert results
    For tabela = 1 To ActiveDocument.Tables(2).Tables.Count
        If ThisDocument.Tables(2).Tables.Item(tabela).Cell(1, 1).range.text Like "*Repaired*" Then
            ActiveDocument.Tables(2).Tables(tabela).Cell(1, 2).range.text = repaired
            ActiveDocument.Tables(2).Tables(tabela).Cell(2, 2).range.text = notRepaired
        End If
    Next tabela
End Sub

Sub HighlightRows()
Dim tabela As Byte
Dim wiersz As Byte
    Application.ScreenUpdating = False
    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
            For wiersz = 2 To ActiveDocument.Tables(tabela).Rows.Count
                If ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text Like "*Repaired*" Then
                    ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorGray15
                    'ThisDocument.Tables.Item(tabela).Rows(wiersz).Select
                    'Selection.Font.ColorIndex = wdAuto
                ElseIf ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text Like "*New entry*" Then
                    ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorLightGreen
                Else:
                    ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorWhite
                End If
            Next wiersz
        End If
    Next tabela
    Application.ScreenUpdating = True
End Sub
Sub ChangeStatusWording()
Dim tabela As Byte
Dim wiersz As Byte
    Application.ScreenUpdating = False
    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
            For wiersz = 2 To ActiveDocument.Tables(tabela).Rows.Count
                If ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text Like "*Restated*" Then
                    If Not ActiveDocument.Tables(tabela).Cell(wiersz, 2).range.text Like "*patch*" Then
                        ActiveDocument.Tables(tabela).Cell(wiersz, 5).range.text = "Not repaired"
                    End If
                Else:
                    'ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorWhite
                End If
            Next wiersz
        End If
    Next tabela
    Application.ScreenUpdating = True
End Sub
Sub lockAspectRatio()
' Sets all selected shapes to Locked Aspect Ratio
Dim i As Integer
    For i = 1 To ActiveDocument.InlineShapes.Count
        ActiveDocument.InlineShapes(i).lockAspectRatio = msoTrue
    Next
End Sub
Sub adjustColumns()
Dim tabela As Byte
    'Application.ScreenUpdating = False
    'ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1.7)

    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
            ActiveDocument.Tables.Item(tabela).Rows.SetLeftIndent LeftIndent:=-10, RulerStyle:=wdAdjustFirstColumn  'przesun cala tabele w lewo 10 jednostek
        End If
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).range.text Like "*Code*" Then
            ActiveDocument.Tables.Item(tabela).Columns(1).Width = 40  'code
            ActiveDocument.Tables.Item(tabela).Columns(2).Width = 110 'condition
            ActiveDocument.Tables.Item(tabela).Columns(3).Width = 80  'assessment
            ActiveDocument.Tables.Item(tabela).Columns(4).Width = 40  'enetered on
            ActiveDocument.Tables.Item(tabela).Columns(5).Width = 50  'status
            ActiveDocument.Tables.Item(tabela).Columns(6).Width = 90  'properties
            If ActiveDocument.Tables.Item(tabela).Cell(7, 1).range.text = Empty Then 'mozna sie pozbyc jak to juz bedzie 100% dzialalo
                ActiveDocument.Tables.Item(tabela).Columns(7).Width = 110 'treatment/comments
            End If
        End If
    Next tabela
    Application.ScreenUpdating = True
    Application.ScreenRefresh
End Sub
Sub TblCellPadding()
Dim myCell As Cell
Dim myRow As Row
Dim myTable As Table
Dim tabele As Byte
    Application.ScreenUpdating = False
    For tabele = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabele).Cell(1, 1).range.text Like "*Code*" Then
            Set myTable = ThisDocument.Tables.Item(tabele)
            myTable.Rows(1).Select
            Selection.Rows.HeadingFormat = True
            myTable.Rows.AllowBreakAcrossPages = False
            For Each myRow In myTable.Rows
                For Each myCell In myRow.Cells
                'myCell.TopPadding = CentimetersToPoints(0)
                'myCell.BottomPadding = CentimetersToPoints(0)
                myCell.LeftPadding = CentimetersToPoints(0)
                myCell.RightPadding = CentimetersToPoints(0.15)
                Next
            Next
        End If
    Next
    Application.ScreenUpdating = True
End Sub
Sub Draw_the_level()
Dim mark_up_level As Byte
Dim vOffset, hOffset As Integer
Dim Drop As Byte
Dim index As Integer
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    vOffset = 188
    hOffset = 87 'how far from the top drawing start
    Drop = InputBox("What is the drop number?", Drop)
    If Drop <> 2 And Drop <> 4 And Drop <> 6 And Drop <> 8 Then
        MsgBox ("Wrong drop number, please try again.")
        Exit Sub
    End If
    mark_up_level = InputBox("What is the report level number?", Level)
    If mark_up_level > 68 Or mark_up_level < 7 Then
        MsgBox ("Wrong level number, please try again.")
        Exit Sub
    End If
    '--------------
    Level = 248.825 * 2 - 27.127 * 2
    'add calibration line
    'With ThisDocument.Shapes.AddLine(hOffset, vOffset, hOffset + 50 * 2, vOffset).Line
    '    .Visible = msoTrue
    '    .Weight = 0.55
    '    .DashStyle = msoLineSolid
    '    .ForeColor.RGB = RGB(255, 0, 0)
    'End With
    '--------------
    For index = 7 To mark_up_level
        Height = 3.42 * 2
        If index = 7 Then Height = 6.516 * 2
        If index = 26 Or index = 54 Then Height = 3.42 * 2 + 3.08 * 2
        If index = 68 Then Height = 4 * 2
        Level = Level - Height
    Next index
    If Drop = 8 Or Drop = 4 Then drop_draw = hOffset
    If Drop = 6 Or Drop = 2 Then drop_draw = hOffset + 26 * 2
    With ThisDocument.Shapes.AddShape(msoShapeRectangle, drop_draw, Level + vOffset, 26 * 2, Height)
        .Fill.Visible = msoTrue
        With .Line
            .Weight = 0.55
            .DashStyle = msoLineSquareDot
            .Style = msoLineSingle
            .Transparency = 0#
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .BackColor.RGB = RGB(255, 0, 0)
        End With
        .Fill.ForeColor.RGB = RGB(200, 110, 110)
        .Fill.BackColor.RGB = RGB(255, 170, 170)
        .Fill.Transparency = 0.34
    End With
    'draw arrow
    'With ThisDocument.Shapes.AddLine(drop_draw + 26, Level + vOffset + Height / 2, 150, 100 + vOffset).Line
    '    .DashStyle = msoLineDashDotDot
    '    .ForeColor.RGB = RGB(50, 0, 128)
    '    .EndArrowheadStyle = msoArrowheadOpen
    'End With
    ThisDocument.Tables.Item(3).Cell(2, 2).range.text = "Level " & mark_up_level + 1
    ThisDocument.Tables.Item(3).Cell(3, 2).range.text = "Level " & mark_up_level
    If Drop = 2 Then
        ThisDocument.Tables.Item(3).Cell(1, 1).range.text = "Drop 1"
        ThisDocument.Tables.Item(3).Cell(1, 2).range.text = "Column 1/2"
        ThisDocument.Tables.Item(3).Cell(1, 3).range.text = "Drop 2"
        ThisDocument.Tables.Item(3).Cell(1, 4).range.text = "Column 2/3"
        ThisDocument.Tables.Item(3).Cell(1, 5).range.text = "Drop 3"
    End If
    If Drop = 4 Then
        ThisDocument.Tables.Item(3).Cell(1, 1).range.text = "Drop 3"
        ThisDocument.Tables.Item(3).Cell(1, 2).range.text = "Column 3/4"
        ThisDocument.Tables.Item(3).Cell(1, 3).range.text = "Drop 4"
        ThisDocument.Tables.Item(3).Cell(1, 4).range.text = "Column 4/5"
        ThisDocument.Tables.Item(3).Cell(1, 5).range.text = "Drop 5"
    End If
    If Drop = 6 Then
        ThisDocument.Tables.Item(3).Cell(1, 1).range.text = "Drop 5"
        ThisDocument.Tables.Item(3).Cell(1, 2).range.text = "Column 5/6"
        ThisDocument.Tables.Item(3).Cell(1, 3).range.text = "Drop 6"
        ThisDocument.Tables.Item(3).Cell(1, 4).range.text = "Column 6/7"
        ThisDocument.Tables.Item(3).Cell(1, 5).range.text = "Drop 7"
    End If
    If Drop = 8 Then
        ThisDocument.Tables.Item(3).Cell(1, 1).range.text = "Drop 7"
        ThisDocument.Tables.Item(3).Cell(1, 2).range.text = "Column 7/8"
        ThisDocument.Tables.Item(3).Cell(1, 3).range.text = "Drop 8"
        ThisDocument.Tables.Item(3).Cell(1, 4).range.text = "Column 8/1"
        ThisDocument.Tables.Item(3).Cell(1, 5).range.text = "Drop 1"
    End If
    If Drop = 2 Or Drop = 8 Then
        With Selection.Find
            .ClearFormatting
            .text = "South"
            .Replacement.ClearFormatting
            .Replacement.text = "North"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End With
    End If
    If Drop = 4 Or Drop = 6 Then
        With Selection.Find
            .ClearFormatting
            .text = "North"
            .Replacement.ClearFormatting
            .Replacement.text = "South"
            .Execute Replace:=wdReplaceAll, Forward:=True, _
            Wrap:=wdFindContinue
        End With
    End If
    With Selection.Find
        .ClearFormatting
        .text = "D*L**"
        .Replacement.ClearFormatting
        .Replacement.text = "D" & Drop & "L" & mark_up_level
        .Execute Replace:=wdReplaceAll, Forward:=True, _
        Wrap:=wdFindContinue
    End With
    'ChangeTableInFooters()
    Dim sec As Section
    Dim ftr As HeaderFooter
    Dim i As Integer
        For Each sec In ActiveDocument.Sections
            For Each ftr In sec.Footers
                If ftr.Exists Then
                    For i = ftr.range.Tables.Count To 1 Step -1
                        ftr.range.Tables(i).Cell(1, 1).range.text = "Campaign: As-built D" & Drop & "L" & mark_up_level
                    Next i
                End If
        Next ftr
    Next sec
    ThisDocument.Tables.Item(1).Cell(2, 2).range.text = "As-built records from drop " & Drop & ", level " & mark_up_level & "-" & mark_up_level + 1 & " work platform location."
    '----------
Dim oRng As range
    Application.ScreenUpdating = False
    Documents.Open FileName:="T:\QUALITY\SCANPRINT REPORTS\key_back.docx"
    'MsgBox (Documents(1).shapes.Count)
    Documents(1).Shapes.Item(1).Select
    Selection.Copy
    Documents("T:\QUALITY\SCANPRINT REPORTS\key_back.docx").Close SaveChanges:=wdDoNotSaveChanges
    Selection.Find.Replacement.ClearFormatting
    Selection.Find.text = "Drawing"
    Selection.Collapse Direction:=wdCollapseEnd
    'ActiveDocument.GoTo(What:=wdGoToPage, Count:=41).Select
    Selection.Find.text = "Drawing"
    Selection.Find.Execute
    Selection.Find.Wrap = wdFindContinue
    For i = 2 To ActiveDocument.Sections.Count - 1
        Set oRng = ActiveDocument.Sections.Item(i).range
        oRng.Collapse wdCollapseStart
        oRng.Select
        Selection.Paste
        'heading text
        Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        While heading.Style <> "Heading 6"
            Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        Wend
        heading.Expand Unit:=wdParagraph
        NaglowekTekst = Left(heading.text, Len(heading.text) - 1)
        While heading.Style <> "Heading 7"
            Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToNext)
        Wend
        extraVOffset = 42
        Select Case NaglowekTekst
            Case "LS FPC"
                vOffset = 504 + extraVOffset
                hOffset = 663
                panelHeight = 8
                panelWidth = 25
            Case "LS FPL"
                vOffset = 504 + extraVOffset
                hOffset = 663 - 25
                panelHeight = 8
                panelWidth = 25
            Case "LS FPR"
                vOffset = 504 + extraVOffset
                hOffset = 663 + 25
                panelHeight = 8
                panelWidth = 25
            Case "LS MBC"
                vOffset = 499 + extraVOffset
                hOffset = 663
                panelHeight = 5
                panelWidth = 25
            Case "LS MBL"
                vOffset = 499 + extraVOffset
                hOffset = 663 - 25
                panelHeight = 5
                panelWidth = 25
            Case "LS MBR"
                vOffset = 499 + extraVOffset
                hOffset = 663 + 25
                panelHeight = 5
                panelWidth = 25
            Case "LS SFC"
                vOffset = 490 + extraVOffset
                hOffset = 663
                panelHeight = 4
                panelWidth = 25
            Case "LS SFL"
                vOffset = 490 + extraVOffset
                hOffset = 663 - 25
                panelHeight = 4
                panelWidth = 25
            Case "LS SFR"
                vOffset = 490 + extraVOffset
                hOffset = 663 + 25
                panelHeight = 4
                panelWidth = 25
            Case "SS FPR"
                vOffset = 504 + extraVOffset
                hOffset = 550
                panelHeight = 8
                panelWidth = 22
            Case "SS MBR"
                vOffset = 499 + extraVOffset
                hOffset = 550
                panelHeight = 5
                panelWidth = 22
            Case "SS SFR"
                vOffset = 490 + extraVOffset
                hOffset = 550
                panelHeight = 4
                panelWidth = 22
            Case "SS FPL"
                vOffset = 504 + extraVOffset
                hOffset = 779
                panelHeight = 8
                panelWidth = 22
            Case "SS MBL"
                vOffset = 499 + extraVOffset
                hOffset = 779
                panelHeight = 5
                panelWidth = 22
            Case "SS SFL"
                vOffset = 490 + extraVOffset
                hOffset = 779
                panelHeight = 4
                panelWidth = 22
        End Select
        '----------------- oart for columns --------
        'mark_up_level = 51 'this to be inactivated when pasted to main procedure
        Select Case Right(NaglowekTekst, 2)
            Case "-1"
                'MsgBox (Right(Left(naglowekTekst, 6), 2) & " vs. " & mark_up_level)
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                    hOffset = 585
                    panelWidth = 10
                Else
                    vOffset = 489 + extraVOffset
                    hOffset = 581
                    panelWidth = 14
                End If
                panelHeight = 7
            Case "-2"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 595
                panelWidth = 5
                panelHeight = 7
            Case "-3"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 599
                panelWidth = 15
                panelHeight = 7
            Case "-4"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 614
                panelWidth = 10
                panelHeight = 7
            Case "-5"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    MsgBox ("Column panel #5 does not exist on this level, please check the content")
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 623
                panelWidth = 5
                panelHeight = 7
            Case "-6"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    MsgBox ("Column panel #6 does not exist on this level, please check the content")
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 722
                panelWidth = 5
                panelHeight = 7
            Case "-7"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 727
                panelWidth = 10
                panelHeight = 7
            Case "-8"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 736
                panelWidth = 15
                panelHeight = 7
            Case "-9"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                Else
                    vOffset = 489 + extraVOffset
                End If
                hOffset = 751
                panelWidth = 5
                panelHeight = 7
            Case "10"
                If Str(mark_up_level) = Str(Right(Left(NaglowekTekst, 6), 2)) Then
                    vOffset = 503 + extraVOffset
                    hOffset = 755
                    panelWidth = 14
                Else
                    vOffset = 489 + extraVOffset
                    hOffset = 755
                    panelWidth = 14
                End If
                panelHeight = 7
        End Select
        'MsgBox heading.text
        With ThisDocument.Shapes.AddShape(msoShapeRectangle, hOffset, vOffset, panelWidth, panelHeight)
            .Fill.Visible = msoTrue
            With .Line
                .Weight = 0.55
                .DashStyle = msoLineSquareDot
                .Style = msoLineSingle
                .Transparency = 0#
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 0, 0)
                .BackColor.RGB = RGB(255, 0, 0)
            End With
            .Fill.ForeColor.RGB = RGB(200, 110, 110)
            .Fill.BackColor.RGB = RGB(255, 170, 170)
            .Fill.Transparency = 0.34
        End With
        i = i + 1
    Next i
    'Next myStoryRange
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    Application.ScreenUpdating = True
End Sub
Sub delete_shapes()
    'MsgBox (ActiveDocument.shapes.Count)
    While ActiveDocument.Shapes.Count <> 0
        For Each shp In ActiveDocument.Shapes
            shp.Delete
        Next shp
    Wend
End Sub
Sub CopyHeader()
Dim docTemplate As Document
Dim strTemplate As String
Dim hdr1, ftr1 As HeaderFooter
Dim hdr2, ftr2 As HeaderFooter
Dim doc As Document
    Set doc = ActiveDocument
    strTemplate = "T:\QUALITY\SCANPRINT REPORTS\MLC-SPR-0xxx_X = As-built Drop x, level xx.docm"
    Set docTemplate = Documents.Open(strTemplate)
    Set hdr1 = docTemplate.Sections(1).Headers(wdHeaderFooterPrimary)
    Set ftr1 = docTemplate.Sections(1).Footers(wdHeaderFooterPrimary)
    Set hdr2 = doc.Sections(1).Headers(wdHeaderFooterPrimary)
    Set ftr2 = doc.Sections(1).Footers(wdHeaderFooterPrimary)
    hdr1.range.Copy
    hdr2.range.Select
    Selection.Delete
    Selection.Paste
    Selection.Collapse (Forward)
    Selection.TypeBackspace
    Selection.Font.Size = 6
    ftr1.range.Copy
    ftr2.range.Select
    Selection.Delete
    Selection.Paste
    Selection.Collapse (Forward)
    Selection.TypeBackspace
    Selection.Font.Size = 6
    docTemplate.Close False
    For J = 3 To ActiveDocument.Sections.Count
        For K = 1 To ActiveDocument.Sections(J).Headers.Count
            ActiveDocument.Sections(J).Headers(K).LinkToPrevious = True
        Next K
        For K = 1 To ActiveDocument.Sections(J).Footers.Count
            ActiveDocument.Sections(J).Footers(K).LinkToPrevious = True
        Next K
    Next J
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    ActiveDocument.ActiveWindow.View.Type = wdPrintView
    doc.ActiveWindow.View.Type = wdPrintView
End Sub
Sub CropDrawings()
Dim shp As InlineShape
Dim myCrop As Crop
Dim ReportLevel As Byte
Dim reportleveltxt As String
Dim columnSameLevel As Byte
Dim NaglowekTekst As String
    ThisDocument.InlineShapes(22).Select
    'MsgBox (ThisDocument.InlineShapes.Count)
    ReportLevel = Val(Left(Right(ThisDocument.Tables(1).Cell(1, 2).range.text, 4), 2))
    'MsgBox (ReportLevel)
    For i = 6 To ActiveDocument.InlineShapes.Count
        ThisDocument.InlineShapes(i).lockAspectRatio = msoTrue
        ThisDocument.InlineShapes(i).Select
        Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        While heading.Style <> "Heading 6"
            Set heading = Selection.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
        Wend
        heading.Expand Unit:=wdParagraph
        NaglowekTekst = Left(heading.text, Len(heading.text) - 1)
        'MsgBox (naglowektekst)
        reportleveltxt = "*" & ReportLevel & "*"
        reportlevelplusonetxt = "*" & ReportLevel + 1 & "*"
        If Right(NaglowekTekst, 2) = "-5" Or Right(NaglowekTekst, 2) = "-6" Then
            columnSameLevel = 1
        ElseIf NaglowekTekst Like reportleveltxt Then
            If Right(NaglowekTekst, 2) = "-2" Or Right(NaglowekTekst, 2) = "-9" Then
                columnSameLevel = 2
            Else
                columnSameLevel = 3
            End If
        ElseIf NaglowekTekst Like reportlevelplusonetxt Then
            If Right(NaglowekTekst, 2) = "-2" Or Right(NaglowekTekst, 2) = "-9" Then
                columnSameLevel = 4
            Else
                columnSameLevel = 5
            End If
        End If
        'MsgBox (columnSameLevel)
        If NaglowekTekst = "LS FPC" Or NaglowekTekst = "LS FPL" Or NaglowekTekst = "LS FPR" Or NaglowekTekst = "SS FPL" Or NaglowekTekst = "SS FPR" Then
                ThisDocument.InlineShapes(i).PictureFormat.CropBottom = 70
                ThisDocument.InlineShapes(i).PictureFormat.CropTop = 70
                If ThisDocument.InlineShapes(i).Width <> 700 Then
                    ThisDocument.InlineShapes(i).Width = 700
                End If
        End If
        If NaglowekTekst = "LS MBC" Or NaglowekTekst = "LS MBL" Or NaglowekTekst = "LS MBR" Or NaglowekTekst = "LS SFC" Or NaglowekTekst = "LS SFL" Or NaglowekTekst = "LS SFR" Or NaglowekTekst = "SS SFL" Or NaglowekTekst = "SS SFR" Or NaglowekTekst = "SS MBL" Or NaglowekTekst = "SS MBR" Then
                ThisDocument.InlineShapes(i).PictureFormat.CropBottom = 150
                ThisDocument.InlineShapes(i).PictureFormat.CropTop = 150
                If ThisDocument.InlineShapes(i).Width <> 700 Then
                    ThisDocument.InlineShapes(i).Width = 700
                End If
        End If
        If columnSameLevel = 1 Then 'panel #5, #6
            ThisDocument.InlineShapes(i).PictureFormat.CropBottom = 0
            ThisDocument.InlineShapes(i).PictureFormat.CropTop = 0
            ThisDocument.InlineShapes(i).PictureFormat.CropLeft = 400
            ThisDocument.InlineShapes(i).PictureFormat.CropRight = 400
            If ThisDocument.InlineShapes(i).Width <> 200 Then
                ThisDocument.InlineShapes(i).Width = 200
            End If
        ElseIf columnSameLevel = 2 Or columnSameLevel = 3 Then 'column level the same
            ThisDocument.InlineShapes(i).PictureFormat.CropBottom = 300
            ThisDocument.InlineShapes(i).PictureFormat.CropTop = 0
            ThisDocument.InlineShapes(i).PictureFormat.CropLeft = 300
            ThisDocument.InlineShapes(i).PictureFormat.CropRight = 300
            If ThisDocument.InlineShapes(i).Width <> 700 Then
                ThisDocument.InlineShapes(i).Width = 700
            End If
        ElseIf columnSameLevel = 4 Or columnSameLevel = 5 Then 'column level + 1
            ThisDocument.InlineShapes(i).PictureFormat.CropTop = 300
            ThisDocument.InlineShapes(i).PictureFormat.CropBottom = 0
            ThisDocument.InlineShapes(i).PictureFormat.CropLeft = 300
            ThisDocument.InlineShapes(i).PictureFormat.CropRight = 300
            If ThisDocument.InlineShapes(i).Width <> 700 Then
                ThisDocument.InlineShapes(i).Width = 700
            End If
        End If
        If columnSameLevel = 2 Or columnSameLevel = 4 Then 'panel #2, #9
            ThisDocument.InlineShapes(i).PictureFormat.CropLeft = 500
            ThisDocument.InlineShapes(i).PictureFormat.CropRight = 580
            'If ThisDocument.InlineShapes(i).Width <> 200 Then
            '    ThisDocument.InlineShapes(i).Width = 200
            'End If
        End If
        columnSameLevel = 0
    Next i
End Sub
Sub runAll()
    Call TblCellPadding
    Call adjustColumns
    Call lockAspectRatio
    Call ChangeStatusWording
    Call HighlightRows
    Call CountConditions
    Call CropDrawings
    ActiveDocument.TablesOfContents(1).Update
End Sub


