Attribute VB_Name = "Module1"

Sub CopyHeader()
Dim docTemplate As Document
Dim strTemplate As String
Dim hdr1, ftr1 As HeaderFooter
Dim hdr2, ftr2 As HeaderFooter
Dim doc As Document
    Set doc = ActiveDocument
    strTemplate = "T:\QUALITY\SCANPRINT REPORTS\MLC-SPR-0xxx_X = NDI Drop x, level xx.docm"
    Set docTemplate = Documents.Open(strTemplate)
    Set hdr1 = docTemplate.Sections(1).Headers(wdHeaderFooterPrimary)
    Set ftr1 = docTemplate.Sections(1).Footers(wdHeaderFooterPrimary)
    Set hdr2 = doc.Sections(1).Headers(wdHeaderFooterPrimary)
    Set ftr2 = doc.Sections(1).Footers(wdHeaderFooterPrimary)
    hdr1.Range.Copy
    hdr2.Range.Select
    Selection.Delete
    Selection.Paste
    Selection.Collapse (Forward)
    Selection.TypeBackspace
    Selection.Font.Size = 6
    ftr1.Range.Copy
    ftr2.Range.Select
    Selection.Delete
    Selection.Paste
    Selection.Collapse (Forward)
    Selection.TypeBackspace
    Selection.Font.Size = 6
    docTemplate.Close False
    For J = 1 To ActiveDocument.Sections.Count
        For K = 1 To ActiveDocument.Sections(J).Headers.Count
            ActiveDocument.Sections(J).Headers(K).LinkToPrevious = True
        Next K
        For K = 1 To ActiveDocument.Sections(J).Footers.Count
            ActiveDocument.Sections(J).Footers(K).LinkToPrevious = True
        Next K
    Next J
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    ActiveDocument.ActiveWindow.View.Type = wdPrintView
End Sub
Sub HighlightRows()
Dim tabela As Byte
Dim wiersz As Byte
    Application.ScreenUpdating = False
    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).Range.text Like "*Code*" Then
            For wiersz = 2 To ActiveDocument.Tables(tabela).Rows.Count
                If ActiveDocument.Tables(tabela).Cell(wiersz, 4).Range.text Like "*Repaired*" Then
                    ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorGray15
                    'ThisDocument.Tables.Item(tabela).Rows(wiersz).Select
                    'Selection.Font.ColorIndex = wdAuto
                ElseIf ActiveDocument.Tables(tabela).Cell(wiersz, 4).Range.text Like "*New entry*" Then
                    ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorLightGreen
                Else:
                    ThisDocument.Tables.Item(tabela).Rows(wiersz).Shading.BackgroundPatternColor = wdColorWhite
                End If
            Next wiersz
        End If
    Next tabela
    Application.ScreenUpdating = True
End Sub
Sub wstaw_dropdownlist(DN As Byte)
    Selection.Range.ContentControls.Add (wdContentControlDropdownList)
    With Selection.ParentContentControl.DropdownListEntries
        .Clear
        .Add text:="Choose an item.", Value:="0"                            '1  Choose an item.
        .Add text:="Investigate (1)", Value:="Investigate (1)"              '2  Investigate (1) - Potential gradient >0.4V/m and Base Potential < - 200mV. Confirm corrosion condition of rebar by inspection. If inspection confirms corrosion of rebar either repair by patch repair or HCP. If inspection reveals no corrosion, patch repair inspection breakout.
        .Add text:="Investigate (2)", Value:="Investigate (2)"              '3  Investigate (2) - Crack >0.15mm and Base Potential < - 200mV. Confirm corrosion condition of rebar by inspection. If inspection confirms corrosion of rebar either repair by patch repair or HCP. If inspection reveals no corrosion, patch repair inspection breakout.
        .Add text:="Investigate", Value:="Investigate"                      '4  Investigate
        .Add text:="Do nothing (1)", Value:="Do nothing (1)"                '5  Do nothing  (1) - Drummy area assessed as not posing a detachment risk
        .Add text:="Do nothing (2)", Value:="Do nothing (2)"                '6  Do nothing  (2) - Actual low cover record as ?20mm
        .Add text:="Do nothing (3)", Value:="Do nothing (3)"                '7  Do nothing  (3) - Existing patch assessed as not posing a detachment risk
        .Add text:="Do nothing (4)", Value:="Do nothing (4)"                '8  Do nothnig  (4) - Existing crack in marble assessed as not requiring repair
        .Add text:="Do nothing (5)", Value:="Do nothing (5)"                '9  Do nothing  (5) - 2010 audit condition - could not be identified/found on site
        .Add text:="Do nothing (6)", Value:="Do nothing (6)"                '10 Do nothing  (6) - Metallic aggregate corrosion assessed as not posing detachment risk  - TBC
        .Add text:="Do nothing (7)", Value:="Do nothing (7)"                '11 Do nothing  (7) - New condition does not pose detachment risk
        .Add text:="Do nothing (8)", Value:="Do nothing (8)"                '12 Do nothing  (8) - Surface erosion assessed as not compromising rebar cover
        .Add text:="Do nothing (9)", Value:="Do nothing (9)"                '13 Do nothing  (9) - Actual low cover above ?15mm, not in potential hot spot area
        .Add text:="Do nothing (10)", Value:="Do nothing (10)"              '14 Do nothing  (10) - Width of crack below 0.15mm, no risk of detachment
        .Add text:="Do nothing (11)", Value:="Do nothing (11)"              '15 Do nothing  (11) - Surface erosion assessed as not compromising rebar cover due to HCP
        .Add text:="Do nothing (12)", Value:="Do nothing (12)"              '16 Do nothing  (12) - Corner defect less than 40x40mm - trim face with grinder, unless rebar cover is less than 20mm then breakout
        .Add text:="Do nothing (13)", Value:="Do nothing (13)"              '17 Do nothing  (13) - Level 67-68 lower corner defects less than 40x40 - trim face with grinder,  unless rebar cover is less than 20mm then breakout
        .Add text:="Do nothing (14)", Value:="Do nothing (14)"              '18 Do nothing  (14) - Unless directed by client
        .Add text:="Do nothing (15)", Value:="Do nothing (15)"              '19 Do nothing  (15) - Panel edge defect less than 20mm wide
        .Add text:="Do nothing (16)", Value:="Do nothing (16)"              '20 Do nothing  (16) - If HCP installation does not cause drummy area to delaminate
        .Add text:="Do nothing (17)", Value:="Do nothing (17)"              '21 Do nothing  (17) - If erosion zone in HCP zone
        .Add text:="Do nothing (18)", Value:="Do nothing (18)"              '22 Do nothing  (18) - Subject to column edge repairs at level above/below
        .Add text:="Do nothing (19)", Value:="Do nothing (19)"              '23 Do nothing  (19) - Actual low cover record as ?15mm , not in hot spot area
        .Add text:="Do nothing (20)", Value:="Do nothing (20)"              '24 Do nothing  (20) - Crack has been corromapped - Not in hot spot area
        .Add text:="Do nothing (21)", Value:="Do nothing (21)"              '25 Do nothing  (21) - Base potential reading ?-200mV
        .Add text:="Do nothing (22)", Value:="Do nothing (22)"              '26 Do nothing  (22) - Panel integrity in sound, no other defects in the area
        .Add text:="Do nothing (23)", Value:="Do nothing (23)"              '27 Do nothing  (23) - ????
        .Add text:="Patch repair", Value:="Patch repair"                    '28 repair
        .Add text:="Surface patch repair", Value:="Surface patch repair"    '29 repair
        .Add text:="HCP", Value:="Hybrid cathodic protection"               '30 repair
        .Add text:="Tile replacement", Value:="Tile replacement"            '31 repair
        .Add text:="TBC on site", Value:="To be confirmed on site"          '32 TBC on site
        .Add text:="Do nothing", Value:="Do nothing"                        '33 Do nothing
        .Add text:="Seal underside sill", Value:="Seal underside seal"      '34 repair
        .Add text:="Cut and reseal joints between tiles", Value:="cut&re.." '35 repair
        .Add text:="Tile pinning", Value:="pinning"                         '36 repair
        .Add text:="Megapoxy", Value:="megapoxy"                            '37 repair
    End With
    Set oLE = Selection.ParentContentControl.DropdownListEntries(DN)
    oLE.Select 'wybierz
End Sub
Sub wstaw_dropdownYN(YorN As Byte)
    Selection.Range.ContentControls.Add (wdContentControlDropdownList)
    With Selection.ParentContentControl.DropdownListEntries
        .Clear
        .Add text:="YES/No", Value:="0"                                     '1
        .Add text:="No", Value:="No"                                        '2
        .Add text:="YES", Value:="Yes"                                      '3
        .Add text:="TBC on site", Value:="TBC"                              '4
    End With
    Set oLE = Selection.ParentContentControl.DropdownListEntries(YorN)
    oLE.Select 'wybierz
End Sub
Sub pogrob()
Dim tabela As Byte
Dim wiersz As Byte
    Application.ScreenUpdating = False
    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).Range.text Like "*Code*" Then
            For wiersz = 2 To ActiveDocument.Tables(tabela).Rows.Count
                If ActiveDocument.Tables(tabela).Cell(wiersz, 7).Range.text Like "*No*" Then
                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 0
                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.Font.Bold = 0
                Else:
                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.Font.Bold = 1
                End If
            Next wiersz
        End If
    Next tabela
    Application.ScreenUpdating = True
End Sub
Sub marble_Special(tabela As Byte, wiersz As Byte)
    'cracks and eroded clay seams on sill do nothing by default uless detachement risk
    If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*etach*" Then
            'wstaw yes/no to achieve PPO
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
            Selection.Collapse 1
            wstaw_dropdownYN (4)    '4 - TBC on site
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
            'wstaw do nothing droplist
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
            Selection.Collapse 1
            wstaw_dropdownlist (4) '4 - Investigate
            'wstaw_dropdownlist (31) '31 - tile replacement
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
    ElseIf ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*fat*" Then
            'wstaw yes/no to achieve PPO
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
            Selection.Collapse 1
            wstaw_dropdownYN (3)    '3 - YES
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
            'wstaw do nothing droplist
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
            Selection.Collapse 1
            wstaw_dropdownlist (37) '37 - Megapoxy
            'wstaw_dropdownlist (31) '31 - tile replacement
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
    Else
            'wstaw yes/no to achieve PPO
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
            Selection.Collapse 1
            wstaw_dropdownYN (2)    '2 - No
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
            'wstaw do nothing droplist
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
            Selection.Collapse 1
            wstaw_dropdownlist (33) '33 - Do nothing
            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
    End If
End Sub
Sub main()
Dim text As String
Dim variableText As String
Dim index As Integer
Dim tabela As Byte
Dim wiersz As Byte
Dim coverToReo As Byte
Dim wiersz2 As Byte
Dim marble As Boolean
Dim crackWidth As Single
Dim rng As Range
Dim oLE As ContentControlListEntry
    'wygas odswiezanie
    'Application.ScreenUpdating = False
    If MsgBox("Content of column 7 and 8 will be deleted. Are you sure?", vbYesNo, "Confirm") = vbYes Then
        'znajdz tabele
        For tabela = 1 To ActiveDocument.Tables.Count
            If ThisDocument.Tables.Item(tabela).Cell(1, 1).Range.text Like "*Code*" Then
                'check if it is a marble table
                marble = False
                For wiersz2 = 2 To ActiveDocument.Tables(tabela).Rows.Count
                    If ThisDocument.Tables(tabela).Cell(wiersz2, 2).Range.text Like "*arble*" Then
                        marble = True
                        Exit For
                    Else
                        If ThisDocument.Tables(tabela).Cell(wiersz2, 2).Range.text Like "*Clay*" Then
                            marble = True
                            Exit For
                        Else
                            If ThisDocument.Tables(tabela).Cell(wiersz2, 2).Range.text Like "*Tile*" Then
                                marble = True
                                Exit For
                            Else
                                If ThisDocument.Tables(tabela).Cell(wiersz2, 2).Range.text Like "*Displacement*" Then
                                    marble = True
                                    Exit For
                                End If
                            End If
                        End If
                    End If
                Next wiersz2
                Select Case marble
                Case Is = False        '======================================================================== macro for precast panels
                    For wiersz = 2 To ActiveDocument.Tables.Item(tabela).Rows.Count
                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                        Selection.Delete
                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                        Selection.Delete
                        If ThisDocument.Tables(tabela).Cell(wiersz, 4).Range.text Like "*Repaired*" Then
                            'jezeli nie znaleziony to bezwarunkowo DN5
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (5)"
                        Else
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*DRU*" Then
                                'jezeli drummy to odrazu DN1
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (2)
                                'wstaw do nothing droplist
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (5) '5 - DN1
                            End If
                            If ThisDocument.Tables.Item(tabela).Cell(wiersz, 1).Range.text Like "*ICVR*" Then
                                'low cover decision
                                variableText = LTrim(Left(Right(ThisDocument.Tables.Item(tabela).Cell(wiersz, 5).Range.text, 5), 2))
                                coverToReo = CInt(variableText)
                                'ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = coverToReo
                                Select Case coverToReo
                                    Case Is < 15
                                        'wstaw yes to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                                        'wstaw patch repair
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                                    Case Is >= 20
                                        'wstaw no to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                                        'wstaw DN2
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (2)"
                                        If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*rill*" Then
                                            'wstaw tekst jezeli bylo wiercone
                                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                            Selection.Collapse 1
                                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.InsertAfter Chr(11) & "Patch drill hole"
                                        End If
                                    Case Else
                                        'wstaw yes/no to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownYN (2)
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdRed
                                        'wstaw do nothing droplist
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownlist (13) '13 = DN9
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdRed
                                        If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*rill*" Then
                                            'wstaw tekst jezeli bylo wiercone
                                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                            Selection.Collapse 1
                                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.InsertAfter Chr(11) & "Patch drill hole"
                                        End If
                                End Select
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*CRA*" Then
                                'width decision
                                variableText = LTrim(Left(Right(ThisDocument.Tables.Item(tabela).Cell(wiersz, 5).Range.text, 7), 5))
                                crackWidth = Val(variableText)
                                'ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = crackWidth
                                Select Case crackWidth
                                    Case Is = 0
                                        'wstaw tekst ze brakuje info
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                        Selection.Collapse 1
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.InsertAfter "; missing/incorrect crack width record"
                                        'wstaw yes/no to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownYN (4)  '4 - TBC on site
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdTurquoise
                                        'wstaw repair techinques
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownlist (4) '4 - Investigate
                                        'wstaw_dropdownlist (32) '32 - TBC on site
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdTurquoise
                                    Case Is < 0.1
                                        'wstaw yes/no to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownYN (2) '2 - No
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdTurquoise
                                        'wstaw do nothing droplist
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownlist (14) '14 - do nothing (10)
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdTurquoise
                                    Case Is > 0.15
                                        'wstaw yes/no to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                        Selection.Collapse 1
                                        wstaw_dropdownYN (3) '3 - Yes
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdTurquoise
                                        'wstaw do nothing droplist
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                        Selection.Collapse 1
                                        If ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.text Like "*dge*" Then
                                            wstaw_dropdownlist (28) '28 - Patch repair
                                            ElseIf ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.text Like "*within*" Then
                                                wstaw_dropdownlist (28) '28 - Patch repair
                                                Else:
                                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Select
                                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.InsertBefore " "
                                                    
                                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.InsertBefore "-200mV"
                                                    Selection.InsertSymbol Font:="Tahoma", CharacterNumber:=8805, Unicode:=True
                                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.InsertBefore "Base potential "
                                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                                    Selection.Collapse 1
                                                    wstaw_dropdownlist (29) '29 - Surface patch repair
                                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.HighlightColorIndex = wdTurquoise
                                        End If
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdTurquoise
                                    Case Else
                                        'wstaw yes/no to achieve PPO
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                        Selection.Collapse 1
                                        If ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.text Like "*dge*" Then
                                            wstaw_dropdownYN (4) '4 - TBC on site
                                            ElseIf ThisDocument.Tables.Item(tabela).Cell(wiersz, 6).Range.text Like "*within*" Then
                                                wstaw_dropdownYN (4) '4 - TBC on site
                                                Else: wstaw_dropdownYN (2) '2 - No
                                        End If
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
                                        'wstaw do nothing droplist
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                        Selection.Collapse 1
                                        If ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.text Like "*TBC on site*" Then
                                            wstaw_dropdownlist (4) '4 - Investigate
                                            Else: wstaw_dropdownlist (14) '14 - do nothing (10)
                                        End If
                                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
                                End Select
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NRPG*" Then
                                'jezeli drummy to odrazu DN1
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (2)   '2 - No
                                'wstaw do nothing droplist
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (25) '25 - DN21
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdViolet
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wd
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*OPO*" Then
                                'jezeli old patch appears ok to bezwarunkowo DN3
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (3)"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NEBS*" Then
                                If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*racks*" Then
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
                                Else
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (8)"
                                End If
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*EROB*" Then
                                If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*racks*" Then
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
                                Else
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (8)"
                                End If
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NUSB*" Then
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (22)"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NOPD*" Then
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*OLDF*" Then
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*CNSP*" Then
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (4)    '4 - TBC on site
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (28) '28 - Patch repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdGray25
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdGray25
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NCSC*" Then
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (4)    '4 - TBC on site
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (28) '28 - Patch repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdGray25
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdGray25
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*DISP*" Then
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Tile pinning"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*CUT*" Then
                                'jezeli cut out appears ok to bezwarunkowo patch repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*tie bolt*" Then
                                'jezeli mast tie holes to bezwarunkowo patch repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*hole*" Then
                                'jezeli mast tie holes to bezwarunkowo patch repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*rill*" Then
                                'jezeli wywzwpirzewiercone to bezwarunkowo patch repair --------------------------- trzeba przetestowac
                                If Not ThisDocument.Tables.Item(tabela).Cell(wiersz, 1).Range.text Like "*ICVR*" Then
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                    ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Patch repair"
                                End If
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*only*" Then
                                'jezeli mast tie holes to bezwarunkowo patch repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (8)"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*ie wire*" Then
                                'jezeli tie wire to bezwarunkowo TBC on site
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (4)    '4 - TBC on site
                                'ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
                                'wstaw repair techinque
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (4) '4 - Investigate
                                'wstaw_dropdownlist (10) '10 - Do ntohing (6)
                                'ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
                            End If
                            'anything else
                            If ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.text = Chr(13) & Chr(7) Then
                                'wstaw yes/no to achieve PPO
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (1)    'tu trzeba wybrac recznie
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
                                'wstaw do nothing droplist
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (1)  'tu trzeba wybrac recznie
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
                            End If
                        End If
                    Next wiersz
                Case Is = True        '======================================================================== macro for marble sills
                    For wiersz = 2 To ActiveDocument.Tables.Item(tabela).Rows.Count
                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                        Selection.Delete
                        ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                        Selection.Delete
                        If ThisDocument.Tables(tabela).Cell(wiersz, 4).Range.text Like "*Not found*" Then
                            'jezeli nie znaleziony to bezwarunkowo DN5
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "No"
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Do nothing (5)"
                        ElseIf ThisDocument.Tables(tabela).Cell(wiersz, 6).Range.text Like "*oint*" Then
                            'jezeli on joint to megapoxy
                             'wstaw yes/no to achieve PPO
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                            Selection.Collapse 1
                            wstaw_dropdownYN (3) 'tu trzeba wybrac recznie
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdGreen
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                            'wstaw do nothing droplist
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                            Selection.Collapse 1
                            wstaw_dropdownlist (37) 'tu trzeba wybrac recznie
                            ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdGreen
                        Else
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*CRA*" Then
                                    Call marble_Special(tabela, wiersz)
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NCMS*" Then
                                    Call marble_Special(tabela, wiersz)
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*EROD*" Then
                                    Call marble_Special(tabela, wiersz)
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NECS*" Then
                                    Call marble_Special(tabela, wiersz)
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*FRA*" Then
                                    Call marble_Special(tabela, wiersz)
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*NTES*" Then
                                'missing baffle always repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Seal underside seal"
                            End If
                            If ThisDocument.Tables(tabela).Cell(wiersz, 1).Range.text Like "*DISP*" Then
                                'missing baffle always repair
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range = "YES"
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range = "Cut and reseal joints between tile"
                            End If
                            'anything else
                            If ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.text = Chr(13) & Chr(7) Then
                                'wstaw yes/no to achieve PPO
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Select
                                Selection.Collapse 1
                                wstaw_dropdownYN (1) 'tu trzeba wybrac recznie
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.HighlightColorIndex = wdYellow
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 7).Range.Font.Bold = 1
                                'wstaw do nothing droplist
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Select
                                Selection.Collapse 1
                                wstaw_dropdownlist (1) 'tu trzeba wybrac recznie
                                ThisDocument.Tables.Item(tabela).Cell(wiersz, 8).Range.HighlightColorIndex = wdYellow
                            End If
                        End If
                    Next wiersz
                End Select
            End If
        Next tabela
        Call pogrob
        Call TblCellPadding
        Call CopyHeader
        Call HighlightRows
        Call adjustColumns
        Call CropDrawings
    End If
    'przywroc odswiezanie ekranu
    Application.ScreenUpdating = True
End Sub
Sub adjustColumns()
Dim tabela As Byte
    'Application.ScreenUpdating = False
    'ActiveDocument.PageSetup.LeftMargin = CentimetersToPoints(1.7)

    For tabela = 1 To ActiveDocument.Tables.Count
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).Range.text Like "*Code*" Then
            ActiveDocument.Tables.Item(tabela).Rows.SetLeftIndent LeftIndent:=-10, RulerStyle:=wdAdjustFirstColumn  'przesun cala tabele w lewo 10 jednostek
        End If
        If ThisDocument.Tables.Item(tabela).Cell(1, 1).Range.text Like "*Code*" Then
            ActiveDocument.Tables.Item(tabela).Columns(1).Width = 40
            ActiveDocument.Tables.Item(tabela).Columns(2).Width = 90
            ActiveDocument.Tables.Item(tabela).Columns(3).Width = 40
            ActiveDocument.Tables.Item(tabela).Columns(4).Width = 50
            ActiveDocument.Tables.Item(tabela).Columns(5).Width = 100 'properties
            ActiveDocument.Tables.Item(tabela).Columns(6).Width = 75
            ActiveDocument.Tables.Item(tabela).Columns(7).Width = 35
            ActiveDocument.Tables.Item(tabela).Columns(8).Width = 70
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
        If ThisDocument.Tables.Item(tabele).Cell(1, 1).Range.text Like "*Code*" Then
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
Sub CropDrawings()
Dim shp As InlineShape
Dim myCrop As Crop
Dim ReportLevel As Byte
Dim reportleveltxt As String
Dim columnSameLevel As Byte
Dim NaglowekTekst As String

    'MsgBox (ThisDocument.InlineShapes.Count)
    ReportLevel = Val(Left(Right(ThisDocument.Tables(1).Cell(1, 2).Range.text, 4), 2))
    'MsgBox (ReportLevel)
    For i = 4 To ActiveDocument.InlineShapes.Count
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
            ThisDocument.InlineShapes(i).PictureFormat.CropLeft = 580
            ThisDocument.InlineShapes(i).PictureFormat.CropRight = 580
            If ThisDocument.InlineShapes(i).Width <> 200 Then
                ThisDocument.InlineShapes(i).Width = 200
            End If
        End If
        columnSameLevel = 0
    Next i
End Sub

