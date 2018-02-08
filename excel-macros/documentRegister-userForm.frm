VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserDoc 
   Caption         =   "Creade new document"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4185
   OleObjectBlob   =   "UserDoc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Code_Change()
Codec.List = Array("S", "Q")
End Sub

Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Create_Click()
Dim index As Integer
Dim Poczatek As Integer             'begining
Dim KoncowaKomorka As Integer       'last cell
Dim WierszZnaleziony As String      'last row
Dim ZakresKolumn As String          'column range
Dim DocExist As Integer
Dim Poprzedni As Integer            'previous
Dim DocumentFolder As String        'retrieve filepath for document location
Dim DocumentFolderPartial As String 'remove primary folder location from filepath
Dim Michal As String                'will provide document name and hyperlink as Title

    'error control
    If Kod.Value = "Select Document Code" Then
        MsgBox "Document code not selected.", vbCritical, "wrong!"
        End
    End If
    If Typ.Value = "Type" Then
        MsgBox "Document type not selected.", vbCritical, "wrong!"
        End
    End If
    If Tytul.Value = "Title" Then
        MsgBox "Document title not typed.", vbCritical, "wrong!"
        End
    End If
    Application.ScreenUpdating = False
    
    ' Dialog box to obtain document folder location for hyperlink
    ZakresKolumn = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        cell.EntireRow.Hidden = True
    Next
    index = 5
    DocExist = 0
    For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
        If cell.EntireRow.Hidden = True Then
            'Szukanie czy przypadkiem juz nie istnieje / check if the document exists or not
            If ActiveSheet.Cells(1, 6).Value = cell.Value Then
                DocExist = 1
                cell.EntireRow.Hidden = False
                WierszZnaleziony = cell.Row
                While ActiveSheet.Cells(WierszZnaleziony, 9).Value <> ""
                    ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                    WierszZnaleziony = WierszZnaleziony + 1
                Wend
                ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = False
                MsgBox "Document already registered", vbInformation, "wrong"
                Range("I" & WierszZnaleziony - 1).Select
                Selection.Copy
                End
            End If
       End If
       index = index + 1
    Next
    index = 5
    'Tworzenie dokumentu / create document section
    WierszZnaleziony = 0
    Poprzedni = ActiveSheet.Cells(1, 6).Value
    While WierszZnaleziony = 0 And DocExist = 0 And Poprzedni < 10000 And Poprzedni > 1
        Poprzedni = Poprzedni - 1
        For Each cell In ActiveSheet.Range("F5:F" & ZakresKolumn)
            If Poprzedni = cell.Value Then WierszZnaleziony = cell.Row
        Next
    Wend
    'Wstawianie wiersza / add row
    While ActiveSheet.Cells(WierszZnaleziony, 9).Value <> "" And DocExist = 0
        WierszZnaleziony = WierszZnaleziony + 1
    Wend
    ActiveSheet.Cells(WierszZnaleziony, 6).EntireRow.Hidden = True
    WierszZnaleziony = WierszZnaleziony + 1
    Rows(WierszZnaleziony & ":" & WierszZnaleziony).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows(WierszZnaleziony & ":" & WierszZnaleziony).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C6:Q6").Select
    Selection.Copy
    Range("C" & WierszZnaleziony & ":Q" & WierszZnaleziony).Select
    ActiveSheet.Paste
    'Zapis w komorkach / save info in cells
    ActiveSheet.Cells(WierszZnaleziony, 3).Value = 2040672
    ActiveSheet.Cells(WierszZnaleziony, 4).Value = Kod.Value
    ActiveSheet.Cells(WierszZnaleziony, 5).Value = Typ.Value
    ActiveSheet.Cells(WierszZnaleziony, 6).Value = ActiveSheet.Cells(1, 6).Value
    Tytul.Value = UCase(Tytul.Value)
    ActiveSheet.Cells(WierszZnaleziony, 7).Formula = Michal
    ActiveSheet.Cells(WierszZnaleziony, 9).Value = "A"
    Range("I" & WierszZnaleziony).Select
    Selection.Copy
    'ActiveSheet.Cells(6, 3).Value = Kod.Value
    'hyperlink code
    
    
    Unload Me
End Sub

