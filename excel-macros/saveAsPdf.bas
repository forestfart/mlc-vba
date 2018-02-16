Attribute VB_Name = "Module3"
Sub saveasPDF()
Attribute saveasPDF.VB_ProcData.VB_Invoke_Func = " \n14"
'
' saveasPDF Macro
'
    If Sheets("COLUMN 1-2").Cells(1, 1).Value = "" Then
        Sheets(Array("SS FPR", "SS MBR", "SS SFR", "COLUMN 1-2", "LS FPL", "LS MBL", "LS SFL", "LS FPC", "LS MBC", "LS SFC", "LS FPR", "LS MBR", "LS SFR", "COLUMN 2-3", "SS FPL", "SS MBL", "SS SFL")).Select
        Else: Sheets(Array("SS FPR", "SS MBR", "SS SFR", "COLUMN 7-8", "LS FPL", "LS MBL", "LS SFL", "LS FPC", "LS MBC", "LS SFC", "LS FPR", "LS MBR", "LS SFR", "COLUMN 8-1", "SS FPL", "SS MBL", "SS SFL")).Select
    End If
    Sheets("SS FPR").Activate
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    ActiveWindow.ScrollWorkbookTabs Position:=xlFirst
    Sheets("Control").Select
End Sub
