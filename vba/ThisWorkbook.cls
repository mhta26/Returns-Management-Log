Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    Module1.Initialize
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal target As Excel.Range)
    Module1.HandleChange Sh, target
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal target As Excel.Range, Cancel As Boolean)
    Module1.HandleDoubleClick Sh, target, Cancel
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal target As Excel.Range)
    Module1.HandleSelection Sh, target
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    If Module1.BusyPrinting Then Exit Sub
    If ActiveSheet.name = ThisWorkbook.Worksheets(1).name Then
        Cancel = True
        Module1.ArchiveAndPrint
    ElseIf Module1.IsSavedReceiptSheet(ActiveSheet.name) Then
        Cancel = True
        Module1.ReprintSavedSheet ActiveSheet.name
    End If
End Sub


Private Sub Workbook_Activate()
    Module1.ConfigureEntryMode
End Sub

Private Sub Workbook_Deactivate()
    Module1.RestoreEntryMode
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Module1.ConfigureEntryMode
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Module1.RestoreEntryMode
    Module1.CancelNativeDropdown
    Module1.ReleaseNotesReturn
    Module1.RestoreSearchStatus
End Sub
