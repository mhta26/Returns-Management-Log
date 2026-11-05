Attribute VB_Name = "LookupEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Combo As MSForms.ComboBox

Private Sub Combo_Change()
    Module1.LookupTextChanged
End Sub

Private Sub Combo_Click()
    Module1.LookupClicked
End Sub

Private Sub Combo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Module1.LookupKeyDown KeyCode, Shift
End Sub

Private Sub Combo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Module1.LookupKeyReleased
End Sub

Private Sub Combo_DropButtonClick()
    Module1.LookupDropRequested
End Sub
