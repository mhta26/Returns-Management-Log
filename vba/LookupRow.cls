Attribute VB_Name = "LookupRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents NameField As MSForms.Label
Private WithEvents CodeField As MSForms.Label
Private Owner As SearchWindow
Private ItemIndex As Long

Public Sub Attach(ByVal parent As SearchWindow, ByVal title As MSForms.Label, ByVal code As MSForms.Label)
    Set Owner = parent
    Set NameField = title
    Set CodeField = code
End Sub

Public Sub Display(ByVal index As Long, ByVal title As String, ByVal code As String, ByVal selected As Boolean)
    Dim background As Long, foreground As Long
    ItemIndex = index
    NameField.Visible = (index > 0)
    CodeField.Visible = (index > 0)
    If index = 0 Then Exit Sub
    If selected Then
        background = RGB(226, 237, 249): foreground = RGB(23, 54, 93)
    Else
        If index Mod 2 = 0 Then background = RGB(245, 248, 252) Else background = RGB(255, 255, 255)
        foreground = RGB(38, 55, 74)
    End If
    NameField.BackColor = background: CodeField.BackColor = background
    NameField.ForeColor = foreground: CodeField.ForeColor = foreground
    NameField.Caption = ChrW(&H200F) & title & "  "
    CodeField.Caption = code
End Sub

Public Sub Detach()
    Set Owner = Nothing
    Set NameField = Nothing
    Set CodeField = Nothing
End Sub

Private Sub NameField_Click()
    If ItemIndex > 0 Then Owner.SelectResult ItemIndex, False
End Sub
Private Sub CodeField_Click()
    If ItemIndex > 0 Then Owner.SelectResult ItemIndex, False
End Sub
Private Sub NameField_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    If ItemIndex > 0 Then Owner.SelectResult ItemIndex, True
End Sub
Private Sub CodeField_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    If ItemIndex > 0 Then Owner.SelectResult ItemIndex, True
End Sub
