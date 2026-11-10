Attribute VB_Name = "SearchWindow"
Attribute VB_Base = "0{AD3F73AA-BEED-479C-9896-E82DAA60A2F3}{76A2269E-3A61-4E3E-B83B-B0C799EC8B44}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private WithEvents Results As MSForms.ListBox
Attribute Results.VB_VarHelpID = -1
Private WithEvents AcceptButton As MSForms.CommandButton
Attribute AcceptButton.VB_VarHelpID = -1
Private WithEvents CancelButton As MSForms.CommandButton
Attribute CancelButton.VB_VarHelpID = -1
Private DataRows As Variant, Prepared() As String, FilteredRows() As Long
Private Customer As Boolean, Ready As Boolean, RTLReady As Boolean
Public ChosenRow As Long

Private Sub UserForm_Initialize()
    Dim item As Object
    Ready = False
    Me.Width = 670
    Me.BackColor = RGB(243, 246, 250)
    Me.Height = 475
    Me.StartUpPosition = 1
    Me.Font.name = "Tahoma"
    Me.Font.Size = 11
    For Each item In Me.Controls
        item.Visible = False
    Next item
    With Me.TextBox1
        .Visible = True
        .Left = 18: .Top = 49: .Width = 630: .Height = 32
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(23, 54, 93)
        .BorderColor = RGB(182, 201, 220)
        .Font.name = "Tahoma": .Font.Size = 12
        .TextAlign = 3
        .value = ""
        .TabIndex = 0
    End With
    Set Results = Me.Controls.Add("Forms.ListBox.1", "LookupResults", True)
    With Results
        .Left = 18: .Top = 96: .Width = 630: .Height = 270
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(35, 55, 77)
        .Font.name = "Tahoma": .Font.Size = 11
        .ColumnCount = 2: .ColumnWidths = "470 pt;150 pt"
        .MultiSelect = 0: .TabIndex = 1
    End With
    With Me.Label1
        .Visible = True: .Left = 18: .Top = 378: .Width = 630: .Height = 20
        .ForeColor = RGB(8, 113, 125)
        .TextAlign = 3
    End With
    With Me.Label2
        .Visible = True: .Left = 278: .Top = 414: .Width = 370: .Height = 32
        .ForeColor = RGB(82, 103, 126)
        .TextAlign = 3
        .caption = Module1.LookupCaption(4)
    End With
    Set AcceptButton = Me.Controls.Add("Forms.CommandButton.1", "LookupAccept", True)
    With AcceptButton
        .Left = 18: .Top = 410: .Width = 130: .Height = 32
        .BackColor = RGB(8, 126, 139)
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .caption = Module1.LookupCaption(5): .Default = True: .TabIndex = 2
    End With
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "LookupCancel", True)
    With CancelButton
        .Left = 160: .Top = 410: .Width = 100: .Height = 32
        .BackColor = RGB(225, 233, 243)
        .ForeColor = RGB(23, 54, 93)
        .caption = Module1.LookupCaption(6): .Cancel = True: .TabIndex = 3
    End With
End Sub

Public Sub Configure(ByVal isCustomer As Boolean, ByVal initialQuery As String)
    Dim i As Long
    Ready = False
    ChosenRow = 0
    Customer = isCustomer
    If Customer Then Me.caption = Module1.LookupCaption(1) Else Me.caption = Module1.LookupCaption(2)
    DataRows = Module1.LookupData(Customer)
    If IsArray(DataRows) Then
        ReDim Prepared(1 To UBound(DataRows, 1))
        For i = 1 To UBound(DataRows, 1)
            Prepared(i) = Module1.NormalizeLookup(CStr(DataRows(i, 1)) & " " & CStr(DataRows(i, 2)))
        Next i
    End If
    Me.TextBox1.value = initialQuery
    Ready = True
    FilterResults
End Sub

Private Sub UserForm_Activate()
    If Not RTLReady Then RTLReady = Module1.StyleLookupRTL(Me.Caption)
    Me.TextBox1.SetFocus
    Me.TextBox1.SelStart = 0
    Me.TextBox1.SelLength = Len(Me.TextBox1.text)
End Sub

Private Sub TextBox1_Change()
    If Ready Then FilterResults
End Sub

Private Sub FilterResults()
    Dim query As String, tokens As Variant, token As Variant, suffix As Boolean
    Dim i As Long, j As Long, count As Long, accepted As Boolean, output() As Variant
    query = Module1.NormalizeLookup(CStr(Me.TextBox1.value))
    tokens = Split(query, " ")
    suffix = Customer And (Len(query) = 4 Or Len(query) = 5) And Not (query Like "*[!0-9]*")
    Results.Clear
    AcceptButton.Enabled = False
    If IsArray(DataRows) Then
        ReDim FilteredRows(1 To UBound(DataRows, 1))
        For i = 1 To UBound(DataRows, 1)
            accepted = Len(CStr(DataRows(i, 1))) > 0 And Len(CStr(DataRows(i, 2))) > 0
            If suffix Then
                accepted = accepted And (Right$(Module1.NormalizeLookup(CStr(DataRows(i, 1))), Len(query)) = query)
            Else
                For Each token In tokens
                    If Len(token) > 0 Then
                        If InStr(1, Prepared(i), CStr(token), vbBinaryCompare) = 0 Then accepted = False: Exit For
                    End If
                Next token
            End If
            If accepted Then count = count + 1: FilteredRows(count) = i
        Next i
        If count > 0 Then
            ReDim output(0 To count - 1, 0 To 1)
            For j = 1 To count
                i = FilteredRows(j)
                output(j - 1, 0) = CStr(DataRows(i, 2))
                output(j - 1, 1) = CStr(DataRows(i, 1))
            Next j
            Results.List = output
            Results.ListIndex = 0
            AcceptButton.Enabled = True
        End If
    End If
    Me.Label1.caption = Module1.LookupCaption(3) & CStr(count)
    If count = 0 Then Me.Label1.caption = Module1.LookupCaption(7)
End Sub

Private Sub ConfirmChoice()
    If Results.ListIndex < 0 Then Exit Sub
    ChosenRow = FilteredRows(Results.ListIndex + 1) + 1
    Me.Hide
End Sub

Private Sub AcceptButton_Click()
    ConfirmChoice
End Sub

Private Sub CancelButton_Click()
    ChosenRow = 0
    Me.Hide
End Sub

Private Sub Results_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ConfirmChoice
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown Then
        If Results.ListCount > 0 Then Results.SetFocus
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0: ConfirmChoice
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0: ChosenRow = 0: Me.Hide
    End If
End Sub

Private Sub Results_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0: ConfirmChoice
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0: ChosenRow = 0: Me.Hide
    ElseIf KeyCode = vbKeyUp And Results.ListIndex = 0 Then
        KeyCode = 0: Me.TextBox1.SetFocus
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = 1: ChosenRow = 0: Me.Hide
    End If
End Sub
