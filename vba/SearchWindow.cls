Attribute VB_Name = "SearchWindow"
Attribute VB_Base = "0{AD3F73AA-BEED-479C-9896-E82DAA60A2F3}{76A2269E-3A61-4E3E-B83B-B0C799EC8B44}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Const PAGE_ROWS As Long = 11
Private WithEvents ResultScroll As MSForms.ScrollBar
Private WithEvents AcceptButton As MSForms.CommandButton
Private WithEvents CancelButton As MSForms.CommandButton
Private Rows(1 To PAGE_ROWS) As LookupRow
Private DataRows As Variant, Prepared() As String, FilteredRows() As Long
Private Customer As Boolean, Ready As Boolean, Drawing As Boolean
Private ResultCount As Long, SelectedIndex As Long
Public ChosenRow As Long

Private Function LabelAt(ByVal name As String, ByVal text As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal background As Long, ByVal foreground As Long, Optional ByVal bold As Boolean = False, Optional ByVal size As Long = 11) As MSForms.Label
    Dim item As MSForms.Label
    Set item = Me.Controls.Add("Forms.Label.1", name, True)
    With item
        .Left = x: .Top = y: .Width = width: .Height = height
        .Caption = text: .TextAlign = 3: .WordWrap = False
        .Font.Name = "Tahoma": .Font.Size = size: .Font.Bold = bold
        .BackColor = background: .ForeColor = foreground
        .BorderStyle = 0: .SpecialEffect = 0
    End With
    Set LabelAt = item
End Function

Private Sub UserForm_Initialize()
    Dim item As Object, label As MSForms.Label, code As MSForms.Label, i As Long
    Me.Width = 700: Me.Height = 520: Me.StartUpPosition = 1
    Me.Font.Name = "Tahoma": Me.Font.Size = 11
    Me.BackColor = RGB(246, 248, 252)
    For Each item In Me.Controls
        item.Visible = False
    Next item
    Set label = LabelAt("Heading", "", 20, 12, 652, 28, Me.BackColor, RGB(23, 54, 93), True, 16)
    Set label = LabelAt("Subtitle", Module1.LookupCaption(10), 20, 45, 652, 20, Me.BackColor, RGB(100, 116, 139), False, 10)
    With Me.TextBox1
        .Visible = True: .Left = 20: .Top = 75: .Width = 652: .Height = 32
        .Font.Name = "Tahoma": .Font.Size = 12: .TextAlign = 3
        .BackColor = RGB(255, 255, 255): .ForeColor = RGB(23, 54, 93)
        .BorderStyle = 1: .BorderColor = RGB(202, 213, 226): .SpecialEffect = 0
        .Value = "": .TabIndex = 0
    End With
    Set label = LabelAt("NameHeading", Module1.LookupCaption(11), 193, 120, 479, 22, RGB(230, 236, 244), RGB(23, 54, 93), True, 10)
    Set label = LabelAt("CodeHeading", Module1.LookupCaption(12), 38, 120, 155, 22, RGB(230, 236, 244), RGB(23, 54, 93), True, 10)
    label.TextAlign = 2
    Set ResultScroll = Me.Controls.Add("Forms.ScrollBar.1", "ResultScroll", True)
    With ResultScroll
        .Left = 20: .Top = 146: .Width = 14: .Height = PAGE_ROWS * 24
        .Orientation = 0: .Min = 0: .Max = 0: .SmallChange = 1: .LargeChange = PAGE_ROWS
        .TabStop = False
    End With
    For i = 1 To PAGE_ROWS
        Set label = LabelAt("RowName" & CStr(i), "", 193, 146 + (i - 1) * 24, 479, 24, vbWhite, RGB(38, 55, 74))
        Set code = LabelAt("RowCode" & CStr(i), "", 38, 146 + (i - 1) * 24, 155, 24, vbWhite, RGB(38, 55, 74), False, 10)
        code.TextAlign = 2
        Set Rows(i) = New LookupRow
        Rows(i).Attach Me, label, code
    Next i
    With Me.Label1
        .Visible = True: .Left = 20: .Top = 421: .Width = 652: .Height = 20
        .TextAlign = 3: .BackStyle = 0: .ForeColor = RGB(69, 91, 115)
    End With
    With Me.Label2
        .Visible = True: .Left = 20: .Top = 461: .Width = 350: .Height = 27
        .TextAlign = 3: .BackStyle = 0: .Font.Size = 9
        .ForeColor = RGB(100, 116, 139): .Caption = Module1.LookupCaption(4)
    End With
    Set AcceptButton = Me.Controls.Add("Forms.CommandButton.1", "LookupAccept", True)
    With AcceptButton
        .Left = 552: .Top = 455: .Width = 120: .Height = 32
        .Caption = Module1.LookupCaption(5): .Default = True: .TabIndex = 1
        .BackColor = RGB(23, 54, 93): .ForeColor = vbWhite: .Font.Bold = True
    End With
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "LookupCancel", True)
    With CancelButton
        .Left = 442: .Top = 455: .Width = 98: .Height = 32
        .Caption = Module1.LookupCaption(6): .Cancel = True: .TabIndex = 2
        .BackColor = RGB(230, 236, 244): .ForeColor = RGB(23, 54, 93)
    End With
End Sub

Public Sub Configure(ByVal isCustomer As Boolean, ByVal initialQuery As String)
    Dim i As Long
    Ready = False: ChosenRow = 0: Customer = isCustomer
    If Customer Then Me.Caption = Module1.LookupCaption(8) Else Me.Caption = Module1.LookupCaption(9)
    Me.Controls("Heading").Caption = Me.Caption
    DataRows = Module1.LookupData(Customer)
    If IsArray(DataRows) Then
        ReDim Prepared(1 To UBound(DataRows, 1))
        For i = 1 To UBound(DataRows, 1)
            Prepared(i) = Module1.NormalizeLookup(CStr(DataRows(i, 1)) & " " & CStr(DataRows(i, 2)))
        Next i
    End If
    Me.TextBox1.Value = initialQuery
    Ready = True
    FilterResults
End Sub

Private Sub UserForm_Activate()
    Me.TextBox1.SetFocus
    Me.TextBox1.SelStart = 0: Me.TextBox1.SelLength = Len(Me.TextBox1.Text)
End Sub
Private Sub TextBox1_Change()
    If Ready Then FilterResults
End Sub

Private Sub FilterResults()
    Dim query As String, tokens As Variant, token As Variant, suffix As Boolean
    Dim i As Long, count As Long, accepted As Boolean
    query = Module1.NormalizeLookup(CStr(Me.TextBox1.Value))
    tokens = Split(query, " ")
    suffix = Customer And (Len(query) = 4 Or Len(query) = 5) And Not (query Like "*[!0-9]*")
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
    End If
    ResultCount = count: SelectedIndex = 0
    Drawing = True
    ResultScroll.Min = 0
    If count > PAGE_ROWS Then ResultScroll.Max = count - PAGE_ROWS Else ResultScroll.Max = 0
    ResultScroll.Value = 0: ResultScroll.Enabled = (count > PAGE_ROWS)
    Drawing = False
    If count > 0 Then SelectedIndex = 1
    AcceptButton.Enabled = (count > 0)
    Me.Label1.Caption = Module1.LookupCaption(3) & CStr(count)
    If count = 0 Then Me.Label1.Caption = Module1.LookupCaption(7)
    RenderResults
End Sub

Private Sub RenderResults()
    Dim slot As Long, index As Long, source As Long
    For slot = 1 To PAGE_ROWS
        index = ResultScroll.Value + slot
        If index <= ResultCount Then
            source = FilteredRows(index)
            Rows(slot).Display index, CStr(DataRows(source, 2)), CStr(DataRows(source, 1)), (index = SelectedIndex)
        Else
            Rows(slot).Display 0, "", "", False
        End If
    Next slot
End Sub

Public Sub SelectResult(ByVal index As Long, ByVal confirm As Boolean)
    If index < 1 Or index > ResultCount Then Exit Sub
    SelectedIndex = index
    Drawing = True
    If index <= ResultScroll.Value Then ResultScroll.Value = index - 1
    If index > ResultScroll.Value + PAGE_ROWS Then ResultScroll.Value = index - PAGE_ROWS
    Drawing = False
    RenderResults
    If confirm Then ConfirmChoice Else Me.TextBox1.SetFocus
End Sub
Private Sub ResultScroll_Change()
    If Ready And Not Drawing Then RenderResults
End Sub
Private Sub ConfirmChoice()
    If SelectedIndex < 1 Or SelectedIndex > ResultCount Then Exit Sub
    ChosenRow = FilteredRows(SelectedIndex) + 1
    Me.Hide
End Sub
Private Sub AcceptButton_Click()
    ConfirmChoice
End Sub
Private Sub CancelButton_Click()
    ChosenRow = 0: Me.Hide
End Sub
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDown Then
        If SelectedIndex < ResultCount Then SelectResult SelectedIndex + 1, False
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        If SelectedIndex > 1 Then SelectResult SelectedIndex - 1, False
        KeyCode = 0
    ElseIf KeyCode = vbKeyPageDown Then
        If ResultCount > 0 Then SelectResult WorksheetFunction.Min(ResultCount, SelectedIndex + PAGE_ROWS), False
        KeyCode = 0
    ElseIf KeyCode = vbKeyPageUp Then
        If ResultCount > 0 Then SelectResult WorksheetFunction.Max(1, SelectedIndex - PAGE_ROWS), False
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0: ConfirmChoice
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0: ChosenRow = 0: Me.Hide
    End If
End Sub
Public Sub ReleaseRows()
    Dim i As Long
    For i = 1 To PAGE_ROWS
        If Not Rows(i) Is Nothing Then Rows(i).Detach
        Set Rows(i) = Nothing
    Next i
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = 1: ChosenRow = 0: Me.Hide
    End If
End Sub
