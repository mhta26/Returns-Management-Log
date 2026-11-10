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
Private Const ROW_HEIGHT As Single = 26
Private VisibleRows As Long
Private WithEvents ResultScroll As MSForms.ScrollBar
Private WithEvents AcceptButton As MSForms.CommandButton
Private WithEvents CancelButton As MSForms.CommandButton
Private NameRows(1 To PAGE_ROWS) As MSForms.Label, CodeRows(1 To PAGE_ROWS) As MSForms.Label
Private SlotIndexes(1 To PAGE_ROWS) As Long
Private WithEvents Name01 As MSForms.Label
Private WithEvents Code01 As MSForms.Label
Private WithEvents Name02 As MSForms.Label
Private WithEvents Code02 As MSForms.Label
Private WithEvents Name03 As MSForms.Label
Private WithEvents Code03 As MSForms.Label
Private WithEvents Name04 As MSForms.Label
Private WithEvents Code04 As MSForms.Label
Private WithEvents Name05 As MSForms.Label
Private WithEvents Code05 As MSForms.Label
Private WithEvents Name06 As MSForms.Label
Private WithEvents Code06 As MSForms.Label
Private WithEvents Name07 As MSForms.Label
Private WithEvents Code07 As MSForms.Label
Private WithEvents Name08 As MSForms.Label
Private WithEvents Code08 As MSForms.Label
Private WithEvents Name09 As MSForms.Label
Private WithEvents Code09 As MSForms.Label
Private WithEvents Name10 As MSForms.Label
Private WithEvents Code10 As MSForms.Label
Private WithEvents Name11 As MSForms.Label
Private WithEvents Code11 As MSForms.Label
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
    Dim windowWidth As Single, windowHeight As Single, contentWidth As Single
    Dim nameLeft As Single, nameWidth As Single, codeWidth As Single, footerTop As Single, gridTop As Single
    windowWidth = 760: windowHeight = 560
    If Application.Width - 48 < windowWidth Then windowWidth = Application.Width - 48
    If Application.Height - 48 < windowHeight Then windowHeight = Application.Height - 48
    If windowWidth < 520 Then windowWidth = 520
    If windowHeight < 320 Then windowHeight = 320
    Me.Width = windowWidth: Me.Height = windowHeight: Me.StartUpPosition = 1
    Me.Font.Name = "Tahoma": Me.Font.Size = 11
    Me.BackColor = RGB(248, 250, 253)
    contentWidth = Me.InsideWidth - 40
    codeWidth = 140
    If contentWidth < 560 Then codeWidth = 120
    nameLeft = 40 + codeWidth
    nameWidth = Me.InsideWidth - 20 - nameLeft
    gridTop = 152
    VisibleRows = Int((Me.InsideHeight - gridTop - 94) / ROW_HEIGHT)
    If VisibleRows < 1 Then VisibleRows = 1
    If VisibleRows > PAGE_ROWS Then VisibleRows = PAGE_ROWS
    footerTop = Me.InsideHeight - 54
    For Each item In Me.Controls
        item.Visible = False
    Next item
    Set label = LabelAt("Heading", "", 20, 12, contentWidth, 28, Me.BackColor, RGB(23, 54, 93), True, 15)
    Set label = LabelAt("Subtitle", Module1.LookupCaption(10), 20, 45, contentWidth, 20, Me.BackColor, RGB(100, 116, 139), False, 10)
    With Me.TextBox1
        .Visible = True: .Left = 20: .Top = 78: .Width = contentWidth: .Height = 34
        .Font.Name = "Tahoma": .Font.Size = 12: .TextAlign = 3
        .BackColor = RGB(255, 255, 255): .ForeColor = RGB(23, 54, 93)
        .BorderStyle = 1: .BorderColor = RGB(199, 211, 226): .SpecialEffect = 0
        .Value = "": .TabIndex = 0
    End With
    Set label = LabelAt("NameHeading", Module1.LookupCaption(11), nameLeft, 122, nameWidth, 24, RGB(231, 237, 245), RGB(23, 54, 93), True, 10)
    Set label = LabelAt("CodeHeading", Module1.LookupCaption(12), 40, 122, codeWidth, 24, RGB(231, 237, 245), RGB(23, 54, 93), True, 10)
    label.TextAlign = 2
    Set ResultScroll = Me.Controls.Add("Forms.ScrollBar.1", "ResultScroll", True)
    With ResultScroll
        .Left = 20: .Top = gridTop: .Width = 14: .Height = VisibleRows * ROW_HEIGHT
        .Orientation = 0: .Min = 0: .Max = 0: .SmallChange = 1: .LargeChange = VisibleRows
        .TabStop = False
    End With
    For i = 1 To PAGE_ROWS
        Set label = LabelAt("RowName" & CStr(i), "", nameLeft, gridTop + (i - 1) * ROW_HEIGHT, nameWidth, ROW_HEIGHT, vbWhite, RGB(38, 55, 74))
        Set code = LabelAt("RowCode" & CStr(i), "", 40, gridTop + (i - 1) * ROW_HEIGHT, codeWidth, ROW_HEIGHT, vbWhite, RGB(38, 55, 74), False, 10)
        label.Visible = (i <= VisibleRows): code.Visible = (i <= VisibleRows)
        code.TextAlign = 2
        Set NameRows(i) = label: Set CodeRows(i) = code
        BindRow i, label, code
    Next i
    With Me.Label1
        .Visible = True: .Left = 20: .Top = gridTop + VisibleRows * ROW_HEIGHT + 8
        .Width = contentWidth: .Height = 16
        .Font.Name = "Tahoma": .Font.Size = 10
        .TextAlign = 3: .BackStyle = 0: .ForeColor = RGB(69, 91, 115)
    End With
    Set label = LabelAt("FooterDivider", "", 20, footerTop - 10, contentWidth, 1, RGB(224, 231, 240), RGB(224, 231, 240))
    With Me.Label2
        .Visible = True: .Left = 20: .Top = footerTop + 1
        .Width = contentWidth - 240: .Height = 34
        .TextAlign = 3: .BackStyle = 0: .WordWrap = True
        .Font.Name = "Tahoma": .Font.Size = 9
        .ForeColor = RGB(100, 116, 139): .Caption = Module1.LookupCaption(4)
    End With
    Set AcceptButton = Me.Controls.Add("Forms.CommandButton.1", "LookupAccept", True)
    With AcceptButton
        .Left = Me.InsideWidth - 140: .Top = footerTop: .Width = 120: .Height = 34
        .Font.Name = "Tahoma": .Font.Size = 11
        .Caption = Module1.LookupCaption(5): .Default = True: .TabIndex = 1
        .BackColor = RGB(23, 54, 93): .ForeColor = vbWhite: .Font.Bold = True
    End With
    Set CancelButton = Me.Controls.Add("Forms.CommandButton.1", "LookupCancel", True)
    With CancelButton
        .Left = Me.InsideWidth - 248: .Top = footerTop: .Width = 96: .Height = 34
        .Font.Name = "Tahoma": .Font.Size = 11
        .Caption = Module1.LookupCaption(6): .Cancel = True: .TabIndex = 2
        .BackColor = RGB(231, 237, 245): .ForeColor = RGB(23, 54, 93)
    End With
End Sub

Public Sub Configure(ByVal isCustomer As Boolean, ByVal initialQuery As String)
    Ready = False: ChosenRow = 0: Customer = isCustomer
    If Customer Then Me.Caption = Module1.LookupCaption(8) Else Me.Caption = Module1.LookupCaption(9)
    Me.Controls("Heading").Caption = Me.Caption
    DataRows = Module1.LookupData(Customer)
    Module1.CopyLookupKeys Customer, Prepared
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
    If count > VisibleRows Then ResultScroll.Max = count - VisibleRows Else ResultScroll.Max = 0
    ResultScroll.Value = 0: ResultScroll.Enabled = (count > VisibleRows)
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
        If slot <= VisibleRows And index <= ResultCount Then
            source = FilteredRows(index)
            DisplayRow slot, index, CStr(DataRows(source, 2)), CStr(DataRows(source, 1)), (index = SelectedIndex)
        Else
            DisplayRow slot, 0, "", "", False
        End If
    Next slot
End Sub

Public Sub SelectResult(ByVal index As Long, ByVal confirm As Boolean)
    If index < 1 Or index > ResultCount Then Exit Sub
    SelectedIndex = index
    Drawing = True
    If index <= ResultScroll.Value Then ResultScroll.Value = index - 1
    If index > ResultScroll.Value + VisibleRows Then ResultScroll.Value = index - VisibleRows
    Drawing = False
    RenderResults
    If confirm Then ConfirmChoice Else Me.TextBox1.SetFocus
End Sub
Private Sub ResultScroll_Change()
    If Not Ready Or Drawing Then Exit Sub
    If ResultCount > 0 Then
        If SelectedIndex <= ResultScroll.Value Or SelectedIndex > ResultScroll.Value + VisibleRows Then
            SelectedIndex = ResultScroll.Value + 1
        End If
    End If
    RenderResults
End Sub
Private Sub ResultScroll_Scroll()
    ResultScroll_Change
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
        If ResultCount > 0 Then SelectResult WorksheetFunction.Min(ResultCount, SelectedIndex + VisibleRows), False
        KeyCode = 0
    ElseIf KeyCode = vbKeyPageUp Then
        If ResultCount > 0 Then SelectResult WorksheetFunction.Max(1, SelectedIndex - VisibleRows), False
        KeyCode = 0
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = 0: ConfirmChoice
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0: ChosenRow = 0: Me.Hide
    End If
End Sub
Private Sub BindRow(ByVal slot As Long, ByVal title As MSForms.Label, ByVal code As MSForms.Label)
    Select Case slot
    Case 1: Set Name01 = title: Set Code01 = code
    Case 2: Set Name02 = title: Set Code02 = code
    Case 3: Set Name03 = title: Set Code03 = code
    Case 4: Set Name04 = title: Set Code04 = code
    Case 5: Set Name05 = title: Set Code05 = code
    Case 6: Set Name06 = title: Set Code06 = code
    Case 7: Set Name07 = title: Set Code07 = code
    Case 8: Set Name08 = title: Set Code08 = code
    Case 9: Set Name09 = title: Set Code09 = code
    Case 10: Set Name10 = title: Set Code10 = code
    Case 11: Set Name11 = title: Set Code11 = code
    End Select
End Sub

Private Sub DisplayRow(ByVal slot As Long, ByVal index As Long, ByVal title As String, ByVal code As String, ByVal selected As Boolean)
    Dim background As Long, foreground As Long
    SlotIndexes(slot) = index
    NameRows(slot).Visible = (index > 0): CodeRows(slot).Visible = (index > 0)
    If index = 0 Then Exit Sub
    If selected Then
        background = RGB(226, 237, 249): foreground = RGB(23, 54, 93)
    Else
        If index Mod 2 = 0 Then background = RGB(245, 248, 252) Else background = RGB(255, 255, 255)
        foreground = RGB(38, 55, 74)
    End If
    NameRows(slot).BackColor = background: CodeRows(slot).BackColor = background
    NameRows(slot).ForeColor = foreground: CodeRows(slot).ForeColor = foreground
    NameRows(slot).Caption = ChrW(&H200F) & title & "  "
    CodeRows(slot).Caption = code
    NameRows(slot).ControlTipText = title
    CodeRows(slot).ControlTipText = code
End Sub

Private Sub ChooseSlot(ByVal slot As Long, ByVal confirm As Boolean)
    If SlotIndexes(slot) > 0 Then SelectResult SlotIndexes(slot), confirm
End Sub

Private Sub Name01_Click()
    ChooseSlot 1, False
End Sub

Private Sub Name01_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 1, True
End Sub

Private Sub Code01_Click()
    ChooseSlot 1, False
End Sub

Private Sub Code01_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 1, True
End Sub

Private Sub Name02_Click()
    ChooseSlot 2, False
End Sub

Private Sub Name02_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 2, True
End Sub

Private Sub Code02_Click()
    ChooseSlot 2, False
End Sub

Private Sub Code02_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 2, True
End Sub

Private Sub Name03_Click()
    ChooseSlot 3, False
End Sub

Private Sub Name03_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 3, True
End Sub

Private Sub Code03_Click()
    ChooseSlot 3, False
End Sub

Private Sub Code03_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 3, True
End Sub

Private Sub Name04_Click()
    ChooseSlot 4, False
End Sub

Private Sub Name04_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 4, True
End Sub

Private Sub Code04_Click()
    ChooseSlot 4, False
End Sub

Private Sub Code04_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 4, True
End Sub

Private Sub Name05_Click()
    ChooseSlot 5, False
End Sub

Private Sub Name05_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 5, True
End Sub

Private Sub Code05_Click()
    ChooseSlot 5, False
End Sub

Private Sub Code05_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 5, True
End Sub

Private Sub Name06_Click()
    ChooseSlot 6, False
End Sub

Private Sub Name06_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 6, True
End Sub

Private Sub Code06_Click()
    ChooseSlot 6, False
End Sub

Private Sub Code06_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 6, True
End Sub

Private Sub Name07_Click()
    ChooseSlot 7, False
End Sub

Private Sub Name07_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 7, True
End Sub

Private Sub Code07_Click()
    ChooseSlot 7, False
End Sub

Private Sub Code07_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 7, True
End Sub

Private Sub Name08_Click()
    ChooseSlot 8, False
End Sub

Private Sub Name08_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 8, True
End Sub

Private Sub Code08_Click()
    ChooseSlot 8, False
End Sub

Private Sub Code08_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 8, True
End Sub

Private Sub Name09_Click()
    ChooseSlot 9, False
End Sub

Private Sub Name09_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 9, True
End Sub

Private Sub Code09_Click()
    ChooseSlot 9, False
End Sub

Private Sub Code09_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 9, True
End Sub

Private Sub Name10_Click()
    ChooseSlot 10, False
End Sub

Private Sub Name10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 10, True
End Sub

Private Sub Code10_Click()
    ChooseSlot 10, False
End Sub

Private Sub Code10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 10, True
End Sub

Private Sub Name11_Click()
    ChooseSlot 11, False
End Sub

Private Sub Name11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 11, True
End Sub

Private Sub Code11_Click()
    ChooseSlot 11, False
End Sub

Private Sub Code11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ChooseSlot 11, True
End Sub

Public Sub ReleaseRows()
    Dim i As Long
    On Error Resume Next
    For i = 1 To PAGE_ROWS
        BindRow i, Nothing, Nothing
        Set NameRows(i) = Nothing: Set CodeRows(i) = Nothing
    Next i
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = 1: ChosenRow = 0: Me.Hide
    End If
End Sub
