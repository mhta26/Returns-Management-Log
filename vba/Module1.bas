Attribute VB_Name = "Module1"
Option Explicit
Public BusyPrinting As Boolean
Private Const MAIN As Long = 1
Private Const CLIENTS As Long = 2
Private Const PRODUCTS As Long = 3
Private Const SETTINGS As Long = 4

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal text As LongPtr, ByVal caption As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal path As LongPtr) As Long
Private Declare PtrSafe Function ShellExecuteW Lib "shell32" (ByVal hwnd As LongPtr, ByVal operation As LongPtr, ByVal file As LongPtr, ByVal parameters As LongPtr, ByVal directory As LongPtr, ByVal show As Long) As LongPtr
#Else
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal text As Long, ByVal caption As Long, ByVal flags As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal path As Long) As Long
Private Declare Function ShellExecuteW Lib "shell32" (ByVal hwnd As Long, ByVal operation As Long, ByVal file As Long, ByVal parameters As Long, ByVal directory As Long, ByVal show As Long) As Long
#End If

Private Function UnicodeMessage(ByVal prompt As String, Optional ByVal buttons As Long = 0) As Long
    Dim title As String
    title = ThisWorkbook.Name
    UnicodeMessage = MessageBoxW(Application.hwnd, StrPtr(prompt), StrPtr(title), buttons Or &H180000)
End Function

Private Sub OpenUnicodePDF(ByVal path As String)
    Dim operation As String
#If VBA7 Then
    Dim result As LongPtr
#Else
    Dim result As Long
#End If
    operation = "open"
    result = ShellExecuteW(Application.hwnd, StrPtr(operation), StrPtr(path), 0, 0, 1)
    If result <= 32 Then Err.Raise vbObjectError + 926, "OpenUnicodePDF", UI(26)
End Sub


Private Function WS(ByVal n As Long) As Worksheet
    Set WS = ThisWorkbook.Worksheets(n)
End Function

Private Function UI(ByVal n As Long) As String
    UI = CStr(WS(SETTINGS).Cells(n, 8).value)
End Function

Private Function LastRow(ByVal Sh As Worksheet, ByVal cols As Long) As Long
    Dim c As Long, r As Long
    LastRow = 1
    For c = 1 To cols
        r = Sh.Cells(Sh.Rows.count, c).End(xlUp).row
        If r > LastRow Then LastRow = r
    Next c
End Function

Private Function Normal(ByVal s As String) As String
    Dim i As Long
    s = LCase$(Trim$(s))
    s = Replace(s, ChrW(1571), ChrW(1575))
    s = Replace(s, ChrW(1573), ChrW(1575))
    s = Replace(s, ChrW(1570), ChrW(1575))
    s = Replace(s, ChrW(1609), ChrW(1610))
    s = Replace(s, ChrW(1577), ChrW(1607))
    s = Replace(s, ChrW(1600), "")
    For i = 1611 To 1618
        s = Replace(s, ChrW(i), "")
    Next i
    Normal = s
End Function

Private Function Matches(ByVal haystack As String, ByVal query As String) As Boolean
    Dim token As Variant
    haystack = Normal(haystack)
    For Each token In Split(Normal(query), " ")
        If Len(token) > 0 Then
            If InStr(1, haystack, CStr(token), vbTextCompare) = 0 Then Exit Function
        End If
    Next token
    Matches = True
End Function

Private Function CodeRow(ByVal src As Worksheet, ByVal code As String) As Long
    Dim r As Long, found As Long
    If Len(code) = 0 Then Exit Function
    For r = 2 To LastRow(src, 2)
        If CStr(src.Cells(r, 1).Value2) = code And Len(CStr(src.Cells(r, 2).Value2)) > 0 Then
            If found > 0 Then CodeRow = -1: Exit Function
            found = r
        End If
    Next r
    CodeRow = found
End Function

Private Function CellText(ByVal cell As Range) As String
    If IsError(cell.Value2) Then
        CellText = "#ERROR"
    Else
        CellText = Trim$(Replace(Replace(CStr(cell.Value2), ChrW(160), " "), vbTab, " "))
    End If
End Function

Private Sub FormRows(ByRef firstRow As Long, ByRef totalRow As Long)
    Dim r As Long, lastLabelRow As Long, title As String, header As Range
    firstRow = 0: totalRow = 0
    lastLabelRow = WS(MAIN).Cells(WS(MAIN).Rows.count, 1).End(xlUp).row
    For r = 1 To lastLabelRow
        title = Normal(CellText(WS(MAIN).Cells(r, 1)))
        If title = Normal(UI(31)) Then
            If totalRow <> 0 Then Err.Raise vbObjectError + 901, "FormRows", UI(29)
            totalRow = r
        End If
    Next r
    If totalRow = 0 Then Err.Raise vbObjectError + 902, "FormRows", UI(29)
    For r = 1 To totalRow - 1
        If Normal(CellText(WS(MAIN).Cells(r, 2))) = Normal(UI(30)) Then
            Set header = WS(MAIN).Cells(r, 2).MergeArea
            firstRow = header.row + header.Rows.count
            Exit For
        End If
    Next r
    If firstRow = 0 Or firstRow >= totalRow Then Err.Raise vbObjectError + 903, "FormRows", UI(29)
End Sub

Private Function FirstItemRow() As Long
    Dim firstRow As Long, totalRow As Long
    FormRows firstRow, totalRow
    FirstItemRow = firstRow
End Function

Private Function LastItemRow() As Long
    Dim firstRow As Long, totalRow As Long
    FormRows firstRow, totalRow
    LastItemRow = totalRow - 1
End Function

Private Function ItemCells(ByVal firstCol As Long, ByVal lastCol As Long) As Range
    Dim firstRow As Long, totalRow As Long
    FormRows firstRow, totalRow
    Set ItemCells = WS(MAIN).Range(WS(MAIN).Cells(firstRow, firstCol), WS(MAIN).Cells(totalRow - 1, lastCol))
End Function

Private Function RowHasData(ByVal r As Long) As Boolean
    Dim c As Long, value As String
    For c = 2 To 13
        If c <> 11 Then
            value = CellText(WS(MAIN).Cells(r, c))
            If Len(value) > 0 Then
                If c >= 7 And c <= 9 Then
                    If Not IsNumeric(value) Then RowHasData = True: Exit Function
                    If CDbl(value) <> 0 Then RowHasData = True: Exit Function
                Else
                    RowHasData = True
                    Exit Function
                End If
            End If
        End If
    Next c
End Function

Private Function HasList(ByVal cell As Range) As Boolean
    On Error Resume Next
    HasList = (cell.Validation.Type = xlValidateList)
    On Error GoTo 0
End Function

Private Sub ColorStatuses(ByVal firstRow As Long, ByVal totalRow As Long)
    Dim target As Range, rule As FormatCondition, i As Long, colors As Variant
    Set target = WS(MAIN).Range(WS(MAIN).Cells(firstRow, 11), WS(MAIN).Cells(totalRow - 1, 11))
    colors = Array(RGB(255, 199, 206), RGB(248, 203, 173), RGB(255, 235, 156), RGB(198, 239, 206))
    target.FormatConditions.Delete
    For i = 0 To 3
        Set rule = target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & Chr$(34) & UI(62 + i) & Chr$(34))
        rule.Interior.Color = colors(i)
        rule.Font.Color = RGB(0, 0, 0)
    Next i
End Sub

Private Sub EnsureItemRows(ByVal firstRow As Long, ByVal totalRow As Long)
    Dim r As Long
    For r = firstRow To totalRow - 1
        WS(MAIN).Cells(r, 1).Value2 = r - firstRow + 1
        WS(MAIN).Cells(r, 11).FormulaR1C1 = UI(37)
        If Not HasList(WS(MAIN).Cells(r, 2)) Then SetList WS(MAIN).Cells(r, 2), "InlineProducts"
        If Not HasList(WS(MAIN).Cells(r, 3)) Then SetList WS(MAIN).Cells(r, 3), "InlineProducts"
        With WS(MAIN).Cells(r, 12).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="=ReturnReasons"
            .IgnoreBlank = True
            .ShowError = True
        End With
    Next r
    ColorStatuses firstRow, totalRow
End Sub

Private Sub RefreshTotals(ByVal firstRow As Long, ByVal totalRow As Long)
    Dim c As Long
    WS(MAIN).Range("I6").Formula = "=NOW()"
    EnsureItemRows firstRow, totalRow
    For c = 7 To 9
        WS(MAIN).Cells(totalRow, c).Formula = "=SUM(" & WS(MAIN).Range(WS(MAIN).Cells(firstRow, c), WS(MAIN).Cells(totalRow - 1, c)).address(False, False) & ")"
    Next c
    WS(MAIN).Range("J8").Formula = "=SUMPRODUCT(--(" & WS(MAIN).Range(WS(MAIN).Cells(firstRow, 2), WS(MAIN).Cells(totalRow - 1, 2)).address(False, False) & "<>""""))"
    WS(MAIN).PageSetup.PrintArea = WS(MAIN).Range("A1:M" & totalRow).address
    WS(MAIN).Calculate
End Sub

Private Function HasTopSearch() As Boolean
    HasTopSearch = False
End Function

Private Function SearchInputCells() As Range
    Set SearchInputCells = Application.Union(WS(MAIN).Range("A8,B8"), ItemCells(2, 3))
    If HasTopSearch() Then Set SearchInputCells = Application.Union(SearchInputCells, WS(MAIN).Range("D12,G12"))
End Function

Private Function ProductSearchTarget() As Range
    Dim row As Long
    If HasTopSearch() Then
        Set ProductSearchTarget = WS(MAIN).Range("G12")
        Exit Function
    End If
    For row = FirstItemRow() To LastItemRow()
        If Len(CellText(WS(MAIN).Cells(row, 2))) = 0 And Len(CellText(WS(MAIN).Cells(row, 3))) = 0 Then
            Set ProductSearchTarget = WS(MAIN).Cells(row, 3)
            Exit Function
        End If
    Next row
    Set ProductSearchTarget = WS(MAIN).Cells(FirstItemRow(), 3)
End Function

Private Sub RepairTableHeadings(ByVal firstRow As Long)
    ' Headings belong to the final design and must never be overwritten.
End Sub

Private Function IsClientCell(ByVal dest As Range) As Boolean
    IsClientCell = (dest.row = 8 And (dest.Column = 1 Or dest.Column = 2))
End Function

Private Function CandidateSource(ByVal customer As Boolean) As Worksheet
    If customer Then Set CandidateSource = WS(CLIENTS) Else Set CandidateSource = WS(PRODUCTS)
End Function

Private Sub SetList(ByVal dest As Range, ByVal listName As String)
    With dest.MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
        .ShowInput = True
        .InputTitle = UI(28)
        .InputMessage = Left$(UI(4), 255)
    End With
End Sub

Private Function BuildList(ByVal customer As Boolean, ByVal query As String) As Long
    Dim src As Worksheet, data As Variant, result() As Variant
    Dim i As Long, n As Long, count As Long, c As Long, oldLast As Long, listName As String
    Set src = CandidateSource(customer)
    If customer Then c = 27: listName = "InlineClients" Else c = 28: listName = "InlineProducts"
    With WS(MAIN)
        oldLast = .Cells(.Rows.count, c).End(xlUp).row
        If oldLast >= 2 Then .Range(.Cells(2, c), .Cells(oldLast, c)).ClearContents
    End With
    n = LastRow(src, 2)
    If n >= 2 Then
        data = src.Range("A2:B" & n).Value2
        For i = 1 To UBound(data, 1)
            If Len(CStr(data(i, 1))) > 0 And Len(CStr(data(i, 2))) > 0 Then
                If Matches(CStr(data(i, 2)) & " " & CStr(data(i, 1)), query) Then count = count + 1
            End If
        Next i
        If count > 0 Then
            ReDim result(1 To count, 1 To 1)
            count = 0
            For i = 1 To UBound(data, 1)
                If Len(CStr(data(i, 1))) > 0 And Len(CStr(data(i, 2))) > 0 Then
                    If Matches(CStr(data(i, 2)) & " " & CStr(data(i, 1)), query) Then
                        count = count + 1
                        result(count, 1) = Left$(CStr(data(i, 2)), 220) & " | " & CStr(data(i, 1))
                    End If
                End If
            Next i
            WS(MAIN).Cells(2, c).Resize(count, 1).Value2 = result
        End If
    End If
    If count = 0 Then WS(MAIN).Cells(2, c).value = UI(5)
    On Error Resume Next
    ThisWorkbook.Names(listName).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=listName, RefersTo:="='" & Replace(WS(MAIN).Name, "'", "''") & "'!" & WS(MAIN).Cells(2, c).Resize(IIf(count > 0, count, 1), 1).address
    WS(MAIN).Range("AA:AB").EntireColumn.Hidden = True
    BuildList = count
End Function

Private Sub ShowList(ByVal dest As Range)
    dest.Select
    If ActiveSheet.Name = WS(MAIN).Name Then Application.SendKeys "%{DOWN}", False
End Sub

Private Sub FillNames()
    Dim r As Long, found As Long
    found = CodeRow(WS(CLIENTS), CStr(WS(MAIN).Range("A8").Value2))
    If found > 0 Then WS(MAIN).Range("B8").value = WS(CLIENTS).Cells(found, 2).Value2
    If HasTopSearch() Then
        found = CodeRow(WS(PRODUCTS), CStr(WS(MAIN).Range("D12").Value2))
        If found > 0 Then WS(MAIN).Range("G12").value = WS(PRODUCTS).Cells(found, 2).Value2
    End If
    For r = FirstItemRow() To LastItemRow()
        found = CodeRow(WS(PRODUCTS), CStr(WS(MAIN).Cells(r, 2).Value2))
        If found > 0 Then WS(MAIN).Cells(r, 3).value = WS(PRODUCTS).Cells(found, 2).Value2
    Next r
End Sub

Private Sub ApplyChoice(ByVal dest As Range, ByVal sourceRow As Long)
    Dim customer As Boolean, src As Worksheet, r As Long
    customer = IsClientCell(dest)
    Set src = CandidateSource(customer)
    If customer Then
        WS(MAIN).Range("A8").NumberFormat = "@"
        WS(MAIN).Range("A8").value = CStr(src.Cells(sourceRow, 1).Value2)
        WS(MAIN).Range("B8").value = src.Cells(sourceRow, 2).Value2
    ElseIf HasTopSearch() And dest.row = 12 Then
        WS(MAIN).Range("D12").NumberFormat = "@"
        WS(MAIN).Range("D12").value = CStr(src.Cells(sourceRow, 1).Value2)
        WS(MAIN).Range("G12").value = src.Cells(sourceRow, 2).Value2
        For r = FirstItemRow() To LastItemRow()
            If Len(CStr(WS(MAIN).Cells(r, 2).Value2)) = 0 And Len(CStr(WS(MAIN).Cells(r, 3).Value2)) = 0 Then Exit For
        Next r
        If r > LastItemRow() Then
            UnicodeMessage UI(2), vbInformation
        Else
            WS(MAIN).Cells(r, 2).NumberFormat = "@"
            WS(MAIN).Cells(r, 2).value = CStr(src.Cells(sourceRow, 1).Value2)
            WS(MAIN).Cells(r, 3).value = src.Cells(sourceRow, 2).Value2
            WS(MAIN).Cells(r, 7).Select
        End If
    Else
        r = dest.row
        WS(MAIN).Cells(r, 2).NumberFormat = "@"
        WS(MAIN).Cells(r, 2).value = CStr(src.Cells(sourceRow, 1).Value2)
        WS(MAIN).Cells(r, 3).value = src.Cells(sourceRow, 2).Value2
        WS(MAIN).Cells(r, 7).Select
    End If
    BuildList customer, ""
    WS(MAIN).Calculate
End Sub

Private Sub SearchChanged(ByVal dest As Range)
    Dim customer As Boolean, src As Worksheet, query As String, code As String
    Dim p As Long, found As Long, count As Long, c As Long, listName As String
    customer = IsClientCell(dest)
    Set src = CandidateSource(customer)
    query = CellText(dest)
    If customer Then c = 27: listName = "InlineClients" Else c = 28: listName = "InlineProducts"
    If customer Then
        If dest.Column = 2 Then WS(MAIN).Range("A8").MergeArea.ClearContents Else WS(MAIN).Range("B8").MergeArea.ClearContents
    ElseIf HasTopSearch() And dest.row = 12 Then
        If dest.Column = 7 Then WS(MAIN).Range("D12").MergeArea.ClearContents Else WS(MAIN).Range("G12").MergeArea.ClearContents
    Else
        If dest.Column = 3 Then WS(MAIN).Cells(dest.row, 2).ClearContents Else WS(MAIN).Cells(dest.row, 3).MergeArea.ClearContents
    End If
    If Len(query) = 0 Then
        BuildList customer, ""
        SetList dest, listName
        Exit Sub
    End If
    code = query
    p = InStrRev(query, " | ")
    If p > 0 Then code = Mid$(query, p + 3)
    found = CodeRow(src, code)
    If found = -1 Then UnicodeMessage UI(8), vbExclamation: Exit Sub
    If found > 0 Then
        ApplyChoice dest, found
        Exit Sub
    End If
    count = BuildList(customer, query)
    SetList dest, listName
    If count = 0 Then
        UnicodeMessage UI(5), vbInformation
    ElseIf count = 1 Then
        code = CStr(WS(MAIN).Cells(2, c).Value2)
        code = Mid$(code, InStrRev(code, " | ") + 3)
        found = CodeRow(src, code)
        If found > 0 Then ApplyChoice dest, found Else UnicodeMessage UI(8), vbExclamation
    Else
        ShowList dest
    End If
End Sub

Private Sub StartSearch(ByVal dest As Range)
    Dim oldEvents As Boolean, listName As String, customer As Boolean
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    customer = IsClientCell(dest)
    If customer Then listName = "InlineClients" Else listName = "InlineProducts"
    BuildList customer, ""
    SetList dest, listName
    ShowList dest
Done:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    UnicodeMessage UI(6) & vbCrLf & Err.Description, vbExclamation
    Resume Done
End Sub

Public Sub SearchProducts()
    If ActiveSheet.Name <> WS(MAIN).Name Then Exit Sub
    If ActiveCell.row >= FirstItemRow() And ActiveCell.row <= LastItemRow() And (ActiveCell.Column = 2 Or ActiveCell.Column = 3) Then
        StartSearch WS(MAIN).Cells(ActiveCell.row, 3)
    Else
        StartSearch ProductSearchTarget()
    End If
End Sub

Public Sub SearchClients()
    If ActiveSheet.Name = WS(MAIN).Name Then StartSearch WS(MAIN).Range("B8")
End Sub

Public Sub HandleChange(ByVal Sh As Object, ByVal target As Range)
    Dim oldEvents As Boolean, dest As Range, firstRow As Long, totalRow As Long
    Dim changed As Range, cell As Range, found As Long, customer As Boolean, src As Worksheet
    Dim query As String, splitAt As Long
    If BusyPrinting Then Exit Sub
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    If Sh.Name = WS(MAIN).Name Then
        FormRows firstRow, totalRow
        Set dest = target.Cells(1, 1).MergeArea.Cells(1, 1)
        Set changed = Intersect(target, SearchInputCells())
        If Not changed Is Nothing Then
            If target.CountLarge = 1 Or target.address = dest.MergeArea.address Then
                SearchChanged dest
            Else
                For Each cell In changed.Cells
                    If cell.address = cell.MergeArea.Cells(1, 1).address Then
                        customer = IsClientCell(cell)
                        Set src = CandidateSource(customer)
                        query = CellText(cell)
                        splitAt = InStrRev(query, " | ")
                        If splitAt > 0 Then query = Mid$(query, splitAt + 3)
                        If Len(query) = 0 Then
                            If customer Then
                                If cell.Column = 1 Then Sh.Range("B8").MergeArea.ClearContents Else Sh.Range("A8").ClearContents
                            Else
                                If cell.Column = 2 Then Sh.Cells(cell.row, 3).MergeArea.ClearContents Else Sh.Cells(cell.row, 2).ClearContents
                            End If
                        Else
                            found = CodeRow(src, query)
                            If found > 0 Then ApplyChoice cell, found
                        End If
                    End If
                Next cell
            End If
        End If
        RefreshTotals firstRow, totalRow
    ElseIf Sh.Name = WS(CLIENTS).Name Or Sh.Name = WS(PRODUCTS).Name Then
        FillNames
        BuildList True, ""
        BuildList False, ""
        WS(MAIN).Calculate
    End If
Done:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    UnicodeMessage UI(6) & vbCrLf & Err.Description, vbExclamation
    Resume Done
End Sub

Public Sub HandleDoubleClick(ByVal Sh As Object, ByVal target As Range, ByRef Cancel As Boolean)
    If Sh.Name = WS(MAIN).Name Then
        If Not Intersect(target, Sh.Range("L6").MergeArea) Is Nothing Then
            Cancel = True
            ArchiveAndPrint
        ElseIf Not Intersect(target, Sh.Range("A8:D8")) Is Nothing Then
            Cancel = True
            SearchClients
        ElseIf HasTopSearch() And Not Intersect(target, Sh.Range("A12:M12")) Is Nothing Then
            Cancel = True
            StartSearch ProductSearchTarget()
        ElseIf Not Intersect(target, ItemCells(2, 3)) Is Nothing Then
            Cancel = True
            StartSearch WS(MAIN).Cells(target.row, 3)
        End If
    ElseIf IsSavedReceiptSheet(Sh.Name) Then
        If Not Intersect(target, Sh.Range("J11:K11")) Is Nothing Then
            Cancel = True
            ReprintSavedSheet Sh.Name
        End If
    ElseIf Sh.Name = "Archive" Then
        If target.row >= 4 And target.Column = 9 Then
            Cancel = True
            ReprintArchive
        End If
    End If
End Sub

Public Sub Initialize()
    Dim oldEvents As Boolean, r As Long, firstRow As Long, totalRow As Long
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    ThisWorkbook.Names.Add Name:="ReturnReasons", RefersTo:="='" & Replace(WS(SETTINGS).Name, "'", "''") & "'!$A$2:$A$7"
    FormRows firstRow, totalRow
    RepairTableHeadings firstRow
    RefreshTotals firstRow, totalRow
    WS(SETTINGS).Range("D1").value = NextNumber()
    WS(MAIN).Range("L8").Formula = "=""#""&'" & Replace(WS(SETTINGS).Name, "'", "''") & "'!$D$1"
    FillNames
    BuildList True, ""
    BuildList False, ""
    SetList WS(MAIN).Range("A8"), "InlineClients"
    SetList WS(MAIN).Range("B8"), "InlineClients"
    If HasTopSearch() Then
        SetList WS(MAIN).Range("D12"), "InlineProducts"
        SetList WS(MAIN).Range("G12"), "InlineProducts"
    End If
    For r = FirstItemRow() To LastItemRow()
        SetList WS(MAIN).Cells(r, 2), "InlineProducts"
        SetList WS(MAIN).Cells(r, 3), "InlineProducts"
    Next r
    WS(MAIN).Calculate
Done:
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    UnicodeMessage UI(6) & vbCrLf & Err.Description, vbExclamation
    Resume Done
End Sub

Private Function NextNumber() As Long
    Dim Sh As Worksheet, n As Long, suffix As String, log As Worksheet, row As Long
    n = 1
    If IsNumeric(WS(SETTINGS).Range("D1").value) Then
        If CDbl(WS(SETTINGS).Range("D1").value) >= 1 Then n = CLng(WS(SETTINGS).Range("D1").value)
    End If
    If IsNumeric(WS(SETTINGS).Range("C2").Value2) Then
        If CDbl(WS(SETTINGS).Range("C2").Value2) > n Then n = CLng(WS(SETTINGS).Range("C2").Value2)
    End If
    For Each Sh In ThisWorkbook.Worksheets
        If Left$(Sh.Name, 4) = "RET_" Then
            suffix = Mid$(Sh.Name, 5)
            If IsNumeric(suffix) Then
                If CLng(suffix) >= n Then n = CLng(suffix) + 1
            End If
        End If
    Next Sh
    Set log = ThisWorkbook.Worksheets("Archive")
    For row = 4 To LastRow(log, 10)
        If IsNumeric(log.Cells(row, 1).Value2) Then
            If CDbl(log.Cells(row, 1).Value2) >= n Then n = CLng(log.Cells(row, 1).Value2) + 1
        End If
    Next row
    NextNumber = n
End Function

Private Function ValidForm() As Boolean
    Dim r As Long, c As Long, firstRow As Long, totalRow As Long
    Dim quantity As Double, entries As Long, code As String, value As String
    code = CellText(WS(MAIN).Range("A8"))
    If Len(code) = 0 Then UnicodeMessage UI(10), vbExclamation: Exit Function
    If CodeRow(WS(CLIENTS), code) <= 0 Then UnicodeMessage UI(8), vbExclamation: Exit Function
    If Len(CellText(WS(MAIN).Range("F6"))) > 0 Then
        If Not IsNumeric(WS(MAIN).Range("F6").Value2) Then UnicodeMessage UI(58), vbExclamation: Exit Function
        If CDbl(WS(MAIN).Range("F6").Value2) <= 0 Or CDbl(WS(MAIN).Range("F6").Value2) > CDbl(Date) Then UnicodeMessage UI(58), vbExclamation: Exit Function
    End If
    FormRows firstRow, totalRow
    For r = firstRow To totalRow - 1
        If RowHasData(r) Then
            code = CellText(WS(MAIN).Cells(r, 2))
            If Len(code) = 0 Then UnicodeMessage UI(11) & " " & r, vbExclamation: Exit Function
            If CodeRow(WS(PRODUCTS), code) <= 0 Then UnicodeMessage UI(11) & " " & r, vbExclamation: Exit Function
            If Not IsNumeric(WS(MAIN).Cells(r, 10).Value2) Or Len(CellText(WS(MAIN).Cells(r, 10))) = 0 Then
                UnicodeMessage UI(57) & " " & r, vbExclamation: Exit Function
            End If
            If CDbl(WS(MAIN).Cells(r, 10).Value2) <= 0 Or CDbl(WS(MAIN).Cells(r, 10).Value2) > 2958465 Then
                UnicodeMessage UI(57) & " " & r, vbExclamation: Exit Function
            End If
            entries = entries + 1
            quantity = 0
            For c = 7 To 9
                value = CellText(WS(MAIN).Cells(r, c))
                If Len(value) > 0 Then
                    If Not IsNumeric(value) Then UnicodeMessage UI(12) & " " & r, vbExclamation: Exit Function
                    If CDbl(value) < 0 Then UnicodeMessage UI(12) & " " & r, vbExclamation: Exit Function
                    If CDbl(value) <> Fix(CDbl(value)) Then UnicodeMessage UI(12) & " " & r, vbExclamation: Exit Function
                    quantity = quantity + CDbl(value)
                End If
            Next c
            If quantity = 0 Then UnicodeMessage UI(12) & " " & r, vbExclamation: Exit Function
        End If
    Next r
    If entries = 0 Then UnicodeMessage UI(13), vbExclamation: Exit Function
    ValidForm = True
End Function

Private Function ReceiptFileName(ByVal number As Long) As String
    Dim baseName As String, suffix As String, bad As Variant, i As Long
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then suffix = LCase$(Mid$(baseName, InStrRev(baseName, ".")))
    If suffix = ".xlsm" Or suffix = ".xlsx" Or suffix = ".xlsb" Or suffix = ".xls" Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    For Each bad In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        baseName = Replace(baseName, CStr(bad), "_")
    Next bad
    For i = 0 To 31
        baseName = Replace(baseName, Chr$(i), " ")
    Next i
    baseName = Trim$(baseName)
    Do While Len(baseName) > 0
        If Right$(baseName, 1) = "." Or Right$(baseName, 1) = " " Then baseName = Left$(baseName, Len(baseName) - 1) Else Exit Do
    Loop
    If Len(baseName) = 0 Then baseName = UI(38)
    ReceiptFileName = baseName & " - " & CStr(number) & ".pdf"
End Function

Private Function FolderPath() As String
    If Len(ThisWorkbook.path) = 0 Or InStr(1, ThisWorkbook.path, "://", vbTextCompare) > 0 Then Err.Raise vbObjectError + 920, "FolderPath", UI(49)
    FolderPath = ThisWorkbook.path
    If Right$(FolderPath, 1) <> Application.PathSeparator Then FolderPath = FolderPath & Application.PathSeparator
End Function

Private Function FileExists(ByVal path As String) As Boolean
    Dim attributes As Long
    attributes = GetFileAttributesW(StrPtr(path))
    FileExists = (attributes <> -1) And ((attributes And vbDirectory) = 0)
End Function

Public Function IsSavedReceiptSheet(ByVal sheetName As String) As Boolean
    Dim suffix As String
    If Left$(sheetName, 4) <> "RET_" Then Exit Function
    suffix = Mid$(sheetName, 5)
    If Len(suffix) = 0 Then Exit Function
    If IsNumeric(suffix) Then IsSavedReceiptSheet = (CDbl(suffix) > 0)
End Function

Private Sub CloseReceiptCopy(ByRef temporary As Workbook)
    On Error Resume Next
    Application.CutCopyMode = False
    If Not temporary Is Nothing Then temporary.Close SaveChanges:=False
    Set temporary = Nothing
    On Error GoTo 0
End Sub

Private Sub PrepareReceiptCopy(ByVal number As Long, ByVal totalRow As Long, ByRef temporary As Workbook, ByRef receipt As Worksheet)
    Dim address As String, i As Long
    address = "A1:M" & totalRow
    WS(MAIN).Calculate
    WS(MAIN).Copy
    Set temporary = Application.ActiveWorkbook
    Set receipt = temporary.Worksheets(1)
    receipt.Range(address).Copy
    receipt.Range(address).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    receipt.Range("I6").value = Now
    receipt.Range("L8").value = "#" & CStr(number)
    receipt.Range("L6").MergeArea.ClearContents
    receipt.Range("AA:AB").ClearContents
    receipt.Range(address).Validation.Delete
    For i = receipt.Shapes.count To 1 Step -1
        If Left$(receipt.Shapes(i).Name, 4) = "btn_" Or receipt.Shapes(i).Type = msoFormControl Then receipt.Shapes(i).Delete
    Next i
    receipt.PageSetup.PrintArea = receipt.Range(address).address
End Sub

Private Sub ExportReceiptPDF(ByVal receipt As Worksheet, ByVal path As String)
    receipt.ExportAsFixedFormat Type:=xlTypePDF, fileName:=path, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Not FileExists(path) Then Err.Raise vbObjectError + 921, "ExportReceiptPDF", UI(56)
    If CreateObject("Scripting.FileSystemObject").GetFile(path).Size = 0 Then Err.Raise vbObjectError + 921, "ExportReceiptPDF", UI(56)
End Sub

Private Sub RestoreCells(ByVal target As Range, ByVal saved As Variant)
    Dim cell As Range
    For Each cell In target.Cells
        If cell.address = cell.MergeArea.Cells(1, 1).address Then
            cell.Formula = saved(cell.row - target.row + 1, cell.Column - target.Column + 1)
        End If
    Next cell
End Sub

Public Sub ArchiveAndPrint()
    Dim form As Worksheet, log As Worksheet, receipt As Worksheet, temporary As Workbook
    Dim firstRow As Long, totalRow As Long, number As Long, logRow As Long
    Dim pdfPath As String, relativeName As String, paper As Boolean, errorText As String, paperError As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCounter As Variant
    Dim itemBackup As Variant, noteBackup As Variant, clientCode As Variant, clientName As Variant, topCode As Variant, topName As Variant
    Dim copiesRequested As Variant, copyCount As Long, returnBackup As Variant, invoiceBackup As Variant
    Dim backupReady As Boolean, resetStarted As Boolean, pdfSaved As Boolean, committed As Boolean
    If BusyPrinting Then Exit Sub
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Set form = WS(MAIN)
    Set log = ThisWorkbook.Worksheets("Archive")
    If ThisWorkbook.ReadOnly Then UnicodeMessage UI(14), vbExclamation: Exit Sub
    pdfPath = FolderPath()
    FormRows firstRow, totalRow
    RepairTableHeadings firstRow
    RefreshTotals firstRow, totalRow
    If Not ValidForm() Then Exit Sub
    Do
        copiesRequested = Application.InputBox(Prompt:=UI(59), Title:=UI(60), Default:=1, Type:=1)
        If VarType(copiesRequested) = vbBoolean Then Exit Sub
        If copiesRequested >= 1 And copiesRequested <= 999 Then
            If copiesRequested = Fix(copiesRequested) Then Exit Do
        End If
        UnicodeMessage UI(61), vbExclamation
    Loop
    copyCount = CLng(copiesRequested)
    paper = True
    number = NextNumber()
    Do While FileExists(pdfPath & ReceiptFileName(number))
        number = number + 1
    Loop
    relativeName = ReceiptFileName(number)
    pdfPath = pdfPath & relativeName
    ' Keep a durable filled-in copy until the PDF is successfully exported.
    ThisWorkbook.Save
    oldCounter = WS(SETTINGS).Range("D1").Value2
    itemBackup = form.Range(form.Cells(firstRow, 2), form.Cells(totalRow - 1, 10)).Formula
    noteBackup = form.Range(form.Cells(firstRow, 12), form.Cells(totalRow - 1, 13)).Formula
    clientCode = form.Range("A8").Formula
    clientName = form.Range("B8").Formula
    If HasTopSearch() Then
        topCode = form.Range("D12").Formula
        topName = form.Range("G12").Formula
    End If
    returnBackup = form.Range("D6").Formula
    invoiceBackup = form.Range("F6").Formula
    backupReady = True
    BusyPrinting = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    WS(SETTINGS).Range("D1").Value2 = number
    form.Calculate
    PrepareReceiptCopy number, totalRow, temporary, receipt
    ExportReceiptPDF receipt, pdfPath
    pdfSaved = True
    logRow = LastRow(log, 10) + 1
    If logRow < 4 Then logRow = 4
    log.Cells(logRow, 1).Value2 = number
    log.Cells(logRow, 2).value = Now
    log.Cells(logRow, 2).NumberFormat = "yyyy-mm-dd hh:mm"
    log.Cells(logRow, 3).NumberFormat = "@"
    log.Cells(logRow, 3).Value2 = CStr(form.Range("A8").Value2)
    log.Cells(logRow, 4).Value2 = form.Range("B8").Value2
    log.Cells(logRow, 5).Value2 = form.Range("J8").Value2
    log.Cells(logRow, 6).Value2 = form.Cells(totalRow, 7).Value2
    log.Cells(logRow, 7).Value2 = form.Cells(totalRow, 8).Value2
    log.Cells(logRow, 8).Value2 = form.Cells(totalRow, 9).Value2
    log.Cells(logRow, 9).Value2 = UI(42)
    log.Hyperlinks.Add Anchor:=log.Cells(logRow, 10), address:=relativeName, TextToDisplay:=relativeName
    resetStarted = True
    form.Range("A8").MergeArea.ClearContents
    form.Range("B8").MergeArea.ClearContents
    If HasTopSearch() Then
        form.Range("D12").MergeArea.ClearContents
        form.Range("G12").MergeArea.ClearContents
    End If
    form.Range(form.Cells(firstRow, 2), form.Cells(totalRow - 1, 10)).ClearContents
    form.Range(form.Cells(firstRow, 12), form.Cells(totalRow - 1, 13)).ClearContents
    form.Range("D6").ClearContents
    form.Range("F6").ClearContents
    WS(SETTINGS).Range("D1").Value2 = number + 1
    BuildList True, ""
    BuildList False, ""
    RefreshTotals firstRow, totalRow
    form.Activate
    ThisWorkbook.Save
    committed = True
    If paper Then
        On Error Resume Next
        receipt.PrintOut Copies:=copyCount, Collate:=True
        If Err.number <> 0 Then paperError = Err.Description
        Err.Clear
        On Error GoTo Failed
        If Len(paperError) = 0 Then log.Cells(logRow, 9).Value2 = UI(45) Else log.Cells(logRow, 9).Value2 = UI(46)
        On Error Resume Next
        ThisWorkbook.Save
        On Error GoTo Failed
    End If
    CloseReceiptCopy temporary
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    BusyPrinting = False
    form.Activate
    If Len(paperError) > 0 Then
        UnicodeMessage UI(46) & vbCrLf & pdfPath & vbCrLf & paperError & vbCrLf & UI(20), vbExclamation
    End If
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not committed And backupReady Then
        If resetStarted Then
            RestoreCells form.Range(form.Cells(firstRow, 2), form.Cells(totalRow - 1, 10)), itemBackup
            RestoreCells form.Range(form.Cells(firstRow, 12), form.Cells(totalRow - 1, 13)), noteBackup
            form.Range("D6").Formula = returnBackup
            form.Range("F6").Formula = invoiceBackup
            form.Range("A8").Formula = clientCode
            form.Range("B8").Formula = clientName
            If HasTopSearch() Then
                form.Range("D12").Formula = topCode
                form.Range("G12").Formula = topName
            End If
        End If
        WS(SETTINGS).Range("D1").Value2 = oldCounter
        If logRow >= 4 Then log.Range(log.Cells(logRow, 1), log.Cells(logRow, 10)).Clear
        form.Calculate
    End If
    CloseReceiptCopy temporary
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    BusyPrinting = False
    If Not form Is Nothing Then form.Activate
    On Error GoTo 0
    If committed Then
        UnicodeMessage UI(46) & vbCrLf & pdfPath & vbCrLf & errorText & vbCrLf & UI(20), vbExclamation
    ElseIf pdfSaved Then
        UnicodeMessage UI(52) & vbCrLf & pdfPath & vbCrLf & errorText, vbExclamation
    Else
        UnicodeMessage UI(56) & vbCrLf & errorText, vbExclamation
    End If
End Sub

Private Sub OpenLogPDF(ByVal row As Long, Optional ByVal legacyName As String = "")
    Dim log As Worksheet, address As String, path As String, number As Long
    Dim receipt As Worksheet, oldEvents As Boolean, oldBusy As Boolean, errorText As String
    oldEvents = Application.EnableEvents
    oldBusy = BusyPrinting
    On Error GoTo Failed
    Set log = ThisWorkbook.Worksheets("Archive")
    If row < 4 Then Err.Raise vbObjectError + 922, "OpenLogPDF", UI(54)
    number = CLng(log.Cells(row, 1).Value2)
    If log.Cells(row, 10).Hyperlinks.count > 0 Then address = log.Cells(row, 10).Hyperlinks(1).address
    If Len(address) > 0 Then
        If LCase$(Right$(address, 4)) <> ".pdf" Then Err.Raise vbObjectError + 925, "OpenLogPDF", UI(54)
        path = address
        If InStr(1, path, ":", vbTextCompare) = 0 And Left$(path, 2) <> "\\" Then path = FolderPath() & path
        If Not FileExists(path) Then Err.Raise vbObjectError + 923, "OpenLogPDF", UI(53)
    Else
        If Len(legacyName) = 0 Then legacyName = CStr(log.Cells(row, 10).Value2)
        If Not IsSavedReceiptSheet(legacyName) Then Err.Raise vbObjectError + 924, "OpenLogPDF", UI(54)
        Set receipt = ThisWorkbook.Worksheets(legacyName)
        address = ReceiptFileName(number)
        path = FolderPath() & address
        Application.EnableEvents = False
        BusyPrinting = True
        If Not FileExists(path) Then ExportReceiptPDF receipt, path
        If Not ThisWorkbook.ReadOnly Then
            log.Cells(row, 10).Hyperlinks.Delete
            log.Hyperlinks.Add Anchor:=log.Cells(row, 10), address:=address, TextToDisplay:=address
            log.Cells(row, 9).Value2 = UI(42)
            ThisWorkbook.Save
        End If
    End If
    Application.EnableEvents = oldEvents
    BusyPrinting = oldBusy
    OpenUnicodePDF path
    Exit Sub
Failed:
    errorText = Err.Description
    Application.EnableEvents = oldEvents
    BusyPrinting = oldBusy
    UnicodeMessage UI(26) & vbCrLf & errorText, vbExclamation
End Sub

Public Sub ReprintArchive()
    Dim log As Worksheet
    Set log = ThisWorkbook.Worksheets("Archive")
    If ActiveSheet.Name <> log.Name Or ActiveCell.row < 4 Then UnicodeMessage UI(25), vbInformation: Exit Sub
    OpenLogPDF ActiveCell.row
End Sub

Public Sub ReprintSavedSheet(ByVal sheetName As String)
    Dim log As Worksheet, row As Long, number As Long
    If Not IsSavedReceiptSheet(sheetName) Then Exit Sub
    On Error GoTo Failed
    number = CLng(Mid$(sheetName, 5))
    Set log = ThisWorkbook.Worksheets("Archive")
    For row = 4 To LastRow(log, 10)
        If IsNumeric(log.Cells(row, 1).Value2) Then
            If CLng(log.Cells(row, 1).Value2) = number Then
                OpenLogPDF row, sheetName
                Exit Sub
            End If
        End If
    Next row
    UnicodeMessage UI(54), vbExclamation
    Exit Sub
Failed:
    UnicodeMessage UI(26) & vbCrLf & Err.Description, vbExclamation
End Sub
