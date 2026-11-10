Attribute VB_Name = "Module1"
Option Explicit
Public BusyPrinting As Boolean
Private Const MAIN As Long = 1
Private Const CLIENTS As Long = 2
Private Const PRODUCTS As Long = 3
Private Const SETTINGS As Long = 5
Private Const DRIVERS As Long = 4
Public BusySearching As Boolean
Private EntryChangeActive As Boolean, MouseCommit As Boolean

#If VBA7 Then

Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal key As Long) As Integer
Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hwnd As LongPtr, ByVal text As LongPtr, ByVal caption As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal path As LongPtr) As Long
Private Declare PtrSafe Function ShellExecuteW Lib "shell32" (ByVal hwnd As LongPtr, ByVal operation As LongPtr, ByVal file As LongPtr, ByVal parameters As LongPtr, ByVal directory As LongPtr, ByVal show As Long) As LongPtr
#Else
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal key As Long) As Integer
Private Declare Function MessageBoxW Lib "user32" (ByVal hwnd As Long, ByVal text As Long, ByVal caption As Long, ByVal flags As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal path As Long) As Long
Private Declare Function ShellExecuteW Lib "shell32" (ByVal hwnd As Long, ByVal operation As Long, ByVal file As Long, ByVal parameters As Long, ByVal directory As Long, ByVal show As Long) As Long
#End If

Private EntryModeActive As Boolean, PreviousReturnEnabled As Boolean, PreviousReturnDirection As Long
Private PendingEntryAddress As String
Private NotesReturnBound As Boolean

Private ClientCache As Variant, ProductCache As Variant
Private ClientCacheKeys() As String, ProductCacheKeys() As String
Private ClientCodeIndex As Object, ProductCodeIndex As Object
Private ClientCacheReady As Boolean, ProductCacheReady As Boolean
Private ClientCacheLast As Long, ProductCacheLast As Long
Private ClientListReady As Boolean, ProductListReady As Boolean
Private ClientListQuery As String, ProductListQuery As String
Private ClientListCount As Long, ProductListCount As Long
Private SearchStatusOwned As Boolean, OriginalSearchStatus As Variant


Private Function UnicodeMessage(ByVal prompt As String, Optional ByVal buttons As Long = 0) As Long
    Dim title As String
    title = ThisWorkbook.name
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


Private Function WS(ByVal n As Long) As Excel.Worksheet
    Dim sheet As Excel.Worksheet, expected As String
    Select Case n
    Case MAIN: expected = "Sheet1"
    Case CLIENTS: expected = "Sheet2"
    Case PRODUCTS: expected = "Sheet3"
    Case SETTINGS: expected = "Sheet4"
    Case DRIVERS
        For Each sheet In ThisWorkbook.Worksheets
            If sheet.name = DriverSheetName() Then Set WS = sheet: Exit Function
        Next sheet
    End Select
    For Each sheet In ThisWorkbook.Worksheets
        If sheet.CodeName = expected Then Set WS = sheet: Exit Function
    Next sheet
    Err.Raise vbObjectError + 940, "WS", "Required worksheet is missing."
End Function

Private Function DecodeText(ByVal encoded As String) As String
    Dim unit As Variant
    For Each unit In Split(encoded, " ")
        If Len(unit) > 0 Then DecodeText = DecodeText & ChrW(CLng("&H" & unit))
    Next unit
End Function

Private Function UI(ByVal n As Long) As String
    If n = 81 Then UI = DecodeText("062A 0645 0020 062D 0641 0638 0020 0627 0644 0633 0646 062F 0020 0648 0625 0631 0633 0627 0644 0647 0020 0644 0644 0637 0627 0628 0639 0629 060C 0020 0644 0643 0646 0020 062A 0639 0630 0631 0020 062D 0641 0638 0020 062A 062D 062F 064A 062B 0020 062D 0627 0644 0629 0020 0627 0644 0637 0628 0627 0639 0629 0020 0641 064A 0020 0627 0644 0623 0631 0634 064A 0641 002E "): Exit Function
    If n = 75 Then UI = DecodeText("0646 062A 0627 0626 062C 0020 0627 0644 0628 062D 062B 0020 2014 0020 0627 062E 062A 0631 0020 0645 0646 0020 0627 0644 0642 0627 0626 0645 0629 003A 0020"): Exit Function
    If n = 37 Then UI = WarehouseFormula(): Exit Function
    If n >= 62 And n <= 65 Then UI = Trim$(CStr(WS(SETTINGS).Cells(n - 60, 2).Value2)): Exit Function
    If n = 66 Then UI = DecodeText("0631 0627 062C 0639 0020 0642 0648 0627 0639 062F 0020 0627 0644 0645 062E 0627 0632 0646 003A 0020 062B 0644 0627 062B 0629 0020 062D 062F 0648 062F " & _
            "0020 0623 064A 0627 0645 0020 0635 062D 064A 062D 0629 0020 0645 062A 0632 0627 064A 062F 0629 0020 0648 0623 0631 0628 0639 0629 0020 0623 0633 0645 " & _
            "0627 0621 0020 0645 062E 062A 0644 0641 0629 0020 063A 064A 0631 0020 0641 0627 0631 063A 0629 002E"): Exit Function
    If n = 67 Then UI = DecodeText("062A 0627 0631 064A 062E 0020 063A 064A 0631 0020 0635 062D 064A 062D"): Exit Function
    If n = 68 Then UI = DecodeText("0627 062E 062A 0631 0020 0631 0642 0645 0020 0633 0627 0626 0642 0020 0635 062D 064A 062D 064B 0627 0020 0645 0646 0020 0642 0627 0639 062F 0629 0020 " & _
            "0627 0644 0633 0627 0626 0642 064A 0646 002E"): Exit Function
    If n = 69 Then UI = DecodeText("0623 062F 062E 0644 0020 0631 0642 0645 0020 0627 0644 0625 0631 062C 0627 0639 0020 0623 0648 0644 064B 0627 002E"): Exit Function
    If n = 70 Then UI = DecodeText("0623 062F 062E 0644 0020 062A 0627 0631 064A 062E 0020 0627 0644 0641 0627 062A 0648 0631 0629 0020 0623 0648 0644 064B 0627 002E"): Exit Function
    If n = 71 Then UI = DecodeText("0628 062D 062B 0020 0627 0644 0639 0645 064A 0644 0020 2014 0020 0627 0643 062A 0628 0020 0627 0644 0627 0633 0645 0020 0623 0648 0020 0627 0644 0643 " & _
            "0648 062F"): Exit Function
    If n = 72 Then UI = DecodeText("0628 062D 062B 0020 0627 0644 0635 0646 0641 0020 2014 0020 0627 0643 062A 0628 0020 0627 0644 0627 0633 0645 0020 0623 0648 0020 0627 0644 0643 0648 " & _
            "062F"): Exit Function
    If n = 73 Then UI = DecodeText("0627 062E 062A 064A 0627 0631"): Exit Function
    If n = 74 Then UI = DecodeText("0625 0644 063A 0627 0621"): Exit Function
    UI = CStr(WS(SETTINGS).Cells(n, 8).value)
    If Len(UI) > 0 Then Exit Function
    Select Case n
    Case 5
        UI = DecodeText("0644 0627 0020 062A 0648 062C 062F 0020 0646 062A 0627 0626 062C 0020 0645 0637 0627 0628 0642 0629 002E 0020 062A 062D 0642 0642 0020 0645 0646 0020 " & _
            "0627 0644 0627 0633 0645 0020 0623 0648 0020 0623 0636 0641 0020 0627 0644 0633 062C 0644 0020 0625 0644 0649 0020 0642 0627 0639 062F 0629 0020 0627 " & _
            "0644 0628 064A 0627 0646 0627 062A 002E")
    Case 6
        UI = DecodeText("062A 0639 0630 0631 0020 062A 062D 062F 064A 062B 0020 0646 062A 0627 0626 062C 0020 0627 0644 0628 062D 062B 002E")
    Case 8
        UI = DecodeText("0643 0648 062F 0020 0627 0644 0639 0645 064A 0644 0020 0623 0648 0020 0627 0644 0635 0646 0641 0020 0645 0641 0642 0648 062F 0020 0623 0648 0020 0645 " & _
            "0643 0631 0631 002E 0020 0631 0627 062C 0639 0020 0642 0627 0639 062F 0629 0020 0627 0644 0628 064A 0627 0646 0627 062A 0020 0648 062D 062F 0651 062B " & _
            "0020 0627 0644 0628 062D 062B 002E")
    Case 10
        UI = DecodeText("0627 062E 062A 0631 0020 0627 0644 0639 0645 064A 0644 0020 0642 0628 0644 0020 0627 0644 0637 0628 0627 0639 0629 002E")
    Case 11
        UI = DecodeText("0643 0648 062F 0020 0635 0646 0641 0020 0645 0641 0642 0648 062F 0020 0623 0648 0020 063A 064A 0631 0020 0635 062D 064A 062D 0020 0623 0648 0020 0645 " & _
            "0643 0631 0631 0020 0641 064A 0020 0627 0644 0635 0641 003A")
    Case 12
        UI = DecodeText("0623 062F 062E 0644 0020 0643 0645 064A 0627 062A 0020 0635 062D 064A 062D 0629 0020 063A 064A 0631 0020 0633 0627 0644 0628 0629 0020 0648 0628 0625 " & _
            "062C 0645 0627 0644 064A 0020 0623 0643 0628 0631 0020 0645 0646 0020 0635 0641 0631 0020 0641 064A 0020 0627 0644 0635 0641 003A")
    Case 13
        UI = DecodeText("0623 0636 0641 0020 0635 0646 0641 064B 0627 0020 0648 0627 062D 062F 064B 0627 0020 0639 0644 0649 0020 0627 0644 0623 0642 0644 0020 0642 0628 0644 " & _
            "0020 0627 0644 0637 0628 0627 0639 0629 002E")
    Case 14
        UI = DecodeText("0627 0644 0645 0644 0641 0020 0644 0644 0642 0631 0627 0621 0629 0020 0641 0642 0637 002E 0020 0627 062D 0641 0638 0020 0646 0633 062E 0629 0020 0642 " & _
            "0627 0628 0644 0629 0020 0644 0644 062A 0639 062F 064A 0644 0020 0623 0648 0644 0627 064B 002E")
    Case 20
        UI = DecodeText("062A 0645 0020 062A 0641 0631 064A 063A 0020 062E 0627 0646 0627 062A 0020 0627 0644 0633 0646 062F 0020 0648 062D 0641 0638 0020 0627 0644 0645 0644 " & _
            "0641 0020 0627 0644 0631 0626 064A 0633 064A 0020 0628 0627 0644 0631 0642 0645 0020 0627 0644 062A 0627 0644 064A 002E")
    Case 25
        UI = DecodeText("0627 0646 0642 0631 0020 0631 0627 0628 0637 0020 0050 0044 0046 0020 0641 064A 0020 0627 0644 0623 0631 0634 064A 0641 0020 0644 0641 062A 062D 0020 " & _
            "0627 0644 0633 0646 062F 0020 0648 0625 0639 0627 062F 0629 0020 0637 0628 0627 0639 062A 0647 002E")
    Case 26
        UI = DecodeText("062A 0639 0630 0631 0020 0641 062A 062D 0020 0627 0644 0633 0646 062F 0020 0627 0644 0645 062D 0641 0648 0638 0020 0644 0644 0637 0628 0627 0639 0629 " & _
            "002E")
    Case 29
        UI = DecodeText("062A 0639 0630 0631 0020 062A 062D 062F 064A 062F 0020 0628 062F 0627 064A 0629 0020 0627 0644 0623 0635 0646 0627 0641 0020 0648 0635 0641 0020 0627 " & _
            "0644 0625 062C 0645 0627 0644 064A 0020 0628 0623 0645 0627 0646 002E 0020 062A 0623 0643 062F 0020 0645 0646 0020 0628 0642 0627 0621 0020 0639 0646 " & _
            "0648 0627 0646 0020 0643 0648 062F 0020 0627 0644 0635 0646 0641 0020 0648 0643 0644 0645 0629 0020 0627 0644 0625 062C 0645 0627 0644 064A 061B 0020 " & _
            "0644 0645 0020 062A 064F 0645 0633 062D 0020 0628 064A 0627 0646 0627 062A 0643 002E")
    Case 30
        UI = DecodeText("0643 0648 062F 0020 0627 0644 0635 0646 0641")
    Case 31
        UI = DecodeText("0627 0644 0625 062C 0645 0627 0644 064A")
    Case 38
        UI = DecodeText("0633 062C 0644 0020 0645 0631 062A 062C 0639")
    Case 42
        UI = DecodeText("062D 064F 0641 0638 0020 0645 0644 0641 0020 0050 0044 0046")
    Case 45
        UI = DecodeText("0050 0044 0046 0020 0645 062D 0641 0648 0638 0020 2014 0020 0623 064F 0631 0633 0644 0020 0644 0644 0637 0627 0628 0639 0629")
    Case 46
        UI = DecodeText("0050 0044 0046 0020 0645 062D 0641 0648 0638 0020 2014 0020 062A 0639 0630 0631 062A 0020 0627 0644 0637 0628 0627 0639 0629 0020 0627 0644 0648 0631 " & _
            "0642 064A 0629")
    Case 49
        UI = DecodeText("0627 062D 0641 0638 0020 0627 0644 0645 0635 0646 0641 0020 0641 064A 0020 0645 062C 0644 062F 0020 0645 062D 0644 064A 0020 0623 0648 0644 064B 0627 " & _
            "0020 062D 062A 0649 0020 064A 0645 0643 0646 0020 062D 0641 0638 0020 0645 0644 0641 0627 062A 0020 0050 0044 0046 0020 0628 062C 0627 0646 0628 0647 " & _
            "002E")
    Case 52
        UI = DecodeText("062D 064F 0641 0638 0020 0050 0044 0046 060C 0020 0644 0643 0646 0020 062A 0639 0630 0631 0020 0625 0643 0645 0627 0644 0020 062D 0641 0638 0020 0627 " & _
            "0644 0645 0635 0646 0641 0020 0628 0627 0644 062D 0627 0644 0629 0020 0627 0644 062C 062F 064A 062F 0629 002E 0020 0623 064F 0639 064A 062F 062A 0020 " & _
            "0645 062F 062E 0644 0627 062A 0020 0627 0644 0633 0646 062F 0020 0648 0627 0644 0631 0642 0645 0020 0627 0644 0633 0627 0628 0642 0020 0641 064A 0020 " & _
            "0627 0644 0630 0627 0643 0631 0629 061B 0020 0627 062D 062A 0641 0638 0020 0628 0627 0644 0645 0644 0641 0020 0648 0631 0627 062C 0639 0020 0627 0644 " & _
            "062D 0641 0638 002E")
    Case 53
        UI = DecodeText("0644 0645 0020 064A 064F 0639 062B 0631 0020 0639 0644 0649 0020 0050 0044 0046 0020 0627 0644 0645 0631 062A 0628 0637 0020 0628 0627 0644 0633 0646 " & _
            "062F 002E 0020 064A 062C 0628 0020 0623 0646 0020 064A 0628 0642 0649 0020 0645 0644 0641 0020 0050 0044 0046 0020 0628 062C 0627 0646 0628 0020 0627 " & _
            "0644 0645 0635 0646 0641 0020 0623 0648 0020 062A 0639 0627 062F 0020 0627 0644 0645 0644 0641 0627 062A 0020 0625 0644 0649 0020 0645 0643 0627 0646 " & _
            "0647 0627 002E")
    Case 54
        UI = DecodeText("0644 0645 0020 064A 064F 0639 062B 0631 0020 0639 0644 0649 0020 0633 0646 062F 0020 0635 0627 0644 062D 0020 0623 0648 0020 0631 0627 0628 0637 0020 " & _
            "0050 0044 0046 0020 0644 0647 0630 0627 0020 0627 0644 0633 062C 0644 002E")
    Case 56
        UI = DecodeText("0644 0645 0020 064A 0643 062A 0645 0644 0020 062D 0641 0638 0020 0050 0044 0046 061B 0020 0644 0645 0020 062A 064F 0645 0633 062D 0020 0645 062F 062E " & _
            "0644 0627 062A 0020 0627 0644 0633 0646 062F 0020 0648 0644 0645 0020 064A 064F 0639 062A 0645 062F 0020 0631 0642 0645 0020 062C 062F 064A 062F 002E")
    Case 57
        UI = DecodeText("0623 062F 062E 0644 0020 062A 0627 0631 064A 062E 0020 0627 0646 062A 0647 0627 0621 0020 0635 0644 0627 062D 064A 0629 0020 0635 062D 064A 062D 064B " & _
            "0627 0020 0641 064A 0020 0627 0644 0635 0641 003A")
    Case 58
        UI = DecodeText("0623 062F 062E 0644 0020 062A 0627 0631 064A 062E 0020 0641 0627 062A 0648 0631 0629 0020 0635 062D 064A 062D 064B 0627 0020 0644 0627 0020 064A 062A " & _
            "062C 0627 0648 0632 0020 062A 0627 0631 064A 062E 0020 0627 0644 064A 0648 0645 002E")
    Case 59
        UI = DecodeText("0643 0645 0020 0646 0633 062E 0629 0020 062A 0631 064A 062F 0020 0637 0628 0627 0639 062A 0647 0627 061F 0020 0623 062F 062E 0644 0020 0639 062F 062F " & _
            "064B 0627 0020 0635 062D 064A 062D 064B 0627 0020 0645 0646 0020 0031 0020 0625 0644 0649 0020 0039 0039 0039 002E")
    Case 60
        UI = DecodeText("0639 062F 062F 0020 0646 0633 062E 0020 0627 0644 0637 0628 0627 0639 0629")
    Case 61
        UI = DecodeText("0623 062F 062E 0644 0020 0639 062F 062F 0020 0646 0633 062E 0020 0635 062D 064A 062D 064B 0627 0020 0645 0646 0020 0031 0020 0625 0644 0649 0020 0039 " & _
            "0039 0039 060C 0020 0623 0648 0020 0627 0636 063A 0637 0020 0625 0644 063A 0627 0621 0020 0644 0644 0631 062C 0648 0639 002E")
    End Select
End Function

Private Function DriverSheetName() As String
    DriverSheetName = DecodeText("0642 0627 0639 062F 0629 0020 0627 0644 0633 0627 0626 0642 064A 0646")
End Function

Private Function WarehouseFormula() As String
    Dim s As String, d As String, remain As String, bad As String, quote As String
    quote = Chr$(34)
    s = "'" & Replace(WS(SETTINGS).name, "'", "''") & "'!"
    d = "RC[-1]": remain = "INT(" & d & ")-INT(R6C9)"
    bad = quote & UI(67) & quote
    WarehouseFormula = "=IF(" & d & "=" & quote & quote & "," & quote & quote & ",IFERROR(IF(OR(NOT(ISNUMBER(" & d & "))," & d & "<=0," & d & ">2958465)," & bad & ",IF(" & remain & "<=VALUE(" & s & "R2C1),TRIM(" & s & "R2C2),IF(" & remain & "<=VALUE(" & s & "R3C1),TRIM(" & s & "R3C2),IF(" & remain & "<=VALUE(" & s & "R4C1),TRIM(" & s & "R4C2),TRIM(" & s & "R5C2)))))," & bad & "))"
End Function

Private Function RulesValid() As Boolean
    Dim i As Long, j As Long, value As Variant, previous As Double, label As String
    previous = -1
    For i = 2 To 4
        value = WS(SETTINGS).Cells(i, 1).Value2
        If Not IsNumeric(value) Then Exit Function
        If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) <= previous Then Exit Function
        previous = CDbl(value)
    Next i
    For i = 2 To 5
        label = Trim$(CellText(WS(SETTINGS).Cells(i, 2)))
        If Len(label) = 0 Then Exit Function
        For j = 2 To i - 1
            If StrComp(label, Trim$(CellText(WS(SETTINGS).Cells(j, 2))), vbTextCompare) = 0 Then Exit Function
        Next j
    Next i
    RulesValid = True
End Function

Private Function DriverValid() As Boolean
    DriverValid = (CodeRow(WS(DRIVERS), CellText(WS(MAIN).Range("B6"))) > 0)
End Function

Private Sub RefreshDriver()
    Dim found As Long, formula As String, oldEvents As Boolean
    Dim errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    found = CodeRow(WS(DRIVERS), CellText(WS(MAIN).Range("B6")))
    If found > 0 Then
        If CStr(WS(MAIN).Range("B6").Value2) <> CStr(WS(DRIVERS).Cells(found, 1).Value2) Then
            WS(MAIN).Range("B6").NumberFormat = "@"
            WS(MAIN).Range("B6").Value2 = CStr(WS(DRIVERS).Cells(found, 1).Value2)
        End If
    End If
    formula = DriverFormula()
    If WS(MAIN).Range("L35").formula <> formula Then WS(MAIN).Range("L35").formula = formula
    WS(MAIN).Range("L35").Calculate
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    errorNumber = Err.number
    errorText = Err.Description
    Application.EnableEvents = oldEvents
    Err.Raise errorNumber, "RefreshDriver", errorText
End Sub

Private Function DriverFormula() As String
    Dim name As String, quote As String
    quote = Chr$(34)
    name = "'" & Replace(WS(DRIVERS).name, "'", "''") & "'!$A:$B"
    DriverFormula = "=IF(B6=" & quote & quote & "," & quote & quote & ",IFERROR(VLOOKUP(B6&" & quote & quote & "," & name & ",2,FALSE),IFERROR(VLOOKUP(IFERROR(VALUE(B6),B6)," & name & ",2,FALSE)," & quote & quote & ")))"
End Function

Private Function ReceiptWarehouses(ByVal firstRow As Long, ByVal totalRow As Long) As String
    Dim r As Long, name As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For r = firstRow To totalRow - 1
        If Len(CellText(WS(MAIN).Cells(r, 2))) > 0 Then
            name = CellText(WS(MAIN).Cells(r, 11))
            If Len(name) > 0 And Not seen.Exists(name) Then
                seen.Add name, True
                If Len(ReceiptWarehouses) > 0 Then ReceiptWarehouses = ReceiptWarehouses & " | "
                ReceiptWarehouses = ReceiptWarehouses & name
            End If
        End If
    Next r
End Function

Public Sub ReleaseNotesReturn()
    If NotesReturnBound Then
        Application.OnKey "~"
        Application.OnKey "{ENTER}"
    End If
    NotesReturnBound = False
End Sub

Public Sub EnterFromNotes()
    Dim current As Excel.Range, destination As Excel.Range
    On Error GoTo Failed
    If Not ActiveWorkbook Is ThisWorkbook Then RestoreEntryMode: Exit Sub
    If ActiveSheet.name <> WS(MAIN).name Then RestoreEntryMode: Exit Sub
    If BusyPrinting Or BusySearching Then Exit Sub
    MouseCommit = False
    If Len(PendingEntryAddress) > 0 Then
        FlushEntry
        Exit Sub
    End If
    Set current = ActiveCell.MergeArea.Cells(1, 1)
    If current.address = WS(MAIN).Range("L6").address Then
        ArchiveAndPrint
        Exit Sub
    End If
    Set destination = NextEntryCell(current)
    If Not destination Is Nothing Then QueueEntry destination
    FlushEntry
    Exit Sub
Failed:
    UnicodeMessage UI(6) & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub BeginEntryMode()
    Dim macroName As String
    If Not EntryModeActive Then
        PreviousReturnEnabled = Application.MoveAfterReturn
        PreviousReturnDirection = Application.MoveAfterReturnDirection
        EntryModeActive = True
    End If
    Application.MoveAfterReturn = False
    If Not NotesReturnBound Then
        macroName = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!Module1.EnterFromNotes"
        Application.OnKey "~", macroName
        Application.OnKey "{ENTER}", macroName
        NotesReturnBound = True
    End If
End Sub

Public Sub RestoreEntryMode()
    ReleaseNotesReturn
    RestoreSearchStatus
    If EntryModeActive Then
        Application.MoveAfterReturn = PreviousReturnEnabled
        Application.MoveAfterReturnDirection = PreviousReturnDirection
    End If
    EntryModeActive = False
    EntryChangeActive = False
    PendingEntryAddress = ""
    MouseCommit = False
End Sub

Public Sub ConfigureEntryMode()
    If BusyPrinting Or BusySearching Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Then
        RestoreEntryMode
        Exit Sub
    End If
    If ActiveSheet.name = WS(MAIN).name Then BeginEntryMode Else RestoreEntryMode
End Sub

Private Sub QueueEntry(ByVal destination As Excel.Range)
    If MouseCommit Then Exit Sub
    PendingEntryAddress = destination.MergeArea.Cells(1, 1).address
End Sub

Private Sub FlushEntry()
    Dim oldEvents As Boolean, address As String
    If Len(PendingEntryAddress) = 0 Then Exit Sub
    address = PendingEntryAddress
    PendingEntryAddress = ""
    If MouseCommit Or GetAsyncKeyState(1) < 0 Then Exit Sub
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    If ActiveSheet.name <> WS(MAIN).name Then Exit Sub
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    Application.EnableEvents = False
    WS(MAIN).Range(address).Select
Done:
    Application.EnableEvents = oldEvents
End Sub

Private Function NextEntryCell(ByVal current As Excel.Range) As Excel.Range
    Dim firstRow As Long, totalRow As Long, row As Long, col As Long
    FormRows firstRow, totalRow
    row = current.row: col = current.column
    If row = 6 Then
        Select Case col
        Case 2: Set NextEntryCell = WS(MAIN).Range("D6")
        Case 4: Set NextEntryCell = WS(MAIN).Range("F6")
        Case 6: Set NextEntryCell = WS(MAIN).Range("A8")
        End Select
    ElseIf row = 8 And (col = 1 Or col = 2) Then
        Set NextEntryCell = WS(MAIN).Cells(firstRow, 3)
    ElseIf row >= firstRow And row < totalRow Then
        Select Case col
        Case 2, 3: Set NextEntryCell = WS(MAIN).Cells(row, 7)
        Case 7, 8, 9: Set NextEntryCell = WS(MAIN).Cells(row, col + 1)
        Case 10, 11: Set NextEntryCell = WS(MAIN).Cells(row, 12)
        Case 12: Set NextEntryCell = WS(MAIN).Cells(row, 13)
        Case 13
            If row < totalRow - 1 Then
                Set NextEntryCell = WS(MAIN).Cells(row + 1, 3)
            Else
                Set NextEntryCell = WS(MAIN).Range("L6")
            End If
        End Select
    End If
End Function

Public Sub HandleSelection(ByVal Sh As Object, ByVal target As Excel.Range)
    If Sh.name <> WS(MAIN).name Then Exit Sub
    If BusyPrinting Or BusySearching Then Exit Sub
    PendingEntryAddress = ""
End Sub


Private Function LastRow(ByVal Sh As Excel.Worksheet, ByVal cols As Long) As Long
    Dim c As Long, r As Long
    LastRow = 1
    For c = 1 To cols
        r = Sh.Cells(Sh.rows.count, c).End(xlUp).row
        If r > LastRow Then LastRow = r
    Next c
End Function

Private Function Normal(ByVal s As String) As String
    Dim i As Long
    s = LCase$(Trim$(s))
    For i = 0 To 9
        s = Replace(s, ChrW(&H660 + i), Chr$(48 + i))
        s = Replace(s, ChrW(&H6F0 + i), Chr$(48 + i))
    Next i
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

Private Sub InvalidateSearchCaches()
    ClientCacheReady = False
    ProductCacheReady = False
    ClientListReady = False
    ProductListReady = False
End Sub

Private Sub EnsureSearchCache(ByVal Customer As Boolean)
    Dim src As Excel.Worksheet, data As Variant, Prepared() As String, index As Object
    Dim last As Long, i As Long, key As String, Ready As Boolean, cachedLast As Long
    If Customer Then
        Ready = ClientCacheReady: cachedLast = ClientCacheLast
    Else
        Ready = ProductCacheReady: cachedLast = ProductCacheLast
    End If
    Set src = CandidateSource(Customer)
    last = LastRow(src, 2)
    If Ready And cachedLast = last Then Exit Sub
    If Customer Then ClientListReady = False Else ProductListReady = False
    Set index = CreateObject("Scripting.Dictionary")
    index.CompareMode = vbBinaryCompare
    If last >= 2 Then
        data = src.Range("A2:B" & last).Value2
        ReDim Prepared(1 To UBound(data, 1))
        For i = 1 To UBound(data, 1)
            If IsError(data(i, 1)) Or IsError(data(i, 2)) Then
                data(i, 1) = "": data(i, 2) = ""
            End If
            If Len(Trim$(CStr(data(i, 1)))) > 0 And Len(Trim$(CStr(data(i, 2)))) > 0 Then
                key = Normal(CStr(data(i, 1)))
                Prepared(i) = Normal(CStr(data(i, 2)) & " " & CStr(data(i, 1)))
                If index.Exists(key) Then index(key) = -1 Else index.Add key, i + 1
            End If
        Next i
    Else
        ReDim Prepared(1 To 1)
    End If
    If Customer Then
        ClientCache = data: ClientCacheKeys = Prepared
        Set ClientCodeIndex = index
        ClientCacheLast = last: ClientCacheReady = True
    Else
        ProductCache = data: ProductCacheKeys = Prepared
        Set ProductCodeIndex = index
        ProductCacheLast = last: ProductCacheReady = True
    End If
End Sub

Private Function FilterSearchData(ByVal Customer As Boolean, ByVal query As String, ByRef data As Variant, ByRef Prepared() As String) As Long
    Dim i As Long, j As Long, count As Long, oldLast As Long, column As Long, rows As Long
    Dim matches() As Long, result() As Variant, tokens As Variant, token As Variant, accepted As Boolean
    Dim listName As String, reference As String, existing As Excel.name, suffix As Boolean
    query = Normal(query)
    suffix = Customer And (Len(query) = 4 Or Len(query) = 5) And Not (query Like "*[!0-9]*")
    If Customer Then column = 27: listName = "InlineClients" Else column = 28: listName = "InlineProducts"
    tokens = Split(Normal(query), " ")
    If IsArray(data) Then
        rows = UBound(data, 1)
        ReDim matches(1 To rows)
        For i = 1 To rows
            If Len(Prepared(i)) > 0 Then
                accepted = True
                For Each token In tokens
                    If Len(token) > 0 Then
                        If InStr(1, Prepared(i), CStr(token), vbBinaryCompare) = 0 Then accepted = False: Exit For
                    End If
                Next token
                If suffix Then accepted = (Right$(Normal(CStr(data(i, 1))), Len(query)) = query)
                If accepted Then count = count + 1: matches(count) = i
            End If
        Next i
    End If
    With WS(MAIN)
        oldLast = .Cells(.rows.count, column).End(xlUp).row
        If oldLast >= 2 Then .Range(.Cells(2, column), .Cells(oldLast, column)).ClearContents
        If count > 0 Then
            ReDim result(1 To count, 1 To 1)
            For j = 1 To count
                i = matches(j)
                result(j, 1) = CStr(data(i, 1)) & " | " & Left$(CStr(data(i, 2)), 220)
            Next j
            .Cells(2, column).Resize(count, 1).Value2 = result
        Else
            .Cells(2, column).Value2 = UI(5)
        End If
        reference = "='" & Replace(.name, "'", "''") & "'!" & .Cells(2, column).Resize(IIf(count > 0, count, 1), 1).address
        .Range("AA:AB").EntireColumn.Hidden = True
    End With
    On Error Resume Next
    Set existing = ThisWorkbook.Names(listName)
    On Error GoTo 0
    If existing Is Nothing Then
        ThisWorkbook.Names.Add name:=listName, RefersTo:=reference
    Else
        existing.RefersTo = reference
    End If
    FilterSearchData = count
End Function

Private Sub SearchStatus(ByVal text As String)
    If Not SearchStatusOwned Then
        OriginalSearchStatus = Application.StatusBar
        SearchStatusOwned = True
    End If
    Application.StatusBar = text
End Sub

Public Sub RestoreSearchStatus()
    If SearchStatusOwned Then
        Application.StatusBar = OriginalSearchStatus
        SearchStatusOwned = False
    End If
End Sub

Private Function CodeRow(ByVal src As Excel.Worksheet, ByVal code As String) As Long
    Dim key As String, index As Object, r As Long, found As Long
    key = Normal(code)
    If Len(key) = 0 Then Exit Function
    If src.name = WS(CLIENTS).name Then
        EnsureSearchCache True
        Set index = ClientCodeIndex
    ElseIf src.name = WS(PRODUCTS).name Then
        EnsureSearchCache False
        Set index = ProductCodeIndex
    End If
    If Not index Is Nothing Then
        If index.Exists(key) Then CodeRow = CLng(index(key))
        Exit Function
    End If
    For r = 2 To LastRow(src, 2)
        If Not IsError(src.Cells(r, 1).Value2) And Not IsError(src.Cells(r, 2).Value2) Then
            If Normal(CStr(src.Cells(r, 1).Value2)) = key And Len(Trim$(CStr(src.Cells(r, 2).Value2))) > 0 Then
                If found > 0 Then CodeRow = -1: Exit Function
                found = r
            End If
        End If
    Next r
    CodeRow = found
End Function

Private Function CellText(ByVal cell As Excel.Range) As String
    If IsError(cell.Value2) Then
        CellText = "#ERROR"
    Else
        CellText = Trim$(Replace(Replace(CStr(cell.Value2), ChrW(160), " "), vbTab, " "))
    End If
End Function

Private Sub FormRows(ByRef firstRow As Long, ByRef totalRow As Long)
    Dim r As Long, lastLabelRow As Long, title As String, header As Excel.Range
    firstRow = 0: totalRow = 0
    lastLabelRow = WS(MAIN).Cells(WS(MAIN).rows.count, 1).End(xlUp).row
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
            firstRow = header.row + header.rows.count
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

Private Function ItemCells(ByVal firstCol As Long, ByVal lastCol As Long) As Excel.Range
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

Private Function HasList(ByVal cell As Excel.Range) As Boolean
    On Error Resume Next
    HasList = (cell.Validation.Type = xlValidateList)
    On Error GoTo 0
End Function

Private Sub ColorStatuses(ByVal firstRow As Long, ByVal totalRow As Long)
    Dim target As Excel.Range, rule As Excel.FormatCondition, i As Long, colors As Variant
    Set target = WS(MAIN).Range(WS(MAIN).Cells(firstRow, 11), WS(MAIN).Cells(totalRow - 1, 11))
    colors = Array(RGB(255, 199, 206), RGB(248, 203, 173), RGB(255, 235, 156), RGB(198, 239, 206))
    Dim unchanged As Boolean
    unchanged = (target.FormatConditions.count = 4)
    If unchanged Then
        For i = 1 To 4
            Set rule = target.FormatConditions(i)
            If rule.Type <> xlCellValue Then
                unchanged = False
                Exit For
            End If
            If rule.AppliesTo.address <> target.address Then unchanged = False
            If rule.Operator <> xlEqual Then unchanged = False
            If rule.Formula1 <> "=" & Chr$(34) & UI(61 + i) & Chr$(34) Then unchanged = False
        Next i
        If unchanged Then Exit Sub
    End If
    target.FormatConditions.Delete
    For i = 0 To 3
        Set rule = target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=" & Chr$(34) & UI(62 + i) & Chr$(34))
        rule.Interior.Color = colors(i)
        rule.Font.Color = RGB(0, 0, 0)
    Next i
End Sub

Private Sub EnsureItemRows(ByVal firstRow As Long, ByVal totalRow As Long)
    Dim r As Long, template As String
    template = UI(37)
    For r = firstRow To totalRow - 1
        If WS(MAIN).Cells(r, 1).Value2 <> r - firstRow + 1 Then WS(MAIN).Cells(r, 1).Value2 = r - firstRow + 1
        If WS(MAIN).Cells(r, 11).FormulaR1C1 <> template Then WS(MAIN).Cells(r, 11).FormulaR1C1 = template
        If Not HasList(WS(MAIN).Cells(r, 2)) Then SetList WS(MAIN).Cells(r, 2), "InlineProducts"
        If Not HasList(WS(MAIN).Cells(r, 3)) Then SetList WS(MAIN).Cells(r, 3), "InlineProducts"
        If Not HasList(WS(MAIN).Cells(r, 12)) Then
        With WS(MAIN).Cells(r, 12).Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="=ReturnReasons"
            .IgnoreBlank = True
            .ShowError = True
        End With
        End If
    Next r
    ColorStatuses firstRow, totalRow
End Sub

Private Sub FitReceiptPage(ByVal receipt As Excel.Worksheet)
    Dim previousCommunication As Boolean, errorNumber As Long, errorText As String
    previousCommunication = Application.PrintCommunication
    On Error GoTo Failed
    Application.PrintCommunication = False
    With receipt.PageSetup
        .PrintArea = "$A$1:$M$36"
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    Application.PrintCommunication = previousCommunication
    Exit Sub
Failed:
    errorNumber = Err.number
    errorText = Err.Description
    On Error Resume Next
    Application.PrintCommunication = previousCommunication
    On Error GoTo 0
    Err.Raise errorNumber, "FitReceiptPage", errorText
End Sub

Private Sub RefreshTotals(ByVal firstRow As Long, ByVal totalRow As Long)
    Dim c As Long, oldEvents As Boolean, errorNumber As Long, errorText As String
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    If Not RulesValid() Then Err.Raise vbObjectError + 941, "RefreshTotals", UI(66)
    WS(MAIN).Range("I6").formula = "=NOW()"
    EnsureItemRows firstRow, totalRow
    For c = 7 To 9
        WS(MAIN).Cells(totalRow, c).formula = "=SUM(" & WS(MAIN).Range(WS(MAIN).Cells(firstRow, c), WS(MAIN).Cells(totalRow - 1, c)).address(False, False) & ")"
    Next c
    WS(MAIN).Range("J8").formula = "=SUMPRODUCT(--(" & WS(MAIN).Range(WS(MAIN).Cells(firstRow, 2), WS(MAIN).Cells(totalRow - 1, 2)).address(False, False) & "<>""""))"
    WS(MAIN).Calculate
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    errorNumber = Err.number
    errorText = Err.Description
    Application.EnableEvents = oldEvents
    Err.Raise errorNumber, "RefreshTotals", errorText
End Sub

Private Function SearchInputCells() As Excel.Range
    Set SearchInputCells = Application.Union(WS(MAIN).Range("A8,B8"), ItemCells(2, 3))
End Function

Private Function ProductSearchTarget() As Excel.Range
    Dim r As Long
    For r = FirstItemRow() To LastItemRow()
        If Len(CellText(WS(MAIN).Cells(r, 2))) = 0 And Len(CellText(WS(MAIN).Cells(r, 3))) = 0 Then
            Set ProductSearchTarget = WS(MAIN).Cells(r, 3)
            Exit Function
        End If
    Next r
    Set ProductSearchTarget = WS(MAIN).Cells(FirstItemRow(), 3)
End Function

Private Function IsClientCell(ByVal dest As Excel.Range) As Boolean
    IsClientCell = (dest.row = 8 And (dest.column = 1 Or dest.column = 2))
End Function

Private Function CandidateSource(ByVal Customer As Boolean) As Excel.Worksheet
    If Customer Then Set CandidateSource = WS(CLIENTS) Else Set CandidateSource = WS(PRODUCTS)
End Function

Private Sub SetList(ByVal dest As Excel.Range, ByVal listName As String)
    Dim protected As Boolean
    protected = dest.Worksheet.ProtectContents
    If protected Then dest.Worksheet.Unprotect
    With dest.MergeArea.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Operator:=xlBetween, Formula1:="=" & listName
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
        .ShowInput = False
        .InputTitle = ""
        .InputMessage = ""
    End With
    If protected Then dest.Worksheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Private Function BuildList(ByVal Customer As Boolean, ByVal query As String) As Long
    Dim normalized As String
    EnsureSearchCache Customer
    normalized = Normal(query)
    If Customer Then
        If ClientListReady And ClientListQuery = normalized Then BuildList = ClientListCount: Exit Function
        ClientListCount = FilterSearchData(True, query, ClientCache, ClientCacheKeys)
        ClientListQuery = normalized: ClientListReady = True
        BuildList = ClientListCount
    Else
        If ProductListReady And ProductListQuery = normalized Then BuildList = ProductListCount: Exit Function
        ProductListCount = FilterSearchData(False, query, ProductCache, ProductCacheKeys)
        ProductListQuery = normalized: ProductListReady = True
        BuildList = ProductListCount
    End If
End Function

Private Sub ShowList(ByVal dest As Excel.Range)
    Dim dialog As SearchWindow, sourceRow As Long, Customer As Boolean
    Dim oldBusy As Boolean, oldEvents As Boolean, oldScreen As Boolean, errorText As String
    oldScreen = Application.ScreenUpdating
    oldBusy = BusySearching
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    BusySearching = True
    Application.EnableEvents = False
    Customer = IsClientCell(dest)
    Set dialog = New SearchWindow
    dialog.Configure Customer, CellText(dest)
    Application.ScreenUpdating = True
    dialog.show vbModal
    sourceRow = dialog.ChosenRow
    dialog.ReleaseRows
    Unload dialog
    Set dialog = Nothing
    If sourceRow > 0 Then ApplyChoice dest, sourceRow
Done:
    On Error Resume Next
    If Not dialog Is Nothing Then
        dialog.ReleaseRows
        Unload dialog
    End If
    Set dialog = Nothing
    Application.EnableEvents = oldEvents
    BusySearching = oldBusy
    Application.ScreenUpdating = oldScreen
    If sourceRow > 0 And Not EntryChangeActive Then FlushEntry
    Exit Sub
Failed:
    errorText = Err.Description
    UnicodeMessage UI(6) & vbCrLf & errorText, vbExclamation
    Resume Done
End Sub

Private Sub FillNames()
    Dim r As Long, found As Long
    found = CodeRow(WS(CLIENTS), CellText(WS(MAIN).Range("A8")))
    If found > 0 Then WS(MAIN).Range("B8").Value2 = WS(CLIENTS).Cells(found, 2).Value2
    For r = FirstItemRow() To LastItemRow()
        found = CodeRow(WS(PRODUCTS), CellText(WS(MAIN).Cells(r, 2)))
        If found > 0 Then WS(MAIN).Cells(r, 3).Value2 = WS(PRODUCTS).Cells(found, 2).Value2
    Next r
End Sub

Private Sub ApplyChoice(ByVal dest As Excel.Range, ByVal sourceRow As Long)
    Dim Customer As Boolean, src As Excel.Worksheet
    Customer = IsClientCell(dest)
    Set src = CandidateSource(Customer)
    RestoreSearchStatus
    If Customer Then
        WS(MAIN).Range("A8").NumberFormat = "@"
        WS(MAIN).Range("A8").Value2 = CStr(src.Cells(sourceRow, 1).Value2)
        WS(MAIN).Range("B8").Value2 = src.Cells(sourceRow, 2).Value2
    Else
        WS(MAIN).Cells(dest.row, 2).NumberFormat = "@"
        WS(MAIN).Cells(dest.row, 2).Value2 = CStr(src.Cells(sourceRow, 1).Value2)
        WS(MAIN).Cells(dest.row, 3).Value2 = src.Cells(sourceRow, 2).Value2

    End If
    BuildList Customer, ""
    WS(MAIN).Calculate

    If Customer Then
        QueueEntry WS(MAIN).Cells(FirstItemRow(), 3)
    Else
        QueueEntry WS(MAIN).Cells(dest.row, 7)
    End If
End Sub

Private Sub SearchChanged(ByVal dest As Excel.Range)
    Dim Customer As Boolean, src As Excel.Worksheet, query As String, code As String
    Dim p As Long, found As Long, count As Long, c As Long, listName As String
    Customer = IsClientCell(dest)
    Set src = CandidateSource(Customer)
    query = CellText(dest)
    If Customer Then c = 27: listName = "InlineClients" Else c = 28: listName = "InlineProducts"
    If Customer Then
        If dest.column = 2 Then WS(MAIN).Range("A8").MergeArea.ClearContents Else WS(MAIN).Range("B8").MergeArea.ClearContents
    Else
        If dest.column = 3 Then WS(MAIN).Cells(dest.row, 2).ClearContents Else WS(MAIN).Cells(dest.row, 3).MergeArea.ClearContents
    End If
    If Len(query) = 0 Then
        RestoreSearchStatus
        BuildList Customer, ""
        SetList dest, listName
        Exit Sub
    End If
    code = query
    p = InStrRev(query, " | ")
    If p > 0 Then code = Left$(query, p - 1)
    found = CodeRow(src, code)
    If found = -1 Then UnicodeMessage UI(8), vbExclamation: Exit Sub
    If found > 0 Then
        ApplyChoice dest, found
        Exit Sub
    End If
    count = BuildList(Customer, query)
    SetList dest, listName
    If count = 0 Then
        SearchStatus UI(5)
        QueueEntry dest
    ElseIf count = 1 Then
        code = CStr(WS(MAIN).Cells(2, c).Value2)
        code = Left$(code, InStrRev(code, " | ") - 1)
        found = CodeRow(src, code)
        If found > 0 Then ApplyChoice dest, found Else UnicodeMessage UI(8), vbExclamation
    Else
        SearchStatus UI(75) & CStr(count)
    End If
End Sub

Private Sub StartSearch(ByVal dest As Excel.Range)
    MouseCommit = False
    ShowList dest
End Sub

Public Sub SearchProducts()
    If ActiveSheet.name <> WS(MAIN).name Then Exit Sub
    If ActiveCell.row >= FirstItemRow() And ActiveCell.row <= LastItemRow() Then
        StartSearch WS(MAIN).Cells(ActiveCell.row, 3)
    Else
        StartSearch ProductSearchTarget()
    End If
End Sub

Public Sub SearchClients()
    If ActiveSheet.name = WS(MAIN).name Then StartSearch WS(MAIN).Range("B8")
End Sub

Public Sub HandleChange(ByVal Sh As Object, ByVal target As Excel.Range)
    Dim oldScreen As Boolean
    Dim oldEvents As Boolean, dest As Excel.Range, firstRow As Long, totalRow As Long
    Dim changed As Excel.Range, cell As Excel.Range, found As Long, Customer As Boolean, src As Excel.Worksheet
    Dim query As String, splitAt As Long
    If BusyPrinting Then Exit Sub
    PendingEntryAddress = ""
    MouseCommit = (GetAsyncKeyState(1) < 0)
    EntryChangeActive = True
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If Sh.name = WS(MAIN).name Then
        BeginEntryMode
        If Not Intersect(target, Sh.Range("B6")) Is Nothing Then
            RefreshDriver
            If Len(CellText(Sh.Range("B6"))) > 0 Then
                If DriverValid() Then QueueEntry Sh.Range("D6") Else UnicodeMessage UI(68), vbExclamation
            End If
            GoTo Done
        ElseIf Not Intersect(target, Sh.Range("D6")) Is Nothing Then
            If Len(CellText(Sh.Range("D6"))) > 0 Then QueueEntry Sh.Range("F6")
            GoTo Done
        ElseIf Not Intersect(target, Sh.Range("F6")) Is Nothing Then
            If Len(CellText(Sh.Range("F6"))) > 0 Then
                If IsNumeric(Sh.Range("F6").Value2) Then
                    If CDbl(Sh.Range("F6").Value2) > 0 And CDbl(Sh.Range("F6").Value2) <= CDbl(Date) Then
                        QueueEntry Sh.Range("A8")
                    Else
                        UnicodeMessage UI(58), vbExclamation
                    End If
                Else
                    UnicodeMessage UI(58), vbExclamation
                End If
            End If
            GoTo Done
        End If
        FormRows firstRow, totalRow
        Set dest = target.Cells(1, 1).MergeArea.Cells(1, 1)
        Set changed = Intersect(target, SearchInputCells())
        If Not changed Is Nothing Then
            If target.CountLarge = 1 Or target.address = dest.MergeArea.address Then
                SearchChanged dest
            Else
                For Each cell In changed.Cells
                    If cell.address = cell.MergeArea.Cells(1, 1).address Then
                        Customer = IsClientCell(cell)
                        Set src = CandidateSource(Customer)
                        query = CellText(cell)
                        splitAt = InStrRev(query, " | ")
                        If splitAt > 0 Then query = Left$(query, splitAt - 1)
                        If Len(query) = 0 Then
                            If Customer Then
                                If cell.column = 1 Then Sh.Range("B8").MergeArea.ClearContents Else Sh.Range("A8").ClearContents
                            Else
                                If cell.column = 2 Then Sh.Cells(cell.row, 3).MergeArea.ClearContents Else Sh.Cells(cell.row, 2).ClearContents
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
        If target.CountLarge = 1 And dest.row >= firstRow And dest.row < totalRow Then
            If dest.column >= 7 And dest.column <= 9 Then
                If Len(CellText(dest)) > 0 Then
                    If Not IsNumeric(dest.Value2) Then UnicodeMessage UI(12) & " " & dest.row, vbExclamation: QueueEntry dest: GoTo Done
                    If CDbl(dest.Value2) < 0 Or CDbl(dest.Value2) <> Fix(CDbl(dest.Value2)) Then UnicodeMessage UI(12) & " " & dest.row, vbExclamation: QueueEntry dest: GoTo Done
                End If
                QueueEntry Sh.Cells(dest.row, dest.column + 1)
            ElseIf dest.column = 10 Then
                If Len(CellText(dest)) > 0 Then
                    If Not IsNumeric(dest.Value2) Then UnicodeMessage UI(57) & " " & dest.row, vbExclamation: QueueEntry dest: GoTo Done
                    If CDbl(dest.Value2) <= 0 Or CDbl(dest.Value2) > 2958465 Then UnicodeMessage UI(57) & " " & dest.row, vbExclamation: QueueEntry dest: GoTo Done
                End If
                QueueEntry Sh.Cells(dest.row, 12)
            ElseIf dest.column = 12 Then
                QueueEntry Sh.Cells(dest.row, 13)
            ElseIf dest.column = 13 Then
                If dest.row < totalRow - 1 Then
                    QueueEntry Sh.Cells(dest.row + 1, 3)
                Else
                    QueueEntry Sh.Range("L6")
                End If
            End If
        End If
    ElseIf Sh.name = WS(DRIVERS).name Then
        RefreshDriver
    ElseIf Sh.name = WS(SETTINGS).name Then
        If Not Intersect(target, Sh.Range("A2:B5")) Is Nothing Then
            If Not RulesValid() Then UnicodeMessage UI(66), vbExclamation: GoTo Done
            WS(MAIN).Range("K11:K32").FormulaR1C1 = WarehouseFormula()
            FormRows firstRow, totalRow
            ColorStatuses firstRow, totalRow
            WS(MAIN).Calculate
        End If
    ElseIf Sh.name = WS(CLIENTS).name Or Sh.name = WS(PRODUCTS).name Then
        InvalidateSearchCaches
        FillNames
        BuildList True, ""
        BuildList False, ""
        WS(MAIN).Calculate
    End If
Done:
    If Not MouseCommit Then FlushEntry
    EntryChangeActive = False
    MouseCommit = False
    Application.ScreenUpdating = oldScreen
    Application.EnableEvents = oldEvents
    Exit Sub
Failed:
    Application.ScreenUpdating = oldScreen
    UnicodeMessage UI(6) & vbCrLf & Err.Description, vbExclamation
    Resume Done
End Sub

Public Sub HandleDoubleClick(ByVal Sh As Object, ByVal target As Excel.Range, ByRef Cancel As Boolean)
    Dim clicked As Excel.Range
    If Cancel Or BusyPrinting Or BusySearching Then Exit Sub
    On Error GoTo Failed
    Set clicked = target.Cells(1, 1).MergeArea.Cells(1, 1)
    If Sh.CodeName = WS(MAIN).CodeName Then
        If Not Intersect(target, Sh.Range("L6").MergeArea) Is Nothing Then
            Cancel = True
            ArchiveAndPrint
        ElseIf Not Intersect(clicked, Sh.Range("A8:D8")) Is Nothing Then
            Cancel = True
            PendingEntryAddress = ""
            MouseCommit = False
            StartSearch Sh.Range("B8")
        ElseIf Not Intersect(clicked.MergeArea, ItemCells(2, 3)) Is Nothing Then
            Cancel = True
            PendingEntryAddress = ""
            MouseCommit = False
            StartSearch Sh.Cells(clicked.row, 3)
        End If
    ElseIf IsSavedReceiptSheet(Sh.name) Then
        If Not Intersect(target, Sh.Range("J11:K11")) Is Nothing Then
            Cancel = True
            ReprintSavedSheet Sh.name
        End If
    ElseIf Sh.name = "Archive" Then
        If target.row >= 5 And target.column = 9 Then
            Cancel = True
            ReprintArchive
        End If
    End If
    Exit Sub
Failed:
    UnicodeMessage UI(6) & vbCrLf & Err.Description, vbExclamation
End Sub

Public Sub ProtectFormulaCells()
    Dim sheet As Excel.Worksheet, receipt As Excel.Worksheet
    Set receipt = WS(MAIN)
    For Each sheet In ThisWorkbook.Worksheets
        If Not IsSavedReceiptSheet(sheet.name) Then
            If sheet.CodeName = receipt.CodeName Or sheet.ProtectContents Then
                sheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True, AllowFiltering:=True
                sheet.EnableSelection = xlNoRestrictions
            End If
        End If
    Next sheet
End Sub

Private Sub BindSearchButtons()
    Dim prefix As String, button As Object
    prefix = "'" & Replace(ThisWorkbook.name, "'", "''") & "'!Module1."
    On Error Resume Next
    Set button = WS(MAIN).Shapes("btn_SearchClient")
    On Error GoTo 0
    If Not button Is Nothing Then button.OnAction = prefix & "SearchClients"
    Set button = Nothing
    On Error Resume Next
    Set button = WS(MAIN).Shapes("btn_SearchProduct")
    On Error GoTo 0
    If Not button Is Nothing Then button.OnAction = prefix & "SearchProducts"
End Sub

Public Sub Initialize()
    Dim oldEvents As Boolean, oldScreen As Boolean, errorText As String
    Dim receipt As Excel.Worksheet
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Set receipt = WS(MAIN)
    receipt.Select Replace:=True
    receipt.Range("B6").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    receipt.Unprotect
    BindSearchButtons
    InvalidateSearchCaches
    RestoreSearchStatus
    ProtectFormulaCells
    BeginEntryMode
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    RestoreEntryMode
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    On Error GoTo 0
    UnicodeMessage UI(6) & vbCrLf & errorText, vbExclamation
End Sub

Private Function NextNumber() As Long
    Dim Sh As Excel.Worksheet, n As Long, suffix As String, log As Excel.Worksheet, row As Long
    n = 1
    If IsNumeric(WS(SETTINGS).Range("E2").value) Then
        If CDbl(WS(SETTINGS).Range("E2").value) >= 1 Then n = CLng(WS(SETTINGS).Range("E2").value)
    End If
    For Each Sh In ThisWorkbook.Worksheets
        If Left$(Sh.name, 4) = "RET_" Then
            suffix = Mid$(Sh.name, 5)
            If IsNumeric(suffix) Then
                If CLng(suffix) >= n Then n = CLng(suffix) + 1
            End If
        End If
    Next Sh
    Set log = ThisWorkbook.Worksheets("Archive")
    For row = 5 To LastRow(log, 15)
        If IsNumeric(log.Cells(row, 1).Value2) Then
            If CDbl(log.Cells(row, 1).Value2) >= n Then n = CLng(log.Cells(row, 1).Value2) + 1
        End If
    Next row
    NextNumber = n
End Function

Private Function ValidForm() As Boolean
    Dim r As Long, c As Long, firstRow As Long, totalRow As Long
    Dim quantity As Double, entries As Long, code As String, value As String
    If Not RulesValid() Then UnicodeMessage UI(66), vbExclamation: Exit Function
    If Not DriverValid() Then UnicodeMessage UI(68), vbExclamation: Exit Function
    If Len(CellText(WS(MAIN).Range("D6"))) = 0 Then UnicodeMessage UI(69), vbExclamation: Exit Function
    If Len(CellText(WS(MAIN).Range("F6"))) = 0 Then UnicodeMessage UI(70), vbExclamation: Exit Function
    RefreshDriver
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
            If quantity = 0 And StrComp(Normal(code), "m500", vbBinaryCompare) <> 0 Then UnicodeMessage UI(12) & " " & r, vbExclamation: Exit Function
        End If
    Next r
    If entries = 0 Then UnicodeMessage UI(13), vbExclamation: Exit Function
    ValidForm = True
End Function

Private Function ReceiptFileName(ByVal number As Long, Optional ByVal customerName As String = "") As String
    Dim baseName As String, bad As Variant, i As Long
    baseName = Trim$(customerName)
    If Len(baseName) = 0 Then baseName = CellText(WS(MAIN).Range("B8"))
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

Private Sub CloseReceiptCopy(ByRef temporary As Excel.Workbook)
    On Error Resume Next
    Application.CutCopyMode = False
    If Not temporary Is Nothing Then temporary.Close SaveChanges:=False
    Set temporary = Nothing
    On Error GoTo 0
End Sub

Private Sub PrepareReceiptCopy(ByVal number As Long, ByVal totalRow As Long, ByRef temporary As Excel.Workbook, ByRef receipt As Excel.Worksheet)
    Dim address As String, i As Long, lookupLast As Long, usedLast As Long
    address = "A1:M36"
    WS(MAIN).Calculate
    WS(MAIN).Copy
    Set temporary = Application.ActiveWorkbook
    Set receipt = temporary.Worksheets(1)
    receipt.Unprotect
    receipt.Range(address).Copy
    receipt.Range(address).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    receipt.Range("I6").value = Now
    receipt.Range("L8").value = "#" & CStr(number)
    receipt.Range("L6").MergeArea.ClearContents
    lookupLast = Application.Max(receipt.Cells(receipt.rows.count, 27).End(xlUp).row, receipt.Cells(receipt.rows.count, 28).End(xlUp).row)
    If lookupLast >= 1 Then receipt.Range("AA1:AB" & lookupLast).ClearContents
    usedLast = receipt.UsedRange.row + receipt.UsedRange.rows.count - 1
    If usedLast > 36 Then
        If Application.CountA(receipt.Range("A37:M" & usedLast)) = 0 Then
            receipt.rows("37:" & usedLast).Delete
        End If
    End If
    receipt.Range(address).Validation.Delete
    For i = receipt.Shapes.count To 1 Step -1
        If Left$(receipt.Shapes(i).name, 4) = "btn_" Or receipt.Shapes(i).Type = msoFormControl Then receipt.Shapes(i).Delete
    Next i
    FitReceiptPage receipt
End Sub

Private Sub ExportReceiptPDF(ByVal receipt As Excel.Worksheet, ByVal path As String)
    receipt.ExportAsFixedFormat Type:=xlTypePDF, fileName:=path, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    If Not FileExists(path) Then Err.Raise vbObjectError + 921, "ExportReceiptPDF", UI(56)
    If CreateObject("Scripting.FileSystemObject").GetFile(path).size = 0 Then Err.Raise vbObjectError + 921, "ExportReceiptPDF", UI(56)
End Sub

Private Sub RestoreCells(ByVal target As Excel.Range, ByVal saved As Variant)
    Dim cell As Excel.Range
    For Each cell In target.Cells
        If cell.address = cell.MergeArea.Cells(1, 1).address Then
            cell.formula = saved(cell.row - target.row + 1, cell.column - target.column + 1)
        End If
    Next cell
End Sub

Public Sub ArchiveAndPrint()
    Dim form As Excel.Worksheet, log As Excel.Worksheet, receipt As Excel.Worksheet, temporary As Excel.Workbook
    Dim firstRow As Long, totalRow As Long, number As Long, logRow As Long
    Dim pdfPath As String, relativeName As String, paper As Boolean, errorText As String, paperError As String, statusError As String
    Dim oldEvents As Boolean, oldScreen As Boolean, oldCounter As Variant
    Dim oldCalculation As Excel.XlCalculation
    Dim itemBackup As Variant, noteBackup As Variant, clientCode As Variant, clientName As Variant
    Dim copiesRequested As Variant, copyCount As Long, returnBackup As Variant, invoiceBackup As Variant, driverBackup As Variant
    Dim backupReady As Boolean, resetStarted As Boolean, pdfSaved As Boolean, committed As Boolean
    If BusyPrinting Then Exit Sub
    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating
    oldCalculation = Application.Calculation
    On Error GoTo Failed
    Set form = WS(MAIN)
    Set log = ThisWorkbook.Worksheets("Archive")
    If ThisWorkbook.ReadOnly Then UnicodeMessage UI(14), vbExclamation: Exit Sub
    pdfPath = FolderPath()
    FormRows firstRow, totalRow
    RefreshTotals firstRow, totalRow
    If Not ValidForm() Then Exit Sub
    Do
        copiesRequested = Application.InputBox(prompt:=UI(59), title:=UI(60), Default:=1, Type:=1)
        If VarType(copiesRequested) = vbBoolean Then Exit Sub
        If copiesRequested >= 1 And copiesRequested <= 999 Then
            If copiesRequested = Fix(copiesRequested) Then Exit Do
        End If
        UnicodeMessage UI(61), vbExclamation
    Loop
    copyCount = CLng(copiesRequested)
    If Not Application.Dialogs(xlDialogPrinterSetup).show Then Exit Sub
    paper = True
    number = NextNumber()
    Do While FileExists(pdfPath & ReceiptFileName(number))
        number = number + 1
    Loop
    relativeName = ReceiptFileName(number)
    pdfPath = pdfPath & relativeName
    ' Keep a durable filled-in copy until the PDF is successfully exported.
    ThisWorkbook.Save
    oldCounter = WS(SETTINGS).Range("E2").Value2
    itemBackup = form.Range(form.Cells(firstRow, 2), form.Cells(totalRow - 1, 10)).formula
    noteBackup = form.Range(form.Cells(firstRow, 12), form.Cells(totalRow - 1, 13)).formula
    clientCode = form.Range("A8").formula
    clientName = form.Range("B8").formula
    driverBackup = form.Range("B6").formula
    returnBackup = form.Range("D6").formula
    invoiceBackup = form.Range("F6").formula
    backupReady = True
    BusyPrinting = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    WS(SETTINGS).Range("E2").Value2 = number
    form.Calculate
    PrepareReceiptCopy number, totalRow, temporary, receipt
    ExportReceiptPDF receipt, pdfPath
    pdfSaved = True
    logRow = LastRow(log, 15) + 1
    If logRow < 5 Then logRow = 5
    If logRow > 5 Then
        log.Range("A5:O5").Copy
        log.Range(log.Cells(logRow, 1), log.Cells(logRow, 15)).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End If
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
    log.Cells(logRow, 11).NumberFormat = "@"
    log.Cells(logRow, 11).Value2 = CStr(form.Range("D6").Value2)
    log.Cells(logRow, 12).NumberFormat = "yyyy-mm-dd"
    If Len(CellText(form.Range("F6"))) > 0 Then log.Cells(logRow, 12).Value2 = form.Range("F6").Value2
    log.Cells(logRow, 13).NumberFormat = "@"
    log.Cells(logRow, 13).Value2 = CStr(form.Range("B6").Value2)
    log.Cells(logRow, 14).Value2 = form.Range("L35").Value2
    log.Cells(logRow, 15).Value2 = ReceiptWarehouses(firstRow, totalRow)
    log.Cells(logRow, 15).WrapText = True
    log.rows(logRow).AutoFit
    log.Cells(logRow, 9).Value2 = UI(42)
    log.Hyperlinks.Add anchor:=log.Cells(logRow, 10), address:=relativeName, TextToDisplay:=relativeName
    resetStarted = True
    form.Range("A8").MergeArea.ClearContents
    form.Range("B8").MergeArea.ClearContents
    form.Range(form.Cells(firstRow, 2), form.Cells(totalRow - 1, 10)).ClearContents
    form.Range(form.Cells(firstRow, 12), form.Cells(totalRow - 1, 13)).ClearContents
    form.Range("B6").ClearContents
    form.Range("D6").ClearContents
    form.Range("F6").ClearContents
    WS(SETTINGS).Range("E2").Value2 = number + 1
    BuildList True, ""
    BuildList False, ""
    RefreshTotals firstRow, totalRow
    Application.Calculation = oldCalculation
    form.Activate
    ThisWorkbook.Save
    committed = True
    If paper Then
        On Error Resume Next
        receipt.PrintOut copies:=copyCount, Collate:=True
        If Err.number <> 0 Then paperError = Err.Description
        Err.Clear
        On Error GoTo Failed
        If Len(paperError) = 0 Then log.Cells(logRow, 9).Value2 = UI(45) Else log.Cells(logRow, 9).Value2 = UI(46)
        On Error Resume Next
        ThisWorkbook.Save
        If Err.number <> 0 Then
            statusError = Err.Description
            Err.Clear
        End If
        On Error GoTo Failed
    End If
    CloseReceiptCopy temporary
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    BusyPrinting = False
    form.Activate
    form.Range("B6").Select
    If Len(paperError) > 0 Then
        UnicodeMessage UI(46) & vbCrLf & pdfPath & vbCrLf & paperError & vbCrLf & statusError & vbCrLf & UI(20), vbExclamation
    ElseIf Len(statusError) > 0 Then
        UnicodeMessage UI(81) & vbCrLf & statusError, vbExclamation
    End If
    Exit Sub
Failed:
    errorText = Err.Description
    On Error Resume Next
    If Not committed And backupReady Then
        If resetStarted Then
            RestoreCells form.Range(form.Cells(firstRow, 2), form.Cells(totalRow - 1, 10)), itemBackup
            RestoreCells form.Range(form.Cells(firstRow, 12), form.Cells(totalRow - 1, 13)), noteBackup
            form.Range("B6").formula = driverBackup
            RefreshDriver
            form.Range("D6").formula = returnBackup
            form.Range("F6").formula = invoiceBackup
            form.Range("A8").formula = clientCode
            form.Range("B8").formula = clientName
        End If
        WS(SETTINGS).Range("E2").Value2 = oldCounter
        If logRow >= 5 Then
            log.Range(log.Cells(logRow, 1), log.Cells(logRow, 15)).Hyperlinks.Delete
            log.Range(log.Cells(logRow, 1), log.Cells(logRow, 15)).ClearContents
        End If
        form.Calculate
    End If
    CloseReceiptCopy temporary
    Application.Calculation = oldCalculation
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
    Dim log As Excel.Worksheet, address As String, path As String, number As Long
    Dim receipt As Excel.Worksheet, oldEvents As Boolean, oldBusy As Boolean, errorText As String
    oldEvents = Application.EnableEvents
    oldBusy = BusyPrinting
    On Error GoTo Failed
    Set log = ThisWorkbook.Worksheets("Archive")
    If row < 5 Then Err.Raise vbObjectError + 922, "OpenLogPDF", UI(54)
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
        address = ReceiptFileName(number, CStr(log.Cells(row, 4).Value2))
        path = FolderPath() & address
        Application.EnableEvents = False
        BusyPrinting = True
        If Not FileExists(path) Then ExportReceiptPDF receipt, path
        If Not ThisWorkbook.ReadOnly Then
            log.Cells(row, 10).Hyperlinks.Delete
            log.Hyperlinks.Add anchor:=log.Cells(row, 10), address:=address, TextToDisplay:=address
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
    Dim log As Excel.Worksheet
    Set log = ThisWorkbook.Worksheets("Archive")
    If ActiveSheet.name <> log.name Or ActiveCell.row < 5 Then UnicodeMessage UI(25), vbInformation: Exit Sub
    OpenLogPDF ActiveCell.row
End Sub

Public Sub ReprintSavedSheet(ByVal sheetName As String)
    Dim requested As Variant, copies As Long, previousBusy As Boolean, errorText As String
    If Not IsSavedReceiptSheet(sheetName) Then Exit Sub
    previousBusy = BusyPrinting
    On Error GoTo Failed
    Do
        requested = Application.InputBox(prompt:=UI(59), title:=UI(60), Default:=1, Type:=1)
        If VarType(requested) = vbBoolean Then Exit Sub
        If requested >= 1 And requested <= 999 Then
            If requested = Fix(requested) Then Exit Do
        End If
        UnicodeMessage UI(61), vbExclamation
    Loop
    copies = CLng(requested)
    If Not Application.Dialogs(xlDialogPrinterSetup).show Then Exit Sub
    BusyPrinting = True
    FitReceiptPage ThisWorkbook.Worksheets(sheetName)
    ThisWorkbook.Worksheets(sheetName).PrintOut copies:=copies, Collate:=True
    BusyPrinting = previousBusy
    Exit Sub
Failed:
    errorText = Err.Description
    BusyPrinting = previousBusy
    UnicodeMessage UI(26) & vbCrLf & errorText, vbExclamation
End Sub

Public Function LookupData(ByVal Customer As Boolean) As Variant
    EnsureSearchCache Customer
    If Customer Then LookupData = ClientCache Else LookupData = ProductCache
End Function

Public Function NormalizeLookup(ByVal text As String) As String
    NormalizeLookup = Normal(text)
End Function

Public Function LookupCaption(ByVal item As Long) As String
    Select Case item
    Case 12: LookupCaption = DecodeText("0627 0644 0643 0648 062F")
    Case 11: LookupCaption = DecodeText("0627 0644 0627 0633 0645")
    Case 10: LookupCaption = DecodeText("0627 0628 062D 062B 0020 0628 0627 0644 0627 0633 0645 0020 0623 0648 0020 0627 0644 0643 0648 062F 0020 062B 0645 0020 0627 062E 062A 0631 0020 0645 0646 0020 0627 0644 0646 062A 0627 0626 062C")
    Case 9: LookupCaption = DecodeText("0628 062D 062B 0020 0627 0644 0623 0635 0646 0627 0641")
    Case 8: LookupCaption = DecodeText("0628 062D 062B 0020 0627 0644 0639 0645 0644 0627 0621")
    Case 1: LookupCaption = DecodeText("0628 062D 062B 0020 0627 0644 0639 0645 0644 0627 0621 0020 2014 0020 0627 0644 0627 0633 0645 0020 0623 0648 0020 0627 0644 0643 0648 062F 0020 0623 0648 0020 0622 062E 0631 0020 0034 2013 0035 0020 0623 0631 0642 0627 0645")
    Case 2: LookupCaption = DecodeText("0628 062D 062B 0020 0627 0644 0623 0635 0646 0627 0641 0020 2014 0020 0627 0644 0627 0633 0645 0020 0623 0648 0020 0627 0644 0643 0648 062F")
    Case 3: LookupCaption = DecodeText("0639 062F 062F 0020 0627 0644 0646 062A 0627 0626 062C 003A 0020")
    Case 4: LookupCaption = DecodeText("0627 0643 062A 0628 0020 0644 0644 0628 062D 062B 0020 2022 0020 0627 0644 0623 0633 0647 0645 0020 0644 0644 062A 0646 0642 0644 0020 2022 0020 0045 006E 0074 0065 0072 0020 0644 0644 0627 062E 062A 064A 0627 0631 0020 2022 0020 0045 0073 0063 0020 0644 0644 0625 0644 063A 0627 0621")
    Case 5: LookupCaption = DecodeText("0627 062E 062A 064A 0627 0631")
    Case 6: LookupCaption = DecodeText("0625 0644 063A 0627 0621")
    Case 7: LookupCaption = DecodeText("0644 0627 0020 062A 0648 062C 062F 0020 0646 062A 0627 0626 062C 0020 0645 0637 0627 0628 0642 0629 061B 0020 0639 062F 0651 0644 0020 0646 0635 0020 0627 0644 0628 062D 062B 002E")
    End Select
End Function



Public Function IsMainSheet(ByVal sheet As Object) As Boolean
    If TypeOf sheet Is Excel.Worksheet Then IsMainSheet = (sheet.CodeName = WS(MAIN).CodeName)
End Function

Public Sub CopyLookupKeys(ByVal Customer As Boolean, ByRef keys() As String)
    EnsureSearchCache Customer
    If Customer Then
        keys = ClientCacheKeys
    Else
        keys = ProductCacheKeys
    End If
End Sub
