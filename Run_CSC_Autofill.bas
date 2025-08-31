Option Explicit

'Friendly CSC autofill macro with interactive prompts and error handling
Public Sub Run_CSC_Autofill()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsControl As Worksheet, wsTickers As Worksheet
    Dim wsScrub As Worksheet, wsBench As Worksheet, wsInputs As Worksheet
    Dim arrTicks() As String, tick As String
    Dim dict As Object, nameStr As Variant
    Dim i As Long, j As Long, row As Long, t As Single, blk As Variant
    Dim blocks As Variant, sDate As Variant, comp As Variant, idx As Long
    Dim retry As VbMsgBoxResult

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    'get required sheets
    On Error Resume Next
    Set wsControl = wb.Worksheets("Control")
    Set wsTickers = wb.Worksheets("Tickers")
    Set wsScrub = wb.Worksheets("CSC Scrub Adjustments")
    Set wsBench = wb.Worksheets("Benchmarking")
    On Error GoTo ErrHandler
    If wsControl Is Nothing Or wsTickers Is Nothing Or _
       wsScrub Is Nothing Or wsBench Is Nothing Then
        MsgBox "One or more required sheets are missing.", vbCritical, "CSC Autofill"
        GoTo Cleanup
    End If

    'ensure Inputs sheet exists
    On Error Resume Next
    Set wsInputs = wb.Worksheets("Inputs")
    On Error GoTo ErrHandler
    If wsInputs Is Nothing Then
        Set wsInputs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsInputs.Name = "Inputs"
    End If
    wsInputs.Columns("A:B").ClearContents
    wsInputs.Range("A1:B1").Value = Array("Company_Name", "Ticker")

    MsgBox "Welcome to the CSC autofill wizard!" & vbCrLf & _
           "You'll be prompted for a pricing date and company names.", _
           vbInformation, "CSC Autofill"

    'pricing date
    Do
        sDate = Application.InputBox("Pricing Date? (mm/dd/yyyy)", _
                                      "CSC Autofill", Date, Type:=2)
        If sDate = False Then GoTo Cleanup
        If IsDate(sDate) Then
            wsControl.Range("D4").Value = CDate(sDate)
            Exit Do
        End If
        MsgBox "Please enter a valid date.", vbExclamation, "CSC Autofill"
    Loop
    wsControl.Range("D5").Value = 1
    wsControl.Range("D6").Value = "USD"

    'gather company names interactively
    Set dict = CreateObject("Scripting.Dictionary")
    idx = 1
    Do
        comp = Application.InputBox("Enter company name #" & idx & _
               " (blank or Cancel to finish)", "Company Entry", Type:=2)
        If comp = False Then Exit Do
        comp = Trim(comp)
        If comp <> "" Then
            If Not dict.Exists(comp) Then
                dict.Add comp, comp
                idx = dict.Count + 1
            Else
                MsgBox "Company already entered.", vbExclamation, "CSC Autofill"
            End If
        ElseIf dict.Count = 0 Then
            MsgBox "Please enter at least one company.", vbExclamation, "CSC Autofill"
        Else
            Exit Do
        End If
    Loop
    If dict.Count = 0 Then GoTo Cleanup

    'query FactSet tickers
    For Each nameStr In dict.Items
RetryLookup:
        row = wsTickers.Cells(wsTickers.Rows.Count, "B").End(xlUp).Row + 1
        If row < 9 Then row = 9
        wsTickers.Cells(row, "B").Value = nameStr
        t = Timer
        Do While wsTickers.Cells(row, "C").Value = "" And Timer - t < 5
            DoEvents
        Loop
        tick = wsTickers.Cells(row, "C").Value
        If tick = "" Then
            retry = MsgBox("No ticker found for '" & nameStr & "'. Retry?", _
                           vbRetryCancel + vbExclamation, "CSC Autofill")
            If retry = vbRetry Then
                comp = Application.InputBox("Enter a different name", "Retry", nameStr, Type:=2)
                If comp <> False And Trim(comp) <> "" Then
                    nameStr = comp
                    GoTo RetryLookup
                End If
            End If
        Else
            If MsgBox("'" & nameStr & "' resolved to '" & tick & _
                      "'. Use this?", vbYesNo + vbQuestion, "CSC Autofill") = vbNo Then
                comp = Application.InputBox("Enter the correct name", "Correct", nameStr, Type:=2)
                If comp <> False And Trim(comp) <> "" Then
                    nameStr = comp
                    GoTo RetryLookup
                Else
                    tick = ""
                End If
            End If
        End If
        If tick <> "" Then
            i = i + 1
            ReDim Preserve arrTicks(1 To i)
            arrTicks(i) = tick
            wsInputs.Cells(i + 1, "A").Value = nameStr
            wsInputs.Cells(i + 1, "B").Value = tick
        End If
    Next nameStr
    wsInputs.Columns("A:B").AutoFit
    wsTickers.Columns("B:C").AutoFit
    If i = 0 Then
        MsgBox "No valid tickers captured.", vbExclamation, "CSC Autofill"
        GoTo Cleanup
    End If

    'populate CSC Scrub Adjustments
    For j = 1 To i
        wsScrub.Cells(7, 3 + j).Value = arrTicks(j)
    Next j
    wsScrub.Rows(7).EntireRow.AutoFit

    'trim benchmarking blocks
    blocks = Array("B16:B31", "B42:B55", "B64:B79", "B88:B101", _
                   "B111:B127", "B136:B151", "B160:B175", "B184:B189", _
                   "B208:B223", "B232:B247", "B256:B271", "B280:B293")
    For Each blk In blocks
        Dim rng As Range, total As Long, rStart As Long
        Set rng = wsBench.Range(blk)
        total = rng.Rows.Count
        rStart = rng.Row
        For j = 1 To Application.Min(i, total)
            rng.Cells(j, 1).Value = arrTicks(j)
        Next j
        If i < total Then
            wsBench.Rows(rStart + i & ":" & rStart + total - 1).Delete
        End If
    Next blk
    wsBench.Columns("B").AutoFit

    'trigger FactSet refresh
    On Error Resume Next
    wsControl.Shapes("Search").OLEFormat.Object.Verb xlPrimary
    If Err.Number <> 0 Then
        Err.Clear
        Application.Run "Search"
    End If
    On Error GoTo ErrHandler

    MsgBox "CSC autofill complete for " & i & " companies.", _
           vbInformation, "CSC Autofill"

Cleanup:
    wsControl.Range("D5").Value = 0
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "CSC Autofill"
    Resume Cleanup
End Sub
