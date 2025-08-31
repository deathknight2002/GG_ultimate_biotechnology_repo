Attribute VB_Name = "test_NBIRevenue"
Option Explicit
Option Private Module
'@folder("SeleniumVBA.Testing")

'Benchmarks LTM revenue composition for provided tickers.
'Users should supply the list of NBI component tickers in the tickers array or
'update the procedure to read them from a worksheet range.
Sub test_BenchmarkNBIRevenue()
    Dim tickers As Variant
    tickers = Array("REGN", "GILD") 'TODO: populate with all NBI component tickers

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    Dim cikMap As Object
    Set cikMap = LoadTickerCikMap(http)

    Dim t As Variant
    For Each t In tickers
        Dim cik As String
        cik = LookupCik(CStr(t), cikMap)
        If cik <> "" Then
            Dim json As Object
            Set json = GetCompanyFacts(http, cik)
            If Not json Is Nothing Then
                Dim totalRev As Double
                totalRev = SumLTM(json, "us-gaap", "RevenueFromContractWithCustomerExcludingAssessedTax")

                Dim partnerRev As Double
                partnerRev = SumLTM(json, "us-gaap", Array("LicensesRevenue", _
                                                            "RoyaltyRevenue", _
                                                            "RevenueRecognitionMilestoneMethodRevenueRecognized"))
                Dim productRev As Double
                productRev = totalRev - partnerRev
                Dim pct As Double
                If totalRev <> 0 Then pct = partnerRev / totalRev
                Debug.Print t & ": Total=" & FormatNumber(totalRev, 0) & _
                            " Partner=" & FormatNumber(partnerRev, 0) & _
                            " Product=" & FormatNumber(productRev, 0) & _
                            " Partner%=" & FormatPercent(pct, 2)
            End If
        End If
    Next t
End Sub

'Loads SEC ticker to CIK mapping into a dictionary
Private Function LoadTickerCikMap(http As Object) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim url As String
    url = "https://www.sec.gov/include/ticker.txt"
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (compatible; NBIRevenueMacro/1.0; +https://example.com)"
    http.send
    Dim lines As Variant
    lines = Split(http.responseText, vbLf)
    Dim line As Variant, parts As Variant
    For Each line In lines
        line = Trim(line)
        If Len(line) > 0 Then
            parts = Split(line, vbTab)
            If UBound(parts) >= 1 Then
                dict(LCase(parts(0))) = parts(1)
            End If
        End If
    Next line
    Set LoadTickerCikMap = dict
End Function

'Looks up a ticker in the mapping and returns a zero padded CIK
Private Function LookupCik(ticker As String, dict As Object) As String
    Dim key As String
    key = LCase(ticker)
    If dict.Exists(key) Then
        LookupCik = Right("0000000000" & dict(key), 10)
    End If
End Function

'Fetches company facts JSON for a CIK
Private Function GetCompanyFacts(http As Object, cik As String) As Object
    Dim url As String
    url = "https://data.sec.gov/api/xbrl/companyfacts/CIK" & cik & ".json"
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (compatible; NBIRevenueMacro/1.0; +https://example.com)"
    http.send
    If http.Status = 200 Then
        Set GetCompanyFacts = WebJsonConverter.ParseJson(http.responseText)
    End If
End Function

'Sums last four reported values for the specified tag or array of tags
Private Function SumLTM(json As Object, taxonomy As String, tags As Variant) As Double
    Dim total As Double
    Dim tag As Variant
    If IsArray(tags) Then
        For Each tag In tags
            total = total + SumLTM(json, taxonomy, CStr(tag))
        Next tag
    Else
        On Error Resume Next
        Dim col As Collection
        Set col = json("facts")(taxonomy)(tags)("units")("USD")
        On Error GoTo 0
        If Not col Is Nothing Then total = total + SumLastFour(col)
    End If
    SumLTM = total
End Function

'Sums the last four entries in a collection of period data
Private Function SumLastFour(col As Collection) As Double
    Dim i As Long, cnt As Long
    For i = col.Count To 1 Step -1
        SumLastFour = SumLastFour + CDbl(col(i)("val"))
        cnt = cnt + 1
        If cnt = 4 Then Exit For
    Next i
End Function
