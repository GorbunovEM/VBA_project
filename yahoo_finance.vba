Option Explicit

Public Sub GetYahooInfo()
    Dim tickers(), ticker As Long, lastRow As Long, headers()
    Dim wsSource As Worksheet, http As clsHTTP, html As HTMLDocument

    Application.ScreenUpdating = False

    Set wsSource = ThisWorkbook.Worksheets("Àêöèè Sum") '<== Change as appropriate to sheet containing the tickers
    Set http = New clsHTTP

    headers = Array("Ticker", "Previous Close", "Open", "Bid", "Ask", "Day's Range", "52 Week Range", "Volume", "Avg. Volume", "Market Cap", "Beta", "PE Ratio (TTM)", "EPS (TTM)", _
                    "Earnings Date", "Forward Dividend & Yield", "Ex-Dividend Date", "1y Target Est")

    With wsSource
        lastRow = GetLastRow(wsSource, 1)
        Select Case lastRow
        Case Is < 3
            Exit Sub
        Case 3
            ReDim tickers(1, 1): tickers(1, 1) = .Range("A3").Value
        Case Is > 3
            tickers = .Range("A3:A" & lastRow).Value
        End Select

        ReDim Results(0 To UBound(tickers, 1) - 1)
        Dim i As Long, endPoint As Long
        endPoint = UBound(headers)

        For ticker = LBound(tickers, 1) To UBound(tickers, 1)
            On Error Resume Next
            If Not IsEmpty(tickers(ticker, 1)) Then
                Set html = http.GetHTMLDoc("https://finance.yahoo.com/quote/" & tickers(ticker, 1) & "/?p=" & tickers(ticker, 1))
                Results(ticker - 1) = http.GetInfo(html, endPoint)
                Set html = Nothing
            Else
                Results(ticker) = vbNullString
            End If
        Next

        .Cells(2, 1).Resize(1, UBound(headers) + 1) = headers
        For i = LBound(Results) To UBound(Results)
            .Cells(3 + i, 2).Resize(1, endPoint - 1) = Results(i)
        Next
    End With
    Application.ScreenUpdating = True
End Sub

Public Function GetLastRow(ByVal ws As Worksheet, Optional ByVal columnNumber As Long = 1) As Long
    With ws
        GetLastRow = .Cells(.Rows.Count, columnNumber).End(xlUp).Row
    End With
End Function
