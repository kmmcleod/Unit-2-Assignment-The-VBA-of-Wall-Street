Sub StockWS()
    For Each ws In Worksheets
    ws.Activate
    Call StockAnalysis
    Next ws
End Sub