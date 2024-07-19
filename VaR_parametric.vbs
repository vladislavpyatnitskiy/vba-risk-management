Sub CalculateVaR()
    Dim ws As Worksheet
    Dim assetPrices As Range
    Dim returns As Range
    Dim VaR As Double
    Dim confidenceLevel As Double
    Dim numDays As Integer
    Dim i As Integer
    
    ' Set worksheet and data ranges
    Set ws = ThisWorkbook.Sheets("Data")
    Set assetPrices = ws.Range("A2:A101")
    Set returns = ws.Range("B2:B101")
    
    ' Calculate daily returns
    For i = 2 To assetPrices.Rows.Count
        returns.Cells(i - 1, 1).Value = (assetPrices.Cells(i, 1).Value / assetPrices.Cells(i - 1, 1).Value) - 1
    Next i
    
    ' Set confidence level (e.g., 95%)
    confidenceLevel = 0.95
    numDays = returns.Rows.Count
    
    ' Sort returns to find the VaR
    returns.Sort Key1:=returns.Cells(1, 1), Order1:=xlAscending
    VaR = returns.Cells(Application.WorksheetFunction.RoundUp((1 - confidenceLevel) * numDays, 0), 1).Value
    
    ' Output VaR
    ws.Range("E2").Value = "VaR (95% Confidence Level)"
    ws.Range("E3").Value = VaR
    
    MsgBox "VaR Calculation Complete"
End Sub
