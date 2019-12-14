Sub VBAStocks()

'Loop through all worksheets
For Each ws In Worksheets

    
'Add column headers to summary table
    Dim Summary_Table_Headers() As Variant
    
    Summary_Table_Headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
        ws.Range("I1:L1") = Summary_Table_Headers
        
'Add row and column headers to kpi table
    Dim KPI_Table_Column_Header() As Variant
    
    KPI_Table_Column_Header = Array("Ticker", "Value")
    
        ws.Range("P1:Q1") = KPI_Table_Column_Header
        
        ws.Range("O2") = "Greatest % Increase"
        
        ws.Range("O3") = "Greatest % Decrease"
        
        ws.Range("O4") = "Greatest Total Volume"

        
'Compile summary table with data
    Dim TickerSymbol As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim Summary_Table_Row As Integer
    Dim Max_Percent_Chg As Double
    Dim Min_Percent_Chg As Double
    Dim Max_Total_Volume As Double
    Dim Match_Percent_Chg1 As Double
    Dim Match_Percent_Chg2 As Double
    Dim Match_Total_Volume As Double


    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Table_Row = 2
    TotalStockVolume = 0
    

'Loop through all stocks data and compile data
    For I = 2 To LastRow
    
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
            'Set/calculate values
            TickerSymbol = ws.Cells(I, 1).Value
            
            OpeningPrice = ws.Cells(2, 3).Value
            
            ClosingPrice = ws.Cells(I, 6).Value
            
            YearlyChange = ClosingPrice - OpeningPrice
            
            PercentChange = YearlyChange / OpeningPrice
            
            TotalStockVolume = TotalStockVolume + ws.Cells(I, 7).Value
            

            ws.Range("I" & Summary_Table_Row).Value = TickerSymbol
            
            ws.Range("J" & Summary_Table_Row).Value = YearlyChange
            
            ws.Range("K" & Summary_Table_Row).Value = Format(PercentChange, "Percent")
            
            ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
            
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            TotalStockVolume = 0
            
            
            
            Max_Percent_Chg = WorksheetFunction.Max(ws.Range("K:K"))
            
            Min_Percent_Chg = WorksheetFunction.Min(ws.Range("K:K"))
            
            Max_Total_Volume = WorksheetFunction.Max(ws.Range("L:L"))
            
            ws.Range("Q2").Value = Format(Max_Percent_Chg, "Percent")
            
            ws.Range("Q3").Value = Format(Min_Percent_Chg, "Percent")
            
            ws.Range("Q4").Value = Format(Max_Total_Volume, "General Number")
            
            
            
            Match_Percent_Chg1 = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0)
            
            Match_Percent_Chg2 = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0)
            
            Match_Total_Volume = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0)

            
            ws.Range("P2").Value = ws.Cells(Match_Percent_Chg1, 9)
            
            ws.Range("P3").Value = ws.Cells(Match_Percent_Chg2, 9)
            
            ws.Range("P4").Value = ws.Cells(Match_Total_Volume, 9)
            
           
        Else
        
            TotalStockVolume = TotalStockVolume + ws.Cells(I, 7).Value
            
        
        End If
    
    Next I


'Formatting
    
        ws.Columns("A:Q").AutoFit
        ws.Range("Q4").Value = Format(Max_Total_Volume, "Scientific")
        
        Dim YearlyChange_RG As Range
        Dim cond1 As FormatCondition
        Dim cond2 As FormatCondition
        Set YearlyChange_RG = ws.Range("J2", ws.Range("J2").End(xlDown))
        
        YearlyChange_RG.FormatConditions.Delete
        
        'define the rule for each conditional format
        Set cond1 = YearlyChange_RG.FormatConditions.Add(xlCellValue, xlGreater, 0)
        Set cond2 = YearlyChange_RG.FormatConditions.Add(xlCellValue, xlLess, 0)
        
        'define the format applied for each conditional format
        With cond1
        .Interior.Color = vbGreen
        End With
        
        With cond2
        .Interior.Color = vbRed
        End With
    
Next ws

End Sub