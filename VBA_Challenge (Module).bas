Attribute VB_Name = "Module1"
Sub Stock_Analysis()
    ' Declare variables:
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Total_Stock_Volume As Double
    
    ' Assign values
    OpenPrice = 0
    ClosePrice = 0
    YearlyChange = 0
    PercentChange = 0
    Total_Stock_Volume = 0
 
    ' Keep track of the summary table row location:
    Dim Summary_Table_Row As Integer
    
 
        ' Loop through each worksheet in this workbook:
        For Each ws In Worksheets
        
            Summary_Table_Row = 2
            
            ' Create column headings:
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
                ' Retrieve the last row in each worksheet
                LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
                ' Set the open price value:
                OpenPrice = ws.Cells(2, 3).Value
        
                ' Loop through all stock market data:
                For i = 2 To LastRow
        
                ' Evaluate if the next match is the last for a specific ticker:
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
                    ' Set the ticker and close price value:
                    Ticker = ws.Cells(i, 1).Value
                    ClosePrice = ws.Cells(i, 6).Value
    
                    ' Calculate yearly change, percent change, and total stock volume:
                    YearlyChange = ClosePrice - OpenPrice
                    PercentChange = YearlyChange / OpenPrice
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
                    ' Print the values into the summary table:
                    ws.Range("I" & Summary_Table_Row).Value = Ticker
                    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                    ws.Range("K" & Summary_Table_Row).Value = PercentChange
                    ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                                   
                    If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                
                    ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    
                    End If
                    
                    ' Reset the open price, yearly change, percent change, and total stock volume:
                    OpenPrice = ws.Cells(i + 1, 3).Value
                    YearlyChange = 0
                    PercentChange = 0
                    Total_Stock_Volume = 0
                    
                    ' Increment the summary table row:
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                ' If the cell immediately following a row is the same ticker:
                Else
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                End If
                 
            Next i
    
                ' Declare variables for summary table two:
                Dim maxIncrease As Double
                Dim maxDecrease As Double
                Dim maxVolume As Double
                Dim maxInc_Ticker As String
                Dim maxDec_Ticker As String
                Dim maxVol_Ticker As String
                
                'Set variable values:
                maxIncrease = ws.Cells(2, 11).Value
                maxDecrease = ws.Cells(2, 11).Value
                maxVolume = ws.Cells(2, 12).Value
                maxInc_Ticker = ws.Cells(2, 9).Value
                maxDec_Ticker = ws.Cells(2, 9).Value
                maxVol_Ticker = ws.Cells(2, 9).Value
                
                ' Retrieve the summary table one last row in each worksheet:
                LastRow_Summary_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
                    
                    ' Loop through summary table one:
                    For i = 2 To LastRow_Summary_Table
            
                        ' Create row heading/titles:
                        ws.Range("O1").Value = "Ticker"
                        ws.Range("P1").Value = "Value"
                        ws.Range("N2").Value = "Greatest % increase"
                        ws.Range("N3").Value = "Greatest % decrease"
                        ws.Range("N4").Value = "Greatest total volume"
                
                            'Return the greatest % increase, greatest % decrease, and greatest total volume tickers and values:
                            If maxIncrease < ws.Cells(i, 11).Value Then
                            maxIncrease = ws.Cells(i, 11).Value
                            maxInc_Ticker = ws.Cells(i, 9).Value
                            End If
                        
                            If maxDecrease > ws.Cells(i, 11).Value Then
                            maxDecrease = ws.Cells(i, 11).Value
                            maxDec_Ticker = ws.Cells(i, 9).Value
                            End If
                        
                            If maxVolume < ws.Cells(i, 12).Value Then
                            maxVolume = ws.Cells(i, 12).Value
                            maxVol_Ticker = ws.Cells(i, 9).Value
                    
                    End If
        
                Next i
        
                ' Print greatest % increase , greatest % decrease, and greatest total volume tickers and values:
                ws.Range("O2").Value = maxInc_Ticker
                ws.Range("O3").Value = maxDec_Ticker
                ws.Range("O4").Value = maxVol_Ticker
                ws.Range("P2").Value = maxIncrease
                ws.Range("P3").Value = maxDecrease
                ws.Range("P4").Value = maxVolume
 
    Next ws
 
End Sub



