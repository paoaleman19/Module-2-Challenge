
Sub Stocks()

    Dim wkst As Worksheet

        
    For Each wkst In Worksheets

        'Heading names 1st part
            wkst.Range("I1").Value = "Ticker"
            wkst.Range("J1").Value = "Yearly Change"
            wkst.Range("K1").Value = "Percent Change"
            wkst.Range("L1").Value = "Total Stock Volume"

        'Heading names 2nd part
            wkst.Range("Q1").Value = "Ticker"
            wkst.Range("R1").Value = "Value"
            wkst.Range("P2").Value = "Greatest  % Increase"
            wkst.Range("P3").Value = "Greatest  % Decrease"
            wkst.Range("P4").Value = "Greatest  Total Volume"
        
        Dim Ticker As String
        Dim CurrentRow As Integer
                CurrentRow = 2
        
        Dim BegPrice As Double
                BegPrice = 0
                BegPrice = wkst.Cells(2, 3).Value
        
        Dim EndPrice As Double
                EndPrice = 0
                
        Dim YearlyChange As Double
                YearlyChange = 0
        
        Dim PercentChange As Double
                PercentChange = 0
        
        Dim TotalVolume As Double
                
                
             
             'To get the tickers listed!!
            For i = 2 To wkst.Cells(Rows.Count, 1).End(xlUp).Row
                
                If wkst.Cells(i, 1).Value <> wkst.Cells(i + 1, 1).Value Then
                
                Ticker = wkst.Cells(i, 1).Value
                wkst.Cells(CurrentRow, 9).Value = Ticker
                
                
                YearlyChange = wkst.Cells(i, "F") - BegPrice
                
                wkst.Cells(CurrentRow, 10).Value = YearlyChange
                    
                    'Color format the cells for Yearly Change
                     
                     If YearlyChange > 0 Then
                        wkst.Cells(CurrentRow, 10).Interior.ColorIndex = 4
                    
                    ElseIf YearlyChange <= 0 Then
                        wkst.Cells(CurrentRow, 10).Interior.ColorIndex = 3
                        
                    End If
            
                
                'To calculate the percent change. The formula works but the numbers are not correct!!
                        If BegPrice <> 0 Then
                            PercentChange = (YearlyChange / BegPrice) * 100
                            
                            wkst.Cells(CurrentRow, 11).Value = PercentChange
                       
                       End If
                
                BegPrice = wkst.Cells(i + 1, "C")
                'To calculate the TotalVolume
                
                TotalVolume = TotalVolume + wkst.Cells(i, 7).Value
                wkst.Cells(CurrentRow, 12).Value = TotalVolume
                
                'This would be tto change ticker!
                
                CurrentRow = CurrentRow + 1
                TotalVolume = 0
                
                Else
                
                TotalVolume = TotalVolume + wkst.Cells(i, 7).Value
                
                End If
            
            
            Next i

    Next wkst

End Sub

