Sub vbastocks():

    ' Loop through all the worksheets
    For Each ws In Worksheets

        ' Print headers for all the worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
  
  
        ' Set an initial variable for holding the ticker symbol
        Dim Ticker_Symbol As String
        
         ' Set an initial variable for holding the opening price per ticker symbol
        Dim Ticker_Opening As Double
        
        ' Set an initial variable for holding the closing price per ticker symbol
        Dim Ticker_Closing As Double
        
        ' Set an initial variable for holding the yearly change per ticker symbol
        Dim Yearly_Change As Double
        
        ' Set an initial variable for holding the percentage change per ticker symbol
        Dim Percentage_Change As Double
        
        ' Set an initial variable for holding the total volume per ticker symbol
        Dim Ticker_Total As Double
        Ticker_Total = 0
        
        ' Keep track of the location of the opening row per ticker symbol
        Dim Opening_Row As Double
        Opening_Row = 2
        
        ' Keep track of the location for each ticker symbol in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Find the last row of the sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        
        ' Loop through all ticker symbols
        For i = 2 To LastRow
               
        ' Check if we are still within the same ticker symbol, if not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ' Set the ticker symbol
            Ticker_Symbol = ws.Cells(i, 1).Value
            
            ' Print the ticker symbol in the summary table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
            
            ' Set the ticker opening value
            Ticker_Opening = ws.Range("C" & Opening_Row).Value
            
            ' Set the ticker closing value
            Ticker_Closing = ws.Cells(i, 6).Value
            
            ' Set the yearly change per ticker symbol
            Yearly_Change = Ticker_Closing - Ticker_Opening
            
            ' Print the yearly change in the summary table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            ' Set the percentage change per ticker symbol
            If Ticker_Opening = 0 Then
            
                ' Avoids divisions by zero
                Percentage_Change = 0
            Else
                Percentage_Change = Yearly_Change / Ticker_Opening
            End If
            
            ' Change percentage change to percentage format
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        
            ' Print the percentage change in the summary table
            ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
                 
            ' Add to the ticker total
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            
            ' Print the ticker total to the summary table
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
            
            ' Highlight positive change in green and negative change in red
            If ws.Range("J" & Summary_Table_Row).Value > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                
            ElseIf ws.Range("J" & Summary_Table_Row).Value < 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                
            ElseIf ws.Range("J" & Summary_Table_Row).Value = 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 48
                
            End If
            
            ' Update the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Update the opening row
            Opening_Row = i + 1
            
            ' Reset the ticker total
            Ticker_Total = 0
        
        ' If the cell immediately following a row is the same ticker symbol...
        Else
        
        ' Add to the ticker total
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        
        End If
    
      Next i
      
      'Find the last row of the percentage changes
      LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ' Loop through all percentage changes
        For i = 2 To LastRow
        
            ' Check for the greater percentage change
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                
                ' Print greater percentage change to max values summary table
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                
                ' Print corresponding ticker symbol
                ws.Range("P2").Value = ws.Range("I" & i).Value
                
            End If
        
            ' Check for lesser percentage change
            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            
                ' Print lesser percentage change to max values summary table
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                
                ' Print corresponding ticker symbol
                ws.Range("P3").Value = ws.Range("I" & i).Value
                
            End If
        
            ' Check for greater total volume
            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            
                ' Print greater total volume to max values summary table
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                
                ' Print corresponding ticker symbol
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If
            
            ' Change greatest percentage increase and decrease to percentage format
            ws.Range("Q2:Q3").NumberFormat = "0.00%"

        Next i
      
    Next ws

End Sub
