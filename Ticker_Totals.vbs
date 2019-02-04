Sub Ticker_Totals()

    ' Loop through all sheets
    For Each ws In Worksheets

  ' Set an initial variable for holding the ticker symbol
  Dim Ticker_Symbol As String

  ' Set an initial variable for holding the total stock volume
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Set an initial variable for holding the ticker's opening price for the year
  Dim Ticker_Year_Opening_Date_Price(1) As String
  Ticker_Year_Opening_Date_Price(0) = "30001231"
  Ticker_Year_Opening_Date_Price(1) = ""

  ' Set an initial variable for holding the ticker's closing price for the year
  Dim Ticker_Year_Closing_Date_Price(1) As String
  Ticker_Year_Closing_Date_Price(0) = "19000101"
  Ticker_Year_Closing_Date_Price(1) = ""

  ' Set an initial variable for holding the yearly change
  Dim Year_Change As Double
  Year_Change = 0

  ' Set an initial variable for holding the percent change
  Dim Pct_Change As Double
  Pct_Change = 0

  ' Set an initial variable for holding greatest values
  Dim Greatest_Values(5) As String
  Greatest_Values(0) = ""
  Greatest_Values(1) = "0"
  Greatest_Values(2) = ""
  Greatest_Values(3) = "0"
  Greatest_Values(4) = ""
  Greatest_Values(5) = "0"

  ' Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Print Headers
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
  ' Determine the Last Row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all ticker symbol
  For i = 2 To LastRow

    ' Check if we are still within the same ticker symbol, if it is not...
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

      ' Set the Ticker symbol
      Ticker_Symbol = ws.Cells(i, 1).Value

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

      ' Print the Ticket Symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the Stock Volume Total to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Volume_Total

      ' Get year opening price
      If Int(Str(ws.Cells(i, 2).Value)) < Int(Ticker_Year_Opening_Date_Price(0)) Then
    
        ' Set the Ticker symbol
        Ticker_Year_Opening_Date_Price(0) = Str(ws.Cells(i, 2).Value)
        Ticker_Year_Opening_Date_Price(1) = Str(ws.Cells(i, 3).Value)
    
      End If

      ' Get year closing price
      If Int(Str(ws.Cells(i, 2).Value)) > Int(Ticker_Year_Closing_Date_Price(0)) Then
    
        ' Set the Ticker symbol
        Ticker_Year_Closing_Date_Price(0) = Str(ws.Cells(i, 2).Value)
        Ticker_Year_Closing_Date_Price(1) = Str(ws.Cells(i, 6).Value)
    
      End If
  
      Year_Change = CDbl(Ticker_Year_Closing_Date_Price(1)) - CDbl(Ticker_Year_Opening_Date_Price(1))
      
      If CDbl(Ticker_Year_Opening_Date_Price(1)) <> 0 Then

        Pct_Change = Year_Change / CDbl(Ticker_Year_Opening_Date_Price(1))
     
      Else

        Pct_Change = 0
        
      End If
      
      ws.Range("J" & Summary_Table_Row).Value = Year_Change
      
      If Year_Change > 0 Then
      
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
      ElseIf Year_Change < 0 Then
      
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
      
      End If
      
      ws.Range("K" & Summary_Table_Row).Value = Pct_Change
      
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' ws.Range("M" & Summary_Table_Row).Value = Ticker_Year_Opening_Date_Price(0) + " $" + Ticker_Year_Opening_Date_Price(1)
      
      ' ws.Range("N" & Summary_Table_Row).Value = Ticker_Year_Closing_Date_Price(0) + " $" + Ticker_Year_Closing_Date_Price(1)
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Volume_Total = 0
      
      ' Reset Opening
      Ticker_Year_Opening_Date_Price(0) = "30001231"
      Ticker_Year_Opening_Date_Price(1) = ""
      
      ' Reset Closing
      Ticker_Year_Closing_Date_Price(0) = "19000101"
      Ticker_Year_Closing_Date_Price(1) = ""
      
      ' Reset Change variables
      Year_Change = 0
      Pct_Change = 0

    ' If the cell immediately following a row is the same ticker symbol...
    Else

      ' Add to the Stock Volume Total
      Volume_Total = Volume_Total + ws.Cells(i, 7).Value

      ' Get year opening price
      If Int(Str(ws.Cells(i, 2).Value)) < Int(Ticker_Year_Opening_Date_Price(0)) Then
    
      ' Set the Ticker symbol
      Ticker_Year_Opening_Date_Price(0) = Str(ws.Cells(i, 2).Value)
      Ticker_Year_Opening_Date_Price(1) = Str(ws.Cells(i, 3).Value)
    
      End If

      ' Get year closing price
      If Int(Str(ws.Cells(i, 2).Value)) > Int(Ticker_Year_Closing_Date_Price(0)) Then
    
      ' Set the Ticker symbol
      Ticker_Year_Closing_Date_Price(0) = Str(ws.Cells(i, 2).Value)
      Ticker_Year_Closing_Date_Price(1) = Str(ws.Cells(i, 6).Value)
    
      End If

    End If

  Next i
  
  ' Determine the Last Row
  LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

  ' Loop through all ticker symbol
  For i = 2 To LastRow
    
    ' Determine Greatest Increase
    If CDbl(Str(ws.Cells(i, 11).Value)) > CDbl(Greatest_Values(1)) Then
    
        ' Set the Ticker symbol
        Greatest_Values(0) = ws.Cells(i, 9).Value
        Greatest_Values(1) = CDbl(ws.Cells(i, 11).Value)
    
    End If
    
    ' Determine Greatest Decrease
    If CDbl(Str(ws.Cells(i, 11).Value)) < CDbl(Greatest_Values(3)) Then
    
        ' Set the Ticker symbol
        Greatest_Values(2) = ws.Cells(i, 9).Value
        Greatest_Values(3) = CDbl(ws.Cells(i, 11).Value)
    
    End If
    
    ' Determine Greatest Volume
    If CDbl(Str(ws.Cells(i, 12).Value)) > CDbl(Greatest_Values(5)) Then
    
        ' Set the Ticker symbol
        Greatest_Values(4) = ws.Cells(i, 9).Value
        Greatest_Values(5) = CDbl(ws.Cells(i, 12).Value)
    
    End If
    
 Next i
 
 ws.Range("P2").Value = Greatest_Values(0)
 ws.Range("Q2").Value = CDbl(Greatest_Values(1))
 ws.Range("Q2").NumberFormat = "0.00%"
 ws.Range("P3").Value = Greatest_Values(2)
 ws.Range("Q3").Value = CDbl(Greatest_Values(3))
 ws.Range("Q3").NumberFormat = "0.00%"
 ws.Range("P4").Value = Greatest_Values(4)
 ws.Range("Q4").Value = CDbl(Greatest_Values(5))
 
 ws.Columns("A:Q").AutoFit
 
 Next ws

End Sub
