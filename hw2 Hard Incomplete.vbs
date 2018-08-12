
Sub AlphabticalTesting():

  For Each WS In Worksheets

    Dim WorksheetName As String
    LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Ticker_Total As Double
    Ticker_Total = 0

    Dim Yearly_Change As Double
    
  'Set Summary Table Row
    Dim ST_Row As Integer
    ST_Row = 2

    WS.Range("I1").Value = "Ticker Name"
    WS.Range("J1").Value = "Yearly Change"
    WS.Range("K1").Value = "Percent Change"
    WS.Range("L1").Value = "Total Stock Volumn"


      For i = 2 To LastRow

          If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 
          Ticker_Total = Ticker_Total + Cells(i, 7).Value

          WS.Range("I" & ST_Row).Value = Cells(i, 1).Value
          WS.Range("J" & ST_Row).Value = Round(Yearly_Change, 2)
          WS.Range("K" & ST_Row).Value = Percent_Change
          WS.Range("K" & ST_Row).NumberFormat = "0.00%"
          WS.Range("L" & ST_Row).Value = Ticker_Total

          ST_Row = ST_Row + 1
      
          Ticker_Total = 0

          Else

          Ticker_Total = Ticker_Total + Cells(i, 7).Value

        End If

        Open_Price = Cells(i, 3).Value
        Close_Price = Cells(i, 6).Value
    
        Yearly_Change = Close_Price - Open_Price

        Percent_Change = Yearly_Change / Open_Price


                If Yearly_Change > 0 Then
                    WS.Range("J" & ST_Row).Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    WS.Range("J" & ST_Row).Interior.ColorIndex = 3
                Else
                    WS.Range("J" & ST_Row).Interior.ColorIndex = 0
                End If
        
      Next i
  


    'Hard: Incompelete

      WS.Range("O1").Value = "Ticker_Name"
      WS.Range("P1").Value = "Value"
      WS.Cells(2, 14).Value = "Greatest % increase"
      WS.Cells(3, 14).Value = "Greatest % decrease"
      WS.Cells(4, 14).Value = "Greatest Total Volume "
      
      

      'Have bug in In
      
      If WS.Range("P2").Value = Application.WorksheetFunction.Max(Range("K2:K") Then
        WS.Range("O2").Value = WS.Range("I2:I").Value

      ElseIf WS.Range("P3").Value = Application.WorksheetFunction.Min(Range("K2:K") Then
        WS.Range("O3").Value = WS.Range("I2: I").Value

      ElseIf WS.Range("P4").Value = Application.WorksheetFunction.Max(Range("L2:L") Then
        WS.Range("O4").Value = WS.Range("I2:I").Value

     End If



  Next WS
  
  End Sub
