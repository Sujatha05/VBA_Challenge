
Sub Stock_data()

Dim sheet As Integer

For sheet = 1 To 3

For Each ws In Worksheets

    ws.Range("K1").Value = "Ticker"
    ws.Range("L1").Value = "Yearly Change"
    Range("L:L").NumberFormat = "0.00"
    ws.Range("M1").Value = "Percent Change"
    Range("M:M").NumberFormat = "0.00%"
    ws.Range("N1").Value = "Total Stock Volume"
    
    ws.Range("S1").Value = "Ticker"
    ws.Range("T1").Value = "Value"
    ws.Range("P2").Value = "Greatest % increase"
    ws.Range("P3").Value = "Greatest % decrease"
    ws.Range("P4").Value = "Greatest Total volume"


If sheet = 1 Then
 Worksheets("2018").Activate
   ElseIf sheet = 2 Then
Worksheets("2019").Activate
ElseIf sheet = 3 Then
Worksheets("2020").Activate
End If

 'Determine the Last Row
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
 'Set an initial variable for holding the Ticker Name
  Dim Ticker_Name As String
  
 'Set an initial variable for holding the Open Value
 
 Dim Open_value As Double

 'Set an initial variable for holding the Close Value
  Dim Close_value As Double

  
 'Set an initial variable for total stock volume per ticker
  Dim Stock_Volume As Double
  
  
  Stock_Volume = 0

  ' Keep track of the location for each ticker name in the stock data table

  Dim Stock_Table_Row As Integer
  Stock_Table_Row = 2
  
  'Set an initial variable for holding the value of the yearly change
 
  Dim Yearly_change As Double

 'Set an initial variable for holding the value of the percentage change
  Dim Percentage_Change As Double

  'Loop through all Ticker
    For i = 2 To LastRow
   
   If sheet = 1 And Cells(i, 2).Value = "20180102" Then
   Open_value = Cells(i, 3).Value
  
    ElseIf sheet = 1 And Cells(i, 2).Value = "20181231" Then
    Close_value = Cells(i, 6).Value
   
  
   ElseIf sheet = 2 And Cells(i, 2).Value = "20190102" Then
   Open_value = Cells(i, 3).Value
   
   ElseIf sheet = 2 And Cells(i, 2).Value = "20191231" Then
    Close_value = Cells(i, 6).Value
   
   
   ElseIf sheet = 3 And Cells(i, 2).Value = "20200102" Then
   Open_value = Cells(i, 3).Value
   
   ElseIf sheet = 3 And Cells(i, 2).Value = "20201231" Then
    Close_value = Cells(i, 6).Value
   
   End If
    

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'Close_value = Cells(i, 6).Value
    
 Yearly_change = Close_value - Open_value
 Percentage_Change = (Yearly_change / Open_value)
 
      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Stock Volume
      Stock_Volume = Stock_Volume + Cells(i, 7).Value

      ' Print the Ticker Name in the Stock Table
      Range("K" & Stock_Table_Row).Value = Ticker_Name
      
       'Print the Yearly Change to the Stock Table
      Range("L" & Stock_Table_Row).Value = Yearly_change
      
      
       'Print the Percentage to the Stock Table
      Range("M" & Stock_Table_Row).Value = Percentage_Change
     

     'Print the Stock Volume to the Stock Table
      Range("N" & Stock_Table_Row).Value = Stock_Volume
      
       
      
    


' Add one to the Stock Table Row
      Stock_Table_Row = Stock_Table_Row + 1
      
      ' Reset the Stock Volume
      Stock_Volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the Brand Total
      Stock_Volume = Stock_Volume + Cells(i, 7).Value
      
      If Cells(i, 12).Value < "0" Then
             Cells(i, 12).Interior.ColorIndex = 3
       Else
       Cells(i, 12).Interior.ColorIndex = 4
       End If

    
    End If




  Next i
  

  
  Next ws
  
   
  Next sheet
  

End Sub


