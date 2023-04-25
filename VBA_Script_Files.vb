Sub Ticker_ticker()

'Loop through all sheets
For Each ws In Worksheets

    ' Add the words Ticker/Yearly Change/ Percentage Change/ Total Stock Volume to the First Column Header(Summary table)
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Add the addtional titles for the hard solutions
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
        
    ' Set an initial variable for holding the ticker symbol
    Dim Ticker As String

    ' Set an initial variable for holding the total stock volume/yearly change/percentage change per ticker
    Dim Total_Stock_Volume As Double
    Dim MAX_Total_Stock_Volume As Double
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    Dim MAX_Percentage_Change As Double
    Dim MIN_Percentage_Change As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Total_Stock_Volume = 0
    MAX_Total_Stock_Volume = 0
    Yearly_Change = 0
    Percentage_Change = 0
    MAX_Percentage_Change = 0
    MIN_Percentage_Change = 0
    Open_Price = 0
    Close_Price = 0
    
    ' Determine the Last Row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Set initial open price for first ticker
    Open_Price = ws.Cells(2, 3).Value
    
    ' Loop through all the stocks
    For i = 2 To lastRow

        ' Check if we are still within the same ticker, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
         ' Set the Ticker symbol
        Ticker = ws.Cells(i, 1).Value
        
        'Calculate the yearly change and percentage change
        Close_Price = ws.Cells(i, 6).Value
        Yearly_Change = Close_Price - Open_Price
        Percentage_Change = (Yearly_Change / Open_Price)
        
        ' Add to the Total Stock Volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        ' Print the Ticker Symbol & Total Stock Volume in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
        ' Print the Percentage change and correct its number format in the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        ' Print the Yearly Change in the Summary Table and conditionals applied
        ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            If Yearly_Change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
         
         'Set the open price for the next ticker
         Open_Price = ws.Cells(i + 1, 3).Value
         
         'Determine the Greatest % Increase and Greatest % Decrease of stock
         If Percentage_Change > MAX_Percentage_Change Then
         MAX_Percentage_Change = Percentage_Change
         ws.Range("Q2").Value = MAX_Percentage_Change
         ws.Range("Q2").NumberFormat = "0.00%"
         ws.Range("P2").Value = Ticker
         ElseIf Percentage_Change < MIN_Percentage_Change Then
         MIN_Percentage_Change = Percentage_Change
         ws.Range("Q3").Value = MIN_Percentage_Change
         ws.Range("Q3").NumberFormat = "0.00%"
         ws.Range("P3").Value = Ticker
         End If
         
         'Determine the Greatest Total Stock Volume
         If Total_Stock_Volume > MAX_Total_Stock_Volume Then
         MAX_Total_Stock_Volume = Total_Stock_Volume
         ws.Range("Q4") = MAX_Total_Stock_Volume
         ws.Range("P4") = Ticker
         End If
         
        ' Reset the yearly change/percentage change/total volume stock
         Total_Stock_Volume = 0
         Yearly_Change = 0
         Percentage_Change = 0

        ' If the cell immediately following a row is the same ticker...
        Else

        ' Add to the total stock volume
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        End If

    Next i
    
    ' Autofit to display data
    ws.Columns("I:Q").AutoFit

Next ws

MsgBox ("Ah-ha! Mission Complete")

End Sub