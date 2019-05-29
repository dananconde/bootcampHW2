VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_market_analyst_HARD()

' Last row in columns

Dim lrow As Long
lrow_1 = Cells(Rows.Count, 1).End(xlUp).Row

' Printing appropriate titles

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

' ticker, yearly change, percent change, total volume, and current row in summary table

Dim ticker As String

Dim yearly_change As Double

Dim percent_change As Double

Dim total_volume As Double

Dim summary_table_row As Long
summary_table_row = 2

' New Ticker Rows

Dim new_ticker_row As Long
new_ticker_row = 2

' Populating summary table (1st loop w 'i')

For i = 2 To lrow_1

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ticker = Cells(i, 1).Value
      
      yearly_change = (Cells(i, 6).Value - Cells(new_ticker_row, 3).Value)
      
        ' Trying to avoid Dividing by 0 error
      
        If Cells(new_ticker_row, 3).Value = 0 Then
      
            percent_change = 0
            
        Else
        
            percent_change = Round((yearly_change / Cells(new_ticker_row, 3).Value) * 100, 2)
            
        End If

      total_volume = total_volume + Cells(i, 7).Value

      ' Print ticker, percent change, and total volume in summary table
      Range("I" & summary_table_row).Value = ticker
      Range("K" & summary_table_row).Value = percent_change
      Range("L" & summary_table_row).Value = total_volume
      
      ' Conditional formatting based on positive and negative yearly change
      
        If yearly_change < 0 Then
        
            Range("J" & summary_table_row).Value = yearly_change
            Range("J" & summary_table_row).Interior.ColorIndex = 3
        
        Else
        
            Range("J" & summary_table_row).Value = yearly_change
            Range("J" & summary_table_row).Interior.ColorIndex = 4
            
        End If
            
      ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      ' Assigning new ticker rows
      new_ticker_row = i + 1
      
      ' Resetting total volume
      total_volume = 0

    ' Same ticker
    Else

      ' Just add to total volume
      total_volume = total_volume + Cells(i, 7).Value

    End If

Next i

' Table for greatest values

Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

Dim gp_increase As Double
gp_increase = 0
Dim gp_decrease As Double
gp_decrease = 0
Dim gt_volume As Double
gt_volume = 0

Dim gpi_ticker As String
Dim gpd_ticker As String
Dim gtv_ticker As String

' Searching for greatest values in summary table (2nd loop w 'j')

Dim lrow_2 As Long
lrow_2 = Cells(Rows.Count, 9).End(xlUp).Row

For j = 2 To lrow_2

    ' Negative case
    
    If Cells(j, 11).Value < 0 Then
    
        If Cells(j, 11).Value < gp_decrease Then
        
            gp_decrease = Cells(j, 11).Value
            gpd_ticker = Cells(j, 9).Value
            
        End If
            
    ' Positive case
            
    Else
    
        If Cells(j, 11).Value > gp_increase Then
        
            gp_increase = Cells(j, 11).Value
            gpi_ticker = Cells(j, 9).Value
            
        End If
            
    End If
    
Next j

Range("P2").Value = gpi_ticker
Range("Q2").Value = gp_increase
Range("P3").Value = gpd_ticker
Range("Q3").Value = gp_decrease
        
For k = 2 To lrow_2

    If Cells(k, 12).Value > gt_volume Then
    
        gt_volume = Cells(k, 12).Value
        gtv_ticker = Cells(k, 9).Value
        
    End If

Next k

Range("P4").Value = gtv_ticker
Range("Q4").Value = gt_volume

Range("O6").Value = "Proof of"
Range("O7").Value = "Legitimate Screenshot"
Range("O8").Value = "by Dan Conde"

End Sub





