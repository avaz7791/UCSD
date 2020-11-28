Attribute VB_Name = "TickerLoop"
Sub RunTickerLoop()
'Program Script:
'   This script looks through all the worksheets and
'   sumarizes the year effort of a ticker symbol
'Calculations:
'     1)Yearly Change
'     2)Percent Change
'     3)Total Stock Volume
'     4)Conditional formatting to the yearly Change RED<0, GREEN>=0
'Bonus:
'     5) Find the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

'Developer: Amir Vazquez
'Date: 11/28/2020

'-------------------------------------------------------------------------------------------------------
'Variable declaration
Dim Ticker_Symbol As String
Dim Volume As Double
Dim Ticker_YrOpen As Double
Dim Ticker_YrClose As Double
Dim Yearly_Change As Double
Dim Ticker_ChangePer As Double
Dim Table_row As Integer 'variable to track row of ticker in summary table
Dim ws As Worksheet

Dim maxPercIncrease, maxPercDecrease, maxVolume As Double
Dim maxPercIncTicker, maxPercDecTicker, maxVolTicker As String

For Each ws In Worksheets

ws.Select
' Initialize variables
Volume = 0
Ticker_YrOpen = 0
Ticker_YrClose = 0
Table_row = 2
maxPercIncrease = 0
maxPercDecrease = 0
maxVolume = 0

'Add Rows
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "% Change"
Range("L1").Value = "Total Stock Volume"

Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"


Range("i1:l1").Font.Bold = True
Set Ticker_Sheet = Worksheets.Application.ActiveSheet
 'Find the last row in the sheet
LastRowSheet = Ticker_Sheet.Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRowSheet
 
 'check for change in ticker symbol
 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker_Symbol = Cells(i, 1).Value
    Ticker_YrClose = Cells(i, 6).Value
    Range("I" & Table_row).Value = Ticker_Symbol
 
    Volume = Volume + Cells(i, 7).Value
    
    'Yearly Change
    Yearly_Change = Ticker_YrClose - Ticker_YrOpen
    Range("J" & Table_row).Value = Yearly_Change
    
    '% Change
    If Ticker_YrOpen = 0 Then
        Ticker_ChangePer = 0
    Else
        Ticker_ChangePer = Yearly_Change / Ticker_YrOpen
    End If
    
    '-Formatting
    Range("K" & Table_row).Value = Ticker_ChangePer
    Range("K" & Table_row).NumberFormat = "0.00%"     'Style = "Percent"
         
    'Total Stock Volume
    Range("L" & Table_row).Value = Volume

    'check for Greatest % Increase
    If Ticker_ChangePer >= maxPercIncrease Then
        maxPercIncrease = Ticker_ChangePer
        maxPercIncTicker = Ticker_Symbol
    End If
    
    'check for Greaest % Decrease
    If Ticker_ChangePer <= maxPercDecrease Then
        maxPercDecrease = Ticker_ChangePer
        maxPercDecTicker = Ticker_Symbol
    End If
    
    'check for greatest total volume
    If Volume >= maxVolume Then
        maxVolume = Volume
        maxVolTicker = Ticker_Symbol
    End If

    'Add one to the table row for the next ticker symbol
    Table_row = Table_row + 1
    
    'Reset variables for next Ticker
    Volume = 0
    Ticker_YrOpen = 0
    Ticker_YrClose = 0
 Else
 
    If Ticker_YrOpen = 0 Then
        Ticker_YrOpen = Cells(i, 3).Value
    End If
    Volume = Volume + Cells(i, 7).Value
 End If
  
Next i

'create conditional formatting for Yearly Change
    With Range("J2:J" & Table_row - 1).FormatConditions.Add(xlCellValue, xlLessEqual, "0")
    .Interior.Color = RGB(255, 0, 0) 'RED
    .StopIfTrue = False
    End With

    With Range("J2:J" & Table_row - 1).FormatConditions.Add(xlCellValue, xlGreater, "0")
    .Interior.Color = RGB(0, 255, 0) 'Green
    .StopIfTrue = False
    End With
   
'Bonus print
'"Greatest % Increase"
    Range("P2").Value = maxPercIncTicker
    Range("Q2").Value = maxPercIncrease
    Range("Q2").NumberFormat = "0.00%"     'Style = "Percent"
   
'"Greatest % Decrease"
    Range("P3").Value = maxPercDecTicker
    Range("Q3").Value = maxPercDecrease
    Range("Q3").NumberFormat = "0.00%"     'Style = "Percent"
    
'"Greatest Total Volume"
    Range("P4").Value = maxVolTicker
    Range("Q4").Value = maxVolume
    
    Range("A:Q").Columns.AutoFit 'Make it beautiful to read
    
Next ws

End Sub






