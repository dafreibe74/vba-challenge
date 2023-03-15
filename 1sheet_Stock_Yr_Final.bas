Attribute VB_Name = "Module2"
Sub Stock_Yr_Final()

'set variable to hold the ticker symbol
Dim Ticker_Symbol As String
'set variable to hold the total per stock ticker
Dim Stock_Volume As Double
'set variable to hold the yearly change
Dim Yearly_Change As Double
'set variable to grab the open price value in beginning
Dim Open_Def As Boolean
'set variable to ensure last row works for longer sheets
Dim last_row As Long
Dim i As Long
'set variables for the min and max, both value and symbol
Dim Min_percent As Double
Dim Max_percent As Double
Dim Max_Volume As Double
Dim Min_percent_symbol As String
Dim Max_percent_symbol As String
Dim Max_Volume_symbol As String

'set variable to track location of each ticker symbol in the summary table
Dim tally As Integer
tally = 2

'set variable to grab open price at beginning
Open_Def = True

'create column headers
    Range("I1, P1").Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    

'automate the loop to run to last row
last_row = Cells(Rows.Count, 1).End(xlUp).Row

'loop through all the stocks
For i = 2 To last_row

                'check if we are still within the the same year value, if not. . .
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                    'set the stock name
                    Range("I" & tally).Value = Cells(i, 1).Value
                    
                    'determine the yearly change and place in summary table
                    Close_amt = Cells(i, 6).Value
                    Yearly_Change = Close_amt - open_amt
                    Range("J" & tally).Value = Yearly_Change
                    
                        'conditionally format the two change columns
                        If Yearly_Change < 0 Then
                            Range("J" & tally).Interior.ColorIndex = 3
                            Range("K" & tally).Interior.ColorIndex = 3
                         
                         ElseIf Yearly_Change > 0 Then
                            Range("J" & tally).Interior.ColorIndex = 4
                            Range("K" & tally).Interior.ColorIndex = 4
                        End If
                    
                        'determine the percent change and place in summary table - compensate for zero
                        If Yearly_Change = 0 Or open_amt = 0 Then
                            Range("K" & tally).Value = 0
                        Else
                            Range("K" & tally).Value = Format(Yearly_Change / open_amt, "#.##%")
                        End If
                        
                        'add to the ticker volume
                        Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                        'print the stock volume in the summary table
                        Range("L" & tally).Value = Stock_Volume
                            
                       'get values for the highest and lowest
                        If Cells(tally, 11).Value > Max_percent Then
                            If Cells(tally, 11).Value = ".%" Then
                            Else
                                Max_percent = Cells(tally, 11).Value
                                Max_percent_symbol = Cells(tally, 9).Value
                            End If
                            ElseIf Cells(tally, 11).Value < Min_percent Then
                                Min_percent = Cells(tally, 11).Value
                                Min_percent_symbol = Cells(tally, 9).Value
                            ElseIf Cells(tally, 12).Value > Max_Volume Then
                                Max_Volume = Cells(tally, 12).Value
                                Max_Volume_symbol = Cells(tally, 9).Value
                            End If
                        
                    'advance down in summary table row
                    tally = tally + 1
                    
                    'Reset the volume total
                    Stock_Volume = 0
                    
                    'add to the ticker volume
                    Stock_Volume = Stock_Volume + Cells(i, 7).Value
                    
                    'print the stock volume in the summary table
                    Range("L" & tally).Value = Stock_Volume
                    
                Else
                'Use default variable to price value at the beginning.
                 If Open_Def Then
                    open_amt = Cells(i, 3).Value
                    Open_Def = False
                End If
                
                ' If next ticker hasn't changed, bank the volume
                Stock_Volume = Stock_Volume + Cells(i, 7).Value
        
        End If

Next i
                    
 ' Print symbol to summary table
    Cells(2, 16).Value = Max_percent_symbol
    Cells(3, 16).Value = Min_percent_symbol
    Cells(4, 16).Value = Max_Volume_symbol

' Add values for greatest decrease/increase and greatest volume.
    Cells(2, 17).Value = Format(Max_percent, "#.##%")
    Cells(3, 17).Value = Format(Min_percent, "#.##%")
    Cells(4, 17).Value = Max_Volume
    
End Sub

