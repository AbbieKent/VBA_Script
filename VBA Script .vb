Sub VBAHomework()
    For Each ws In Worksheets
        ws.Activate
        Call SetTitle
    Next ws
End Sub
Sub SetTitle()
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0
' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    'this is for challenge only
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("I:O").Columns.AutoFit
    
Call CalculateSummary
End Sub

Sub CalculateSummary()
    ' Start writing your code here
    'declare variables
Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim yearly As Double
Dim percentchange As Double
Dim totalvolume As Double
Dim j As Double
open_price = Cells(2, 3).Value
close_price = 0
yearly = 0
percent_change = 0
total_volume = 0
ticker = ""

'find last row
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row

'ticker loop
j = 2
For i = 2 To last_row + 1
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        
        'total volume
            total_volume = total_volume + Cells(i, 7).Value
            ticker = Cells(i, 1).Value
            
    Else
        'yearly
            close_price = Cells(i, 6).Value
            yearly = close_price - open_price
            Cells(j, 10).Value = yearly
        'percent change
            If close_price = 0 And open_price <> 0 Then
                percent_change = -100
            ElseIf close_price = 0 And open_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly / open_price)
            End If
            Cells(j, 11).Value = percent_change
            Cells(j, "K").NumberFormat = "00.00%'"
            Cells(j, 12).Value = total_volume
            'add ticker
            Cells(j, "I").Value = ticker
    
'format pos neg
    If Cells(j, "j") > 0 Then
       Cells(j, "j").Interior.Color = vbGreen
    ElseIf Cells(j, "j") < 0 Then
       Cells(j, "j").Interior.Color = vbRed
    Else
       Cells(j, "j").Interior.Color = vbBlue
    End If
    
    If Cells(j, "k").Value > 0 Then
       Cells(j, "k").Interior.Color = vbGreen
    ElseIf Cells(j, "k") < 0 Then
       Cells(j, "k").Interior.Color = vbRed
    Else
       Cells(j, "k").Interior.Color = vbBlue
    End If
open_price = Cells(i + 1, 3).Value
close_price = 0
yearly = 0
percent_change = 0
j = j + 1
End If

Next i
'challenge
last_row_chal = Cells(Rows.Count, 9).End(xlUp).Row
Dim increase As Double
Dim increase_ticker As String
Dim decrease As Double
Dim decrease_ticker As String
Dim great_total As Double
Dim total_ticker As String
increase = Cells(2, 10).Value
decrease = Cells(2, 10).Value
great_total = Cells(2, 12).Value
For i = 2 To last_row_chal

 If increase < Cells(i, 10).Value Then
     increase = Cells(i, 10).Value
  increase_ticker = Cells(i, 9).Value

End If
   If decrease > Cells(i, 10).Value Then
      decrease = Cells(i, 10).Value
     decrease_ticker = Cells(i, 9).Value
    End If
    
   If great_total < Cells(i, 12).Value Then
       great_total = Cells(i, 12).Value
       total_ticker = Cells(i, 9).Value
    End If
Next i
Cells(2, 16).Value = increase_ticker
Cells(2, 17).Value = increase
Cells(3, 16).Value = decrease_ticker
Cells(3, 17).Value = decrease
Cells(4, 16).Value = total_ticker
Cells(4, 17).Value = great_total
End Sub





