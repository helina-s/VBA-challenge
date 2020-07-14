Attribute VB_Name = "Module1"
Sub Stockyear()

Dim ticker As Variant
Dim opening_price As Double
    opening_price = 0
Dim closing_price As Double
    closing_price = 0
Dim total_volume As Long
    total_volume = 0
Dim yearly_change As Double
Dim percent_change As Variant

Dim greatest_percent_increase As Double
Dim greatest_percent_decrease As Double
Dim geatest_total_volume As Double

Dim lastrow As Variant
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
       
Dim summary_table As Long
    summary_table = 2

For Row = 2 To lastrow

If Cells(Row, 1).Value <> Cells(Row + 1, 1).Value Then
    
    ticker = Cells(Row, 1).Value
    Range("I" & summary_table).Value = ticker
    
    closing_price = Cells(Row, 6).Value
    opening_price = Cells(Row, 3).Value
    
    yearly_change = closing_price - opening_price
    Range("J" & summary_table).Value = yearly_change
    

    percent_change = (yearly_change / opening_price) * 100
    Range("K" & summary_table).Value = percent_change
        'format the percent change
    percent_change = Format(percent_change, "%0.00")

    total_volume = total_volume + Cells(Row, 7).Value
    Range("L" & summary_table).Value = total_volume
    
    summary_table = summary_table + 1
    
    End If
    
        'positive and negative yearlychange
    If Range("J" & summary_table).Value > 0 Then
        Range("J" & summary_table).Interior.ColorIndex = 4
    
    ElseIf Range("J" & summary_table).Value < 0 Then
        Range("J" & summary_table).Interior.ColorIndex = 3

End If
Next Row

    'calculate the greatest % increase, decrease and total volume
For Row = 2 To lastrow
 
If Cells(Row + 1, 11).Value > greatest_percent_increase Then
    greatest_percent_increase = Cells(Row + 1, 11).Value
    Range("O2").Value = Cells(Row + 1, 9).Value
    Range("P2") = greatest_percent_increase
End If
If Cells(Row + 1, 11).Value < greatest_percent_decrease Then
    greatest_percent_decrease = Cells(Row + 1, 11).Value
    Range("O3").Value = Cells(Row + 1, 9).Value
    Range("P3") = greatest_percent_decrease
End If

If Cells(Row + 1, 12).Value > greatest_total_volume Then
    greatest_total_volume = Cells(Row + 1, 12).Value
    greatest_total_volume = Cells(Row + 1, 12).Value
    Range("P4") = greatest_total_volume
    Range("O4").Value = Cells(Row + 1, 9).Value
End If

Next Row

End Sub
