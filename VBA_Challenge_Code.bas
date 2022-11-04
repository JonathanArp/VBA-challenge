Sub Module2()

Dim Ticker As String

Dim Stock_Volume As Double
Stock_Volume = 0

Dim Yearly_Change As Single

Dim Percent_Change As Single

Dim Start_Amount As Single

Dim End_Amount As Single

Dim Row_Number As Integer
Row_Number = 2

Range("I1").Value = "Ticker"

Range("J1").Value = "Yearly_Change"

Range("K1").Value = "Percent_Change"

Range("L1").Value = "Total Stock Volume"

For i = 2 To 753001

If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    Start_Amount = Cells(i, 3).Value

End If

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    Ticker = Cells(i, 1).Value
    
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
    Range("I" & Row_Number).Value = Ticker
    
    Range("L" & Row_Number).Value = Stock_Volume
    
    End_Amount = Cells(i, 6).Value
    
    Yearly_Change = End_Amount - Start_Amount
    
    Range("J" & Row_Number).Value = Yearly_Change

    Percent_Change = Yearly_Change / Start_Amount
    
    Range("K" & Row_Number).Value = Percent_Change
    
    Row_Number = Row_Number + 1
    
    Stock_Volume = 0
    
Else
    Stock_Volume = Stock_Volume + Cells(i, 7).Value
    
End If

If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4

ElseIf Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    

End If


Next i

    
End Sub

