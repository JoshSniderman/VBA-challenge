Sub MaxMinBonus()

    Dim max As Double
    Dim min As Double
    Dim tot As LongLong
    Dim j As Long
    
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    max = Range("K2").Value
    min = Range("K2").Value
    tot = Range("L2").Value
    
    For j = 2 To 4000
        If min > Cells(j, 11) Then
            min = Cells(j, 11)
            Range("P3") = Cells(j, 9)
            Range("Q3") = Format(Cells(j, 11), "Percent")
        End If
        If max < Cells(j, 11) Then
            max = Cells(j, 11)
            Range("P2") = Cells(j, 9)
            Range("Q2") = Format(Cells(j, 11), "Percent")
        End If
        If tot < Cells(j, 12).Value Then
            tot = Cells(j, 12).Value
            Range("P4").Value = Cells(j, 9)
            Range("Q4").Value = Cells(j, 12)
        End If
    Next j
End Sub