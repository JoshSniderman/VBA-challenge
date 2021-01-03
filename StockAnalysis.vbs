Sub TestData()

    Dim i As Long
    Dim ttl As LongLong
    Dim ticker As String
    Dim strt As Double
    Dim nd As Double
    Dim agg As Integer
    
    ttl = 0
    agg = 2
    strt = Range("C2").Value
    
    For i = 2 To 800000
        ttl = ttl + Cells(i, 7).Value
        If Cells(i, 1) <> Cells(i + 1, 1) Then
            ticker = Cells(i, 1).Value
            nd = Cells(i, 6).Value
            Cells(agg, 9).Value = ticker
            Cells(agg, 10).Value = nd - strt
            If strt = 0 Then
                Cells(agg, 11).Value = Format(0, "Percent")
            Else
                Cells(agg, 11).Value = Format((nd - strt) / strt, "Percent")
            End If
            Cells(agg, 12).Value = ttl
            If Cells(agg, 10).Value >= 0 Then
                Cells(agg, 10).Interior.ColorIndex = 4
            Else
                Cells(agg, 10).Interior.ColorIndex = 3
            End If
            agg = agg + 1
            ttl = 0
            strt = Cells(i + 1, 3).Value
        End If
    Next i

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("I1:L1").Font.Bold = True
End Sub