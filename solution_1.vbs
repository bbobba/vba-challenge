Sub vba_stock_analysis():

    ' set dimensions to double data type
    Dim total As Double

    ' find how many rows contain data
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    ' set header labels for results
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Total Stock Volume"

    For i = 2 To RowCount

        ' if the ticker value changes, then print the results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' store results
            total = total + Cells(i, 7).Value

            ' print ticker
            Range("I" & 2 + j).Value = Cells(i, 1).Value

            ' print total
            Range("J" & 2 + j).Value = total

            ' reset total
            total = 0

            ' continue to next row
            j = j + 1

        ' else add to the total
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

End Sub