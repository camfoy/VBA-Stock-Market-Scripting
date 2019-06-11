Sub Stocks_Easy

    For each ws in Worksheets

        Dim Ticker as String
        
        Dim Volume as Double
            Volume = 0

        Dim STR as Integer
            STR = 2

        Dim LastRow as Long
            LastRow = ws.Cells(Rows.Count, 1).End(xlUP).Row

        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Total Stock Volume"

        For i = 2 to LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

                Ticker = ws.Cells(i, 1).Value
                Volume = Volume + ws.Cells(i, 7).Value
                ws.Range("I" & STR).Value = Ticker
                ws.Range("J" & STR).Value = Volume
                STR = STR + 1
                Volume = 0

            Else

                Volume = Volume + ws.Cells(i, 7).Value

            End If

        Next i

        ws.Cells.Columns.Autofit

    Next ws

End Sub