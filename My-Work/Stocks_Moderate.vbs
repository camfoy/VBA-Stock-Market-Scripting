Sub Stocks_Moderate

    For each ws in Worksheets

        Dim Ticker as String
        
        Dim Volume as Double
            Volume = 0

        Dim STR as Integer
            STR = 2

        Dim LastRow as Long
            LastRow = ws.Cells(Rows.Count, 1).End(xlUP).Row

        Dim Yearly_Change as Double

        Dim Opn as Double
            Opn = ws.Cells(2, 3).Value

        Dim Clos as Double

        Dim Percent_Change as Double

        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly Change"

        ws.Range("K1").Value = "Percent Change"

        ws.Range("L1").Value = "Total Stock Volume"

        For i = 2 to LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

                Ticker = ws.Cells(i, 1).Value
                Clos = ws.Cells(i, 6).Value
                Yearly_Change = Clos - Opn

                If Opn > 0 Then

                    Percent_Change = (Clos - Opn)/Opn

                Else

                    Percent_Change = 0

                End if

                Volume = Volume + ws.Cells(i, 7).Value
                ws.Range("I" & STR).Value = Ticker
                ws.Range("J" & STR).Value = Yearly_Change
                
                    If Yearly_Change >= 0 Then
                
                        ws.Range("J" & STR).Interior.Color = vbGreen

                    Else

                        ws.Range("J" & STR).Interior.Color = vbRed

                    End if

                ws.Range("K" & STR).Value = Percent_Change
                ws.Range("K" & STR).NumberFormat = "0.00%"
                ws.Range("L" & STR).Value = Volume
                STR = STR + 1
                Yearly_Change = 0
                Percent_Change = 0
                Opn = ws.Cells(i + 1, 3).Value
                Volume = 0

            Else

                Volume = Volume + ws.Cells(i, 7).Value

            End If

        Next i

         ws.Cells.Columns.Autofit

    Next ws

End Sub