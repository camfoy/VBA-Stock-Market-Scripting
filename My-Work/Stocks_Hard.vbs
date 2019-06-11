Sub Stocks_Hard

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

        Dim Max_Percent As Double

        Dim Min_Percent as Double

        Dim Max_Vol as Double

        Max_Percent = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))

        Min_Percent = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))

        Max_Vol = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q2").Value = Max_Percent

        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Min_Percent

        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("Q4").Value = Max_Vol

        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        Dim Ticker_2 As String

        For i = 2 to LastRow

            If ws.Cells(i, 11).Value = Max_Percent Then

                Ticker_2 = ws.Cells(i, 9).Value
                ws.Range("P2").Value = Ticker_2

            End If

            If ws.Cells(i,11).Value = Min_Percent Then

                Ticker_2 = ws.Cells(i, 9).Value
                ws.Range("P3").Value = Ticker_2

            End If

            If ws.Cells(i, 12).Value = Max_Vol Then

                Ticker_2 = ws.Cells(i, 9).Value
                ws.Range("P4").Value = Ticker_2

            End If

        Next i

        ws.Cells.Columns.Autofit

    Next ws

End Sub