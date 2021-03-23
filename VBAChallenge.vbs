VBAChallenge.vbs

Sub Prueba_1()

For Each ws In Worksheets

    Dim ticker as String

    Dim Yearly_change as Double
    Yearly_change = 0

    Dim Total_Stock as Double
    Total_Stock = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    Dim LastRow as Long
    With ActiveSheet.UsedRange
        LastRow = .Rows(.Rows.Count).Row
    End With

    Dim Percent_change as Double
    Dim Open_value As Double
    Dim Close_value As Double
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Total Stock Volume counter:

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                Total_Stock = Total_Stock + ws.Cells(i, 7).Value

                ws.Range("L" & Summary_Table_Row).Value = Total_Stock
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Total_Stock = 0

            Else

                Total_Stock = Total_Stock + ws.Cells(i, 7).Value

            End If

        Next i

    'Summary Table row reset:

    Summary_Table_Row = 2

    'Ticker symbol, yearly change from opening price to the closing price, percent change

        For i = 2 To LastRow    
            Open_value = ws.Cells(i, 3).Value

            ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = ticker


            While ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value

                i = i + 1
                
            Wend 
            

            Close_value = ws.Cells(i, 6).Value
         

            Yearly_change = Close_value - Open_value

            '*Prints the annual oppening and close in J and K

            'Range("J" & Summary_Table_Row).Value = Open_value

            'Range("K" & Summary_Table_Row).Value = Close_value

            Yearly_change = Close_value - Open_value

            ws.Range("J" & Summary_Table_Row).Value = Yearly_change

            If Open_value = 0 Then

                Percent_change = 0

            Else
                Percent_change = (Close_value - Open_value ) / Open_value

            End If
            
            'Percent_change = ((Close_value * 100) / Open_value) - 100
            ws.Range("K" & Summary_Table_Row).Value = Percent_change
            ws.Range("K:K").NumberFormat = "0.00%"

            
            Summary_Table_Row = Summary_Table_Row + 1

            

        Next i

    'Conditional formatting:
    
    Dim cond1 As FormatCondition
    Dim cond2 As FormatCondition

    Dim Myrange As Range
    Set Myrange = ws.Range("J2", ws.Range("J2").End(xlDown))

    Set cond1 = Myrange.FormatConditions.Add(xlCellValue, xlGreater, "0")
    Set cond2 = Myrange.FormatConditions.Add(xlCellValue, xlLess, "0")

    With cond1
        .Interior.Color = vbGreen
    End With

    With cond2
        .Interior.Color = vbRed
    End With

    'Bonus:

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Range("O2").Value = "Greatest Increase %"
    ws.Range("O3").Value = "Greatest Decrease %"
    ws.Range("O4").Value = "Gretest Total Volume"
    
    Dim Max_inc as Double
    Dim Min_inc as Double
    Dim Max_vol as Double

    Max_inc = Application.worksheetfunction.max(ws.range("J:J"))
    ws.cells(2, 17).Value = Max_inc

    Min_inc = Application.worksheetfunction.min(ws.range("J:J"))
    ws.cells(3, 17).Value = Min_inc

    Max_vol = Application.worksheetfunction.max(ws.range("L:L"))
    ws.cells(4, 17).Value = Max_vol

Next ws

End Sub   

        