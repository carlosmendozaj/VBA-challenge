VBAChallenge.vbs

Sub Prueba_1()

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
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'Total Stock Volume counter:

        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                Total_Stock = Total_Stock + Cells(i, 7).Value

                Range("L" & Summary_Table_Row).Value = Total_Stock
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Total_Stock = 0

            Else

                Total_Stock = Total_Stock + Cells(i, 7).Value

            End If

        Next i

    'Summary Table row reset:

    Summary_Table_Row = 2

    'Ticker symbol, yearly change from opening price to the closing price, percent change

        For i = 2 To LastRow    
            Open_value = Cells(i, 3).Value

            ticker = Cells(i, 1).Value
            Range("I" & Summary_Table_Row).Value = ticker


            While Cells(i, 1).Value = Cells(i + 1, 1).Value

                i = i + 1
                
            Wend 
            

            Close_value = Cells(i, 6).Value
         

            Yearly_change = Close_value - Open_value

            '*Prints the annual oppening and close in J and K

            'Range("J" & Summary_Table_Row).Value = Open_value

            'Range("K" & Summary_Table_Row).Value = Close_value

            Yearly_change = Close_value - Open_value

            Range("J" & Summary_Table_Row).Value = Yearly_change

            Percent_change = ((Close_value * 100) / Open_value) - 100
            Range("K" & Summary_Table_Row).Value = Percent_change

            
            Summary_Table_Row = Summary_Table_Row + 1

            

        Next i

    'Conditional formatting:
    
    Dim cond1 As FormatCondition
    Dim cond2 As FormatCondition

    Dim Myrange As Range
    Set Myrange = Range("J2", Range("J2").End(xlDown))

    Set cond1 = Myrange.FormatConditions.Add(xlCellValue, xlGreater, "0")
    Set cond2 = Myrange.FormatConditions.Add(xlCellValue, xlLess, "0")

    With cond1
        .Interior.Color = vbGreen
    End With

    With cond2
        .Interior.Color = vbRed
    End With


End Sub   

        