VBAChallenge.vbs

Sub Prueba_1()

    Dim ticker as String

    Dim Yearly_change as Double
    Yearly_change = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    Dim LastRow as Long
    With ActiveSheet.UsedRange
        LastRow = .Rows(.Rows.Count).Row
    End With

    Dim Open_value As Double

    Dim Close_value As Double
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"

        For i = 2 To LastRow


            Open_value = Cells(i, 3).Value

            ticker = Cells(i, 1).Value
            Range("I" & Summary_Table_Row).Value = ticker


            While Cells(i, 1).Value = Cells(i + 1, 1).Value

                i = i + 1
                
            Wend 
            

            Close_value = Cells(i, 6).Value
         

            Yearly_change = Close_value - Open_value

            'Prints the annual oppening and close in J and K

            'Range("J" & Summary_Table_Row).Value = Open_value

            'Range("K" & Summary_Table_Row).Value = Close_value

            Yearly_change = Close_value - Open_value

            Range("J" & Summary_Table_Row).Value = Yearly_change

            
            Summary_Table_Row = Summary_Table_Row + 1

            

        Next i

End Sub  
