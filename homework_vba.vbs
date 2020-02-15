Sub first_try():



Dim ws As Worksheet
Dim LastRow As Double
Dim ticker As String


For Each ws In Worksheets


    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Price Change"
    Cells(1, 11).Value = "Price Change Percentage"
    Cells(1, 12).Value = "Total Stock Volume"


    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Dim i As Long
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        

        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ticker = Cells(i, 1).Value
            
            Range("I" & Summary_Table_Row).Value = ticker
            
            
            Summary_Table_Row = Summary_Table_Row + 1
    
            End If
            
        Next i
        
    
    



Next ws


End Sub