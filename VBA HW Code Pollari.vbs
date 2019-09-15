Sub StockVolume()
  
  ' variables
    Dim Ticker As String
    Dim TotalVolume As Double
    Dim NextTick As String
    Dim i As Double
    Dim UniqueTick As Double
    Dim LastRow As Double
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    TotalVolume = 0
    UniqueTick = 1
    
    ' For Loops
    
            For i = 1 To LastRow
        
            Ticker = Cells(i, 1).Value
            NextTick = Cells(i + 1, 1).Value
            
            If Ticker = NextTick Then
            
                TotalVolume = TotalVolume + Cells(i + 1, 7).Value
            
                ElseIf Ticker <> NextTick Then
                Cells(UniqueTick, 10).Value = TotalVolume
                Cells(UniqueTick, 9).Value = Ticker
            
            UniqueTick = UniqueTick + 1
            
                TotalVolume = Cells(i + 1, 7).Value
                   
                End If
          
        Next i
        

Cells(1, 10).Value = "Total Stock Volume"
Cells(1, 9).Value = "Ticker Symbol"

End Sub

' Running the script for each sheet so don't have to run it multiple times
Sub StockWS()
    
    Dim WS As Worksheet
    For Each WS In Worksheets
    WS.Activate
    
    Call StockVolume
    Next WS
    
End Sub
