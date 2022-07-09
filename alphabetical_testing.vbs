Attribute VB_Name = "Module1"
Sub alphabetical_testing()

    For Each ws In Worksheets

    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    Dim Ticker As String
    Dim Volume As Double
    Dim Counter As Integer
    Dim Index As Integer
    
    
    Volume = 0
    Counter = 0
    Index = 2
    
    Dim ColorRed As Integer
    Dim ColorGreen As Integer
    Dim StartofYear As Long
    
    ColorRed = 3
    ColorGreen = 4
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(Index, 11).NumberFormat = "0.00%"

    StartofYear = 2
    
    
        For i = 2 To LastRow
        
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(Index, 9).Value = Ticker
                
                Volume = Volume + ws.Cells(i, 7).Value
                ws.Cells(Index, 12).Value = Volume
                
                ws.Cells(Index, 10).Value = ws.Cells(i, 6).Value - ws.Cells(StartofYear, 3).Value
                ws.Cells(Index, 11).Value = (ws.Cells(i, 6).Value / ws.Cells(StartofYear, 3).Value) - 1
                
            
                
                    If ws.Cells(Index, 10).Value >= 0 Then
                    
                        ws.Cells(Index, 10).Interior.ColorIndex = ColorGreen
                        
                    Else
                        
                        ws.Cells(Index, 10).Interior.ColorIndex = ColorRed
                        
                    End If
                    
                Index = Index + 1
                StartofYear = i + 1
                
                
            Else
                
                Counnter = Counter + 1
                Volume = Volume + ws.Cells(i, 7).Value
            End If
            
        Next i
        
    Next ws


End Sub
