Attribute VB_Name = "Module1"
Sub Module_2_Challenge()

    Dim ws As Worksheet
    Dim lastrow As Long
    Dim summary_row As Integer
    Dim ticker As String
    Dim stock_volume As Double
    Dim opening As Double
    Dim closing As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
        
        ws.Range("P1").Value = "Ticker"
        
        ws.Range("Q1").Value = "Value"
           
        ws.Range("I1").Value = "Ticker"
        
        ws.Range("J1").Value = "Yearly Change"
        
        ws.Range("K1").Value = "Percent Change"
        
        ws.Range("L1").Value = "Total Stock Volume"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        summary_row = 2
        
        stock_volume = 0
        
        
        For i = 2 To lastrow
        
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            
                opening = ws.Cells(i, 3).Value
                
            End If
            
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value
                
                closing = ws.Cells(i, 6).Value
                
                yearly_change = closing - opening
                
                    If opening <> 0 Then
                    
                        percent_change = yearly_change / opening
                    
                    Else
                        
                        percent_change = 0
                        
                    End If
                    
                    
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                ws.Cells(summary_row, 9).Value = ticker
                
                ws.Cells(summary_row, 10).Value = yearly_change
                
                    If yearly_change >= 0 Then
                    
                        ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                        
                    Else
                        
                        ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                    
                    End If
                        
                    
                
                ws.Cells(summary_row, 11).Value = percent_change
                
                ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
                
                
                ws.Cells(summary_row, 12).Value = stock_volume
                
                summary_row = summary_row + 1
                
                stock_volume = 0
                
                
            Else
                
                stock_volume = stock_volume + ws.Cells(i, 7).Value
                
                
            End If
            
        Next i
        
        Dim lastrow_ticker As Long
        
        ws.Range("O2").Value = "Greatest % Increase"
        
        ws.Range("O3").Value = "Greatest % Decrease"
        
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ws.Range("P1").Value = "Ticker"
        
        ws.Range("Q1").Value = "Value"
        
        lastrow_ticker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
            For i = 2 To lastrow_ticker
            
                If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow_ticker)) Then
                    
                    ws.Range("Q2").Value = ws.Cells(i, 11).Value
                    
                    ws.Range("P2").Value = ws.Cells(i, 9).Value
                    
                ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow_ticker)) Then
                    
                    ws.Range("Q3").Value = ws.Cells(i, 11).Value
                    
                    ws.Range("P3").Value = ws.Cells(i, 9).Value
                
                ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow_ticker)) Then
                
                    ws.Range("Q4").Value = ws.Cells(i, 12).Value
                    
                    ws.Range("P4").Value = ws.Cells(i, 9).Value
                
                ws.Range("Q2:Q3").NumberFormat = "0.00%"
                
                
                End If
                
            Next i
            
    Next ws
    
    
End Sub


