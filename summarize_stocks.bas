Attribute VB_Name = "Module1"
Sub SummaryData()

    For Each ws In Worksheets
        Dim lr As Long
        lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        Dim Ticker As String
        Dim OV As Currency
        Dim CV As Currency
        Dim Vol As LongLong
        
        Ticker = ws.Cells(2, 1).Value
        OV = ws.Cells(2, 3).Value
        CV = ws.Cells(2, 6).Value
        Vol = ws.Cells(2, 7).Value
        
        Dim nr As Long
  
        For r = 3 To lr + 1
            
            If ws.Cells(r, 1).Value = Ticker Then
                CV = ws.Cells(r, 6).Value
                Vol = Vol + ws.Cells(r, 7).Value
            Else
                nr = ws.Cells(Rows.Count, 9).End(xlUp).Row + 1
                ws.Cells(nr, 9).Value = Ticker
                'conditional format green/red
                If CV - OV > 0 Then
                    ws.Cells(nr, 10).Interior.ColorIndex = 4
                    ws.Cells(nr, 10).Value = CV - OV
                ElseIf CV - OV < 0 Then
                    ws.Cells(nr, 10).Interior.ColorIndex = 3
                    ws.Cells(nr, 10).Value = CV - OV
                End If
                ws.Cells(nr, 12).Value = Vol
                'avoid divide by zero error
                If OV = 0 Then
                    ws.Cells(nr, 11).Value = Null
                Else
                    ws.Cells(nr, 11).Value = (CV - OV) / OV
                End If
                
                Ticker = ws.Cells(r, 1).Value
                OV = ws.Cells(r, 3).Value
                CV = ws.Cells(r, 6).Value
                Vol = ws.Cells(r, 7).Value
            End If
        Next r

        'format ranges
        Dim fl As Long
        fl = ws.Cells(Rows.Count, 9).End(xlUp).Row
        ws.Range("J2:J" & fl).NumberFormat = "$#,##0.00; -$#,##0.00"
        ws.Range("K2:K" & fl).NumberFormat = "0.00%; -0.00%"
        ws.Range("L2:L" & fl).NumberFormat = "#,##0"
    
    Next ws

End Sub
