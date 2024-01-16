Attribute VB_Name = "Module2"

Sub stocks()

  Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
         
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "percent change"
    [l1] = " Total Stock Volume"
    [O2] = "Greatest%Increase"
    [O3] = "Greatest%Decrease"
    [o4] = "Greatest Total Volume"
    [P1] = "Ticker"
    [q1] = "Value"
    
    Columns.AutoFit
    
    si = 2
    openPrice = 0
    Dim volume As Double
    volume = 0
    
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    
    For I = 2 To LastRow
'Begining of the ticker year
        If openPrice = 0 Then
            openPrice = ws.Cells(I, "C")
        End If
        
'End of the ticker year
        If ws.Cells(I, "A") <> ws.Cells(I + 1, "A") Then
            ws.Cells(si, "I") = ws.Cells(I, "A")
            yearly_ch = ws.Cells(I, "F") - openPrice
            ws.Cells(si, "J") = yearly_ch
            If yearly_ch > 0 Then
                ws.Cells(si, "J").Interior.ColorIndex = 4
            Else
                ws.Cells(si, "J").Interior.ColorIndex = 3
            End If
            
    'percentige change
    
    precentigachange = yearly_ch / openPrice
     ws.Cells(si, "k") = precentigachange
     ws.Cells(si, "k").NumberFormat = "0.00%"
     
      volume = volume + ws.Cells(I, "G")
      
      ws.Cells(si, "l") = volume
    
'reset area
        si = si + 1
        
        openPrice = ws.Cells(I + 1, "C")
        
        Else
        
    volume = volume + ws.Cells(I, "G")
        
        End If
    
    Next I
    
    
    ws.Cells(2, "Q").Value = WorksheetFunction.Max(ws.Range("k2:k" & LastRow))
    
      ws.Cells(2, "Q").NumberFormat = "0.00%"
     ws.Cells(3, "Q").Value = WorksheetFunction.Min(ws.Range("k2:k" & LastRow))
     
       ws.Cells(3, "Q").NumberFormat = "0.00%"
     
      ws.Cells(4, "Q").Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
      
      max_index = WorksheetFunction.Match(ws.Cells(2, "Q").Value, ws.Range("k2:k" & LastRow), 0)
      
      min_index = WorksheetFunction.Match(ws.Cells(3, "Q").Value, ws.Range("k2:k" & LastRow), 0)
      
      vol_index = WorksheetFunction.Match(ws.Cells(4, "Q").Value, ws.Range("L2:L" & LastRow), 0)
      
      ws.Cells(2, "P") = ws.Cells(max_index + 1, "I")
      
       ws.Cells(3, "P") = ws.Cells(min_index + 1, "I")
       
        ws.Cells(4, "P") = ws.Cells(vol_index + 1, "I")
      
      
       
    Next ws
    
    
End Sub

