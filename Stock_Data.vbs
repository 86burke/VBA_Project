Sub stock()
'Declare everything
    Dim ws As Worksheet
    Dim Ticker As String
    Dim vol As Double
    Dim po As Double
    Dim pc As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Summary_Table_Row As Double
'For every worksheet
    For Each ws In ThisWorkbook.Worksheets
'Summary table start
    Summary_Table_Row = 2
'New column names
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

'lastrow
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'loop for volume
    For i = 2 To lastrow
    po = ws.Cells(i, 3).Value
    vol = vol + Cells(i, 7).Value
   
           If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
           Ticker = ws.Cells(i, 1).Value
           pc = ws.Cells(i, 6).Value
           
           yearly_change = pc - po
           'percent_change = (pc - po) / pc
           
           ws.Range("I" & Summary_Table_Row).Value = Ticker
           ws.Range("J" & Summary_Table_Row).Value = yearly_change
           ws.Range("K" & Summary_Table_Row).Value = percent_change
           ws.Range("L" & Summary_Table_Row).Value = vol
           
           Summary_Table_Row = Summary_Table_Row + 1
           vol = 0
           
           ElseIf po <> 0 Then
                    'Dim pc As Double
                    percent_change = (pc - po) / po
                    ws.Range("K" & Summary_Table_Row) = Format(percent_change, "percent")
                    Else
                    Cells(2, 11).Value = Format(0, "percent")
                   
        End If
    Next i

    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = ws.Range("J2", ws.Range("J2").End(xlDown))

    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.ColorIndex = 4
            End With
        Case Is < 0
            With color_cell
                .Interior.ColorIndex = 3
            End With
       End Select
    Next g
        
Next ws
End Sub
