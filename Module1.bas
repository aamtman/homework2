Attribute VB_Name = "Module1"
Sub nubmers():

For Each ws In Worksheets
 'ws.Activate
 
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Double

'run through each worksheet

    'set headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'setup integers for loop
    Summary_Table_Row = 2
    
    'determine the last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'loop
        ' Account for first recrod / row and set year_open
        
        year_open = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
            vol = ws.Cells(i, 7).Value + vol
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' we have year end now read ticker and calculate the changes
                
                ticker = ws.Cells(i, 1).Value
                ws.Cells(Summary_Table_Row, 9).Value = ticker
                
                year_close = ws.Cells(i, 6).Value
                yearly_change = year_close - year_open
              
                ' Account for zero cases
                
             If year_open = 0 Then
                percent_change = year_close
             Else
                percent_change = (year_close - year_open) / year_open
             End If
                 
                    ' insert values into summary for ticker
                    ws.Cells(Summary_Table_Row, 10).Value = yearly_change
                    ws.Cells(Summary_Table_Row, 11).Value = percent_change
                    ws.Cells(Summary_Table_Row, 12).Value = vol
                    ' ws.Cells(Summary_Table_Row, 13).Value = I
                    If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
                        ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(0, 255, 0)
                    Else
                        ws.Cells(Summary_Table_Row, 10).Interior.Color = RGB(255, 0, 0)
                End If
                    
                
                ' Always set the next year_open when we find a new ticker and zero the volume
                year_open = ws.Cells(i + 1, 3).Value
                vol = 0
                
                ' Iterate summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
            End If
                 
            
              
            'finish loop
            Next i
    
' ws.Columns("K").NumberFormat = "0.00%"




'move through to next worksheet

Next ws

End Sub
