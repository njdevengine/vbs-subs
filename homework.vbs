Sub output()
For Each ws In Worksheets
Dim end_row, new_end, opening, closing, total_volume, year_change, percent_change As Double
Dim Summary_Table_Row As Integer
Dim zero As String
zero = "0%"
end_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
total_volume = 0
Summary_Table_Row = 2

ws.Range("I1").Value = "ticker"
ws.Range("J1").Value = "year_change"
ws.Range("K1").Value = "%_change"
ws.Range("L1").Value = "total_volume"

For i = 2 To end_row

If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
opening = ws.Cells(i, 3).Value
End If

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
total_volume = total_volume + ws.Cells(i, 7).Value
closing = ws.Cells(i, 6).Value
year_change = closing - opening

If opening = 0 Then
ws.Cells(Summary_Table_Row, 11).Value = zero
Else
percent_change = (closing - opening) / opening
ws.Cells(Summary_Table_Row, 11).Value = percent_change
ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
End If

ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i, 1).Value
ws.Cells(Summary_Table_Row, 10).Value = year_change
ws.Cells(Summary_Table_Row, 12).Value = total_volume
Summary_Table_Row = Summary_Table_Row + 1
total_volume = 0
Else
total_volume = total_volume + ws.Cells(i, 7)
End If
Next i

new_end = ws.Cells(Rows.Count, 9).End(xlUp).Row
For i = 2 To new_end
If ws.Cells(i, 10) >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 43
Else
ws.Cells(i, 10).Interior.ColorIndex = 38
End If
Next i

Next ws
End Sub
