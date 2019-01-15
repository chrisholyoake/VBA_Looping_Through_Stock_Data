Sub ChangeStock()

'variable declaration
Dim Volume_counter As Double
Dim Annual_change As Double
Dim Percent_Change As Double
Dim Opening As Double
Dim Closing As Double
Dim Name As String
Dim i As Long
Dim j As Integer
Dim rowCounter As Long
j = 1
i = 2

'count number of rows
rowCounter = Range("A1").End(xlDown).Row

Opening = Cells(i, 3).Value

'Start loop
For i = 2 To rowCounter

'current name and value
Name = Cells(i, 1).Value
Volume_counter = Volume_counter + Cells(i, 7).Value

'statement to sum value and outname value and name to output array if name changes
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   Closing = Cells(i, 6).Value
   Annual_change = Opening - Closing
   Percent_Change = Annual_change / Opening
   j = j + 1
   Cells(j, 9).Value = Name
   Cells(j, 10).Value = Annual_change
   Cells(j, 10).NumberFormat = "0.0000000"
   If Annual_change < 0 Then
       Cells(j, 10).Interior.ColorIndex = 3
   Else
       Cells(j, 10).Interior.ColorIndex = 4
   End If
   Cells(j, 11).Value = Percent_Change
   Cells(j, 11).NumberFormat = "0.0000%"
   Cells(j, 12).Value = Volume_counter
   Cells(j, 12).NumberFormat = "#,###,##0"
   Opening = Cells(i + 1, 3).Value
End If
Next i
End Sub

