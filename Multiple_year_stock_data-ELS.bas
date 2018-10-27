Attribute VB_Name = "Module1"
Sub tickervolume():

'What is the last row on my spreadsheet?
  LastRow = Cells(Rows.Count, "A").End(xlUp).Row

'Initialize a variable to store inside
Dim total As Double
total = 0

'Refer to row number for information recorded for total volume to keep track of stock ticker
Dim j As Integer
j = 0

'i is the index or a temporary variable and does not need initializing
For i = 2 To LastRow

'If the next cell ticker does not equal the current cell ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Then we do final calculation for final total the last row for ticker
        total = total + Cells(i, 7).Value
        
'Input the ticker into I (9) column to record the ticker value extracting from current cell
        Cells(2 + j, 9).Value = Cells(i, 1).Value
        
'Input total volume into J (10) column
        Cells(2 + j, 10).Value = total
        
'Reset total volume to zero
        total = 0
        
'j is equal to the number of tickers we've gone through so we add j plus one to go through the tickers
        j = j + 1
        
    Else
'otherwise we keep adding volume to total volume
        total = total + Cells(i, 7).Value
    End If

Next i

End Sub

