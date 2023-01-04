Sub stockLoop()

'creating variables
Dim ws As Worksheet
Dim Ticker As String
Dim opening As Double
Dim closing As Double
Dim start_row As Integer
Dim stockVolume As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

'1st For-loop to go through all the worksheets
For Each ws In Worksheets

'initializing variables
start_row = 2
stockVolume = 0
previous_value = 1

    'naming the new columns we will be working on
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'last row of coumn A
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        '2nd for-loop to go through <ticker> data
        For i = 2 To LastRow

            'comparing ticker(s)-->next ticker =! previous ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'getting the ticker symbol
            Ticker = ws.Cells(i, 1).Value
            previous_value = previous_value + 1

            'opening price at beginning of given year (column 3 or 'C')
            opening = ws.Cells(previous_value, 3).Value
        
            'closing price at end of that year (column 6 or 'F')
            closing = ws.Cells(i, 6).Value

            '3rd For-loop for the total stock volume of the stock
            For j = previous_value To i

                'calculating the stock volume (column 7 or 'G')
                stockVolume = stockVolume + ws.Cells(j, 7).Value
                
            'next for 3rd loop
            Next j

            'calculating the yearly change and the percent change
            If opening = 0 Then
            Percent_Change = closing

            Else
            Yearly_Change = closing - opening
            Percent_Change = Yearly_Change / opening

            End If

            'PRINTING VALUES
            ws.Cells(start_row, 9).Value = Ticker
            ws.Cells(start_row, 10).Value = Yearly_Change
            ws.Cells(start_row, 11).Value = Percent_Change
            ws.Cells(start_row, 11).NumberFormat = "0.00%" 'changing format to %(column 11 or 'K')
            ws.Cells(start_row, 12).Value = stockVolume

            'going to next
            start_row = start_row + 1
            
            'Reset to zero
            stockVolume = 0
            Yearly_Change = 0
            Percent_Change = 0
            previous_value = i

            End If

        'next for 2nd for-loop
        Next i
        
'===========================================================
'Conditional formatting = highlighting Yearly Change Values
'===========================================================

'getting last row of the 10th column or column J
jLastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row

'4th for-loop for conditional formatting
For j = 2 To jLastRow

    If ws.Cells(j, 10) > 0 Then
    'if value greater than 0 then color the cell BRIGHT GREEN
    ws.Cells(j, 10).Interior.ColorIndex = 4
    
    Else
    'else value less than 0 then color the cell RED
    ws.Cells(j, 10).Interior.ColorIndex = 3
    
    End If
    
'next for 4th for-loop
Next j

'===========================
        'BONUS?
'===========================

'last row of column K
kLastrow = ws.Cells(Rows.Count, "K").End(xlUp).Row

'naming the new columns and rows we will be working on
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'initializing variables to 0
greatest_increase = 0
greatest_decrease = 0
greatest_total = 0

'5th for-loop to get greatest % and total volume
For k = 2 To kLastrow

    'retrieving the greatest increase
    If ws.Cells(k, 11).Value > greatest_increase Then
    greatest_increase = ws.Cells(k, 11).Value
    
    'Printing greatest increase
    ws.Cells(2, 17).Value = greatest_increase
    ws.Cells(2, 17).NumberFormat = "0.00%" 'printing as percentage format
    ws.Cells(2, 16).Value = ws.Cells(k, 9).Value 'printing the ticker
   
    End If
   
    'retrieving the greatest decrease
    If ws.Cells(k, 11).Value < greatest_decrease Then
    greatest_decrease = ws.Cells(k, 11).Value
    
    'Printing greatest decrease
    ws.Cells(3, 17).Value = greatest_decrease
    ws.Cells(3, 17).NumberFormat = "0.00%" 'printing as percentage format
    ws.Cells(3, 16).Value = ws.Cells(k, 9).Value 'printing the ticker
  
    End If
        
    'Retrieving greatest total volume
    If ws.Cells(k, 12).Value > greatest_total Then
    greatest_total = ws.Cells(k, 12).Value
    
    'Printing greatest total volume
    ws.Cells(4, 17).Value = greatest_total
    ws.Cells(4, 16).Value = ws.Cells(k, 9).Value 'printing the ticker
   
    End If
        
'Next 5th for-loop
Next k
    
'next worksheet
Next ws

'end of the sub
End Sub