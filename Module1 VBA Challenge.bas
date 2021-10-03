Attribute VB_Name = "Module1"
Sub stock():
' Setting Varables
' Setting ticker symbols as string b/c alphabetical letters, this will output every ticker variable in column
Dim ticker As Double

' Setting yearly change as double b/c it will be a decimal number from opening price to closing price
Dim yearly_change As Double

' Setting percent change as double b/c it will be a decimal number from opening and closing prices
Dim percent_change As Double

' Setting total stock volume as integer
Dim total_stock_volume As Integer

' settting the yearly open price as double
Dim yearly_open_price As Double
' Setting the yearly closed price as double
Dim yearly_closed_price As Double

' setting i ,this is current row
Dim i As Double
' setting j, this is the start row of ticker (not sure if I need this?)
Dim j As Integer
' new_ticker for new ticker row
Dim new_ticker As Double

'setting up loop to run through every worksheet
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate
'The last row of each sheet
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row


'setting headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volumn"

    ticker = 2
    j = 2
' (this allows the code to loop through all rows)
    For i = 2 To (last_row + 1)
    
    'Outputting ticker symbol
        If ws.Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'ticker values for col I
        ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
        End If
        
     'Outputting Yearly change (col J) from opening price at the beginning of a given year to teh closing price a the end of year
        ws.Cells(ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value  'Close value - open value (F2-C2)
         
                
       'Conditional formating for yearly change
            If ws.Cells(ticker, 10).Value < 0 Then
            ws.Cells(ticker, 10).Interior.ColorIndex = 3
        
            Else
            ws.Cells(ticker, 10).Interior.ColorIndex = 4
            End If
      
         
     'Outputting percent change (Col K) from opening price
            If ws.Cells(j, 3).Value <> 0 Then
            percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3))
            ws.Cells(ticker, 11).Value = Format(percent_change, "Percent Change")
        
            Else
                ws.Cells(ticker, 11).Value = Format(0, "Percent Change")
            End If
     'Outputting the total stock volume (col L) (use worksheet function sum!!!!!)
        
        ws.Cells(ticker, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(j, 7)))
            
        ticker = ticker + 1
        j = i + 1
        
    
      
        

Next i
Next ws

End Sub
