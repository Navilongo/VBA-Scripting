Sub wallstreet ()

'Ticker name variable
Dim Ticker_name As String

'Label Tickers Column
Cells(1, 9).Value = "Ticker"

'Label Total Volume column
Cells(1, 12).Value = "Volume"

'Label Change Columns
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"

'Set volume total for Ticker
Dim Ticker_total As Double
Ticker_total = 0

'Note which row we are working on
Dim Ticker_Name_Row As Integer
Ticker_Name_Row = 2

'Set each variable for changes
Dim Open_price As Double
Dim Close_price As Double
Dim Year_Change As Double
Dim Year_Percent As Double

'Note how many rows down we've moved to calculate change
Dim Open_row As Integer


'Loop Worksheets

Dim ws_count As Integer
Dim w As Integer

'Count worksheets active in workbook
ws_count = ActiveWorkbook.Worksheets.Count

'Loop worksheets

For w = 1 To ws_count


    'Reset Open count
    Open_row = 0

    ' Determine the Last Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Format Percent Column
    Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
    'Loop through each ticker
    For i = 2 To LastRow
    
        'Check if next row is different
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            'Set ticker name
            Ticker_name = Cells(i, 1).Value
    
            'Add Ticker Total Volume
            Ticker_total = Ticker_total + Cells(i, 7).Value
    
            'Pull first open in ticker
            Open_price = Cells(i - Open_row, 3).Value
            'Pull closing price
            Close_price = Cells(i, 6).Value
    
            'Set Year Change
            Year_Change = Close_price - Open_price
    
            'Set Year Change Percent
            Year_Percent = (Year_Change) / Open_price
    
            'Print Ticker Name
            Range("I" & Ticker_Name_Row).Value = Ticker_name
    
            'Print the Change from Close - Open
            Range("J" & Ticker_Name_Row).Value = Year_Change
    
                
            'Print the Change from Close - Open
            Range("K" & Ticker_Name_Row).Value = Year_Percent
    
            'Print the total volume for that ticker
            Range("L" & Ticker_Name_Row).Value = Ticker_total
            
            'Move to next row
            Ticker_Name_Row = Ticker_Name_Row + 1
    
            'Reset Volume total
            Ticker_total = 0
    
            'Reset Year_Change Total
            Year_Change = 0
            'Reset how many rows down we've moved to calculate change
            Open_row = 0
    
    
        'If the next cell is the same ticker
        Else
            'Add to the Volume total
            Ticker_total = Ticker_total + Cells(i, 7).Value
        
        End If
    'If on the same ticker, keep adding rows to calculate change in the year
    Open_row = Open_row + 1
    
    Next i
    
Next w

' Determine the Last Row
LastRow_Change = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To LastRow_Change

If Cells(j, 10).Value >= 0 Then
    Cells(j, 10).Interior.ColorIndex = 4
    
Else
Cells(j, 10).Interior.ColorIndex = 3
            
End If
Next j

Columns("I:L").Autofit

End Sub