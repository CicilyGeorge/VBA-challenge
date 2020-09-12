Attribute VB_Name = "Module2"
Sub stockAnalysis()

'Loop through each worksheets
 Dim ws As Worksheet

 For Each ws In ActiveWorkbook.Worksheets
    'Sort table first, if in case any ticker symbols are not in order
    Call sortTable(ws)
 

    'Create the column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Fill the columns Yearly Change, Percent Change and Total Volume for each Ticker
    Call summarizeTicker(ws)
    

    'To Calculate Greatest Increase and Greatest Decrease
    last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim max_percent, min_percent, max_volume As Double
    Dim row_max, row_min, row_max_vol As Long
    
    'Setting First row as max and min
    row_max = 2
    row_min = 2
    row_max_vol = 2
    
    max_percent = ws.Cells(2, 11).Value
    min_percent = ws.Cells(2, 11).Value
    max_volume = ws.Cells(2, 12).Value
    
    
    'Loop through each filtered row (column I)
    For i = 2 To last_row
        
        If ws.Cells(i + 1, 11).Value > max_percent Then
            max_percent = ws.Cells(i + 1, 11).Value
            row_max = i + 1
        End If
        If ws.Cells(i + 1, 11).Value < min_percent Then
            min_percent = ws.Cells(i + 1, 11).Value
            row_min = i + 1
        End If
        If ws.Cells(i + 1, 12).Value > max_volume Then
            max_volume = ws.Cells(i + 1, 12).Value
            row_max_vol = i + 1
        End If
        
    Next i
    
    'Insert Ticker and Values
    ws.Range("P2").Value = ws.Cells(row_max, 9).Value
    ws.Range("Q2").Value = Format(max_percent, "Percent")
    ws.Range("P3").Value = ws.Cells(row_min, 9).Value
    ws.Range("Q3").Value = Format(min_percent, "Percent")
    ws.Range("P4").Value = ws.Cells(row_max_vol, 9).Value
    ws.Range("Q4").Value = max_volume

Next ws

End Sub

Sub sortTable(ws As Worksheet)
    
    'sort_Table Macro
    Columns("A:G").Select
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add2 Key:=Range( _
        "A2:A705714"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ws.Sort
        .SetRange Range("A1:G705714")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 
End Sub

Sub summarizeTicker(ws As Worksheet)
    
    'Defining Column variables
    Dim ticker, percent As String
    Dim open_price, close_price, price_change, price_change_percent, ticker_volume As Double
    open_price = 0
    close_price = 0
    price_change = 0
    price_change_percent = 0
    ticker_volume = 0
    
    'Defining iteration variables
    Dim i, j, last_row As Long
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim current_row As Long
    current_row = 1
    j = 2
    
    'Loop through each row (column A)
        For i = 2 To last_row
        
        'Assigning open_price for each Ticker symbol
        open_price = ws.Cells(j, 3).Value
        
        'Calculate stock volume
        ticker_volume = ticker_volume + ws.Cells(i, 7)
        
        'Filter Ticker symbol
        'if cell A1 not equal to A2
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
           
            'Insert Ticker symbol column values
            current_row = current_row + 1
            ticker = ws.Cells(i, 1).Value
            ws.Cells(current_row, "I").Value = ticker
              
            'Calculate change in Price for each Ticker symbol
            close_price = ws.Cells(i, 6).Value
            price_change = close_price - open_price
            ws.Cells(current_row, "J").Value = price_change
            
            'Conditional formatting for positive and negative values
            If price_change < 0 Then
                ws.Cells(current_row, "J").Interior.ColorIndex = 3
            ElseIf price_change > 0 Then
                ws.Cells(current_row, "J").Interior.ColorIndex = 4
            Else
                ws.Cells(current_row, "J").Interior.ColorIndex = xlNone
            End If

            'Calculate Price change percent for each Ticker symbol
            If open_price <> 0 Then
                price_change_percent = price_change / open_price * 100
                ws.Cells(current_row, "K").Value = Format(price_change_percent, "0.00\%")
            End If
            
            'Insert stock volume for each Ticker symbol
            ws.Cells(current_row, "L").Value = ticker_volume
            ticker_volume = 0
            
            j = i + 1
            
        End If
        
    Next i
    
End Sub


