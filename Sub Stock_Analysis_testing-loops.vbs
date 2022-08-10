Sub Stock_Analysis():

' COMPLETELY REDONE
' Start by listing all variable needed
' Just do loops with counters
    'Loops
        'Sheets
        'Ticker
            'count times through loop
            'get ticker symbol value
            'get first and last value of current range
' This is just like Module 4 except just doing the outer loops - no formatting.

' Loop through all worksheets
' Create variable for worksheets
Dim ws_number As Integer
Dim a As Integer    ' counter for worksheets
Dim i As Integer, Dim j as Integer  ' for loops through ticker
Dim ws As Worksheet ' DON'T THINK I NEED THIS
Dim LastRow As Long
Dim ticker As String
Dim open_value As Double
Dim close_value As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim total_volume As Long
Dim Summary_Table_Row As Integer
Dim row_count As Long
Dim ticker_range As Range
Dim start_rng As Range
Dim end_rng As Range
Dim start_date As Double
Dim end_date As Double


Set ws = ActiveSheet    ' remember which worksheet is active in the beginning

ws_number = ThisWorkbook.Worksheets.Count   ' total number of worksheets


    ' FIRST LOOP THROUGH ALL WORKSHEETS
    For a = 1 To ws_number
        ThisWorkbook.Worksheets(a).Activate
        ' MsgBox a
        
        ' DETERMINE THE LAST ROW
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' SPECIFY LOCATION FOR SOLUTIONS
        Summary_Table_Row = 2


        With Worksheets(a)
            '---------------------------------
            ' FIRST GROUP OF COMMANDS IS TO ALLOW TESTING OF ITERATIONS
            ' COMMENT THESE OUT WHEN HAVE EVERYTHING WORKING.

            ' CLEAR CONTENTS TO KEEP TESTING LOOPS
            .Range("H1:L1").EntireColumn.ClearContents
            .Cells(1, 1).ClearContents
            ' --------------------------------

            ' ADD ALL NEW COLUMNS ONE LINE
            .Range("H1:L1").EntireColumn.Insert
            
            ' ADD COLUMN HEADERS
            .Cells(1, 8) = "Count"
            .Cells(1, 9) = "Ticker"
            .Cells(1, 10) = "Yearly Change"
            .Cells(1, 11) = "Percent Change"
            .Cells(1, 12) = "Total Stock Volume"
        End With
 
        ' SECOND LOOP THROUGH ALL ROWS ON WORKSHEET:
        For i = 2 To LastRow

            ' FROM COLUMN A, GET UNIQUE TICKER SYMBOL:
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Set the ticker symbol
                ticker = Cells(i, 1).Value
'                        If Cells(i, 1) = Cells(i, 1) Then  ' this worked but gave me the count showing how many ticker symbols - not how many times for each
'                        row_count = row_count + 1
'                        End If
                ' Print the ticker symbol in the Summary Table in column I:
                Range("I" & Summary_Table_Row).Value = ticker

            j = 0
            Do Until i = i + 1
                Cells(i, 1).Count = j
                j = j + 1
            Loop
                
'                    Dim rng As Range
'                    Dim rngCell As Range
'                    Next i
                
'                        If ticker = ticker Then     ' this worked but gave me the count showing how many ticker symbols - not how many times for each
'                        row_count = row_count + 1
'                        End If
'                        row_count = Application.WorksheetFunction.CountIf(Range("A" & i).Value, Range("I" & Summary_Table_Row).Value) = 1
                
                    
'                    If Worksheets(a).Range("A" & i).Value = Worksheets(a).Range("A" & i + 1).Value Then
'                     While Worksheets(a).Range("A" & i).Value = Worksheets(a).Range("A" & i + 1).Value
'                        row_count = Worksheets(a).Range("A" & i).Rows.Count
'                        row_count = row_count + 1
'                    Wend
'                    If WorksheetFunction.CountIf(ws.Range("A" & LastRow), ws.Range("A" & i)) > 1 Then
'                        If WorksheetFunction.CountIf(ws.Range("A" & i), ws.Range("A" & i)) = 1 Then
'                       start_rng = ws.Range("A" & i)
'                       Else
'                       end_rng = ws.Range(("A" & i + 1) - 1)
'                       End If
'                    End If
'                    ticker_range = end_date - start_date
'                    row_count = ws.Range(ticker_range).Rows.Count

'                    start_date = Cells(i, 2).Value
'                    end_date = Cells(i + 1, 2).Value
'                    start_rng = Range(Cells(i, 2))
'                    end_rng = Range(Cells(i + 1, 2))
'                    ticker_range = end_rng - start_rng
'                    row_count = Range(ticker_range).Rows.Count

'                    Do While Cells(i, 1).Value
'                        start_date = Application.Min(Range(Cells(i, 2).Value)
'                         end_date = Application.Max(Cells(i, 2).Value)
'                         ticker_range = end_date - start_date
'                        row_count = Range(ticker_range).Rows.Count
'                    Loop
'
'                    row_count = 0
'                    row_count = Application.WorksheetFunction.CountIf(ticker_range, ticker)
'                    If WorksheetFunction.CountIf(ws.Range("A" & i), ws.Range("A" & i)) = 1 Then
'                        row_count = Cells(i, 1).Rows.Count
'                    End If
'                    row_count = row_count + 1
                                    
'                    row_count = Application.WorksheetFunction.CountIf(ticker_range, ticker)
'                    Do While WorksheetFunction.CountIf(ws.Range("A" & i), ws.Range("A" & i)) = 1
'                        row_count = Cells(i, 1).Rows.Count + 1
'                       row_count = row_count + 1
'                    Loop
            
                    
                    
                    
                    
'                    Do While Cells(i + 1, 1).Value = Cells(i, 1).Value
'                        row_count = Cells(i, 1).Value.Count + 1
'                    Loop
                    Range("H" & Summary_Table_Row).Value = row_count
                    
'                        For j = 2 And While Cells(i +1, 1).Value = Cells(i, 1).Value


                ' If the cell immediately following a row is the same ticker symbol...
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
'
'                ' For column B, offset column range A
'                'Range("A1").Offset(1, 1).Activate
'                ' Use functions and offsets to get opening price, closing price, and total volume of stock traded.
'                'open_value = Application.WorksheetFunction.Min("B1").Offset(1)
'                'close_value = Application.WorksheetFunction.Max("B1").Offset(4)
'                'total_volume = Application.WorksheetFunction.Sum("B1")
'
'                ' Calculate the Yearly Change in stock value:
'                'yearly_change = close_value - open_value
'                ' Print the Yearly Change to the Summary Table
'                Range("J" & Summary_Table_Row).Value = yearly_change
'                ' Format as Number with 2 decimal places to Yearly Change
'                With ws.Range("J1")
'                .NumberFormat = "0.00"
'                .HorizontalAlignment = xlRight
'                 End With
'                 ' Apply Conditional Formatting to Yearly Change
''                        If Range("J" & yearly_change).Value >= 0 Then
''                            Range("J" & yearly_change).Interior.ColorIndex = 4
''                        Else
''                            Range("J" & yearly_change).Interior.ColorIndex = 3
''                        End If
'                For Each Cell In .Range("J1")
'                    If Cell.Value < 0 Then
'                        Cell.Interior.Color = vbRed
'                    Else
'                        Cell.Interior.Color = vbGreen
'                        Exit For
'                    End If

'
'                ' Calculate the Percentage Change in stock value:
''                percentage_change = (yearly_change / open_value) * 100
'                ' Print the Percentage Change to the Summary Table
'                Range("K" & Summary_Table_Row).Value = percentage_change
'                    ' Format new column Yearly Change as Number with 2 decimal places and Right Justified
'                    With ws.Range("K1")
'                    .NumberFormat = "0.00%"
'                    .HorizontalAlignment = xlRight
'                    End With
'
'                ' Add to the total volume of trades for that ticker symbol
'                    ' Format as General with contents right justified
'                    With ws.Range("L1")
'                    .NumberFormat = "0"
'                    .HorizontalAlignment = xlRight
'                    End With
'
'
'
'
'                Else
'
                End If

            Next i

    ' FORMAT THE CONTENT OF THE ADDED COLUMNS
    With .Range("H1")
        .EntireColumn.AutoFit
        ' Format column H
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
    End With

    With .Range("I1")
        .EntireColumn.AutoFit
        ' Format column I
        .NumberFormat = "General"
        .HorizontalAlignment = xlLeft
    End With

    With .Range("J1")
        .EntireColumn.AutoFit
        ' Format column J
        .NumberFormat = "0.00"
        .HorizontalAlignment = xlRight
    End With

    With .Range("K1")
        .EntireColumn.AutoFit
        ' Format column K
        .NumberFormat = "0.00%"
        .HorizontalAlignment = xlRight
    End With

    With .Range("L1")
        .EntireColumn.AutoFit
        ' Format column L
        .NumberFormat = "0"
        .HorizontalAlignment = xlRight
    End With







' To show this loop for Ticker values has run, this sets cell A1 of each sheet to "1".
ThisWorkbook.Worksheets(a).Cells(1, 1) = a
' THIS WORKS THE WAY IT CURRENTLY IS COMMENTED OUT.  7/9/2022 10:45 AM  ' Ticker column fine,Yearly Change and Percent Change all zeroes
' total stock volume all blanks, no formatting of cells in columns I through L
 
Next
 
        
End Sub



