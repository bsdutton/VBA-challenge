Multiple_year_stock_data_VBA-code

Option Explicit

Sub Stock_Analysis():

' Loop through all worksheets
' Create variable for worksheets
Dim ws_number As Integer
Dim a As Integer    ' counter for worksheets
Dim i As Long    ' for loops through ticker
Dim row_count As Long
Dim offset_rows As Long

'Dim j As Integer    ' for loops to count times through each ticker
Dim ws As Worksheet
Dim LastRow As Long

Dim start_row As Long
Dim end_row As Long

Dim ticker As String
Dim min_value As Double
Dim max_value As Double

Dim open_value As Double
Dim close_value As Double

Dim yearly_change As Double
Dim percentage_change As Double

Dim great_percent As Double
Dim least_percent As Double

Dim great_percent_ticker As String
Dim least_percent_ticker As String

Dim greatest_percent As Double
Dim worst_percent As Double

Dim greatest_percent_ticker As String
Dim worst_percent_ticker As String

Dim yearly_change1 As Double
Dim percentage_change1 As Double

Dim total_volume As Double
Dim great_vol As Double
Dim greatest_vol As Double
Dim great_vol_ticker As String
Dim greatest_vol_ticker As String

Dim Summary_Table_Row As Long

Dim ticker_range As Range
Dim start_rng As Range
Dim end_rng As Range
Dim start_date As Long
Dim end_date As Long

Dim AddRow As Integer

Dim after_header As Integer




Set ws = ActiveSheet    ' remember which worksheet is active in the beginning

ws_number = ThisWorkbook.Worksheets.Count   ' total number of worksheets


    ' FIRST LOOP THROUGH ALL WORKSHEETS
    
    For a = 1 To ws_number
        ThisWorkbook.Worksheets(a).Activate
        
        ' DETERMINE THE LAST ROW
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' SPECIFY LOCATION FOR SOLUTIONS
        Summary_Table_Row = 2

        With Worksheets(a)  ' THAT ADDS COLUMNS AND HEADERS
            '---------------------------------
            ' CLEAR CONTENTS TO KEEP TESTING LOOPS
            .Range("H1:P1").EntireColumn.ClearContents
            .Cells(1, 1).ClearContents
            ' --------------------------------

            ' Format column G - just to check
            With .Range("G1")
             .EntireColumn.AutoFit
            .NumberFormat = "0"
            .HorizontalAlignment = xlRight
            End With

            ' ADD ALL NEW COLUMNS ADDED
            .Range("H1:P1").EntireColumn.Insert
            
            ' ADD COLUMN HEADERS
            .Cells(1, 8) = "Count"
            .Cells(1, 9) = "Ticker"
            .Cells(1, 10) = "Yearly Change"
            .Cells(1, 11) = "Percent Change"
            .Cells(1, 12) = "Total Stock Volume"
        End With    ' THAT ADDS COLUMNS AND HEADERS
 
        row_count = 0
        total_volume = 0
        after_header = 2
      
        
        ' SECOND LOOP THROUGH ALL ROWS ON WORKSHEET:
        For i = 2 To LastRow    ' TO LOOP THROUGH ALL ROWS OF COLUMN A TO GET TICKER SYMBOLS
            
            row_count = row_count + 1   ' FOR START OF ROW COUNTER FOR TICKER SYMBOL
            start_row = i   ' This works
            total_volume = total_volume + Cells(i, 7).Value
                        
            ' FROM COLUMN A, GET UNIQUE TICKER SYMBOL:
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then  ' CONDITIONAL CHECK TO GET ALL DUPLICATE TICKER VALUES
               
                ' Set the ticker symbol
                ticker = Cells(i, 1).Value
                offset_rows = i - (row_count - 1)   ' had to account for adding one to row_count
                open_value = Cells(offset_rows, 3).Value
                close_value = Cells(i, 6).Value
                yearly_change = close_value - open_value
                percentage_change = (yearly_change / open_value) ' * 100
                
                ' ADD TICKER SYMBOLS TO TICKER COLUMN, COUNT FOR EACH TICKER SYMBOL TO COUNT COLUMN:
                Range("H" & Summary_Table_Row).Value = row_count
                Range("I" & Summary_Table_Row).Value = ticker
                
                Range("J" & Summary_Table_Row).Value = yearly_change
                Range("K" & Summary_Table_Row).Value = percentage_change

                Range("L" & Summary_Table_Row).Value = total_volume
                
                great_percent = Application.WorksheetFunction.Max(Range(Cells(after_header, 11), Cells(Summary_Table_Row - 1, 11)))
                least_percent = Application.WorksheetFunction.Min(Range(Cells(after_header, 11), Cells(Summary_Table_Row - 1, 11)))
                
'                great_percent_ticker = Range(Cells(great_percent, 11).Offset(0, -2)).Value
'                least_percent_ticker = Range(Cells(least_percent, 11).Offset(0, -2)).Value
                
                Range("K" & Summary_Table_Row + 1).Value = great_percent
                Range("K" & Summary_Table_Row + 2).Value = least_percent
                Range("I" & Summary_Table_Row + 1).Value = great_percent_ticker
                Range("I" & Summary_Table_Row + 2).Value = least_percent_ticker
               
                great_vol = Application.WorksheetFunction.Max(Range(Cells(after_header, 12), Cells(Summary_Table_Row - 1, 12)))
'                great_vol_ticker = Range(Cells(great_vol, 12).Offset(0, -3)).Value

                Range("L" & Summary_Table_Row + 1).Value = great_vol
'                Range("M" & Summary_Table_Row + 1).Value = great_vol_ticker
               
                Summary_Table_Row = Summary_Table_Row + 1   ' ADVANCE TO NEXT LINE IN SUMMARY TABLE ROW WITH ADVANCE TO NEXT TICKER SYMBOL
                row_count = 0   'FOR END OF ROW COUNTER FOR TICKER SYMBOL
                end_row = 0
                total_volume = 0
                
            End If  ' END LOOP FOR DUPLICATE TICKER VALUES
            
        Next i  ' FOR NEXT TICKER SYMBOL
        
        With Worksheets(a)  ' THAT FORMATS COLUMNS
                
            ' Format column H
            With .Range("H1")
            .EntireColumn.AutoFit
            .NumberFormat = "0"
            .HorizontalAlignment = xlCenter
            End With
            
            ' Format column I
            With .Range("I1")
            .EntireColumn.AutoFit
            .NumberFormat = "General"
            .HorizontalAlignment = xlLeft
            End With
            
            ' Format column J
            With .Range("J1")
            .EntireColumn.AutoFit
            .NumberFormat = "0.00"
            .HorizontalAlignment = xlRight
            End With
            
            ' Format column K
            With .Range("K1")
            .EntireColumn.AutoFit
            .NumberFormat = "0.00%" ' This did nothing
            .HorizontalAlignment = xlRight
            End With
            
            ' Format column L
            With .Range("L1")
            .EntireColumn.AutoFit
            .NumberFormat = "0"
            .HorizontalAlignment = xlRight
            End With
       
            ' I got the example to follow for the following conditional formatting from:
            ' https://www.wallstreetmojo.com/vba-conditional-formatting/
                       
            Dim condition1 As FormatCondition
            Dim condition2 As FormatCondition
            Dim condition3 As FormatCondition
            Dim condition4 As FormatCondition

            
            Set condition1 = .Range(Cells(after_header, 10), Cells(Summary_Table_Row - 1, 10)).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
            Set condition2 = .Range(Cells(after_header, 10), Cells(Summary_Table_Row - 1, 10)).FormatConditions.Add(xlCellValue, xlLess, "=0")
            Set condition3 = .Range(Cells(after_header, 11), Cells(Summary_Table_Row + 1, 11)).FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
            Set condition4 = .Range(Cells(after_header, 11), Cells(Summary_Table_Row + 1, 11)).FormatConditions.Add(xlCellValue, xlLess, "=0")
    
    
            With condition1
                .Interior.ColorIndex = 4
            End With
            
            With condition2
                .Interior.ColorIndex = 3
            End With
                   
            With condition3
                .Interior.ColorIndex = 4
                .NumberFormat = "0.00%"
            End With
            
            With condition4
                .Interior.ColorIndex = 3
                .NumberFormat = "0.00%"
            End With
                       
                       
        End With    ' THAT FORMATS COLUMNS
        
        
        ThisWorkbook.Worksheets(a).Cells(1, 1) = a
        
    Next
    
'    greatest_percent = Application.WorksheetFunction.Max(Worksheets("Sheet 1:Sheet 6").Range(Cells(Summary_Table_Row + 1, 11)))
'    worst_percent = Application.WorksheetFunction.Min(Worksheets("Sheet 1:Sheet 6").Range(Cells(Summary_Table_Row + 2, 11)))
'
'    Worksheets("Sheet 1").Cells(3, 14).Value = greatest_percent
'    Worksheets("Sheet 1").Cells(4, 14).Value = worst_percent
'
'    greatest_vol = Application.WorksheetFunction.Max(Worksheets("Sheet 1:Sheet 6").Range(Cells(Summary_Table_Row + 1, 12)))
'
'    Worksheets("Sheet 1").Cells(5, 14).Value = worst_percent
    
    
    
    
'    greatest_percent = Application.WorksheetFunction.Max(Worksheets("Sheet 1:Sheet 6")(great_percent))
'    worst_percent = Application.WorksheetFunction.Min(Worksheets("Sheet 1:Sheet 6")(least_percent))
'
'    Worksheets("Sheet 1").Cells(3, 14).Value = greatest_percent
'    Worksheets("Sheet 1").Cells(4, 14).Value = worst_percent
'
'    greatest_vol = Application.WorksheetFunction.Max(Worksheets("Sheet 1:Sheet 6")(great_vol))
'    Worksheets("Sheet 1").Cells(5, 14).Value = worst_percent
   
    
    End Sub

            
                

