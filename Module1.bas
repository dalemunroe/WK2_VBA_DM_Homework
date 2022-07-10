Attribute VB_Name = "Module1"
Sub WALLST()

'---------------------------------------------------------------------------
' Assignment Week 2 - VBA   Student:Dale Munroe
'---------------------------------------------------------------------------

' LOOP THROUGH ALL SHEETS
    For Each ws In Worksheets
        
        'Create a Variable to Hold File Name
            Dim WorksheetName As String
        
        'Create a Variable to Hold Last Row
            last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (last_row)
            
        'Grabbed the WorksheetName
            WorksheetName = ws.Name
            'MsgBox WorksheetName
    
        'Set an initial variable for holding the ticker name
            Dim ticker_name As String
        
        'Set an initial variable for holding the ticker count
            Dim ticker_count As Integer
            ticker_count = 0
        
        'Set an initial variable for holding the ticker_opening_annual
            Dim ticker_opening_annual As Double
            ticker_opening_annual = 0
    
        'Set an initial variable for holding the ticker_close_annual
            Dim ticker_close_annual As Double
            ticker_close_annual = 0
        
        'Set an initial variable for holding the ticker_annual_change
            Dim ticker_annual_change As Double
            ticker_annual_change = 0
    
        'Set an initial variable for holding the ticker_percentage_change
            Dim ticker_percentage_change As Double
            ticker_percentage_change = 0
     
        'Set an initial variable for holding .....
            Dim ticker_vol_total As LongLong
            ticker_vol_total = 0
    
        'Keep track of the location for each ticker name in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
        
        'Prints headings for macra data columns
            Dim Heading_array1(3) As String
            Heading_array1(0) = "Ticker"
            Heading_array1(1) = "Yearly Change"
            Heading_array1(2) = "Percent Change"
            Heading_array1(3) = "Total Stock Volume"
        
            ws.Range("I1:L1").Value = Heading_array1
            ws.Range("I1:L1").Font.Bold = True
            ws.Columns("I:L").EntireColumn.AutoFit
                    
            'last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'MsgBox (last_row)
        
    
        'Loop through all ticker_name records
            For i = 2 To last_row
    
            'Check if we are still within the same ticker code, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            'Set the ticker_name
                ticker_name = ws.Cells(i, 1).Value
            
            'Add to the ticker_vol_total
                ticker_vol_total = (ticker_vol_total) + ws.Cells(i, 7).Value
    
            'Print the ticker_name in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
    
            'Print the ticker_count to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = ticker_vol_total
            
            'Print the annual opening price
            
                ticker_opening_annual = ws.Cells(i - ticker_count, 3).Value
                'ws.Range("M" & Summary_Table_Row).Value = ticker_opening_annual
                           
                'Print the annual closing price
            
                ticker_close_annual = ws.Cells(i, 6).Value
                'ws.Range("N" & Summary_Table_Row).Value = ticker_close_annual
            
                'Print the Annual Change
                ticker_annual_change = ticker_close_annual - ticker_opening_annual
                ticker_annual_change = Application.WorksheetFunction.Round(ticker_annual_change, 2)
                ws.Range("J" & Summary_Table_Row).Value = ticker_annual_change
                ws.Columns("J").NumberFormat = "#,##0.00_)"
                               
                
                If (ticker_annual_change < 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 9
                    ws.Range("J" & Summary_Table_Row).Font.ColorIndex = 2
                
                ElseIf (ticker_annual_change > 0) Then
                    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 10
                    
                End If
                
                
                'Print the Percentage Change
                ticker_percentage_change = ticker_annual_change / ticker_opening_annual * 100
                ticker_percentage_change = Application.WorksheetFunction.Round(ticker_percentage_change, 2)
                ws.Range("K" & Summary_Table_Row).Value = ticker_percentage_change & "%"
            
            
                'Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
          
                'Reset the ticker_vol_total
                ticker_vol_total = 0
           
                'Reset the ticker_vol_count
                ticker_count = 0
    
                'If the cell immediately following a row is the same brand...
        
            Else
    
            'Add to the ticker_vol_total
            ticker_vol_total = (ticker_vol_total) + ws.Cells(i, 7).Value
          
            ticker_count = ticker_count + 1
                
            'Set the annual opening price
            
            'ticker_opening_annual = ws.Cells(i, 3).Value
            'ws.Range("M" & Summary_Table_Row).Value = ticker_opening_annual
        End If
        
    
    
        Next i
    
'----------------------------------------------------------------------------------------------------------

'Bonus Element
    
        'Prints headings for macra data columns
        Dim Heading_array2(2) As String
        Heading_array2(0) = ""
        Heading_array2(1) = "Ticker"
        Heading_array2(2) = "Value"
        
        
        ws.Range("N1:P1").Value = Heading_array2
        ws.Range("N1:P1").Font.Bold = True
        ws.Columns("O:P").ColumnWidth = 15
        ws.Columns("M:N").ColumnWidth = 25
        ws.Range("P2:P3").Style = "Percent"
        ws.Range("P2:P3").NumberFormat = "#,##0.00%"
                        
        ws.Range("N2").Value = "Greatest % Increase:"
        ws.Range("N3").Value = "Greatest % Decrease:"
        ws.Range("N4").Value = "Greatest Total Volume:"
        
        
        Dim max_ticker_value As Double
        Dim min_ticker_value As Double
        Dim max_traded_value As LongLong
        
        Dim Max_ticker As Range
        Dim Min_ticker As Range
        Dim Max_traded As Range
        
        'Set range from which to determines max_ticker_value
        Set Max_ticker = ws.Range("K:K")
        
        'Worksheet function Max returns the Highest value in a range

        max_ticker_value = Application.WorksheetFunction.Max(Max_ticker)
        ws.Range("P2").Value = max_ticker_value

        'Set range from which to determines min_ticker_value
        Set Min_ticker = ws.Range("K:K")
        
        'Worksheet function Min returns the Lowest value in a range

        min_ticker_value = Application.WorksheetFunction.Min(Min_ticker)
        ws.Range("P3").Value = min_ticker_value
        
        'Set range from which to determines max_traded_value
        Set Max_traded = ws.Range("L:L")
        
        'Worksheet function Max returns the Highest Traded volume in a range

        max_traded_value = Application.WorksheetFunction.Max(Max_traded)
        ws.Range("P4").Value = max_traded_value
        
        'Populates the Bonus table with ticker IDs to match....
        
        For Z = 2 To 10000
        
            If ws.Cells(Z, 11) = ws.Cells(2, 16) Then
            ws.Cells(2, 15) = ws.Cells(Z, 9).Value
            
            ElseIf ws.Cells(Z, 11) = ws.Cells(3, 16) Then
            ws.Cells(3, 15) = ws.Cells(Z, 9).Value
            
            ElseIf ws.Cells(Z, 12) = ws.Cells(4, 16) Then
            ws.Cells(4, 15) = ws.Cells(Z, 9).Value
            
            End If
        
        Next Z

    
    Next ws
    
    
MsgBox ("Complete")


End Sub

