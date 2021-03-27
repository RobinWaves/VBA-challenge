Attribute VB_Name = "Module1"
Sub StockMarket():

    Dim ws As Worksheet
    Dim last_row As Long
    Dim print_row As Long
    Dim ticker As String
    Dim open_price As Double
    Dim year_chg As Double
    Dim pct_chg As Double
    Dim tot_vol As Double
    
    For Each ws In Worksheets
        
        ' prints headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        'finds last row in sheet
        last_row = ws.Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
        'set initial values
        print_row = 2
        year_chg = 0
        pct_chg = 0
        tot_vol = 0
                
        open_price = ws.Cells(2, 3).Value
        
        For r = 2 To last_row
            
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
                
                'assign & print ticker
                ticker = ws.Cells(r, 1).Value
                ws.Cells(print_row, 9).Value = ticker
                
                'assigns close price, calculates year change, prints year change
                year_chg = ws.Cells(r, 6).Value - open_price
                ws.Cells(print_row, 10).Value = year_chg
                
                'conditional formatting for year change
                If year_chg < 0 Then
                    ws.Cells(print_row, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(print_row, 10).Interior.ColorIndex = 4
                End If
                
                'calculates, formats and prints percent change
                If open_price = 0 Then
                    pct_chg = 0
                Else
                    pct_chg = year_chg / open_price
                End If
                ws.Cells(print_row, 11).NumberFormat = "0.00%"
                ws.Cells(print_row, 11).Value = pct_chg
               
                'calculate total volume and print total volume
                tot_vol = tot_vol + ws.Cells(r, 7).Value
                ws.Cells(print_row, 12).NumberFormat = "#####################"
                ws.Cells(print_row, 12).Value = tot_vol
                
                'sets open price for next ticker
                open_price = ws.Cells(r + 1, 3).Value
                
                print_row = print_row + 1
                tot_vol = 0
            Else
                tot_vol = tot_vol + ws.Cells(r, 7).Value
            End If
        Next r
        
        Dim temp_row As Long

        'prints headers (Bonus)
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        'finds,formats and prints the greatest % increase, decrease and greatest total volume
        ws.Range("P2").NumberFormat = "#####.##%"
        ws.Range("P2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & print_row - 1))
        ws.Range("P3").NumberFormat = "#####.##%"
        ws.Range("P3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & print_row - 1))
        ws.Range("P4").NumberFormat = "#####################"
        ws.Range("P4").Value = Application.WorksheetFunction.Max(ws.Range("L2:G" & print_row - 1))
        
        'prints the ticker associated with greatest % increase, decrease and greatest total volume
        temp_row = Application.WorksheetFunction.Match((ws.Range("P2").Value), ws.Range("K:K"), 0)
        ws.Range("O2").Value = ws.Cells(temp_row, 9).Value
        temp_row = Application.WorksheetFunction.Match((ws.Range("P3").Value), ws.Range("K:K"), 0)
        ws.Range("O3").Value = ws.Cells(temp_row, 9).Value
        temp_row = Application.WorksheetFunction.Match((ws.Range("P4").Value), ws.Range("L:L"), 0)
        ws.Range("O4").Value = ws.Cells(temp_row, 9)
        
        ws.Columns("A:P").AutoFit
        
    Next ws
     
End Sub
