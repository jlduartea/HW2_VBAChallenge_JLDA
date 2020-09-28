'Homework VBA Challenge by Jose Luis Duarte Alcantara

Sub summary_tickets_assigment()

'variables definition
    Dim ticker_name As String
    Dim last_row, ticker_total, yr_chg_fin, yr_chg_ini, prc_chg As Double
    Dim summary_table_row As Integer
    
    
'variables initialization
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    summary_table_row = 2
    ticker_total = 0
    
'headers of summary
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 13).Value = "Open"
    Cells(1, 14).Value = "Close"
    
    
'counting tickers code
    For i = 2 To last_row
        
        'if the cell inmediately following a row is different ticker...
        If Cells(i + 1, 1) <> Cells(i, 1) Then
            ticker_name = Cells(i, 1).Value
            ticker_total = ticker_total + Cells(i, 7).Value
            yr_chg_fin = Cells(i, 6).Value - yr_chg_ini
            If yr_chg_ini = 0 Then
                prc_chg = 0
            Else
                prc_chg = yr_chg_fin / yr_chg_ini
            End If
    
            Range("I" & summary_table_row).Value = ticker_name
            Range("J" & summary_table_row).Value = yr_chg_fin
            Range("K" & summary_table_row).Value = prc_chg
            Range("L" & summary_table_row).Value = ticker_total
            Range("M" & summary_table_row).Value = yr_chg_ini
            Range("N" & summary_table_row).Value = Cells(i, 6).Value
            
            'set the colors base on positive o negative year change value
            If yr_chg_fin < 0 Then
                Range("J" & summary_table_row).Interior.ColorIndex = 3
            ElseIf (yr_chg_fin = 0) Then
                Range("J" & summary_table_row).Interior.ColorIndex = 6
            Else
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            
            End If
            
        ' Reset all acummulative variables
            summary_table_row = summary_table_row + 1
            yr_chg_ini = Cells(i + 1, 3).Value
            ticket_total = 0

        ' If the cell previously a row is the different ticker...
        ElseIf Cells(i - 1, 1) <> Cells(i, 1) Then
            yr_chg_ini = Cells(i, 3).Value
            ticker_total = ticker_total + Cells(i, 7).Value
        
        ' If the cell immediately following a row is the same ticker...
        Else
        ' Add to the Brand Total
            ticker_total = ticker_total + Cells(i, 7).Value

        End If

    Next i
    
End Sub

Sub reset_summaries()
    Dim max_row2 As Integer
    max_row2 = Cells(Rows.Count, 10).End(xlUp).Row
    
    Range("i2:n" & max_row2).ClearContents
    Range("q2:s10").ClearContents
    Range("i2:n10000").ClearFormats
    
End Sub

Sub summary_greatest_tickers()
    Dim max_tk, min_tk, max_vol As Double
    Dim max_row  As Integer
    Dim max_vol_name As String
    Dim max_vol_row, max_tk_row, min_tk_row As Object
    

    max_row = Cells(Rows.Count, 10).End(xlUp).Row
    max_tk = Application.WorksheetFunction.Max(Range("J2:J" & max_row))
    min_tk = Application.WorksheetFunction.Min(Range("J2:J" & max_row))
    max_vol = Application.WorksheetFunction.Max(Range("L2:L" & max_row))
    
    Set max_tk_row = Range("j2:j" & max_row).Find(max_tk)
    Set min_tk_row = Range("j2:j" & max_row).Find(min_tk)
    Set max_vol_row = Range("l2:l" & max_row).Find(max_vol)

'Headers
    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    
    Cells(2, 18).Value = max_tk
    Cells(3, 18).Value = min_tk
    Cells(4, 18).Value = max_vol
    Cells(2, 17).Value = Cells(max_tk_row.Row, 9).Value
    Cells(3, 17).Value = Cells(min_tk_row.Row, 9).Value
    Cells(4, 17).Value = Cells(max_vol_row.Row, 9).Value
    
End Sub
Sub create_summaries_full()

    Dim WS_Count As Integer
    Dim j As Integer
    'Dim pctdone2 As Single
    ' Set WS_Count equal to the number of worksheets in the active workbook.
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    'ufProgress.LabelProgress.Width = 0
    'ufProgress.Show

    ' Begin the loop.
    
    For j = 1 To WS_Count
    '    pctdone2 = j / WS_Count
        Worksheets(ActiveWorkbook.Worksheets(j).Name).Activate
        summary_tickets_assigment
        summary_greatest_tickers
    '    With ufProgress
    '        .LabelCaption.Caption = "Processing Sheet " & j & "  of  " & WS_Count
    '        .LabelProgress.Width = pctdone2 * (.FrameProgress.Width)
    '    End With
    '    DoEvents
    '    If j = WS_Count Then Unload ufProgress
    Next j
Worksheets(ActiveWorkbook.Worksheets(1).Name).Activate
End Sub
Sub reset_full_summaries()

Dim WS_Count2 As Integer
Dim k As Integer
'Dim pctdone As Single

    ' Set WS_Count equal to the number of worksheets in the active workbook.
    
    WS_Count2 = ActiveWorkbook.Worksheets.Count
    'ufProgress.LabelProgress.Width = 0
    'ufProgress.Show
    
    ' Begin the loop.
    
    For k = 1 To WS_Count2
        'pctdone = k / WS_Count2
        Worksheets(ActiveWorkbook.Worksheets(k).Name).Activate
        reset_summaries
        'With ufProgress
        '    .LabelCaption.Caption = "Processing Sheet " & k & "  of  " & WS_Count2
        '    .LabelProgress.Width = pctdone * (.FrameProgress.Width)
        'End With
        'DoEvents
        'If k = WS_Count2 Then Unload ufProgress
    Next k
Worksheets(ActiveWorkbook.Worksheets(1).Name).Activate
End Sub