Attribute VB_Name = "Module1"
'***************************************************************************************
'**
'**  VBA Homework: The VBA of Wall Street
'**
'**    Author: George Alonzo
'**  Due Date: Sept 11, 2021
'**
'***************************************************************************************

Sub StockAnalysis()

    'Dim wsName As String
    
    Dim BeginningOpen As Double
    Dim EndingClose As Double
    Dim YearlyChange As Double
    Dim PctChange As Double
    Dim OpenValue As Double
    Dim TotStockVol As Double
    
    Dim SummaryRow As Double
    Dim w As Long
    
    'THE FOLLOWING VARIABLES ARE FOR THE GREATEST SUMMARY TABLE
    Dim MaxPctIncTicker As String
    Dim MaxPctIncValue As Double
    Dim MaxPctDecTicker As String
    Dim MaxPctDecValue As Double
    Dim MaxTotVolTicker As String
    Dim MaxTotVolValue As Double
    
    
    For Each ws In Worksheets
        
        'ACTIVATE WORKSHEET TO PRINT SUMMARY ON
        ws.Activate
        'SET FIRST ROW FOR PRINTING SUMMARY TABLE
        SummaryRow = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'INITIALIZE VARIABLES FOR EACH SHEET
        MaxPctIncValue = 0
        MaxPctDecValue = 0
        MaxTotVolValue = 0
        
        BeginningOpen = Range("C2").Value
        '*** KEEP ADVANCING UNTIL FIRST NON-ZERO BEGINNING OPEN IS FOUND, IF ANY
        '*** THIS HELPS PREVENT DIV/0 ERROR WHEN CALCULATING PERCENT CHANGE WHEN
        '*** BEGINNING OPEN VALUE IS $0
        w = 2
        If BeginningOpen = 0 And Not IsEmpty(Range("A" & w + 1).Value) Then
            
            Do Until (Range("A" & w).Value <> Range("A" & w + 1) And Not IsEmpty(Range("A" & w + 1).Value)) Or BeginningOpen > 0
                BeginningOpen = Range("C" & w + 1).Value
                w = w + 1
            Loop
            
        End If
        
        'BUILD STATIC COLUMN HEADERS FOR SUMMARY
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Pct Change"
        Range("L1").Value = "Tot Stock Vol"
        
        'BeginningOpen = Range("C2").Value
        
        For r = 2 To LastRow
        
            If Range("A" & r).Value <> Range("A" & r + 1).Value Then
                
                'PRINT TO SUMMARY TABLE, PERFORM CALCULATIONS PRIOR TO WHERE NECESSARY
                
                'PRINT TICKER NAME TO SUMMARY TABLE
                Range("I" & SummaryRow).Value = Range("A" & r).Value
                
                'PREP AND PRINT YEARLY CHANGE TO SUMMARY TABLE
                EndingClose = Range("F" & r).Value
                YearlyChange = EndingClose - BeginningOpen
                Range("J" & SummaryRow).Value = YearlyChange
                If YearlyChange > 0 Then
                    Range("J" & SummaryRow).Cells.Interior.Color = RGB(0, 255, 0)
                Else
                    Range("J" & SummaryRow).Cells.Interior.Color = RGB(255, 0, 0)
                End If
                
                'PREP AND PRINT PERCENT CHANGE TO SUMMARY TABLE
                
                'IF BEGINNING OPEN VALUE IS $0 EVEN AFTER PREVIOUSLY ATTEMPTING TO FIND THE
                'FIRST NON-$0 VALUE, SUBSTITUTE PERCENT CHANGE WITH $0
                If BeginningOpen = 0 Then
                    Range("K" & SummaryRow).Value = "0"
                Else
                    PctChange = YearlyChange / BeginningOpen
                    Range("K" & SummaryRow).Value = PctChange
                    Range("K" & SummaryRow).NumberFormat = "0.00%"
                End If
                
                'PREP AND PRINT TOTAL STOCK VOLUME TO SUMMARY TABLE
                TotStockVol = TotStockVol + Range("G" & r).Value
                Range("L" & SummaryRow).Value = TotStockVol
    
                'COMPARE AND STORE VALUES FOR GREATEST VALUES IN EACH SHEET
                If PctChange <> 0 And PctChange > MaxPctIncValue Then
                    MaxPctIncTicker = Range("A" & r).Value
                    MaxPctIncValue = PctChange
                End If
                
                If PctChange <> 0 And PctChange < MaxPctDecValue Then
                    MaxPctDecTicker = Range("A" & r).Value
                    MaxPctDecValue = PctChange
                End If
                
                If TotStockVol > MaxTotVolValue Then
                    MaxTotVolTicker = Range("A" & r).Value
                    MaxTotVolValue = TotStockVol
                End If
                
                'INITIALIZE VARIABLES FOR NEXT CHANGE IN TICKER
                TotStockVol = 0
                SummaryRow = SummaryRow + 1
                BeginningOpen = Range("C" & r + 1).Value
                
                
                '*** KEEP ADVANCING UNTIL FIRST NON-ZERO BEGINNING OPEN IS FOUND, IF ANY
                '*** THIS HELPS PREVENT DIV/0 ERROR WHEN CALCULATING PERCENT CHANGE WHEN
                '*** BEGINNING OPEN VALUE IS $0
                If BeginningOpen = 0 And Not IsEmpty(Range("A" & r + 1).Value) Then
               
                    w = r + 1
                
                    Do Until (Range("A" & w).Value <> Range("A" & w + 1) And Not IsEmpty(Range("A" & w + 1).Value)) Or BeginningOpen > 0
                        BeginningOpen = Range("C" & w + 1).Value
                        w = w + 1
                    Loop
                End If
            Else
                TotStockVol = TotStockVol + Range("G" & r).Value
            End If
        
        Next r
    
            'PRINT MAX SUMMARY VALUES
            SummaryRow = 1
            Range("O" & SummaryRow).Value = "Ticker"
            Range("P" & SummaryRow).Value = "Value"
            
            SummaryRow = SummaryRow + 1
            Range("N" & SummaryRow).Value = "Greatest % Increase:"
            Range("O" & SummaryRow).Value = MaxPctIncTicker
            Range("P" & SummaryRow).Value = MaxPctIncValue
            Range("P" & SummaryRow).NumberFormat = "0.00%"
            
            SummaryRow = SummaryRow + 1
            Range("N" & SummaryRow).Value = "Greatest % Decrease:"
            Range("O" & SummaryRow).Value = MaxPctDecTicker
            Range("P" & SummaryRow).Value = MaxPctDecValue
            Range("P" & SummaryRow).NumberFormat = "0.00%"
            
            SummaryRow = SummaryRow + 1
            Range("N" & SummaryRow).Value = "Greatest Total Volume:"
            Range("O" & SummaryRow).Value = MaxTotVolTicker
            Range("P" & SummaryRow).Value = MaxTotVolValue
            
            ws.UsedRange.Columns.AutoFit
    Next ws

End Sub

