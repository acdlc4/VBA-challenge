Attribute VB_Name = "QtrlyStockAnalysis"
Option Explicit

Sub QtrlyStockAnalysis()


'PURPOSE: Analyze a list of daily stock prices (sorted in order by ticker and date ascending)
'         for quarterly changes in price, percentage change, and total shares traded


Dim number_sheets As Integer
Dim number_rows As Long
Dim X As Integer

'Set count of sheets in active workbook
number_sheets = ThisWorkbook.Sheets.Count

'External loop to apply process to each and every worksheet in active workbook

For X = 1 To number_sheets

'Set column headers for analysis
Worksheets(X).Select
Range("A1").Select
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Quarterly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Count of number of rows in daily stock listings in current worksheet
number_rows = Cells(Rows.Count, 1).End(xlUp).Row

'Ticker summarization into Col. I for current worksheet

Dim A As Long
Dim B As Integer
Dim ticker As String
Dim LastRow As Long
Dim BegQtrPrice As Currency
Dim EndQtrPrice As Currency
Dim QtrChg As Currency
Dim PercentChg As Variant
Dim TotVol As LongLong

'Set beginning row value for summary analysis section
B = 2

'Set Qtr Beginning Price for first ticker
BegQtrPrice = Cells(B, 3).Value

'Loop to process through all different tickers
For A = 2 To number_rows
    
   If Cells(A + 1, 1).Value <> Cells(A, 1).Value Then
   
        'Set and report stock ticker
        ticker = Cells(A, 1).Value
        Cells(B, 9).Value = ticker
        
        'Set and report Total Stock Vol by ticker
        TotVol = TotVol + Cells(A, 7).Value
        Cells(B, 12).Value = TotVol
        
        'Set and report Qtr Chg and Pct Chg by ticker
        EndQtrPrice = Cells(A, 6).Value
        QtrChg = EndQtrPrice - BegQtrPrice
        Cells(B, 10).Value = QtrChg
                
        PercentChg = WorksheetFunction.Round(QtrChg / BegQtrPrice, 4)
        Cells(B, 11).Value = PercentChg
        
        'Reset beginning values for next ticker
        BegQtrPrice = Cells(A + 1, 3).Value
        TotVol = 0
        
        'Set next row value for next ticker in summary analysis section
        B = B + 1
        
    Else
        'Accumulation of Total Stock Volume for each unique ticker, looped
        TotVol = TotVol + Cells(A, 7).Value
     
    End If
    
Next A
    

'Loop for Conditional formatting of Column J
Dim C As Long

'Set count of number of rows in analysis section
LastRow = Cells(Rows.Count, 9).End(xlUp).Row

    For C = 2 To LastRow
       
       'No formatting when cell value is exactly ZERO
       If Cells(C, 10) = 0 Then
    
        'Set Cell Color to GREEN if POSITIVE
        ElseIf Cells(C, 10).Value > 0 Then
           Cells(C, 10).Interior.ColorIndex = 4
          
        'Set Cell Color to RED if NEGATIVE
        Else
           Cells(C, 10).Interior.ColorIndex = 3
        
       End If
        
    Next C
     
    'Format Qtrly Chg Col
    Range("J2:J" & LastRow).NumberFormat = "0.00"
    
    'Format PctChg Col
    Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
    'Format TotVol Col
    Range("L2:L" & LastRow).Style = "Comma"
    Range("L2:L" & LastRow).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        
   'AutoFit Columns I through L and reset selected cell to A2
    Columns("I:L").EntireColumn.AutoFit
    Range("A2").Select
     
'Create GREATEST section for each worksheet

    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

Dim GrtPerInc As Variant
Dim GrtPerDec As Variant
Dim GrtTotVol As LongLong
Dim PerChgCol As Range
Dim TotVolCol As Range
Dim TickSymCol As Range

Set TickSymCol = Range(Cells(2, 9), Cells(LastRow, 9))
Set PerChgCol = Range(Cells(2, 11), Cells(LastRow, 11))
Set TotVolCol = Range(Cells(2, 12), Cells(LastRow, 12))

GrtPerInc = WorksheetFunction.Max(PerChgCol)
GrtPerDec = WorksheetFunction.Min(PerChgCol)
GrtTotVol = WorksheetFunction.Max(TotVolCol)

'Set above Dim values

Cells(2, 17).Value = GrtPerInc
Cells(3, 17).Value = GrtPerDec
Cells(4, 17).Value = GrtTotVol

'Lookup and place Tickers for above Dim values
Cells(2, 16).Value = WorksheetFunction.Index(TickSymCol, WorksheetFunction.Match(GrtPerInc, PerChgCol, 0))
Cells(3, 16).Value = WorksheetFunction.Index(TickSymCol, WorksheetFunction.Match(GrtPerDec, PerChgCol, 0))
Cells(4, 16).Value = WorksheetFunction.Index(TickSymCol, WorksheetFunction.Match(GrtTotVol, TotVolCol, 0))


'FORMAT HIGHLIGHTS Chart
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    Range("Q4").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    
    Columns("O:Q").EntireColumn.AutoFit
    Columns("M:N").ColumnWidth = 4
     
    Range("A2").Select


'RESET Dims for next worksheet

B = 0
A = 0
number_rows = 0
LastRow = 0

Range("A2").Select
Range("A1").Select

Next X

'Reset view by selecting first worksheet
Worksheets(1).Select

End Sub













