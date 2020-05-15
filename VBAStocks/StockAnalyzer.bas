Attribute VB_Name = "Module1"
'****************************************************************
Sub StockAnalyzer():
'*****************************************************************
'This Module scans all the Worksheets in a Workbook to find
'(a) Yearly Change of stocks
'(b) Percentage Change of stocks  &
'(c) Total Stock Volume of stocks arranged by their ticker symbols
'(d) Exhibits the stocks that achieved (i) Greatest % Increase
'       (ii) Greatest % decrease & (iii) Greatest Total Volume
'*******************************************************************
'Define ws as Worksheet variable
Dim ws As Worksheet
'Iterate through all the worksheets in this workbook
For Each ws In ThisWorkbook.Worksheets
    'Activate the current worksheet
    ws.Activate
    'Turn off screen refresh and recalculating workbook's formulas
    'which will help the macro to run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    '----------------------------------------------
    ' Define the variables needed in this module
    '-----------------------------------------------
    Dim LastRow As Long 'Last row of data sheet
    Dim LastCol As Long 'Last column of data sheet
    Dim lRow As Long 'Last row of summary sheet
    Dim Start_row As Long 'Starting row of the data sheet
    Dim Next_row As Long 'Next row of the data sheet
    Dim counter As Long 'counter, Row & tck are counter variables
    Dim Row As Long
    Dim tck As Long
    Dim OpenPrice As Double 'Opening price of a stock
    Dim ClosePrice As Double 'Closing price of a stock
    Dim YearlyChange As Double 'Stock price change over an year
    Dim PercentChange As Double 'Percentage change over an year
    Dim TotalStock As Double 'Total stock volume over an year
    Dim GreatestInc As Double 'Greatest % increase value
    Dim tckInc As Long
    Dim GreatestDec As Double 'Greatest % decrease value
    Dim tckDec As Long
    Dim GreatestVol As Double 'Greatest stock volume value
    Dim tckVol As Long
    '----------------------------------------------------------
    'Getting the last row and last column numbers
    '----------------------------------------------------------
    'Getting the row number of last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    'Getting column number of last column
    LastCol = Cells(1, Columns.Count).End(xlToLeft).Column
    '------------------------------------------------------------
    'Initializing the headers for summary columns
    '------------------------------------------------------------
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Columns("I:L").EntireColumn.AutoFit
    '--------------------------------------------------------------
    'Initializing the variables before iteration
    '---------------------------------------------------------------
    Start_row = 2 'Starting row of the data set
    'Total stock value for the 1st stock
    TotalStock = Cells(Start_row, 7).Value
    'Open price value for the 1st stock
    OpenPrice = Range("C" & Start_row).Value
    tck = 0
    counter = 1
    '-------------------------------------------------------------
    ' Create a For loop to find unique sticker symbols
    '-------------------------------------------------------------
    For Row = Start_row To LastRow
       Next_row = Row + 1 'Next row is current row + 1
         If (Cells(Row, 1).Value = Cells(Next_row, 1).Value) Then
            tck = tck + 1 'Increment ticker index
            'Cumulate stock volume
            TotalStock = TotalStock + Cells(Next_row, 7).Value
         Else
            'tck is the number of unique ticker symbol in the sheet
            'Row is the last row (i.e. year-end) of that ticker
      '------------------------------------------------------------
      ' Calculate Yearly change, Percentage change & Total Volume
      '------------------------------------------------------------
           'Calculate the ClosePrice
            ClosePrice = Range("F" & Row).Value
            'Calculate yearly change of the stock price
            YearlyChange = ClosePrice - OpenPrice
            'Calculate percent change of the stock price
            'Select cases only when OpenPrice is not zero
            If (OpenPrice <> 0) Then
                PercentChange = YearlyChange / OpenPrice
            Else
                PercentChange = 0
            End If
        '----------------------------------------------------------
        ' Populate the cells in the Summary Table with above values
        '----------------------------------------------------------
            'Increment Row counter for Summary table
            counter = counter + 1
            'Write the summary values in the Summary table
            Range("I" & counter).Value = Range("A" & Row).Value
            Range("J" & counter).Value = Format(YearlyChange, "0.00")
            '------------------------------------------------------------
            'Color the Cells as Red if the change is negative, else Green
            '------------------------------------------------------------
                If (YearlyChange < 0) Then
                    Range("J" & counter).Interior.ColorIndex = 3
                Else
                    Range("J" & counter).Interior.ColorIndex = 4
                End If
            Range("K" & counter).Value = Format(PercentChange, "Percent")
            Range("L" & counter).Value = Format(TotalStock, "0")
        '---------------------------------------------------------------
        ' Continue with the For loop and find the other stocks
        '---------------------------------------------------------------
            'Reset ticker index for the next stock
            tck = 0
            'Update OpenPrice for the next stock
            OpenPrice = Range("C" & Next_row).Value
            'Update TotalStock for the next stock
            TotalStock = Range("G" & Next_row).Value
         End If
    Next Row
    '-----------------------------------------------------------------------
    ' Summary Table is complete, let's format it nicely
    '------------------------------------------------------------------------
    'Find the last row of the summary section
    lRow = Cells(Rows.Count, 9).End(xlUp).Row
    'Format the Summary Table using the last row value
    Range("I1:L" & lRow).Borders.LineStyle = xlContinuous
    Range("I1:L" & lRow).Borders.Weight = xlThin
    Range("I1:L" & lRow).Font.Name = "Arial Narrow"
    Range("I1:L1").Font.FontStyle = "Bold"
    Range("I1:L1").Font.Size = 12
    Range("I1:L1").Interior.ColorIndex = 44
    Range("I1:L1").HorizontalAlignment = xlCenter
    '-----------------------------------------------------------------------
    'Creating a new table for the "greatest" stocks
    'We shall name it as Greatest Table
    '-----------------------------------------------------------------------
    'Create the row labels
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    'Display the row labels
    Columns("N:P").EntireColumn.AutoFit
    'Initialize the Variables for the Greatest Table
    GreatestInc = 0
    GreatestDec = 0
    GreatestVol = 0
    'Display WorksheetName in the Greatest Table
    Range("N1") = "Sheet Name:" + " " + ws.Name
    '----------------------------------------------------------------------
    'Create a For loop to find the greatest values from the summary table
    '----------------------------------------------------------------------
    'Searching for Greatest Percent Increase
     For Row = Start_row To lRow
       If Cells(Row, 11).Value > GreatestInc Then
            GreatestInc = Cells(Row, 11).Value
            tckInc = Row
        End If
     Next Row
    'Fill the cell with greatest percentage increase
    Range("P2").Value = Format(GreatestInc, "Percent")
    'Fill the cell with corresponding ticker symbol
    Range("O2").Value = Range("I" & tckInc).Value
    'Searching for Greatest Percent Decrease
     For Row = Start_row To lRow
       If Cells(Row, 11).Value < GreatestDec Then
            GreatestDec = Cells(Row, 11).Value
            tckDec = Row
        End If
     Next Row
    'Fill the cell with greatest percentage decrease
    Range("P3").Value = Format(GreatestDec, "Percent")
    'Fill the cell with corresponding ticker symbol
    Range("O3").Value = Range("I" & tckDec).Value
    'Searching for Greatest Total Volume
     For Row = Start_row To lRow
       If Cells(Row, 12).Value > GreatestVol Then
            GreatestVol = Cells(Row, 12).Value
            tckVol = Row
        End If
     Next Row
    'Fill the cell with greatest total volume
    Range("P4").Value = Format(GreatestVol, "Scientific")
    'Fill the cell with corresponding ticker symbol
    Range("O4").Value = Range("I" & tckVol).Value
    '--------------------------------------------------------------
    'Greatest Table is complete, lets format it nicely
    '--------------------------------------------------------------
    Range("N1:P4").Borders.LineStyle = xlContinuous
    Range("N1:P4").Borders.Weight = xlThin
    Range("N1:P4").Font.Name = "Arial Narrow"
    Range("N1:P1").Font.FontStyle = "Bold"
    Range("N1:P1").Font.Size = 12
    Range("N1:P1").Interior.ColorIndex = 36
    Range("N1:P1").HorizontalAlignment = xlCenter
    
    'Turn on screen refresh and automatic calculation of formulas
    'before leaving this worksheet
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
'Let's go to the next worksheet
Next ws
End Sub
