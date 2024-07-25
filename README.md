# VBA-challenge
VBA challenge for KU Bootcamp with Caitlin Moran


'Create a script that loops through all the stocks for each quarter and outputs the following information:
Sub stock()
Dim ws As Worksheet
For Each ws In Worksheets
    With ws.Columns("I:L") 'adam note: base code found here and modified. https://learn.microsoft.com/en-us/office/vba/api/excel.range.columnwidth
        .ColumnWidth = .ColumnWidth * 2
    End With
    With ws.Columns("n:p")
        .ColumnWidth = .ColumnWidth * 3
    End With
    Range("i1") = "Ticker"
    ws.Range("j1") = "Quarterly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    ws.Range("n2") = "Greatest % Increase"
    ws.Range("n3") = "Greatest % Decrease"
    ws.Range("n4") = "Greatest Total Volume"
    ws.Range("o1") = "Ticker"
    ws.Range("p1") = "Value"
    Dim openv As Double
    openv = ws.Cells(2, 3) 'adam note: set initial open state
    Dim percent As Double
    percent = 0
    Dim counter As LongLong
    counter = 0
    Dim ticker As String
    Dim summary_row As String
    summary_row = 1 'adam note: set initial summary row state
    For i = 2 To 100000
        If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then 'adam note: if same ticker, just add to counter
            counter = counter + ws.Cells(i, 7)
        ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then 'adam note: if different ticker, close minus open, then display everything, then set next openv and reset counter
            summary_row = summary_row + 1
'The ticker symbol
            ticker = ws.Cells(i, 1)
            ws.Cells(summary_row, 9) = ticker
'Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
            value = ws.Cells(i, 6) - openv
            ws.Cells(summary_row, 10) = value
            If ws.Cells(summary_row, 10).value < 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(summary_row, 10).value > 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            End If
'The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
            If ws.Cells(summary_row, 10) <> 0 Then 'adam note: will get divide by 0 unless you check for this
                percent = ws.Cells(summary_row, 10) / ws.Cells(i, 6)
                ws.Cells(summary_row, 11) = percent
            Else
                ws.Cells(summary_row, 11) = 0
            End If
'The total stock volume of the stock. The result should match the following image:
            ws.Cells(summary_row, 12) = counter
            openv = ws.Cells(i + 1, 3)
            counter = 0
        End If
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
    Next i
        'adam note: per https://stackoverflow.com/questions/36165887/run-time-error-1004-select-method-of-range-class-failed-using-thisworkbook and https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.activate(method)
        ws.Activate
        'adam note: from https://learn.microsoft.com/en-us/office/vba/api/excel.worksheetfunction.max
        Dim gper As Double
        Dim gdec As Double
        Dim gtv As LongLong
        gper = WorksheetFunction.Max(ws.Range("K2:K1000"))
        ws.Range("p2") = gper
        ws.Range("o2") = Cell
        gdec = WorksheetFunction.Min(ws.Range("K2:K1000"))
        ws.Range("p3") = gdec
        gtv = WorksheetFunction.Max(ws.Range("l2:l1000"))
        ws.Range("p4") = gtv
        'adam note: from https://stackoverflow.com/questions/41008736/how-to-get-cell-address-from-find-function-in-excel-vba
        Dim GPA As Range
        ws.Columns("K:K").Select
        Set GPA = Selection.Find(gper)
        'adam note: from https://stackoverflow.com/questions/35617755/excel-macro-cell-address-increase
        ws.Cells(2, 15) = ws.Range(GPA.Address).Offset(0, -2)
        Dim GDA As Range
        ws.Columns("K:K").Select
        Set GDA = Selection.Find(gdec)
        ws.Cells(3, 15) = ws.Range(GDA.Address).Offset(0, -2)
        Dim GTA As Range
        ws.Columns("L:L").Select
        Set GTA = Selection.Find(gtv)
        ws.Cells(4, 15) = ws.Range(GTA.Address).Offset(0, -3)
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Range("p2").NumberFormat = "0.00%"
        ws.Range("p3").NumberFormat = "0.00%"
'Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.
    percent = 0
    counter = 0
    summary_row = 1
Next ws
Worksheets(1).Activate
End Sub
