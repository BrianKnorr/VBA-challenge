Attribute VB_Name = "Module1"

'Cell(row,column)'
'general notes'
'1. open sub'
'2. define variables'
'3. Logic (Large to smaller)'
    '3.a. Open worksheets(open first loop)'
    '3.b. create index table headers'
    '3.c. go through ticker data (open second loop)'
        '3.c.I find first ticker variable and record in index'
        '3.c.II calculate all needed info (Changes and total vol)'
        '3.c.III conditional formatting'
    '3.d. loop to end of ticker column (column 1) for each ticker variable'
        '3.d.I for every new ticker variable record and repeat 3.c (close second loop)'
    '3.e. loop all of above 3 through all worksheet (close first loop)'
'4. end sub'

Sub stonks()
    'define all values to run through in each sheet (opening 1st loop)'
    For Each ws In Worksheets
        'define worksheets'
        Dim Current_Worksheet As String
        'i starts top of ticker column (starts at i = 2)'
        Dim i As Long
        'j starts the index (starts beneath the headers)'
        Dim j As Long
        'Index counter to fill Ticker row
        Dim Ticker As Long
        'Last row for given data'
        Dim Last_Data As Long
        'last row for Index'
        Dim Last_Index As Long
        'Variable for total volume'
        Dim Volume  As Double
        'variable for percengt change'
        Dim Percent_Change As Double
        'pull worksheet'
        Current_Worksheet = ws.CW
        
        'create headers for results'
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Volume"
        'define starting positions for loops'
        'starting position for ticker is row 2'
        Ticker = 2
        'Index values start at row 2'
        j = 2
        'equation for last row in data (defining/finding last row) (don't have to use equation for range)'
        Last_Data = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'set vtotal volume to 0 for first ticker calc'
        Volume = 0
            'running through data to record and calculate (open 2nd loop) all calcs being done here'
            For i = 2 To Last_Data
                'find first ticker value and determine next for looping'
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'record ticker in index table'
                ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
                'record change in index table'
                ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                'Conditional formating for change. if positive then grees else red)
                If ws.Cells(Ticker, 10).Value < 0 Then
                ws.Cells(Ticker, 10).Interior.ColorIndex = 30
                Else
                ws.Cells(Ticker, 10).Interior.ColorIndex = 50
                End If
                'calculating percent change'
                If ws.Cells(j, 3).Value <> 0 Then
                Percent_Change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                ws.Cells(Ticker, 11).Value = Format(Percent_Change, "Percent")
                Else
                ws.Cells(Ticker, 11).Value = Format(0, "Percent")
                End If
                'calculating sum of all tickers'
                ws.Cells(Volume, 12).Value = Volume + Cells(j, 7).Value
                'next ticker in index'
                Ticker = Ticker + 1
                'next'
                j = j + 1
                End If
                'reset volume for new ticker'
                Volume = 0
            Next i
        Next ws
End Sub

