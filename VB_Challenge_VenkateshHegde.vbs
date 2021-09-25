Attribute VB_Name = "Module1"
' Steps:
' ----------------------------------------------------------------------------

' Part I:

' 1. Loop through Column A and track changes to stock ticker
' When stock ticker has changed then store the new ticker symbol in the array Tickers
' Within a stock track changes in the price and store the opening and closing price
' and store in arrays TickerOpens and TickerCloses
' Track the yearly change and percentage from opening an closing in  arrays
' Track total stock volume in arrays
' 2. Add the State to the first column of each spreadsheet.
' 3. Convert the headers of each row to simply say the year.
' 4. Convert the cells to currency format

Sub ticker_extract()
Dim FirstRow As Long
Dim LastRow As Long
Dim FirstCol As Long
Dim LastCol As Long
Dim CurrentTicker As String
Dim LastTicker As String
Dim TickerOpen As Double
Dim TickerClose As Double
Dim TickerMin As Double
Dim TickerMax As Double
Dim TickerVol As LongLong
Dim MaxPerIncrTicker As String
Dim MaxPerIncrease As Double
Dim MaxPerDecrTicker As String
Dim MaxPerDecrease As Double
Dim MaxVol As LongLong
Dim MaxVolTicker As String

Dim Tickers(5000) As String
Dim TickerOpens(5000) As Double
Dim TickerCloses(5000) As Double
Dim TickerMins(5000) As Double
Dim TickerMaxs(5000) As Double
Dim TickerChanges(5000) As Double
Dim TickerPerChanges(5000) As Double
Dim TickerVols(5000) As LongLong
Dim TickerNumber As Integer

Dim TickerChange As Double
Dim WeAreDone As Boolean

MaxPerIncrTicker = ""
MaxPerIncrease = -9999
MaxPerDecrTicker = ""
MaxPerDecrease = 9999
MaxVol = 0
MaxVolTicker = ""
TickerNumber = 0


' Create the output sheet by adding it to the beginning
'On Error Resume Next
'Application.DisplayAlerts = False
'Sheets("Output").Delete
'Sheets.Add.Name = "Output"
'move created sheet to be first sheet
'On Error GoTo 0
'Application.DisplayAlerts = True
'Sheets("Output").Move Before:=Sheets(1)
' Specify the location of the combined sheet
'Set Output_sheet = Worksheets("Output")

For Each ws In Worksheets
    If ws.Name <> "Output" And ws.Range("A1") = "<ticker>" Then
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(Columns.Count, 1).End(xlToLeft).Column
        ' Clean out output area
        ws.Range("L1:T" & LastRow).Clear
    ' Initialize values
        CurrentTicker = ""
        LastTicker = ""
        TickerOpen = 0
        TickerClose = 0
        TickerMin = 9999999
        TickerMax = 0
        TickerVol = 0
        
        MaxPerIncrTicker = ""
        MaxPerIncrease = -9999
        MaxPerDecrTicker = ""
        MaxPerDecrease = 9999
        MaxVol = 0
        MaxVolTicker = ""
        TickerNumber = 0
        
        
        WeAreDone = False
        Debug.Print ("Initialized Values")
        
        ' --------------------------------------------
        ' LOOP THROUGH ALL ROWS
        ' --------------------------------------------
        For Each CellVal In ws.Range("A2:A" & LastRow)
        
            'Debug.Print ("Cell value " & CellVal.Value)
            ' If we have reached a cell without any data in ticker then we are done
            ' If we did process any records then we need to collect close and other values
            If CellVal.Value = "" Then
                WeAreDone = True
                'If there are no records to process then we exit the loop without doing any work
                If TickerNumber = 0 Then
                    Exit For
                End If
            End If
            
    '        Debug.Print ("Starting ticker " & TickerNumber)
            
            If CellVal.Value <> CurrentTicker Then
            'And CurrentTicker <> "" Then
                ' Check if this is not the very first ticker first
                '  then since the ticker has changed store all values like close for the last ticker
                If CurrentTicker <> "" Then
                    Debug.Print ("Store values for old ticker " & CurrentTicker)
                    Tickers(TickerNumber) = CurrentTicker
                    TickerCloses(TickerNumber) = TickerClose
                    TickerMins(TickerNumber) = TickerMin
                    TickerMaxs(TickerNumber) = TickerMax
                    TickerVols(TickerNumber) = TickerVol
                    ' Check the MaxChanges for last ticker and determine if last ticker qualifies
                    ' In any of the max values for increase decrease or vol
                    ' If this is first ticker then do not calculate these
                    If TickerOpen <> 0 Then
                        If (TickerClose - TickerOpen) / TickerOpen > MaxPerIncrease Then
                            MaxPerIncrease = (TickerClose - TickerOpen) / TickerOpen
                            MaxPerIncrTicker = CurrentTicker
                        End If
                    
                        If (TickerClose - TickerOpen) / TickerOpen < MaxPerDecrease Then
                            MaxPerDecrease = (TickerClose - TickerOpen) / TickerOpen
                            MaxPerDecrTicker = CurrentTicker
                        End If
                    End If
                    If TickerVol > MaxVol Then
                        MaxVol = TickerVol
                        MaxVolTicker = CurrentTicker
                    End If
                    Debug.Print (CurrentTicker & " - " & TickerOpen & " - " & TickerClose & " - " & TickerVol)
                    TickerNumber = TickerNumber + 1
    
                Else
                    Debug.Print ("First ticker ")
    
                End If
                ' If we have gone beyond last row then we are done
                If WeAreDone Then
                    Debug.Print ("Finished last row")
                    Exit For
                End If
                ' Store values for new ticker such as open and initialize values like vol
                TickerOpen = ws.Range("C" & CellVal.Row).Value
                Tickers(TickerNumber) = CellVal.Value
                TickerOpens(TickerNumber) = TickerOpen
                LastTicker = CurrentTicker
                CurrentTicker = CellVal.Value
                TickerMin = ws.Range("E" & CellVal.Row).Value
                TickerMax = ws.Range("D" & CellVal.Row).Value
                TickerVol = 0
                Debug.Print ("Stored values for new ticker " & CurrentTicker)
                Debug.Print (CurrentTicker & " - " & TickerOpen & " - " & TickerClose & " - " & TickerVol)
'                If TickerNumber = 0 Then
'                    TickerNumber = TickerNumber + 1
'                End If
    
    
            End If
           ' for each record being processed for a ticker we need to store min and max
           ' We also need to store the values
            TickerClose = ws.Range("F" & CellVal.Row).Value
            If TickerMin > ws.Range("E" & CellVal.Row).Value Then
                TickerMin = ws.Range("E" & CellVal.Row).Value
            End If
                
            If TickerMax < ws.Range("D" & CellVal.Row).Value Then
                TickerMin = ws.Range("D" & CellVal.Row).Value
            End If
            TickerVol = TickerVol + ws.Range("G" & CellVal.Row).Value
    '        Debug.Print (CurrentTicker & " - " & TickerOpen & " - " & TickerClose & " - " & TickerVol)
'            If CellVal.Row > 900 Then
'                Exit For
'            End If
        Next CellVal
        'we finished last row and we need to store the last values
        Debug.Print ("Store values for old ticker " & CurrentTicker)
        Tickers(TickerNumber) = CurrentTicker
        TickerCloses(TickerNumber) = TickerClose
        TickerMins(TickerNumber) = TickerMin
        TickerMaxs(TickerNumber) = TickerMax
        TickerVols(TickerNumber) = TickerVol
        ' Check the MaxChanges for last ticker and determine if last ticker qualifies
        ' In any of the max values for increase decrease or vol
        ' If this is first ticker then do not calculate these
        If TickerOpen <> 0 Then
            If (TickerClose - TickerOpen) / TickerOpen > MaxPerIncrease Then
                MaxPerIncrease = (TickerClose - TickerOpen) / TickerOpen
                MaxPerIncrTicker = CurrentTicker
            End If
        
            If (TickerClose - TickerOpen) / TickerOpen < MaxPerDecrease Then
                MaxPerDecrease = (TickerClose - TickerOpen) / TickerOpen
                MaxPerDecrTicker = CurrentTicker
            End If
        End If
        
        If TickerVol > MaxVol Then
            MaxVol = TickerVol
            MaxVolTicker = CurrentTicker
        End If
        TickerNumber = TickerNumber + 1
        
    '    If TickerNumber > 80 Then
    '        Exit For
    '    End If
    
    
    Else
        ws.Cells.Clear
    End If


    Debug.Print (" Printing all tickers")
    
    ws.Range("L1") = "Ticker Name"
    ws.Range("M1") = "Open"
    ws.Range("N1") = "Close"
    ws.Range("O1") = "Volume"
    ws.Range("P1") = "Min"
    ws.Range("Q1") = "Max"
    ws.Range("R1") = "Change"
    ws.Range("S1") = "%Change"
    
    
    For i = 0 To TickerNumber - 1
        Debug.Print (Tickers(i) & " - " & TickerOpens(i) & " - " & TickerCloses(i) & _
        " - " & TickerMaxs(i) & " - " & TickerMins(i) & " - " & TickerVols(i))
            ws.Range("L" & i + 2) = Tickers(i)
            ws.Range("M" & i + 2) = TickerOpens(i)
            ws.Range("N" & i + 2) = TickerCloses(i)
            ws.Range("O" & i + 2) = TickerVols(i)
            ws.Range("P" & i + 2) = TickerMins(i)
            ws.Range("Q" & i + 2) = TickerMaxs(i)
            ws.Range("R" & i + 2) = TickerCloses(i) - TickerOpens(i)
            
            If TickerOpens(i) <> 0 Then
                ws.Range("S" & i + 2) = (TickerCloses(i) - TickerOpens(i)) / TickerOpens(i)
            End If
            If TickerCloses(i) > TickerOpens(i) Then
                ws.Range("R" & i + 2).Interior.ColorIndex = 4
                ws.Range("S" & i + 2).Interior.ColorIndex = 4
            Else
                ws.Range("R" & i + 2).Interior.ColorIndex = 3
                ws.Range("S" & i + 2).Interior.ColorIndex = 3
            End If
            ws.Range("S" & i + 2).Value = FormatPercent(ws.Range("S" & i + 2).Value)
                    
            
    Next i
    
    ' Outputting maximums across all tickerrs
    ws.Range("V1") = "Ticker"
    ws.Range("W1") = "Value"
    
    
    ws.Range("U2") = "Ticker with Max Increase%"
    ws.Range("V2") = MaxPerIncrTicker
    ws.Range("W2") = FormatPercent(MaxPerIncrease)
    
    
    ws.Range("U3") = "Ticker with Max Decrease%"
    ws.Range("V3") = MaxPerDecrTicker
    ws.Range("W3") = FormatPercent(MaxPerDecrease)
    
    
    ws.Range("U4") = "Ticker with Max Volume"
    ws.Range("V4") = MaxVolTicker
    ws.Range("W4") = MaxVol
    
Next ws


End Sub






