Sub Hard_Ticker_2()
    'variable dims
    Dim years As Double, l As Double, counter As Double, totals As Double, newtick As Double, apple As Double
    Dim start_vol As Double, end_vol As Double, pmin As Double, pmax As Double, vol_maxl As Double
    Dim pmin_count As Double, pmax_count As Double, vol_count As Double
    Dim ticker As String
    Dim ws As Worksheet
    Dim vValues As Variant, pValues As Variant

   'get this puppy running as quickly as possible
    Application.ScreenUpdating = False
    
    'Start with looping through the worksheets
    For Each ws In Worksheets
    
        'let's sort the damn worksheet cuz ya never know if things will be in order
        'This will sort by the first column and then the second column, so by ticker and then date within each ticker
        ws.Sort.SortFields.Add Key:=Range("A:A"), Order:=xlAscending
        ws.Sort.SortFields.Add Key:=Range("B:B"), Order:=xlAscending
        ws.Sort.SetRange Range("A:G")
        ws.Sort.Header = xlYes
        ws.Sort.Apply
        
        'Figure out how many iterations do we need to do in the For-Loop
        l = ws.Cells(Rows.Count, 2).End(xlUp).Row
        
        'Lets format the space where we'll be putting the results
        ws.Cells(1, 9) = "Ticker in " & Left(ws.Cells(2, 2), 4)
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(2, 15) = "Greatest % Increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Range("K2:K" & l).NumberFormat = "0.00%"
        ws.Columns("O").ColumnWidth = 19.2
        ws.Columns("Q").ColumnWidth = 11.5
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Need to put some start values
        ws.Cells(2, 9) = ws.Cells(2, 1)
        'totals will track the volume of each ticker throughout the loop
        totals = ws.Cells(2, 7)
        'Grabs the price at the year's opening
        start_vol = ws.Cells(2, 3)
        'Tracks how many tickers we have, but is really used to index where things are being placed
        counter = 1
        
        'Start looping through rows upon rows
        For i = 3 To l
            'if the ticker name matches the current ticker, then we just need to add volume 
            If ws.Cells(i, 1) = ws.Cells(counter + 1, 9) Then
                totals = totals + ws.Cells(i, 7)
            Else:
            'but if the ticker doesn't match the current ticker, we need to reset some things
                ws.Cells(counter + 2, 9) = ws.Cells(i, 1)
                ws.Cells(counter + 1, 12) = totals
                'Gotta start back at 0
                totals = 0
                'Grabs the price at the year's closing
                end_vol = ws.Cells(i - 1, 6)
                'Calculate the difference between the start and end of the year
                ws.Cells(counter + 1, 10) = end_vol - start_vol
                
                'Let's make things pretty based on how well the ticker did
                If ws.Cells(counter + 1, 10) > 0 Then
                    ws.Cells(counter + 1, 10).Interior.Color = RGB(46, 139, 87)
                Else: ws.Cells(counter + 1, 10).Interior.Color = RGB(220, 20, 60)
                End If
                
                'Calculate % change. BUT if the start vol = 0 (as in a new ticker?), then we don't do this cuz dividing by zero is bad.
                If start_vol > 0 Then
                    ws.Cells(counter + 1, 11) = (end_vol - start_vol) / start_vol
                End If
                'We have a new opening price.
                start_vol = ws.Cells(i, 3)
                'We have another ticker! note that.
                counter = counter + 1
                
            End If
            
        Next i
        'Finish the stats on the final ticker. Nothing really new here.
        ws.Cells(counter + 1, 12) = totals
        end_vol = ws.Cells(i - 1, 6)
        ws.Cells(counter + 1, 10) = end_vol - start_vol
       
        If ws.Cells(counter + 1, 10) > 0 Then
            ws.Cells(counter + 1, 10).Interior.Color = RGB(46, 139, 87)
        Else: ws.Cells(counter + 1, 10).Interior.Color = RGB(220, 20, 60)
        End If
            
        If start_vol > 0 Then
            ws.Cells(counter + 1, 11) = (end_vol - start_vol) / start_vol
        End If
        
        'Find the max and min percentages by going through the columns that have already been placed.
        pValues = ws.Range("K2:K" & counter + 1)
        pmin = WorksheetFunction.Min(pValues)
        pmax = ws.Application.WorksheetFunction.Max(pValues)
        pmin_count = ws.Application.WorksheetFunction.Match(pmin, pValues, 0)
        pmax_count = ws.Application.WorksheetFunction.Match(pmax, pValues, 0)
        
        'Now find the most volumous ticker in the same way as above.
        vValues = ws.Range("L2:L" & counter + 1)
        vol_max = ws.Application.WorksheetFunction.Max(vValues)
        vol_count = ws.Application.WorksheetFunction.Match(vol_max, vValues, 0)

        'Fill in the tickers.
        ws.Cells(2, 16) = ws.Range("I" & pmax_count + 1)
        ws.Cells(3, 16) = ws.Range("I" & pmin_count + 1)
        ws.Cells(4, 16) = ws.Range("I" & vol_count + 1)
        
        'Fill in the extreme values.
        ws.Cells(2, 17) = pmax
        ws.Cells(3, 17) = pmin
        ws.Cells(4, 17) = vol_max
        
    Next ws
    
    Application.ScreenUpdating = True
    'Altert yourself that this program is done processing. 
    'This was more necessary when the code took hours to run.
    'Seems to have become a bit more efficient now. 
    'But it's still nice to be notified that your program is done and you can see how much you win or fail :-)
    MsgBox ("DONE!" & vbcrlf & "Double check the numbers..." & vbcrlf & "...or go get a drink" )
    
End Sub

