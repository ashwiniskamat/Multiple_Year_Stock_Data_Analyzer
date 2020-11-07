Sub Stock_Market():
    'Declare Variables
    Dim totalstockvolume As Double
    Dim ws As Worksheet
    Dim I As Double
    Dim tickervalue As Integer
    Dim ticker As String
    Dim GVT As String, GIT As String, GDT As String
    Dim GVV As Double, GIV As Double, GDV As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim stockvolume As Double
    Dim lowvalue As Double
    Dim highvalue As Double
    Dim change As Boolean
    
    'Process each worksheet
    For Each ws In Worksheets
        'Sort the columns
        With ws.Sort
         .SortFields.Add Key:=Range("A1"), Order:=xlAscending
         .SortFields.Add Key:=Range("B1"), Order:=xlAscending
         .SetRange Range("A:G")
         .Header = xlYes
         .Apply
        End With
        
        'Create Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest Percent Increase"
        ws.Cells(3, 14).Value = "Greatest Percent Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'Initialize Values
        Count = 2
    
        totalstockvolume = 0
        tickervalue = 1
        ticker = ""
        GVT = ""
        GIT = ""
        GDT = ""
        GVV = 0
        GIV = 0
        GDV = 0
        yearlychange = 0
        percentchange = 0
        stockvolume = 0
        lowvalue = 0
        highvalue = 0
        change = True
    
        'Find the last row
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
        'Process each row
        For I = 2 To RowCount
            'Aggregate the total stock volume for the current ticker
            totalstockvolume = totalstockvolume + ws.Cells(I, 7).Value
            'Change when ticker changes
            If change = True Then
                lowvalue = ws.Cells(I, 3).Value
            End If
            'Reset the change so ticker change can be used when ticker change is detected
            change = False
            'If ticker changes or year changes, then store the highvalue, total stock volume, yearly change, and percent change
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Or Left(ws.Cells(I, 1).Value, 4) <> Left(ws.Cells(I + 1, 1).Value, 4) Then
                highvalue = ws.Cells(I, 6).Value
                
                tickervalue = tickervalue + 1
                ws.Cells(tickervalue, 12).Value = totalstockvolume
                ws.Cells(tickervalue, 9).Value = ws.Cells(I, 1).Value
                ws.Cells(tickervalue, 10).Value = highvalue - lowvalue
                If ws.Cells(tickervalue, 10).Value < 0 Then
                    ws.Cells(tickervalue, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(tickervalue, 10).Interior.Color = RGB(0, 255, 0)
                End If
                'Check if the low value is 0 and prevent division by 0
                If lowvalue <> 0 Then
                    ws.Cells(tickervalue, 11).Value = ((highvalue - lowvalue) / (lowvalue)) * 100
                Else
                    ws.Cells(tickervalue, 11).Value = 0
                End If
                
                'Check if percent change is greater than GIV, then store percentage in GIV
                If GIV < ws.Cells(tickervalue, 11).Value Then
                   GIV = ws.Cells(tickervalue, 11).Value
                   GIT = ws.Cells(I, 1).Value
                End If
                'Check if percent change is less than GDV, then store percentage in GDV
                If GDV > ws.Cells(tickervalue, 11).Value Then
                   GDV = ws.Cells(tickervalue, 11).Value
                   GDT = ws.Cells(I, 1).Value
                End If
                'Check if total stock volume is greater than GVV, then store total stock value in GVV
                If GVV < totalstockvolume Then
                   GVV = ws.Cells(tickervalue, 12).Value
                   GVT = ws.Cells(I, 1).Value
                End If
                totalstockvolume = 0
                'Ticker change or year change is detected
                change = True
            End If
        Next I
        'Populating the summary table
        ws.Cells(2, 15).Value = GIT
        ws.Cells(2, 16).Value = GIV
        ws.Cells(3, 15).Value = GDT
        ws.Cells(3, 16).Value = GDV
        ws.Cells(4, 15).Value = GVT
        ws.Cells(4, 16).Value = GVV
    Next ws
    
End Sub