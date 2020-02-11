Attribute VB_Name = "Module1"
Sub TickerYTDChange()
    
'   Created By: Karl Ramsay
'   Created:    Feb. 8th, 2020
'   Modified:   Feb. 10th, 2020
'   Modified:
'   Description: The purpose of this code is to process stock files for 2014, 2015 and 2016
'   For each year and Ticker we need to find the delta between the First Open Day Price and Last Day Price.
'   Post the results for the tickers that had the biggest Percentage Increase and Decrease
'   Additionally, we will post the ticker that had the largest Traded Stock Volume
'   For each of these metrics identified, we will post the results in a summary table in columns O, P and Q
            
    Dim lastCol As Double
    Dim lastRow As Double
    Dim f As Integer
    Dim rngSummaryTbl As Range
    Dim ProcessedList As String
     
    
    'initialize variables needed for the process log - audit report
    f = 1
    ProcessedList = "The following sheets have been processed: " & vbCrLf
       
    
    'Loop through sheets in the workbook and process each sheet
    For Each sht In Worksheets
        
        sht.Select
        
        'clear the range before setting headers for the answers
         sht.Range("I:Q").Clear
          
        'locate the boundaries for this dataset
         lastRow = sht.Range("A" & Rows.Count).End(xlUp).Row
         lastCol = sht.Cells(1, Columns.Count).End(xlToLeft).Column
          
        'run procedure to sort the data, just to be sure we have ticker and dates in the right order
        SortData sht, Range(Cells(1, 1), Cells(lastRow, lastCol))
        
        'Run procedure to find the Biggest Percentage Increase and Decrease
        FindFirstLastDayPricing sht
         
    
        'Add summary for Extreme Changes (Min / Max)
        sht.Range("O2").ColumnWidth = "25"
        sht.Range("O2").Value = "Greatest % Increase"
        sht.Range("O3").Value = "Greatest % Decrease"
        sht.Range("O4").Value = "Greatest Total Volume"
        sht.Range("P1").Value = "Ticker"
        sht.Range("Q1").Value = "Value"
        sht.Range("Q1").ColumnWidth = "16"
        
        'find the last row in the summary data section and define a range for the data summary table
        
        lastRow = sht.Range("I" & Rows.Count).End(xlUp).Row
        Set rngSummaryTbl = Range(sht.Cells(1, 9), sht.Cells(lastRow, 12))
        
        
        
        'set the values for the ticker showing the greatest percentage increase
        rngSummaryTbl.Sort sht.Range("K1").Value, xlDescending, , , , , , xlYes
          
        
        
        sht.Range("P2").Value = sht.Range("I2").Value
        sht.Range("Q2").Value = sht.Range("K2").Value
        sht.Range("Q2").NumberFormat = "0.00%"
        
        'set the values for the ticker showing the greatest percentage decrease
        rngSummaryTbl.Sort sht.Range("K1").Value, xlAscending, , , , , , xlYes
        
        
        
        sht.Range("P3").Value = sht.Range("I2").Value
        sht.Range("Q3").Value = sht.Range("K2").Value
        sht.Range("Q3").NumberFormat = "0.00%"
        
        
        
        'set the values for the ticker showing the greatest total volume
        rngSummaryTbl.Sort sht.Range("L1").Value, xlDescending, , , , , , xlYes
        
        
        
        sht.Range("P4").Value = sht.Range("I2").Value
        sht.Range("Q4").Value = sht.Range("L2").Value
           
        Cells(2, 9).Select
    
        'keep track of the sheets you processed
        ProcessedList = ProcessedList & f & ") " & sht.Name & vbCrLf
        
        f = f + 1
        
    Next sht
    
    MsgBox "Done!", vbOKOnly
    MsgBox ProcessedList, vbInformation, "Sheets Processed"
    
  
    
End Sub



Sub SortData(Optional ws As Variant, Optional rngSort As Range)

'   Created By: Karl Ramsay
'   Created:    Feb. 8th, 2020
'   Modified:   Feb. 10th, 2020
'   Modified:
'   Description: The purpose of this code is to sort a defined dataset
'   It optionally take two arguments of a worksheet and a sort range.
'   If no arguments are passed in, the activesheet is used and teh dataset on the activesheet used as the range
'   Currently, the code will perform a sort of the data in ascending order.
'   A possible modification is to add a third optional argument for sort order to make it more flexible (i.e. Ascending or Descending)


    Dim lastCol As Long
    Dim lastRow As Long


    'Test to see if an optional worksheet reference was passed in
    If ws Is Nothing Then
    
        Set ws = ThisWorkbook.ActiveSheet
        
    End If
    
    
    'Test to see if an optional range reference was passed in
    If rngSort Is Nothing Then
    
        'if no range was passed in then find the number of columns and rows _
         in the dataset and create a range using the last row / last column cell references
         
        lastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
        lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        Set rngSort = ws.Range(Cells(1, 1), Cells(lastRow, lastCol))
        
    End If
    

    'apply the sort to the defined range
    rngSort.Sort "<ticker>", xlAscending, "<date>", , xlAscending, , , xlYes

End Sub



Sub AnalyzeChange(YrlyChange As Range)

'   Created By: Karl Ramsay
'   Created:    Feb. 8th, 2020
'   Modified:   Feb. 10th, 2020
'   Modified:
'   Description: The purpose of this code is to determine if the Percentage Change meets the criteria for shading
'   the Percentage Change for each ticker based on the range the value falls in (i.e. red for negative, and green for positive




    If YrlyChange.Value > 0 Then
    
        YrlyChange.Interior.ColorIndex = 4
        
        
    ElseIf YrlyChange.Value < 0 Then
    
        YrlyChange.Interior.ColorIndex = 3
    
    Else
        
    End If
    

End Sub


Sub FindFirstLastDayPricing(sht As Variant)
    
'   Created By: Karl Ramsay
'   Created:    Feb. 8th, 2020
'   Modified:   Feb. 10th, 2020
'   Modified:
'   Description: The purpose of this code is to create a temp sheet that will be used to track first Open and last Open
'   for each year and Ticker. Once we find the data for each ticker we calculate the delta and posst results to our temp sheet
'   Next we copy the results to the corresponding sheet and clean up  9 delete the temp sheet
'   Currently the code is designed to delete the temp sheet after each sheet is processed. A possible optimization will included using the same
'   temp sheet to process all files in the dataset or even doing all processing on the individual sheets.

  
Dim wks As Excel.Worksheet
Dim lastR As Double
Dim tickerName As String
Dim totalStockVolume As Double

Application.DisplayAlerts = False


Set wks = ThisWorkbook.Worksheets.Add

'create a reconciliation sheet to check data
wks.Name = "Reconciliation"
wks.Range("A1").Value = "Ticker"
wks.Range("B1").Value = "First Day Open"
wks.Range("C1").Value = "Last Day Close"
wks.Range("D1").Value = "Yearly Change"
wks.Range("E1").Value = "Percentage Change"
wks.Range("F1").Value = "Total Stock Volume"

'set row width
wks.Range("A:F").ColumnWidth = 16

'set the value for the Last row
lastR = sht.Range("A" & Rows.Count).End(xlUp).Row

tickerName = sht.Cells(2, 1).Value
totalStockVolume = 0
roffset = 2
wks.Cells(roffset, 1).Value = sht.Cells(2, 1).Value
wks.Cells(roffset, 2).Value = sht.Cells(2, 3).Value
    
For x = 2 To lastR
   
    'if the ticker symbol changes it means we have to go back 1 row to get the Last Day Close Price
    If sht.Cells(x, 1).Value <> tickerName Then
        'set the Last Day Close Value for ticker
        wks.Cells(roffset, 3).Value = sht.Cells(x - 1, 6).Value
        wks.Cells(roffset, 6).Value = totalStockVolume
        totalStockVolume = 0
        
        If wks.Cells(roffset, 2).Value = 0 Then
        
            'handle divide by zero
            wks.Cells(roffset, 5).Value = 0
        
        Else
        
            wks.Cells(roffset, 4).Value = (wks.Cells(roffset, 3).Value - wks.Cells(roffset, 2).Value)
            AnalyzeChange wks.Cells(roffset, 4)
            wks.Cells(roffset, 5).Value = (wks.Cells(roffset, 3).Value - wks.Cells(roffset, 2).Value) / wks.Cells(roffset, 2).Value
            wks.Cells(roffset, 5).NumberFormat = "0.00%"
                        
            
        End If
    
        
        'now set ticker to the new symbol
        tickerName = sht.Cells(x, 1).Value
        
        'increment the row offset in the reconciliation sheet
        roffset = roffset + 1
        
        wks.Cells(roffset, 1).Value = tickerName
        wks.Cells(roffset, 2).Value = sht.Cells(x, 3).Value
        
        
    Else
    
        'keep adding the Total Stock Volume
        totalStockVolume = totalStockVolume + sht.Cells(x, 7).Value
            
    End If
    

    

Next x


'now copy the results to the appropriate sheet
wks.Range("A:A").Copy sht.Range("I:I")
wks.Range("D:D").Copy sht.Range("J:J")
wks.Range("E:E").Copy sht.Range("K:K")
wks.Range("F:F").Copy sht.Range("L:L")

wks.Delete

Application.DisplayAlerts = True

End Sub
