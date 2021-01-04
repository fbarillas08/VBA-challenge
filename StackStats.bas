Attribute VB_Name = "StackStats"
Sub StockStats()

' VBA Challenge Script Code
' This is the code used to process the multi-year performance of over 3000 stocks
' over three consecutive years (2014,15,16).

' Values captured for analysis include Beginning-of-Year (BOY) and End-of-Year (EOY) Price Points
' Calculated values include (EOY-BOY) values, Annual Price Change, and AGR%

' This code includes the basic and the advanced (bonus) challenge.

' This script includes the data manipulation code and the formatting of the tables code for easy
' reading and executive level appeal.


'Defining Variables and Aggregated Quantities

    Dim LastRow As Long
    Dim tickersymbol As String
    Dim tickervolumetotal As Double
    Dim i As Long
    Dim j As Long
    Dim r As Long

    Dim ws As Worksheet
    Dim SummaryTableRow As Integer
    
    Dim growthrange As Range
    Dim volumerange As Range
    
    Dim maxgrowth As Double
    Dim mingrowth As Double
    Dim maxvolumne As Long
        
'-----------------------------
     
For Each ws In Worksheets

' ----------------------------
   
' Formatting Summary Table and Data Table

    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("j1").Value = "BOY Opening"
    ws.Range("k1").Value = "EOY Closing"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "% Change"
    ws.Range("N1").Value = "Total Volume"
    
    ws.Range("i1:n1, A1:G1").Interior.ColorIndex = 56
    ws.Range("i1:n1, A1:G1").Font.ColorIndex = 2
    ws.Range("i1:n1, A1:G1").Font.FontStyle = "Bold"
    ws.Range("i1:n1, A1:G1").HorizontalAlignment = xlCenter
    ws.Range("i1:n1, A1:G1").ColumnWidth = 15

'Formatting the Headers for the Outlier Table
    
    ws.Range("P1").Value = "Category"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase:"
    ws.Range("P3").Value = "Greatest % Decrease:"
    ws.Range("P4").Value = "Greatest Volume:"
    ws.Range("P2:P4").Font.FontStyle = "Bold"
    ws.Range("P2:P4").HorizontalAlignment = xlRight
    ws.Columns("P").AutoFit
    
    ws.Range("P1:R1").Interior.ColorIndex = 56
    ws.Range("P1:R1").Font.ColorIndex = 2
    ws.Range("P1:R1").Font.FontStyle = "Bold"
    ws.Range("P1:R1").HorizontalAlignment = xlCenter


' Finding the Last Row of of daily ticker table

    LastRow = ws.Cells(Rows.count, "a").End(xlUp).Row
     
'Populating the First BOY Price of First Ticker of the Summary Table

    ws.Range("j2").Value = Cells(2, 3).Value
          
' Rolling Over the Entire Ticket Column

    SummaryTableRow = 2

    For i = 2 To LastRow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
  
        tickersymbol = ws.Cells(i, 1).Value
    
        tickervolumetotal = tickervolumetotal + ws.Cells(i, 7).Value
    
        ws.Range("I" & SummaryTableRow).Value = tickersymbol
    
        ws.Range("n" & SummaryTableRow).Value = tickervolumetotal
    
        SummaryTableRow = SummaryTableRow + 1
    
        tickervolumetotal = 0
      
      Else
         
        tickervolumetotal = tickervolumetotal + ws.Cells(i, 7).Value
    
        ws.Cells(SummaryTableRow, 11).Value = ws.Cells(i + 1, 6).Value 'Records the EOY Price
        
        ws.Cells(SummaryTableRow + 1, 10).Value = ws.Cells(i + 2, 3).Value
    
            
    End If

  Next i

' Rolling over the Summary Table Column for BOY and EOY Prices

For j = 2 To SummaryTableRow - 1
    
    ws.Cells(j, 12).Value = ws.Cells(j, 11).Value - ws.Cells(j, 10).Value
    
   
   If ws.Cells(j, 10).Value = 0 Then
                ws.Cells(j, 13).Value = 0
            Else
                ws.Cells(j, 13).Value = ws.Cells(j, 12).Value / ws.Cells(j, 10).Value
            End If
   
        
    Next j
    
' Formating Cells in the Summary Table

    ws.Range("J2:J" & SummaryTableRow).NumberFormat = "###0.00"
    ws.Range("K2:K" & SummaryTableRow).NumberFormat = "###0.00"
    ws.Range("L2:L" & SummaryTableRow).NumberFormat = "###0.00;[Red] -###0.00"
    ws.Range("m2:m" & SummaryTableRow).NumberFormat = "0.00%"
    ws.Range("N2:N" & SummaryTableRow).NumberFormat = "#,###"
    
'Conditional Coloring Cells of  % Change

For i = 2 To SummaryTableRow

If ws.Cells(i, 13).Value >= 0.15 Then

      ws.Cells(i, 13).Interior.ColorIndex = 35

      
  ElseIf ws.Cells(i, 13).Value >= 0 Then

      ws.Cells(i, 13).Interior.ColorIndex = 36

      
  Else

      ws.Cells(i, 13).Interior.ColorIndex = 38

      
  End If
    
Next i

' Finding the Outliers and Populating Summary Table
      
    Set growthrange = ws.Range("M2:M" & SummaryTableRow)
    Set volumerange = ws.Range("N2:N" & SummaryTableRow)
    
    maxgrowth = Application.WorksheetFunction.Max(growthrange)
    mingrowth = Application.WorksheetFunction.Min(growthrange)
    maxvolume = Application.WorksheetFunction.Max(volumerange)
    
    ws.Range("R2").Value = maxgrowth
    ws.Range("R3").Value = mingrowth
    ws.Range("R4").Value = maxvolume
    
    ws.Range("r2:r3").NumberFormat = "0.00%"
    ws.Range("r4").NumberFormat = "#,###"
    ws.Columns("R").AutoFit
    
    For r = 2 To SummaryTableRow
    If ws.Range("m" & r).Value = maxgrowth Then ws.Range("Q2").Value = ws.Range("i" & r).Value
    If ws.Range("M" & r).Value = mingrowth Then ws.Range("Q3").Value = ws.Range("i" & r).Value
    If ws.Range("N" & r).Value = maxvolume Then ws.Range("Q4").Value = ws.Range("i" & r).Value
    Next r
    
    'Formatting the Bonus Table
    
    ws.Range("Q2:Q4").HorizontalAlignment = xlCenter
    ws.Range("P2:R4").Borders.LineStyle = xlContinuous
    ws.Range("P2:R4").Borders.Weight = 2
    
Next ws
   
    MsgBox ("Process Finished")

    
End Sub




