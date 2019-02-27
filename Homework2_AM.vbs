'**************************************************************************
'UC Berkeley Extension Data Analytics Program
'Homework 2 - VBA
'Submitted by: Alejandro Montesinos
'Febraury 26, 2019
'**************************************************************************
Sub HomeworkTwo():
   For Each ws In Worksheets
      '----------------------------------------------------------
      'Define variables
      Dim LastRow             As Long
      Dim Total_Stock_Vol     As Double
      Dim Open_Price          As Double
      Dim Diff_Ticker_Counter As Integer
      '----------------------------------------------------------
     
      '----------------------------------------------------------
      'Set parameter for easyreference
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      '----------------------------------------------------------
      
      '----------------------------------------------------------
      'Initialize Summary Table
      ws.Range("I1") = "Solutions by Alejandro Montesinos"
      ws.Range("I1").Font.Size = 14
   
      ws.Range("I2") = "Ticker"
      ws.Range("J2") = "Yearly Change"
      ws.Range("K2") = "Percent Change"
      ws.Range("L2") = "Total Stock Volume"
      
      ws.Range("I2:L2").Interior.ColorIndex = 37
      ws.Range("I1:L2").Font.Bold = True
      '----------------------------------------------------------
      
      '----------------------------------------------------------
      'Fill In Summary Table
      ws.Range("A2:A" & LastRow).Sort Key1:=ws.Range("A1"), Order1:=xlAscending              'Sort by <ticker>. We want to make sure <ticker> is properly sorted or we may get a wrong result when looping.
   
      Diff_Ticker_Counter = 2                                                                'Initialize the different <ticker> counter
      Total_Stock_Vol = 0                                                                    'Initalize the <Total Stock Volume> cummulator
      Open_Price = 0                                                                         'Initalize <Open price>
   
      For i = 2 To LastRow
         If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Open_Price = ws.Cells(i, 3).Value                                                'Store initial value for <Open price>
         ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Diff_Ticker_Counter = Diff_Ticker_Counter + 1                                    'Add 1 to the different <ticker> counter every time we find a new ticker code
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value                         ' Dont forget to update the <Total Stock Volume> cummulator
            
            ws.Range("I" & Diff_Ticker_Counter).Value = ws.Cells(i, 1).Value
            ws.Range("J" & Diff_Ticker_Counter).Value = ws.Cells(i, 6).Value - Open_Price
            If Open_Price <> 0 Then
               ws.Range("K" & Diff_Ticker_Counter).Value = ws.Range("J" & Diff_Ticker_Counter).Value / Open_Price
            End If
            ws.Range("L" & Diff_Ticker_Counter).Value = Total_Stock_Vol
            
            Total_Stock_Vol = 0                                                  'Reset the cummulator to zero
            Open_Price = 0                                                       'Reset the open price to zero
         Else
            Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
         End If
      Next i
      '----------------------------------------------------------
   
   
      '----------------------------------------------------------
      'Apply conditional and regular formatting to Summary Table
      LastTabR = ws.Cells(Rows.Count, 9).End(xlUp).Row                          'Define Last row of Summary Table
      
      For j = 3 To LastTabR
         If ws.Cells(j, 10).Value = Empty Then
            ws.Cells(j, 10).Interior.ColorIndex = 2                             'If the cell is empty then no color should be applied
         ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3                             'If change is negative the cell is red
         Else
            ws.Cells(j, 10).Interior.ColorIndex = 4                             'If the change is not negative (zero allowed) then the cell is green
         End If
      Next j
      '----------------------------------------------------------
      
   
      '----------------------------------------------------------
      'Apply format to Summary Table
      ws.Range("J3:J" & LastTabR).NumberFormat = "#.0000"
      ws.Range("K3:K" & LastTabR).NumberFormat = "0.00%"
      ws.Range("L3:L" & LastTabR).NumberFormat = "#,##0"
      ws.Range("I2:L" & LastTabR).Borders.LineStyle = xlContinuous
      '----------------------------------------------------------
      
      
      '----------------------------------------------------------
      'Initialize extra points Table
      ws.Range("N1") = "Extra Table"
      ws.Range("N1").Font.Size = 14
   

      ws.Range("N2") = "Description"
      ws.Range("O2") = "Ticker"
      ws.Range("P2") = "Value"
      
      ws.Range("N2:P2").Interior.ColorIndex = 37
      ws.Range("N1:P2").Font.Bold = True
      ws.Range("N2:P5").Borders.LineStyle = xlContinuous
      
      ws.Range("N3") = "Greatest % Increase"
      ws.Range("N4") = "Greatest % Decrease"
      ws.Range("N5") = "Greatest Total Volume"
      '----------------------------------------------------------
      
      '----------------------------------------------------------
      'Initialize extra points Table
      Dim IncVal As Double
      Dim IncTrack As String
      
      Dim DecVal As Double
      Dim DecTrack As String
      
      IncVal = ws.Cells(3, 11).Value
      IncTrack = ws.Cells(3, 9).Value
      
      DecVal = ws.Cells(3, 11).Value
      DecTrack = ws.Cells(3, 9).Value

      VolVal = ws.Cells(3, 12).Value
      VolTrack = ws.Cells(3, 9).Value
    
      For k = 4 To LastTabR
         If ws.Cells(k, 11).Value > IncVal Then
            IncVal = ws.Cells(k, 11).Value                      'Update the value if it is the largest so far
            IncTrack = ws.Cells(k, 9).Value                     'Keep track of the <Ticker> with the largest value
         End If
         
         If ws.Cells(k, 11).Value < DecVal Then
            DecVal = ws.Cells(k, 11).Value                     'Update the value if it is the lowest so far
            DecTrack = ws.Cells(k, 9).Value                    'Keep track of the <Ticker> with the lowest value
         End If

         If ws.Cells(k, 12).Value > VolVal Then
            VolVal = ws.Cells(k, 12).Value                     'Update the value if it is the largest volume so far
            VolTrack = ws.Cells(k, 9).Value                    'Keep track of the <Ticker> with the largest volume
         End If
      Next k
      
      'Apply some formatting
      ws.Range("O3") = IncTrack
      ws.Range("P3") = IncVal
      ws.Range("O4") = DecTrack
      ws.Range("P4") = DecVal
      ws.Range("O5") = VolTrack
      ws.Range("P5") = VolVal

      ws.Range("P3:P4").NumberFormat = "0.00%"
      ws.Range("P5").NumberFormat = "#,##0"
      ws.Columns("J:P").AutoFit
      '----------------------------------------------------------
   Next ws
End Sub
'**************************************************************************


