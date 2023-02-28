Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data():
    
For Each ws In Worksheets

'Declares the following variables and assigns a data type to each.
Dim WorksheetName As String
Dim TickerCount, LastRowColumnA, LastRowColumnI, i, j As Long
Dim PercentChange As Double


WorksheetName = ws.Name
              
'Set the start row for the ticker counter to row 2.
TickerCount = 2
        
'Sets the start row to row 2.
j = 2

'Calculates the last row in column A / needed to tell the loop where to stop and loop back again.
LastRowColumnA = ws.Range("A" & Rows.Count).End(xlUp).Row
                 
        
        
'Sets the new column headers for first 4 assignment measures.
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

        
                
'Sets range for the calculated columns loop, and begins the loop.
For i = 2 To LastRowColumnA
            
'Fill in ticker name in column I / runs calculation for yearly change and fills in the values in column J.
If ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value Then
ws.Range("I" & TickerCount).Value = ws.Range("A" & i).Value
ws.Range("J" & TickerCount).Value = ws.Range("F" & i).Value - ws.Range("C" & j).Value
                
'Formats the cells in yearly change column to red or green fill based on value above or below 0. Above 0 is green, and below 0 is red.
If ws.Range("J" & TickerCount).Value < 0 Then
ws.Range("J" & TickerCount).Interior.ColorIndex = 3 'red
Else
ws.Range("J" & TickerCount).Interior.ColorIndex = 4 'green
End If
                    
'Runs calculation for percent change and fills in the values in column K / formats the cells to the percentage format.
If ws.Range("C" & j).Value <> 0 Then
PercentChange = ((ws.Range("F" & i).Value - ws.Range("C" & j).Value) / ws.Range("C" & j).Value)
ws.Range("K" & TickerCount).Value = Format(PercentChange, "Percent")
Else
ws.Range("K" & TickCount).Value = Format(0, "Percent")
End If
                    
'Runs calculation for stock volume and fills in the values in column L / tells code to loop back to the beginning and run again for next row.
ws.Range("L" & TickerCount).Value = WorksheetFunction.Sum(Range(ws.Range("G" & j), ws.Range("G" & i)))
TickerCount = TickerCount + 1
j = i + 1
End If

Next i





'Declares the following variables and assigns a data type to each (double for all).
Dim GreatestIncrease, GreatestDecrease, GreatestVolume As Double



'Sets new headers for the following summary measures.
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


        
'Tells where each of the following measures should begin from.
GreatestVolume = ws.Range("L2").Value
GreatestIncrease = ws.Range("K2").Value
GreatestDecrease = ws.Range("K2").Value
        
        
        
'Calculates the last row in column I / needed to tell the loop where to stop and loop back again.
LastRowColumnI = ws.Range("I" & Rows.Count).End(xlUp).Row
        
        
        
'Sets range for the summary loop, and begins the loop.
For i = 2 To LastRowColumnI
            
'Finds the greatest value in the total stock volume column and desgniates that as greatest volume value.
If ws.Range("L" & i).Value > GreatestVolume Then
GreatestVolume = ws.Range("L" & i).Value
ws.Range("P4").Value = ws.Range("I" & i).Value
Else
GreatestVolume = GreatestVolume
End If
                
'Finds the greatest value in the percent change column and desgniates that as greatest increase value.
If ws.Range("K" & i).Value > GreatestIncrease Then
GreatestIncrease = ws.Range("K" & i).Value
ws.Range("P2").Value = ws.Range("I" & i).Value
Else
GreatestIncrease = GreatestIncrease
End If
                
'Finds the lowest value in the percent change column and desgniates that as greatest decrease value.
If ws.Range("K" & i).Value < GreatestDecrease Then
GreatestDecrease = ws.Range("K" & i).Value
ws.Range("P3").Value = ws.Range("I" & i).Value
Else
GreatestDecrease = GreatestDecrease
End If
                
                
                
'Places the following measures into specific cells and formats them accordingly.
ws.Range("Q2").Value = Format(GreatestIncrease, "Percent")
ws.Range("Q3").Value = Format(GreatestDecrease, "Percent")
ws.Range("Q4").Value = Format(GreatestVolume, "Scientific")
            
Next i
            
            
            
'Adjusts all column widths across all worksheets.
Worksheets(WorksheetName).Columns("A:Z").AutoFit
            
Next ws



End Sub
