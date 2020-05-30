Sub AllStockSummary()

Dim Current As Worksheet
For Each Current In Worksheets ' CHALLENGE 2: This loop cycles through and creates the results charts for each sheet

' activate current sheet
Current.Activate

' Create Stocks Summary
Dim i As Long
Dim maxrow As Long
Dim ResultIndex As Integer
Dim tCounter As Long
Dim openPrice As Double
Dim closePrice As Double
Dim percentArray() As Variant


Dim maxDate As Long
Dim minDate As Long

maxrow = Cells(Rows.Count, 1).End(xlUp).Row()

' titles results fields
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

Range("P2").Style = "Percent"
Range("P3").Style = "Percent"


ResultIndex = 2
tCounter = 0

For i = 2 To maxrow ' This loop cycles through and adds a row to the results for each ticker symbole

    If Range("A" & (i)) <> Range("A" & (i + 1)) Then
        'Ticker
        Range("I" & ResultIndex).Value = Range("A" & (i))
        
        ' Yearly Change
        closePrice = Range("F" & i)
        openPrice = Range("C" & (i - tCounter))
        Range("J" & ResultIndex).Value = closePrice - openPrice
      
        
        ' Percent Change - accounts for possibility of Zeros in denominator of calculation or 0% change
        If openPrice = 0 Or (closePrice - openPrice) = 0 Then
            Range("K" & ResultIndex).Value = 0
        Else
            Range("K" & ResultIndex).Value = (closePrice - openPrice) / openPrice
        End If
        
        Range("K" & ResultIndex).Style = "Percent" ' sets style of %change column to Percent
        
        ' Total Stock Volume
        Range("L" & ResultIndex).Value = WorksheetFunction.Sum(Range("G" & i, "G" & (i - tCounter)))
        
        ' Conditional formatting of Yearly Change Column
        If closePrice - openPrice >= 0 Then
         Range("J" & ResultIndex).Interior.ColorIndex = 4
         Else
            Range("J" & ResultIndex).Interior.ColorIndex = 3
        End If
        
        ' Increments Results table
        ResultIndex = ResultIndex + 1
        
        'Resets segment counter
        tCounter = 0
    Else
      tCounter = tCounter + 1
      
    End If
Next i

' CHALLENGE 2: The Code below creates and displays max%Change, min%Change, and maxVolume for each sheet.
Dim minP As Double
Dim maxP As Double
Dim maxV As Double
Dim j As Long
Dim maxPInd As Boolean
Dim minPInd As Boolean
Dim maxVind As Boolean

' These variables indicate if the desired values have been found in the results.
maxPInd = False
minPInd = False
maxVind = False

' Calculates min and max percent change, and max volume based on results produced.
maxP = WorksheetFunction.Max(Range("K2:K" & ResultIndex))
minP = WorksheetFunction.Min(Range("K2:K" & ResultIndex))
maxV = WorksheetFunction.Max(Range("L2:L" & ResultIndex))

' This loop iterates through results 
For j = 2 To ResultIndex
' Searches for Maximum Percent Change
    If maxP = Range("K" & j) And maxPInd = False Then
        maxPInd = True
        Range("O2").Value = Range("I" & j)
        Range("P2").Value = maxP
    Else
        maxPInd = False
    End If
' Searches for Minimum Percent Change
    If minP = Range("K" & j) And minPInd = False Then
        minPInd = True
        Range("O3").Value = Range("I" & j)
        Range("P3").Value = minP
    Else
        minPInd = False
    End If
' Searches for Maximum Volume
    If maxV = Range("L" & j) And maxVind = False Then
        maxPInd = True
        Range("O4").Value = Range("I" & j)
        Range("P4").Value = maxV
    Else
        maxPInd = False
    End If
Next j


Next ' finalizes the master loop that cycles through all sheets

End Sub