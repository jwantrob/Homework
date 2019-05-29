Attribute VB_Name = "Module1"
Sub Sum():
'Loop for all sheets in file found in https://support.microsoft.com/en-us/help/142126/macro-to-loop-through-all-worksheets-in-a-workbook
Dim Sheet As Integer
Dim WS_Count As Integer
WS_Count = ActiveWorkbook.Worksheets.Count
For Sheet = 1 To WS_Count

    'Find number of rows found in https://excel.officetuts.net/en/vba/count-rows-in-excel-vba
    Dim Last_Row As Long
    Last_Row = ActiveWorkbook.Worksheets(Sheet).Cells(Rows.Count, 1).End(xlUp).Row
   
    
    'Make sure the in alphabetical order found https://trumpexcel.com/sort-data-vba/
    
    'Range("A2:A70926").Sort Key1:=Range("A1"), Order1:=xlAscending
    
    
    'Set variables
    
    Dim TickerName As String
    Dim Sum As Double
    Dim Row As Integer
    Dim OpenVal As Double
    Dim CloseVal As Double
    Dim YearChangeVal As Double
    Dim PChangeVal As Double
    Dim PchangePer As String
    Row = 2
    'Set to first value for each variable in the sheet
    TickerName = ActiveWorkbook.Worksheets(Sheet).Range("A2").Value
    Sum = ActiveWorkbook.Worksheets(Sheet).Range("G2").Value
    OpenVal = ActiveWorkbook.Worksheets(Sheet).Range("C2").Value
    'Set Header Names
    ActiveWorkbook.Worksheets(Sheet).Range("J1").Value = "Ticker"
    ActiveWorkbook.Worksheets(Sheet).Range("K1").Value = "Total Stock Volume"
    ActiveWorkbook.Worksheets(Sheet).Range("L1").Value = "Yearly Change"
    ActiveWorkbook.Worksheets(Sheet).Range("M1").Value = "Yearly Percent Change"
    ActiveWorkbook.Worksheets(Sheet).Range("J2").Value = TickerName
    
    'start a for loop to loop through every ticker in column A
    Dim i As Long
    
    For i = 2 To Last_Row
    'Start an if statement to figure out which volumes to sum, percent change, yearly change
    'If the next cell is equal to the current cell then sum add the next cell to the running sum
        If ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 1).Value = ActiveWorkbook.Worksheets(Sheet).Cells(i, 1).Value Then
        Sum = Sum + ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 7).Value
        Else
    'If not then this is the end of the ticker and calculate the close value, year change, percent change and final
        CloseVal = ActiveWorkbook.Worksheets(Sheet).Cells(i, 6).Value
        YearChangeVal = CloseVal - OpenVal
            If OpenVal = 0 Then
            PChangeVal = 0
            Else
            PChangeVal = (CloseVal - OpenVal) / OpenVal
            End If
    'Found format function on https://www.techonthenet.com/excel/formulas/format_string.php
        PchangePer = Format(PChangeVal, "Percent")
    'New If statement to highlight positive changes in green (4) and negative changes in red (3) (both change and pchange)
        If YearChangeVal < 0 Then
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 12).Interior.ColorIndex = 3
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 13).Interior.ColorIndex = 3
        Else
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 12).Interior.ColorIndex = 4
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 13).Interior.ColorIndex = 4
        End If
    'Output all of these into the correct row in the summary table
    
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 11).Value = Sum
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 12).Value = YearChangeVal
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 13).Value = PchangePer
    'Move down one row for the next output
        Row = Row + 1
    'Change the tickename to the next ticker
        TickerName = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 1).Value
    'Output the ticker into the next cell to start the process again
        ActiveWorkbook.Worksheets(Sheet).Cells(Row, 10).Value = TickerName
    'Start the sum over again
        Sum = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 7).Value
    'Start the open value over again
        OpenVal = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 3).Value
        End If
    Next i
    
    'Summarize the Summary
    'Set the variables
    Dim PIncrease As Double
    Dim PDecrease As Double
    Dim PIncreasePer As String
    Dim PDecreasePer As String
    Dim GVolumeVal As Double
    Dim TickerNameIn As String
    Dim TickerNameDe As String
    Dim TickerNameV As String
    
    
    ActiveWorkbook.Worksheets(Sheet).Range("P2").Value = "Greatest % Increase"
    ActiveWorkbook.Worksheets(Sheet).Range("P3").Value = "Greatest % Decrease"
    ActiveWorkbook.Worksheets(Sheet).Range("P4").Value = "Greatest Total Volume"
    ActiveWorkbook.Worksheets(Sheet).Range("Q1").Value = "Ticker"
    ActiveWorkbook.Worksheets(Sheet).Range("R1").Value = "Value"
    'Set Row back to the final Row of the summary table
    Row = Row - 1
    
    
    'If percent de/increase/volume is greater in the third row than second row then that is the largest de/increase/volume otherwise the second row is the greatest
    'Increase
    
    If ActiveWorkbook.Worksheets(Sheet).Cells(3, 13).Value > ActiveWorkbook.Worksheets(Sheet).Cells(2, 13).Value Then
        PIncrease = ActiveWorkbook.Worksheets(Sheet).Cells(3, 13).Value
        TickerNameIn = ActiveWorkbook.Worksheets(Sheet).Cells(3, 10).Value
        Else
        PIncrease = ActiveWorkbook.Worksheets(Sheet).Cells(2, 13).Value
        TickerNameIn = ActiveWorkbook.Worksheets(Sheet).Cells(2, 10).Value
        End If
    'Decrease
    If ActiveWorkbook.Worksheets(Sheet).Cells(3, 13).Value < ActiveWorkbook.Worksheets(Sheet).Cells(2, 13).Value Then
        PDecrease = ActiveWorkbook.Worksheets(Sheet).Cells(3, 13).Value
        TickerNameDe = ActiveWorkbook.Worksheets(Sheet).Cells(3, 10).Value
        Else
        PDecrease = ActiveWorkbook.Worksheets(Sheet).Cells(2, 13).Value
        TickerNameDe = ActiveWorkbook.Worksheets(Sheet).Cells(2, 10).Value
        End If
    'Volume
    If ActiveWorkbook.Worksheets(Sheet).Cells(3, 11).Value > ActiveWorkbook.Worksheets(Sheet).Cells(2, 11).Value Then
        GVolumeVal = ActiveWorkbook.Worksheets(Sheet).Cells(3, 11).Value
        TickerNameV = ActiveWorkbook.Worksheets(Sheet).Cells(3, 10).Value
        Else
        GVolumeVal = ActiveWorkbook.Worksheets(Sheet).Cells(2, 11).Value
        TickerNameV = ActiveWorkbook.Worksheets(Sheet).Cells(2, 10).Value
        End If
    'Create a For loop to go through the summary table to test the rest of the summary table to find greatest percent de/increase and volume
    For i = 3 To Row
    'Increase
        If PIncrease > ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 13) Then
        PIncrease = PIncrease
        TickerNameIn = TickerNameIn
        Else
        PIncrease = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 13)
        TickerNameIn = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 10)
        End If
    'Decrease
        If PDecrease < ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 13) Then
        PDecrease = PDecrease
        TickerNameDe = TickerNameDe
        Else
        PDecrease = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 13)
        TickerNameDe = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 10)
        End If
    'Volume
        If GVolumeVal > ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 11) Then
        GVolumeVal = GVolumeVal
        TickerNameV = TickerNameV
        Else
        GVolumeVal = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 11)
        TickerNameV = ActiveWorkbook.Worksheets(Sheet).Cells(i + 1, 10)
        End If
    Next i
    'Output final answer
    'As Percent with Ticker Name
    PIncreasePer = Format(PIncrease, "Percent")
    PDecreasePer = Format(PDecrease, "Percent")
    ActiveWorkbook.Worksheets(Sheet).Range("R2").Value = PIncreasePer
    ActiveWorkbook.Worksheets(Sheet).Range("Q2").Value = TickerNameIn
    ActiveWorkbook.Worksheets(Sheet).Range("R3").Value = PDecreasePer
    ActiveWorkbook.Worksheets(Sheet).Range("Q3").Value = TickerNameDe
    'As Total Volume with Ticker Name
    ActiveWorkbook.Worksheets(Sheet).Range("R4").Value = GVolumeVal
    ActiveWorkbook.Worksheets(Sheet).Range("Q4").Value = TickerNameV
Next Sheet


End Sub



