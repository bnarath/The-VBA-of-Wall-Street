Attribute VB_Name = "Module3"
Sub Stock_Analysis_Method3()

'Declarations

Dim ws As Worksheet, Row_Index As Long, First_Instance As Long, Count As Long, Error_Count As Long, Last_Row As Long, Yearly_Change As Double, Percent_Change As Double, Total_Stock_Volume As Double, Topper_Flopper_Names(), Topper_Flopper_Values()

'To keep track of time
Dim StartTime As Double, MinutesElapsed As Double
'Log the start time
StartTime = Timer

'Iterate through worksheets and activate each worksheet one after the other.

    For Each ws In ThisWorkbook.Worksheets
        'For Debugging
        'MsgBox (ws.name)
        
        'Activate the current sheet
        Worksheets(ws.Name).Activate
    
        'Sort Each sheets based on column A in ascending order
        Range("A2").End(xlDown).End(xlToRight).Sort [A2], xlAscending, Header:=xlYes
        
        'Last Row
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Adding the label for summary
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        'Initialise the Count as 0 : No Ticker is found towards summary yet ! No Div By Zero yet!: Initialize the Total Stock Volume of the current Ticker as 0
        Count = 0
        Error_Count = 0 'To keep track of Div By Zero
        Total_Stock_Volume = 0
        
        'Start with Row2; Iterate through all the rows; First_Instance corresponds to the first instance of the current Ticker
        First_Instance = 2
        For Row_Index = 2 To Last_Row
            
            If Cells(Row_Index, 1).Value <> Cells(Row_Index + 1, 1).Value Then 'Check if there is a change in value of Ticker'
            
                'Yay, we found a new Ticker
                Count = Count + 1
        
                'Calculate Yearly_Change
                Yearly_Change = Cells(Row_Index, 6).Value - Cells(First_Instance, 3).Value
        
                'Calculate Percent Change - Need to take care of Div by Zero
                If Cells(First_Instance, 3).Value = 0 Then
                    Percent_Change = 0
                    If Error_Count = 0 Then
                        'Labels are added on first Error
                        Cells(6, 16).Value = "Error Tickers(OV is 0)"
                        Cells(6, 16).Interior.ColorIndex = 6
                    End If
                    Cells(7 + Error_Count, 16).Value = Cells(Row_Index, 1).Value
                    Cells(7 + Error_Count, 16).Interior.ColorIndex = 6
                    Error_Count = Error_Count + 1
                Else
                    Percent_Change = Yearly_Change / Cells(First_Instance, 3).Value
                End If
        
                'Calculate Total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(Row_Index, 7).Value
        
                'What are you waiting for? Write down in the summary with color conditioning
                Cells(Count + 1, 9).Value = Cells(Row_Index, 1).Value
                Cells(Count + 1, 10).Value = Yearly_Change
                Cells(Count + 1, 11).Value = Percent_Change
                Cells(Count + 1, 12).Value = Total_Stock_Volume
        
                'Color Conditional Formatting for "Yearly Change"
                If Yearly_Change > 0 Then
                    Cells(Count + 1, 10).Interior.ColorIndex = 4 'Green
                Else  'I cousider 'no change' in stock value is also bad, hence red !
                    Cells(Count + 1, 10).Interior.ColorIndex = 3 'Red
                End If
        
                'Time for Topper-Flopper update
                If Count = 1 Then
                'Initialize the Topper and Flopper as the first Ticker
                    Topper_Flopper_Names = Array(Cells(Row_Index, 1), Cells(Row_Index, 1), Cells(Row_Index, 1))
                    Topper_Flopper_Values = Array(Percent_Change, Percent_Change, Total_Stock_Volume)
                Else
        
                        If Percent_Change > Topper_Flopper_Values(0) Then 'Greatest % increase tracking
                            Topper_Flopper_Values(0) = Percent_Change
                            Topper_Flopper_Names(0) = Cells(Row_Index, 1)
                        End If
                
                        If Percent_Change < Topper_Flopper_Values(1) Then 'Greatest % decrease tracking
                            Topper_Flopper_Values(1) = Percent_Change
                            Topper_Flopper_Names(1) = Cells(Row_Index, 1)
                        End If
                
                        If Total_Stock_Volume > Topper_Flopper_Values(2) Then 'Greatest stock volume tracking
                            Topper_Flopper_Values(2) = Total_Stock_Volume
                            Topper_Flopper_Names(2) = Cells(Row_Index, 1)
                        End If
            
                End If
                
                'Update the First_Occurance of Next Ticker and reset the Total_Stock_Volume
                First_Instance = Row_Index + 1
                Total_Stock_Volume = 0
                
            Else
                'Just update the total Stock Volume
                Total_Stock_Volume = Total_Stock_Volume + Cells(Row_Index, 7).Value
            End If
            
        Next Row_Index
        
        'Fill the results of toppers and floppers
        
        'Label Filling
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        'Results
        Cells(2, 16).Value = Topper_Flopper_Names(0)
        Cells(3, 16).Value = Topper_Flopper_Names(1)
        Cells(4, 16).Value = Topper_Flopper_Names(2)
        
        Cells(2, 17).Value = Topper_Flopper_Values(0)
        Cells(3, 17).Value = Topper_Flopper_Values(1)
        Cells(4, 17).Value = Topper_Flopper_Values(2)
        
        'Number and  Column width formatting
        Call Formatting
        
    Next

    'Log time in seconds the code took to run
    MinutesElapsed = Round((Timer - StartTime) / 60, 2)

    'Notify user the time taken in minutes
    MsgBox "This code ran successfully in " & MinutesElapsed & " minutes", vbInformation

End Sub



Sub Formatting()
'
' Number and  Column width formatting Macro
'

'
    Columns(10).NumberFormat = "0.00"
    Columns(11).NumberFormat = "0.00%"
    Range(Cells(2, 17), Cells(3, 17)).NumberFormat = "0.00%"
    Cells(4, 17).NumberFormat = "0.0000E+00"
        
    Columns("H:H").ColumnWidth = 44
    Columns("I:I").ColumnWidth = 13.17
    Columns("J:J").ColumnWidth = 13.67
    Columns("K:K").ColumnWidth = 13.33
    Columns("L:L").ColumnWidth = 19.83
    Columns("M:M").ColumnWidth = 22
    Columns("N:N").ColumnWidth = 22
    Columns("O:O").ColumnWidth = 26
    Columns("P:P").ColumnWidth = 13.17
    Columns("Q:Q").ColumnWidth = 19.83
    Range("I1:Q1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub

