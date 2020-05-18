Attribute VB_Name = "Module1"
Private Static Function GetAgg(Arr()) As Variant

' Input: Array of 7D shape
' Output:

Dim i, Count As Long, ID As Long, Unique(), D1(), D2(), OV(), CV(), CUMSUM()

Count = 0

ReDim Unique(Count), D1(Count), D2(Count), OV(Count), CV(Count), CUMSUM(Count)

For i = LBound(Arr) To UBound(Arr)
'Note For loop is INCLUSIVE

    If InArray(Arr(i, 1), Unique) = False Then
    
        Unique(Count) = Arr(i, 1) '<-Added Ticker to the reserved place
        D1(Count) = Arr(i, 2) '<- D1 keeps track of the min date
        OV(Count) = Arr(i, 3) '<-OV Keeps track of the Opening stock value (OV is the one corresponds to the min date)
        D2(Count) = Arr(i, 2) '<-D2keeps track of the max date
        CV(Count) = Arr(i, 6) '<-CV Keeps track of the Closing stock value (CV is the one corresponds to the max date)
        CUMSUM(Count) = Arr(i, 7) '<-CUMSUM keeps track of cumulative sum of volumes
        
        Count = Count + 1
        
        ReDim Preserve Unique(Count), D1(Count), D2(Count), OV(Count), CV(Count), CUMSUM(Count)
        '<-Next value reservation
        
       
    Else
    
    'Means, the ticker was seen before
    'Finds the index(ID) of the ticker in the Unique array, all of the 6 arrays are in sync by index
        ID = WhereInArray(Arr(i, 1), Unique) ' First exact match, No duplicates in Unique array, Hence no problem
        
        If D1(ID) > Arr(i, 2) Then 'If a date is found which is less than the tracked open  date
            D1(ID) = Arr(i, 2)
            OV(ID) = Arr(i, 3)
        ElseIf D2(ID) < Arr(i, 2) Then 'If a date is found which is greater than the tracked close date
            D2(ID) = Arr(i, 2)
            CV(ID) = Arr(i, 6)
        End If
        
        'CUMSUM needs to be updated always as it is tracking total transaction volumes
        CUMSUM(ID) = CUMSUM(ID) + Arr(i, 7)
        
    
    End If
    
    

Next i
    

GetAgg = Array(Unique, D1, OV, D2, CV, CUMSUM)

End Function

Private Static Function InArray(val As Variant, Arr() As Variant) As Boolean
'Check if an element is in an array
'Input: Value and Array
'Output: Boolean
'Author: Bincy Narath

    Dim Element As Variant
    
    For Each Element In Arr
    
        If Element = val Then
            InArray = True
            Exit Function
        End If
        
    Next Element
    InArray = False
End Function

Private Static Function WhereInArray(val As Variant, Arr() As Variant) As Long
'Check the index starts with zero of an element is in an array knowing it is in the array
'Input: Value and Array
'Output: Index as Long
'Author: Bincy Narath

    Dim Element As Variant, Index As Long
    Index = 0
    For Each Element In Arr
    
        If Element = val Then
            WhereInArray = Index
            Exit Function
        End If
        
        
        Index = Index + 1
    Next Element
End Function
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


Sub Stock_Analysis()

    'Testing is done on Unit 02 - VBA_Homework_Instructions_Resources_alphabetical_testing.xlsm
    'All the references in the comments are to be related to the same file

    

    
    
    'Declarations
    Dim ws As Worksheet, Data(), TotalRows As Long, ErrCount As Long, TotalColumns As Integer, PercChange As Double, AGG(), Ticker(0 To 2) As String, Value(0 To 2) As Double, DivByZero()
    
    

    For Each ws In ThisWorkbook.Worksheets
        
        'Activate a specific sheet
        Worksheets(ws.Name).Activate
        Debug.Print (ws.Name) '-> Print the sheetname in Immdediate window
        
    
        'Reused variables here
        TotalRows = Range("A2").End(xlDown).Row
        TotalColumns = Range("A1").End(xlToRight).Column
        
        Data = Range(Cells(2, 1), Cells(TotalRows, TotalColumns)).Value
        'Data is stored as a (1-70925,1-7) dimentional array
        'MsgBox (Str(Data(1, 2)) + " " + Str(Data(70925, 7)))
        
        'MsgBox (LBound(Data)) = 1
        'MsgBox (UBound(Data)) = 70925
        
        AGG = GetAgg(Data)
        Unique = AGG(0)
        'Unique, D1, OV, D2, CV, CUMSUM
    
        'Adding the label
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        ErrCount = 0
        
        
        For i = LBound(Unique) To UBound(Unique) - 1 '-1 because the Output of GetAgg has an empty place at the end for all variables
        
            Cells(2 + i, 9).Value = AGG(0)(i)
            Cells(2 + i, 10).Value = AGG(4)(i) - AGG(2)(i) 'CV-OV
            'Color Conditional Formatting for "Yearly Change"
            If Cells(2 + i, 10).Value > 0 Then
                Cells(2 + i, 10).Interior.ColorIndex = 4 'Green
            ElseIf Cells(2 + i, 10).Value < 0 Then
                Cells(2 + i, 10).Interior.ColorIndex = 3 'Red
            Else
                Cells(2 + i, 10).Interior.ColorIndex = xlNone 'No Color
            End If
            
            If AGG(2)(i) = 0 Then
                PercChange = 0
                ReDim DivByZero(ErrCount)
                If InArray(AGG(2)(i), DivByZero) = False Then
                    DivByZero(ErrCount) = AGG(0)(i)
                    ErrCount = ErrCount + 1
                End If
            Else
                PercChange = (AGG(4)(i) - AGG(2)(i)) / AGG(2)(i) '(OV-CV)/OV -> Need be formatted as percentage later
            End If
            Cells(2 + i, 11).Value = PercChange
            
            Cells(2 + i, 12).Value = AGG(5)(i)
            
            
            
            'Ticker keeps track of tickers corresponding to "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Value" at indexes 0, 1 and 2 respectively
            
            'Value keeps track of values corresponding to "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Value" at indexes 0, 1 and 2 respectively
                
               
             If i = 0 Then
             'Initialize to the first value
                Ticker(0) = AGG(0)(i)
                Ticker(1) = AGG(0)(i)
                Ticker(2) = AGG(0)(i)
                Value(0) = PercChange
                Value(1) = PercChange
                Value(2) = AGG(5)(i)
             Else
                If PercChange > Value(0) Then 'Greatest % increase tracking
                    Value(0) = PercChange
                    Ticker(0) = AGG(0)(i)
                End If
                
                If PercChange < Value(1) Then 'Greatest % decrease tracking
                    Value(1) = PercChange
                    Ticker(1) = AGG(0)(i)
                End If
                
                If AGG(5)(i) > Value(2) Then
                    Value(2) = AGG(5)(i)
                    Ticker(2) = AGG(0)(i)
                End If
             End If
             
             
    
        Next i
        
        'Fill the results of toppers and floppers
        Cells(2, 16).Value = Ticker(0)
        Cells(3, 16).Value = Ticker(1)
        Cells(4, 16).Value = Ticker(2)
        
        Cells(2, 17).Value = Value(0)
        Cells(3, 17).Value = Value(1)
        Cells(4, 17).Value = Value(2)
        
        'Fill the Error Tickers where OV is 0 (Div by Zero)
        If ErrCount > 0 Then
            Cells(6, 15).Value = "Error Tickers(OV is 0)"
            Cells(6, 15).Interior.ColorIndex = 6
            For i = LBound(DivByZero) To UBound(DivByZero)
                Cells(7 + i, 15).Value = DivByZero(i)
                Cells(7 + i, 15).Interior.ColorIndex = 6
            Next i
        End If
        
        
        
        'Number and  Column width formatting
        Call Formatting
        
        
    Next

End Sub

