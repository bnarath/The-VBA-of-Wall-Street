Attribute VB_Name = "Module2"
Sub Clear()


Dim ws As Worksheet
    
    

    For Each ws In ThisWorkbook.Worksheets
        
        'Activate a specific sheet
        Worksheets(ws.Name).Activate
        
        Range("I:Q").Value = ""
        Range("I:Q").Interior.ColorIndex = xlNone
        
        
    Next
    
End Sub

