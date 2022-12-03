Attribute VB_Name = "Module1"
Sub Search_Click()
    Dim colRange As String
    Dim colMax As String
    Dim colMaxNumber As Integer
    Dim colPosition As String
    Dim counter As Integer
    Dim startRow As Integer
    Dim startRowPosition As String
    
    colRange = Cells(2, 1) & ":" & Cells(2, 1)
    colMax = Cells(2, 1) & "65535"
    For i = 3 To Worksheets.Count
        Cells(i - 2, 3) = Sheets(i).Name
        On Error Resume Next
        Sheets(i).Range (col)
        colMaxNumber = Sheets(i).Range(colMax).End(xlUp).Row
        Cells(i - 2, 4) = colMaxNumber
        
        counter = 0
        For j = 2 To colMaxNumber
            colPosition = Cells(2, 1) & j
            If Sheets(1).Cells(2, 2).Value = Sheets(i).Range(colPosition).Value Then
                counter = counter + 1
                Sheets(i).Cells(j, 26) = counter
            Else
                 Sheets(i).Cells(j, 26) = 0
            End If
            Cells(i - 2, 5) = counter
        Next j
    Next i
    
    startRow = Sheets(2).Range(colMax).End(xlUp).Row + 1
    For i = 3 To Worksheets.Count
        colMaxNumber = Cells(i - 2, 5)
        For k = 1 To colMaxNumber
            colPosition = Application.WorksheetFunction.Match(k, Sheets(i).Range("Z:Z"), 0) & ":" & Application.WorksheetFunction.Match(k, Sheets(i).Range("Z:Z"), 0)
            startRowPosition = "A" & startRow
            Worksheets(i).Range(colPosition).Copy _
                Destination:=Sheets(2).Range(startRowPosition)
            startRow = startRow + 1
        Next k
    Next i
    
End Sub
