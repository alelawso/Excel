# Automate Sum Funcion

    Public Sub AutomateSum()

        Dim lastCell As String
        Dim i As Integer
    
        i = 1
    
        Do While i <= Worksheets.Count
            Worksheets(i).Select
    
            'selects F2 cell of active sheet'
            Range("F2").Select
    
            'Selected cell goes down the column to the end'
            Selection.End(xlDown).Select
    
            'Variable for last cell that stores the address'
            lastCell = ActiveCell.Address(False, False)
    
            'You want to go down one row and no columns and select the cell'
            ActiveCell.Offset(1, 0).Select
            'The actual formula the sum from F2 to the last cell in the column'
            ActiveCell.Value = "=SUM(F2:" & lastCell & ")"
        
            i = i + 1
        Loop
    End Sub
