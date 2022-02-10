# Sorting Records

    Public Sub UserSortInput()
    Dim userInput As String
    Dim promptMSG As String
    Dim tryAgain As Integer
    
    promptMSG = "Enter a numeric value to sort..." & vbCrLf & _
        "1 --- Sort by Division" & vbCrLf & _
        "2 --- Sort by Category" & vbCrLf & _
        "3 --- Sort by Total"
    
    userInput = InputBox(promptMSG)
    
    If userInput = "1" Then
        DivisionSort
    ElseIf userInput = "2" Then
        CategorySort
    ElseIf userInput = "3" Then
        TotalSort
    Else
        tryAgain = MsgBox("Invalid Value! Try again?", vbYesNo)
        
        If tryAgain = 6 Then
            UserSortInput
        End If
    End If
    
    End Sub

    Sub DivisionSort()
    ' Sort List by Division Ascending'

        Selection.Sort Key1:=Range("A4"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal

    End Sub

    Sub CategorySort()
    ' Sort List by Category Ascending'

        Selection.Sort Key1:=Range("B4"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal

    End Sub

    Sub TotalSort()
    'Sort List by Total Sales Ascending'

        Selection.Sort Key1:=Range("F4"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal

    End Sub
