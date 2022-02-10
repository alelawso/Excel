# Excel VBA

# Inserting & Formatting Text
    Sub AddHeaders()
    'AddHeaders Macro'
    'Automate adding headers to worksheet'
    'Keyboard Shortcut: Ctrl+j'
        
        Range("A3").Select
        ActiveCell.FormulaR1C1 = "Region"
        Range("B3").Select
        ActiveCell.FormulaR1C1 = "Expense"
        Range("C3").Select
        ActiveCell.FormulaR1C1 = "Jan"
        Range("D3").Select
        ActiveCell.FormulaR1C1 = "Feb"
        Range("E3").Select
        ActiveCell.FormulaR1C1 = "March"
        Range("F3").Select
        ActiveCell.FormulaR1C1 = "Totals"
        Range("A3:F3").Select
        With Selection.Font
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        ActiveWindow.Zoom = 100
        ActiveWindow.Zoom = 85
    End Sub

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
