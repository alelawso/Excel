# Excel VBA

# VBA formula for automatically refreshing external stock data whenever the workbook is opened.

Private Sub Workbook_Open()
'PURPOSE: Run Data tab's Refresh All function

ThisWorkbook.RefreshAll

MsgBox "Stock Data has been refreshed!" 'Optional

End Sub

Credit: https://www.thespreadsheetguru.com/blog/add-real-time-stock-prices-and-metrics-to-excel

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

# Cleaning up Data

Public Sub CleanUpData()
    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
            Worksheets(i).Select
            
            AddHeaders
            FormatHeader
            
            i = i + 1
    Loop
End Sub
Sub AddHeaders()
'
' AddHeaders Macro
' Places headers on worksheet
'

'
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "REGION"
    Rows("1:1").Select
    Range("B1").Activate
    ActiveCell.FormulaR1C1 = "C"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "C"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "CATEGORY"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "JAN"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "FEB"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "MAR"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("D4").Select
End Sub
Sub FormatHeader()
'
' FormatHeader Macro
' Formats the first row headers"&chr(13)&"
'

'
    Range("A1:F1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub

# Inserting & Formatting Text
Sub AddHeaders()
'
' AddHeaders Macro
' Automate adding headers to worksheet
'
' Keyboard Shortcut: Ctrl+j
'
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
'
' Sort List by Division Ascending
'

'
    Selection.Sort Key1:=Range("A4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub

Sub CategorySort()
'
' Sort List by Category Ascending
'

'
    Selection.Sort Key1:=Range("B4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub

Sub TotalSort()
'
' Sort List by Total Sales Ascending
'

'
    Selection.Sort Key1:=Range("F4"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

End Sub

