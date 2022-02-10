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
