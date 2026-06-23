Attribute VB_Name = "FiltroAvancado"
Sub Geral()

    Range("F2").Select
    Selection.ClearContents
    Filtrar

End Sub
Sub Estoque()

    Range("F2").Select
    ActiveCell.FormulaR1C1 = ">0"
    Filtrar

End Sub

Sub Filtrar()
Attribute Filtrar.VB_ProcData.VB_Invoke_Func = " \n14"

    Application.ScreenUpdating = False
    
    Planilha1.CriarQueryTable
    
    ThisWorkbook.Sheets("Dados").Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=ThisWorkbook.Sheets("Pesquisa").Range("A1:G2"), _
        CopyToRange:=ThisWorkbook.Sheets("Pesquisa").Range("A6:G6"), _
        Unique:=False
    
    Cells.Select
    With Selection.Font
        .Name = "Consolas"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Font
        .Name = "Consolas"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("A7").Select
    
    Application.ScreenUpdating = True
               
End Sub

Sub Limpar()
    Range("A2:G2").Select
    Selection.ClearContents
    Range("D2").Select
End Sub
