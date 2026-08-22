Attribute VB_Name = "M�dulo1"
Sub sbCria_Diario_dez_aulas()
    
    Dim wbNovo_Diario As Workbook
    Dim wbDados As Workbook
    Set wbNovo_Diario = Workbooks.Add
    Set wbDados = Workbooks("Confirmados Pr�xima rodada.xlsm")
    
    
    sbCabecalho_Dez_Aulas
    sbFormata_Tabela_Chamada
    sbAjusta_Coluna_Chamada
    sbFormata_Tabela_Avaliacao
    sbAjusta_Coluna_Avaliacao
    sbFormulas_Chamada
    sbFormulas_Avaliacao
    sbAjustes_Finais
    sbFormata_Situa��o
    Call sbCopia_Dados(wbNovo_Diario)
    
    Dim wsDados As Worksheet
    Set wsDados = wbDados.ActiveSheet
    
    If wsDados.Range("E2") = "DIGITA��O" Then
        sbAjusta_Aulas_Digitacao
    End If
    
    
End Sub
Sub sbCopia_Dados(wbDiario As Workbook)
    
    'Define as vari�veis do programa
    Dim wbConfirmados As Workbook
    Dim wsNomes As Worksheet
    Dim wsChamada As Worksheet
    Dim wsAvaliacao As Worksheet
    Dim tblChamada As ListObject
    Dim tblAvaliacao As ListObject
    Dim strOficina As String
    Dim strDias As String
    Dim strHorario As String
    Dim strTurno As String
    Dim lTurma As Long
    Dim dInicio As Date
    Dim dFim As Date
    
    
    'Atribui as vari�veis das pastas de trabalho
    Set wbDiario = ActiveWorkbook
    Set wbConfirmados = Workbooks("confirmados-proxima-rodada.xlsm")
    
    'Atribui as vari�veis das planilhas
    Set wsChamada = wbDiario.Sheets("CHAMADA")
    Set wsAvaliacao = wbDiario.Sheets("AVALIA��ES")
    Set wsNomes = wbConfirmados.ActiveSheet
    
    'Atribui as Tabelas
    Set tblAvaliacao = wsAvaliacao.ListObjects("AVALIACAO")
    Set tblChamada = wsChamada.ListObjects("CHAMADA")
    
   
    strOficina = wsNomes.Range("E2")
    strDias = wsNomes.Range("E3")
    strHorario = wsNomes.Range("E4")
    strTurno = wsNomes.Range("E5")
    lTurma = wsNomes.Range("E6")
    dInicio = wsNomes.Range("G2")
    dFim = wsNomes.Range("H2")
    
    MsgBox "Dados copiados para: " & wbDiario.Name
    MsgBox "Oficina: " & strOficina
    MsgBox "Dias: " & strDias
    MsgBox "Hor�rio: " & strHorario
    MsgBox "Turno: " & strTurno
    MsgBox "Turma: " & lTurma

    wsAvaliacao.Range("C5") = strOficina
    wsChamada.Range("E6") = dInicio
    wsChamada.Range("W6") = dFim
    wsChamada.Range("R3") = strDias
    wsChamada.Range("R4") = strHorario
    wsChamada.Range("D3") = strTurno
    wsChamada.Range("O3") = lTurma
    wsChamada.Activate
    
    wbConfirmados.Activate
    Range("A2:B16").Select
    Selection.Copy
    wbDiario.Activate
    Sheets("AVALIA��ES").Select
    Range("B7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    Call sbRenomeia_Tabelas(tblChamada, tblAvaliacao, strTurno, lTurma, strOficina)
   
End Sub



Sub sbRenomeia_Tabelas(tblChamada As ListObject, tblAvaliacao As ListObject, strTurno As String, lTurma As Long, strOficina As String)

    Dim siglaTurno As String
    Dim siglaOficina As String
    Dim nomeTabelaChamada As String
    Dim nomeTabelaAvaliacao As String

    ' Define sigla do turno
    Select Case UCase(Trim(strTurno))
        Case "MANH�": siglaTurno = "M"
        Case "TARDE": siglaTurno = "T"
    End Select

    ' Converte oficina para sigla
    siglaOficina = OficinaParaSigla(strOficina)

    ' Formata n�mero da turma com dois d�gitos
    Dim turmaFormatada As String
    turmaFormatada = Format(lTurma, "00")

    ' Monta os nomes
    nomeTabelaChamada = "CH" & siglaTurno & turmaFormatada & siglaOficina
    nomeTabelaAvaliacao = "AV" & siglaTurno & turmaFormatada & siglaOficina

    ' Renomeia as tabelas
    tblChamada.Name = nomeTabelaChamada
    tblAvaliacao.Name = nomeTabelaAvaliacao

End Sub

Function OficinaParaSigla(oficina As String) As String
    Select Case UCase(Trim(oficina))
        Case "DIGITA��O": OficinaParaSigla = "DG"
        Case "WINDOWS B�SICO": OficinaParaSigla = "WB"
        Case "NO��ES DE WORD": OficinaParaSigla = "NW"
        Case "POWERPOINT": OficinaParaSigla = "PP"
        Case "EXCEL B�SICO": OficinaParaSigla = "EB"
        Case "EXCEL AVAN�ADO": OficinaParaSigla = "EA"
        Case "SMARTPHONE": OficinaParaSigla = "SP"
        Case "REDES SOCIAIS": OficinaParaSigla = "RS"
        Case "CANVA": OficinaParaSigla = "CV"
        Case "PYTHON": OficinaParaSigla = "PY"
    End Select
End Function

Private Sub sbCabecalho_Dez_Aulas()
Attribute sbCabecalho_Dez_Aulas.VB_ProcData.VB_Invoke_Func = " \n14"

    'Aplica a formata��o do cabe�alho do di�rio sem aplicar as tabelas
    Range("A1:X2").Select
    Selection.Merge
    Range("B3:C3").Select
    Selection.Merge
    Range("B4:C4").Select
    Selection.Merge
    Range("D3:I4").Select
    Selection.Merge
    Range("J3:N4").Select
    Selection.Merge
    Range("O3:Q4").Select
    Selection.Merge
    Range("R3:X3").Select
    Selection.Merge
    Range("R4:X4").Select
    Selection.Merge
    Range("B3:C3").Select
    ActiveCell.FormulaR1C1 = "PROFESSOR"
    Range("B4:C4").Select
    ActiveCell.FormulaR1C1 = "PEDRO VOLINO"
    Range("D3:I4").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="MANH�,TARDE"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("J3:N4").Select
    ActiveCell.FormulaR1C1 = "TURMA"
    Range("R3:X3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="TER�AS E QUINTAS,SEGUNDAS E QUARTAS"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("R4:X4").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "08:00 - 10:00,10:00 - 12:00, 13:00 - 15:00,15:00 - 17:00"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("A5:A6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "N�"
    Range("B5:B6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "NOME"
    Range("C5:C6").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "DATA DE NASCIMENTO"
    Range("D5:D6").Select
    Selection.Merge
    ActiveCell.FormulaR1C1 = "IDADE"
    Range("E5").Select
    ActiveCell.FormulaR1C1 = "AULA 1"
    Range("E5:F5,G5:H5,I5:J5,K5:L5,M5:N5,O5:P5,Q5:R5,S5:T5,U5:V5,W5:X5").Select
    Range("W5").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("E5:F5").Select
    Selection.AutoFill Destination:=Range("E5:X5"), Type:=xlFillDefault
    Range("E5:X5").Select
    Range("E6:F6,G6:H6,I6:J6,K6:L6,M6:N6,O6:P6,Q6:R6,S6:T6,U6:V6,W6:X6").Select
    Range("W6").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveSheet.Select
    ActiveSheet.Name = "CHAMADA"
    Range("C23").Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "AVALIA��ES"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "PROFESSOR"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "PEDRO VOLINO"
    Range("C2:E2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "OFICINA"
    Range("C3:E3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "DIGITA��O,WINDOWS B�SICO,NO��ES DE WORD,POWERPOINT,EXCEL B�SICO,EXCEL AVAN�ADO,SMARTPHONE,REDES SOCIAIS,CANVA,PYTHON"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("F2:G2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "CARGA HOR�RIA"
    Range("F3:G3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "20H"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "IN�CIO"
    Range("I3").Select
    ActiveCell.FormulaR1C1 = "FIM"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=CHAMADA!R6C5"
    Range("J3").Select
    ActiveCell.FormulaR1C1 = "=CHAMADA!R6C23"
    Rows("1:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.ColumnWidth = 2.38
    Columns("B:B").Select
    Selection.ColumnWidth = 37.88
    Columns("C:C").Select
    Selection.ColumnWidth = 18.75
    Columns("D:D").Select
    Selection.ColumnWidth = 5.38
    Columns("E:E").Select
    Selection.ColumnWidth = 11.88
    Columns("F:H").Select
    Selection.ColumnWidth = 8.88
    Columns("I:I").Select
    Selection.ColumnWidth = 15.13
    Columns("J:J").Select
    Selection.ColumnWidth = 12.5
    Range("A1:J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 270
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 13382400
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 270
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 13382400
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 8527616
        .TintAndShade = 0
    End With
    Range("A4:J5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13382400
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Size = 10
    Range("A4:B5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C4:E5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("F4:G5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("H4:J5").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B11").Select
    Sheets("CHAMADA").Select
    Range("A1:X2").Select
    With Selection.Interior
        .Pattern = xlPatternLinearGradient
        .Gradient.Degree = 270
        .Gradient.ColorStops.Clear
    End With
    With Selection.Interior.Gradient.ColorStops.Add(0)
        .Color = 13382400
        .TintAndShade = 0
    End With
    With Selection.Interior.Gradient.ColorStops.Add(1)
        .Color = 8527616
        .TintAndShade = 0
    End With
    ActiveCell.FormulaR1C1 = "=IF(AVALIA��ES!R5C3="""","""",AVALIA��ES!R[4]C[2])"
    ActiveCell.FormulaR1C1 = "=IF(AVALIA��ES!R5C3="""","""",AVALIA��ES!R5C3)"
    Range("A3:X6").Select
    Range("R3").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13382400
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Arial"
        .Size = 11
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
    Selection.Font.Size = 10
    Selection.Font.Bold = True
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Range("A1:X2").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A5:A6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B5:B6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("C5:C6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D5:D6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("E5:F6,G5:H6,I5:J6,K5:L6,M5:N6,O5:P6,Q5:R6,S5:T6,U5:V6,W5:X6").Select
    Range("W5").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A3:C4").Select
    Range("B3").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("D3:I4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("J3:Q4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("R3:X4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1:X6").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Range("C5:C6").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
        Range("A1:X2").Select
    With Selection.Font
        .Name = "Berlin Sans FB Demi"
        .Size = 26
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
    Selection.Font.Bold = True
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Sheets("AVALIA��ES").Select
        Range("A1:J3").Select
    ActiveCell.FormulaR1C1 = "ESPA�O NAVE INO�"
    Range("A1:J3").Select
    With Selection.Font
        .Name = "Berlin Sans FB Demi"
        .Size = 36
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    Selection.Font.Bold = True
    
End Sub

Private Sub sbFormata_Tabela_Chamada()
Attribute sbFormata_Tabela_Chamada.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("CHAMADA").Select
    Range("A7:X21").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$7:$X$21"), , xlNo).Name = _
        "CHAMADA"
    Range("CHAMADA[#All]").Select
    ActiveSheet.ListObjects("CHAMADA").TableStyle = "TableStyleLight9"
    ActiveSheet.ListObjects("CHAMADA").ShowHeaders = False
    Rows("7:7").Select
    Selection.Delete Shift:=xlUp
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "1"
    Selection.AutoFill Destination:=Range("CHAMADA[[#All],[Coluna1]]"), Type:= _
        xlFillSeries
    Range("CHAMADA[[#All],[Coluna1]]").Select
    Range("E7:E21,G7:G21,I7:I21,K7:K21,M7:M21,O7:O21,Q7:Q21").Select
    Range("Q7").Activate
    ActiveWindow.Zoom = 85
    ActiveWindow.Zoom = 70
    Range("E7:E21,G7:G21,I7:I21,K7:K21,M7:M21,O7:O21,Q7:Q21,S7:S21,U7:U21,W7:W21"). _
        Select
    Range("W7").Activate
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="P,F,J,R,RJ"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.FormatConditions.Add Type:=xlTextString, String:="P", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = -0.249946592608417
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="F", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16777088
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 7697919
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="J", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -13408615
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10092543
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="R", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -10477568
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
        Range("A7:A21,C7:X21").Select
    Range("C7").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("CHAMADA[[#All],[Coluna2]]").Select
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

Private Sub sbAjusta_Coluna_Chamada()
    Sheets("CHAMADA").Select
    Range("E5:X5").Select
    Selection.ColumnWidth = 3.5
    Range("D5:D6").Select
    Selection.ColumnWidth = 8.13
    Range("C5:C6").Select
    Selection.ColumnWidth = 15
    Range("B5:B6").Select
    Selection.ColumnWidth = 37.88
    Range("A5:A6").Select
    Selection.ColumnWidth = 2.88
End Sub

Private Sub sbFormulas_Chamada()
    Sheets("CHAMADA").Select
    Range("B7").Select
    ActiveCell.Formula = "=IF(AVALIACAO[@Nome]="""","""",AVALIACAO[@Nome])"
    Range("C7").Select
    ActiveCell.Formula = "=IF(AVALIACAO[@[Data de Nascimento]]="""","""",AVALIACAO[@[Data de Nascimento]])"
    Range("D7").Select
    ActiveCell.Formula = "=IF(AVALIACAO[@Idade]="""","""",AVALIACAO[@Idade])"
End Sub

Private Sub sbFormata_Tabela_Avaliacao()
Attribute sbFormata_Tabela_Avaliacao.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("AVALIA��ES").Select
    Range("A6:I20").Select
    Selection.Cut Destination:=Range("B6:J20")
    Range("A6:J20").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$6:$J$20"), , xlNo).Name = _
        "AVALIACAO"
    Range("AVALIACAO[#All]").Select
    ActiveSheet.ListObjects("AVALIACAO").TableStyle = "TableStyleLight9"
    Range("AVALIACAO[[#Headers],[Coluna1]]").Select
    ActiveCell.FormulaR1C1 = "N�"
    Range("AVALIACAO[[#Headers],[Coluna2]]").Select
    ActiveCell.FormulaR1C1 = "Nome"
    Range("AVALIACAO[[#Headers],[Coluna3]]").Select
    ActiveCell.FormulaR1C1 = "Data de Nascimento"
    Range("AVALIACAO[[#Headers],[Coluna4]]").Select
    ActiveCell.FormulaR1C1 = "Idade"
    Range("AVALIACAO[[#Headers],[Coluna5]]").Select
    ActiveCell.FormulaR1C1 = "Documentos"
    Range("AVALIACAO[[#Headers],[Coluna6]]").Select
    ActiveCell.FormulaR1C1 = "Presen�a"
    Range("AVALIACAO[[#Headers],[Coluna7]]").Select
    ActiveCell.FormulaR1C1 = "Av Aulas"
    Range("AVALIACAO[[#Headers],[Coluna8]]").Select
    ActiveCell.FormulaR1C1 = "Av Final"
    Range("AVALIACAO[[#Headers],[Coluna9]]").Select
    ActiveCell.FormulaR1C1 = "Aproveitamento"
    Range("AVALIACAO[[#Headers],[Coluna10]]").Select
    ActiveCell.FormulaR1C1 = "Situa��o"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "1"
    Selection.AutoFill Destination:=Range("AVALIACAO[N�]"), Type:=xlFillSeries
    Range("AVALIACAO[N�]").Select
    Range("D7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF([@[Data de Nascimento]]="""","""",DATEDIF([@[Data de Nascimento]],TODAY(),""y""))"
    Range("AVALIACAO[#Headers]").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13382400
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("C13").Select
    ActiveSheet.ListObjects("AVALIACAO").ShowAutoFilterDropDown = False
    Range("AVALIACAO[[#Headers],[N�]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Nome]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Documentos]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Presen�a]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Av Aulas]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Av Final]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Aproveitamento]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Situa��o]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Data de Nascimento]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("AVALIACAO[[#Headers],[Idade]]").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    ActiveSheet.ListObjects("AVALIACAO").ShowTableStyleColumnStripes = True
    Range("F7").Select
    ActiveCell.FormulaR1C1 = ""
    Range("AVALIACAO[Presen�a]").Select
    Selection.FormatConditions.AddDatabar
    Selection.FormatConditions(Selection.FormatConditions.Count).ShowValue = True
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1)
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
    End With
    With Selection.FormatConditions(1).BarColor
        .Color = 13012579
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).BarFillType = xlDataBarFillSolid
    Selection.FormatConditions(1).Direction = xlContext
    Selection.FormatConditions(1).NegativeBarFormat.ColorType = xlDataBarColor
    Selection.FormatConditions(1).BarBorder.Type = xlDataBarBorderNone
    Selection.FormatConditions(1).AxisPosition = xlDataBarAxisAutomatic
    With Selection.FormatConditions(1).AxisColor
        .Color = 0
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).NegativeBarFormat.Color
        .Color = 255
        .TintAndShade = 0
    End With
    Range("D32").Select
        Range("AVALIACAO[#All]").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("AVALIACAO[Nome]").Select
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

Private Sub sbAjusta_Coluna_Avaliacao()
    Sheets("AVALIA��ES").Select
    Range("A4").Select
    Selection.ColumnWidth = 2.88
    Range("B4").Select
    Selection.ColumnWidth = 37.88
    Range("AVALIACAO[[#Headers],[Data de Nascimento]]").Select
    Selection.ColumnWidth = 18.75
    Range("AVALIACAO[[#Headers],[Idade]]").Select
    Selection.ColumnWidth = 5.38
    Range("AVALIACAO[[#Headers],[Documentos]]").Select
    Selection.ColumnWidth = 11.89
    Range("AVALIACAO[[#Headers],[Presen�a]]").Select
    Selection.ColumnWidth = 8.88
    Range("I4").Select
    Selection.ColumnWidth = 15.13
    Range("J4").Select
    
    Selection.ColumnWidth = 12.5
End Sub

Private Sub sbFormulas_Avaliacao()
    Sheets("AVALIA��ES").Select
    Range("F7").Select
    ActiveCell.Formula = _
        "=IF([@Nome]="""","""",IF((COUNTIF(CHAMADA[@[Coluna5]:[Coluna24]],""P"")+COUNTIF(CHAMADA[@[Coluna5]:[Coluna24]],""R"")+COUNTIF(CHAMADA[@[Coluna5]:[Coluna24]],""RJ""))/10=0,"""",(COUNTIF(CHAMADA[@[Coluna5]:[Coluna24]],""P"")+COUNTIF(CHAMADA[@[Coluna5]:[Coluna24]],""R"")+COUNTIF(CHAMADA[@[Coluna5]:[Coluna24]],""RJ""))/10))"
    Range("G7").Select
    ActiveCell.Formula = _
        "=IF(CHAMADA[@Coluna6]="""","""",AVERAGE(CHAMADA[@Coluna6],CHAMADA[@Coluna8],CHAMADA[@Coluna10],CHAMADA[@Coluna12],CHAMADA[@Coluna14],CHAMADA[@Coluna16],CHAMADA[@Coluna18],CHAMADA[@Coluna20],CHAMADA[@Coluna24]))"
    Range("H7").Select
    ActiveCell.Formula = _
        "=IF(CHAMADA[@Coluna24]="""","""",CHAMADA[@Coluna24])"
    Range("I7").Select
    ActiveCell.Formula = _
        "=IF(AND([@[Av Aulas]]="""",[@[Av Final]]=""""),"""",([@[Av Aulas]]*0.2)+([@[Av Final]]*0.8))"
    Range("AVALIACAO[Situa��o]").Select
    ActiveWindow.SmallScroll Down:=-3
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="Aprovado,Reprovado,Desist�ncia"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Private Sub sbAjustes_Finais()
    Sheets("CHAMADA").Select
    Range("CHAMADA[[#All],[Coluna3]]").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("CHAMADA[[#All],[Coluna4]]").Select
    Selection.NumberFormat = "0"
    Range("CHAMADA[[Coluna5]:[Coluna24]]").Select
    Selection.NumberFormat = "0.0"
    Sheets("AVALIA��ES").Select
    Range("AVALIACAO[Data de Nascimento]").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("AVALIACAO[Idade]").Select
    Selection.NumberFormat = "0"
    Range("AVALIACAO[Presen�a]").Select
    Selection.Style = "Percent"
    Range("AVALIACAO[Av Aulas]").Select
    Selection.NumberFormat = "0.0"
    Range("AVALIACAO[[Av Final]:[Aproveitamento]]").Select
    Selection.NumberFormat = "0.0"
    Sheets("AVALIA��ES").Select
    Range("J4:J5").Select
    Selection.NumberFormat = "m/d/yyyy"
    Sheets("CHAMADA").Select
    Range("E6:X6").Select
    Selection.NumberFormat = "d-mmm"
    
    Sheets("AVALIA��ES").Select
    Range("AVALIACAO[[Nome]:[Situa��o]]").Select
    Application.CutCopyMode = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$J7=""Aprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$J7=""Reprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$J7=""Desist�ncia"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Sheets("CHAMADA").Select
    Range("CHAMADA[[#All],[Coluna2]:[Coluna24]]").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AVALIA��ES!$J7=""Aprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AVALIA��ES!$J7=""Reprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AVALIA��ES!$J7=""Desist�ncia"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Strikethrough = True
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub
Private Sub sbFormata_Situa��o()
' sbFormata_Situa��o Macro'


    Sheets("AVALIA��ES").Select
    Range("AVALIACAO[[Nome]:[Situa��o]]").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$J7=""Desist�ncia"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Strikethrough = True
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$J7=""Reprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .Pattern = xlNone
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=$J7=""Aprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Sheets("CHAMADA").Select
    Range("CHAMADA[[#All],[Coluna2]:[Coluna24]]").Select
    ActiveWindow.SmallScroll Down:=-3
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AVALIA��ES!$J7=""Desist�ncia"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Strikethrough = True
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AVALIA��ES!$J7=""Reprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Application.CutCopyMode = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=AVALIA��ES!$J7=""Aprovado"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = False
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Private Sub sbAjusta_Aulas_Digitacao()

    Dim wbDiario As Workbook
    Dim wbConfirmados As Workbook
    Dim wsNomes As Worksheet
    Dim wsChamada As Worksheet
    Dim wsAvaliacao As Worksheet
    Dim tblChamada As ListObject
    Dim tblAvaliacao As ListObject
    Dim strOficina As String
    Dim strDias As String
    Dim strHorario As String
        
    Sheets("CHAMADA").Select
    Range("A1:X4").Select
    Selection.UnMerge
    
    Range("J4") = Range("O3")
    
    Columns("M:V").Select
    Selection.Delete Shift:=xlToLeft
    
    Columns("E:N").Select
    Selection.ColumnWidth = 4
    
    Range("A1:N2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("D3:I4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("J3:K3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("J4:K4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("L3:N3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("L4:N4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("J3:K4").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Color = -1003520
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    
    Range("L3:N3").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="SEG E QUA,TER E QUI"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("L4:N4").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "08:00 - 10:00,10:00 - 12:00,13:00 - 15:00,15:00 - 17:00"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With


    
    'Atribui as vari�veis das pastas de trabalho
    Set wbDiario = ActiveWorkbook
    Set wbConfirmados = Workbooks("Confirmados Pr�xima rodada.xlsm")
    
    'Atribui as vari�veis das planilhas
    Set wsChamada = wbDiario.Sheets("CHAMADA")
    Set wsAvaliacao = wbDiario.Sheets("AVALIA��ES")
    Set wsNomes = wbConfirmados.ActiveSheet
    
   
    strOficina = wsNomes.Range("E2")
    strDias = wsNomes.Range("E3")
    strHorario = wsNomes.Range("E4")
    
    wsChamada.Range("L4") = strHorario
    
    If strDias = "SEGUNDAS E QUARTAS" Then
        strDias = "SEG E QUA"
        wsChamada.Range("L3") = strDias
    ElseIf strDias = "TER�AS E QUINTAS" Then
        strDias = "TER E QUI"
        wsChamada.Range("L3") = strDias
    End If
    
    
    
End Sub
