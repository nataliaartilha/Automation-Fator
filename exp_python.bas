Attribute VB_Name = "Módulo12"


Sub copiar_python()
Dim linha, ultimalinha As Integer
linha = 2
Worksheets("PYTHON").Activate
'Worksheets("PYTHON").Range("B2:I100").ClearContents
ultimalinha = Sheets("PYTHON").Cells(Rows.Count, 2).End(xlUp).Row

While Worksheets("CALCULADORA").Cells(linha, 5) <> ""
    'Worksheets("python").Range("A"&ultimalinha) = Worksheets("CALCULADORA").Range("D"&linha:"L"&linha).Value
    Worksheets("CALCULADORA").Activate
    Worksheets("CALCULADORA").Cells(linha, 5).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Worksheets("PYTHON").Activate
    
    Worksheets("PYTHON").Cells(ultimalinha + 1, 2).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    linha = linha + 1
    ultimalinha = ultimalinha + 1
Wend
End Sub

Sub exportar_to_python()
'Impede que o Excel atualize a tela
Application.ScreenUpdating = False
'Impede que o Excel exiba alertas
Application.DisplayAlerts = False

'Seta uma variável para se referir a nova pasta de trabalho
Dim NovoWB As Workbook
'Cria esta nova aba
Set NovoWB = Workbooks.Add(xlWBATWorksheet)
With NovoWB
'Copia a aba atual para o novo arquivo, como a segunda aba
ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count)
'Deleta a primeira aba do arquivo criado (Aba em branco)
.Worksheets(1).Delete
'.Worksheets("agora").Columns("T:Z").Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & "excel_python" & ".xlsx"
'Fecha o novo arquivo
'Workbooks("boleta_agora").Columns("T:Z").Delete
.Close SaveChanges:=True
End With
Worksheets("PYTHON").Activate
'Worksheets("PYTHON").Range("B2:I100").ClearContents

'Workbooks.Open "G:\depto\RENDA\Natalia Artilha\boleta_agora.xlsx"
'Columns("T:Z").Delete
'Workbooks("boleta_agora").Close SaveChanges:=True

'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub
