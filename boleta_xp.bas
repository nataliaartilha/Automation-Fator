Attribute VB_Name = "Módulo14"
Sub Todas_XP()
Call AbreMaisRecenteNovo_e_copia_e_cola_xp
Call copiar_colar_xp
Call colar_di_xp
Call SalvarAba_xp
Call Enviar_email_XP
Call copiar_python
End Sub





Sub AbreMaisRecenteNovo_e_copia_e_cola_xp()
Application.ScreenUpdating = False
'Applicationd.DisplayAlerts = False

  Dim arqSys As FileSystemObject
  Dim objArq As File
  Dim minhaPasta
  Dim nomearq As String
  Dim dataArq As Date
Workbooks("Captação CDB - Calculadora.nova versao").Activate
Worksheets("CALCULADORA").Range("E2:M9").ClearContents
Worksheets("XP").Range("A1:N100").ClearContents
        Const Diret As String = "C:\Users\nartilha\Downloads\"
        Set arqSys = New FileSystemObject
        Set minhaPasta = arqSys.GetFolder(Diret)
        dataArq = DateSerial(1900, 1, 1)
For Each objArq In minhaPasta.Files
    If objArq.DateLastModified > dataArq Then
        dataArq = objArq.DateLastModified
        nomearq = objArq
    End If
Next objArq
        ActiveWorkbook.FollowHyperlink Address:=nomearq
        Set arqSys = Nothing
        Set minhaPasta = Nothing
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy

Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("XP").Activate
Range("A1").PasteSpecial

Application.CutCopyMode = False
Application.ScreenUpdating = True
End Sub

Sub copiar_colar_xp()
Application.ScreenUpdating = False
Dim linha As Integer

linha = 2


'data vencimento
While Worksheets("XP").Cells(linha, 8).Value <> ""
    Worksheets("XP").Cells(linha, 8).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 8).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'taxa cliente
While Worksheets("XP").Cells(linha, 10).Value <> ""
    Worksheets("XP").Cells(linha, 10).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 9).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'taxa emissão
While Worksheets("XP").Cells(linha, 9).Value <> ""
    Worksheets("XP").Cells(linha, 9).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 10).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'quantidade
While Worksheets("XP").Cells(linha, 12).Value <> ""
    Worksheets("XP").Cells(linha, 12).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 7).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'DI
While Worksheets("XP").Cells(linha, 11).Value <> ""
    Worksheets("XP").Cells(linha, 11).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 11).PasteSpecial
    linha = linha + 1
Wend
linha = 2
'PU
While Worksheets("XP").Cells(linha, 13).Value <> ""
    Worksheets("XP").Cells(linha, 13).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 12).PasteSpecial
    linha = linha + 1
Wend

linha = 2
'fazer a contraparte
Worksheets("XP").Cells(1, 15) = "Contraparte"
While Worksheets("XP").Cells(linha, 1).Value <> ""
    Worksheets("XP").Cells(linha, 15) = "XP"
    linha = linha + 1
Wend
linha = 2
'contraparte
While Worksheets("XP").Cells(linha, 15).Value <> ""
    Worksheets("XP").Cells(linha, 15).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 5).PasteSpecial
    linha = linha + 1
Wend
'cdi
linha = 2
While Worksheets("XP").Cells(linha, 3).Value <> ""
    Worksheets("XP").Cells(linha, 3).Copy
    Workbooks("Captação CDB - Calculadora.nova versao").Worksheets("CALCULADORA").Cells(linha, 6).PasteSpecial
    linha = linha + 1
Wend
Application.ScreenUpdating = True

End Sub

Sub colar_di_xp()

Application.CutCopyMode = True
linha1 = 2
linha2 = 1
While Worksheets("CALCULADORA").Cells(linha1, 4).Value <> ""
    Worksheets("CALCULADORA").Cells(1, 2).Value = Worksheets("CALCULADORA").Cells(linha1, 4).Value
    Worksheets("CALCULADORA").Cells(15, 2).Value = Worksheets("CALCULADORA").Cells(linha1, 11).Value
    Worksheets("CALCULADORA").Cells(linha1, 14).Value = Worksheets("CALCULADORA").Cells(19, 2).Value
    linha1 = linha1 + 1
    linha2 = linha2 + 1
Wend
Application.CutCopyMode = False
End Sub
'exporta aba e exlui as macros
Sub SalvarAba_xp()
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
.Worksheets("XP").Columns("Q:S").Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & "boleta_xp" & ".xlsx"
'Fecha o novo arquivo
'Workbooks("boleta_agora").Columns("T:Z").Delete
.Close SaveChanges:=True
End With


'Workbooks.Open "G:\depto\RENDA\Natalia Artilha\boleta_agora.xlsx"
'Columns("T:Z").Delete
'Workbooks("boleta_agora").Close SaveChanges:=True

'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False
End Sub
Sub Enviar_email_XP()
Dim txtFileName, nomearq, nomeRel, nomeemail As String
Dim saudacao As String


'Range(Selection, Selection.End(xlToRight)).Select
'Range(Selection, Selection.End(xlDown)).Select
'tabela = Selection

If Hour(Now) < 12 Then
saudacao = "Bom dia."
ElseIf Hour(Now) >= 12 And Hour(Now) <= 18 Then
saudacao = "Boa tarde, prezados!"
ElseIf Hour(Now) > 18 Then
saudacao = "Boa noite, prezados!"
End If




nomeemail = "OPERAÇÕES XP - [BANCO FATOR S/A] - " & Format(Worksheets("CALCULADORA").Range("B7"), "dd.mm.yyyy")



Diretorio = "G:\depto\RENDA\Natalia Artilha\"



Set objeto_outlook = CreateObject("Outlook.Application")
Set Email = objeto_outlook.createitem(0)
nomeRel = "boleta_xp"


With Email
.display
.To = "marcelo.felipe@xpi.com.br;BancoFatorTesouraria@fator.com.br"

.cc = "liquidacao.rf@xpi.com.br;rfmesaclientes@xpi.com.br"
.Subject = nomeemail
.HTMLBody = saudacao & Chr(12) & Chr(12) & "Operação realizada!" & Chr(12) & Chr(12) & "Segue PU no arquivo em anexo." & Chr(12) & Chr(12) & "Atenciosamente," & .HTMLBody
.Attachments.Add (Diretorio & nomeRel & ".xlsx")
'Email.send
End With



End Sub

