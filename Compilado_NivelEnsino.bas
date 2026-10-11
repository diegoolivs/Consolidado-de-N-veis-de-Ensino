Attribute VB_Name = "Módulo1"
Sub Atualizar_base()
 
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
 
'Variáveis
Dim este As Workbook
Dim base As Workbook
Set este = ThisWorkbook
'Definindo usuário para caminhos
este.Sheets("Apoio").Cells(2, 5) = InputBox("Informe o usuário\caminho:", _
"Ajuste de Caminho")
If este.Sheets("Apoio").Cells(2, 5) = "Diego" Then
    caminho = Planilha3.Range("C1").Value
Else
    MsgBox "Erro de usuário para caminho!", vbOKOnly, "Erro de caminho"
    Exit Sub
End If

'Limpando o espaço que irá receber as bases
este.Sheets("Base").Range("A1").CurrentRegion.ClearContents

'Acessar caminho Sistemas EAD
Set base = Workbooks.Open(caminho)
base.Sheets(1).Range("A:B").Copy
este.Sheets(1).Range("A1").PasteSpecial
base.Sheets(1).Range("D:E").Copy
este.Sheets(1).Range("C1").PasteSpecial
'Application.CutCopyMode = False 'Desativa área de transferencia
base.Close
'Atribuindo nível de ensino
linha_este = este.Sheets("Base").Range("A1").CurrentRegion.Rows.Count
este.Sheets("Base").Cells(1, 5).Value = "Nivel_Ensino"
este.Sheets("Base").Range(Cells(2, 5), Cells(linha_este, 5)).Value = "EAD"

'Acessar caminho Sistemas SEMI
caminho = Planilha3.Range("C2").Value
Set base = Workbooks.Open(caminho)
linha_este = este.Sheets("Base").Range("A1").CurrentRegion.Rows.Count + 1
linha_base = base.Sheets(1).Range("A1").CurrentRegion.Rows.Count
base.Sheets(1).Range(Cells(2, 1), Cells(linha_base, 2)).Copy
este.Sheets("Base").Cells(linha_este, 1).PasteSpecial
base.Sheets(1).Range(Cells(2, 4), Cells(linha_base, 5)).Copy
este.Sheets("Base").Cells(linha_este, 3).PasteSpecial
base.Close
'Atribuindo nível de ensino
linha_este = este.Sheets("Base").Range("A1").CurrentRegion.Rows.Count
pri_linha = Sheets("Base").Cells(linha_este, 5).End(xlUp).Row
este.Sheets("Base").Range(Cells(pri_linha + 1, 5), Cells(linha_este, 5)).Value = "SEMI"

'Acessar caminho Sistemas PRESENCIAL
caminho = Planilha3.Range("C3").Value
Set base = Workbooks.Open(caminho)
linha_este = este.Sheets("Base").Range("A1").CurrentRegion.Rows.Count + 1
linha_base = base.Sheets(1).Range("A1").CurrentRegion.Rows.Count
base.Sheets(1).Range(Cells(2, 1), Cells(linha_base, 2)).Copy
este.Sheets("Base").Cells(linha_este, 1).PasteSpecial
base.Sheets(1).Range(Cells(2, 4), Cells(linha_base, 5)).Copy
este.Sheets("Base").Cells(linha_este, 3).PasteSpecial
base.Close
'Atribuindo nível de ensino
linha_este = este.Sheets("Base").Range("A1").CurrentRegion.Rows.Count
pri_linha = Sheets("Base").Cells(linha_este, 5).End(xlUp).Row
este.Sheets("Base").Range(Cells(pri_linha + 1, 5), Cells(linha_este, 5)).Value = "PRESENCIAL"

'Ajustes
Sheets("Base").Activate
este.RefreshAll 'Atualizar tabelas dinâmicas
este.Sheets("Base").Range("A1:E1").Font.Bold = True 'Deixar cabeçalho negrito


Application.DisplayAlerts = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True


MsgBox "Atualização Concluída!"

End Sub







































