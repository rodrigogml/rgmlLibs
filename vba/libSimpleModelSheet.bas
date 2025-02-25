Attribute VB_Name = "libSimpleModelSheet"
'libAppWorksheet - v1 - 28/11/2024

'Esta biblioteca concentra fun��es e m�todos que permitem a constru��o de uma planilha com controle de vers�o e importa��o de dados de uma planilha de vers�o enterior para uma mais nova.
'Esta funcionalidade � �til quando h� a necessidade de v�rias inst�ncias da mesma planilha para cuidar de diferentes assuntos, por exemplo, a mesma planilha para cuidar de obras diferentes ou clientes diferentes. E ao melhorar uma planilha, ela permite que a planilha de outros clientes seja atualizada facilmente.


' Defini��es:
' - Telas para o Usu�rio:
' -- Abas com a cor preta definidas s�o consideradas planilhas auxili�res e n�o s�o exibidas para o usu�rio final. Todas as demais abas s�o consideradas "telas" para o usu�rio e seguem um modelo.
' -- As c�lulas da primeira linha (b1, c1, d1, etc.) devem ter uma cor de fundo definida para indicar que essa coluna deve ser exibida ao usu�rio. As colunas em que a c�lula da primeira linha estejam em branco (sem cor definida), ser�o ocultadas do usu�rio. Podendo ser utilizadas para tabelas e c�lculos auxili�res.

' - A planlha 'plMenu':
' -- � esperado a presen�a de uma planilha menu, cujo (name) esteja definido como 'plModel'. Onde devem constar os links para acessar as demais telas;
' -- A macro 'gotoMenu' deve ter a tecla de atalho CTRL+M definida, para padronizar a chamada do Menu (Recomenda��o)

Option Explicit

'Atribui o foco na aba do Menu
' ATALHO: CTRL+M
Sub gotoMenu()
Attribute gotoMenu.VB_ProcData.VB_Invoke_Func = "m\n14"
    plMenu.Activate
End Sub


'Oculta todas abas com cores definidas de preto
Sub hideBlackSheets()
    Dim ws As Worksheet
    Dim blackColor As Long
    blackColor = RGB(0, 0, 0) ' Define o valor da cor preta
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Tab.ColorIndex <> xlColorIndexNone Then
            If ws.Tab.Color = blackColor Then
                ws.Visible = xlSheetHidden
            End If
        End If
    Next ws
End Sub


'Ocultas todas as colunas cujo cabe�alho n�o tenha cor, de todas as abas n�o coloridas
'!Se obtiver o erro de que n�o � poss�vel "empurar os objetos para fora da planilha", � provavelm que nas colunas escondidas tenham alguma 'nota'. Ao tentar ocultar todas as colunas a direita a nota n�o tem para onde ir e "n�o pode ficar fora da planilha". Exclua a nota.
Sub hideAuxColumns()
    Dim ws As Worksheet, cell As Range
    Dim startCell As Range, endCell As Range
    Dim hide As Boolean

    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Tab.ColorIndex = xlColorIndexNone Then
            hide = False
            For Each cell In ws.Range("1:1")
                If (cell.Interior.TintAndShade = 0 And cell.Interior.Color = 16777215) Then 'Testa a cor branca (sem cor)
                    If (hide) Then
                        Set endCell = cell
                    Else
                        Set startCell = cell
                        Set endCell = cell
                    End If
                    hide = True
                Else
                    If (hide) Then
                        ws.Range(startCell.Address, endCell.Address).EntireColumn.Hidden = True
                        hide = False
                    End If
                End If
            Next cell
            If (hide) Then
                ws.Range(startCell.Address, endCell.Address).EntireColumn.Hidden = True
                hide = False
            End If
        End If
    Next ws
End Sub


'Esta macro aplica todas os m�todos existentes nesse m�dulo respons�veis para deixar a planilha no modo de visualiza��o para o usu�rio final
'Por exemplo: Esconde abas pintadas de preto (auxiliares), copia o menu do modelo para as demais abas, esconde as colunas auxiliares, etc.
Sub applyUserMode()
    hideAuxColumns
    hideBlackSheets
End Sub
