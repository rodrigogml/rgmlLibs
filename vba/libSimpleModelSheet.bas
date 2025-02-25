Attribute VB_Name = "libSimpleModelSheet"
'libAppWorksheet - v1 - 28/11/2024

'Esta biblioteca concentra funções e métodos que permitem a construção de uma planilha com controle de versão e importação de dados de uma planilha de versão enterior para uma mais nova.
'Esta funcionalidade é útil quando há a necessidade de várias instâncias da mesma planilha para cuidar de diferentes assuntos, por exemplo, a mesma planilha para cuidar de obras diferentes ou clientes diferentes. E ao melhorar uma planilha, ela permite que a planilha de outros clientes seja atualizada facilmente.


' Definições:
' - Telas para o Usuário:
' -- Abas com a cor preta definidas são consideradas planilhas auxiliáres e não são exibidas para o usuário final. Todas as demais abas são consideradas "telas" para o usuário e seguem um modelo.
' -- As células da primeira linha (b1, c1, d1, etc.) devem ter uma cor de fundo definida para indicar que essa coluna deve ser exibida ao usuário. As colunas em que a célula da primeira linha estejam em branco (sem cor definida), serão ocultadas do usuário. Podendo ser utilizadas para tabelas e cálculos auxiliáres.

' - A planlha 'plMenu':
' -- É esperado a presença de uma planilha menu, cujo (name) esteja definido como 'plModel'. Onde devem constar os links para acessar as demais telas;
' -- A macro 'gotoMenu' deve ter a tecla de atalho CTRL+M definida, para padronizar a chamada do Menu (Recomendação)

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


'Ocultas todas as colunas cujo cabeçalho não tenha cor, de todas as abas não coloridas
'!Se obtiver o erro de que não é possível "empurar os objetos para fora da planilha", é provavelm que nas colunas escondidas tenham alguma 'nota'. Ao tentar ocultar todas as colunas a direita a nota não tem para onde ir e "não pode ficar fora da planilha". Exclua a nota.
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


'Esta macro aplica todas os métodos existentes nesse módulo responsáveis para deixar a planilha no modo de visualização para o usuário final
'Por exemplo: Esconde abas pintadas de preto (auxiliares), copia o menu do modelo para as demais abas, esconde as colunas auxiliares, etc.
Sub applyUserMode()
    hideAuxColumns
    hideBlackSheets
End Sub
