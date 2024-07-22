Attribute VB_Name = "libModelSheet"
'libAppWorksheet - v1.1 - 21/07/2024

'Esta biblioteca concentra fun��es e m�todos que permitem a constru��o de uma planilha com controle de vers�o e importa��o de dados de uma planilha de vers�o enterior para uma mais nova.
'Esta funcionalidade � �til quando h� a necessidade de v�rias inst�ncias da mesma planilha para cuidar de diferentes assuntos, por exemplo, a mesma planilha para cuidar de obras diferentes ou clientes diferentes. E ao melhorar uma planilha, ela permite que a planilha de outros clientes seja atualizada facilmente.


' Defini��es:
' - Telas para o Usu�rio:
' -- Abas com a cor preta definidas s�o consideradas planilhas auxili�res e n�o s�o exibidas para o usu�rio final. Todas as demais abas s�o consideradas "telas" para o usu�rio e seguem um modelo.
' -- Na telas do usu�rio a coluna "A" � sempre reservada para o menu de navega��o. E as duas primeiras linhas para cabe�alho da tela.
' -- A partir da coluna "B" � considerado as colunas de montagem da "tela" para o usu�rio. Por padr�o, as c�lulas da primeira linha (b1, c1, d1, etc.) devem ter uma cor de fundo definida para indicar que essa coluna deve ser exibida ao usu�rio. As colunas em que a c�lula da primeira linha estejam pintadas de branco (sem cor definida), ser�o ocultadas do usu�rio. Podendo ser utilizadas para tabelas e c�lculos auxili�res.

' - A planlha 'plModel':
' -- � esperado a presen�a de uma planilha modelo (aba preta), cujo (name) esteja definido como 'plModel'.
' -- A planilha modelo serve como base para criar novas telas para o usu�rio, e ter� o seu menu (coluna A), replicado nas demais telas automaticamente (por chamada de macro)
' -- As imagens dessa coluna 'A' s�o copiadas junto para as demais telas, e para evitar a constante duplica��o de imagens, todas as imagens devem ter o texto alternativo definido como "Logo", assim s�o identificadas e exclu�das das abas destino.


Option Explicit


'Replica o menu que est� na planilha plModel para todas as outras abas que n�o tenham a cor preta em sua aba.
'Procura uma imagem com o texto alternativo "Logo" para excluir sempre que for copiar a nova coluna (e copiar novamente o logo)
Sub copyMenuModelToOthers()
    Dim s As Worksheet, img As Shape

    For Each s In ThisWorkbook.Sheets
        If (s.Name <> plModel.Name) Then
            If (s.Tab.ColorIndex = xlColorIndexNone) Then
                'Busca o logotipo e o exclu�. Identifica que � o logo pelo texto alternativo da imagem
                For Each img In s.Shapes
                    If (img.AlternativeText = "Logo") Then
                        img.Delete
                        Exit For
                    End If
                Next img
                
                'tenta excluir o logo e os �cones se existirem
                On Error Resume Next
                s.Shapes("GraphicExternalRefresh").Delete
                s.Shapes("GraphicInternalRefresh").Delete
                s.Shapes("GraphicMenu").Delete
                s.Shapes("TerminalGraph").Delete
                s.Shapes("Logo").Delete
                On Error GoTo 0
                
                'Copia a coluna A do Modelo para a panilha
                plModel.Range("A:A").Copy s.Range("A:A")
            End If
        End If
    Next s
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
    copyMenuModelToOthers
    hideAuxColumns
    hideBlackSheets
End Sub
