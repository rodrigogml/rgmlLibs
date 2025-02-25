Attribute VB_Name = "libMixAuxiliar"
'##################################################
' Métodos auxiliares do Rodrigo para facilitar a utilização do Exce
'##################################################

Option Explicit


'Formata a seleção como o formado de data padrão "dd-MM-yyyy".
' ATALHO: CTRL+SHIFT+D
Public Sub formatColumnsAsDate()
Attribute formatColumnsAsDate.VB_ProcData.VB_Invoke_Func = "D\n14"
    Selection.NumberFormat = "dd-MM-yyyy"
End Sub

'Função para ser utilizadanas planilhas que verifica se o conteúdo de uma célula é uma data válida.
Function IsDateValid(cell As Range) As Boolean
    On Error Resume Next ' Ignora erros
    IsDateValid = IsDate(cell.Value) ' Verifica se é uma data
    On Error GoTo 0 ' Retorna à configuração padrão de erro
End Function
