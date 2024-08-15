Attribute VB_Name = "libMixAuxiliar"
'##################################################
' M�todos auxiliares do Rodrigo para facilitar a utiliza��o do Exce
'##################################################

Option Explicit


'Formata a sele��o como o formado de data padr�o "dd-MM-yyyy".
' ATALHO: CTRL+SHIFT+D
Public Sub formatColumnsAsDate()
Attribute formatColumnsAsDate.VB_ProcData.VB_Invoke_Func = "D\n14"
    Selection.NumberFormat = "dd-MM-yyyy"
End Sub

'Fun��o para ser utilizadanas planilhas que verifica se o conte�do de uma c�lula � uma data v�lida.
Function IsDateValid(cell As Range) As Boolean
    On Error Resume Next ' Ignora erros
    IsDateValid = IsDate(cell.Value) ' Verifica se � uma data
    On Error GoTo 0 ' Retorna � configura��o padr�o de erro
End Function
