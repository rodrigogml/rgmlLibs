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

