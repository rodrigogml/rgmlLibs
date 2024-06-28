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

