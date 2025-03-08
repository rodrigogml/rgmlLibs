Attribute VB_Name = "libExtense"
'libExtense - v1 - 08/03/2025

Function ReaisPorExtenso(ByVal valor As Double) As String
    Dim unidades As Variant
    Dim dezenas As Variant
    Dim centenas As Variant
    Dim num As String
    Dim extenso As String
    Dim reais As Long
    Dim centavos As Long
    
    ' Arredondar para duas casas decimais
    valor = Round(valor, 2)
    
    ' Separar parte inteira e decimal
    reais = Int(valor)
    centavos = Round((valor - reais) * 100)
    
    ' Definição dos números por extenso
    unidades = Array("", "um ", "dois ", "três ", "quatro ", "cinco ", "seis ", "sete ", "oito ", "nove ", "dez ", "onze ", "doze ", "treze ", "quatorze ", "quinze ", "dezesseis ", "dezessete ", "dezoito ", "dezenove ")
    dezenas = Array("", "dez ", "vinte ", "trinta ", "quarenta ", "cinquenta ", "sessenta ", "setenta ", "oitenta ", "noventa ")
    centenas = Array("", "cento ", "duzentos ", "trezentos ", "quatrocentos ", "quinhentos ", "seiscentos ", "setecentos ", "oitocentos ", "novecentos ")
    
    ' Converter reais
    If reais > 0 Then
        extenso = NumeroPorExtensoGrande(reais, unidades, dezenas, centenas) & " "
        If reais > 1 Then
            extenso = extenso & "reais "
        Else
            extenso = extenso & "real "
        End If
    End If
    
    ' Converter centavos
    If centavos > 0 Then
        If extenso <> "" Then extenso = extenso & "e "
        extenso = extenso & NumeroPorExtenso(centavos, unidades, dezenas, centenas) & " centavo "
        If centavos > 1 Then extenso = Replace(extenso, "centavo ", "centavos ")
    End If
    
    ReaisPorExtenso = Trim(extenso)
End Function

Function NumeroPorExtensoGrande(ByVal num As Long, unidades As Variant, dezenas As Variant, centenas As Variant) As String
    Dim resultado As String
    Dim partes As Variant
    Dim nomes As Variant
    Dim i As Integer
    
    partes = Array(1000000000000#, 1000000000, 1000000, 1000, 1)
    nomes = Array("trilhão ", "bilhão ", "milhão ", "mil ", "")
    
    For i = LBound(partes) To UBound(partes)
        If num >= partes(i) Then
            Dim quantidade As Long
            quantidade = Int(num / partes(i))
            num = num Mod partes(i)
            
            If resultado <> "" Then resultado = resultado & "e "
            resultado = resultado & NumeroPorExtenso(quantidade, unidades, dezenas, centenas) & " " & nomes(i)
            
            If quantidade > 1 And nomes(i) <> "" Then
                If nomes(i) = "milhão " Then
                    resultado = Replace(resultado, "milhão ", "milhões ")
                ElseIf nomes(i) = "bilhão " Then
                    resultado = Replace(resultado, "bilhão ", "bilhões ")
                ElseIf nomes(i) = "trilhão " Then
                    resultado = Replace(resultado, "trilhão ", "trilhões ")
                End If
            End If
        End If
    Next i
    
    NumeroPorExtensoGrande = Trim(resultado)
End Function

Function NumeroPorExtenso(ByVal num As Long, unidades As Variant, dezenas As Variant, centenas As Variant) As String
    Dim resultado As String
    
    If num = 100 Then
        NumeroPorExtenso = "cem "
        Exit Function
    End If
    
    If num >= 100 Then
        resultado = centenas(Int(num / 100))
        num = num Mod 100
        If num > 0 Then resultado = resultado & "e "
    End If
    
    If num >= 20 Then
        resultado = resultado & dezenas(Int(num / 10))
        num = num Mod 10
        If num > 0 Then resultado = resultado & "e "
    End If
    
    If num > 0 Then
        resultado = resultado & unidades(num)
    End If
    
    NumeroPorExtenso = Trim(resultado)
End Function


Sub t()
    MsgBox ReaisPorExtenso(1948.25)
End Sub
