Attribute VB_Name = "ModuloIQA"

Function Log10(Number As Double) As Double
    ' Verifica se o número é maior que zero
    If Number > 0 Then
        Log10 = Log(Number) / Log(10)
    Else
        Log10 = CVErr(xlErrNum) ' Retorna erro se o número for menor ou igual a zero
    End If
End Function


Function IQA(rng As Range) As Variant
    Dim Resultados(8) As Double
    Dim pesos_w(8) As Double
    Dim perc_Saturacao As Double
    Dim ConcentracaoSaturacao As Double

    ' Extraindo os valores da faixa (rng)
    Dim Oxigenio As Double: Oxigenio = rng.Cells(1, 1).Value
    Dim Coliformes As Double: Coliformes = rng.Cells(1, 2).Value
    Dim pH As Double: pH = rng.Cells(1, 3).Value
    Dim DBO As Double: DBO = rng.Cells(1, 4).Value
    Dim Nitrato As Double: Nitrato = rng.Cells(1, 5).Value
    Dim Fosfato As Double: Fosfato = rng.Cells(1, 6).Value
    Dim Temperatura As Double: Temperatura = rng.Cells(1, 7).Value
    Dim Turbidez As Double: Turbidez = rng.Cells(1, 8).Value
    Dim SolidosTotais As Double: SolidosTotais = rng.Cells(1, 9).Value
    Dim Altitude As Double: Altitude = rng.Cells(1, 10).Value
    Dim TipoFosfato As String: TipoFosfato = rng.Cells(1, 11).Value

    ' Ajuste do valor do Fosfato
    If TipoFosfato = "fosforo" Then
        Fosfato = Fosfato * 3.066
    End If

    ' Garantir que a Altitude não seja zero
    If Altitude = 0 Then
        Altitude = 1
    End If

    ' Oxigênio Dissolvido
    ConcentracaoSaturacao = (14.62 - 0.3898 * Temperatura + 0.006969 * Temperatura ^ 2 - 0.00005896 * Temperatura ^ 3) * (1 - 0.0000228675 * Altitude) ^ 5.167
    perc_Saturacao = 100 * Oxigenio / ConcentracaoSaturacao

' Cálculo para Oxigênio Dissolvido
If perc_Saturacao > 0 And perc_Saturacao <= 50 Then
    Resultados(0) = 3 + 0.34 * perc_Saturacao + 0.008095 * perc_Saturacao ^ 2 + 1.35252 * 0.00001 * perc_Saturacao ^ 3
ElseIf perc_Saturacao > 50 And perc_Saturacao <= 85 Then
    Resultados(0) = 3 - 1.166 * perc_Saturacao + 0.058 * perc_Saturacao - 3.803435 * 0.00001 * perc_Saturacao ^ 3
ElseIf perc_Saturacao > 85 And perc_Saturacao <= 100 Then
    Resultados(0) = 3 + 3.7745 * perc_Saturacao ^ 0.704889
ElseIf perc_Saturacao > 100 And perc_Saturacao <= 140 Then
    Resultados(0) = 3 + 2.9 * perc_Saturacao - 0.02496 * perc_Saturacao ^ 2 + 5.60919 * 0.00001 * perc_Saturacao ^ 3
Else
    Resultados(0) = 50
End If
pesos_w(0) = Resultados(0) ^ 0.17 ' Mantido como estava


' Coliformes Fecais
If Coliformes > 0 Then
    Coliformes = Log10(Coliformes)
    If Coliformes <= 1 Then
        Resultados(1) = 100 - 33 * Coliformes
    ElseIf Coliformes <= 5 Then
        Resultados(1) = 100 - 37.2 * Coliformes + 3.60743 * Coliformes ^ 2
    Else
        Resultados(1) = 3
    End If
Else
    Resultados(1) = 3 ' Valor padrão
End If
pesos_w(1) = Resultados(1) ^ 0.15 ' Mantido como estava


' pH
Select Case pH
    Case Is <= 2
        Resultados(2) = 2
    Case Is <= 4
        Resultados(2) = 13.6 - 10.6 * pH + 2.4364 * pH ^ 2
    Case Is <= 6.2
        Resultados(2) = 155.5 - 77.36 * pH - 10.2481 * pH ^ 2
    Case Is <= 7
        Resultados(2) = -657.2 + 197.38 * pH - 12.9167 * pH ^ 2
    Case Is <= 8
        Resultados(2) = -427.8 + 142.05 * pH - 9.695 * pH ^ 2
    Case Is <= 8.5
        Resultados(2) = 216 - 16 * pH
    Case Is <= 9
        Resultados(2) = 1415823 * Exp(-1.1507 * pH)
    Case Is <= 10
        Resultados(2) = 228 - 27 * pH
    Case Is <= 12
        Resultados(2) = 633 - 106.5 * pH + 4.5 * pH ^ 2
    Case Else
        Resultados(2) = 3
End Select
pesos_w(2) = Resultados(2) ^ 0.12 ' Mantido como estava


' DBO
If DBO > 0 Then
    If DBO <= 5 Then
        Resultados(3) = 99.96 * Exp(-0.1232728 * DBO)
    ElseIf DBO <= 15 Then
        Resultados(3) = 104.67 - 31.5463 * Log10(DBO)
    ElseIf DBO <= 30 Then
        Resultados(3) = 4394.91 * DBO ^ -1.99809
    Else
        Resultados(3) = 2 ' Para valores acima de 30
    End If
Else
    Resultados(3) = 2 ' Valor padrão
End If
pesos_w(3) = Resultados(3) ^ 0.1 ' Mantido como estava


' Nitrogênio Total
If Nitrato > 0 Then
    If Nitrato <= 10 Then
        Resultados(4) = 100 - 8.169 * Nitrato + 0.3059 * Nitrato ^ 2
    ElseIf Nitrato <= 60 Then
        Resultados(4) = 101.9 - 23.1023 * Log10(Nitrato)
    ElseIf Nitrato <= 100 Then
        Resultados(4) = 159.3148 * Exp(-0.0512842 * Nitrato)
    Else
        Resultados(4) = 1 ' Para valores acima de 100
    End If
Else
    Resultados(4) = 1 ' Valor padrão
End If
pesos_w(4) = Resultados(4) ^ 0.1 ' Mantido como estava


' Fosfatos
If Fosfato > 0 Then
    If Fosfato <= 1 Then
        Resultados(5) = 99 * Exp(-0.91629 * Fosfato) ' Alterado para corresponder à lógica do JS
    ElseIf Fosfato <= 5 Then
        Resultados(5) = 57.6 - 20.178 * Fosfato + 2.1326 * Fosfato ^ 2 ' Ajustado para a faixa de 1 a 5
    ElseIf Fosfato <= 10 Then
        Resultados(5) = 19.8 * Exp(-0.13544 * Fosfato) ' Adicionado o intervalo para 5 a 10
    Else
        Resultados(5) = 5#  ' Para valores acima de 10
    End If
Else
    Resultados(5) = 4 ' Valor padrão
End If
pesos_w(5) = Resultados(5) ^ 0.1 ' Mantido como estava



    pesos_w(6) = 94 ^ 0.1

' Turbidez
If Turbidez > 0 Then
    If Turbidez <= 25 Then ' Alterado para 25
        Resultados(7) = 100.17 - 2.67 * Turbidez + 0.03775 * Turbidez ^ 2 ' Ajuste feito para coincidir com a lógica do JS
    ElseIf Turbidez <= 100 Then ' Alterado para 100
        Resultados(7) = 84.76 * Exp(-0.016206 * Turbidez) ' Função exponencial em VB
    Else
        Resultados(7) = 5 ' Para valores acima de 100
    End If
Else
    Resultados(7) = 2 ' Valor padrão
End If
pesos_w(7) = Resultados(7) ^ 0.08 ' Mantido como estava


' Sólidos Totais
If SolidosTotais > 0 Then
    If SolidosTotais <= 150 Then
        Resultados(8) = 79.75 + 0.166 * SolidosTotais - 0.001088 * SolidosTotais ^ 2
    ElseIf SolidosTotais <= 500 Then
        Resultados(8) = 101.67 - 0.13917 * (SolidosTotais - 150)
    Else
        Resultados(8) = 32
    End If
Else
    Resultados(8) = 2
End If
pesos_w(8) = Resultados(8) ^ 0.08

    IQA = pesos_w(0) * pesos_w(1) * pesos_w(2) * pesos_w(3) * pesos_w(4) * pesos_w(5) * pesos_w(6) * pesos_w(7) * pesos_w(8)
End Function

Function ClassificaIQA(IQA As Double) As String
    Select Case IQA
        Case Is > 79
            ClassificaIQA = "ÓTIMA"
        Case 51 To 79
            ClassificaIQA = "BOA"
        Case 36 To 51
            ClassificaIQA = "REGULAR"
        Case 19 To 36
            ClassificaIQA = "RUIM"
        Case Is <= 19
            ClassificaIQA = "PÉSSIMA"
        Case Else
            ClassificaIQA = "INDEFINIDA"
    End Select
End Function

 


