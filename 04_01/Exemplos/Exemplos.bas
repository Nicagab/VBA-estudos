Option Explicit
Sub Inicio()
Dim nome As String
nome = InputBox("Digite seu nome", "Iniciais")
MsgBox "Seu nome começa com " & Left(nome, 2), vbCritical, "Iniciais"
End Sub
Sub Resultado()
Dim P1 As Single, P2 As Single, F As Single, R As String
Dim EX

P1 = InputBox("Digite a primeira nota", "Final")
P2 = InputBox("Digite a segunda nota", "Final")
EX = InputBox("Digite a nota dos exercicios", "Final")
If EX = "" Then
    F = Final(P1, P2)
Else
    F = Final(P1, P2, EX)
End If
If F >= 6 Then
    R = "Aprovado"
Else
    R = "Reprovado"
End If
MsgBox "Média Final: " & F & Chr(13) & "Resultado: " & R, vbInformation, "Final"
End Sub
Sub Dias()
MsgBox "Hoje é " & Date & vbNewLine & "E faltam " & contagem & " dias para o final deste ano", _
    vbExclamation, "Calendário"
End Sub

Function Final(N1 As Single, N2 As Single, Optional E As Variant) As Single
    If IsMissing(E) Then
        Final = (N1 + N2) / 2
    Else
        Final = (N1 + N2 + E) / 3
    End If
End Function

Function Area(larg As Single, Optional comp As Single = 0)
    If comp = 0 Then
        Area = larg ^ 2
    Else
        Area = larg * comp
    End If
End Function

Function contagem() As Integer
    contagem = #12/31/2025# - Date
End Function
