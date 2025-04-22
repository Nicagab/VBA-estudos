Option Explicit

Sub VotarDirigir()
Const ano As Integer = 2025
Dim nasc As Integer
Dim idade As Integer
Dim msg As String
nasc = InputBox("Informe o ano de nascimento", "Cálculo da Idade")
idade = ano - nasc
msg = "Idade: " & idade & " anos" & vbNewLine
If idade < 16 Then
    MsgBox msg & "Não pode votar nem dirigir", vbCritical, "Análise"
ElseIf idade < 18 Then
    MsgBox msg & "Pode votar, mas não pode dirigir", vbExclamation, "Análise"
Else
    MsgBox msg & "Pode votar e dirigir", vbInformation, "Análise"
End If
End Sub

Sub Calendario()
Dim coluna As Byte
Dim linha As Byte
Do While linha < 1 Or linha > 20
linha = InputBox("Digite um número inteiro entre 1 e 20")
Loop
For coluna = 1 To linha
    With Cells(linha, coluna)
        .Value = Date + coluna
        .NumberFormat = "dd/mm/yyyy"
        .Font.Size = 10
        .Font.Name = "Courie New"
        .Font.Bold = True
        .Font.ColorIndex = coluna + 37
        .HorizontalAlignment = xlCenter
    End With
Next
End Sub
