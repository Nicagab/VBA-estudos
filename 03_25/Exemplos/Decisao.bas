Option Explicit

Sub Situacao()

Dim imc As Single
imc = InputBox("Digite o IMC", "Situação", 0)

If imc < 20 Then
    MsgBox "O IMC é " & imc & vbNewLine & "Baixo", vbCritical, "Situação"
ElseIf imc < 25 Then
    MsgBox "O IMC é " & imc & vbNewLine & "Normal", vbInformation, "Situação"
Else
    MsgBox "O IMC é " & imc & vbNewLine & "Alto", vbCritical, "Situação"
End If
End Sub
Sub Indice()

Worksheets("Massa").Select
Select Case Range("C3").Value
Case Is < 20
    Range("C3").Font.Color = vbBlue
Case Is < 25
    Range("C3").Font.Color = vbGreen
Case Else
    Range("C3").Font.Color = vbRed
End Select
End Sub
Sub Escola()
Dim resp As Integer
resp = MsgBox("Estuda na FATEC?", vbYesNo + vbQuestion, "Pergunta")
If resp = vbYes Then
    MsgBox "Boa!", vbExclamation, "Resposta"
Else
    MsgBox "Que pena!", vbCritical, "Resposta"
End If
End Sub
