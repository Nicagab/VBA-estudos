Option Explicit

Function Final(P1 As Single, P2 As Single) As Single
Final = (P1 * 2 + P2 * 3) / 5
End Function

Sub Exemplo1()
Const valor As Single = 1000
Dim parcela As Variant
parcela = InputBox("Número de parcelas", "Loja")
MsgBox "Valor Parcela: " & valor / parcela, vbInformation, "Loja"
End Sub

Sub Exemplo2()
On Error Resume Next
Const valor As Single = 1000
Dim parcela As Variant
parcela = InputBox("Número de parcelas", "Loja")
MsgBox "Valor Parcela: " & valor / parcela, vbInformation, "Loja"
End Sub

Sub Exemplo3()
On Error GoTo Erro
Const valor As Single = 1000
Dim parcela As Variant
parcela = InputBox("Número de parcelas", "Loja")
MsgBox "Valor Parcela: " & valor / parcela, vbInformation, "Loja"
Exit Sub

Erro:
MsgBox "Número inválido", vbCritical, "Loja"
End Sub

Sub Exemplo4()
On Error GoTo Erro
Const valor As Single = 1000
Dim parcela As Variant
Repetir:
parcela = InputBox("Número de parcelas", "Loja")
MsgBox "Valor Parcela: " & valor / parcela, vbInformation, "Loja"
Exit Sub

Erro:
MsgBox "Número inválido", vbCritical, "Loja"
Resume Repetir
End Sub

Sub Exemplo5()
On Error GoTo Erro
Const valor As Single = 1000
Dim parcela As Variant
Repetir:
parcela = InputBox("Número de parcelas", "Loja")
MsgBox "Valor Parcela: " & valor / parcela, vbInformation, "Loja"
Exit Sub

Erro:
Select Case Err
  Case 11
    MsgBox "Não digite zero", vbCritical, "Loja"
  Case 13
    MsgBox "Digite números", vbCritical, "Loja"
End Select
Resume Repetir
End Sub
