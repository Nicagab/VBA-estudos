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
