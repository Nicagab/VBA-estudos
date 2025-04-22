Option Explicit
Sub Carro()
tanque = InputBox("Digite a capacidade do tanque do carro", "Despesa")
MsgBox "Para completar o tanque irá gastar " & 5 * tanque & " reais", _
    vbInformation, "Despesa"
End Sub
Sub Cotacao()
MsgBox "Amanhã serão necessários " & 100 * amanha & " reais para comprar 100 dolares", _
    vbCritical, "Cotação"
End Sub
