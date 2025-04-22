Option Explicit
Public Const depto As String = "Estágios"

Sub Converter()
Dim km As Single
km = InputBox("Digite a distância em km", "Conversor")
MsgBox km & " km = " & km * 0.62 & " mi", vbInformation, "Conversor"
End Sub
