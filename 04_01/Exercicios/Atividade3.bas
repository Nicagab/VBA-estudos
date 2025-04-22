Option Explicit
Option Compare Text
Sub Multa()
Dim placa As Byte
placa = InputBox("Informe o final da placa", "Rodízio")
Do While placa > 9
    placa = InputBox("Final inválido, digite novamente", "Rodízio")
Loop
If Weekday(Date) = Rodizio(placa) Then
    MsgBox "Placa final " & placa & vbNewLine & "Atenção com a zona de rodízio", vbCritical
Else
    MsgBox "Placa final " & placa & vbNewLine & "Circulação liberada", vbInformation
End If
End Sub
Function Rodizio(final As Byte) As Byte
If final = 1 Or final = 2 Then
    Rodizio = 2
ElseIf final = 3 Or final = 4 Then
    Rodizio = 3
ElseIf final = 5 Or final = 6 Then
    Rodizio = 4
ElseIf final = 7 Or final = 8 Then
    Rodizio = 5
Else
    Rodizio = 6
End If
End Function
Function Gasto(D As Single, Optional C As Variant) As Single
If IsMissing(C) Then
    Gasto = D / 10 * 5
ElseIf C = "G" Then
    Gasto = D / 20 * 6.5
ElseIf C = "A" Then
    Gasto = D / 15 * 4
Else
    Gasto = 0
End If
End Function
