Option Explicit
Sub Lista()
Dim contador As Byte
Sheets("Notas").Select
For contador = 2 To 11
Cells(contador, 1).Value = contador - 1
Next
End Sub
Sub Media()
Sheets("Notas").Select
Range("E2").Select
       
Do While ActiveCell.Value <> ""
    If ActiveCell.Value < 6 Then
        ActiveCell.Font.Color = vbRed
    Else
        ActiveCell.Font.Color = vbBlue
    End If
    ActiveCell.Offset(1, 0).Select
Loop
End Sub
Sub Resultado()
Sheets("Notas").Select
Range("E2").Select

Do Until ActiveCell.Value = ""
    Select Case ActiveCell.Value
        Case Is >= 6
            ActiveCell.Offset(0, 1).Value = "Aprovado"
        Case Is >= 5
            ActiveCell.Offset(0, 1).Value = "Exame"
        Case Else
            ActiveCell.Offset(0, 1).Value = "Reprovado"
        End Select
        ActiveCell.Offset(1, 0).Select
Loop
End Sub
Sub Pendentes()

Dim item As Range
Sheets("Notas").Select
For Each item In Range("F2:F11")
    With item
        If .Value = "Reprovado" Or .Value = "Exame" Then
            .Interior.Color = vbYellow
            .Font.Color = vbRed
            .Font.Bold = True
        End If
    End With
Next
Range("A1").Select

End Sub
