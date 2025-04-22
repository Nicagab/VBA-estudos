Option Explicit

Private Sub cmdSalva_Click()
If txtAluno = "" Then
    MsgBox "Preenchimento obrigatório", vbCritical, "Atenção"
    txtAluno.SetFocus
    Exit Sub
Else
    Sheets("Alunos").Select
    Range("A1").Select
    Do Until ActiveCell = ""
    ActiveCell.Offset(1, 0).Select
    Loop
    ActiveCell = txtAluno
    ActiveCell.Offset(0, 1) = cboDisc
    ActiveCell.Offset(0, 2) = CSng(txtMedia)
    ActiveCell.Offset(0, 3) = lblFim
End If
txtAluno = ""
cboDisc = "Informática"
txtMedia = "0,0"
lblFim = "Reprovado"
lblFim.ForeColor = vbRed
txtAluno.SetFocus
End Sub

Private Sub txtMedia_Exit(ByVal Cancel As MSForms.ReturnBoolean)
txtMedia = FormatNumber(Round(txtMedia, 1), 1)
If txtMedia < 0 Or txtMedia > 10 Then
    MsgBox "Média inválida", vbCritical, "Erro"
    lblFim = "Reprovado"
    lblFim.ForeColor = vbRed
    txtMedia = "0,0"
    Cancel = True
Else
    If txtMedia >= 6 Then
        lblFim = "Aprovado"
        lblFim.ForeColor = vbBlue
    Else
        lblFim = "Reprovado"
        lblFim.ForeColor = vbRed
    End If
End If
End Sub
