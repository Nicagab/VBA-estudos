Option Explicit

Private Sub cboDestino_Change()
lblPreco = FormatNumber(cboDestino, 2)
Total
End Sub

Private Sub chkPref_Click()
If chkPref Then
  lblTot.ForeColor = vbRed
Else
  lblTot.ForeColor = vbBlack
End If
Total
End Sub

Private Sub cmdCadastra_Click()
If txtCliente = "" Then
  MsgBox "Digite o nome do cliente", vbCritical, "Aviso"
  txtCliente.SetFocus
ElseIf cboDestino = 0 Then
  MsgBox "Selecione um destino", vbCritical, "Aviso"
  cboDestino.SetFocus
Else
  Sheets("Lista").Select
  Range("A1").Select
  Do Until ActiveCell = ""
    ActiveCell.Offset(1, 0).Select
  Loop
  ActiveCell = txtCliente
  ActiveCell.Offset(0, 1) = cboDestino.Text
  ActiveCell.Offset(0, 2) = CDbl(lblTot)
  ActiveCell.Offset(0, 2).NumberFormat = "0.00"
  If optVista Then
    ActiveCell.Offset(0, 3) = "Vista"
  Else
    ActiveCell.Offset(0, 3) = "Prazo"
  End If
Columns("A:D").AutoFit
Range("A1").Select
End If
End Sub

Private Sub cmdReinicia_Click()
txtCliente = ""
cboDestino = 0
optVista = True
chkPref = False
lblQt = "1"
spnQt = 1
lblTot = "0,00"
lblPreco = "0,00"
End Sub

Private Sub optPrazo_Click()
Total
End Sub

Private Sub optVista_Click()
Total
End Sub

Private Sub spnQt_Change()
lblQt = spnQt
Total
End Sub

Sub Total()
Dim desconto As Single
If optVista Then
  desconto = 0.9
Else
  desconto = 1
End If
lblTot = FormatNumber(spnQt * cboDestino * desconto, 2)
End Sub
