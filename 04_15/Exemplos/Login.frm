Option Explicit

Private Sub cmdEntra_Click()
If txtSenha = "123" Then
  Unload Me
  Vendas.Show
Else
  MsgBox "Senha inválida", vbCritical, "Login"
  txtSenha = ""
  txtSenha.SetFocus
End If
End Sub

Private Sub cmdSai_Click()
Application.Quit
End Sub

Private Sub txtSenha_Enter()
txtSenha.BackColor = RGB(102, 255, 204)
End Sub

Private Sub txtSenha_Exit(ByVal Cancel As MSForms.ReturnBoolean)
txtSenha.BackColor = RGB(255, 255, 255)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
  Cancel = True
End If
End Sub
