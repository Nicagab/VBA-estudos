Option Explicit

Private Sub UserForm_Activate()
Application.Wait Now + TimeValue("00:00:04")
Unload Me
Cadastro.Show
End Sub

Private Sub UserForm_Initialize()
Me.Height = Application.Height
Me.Width = Application.Width
Me.Left = Application.Left
Me.Top = Application.Top
End Sub
