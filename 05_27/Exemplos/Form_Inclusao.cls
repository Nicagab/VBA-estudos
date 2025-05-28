Option Compare Database
Option Explicit

Private tb As DAO.Recordset


Private Sub cmdProcura_Click()
Set tb = CurrentDb.OpenRecordset("Eventos")
If IsNull(txtCodigo) Then
    MsgBox "Informa o código", vbCritical, "Atenção"
    txtCodigo.SetFocus
Else
    tb.Index = "PrimaryKey"
    tb.Seek "=", txtCodigo
    If tb.NoMatch = True Then
        txtTitulo.Enabled = True
        txtPalestra.Enabled = True
        cboTipo.Enabled = True
        cmdSalva.Enabled = True
        txtTitulo.SetFocus
        txtCodigo.Enabled = False
        cmdProcura.Enabled = False
    Else
        MsgBox "Código já cadastrdo", vbCritical, "Inclusão"
        txtCodigo = Null
        txtCodigo.SetFocus
    End If
End If
End Sub


Private Sub cmdSalva_Click()
    If MsgBox("Confirma inclusão?", vbQuestion + vbYesNo, "Inclusão") = vbNo Then
        MsgBox "Inclusão cancelada", vbCritical, "Inclusão"
    Else
        tb.AddNew
        tb!Codigo = txtCodigo
        tb!Titulo = txtTitulo
        tb!Palestrante = txtPalestra
        tb!Tipo = cboTipo
        tb.Update
        MsgBox "Evento incluído", vbCritical, "Inclusão"
    End If
    
    txtCodigo = Null
    txtPalestra = Null
    txtTitulo = Null
    cboTipo = "Palestra"
    
    txtCodigo.Enabled = True
    cmdProcura.Enabled = True
    txtCodigo.SetFocus
    
    txtTitulo.Enabled = False
    txtPalestra.Enabled = False
    cboTipo.Enabled = False
    cmdSalva.Enabled = False
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
Select Case DataErr
    Case 2279
        MsgBox "Digite uma letra e três números", vbCritical, "Erro"
    Case 2237
        MsgBox "Selecione na lista", vbCritical, "Erro"
End Select
Response = acDataErrContinue
End Sub
