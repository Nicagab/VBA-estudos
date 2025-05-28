Option Compare Database
Option Explicit
Private tbl As DAO.Recordset



Private Sub cmdBusca_Click()
Set tbl = CurrentDb.OpenRecordset("Livros")
If IsNull(txtID) Then
    MsgBox "Digite o código do livro", vbCritical, "Buscar"
    txtID.SetFocus
Else
    tbl.Index = "PrimaryKey"
    tbl.Seek "=", txtID
    If tbl.NoMatch = True Then
        txtTitulo.Visible = True
        txtAutor.Visible = True
        cboArea.Visible = True
        cmdInclui.Visible = True
        
        txtTitulo.SetFocus
        
        txtID.Enabled = False
        cmdBusca.Enabled = False
    Else
        MsgBox "Livro já cadastrado", vbCritical, "Buscar"
        txtID = Null
        txtID.SetFocus
    End If
        
End If
End Sub

Private Sub cmdInclui_Click()

If MsgBox("Cadastrar livro?", vbYesNo + vbQuestion, "Cadastrar") = vbYes Then
    tbl.AddNew
    tbl!ID = txtID
    tbl!Titulo = txtTitulo
    tbl!Autor = txtAutor
    tbl!Area = cboArea
    tbl.Update
    
    MsgBox "Livro cadastrado", vbCritical, "Cadastrar"
Else
    MsgBox "Cadastramento cancelado", vbCritical, "Cadastrar"
End If
txtID = Null
txtTitulo = Null
txtAutor = Null
cboArea = "Economia"

txtID.Enabled = True
cmdBusca.Enabled = True

txtID.SetFocus

txtTitulo.Visible = False
txtAutor.Visible = False
cboArea.Visible = False
cmdInclui.Visible = False
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
If DataErr Then
MsgBox "Dados não preenchidos corretamente", vbExclamation, "Atenção"
End If
Response = acDataErrContinue
End Sub
