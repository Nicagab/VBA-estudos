Option Explicit
Private turno As String
Public tanque As Single
Private Const horario As String = "12h50"
Public Const amanha As Single = 6

Private Sub Mensagem()
MsgBox "Bom curso!", vbExclamation, "Aviso"
End Sub
Sub Hello()
MsgBox "Bem-vindos ao curso de ADS", vbExclamation, "Aviso"
Mensagem
End Sub
Sub Lista()
Dim nome As String
nome = InputBox("Digite seu nome", "Atenção", "Seu nome...")
MsgBox "Olá " & nome, vbExclamation, "Aviso"
End Sub
Sub Presenca()
Dim nome
nome = InputBox("Digite seu nome", "Atenção")
MsgBox "Olá " & nome, vbExclamation, "Aviso"
End Sub
Sub Consumo()
tanque = InputBox("Quilometros por litro?", "Combustível")
MsgBox "Seu carro irá consumir " & 100 / tanque & " litros para andar 100 kilometros", _
    vbCritical, "Despesa"
End Sub
Sub Hora()
turno = InputBox("Informe o turno", "Escola")
MsgBox "O turno da " & turno & " começa as aulas às " & horario, _
    vbInformation, "Escola"
End Sub
Sub Grade()
MsgBox "Na FATEC as aulas do turno da " & turno & " começam às " & horario, _
    vbInformation, "Escola"
End Sub
Sub dolar()
Const hoje As Single = 5
MsgBox "Hoje serão necessários " & 100 * hoje & " reais para comprar 100 dolares", _
    vbCritical, "Cotação"
End Sub
Sub VidaUtil()
Static Numero As Integer
MsgBox "O número inicial é " & Numero
Numero = InputBox("Digite um novo número", "Números")
MsgBox "E mudou para " & Numero, vbQuestion, "Número"
End Sub
