Attribute VB_Name = "Módulo1"
Sub Cadastro()
'Obter dados pessoais do cliente
Dim nome, email, telefone, celula As String
Dim idade As Integer
nome = InputBox("Digite o seu nome", "Caixa de entrada")
email = InputBox("Digite o seu email", "Caixa de entrada")
telefone = InputBox("Digite seu telefone", "Caixa de entrada")
idade = InputBox("Digite sua idade", "caixa de entrada")
celula = InputBox("Digite a célula que você quer colocar o nome")
Range(celula).Value = nome
celula = InputBox("Digite a célula que você quer colocar o email")
Range(celula).Value = email
celula = InputBox("Digite a célula que você quer colocar o telefone")
Range(celula).Value = telefone
celula = InputBox("Digite a célula que você quer colocar o idade")
Range(celula).Value = idade
End Sub
Sub PegaData()
Dim data, semana As String
data = Range("B2").Value
semana = Range("C2").Value

MsgBox data
MsgBox semana


End Sub
