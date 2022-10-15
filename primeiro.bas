Attribute VB_Name = "Módulo1"
Sub primeiro()
'O comando Dim(Dimension) é utilizado para declrar variavel
'A variavel nome foi tipada como String(texto)

Dim nome As String
'O comando InputBox abre uma caixa de entrada de dados
'Assim o usuario digita o nome e aloca na variavel nome

nome = InputBox("Digite o seu nome")
'O comando Range permite selecionar uma celula na planilha do Exel.
'Assim selecionamos a célula A1 e adicionamos o valor que foi
'digitado na caixa de entrada e usado na variável nome

Range("A1").Value = nome

End Sub
