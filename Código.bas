Attribute VB_Name = "Módulo1"
Sub ControleEstoque()
    Dim nomeEquipamento As String
    Dim quantidade As Integer
    Dim preco As Currency
    Dim dataEntrada As String
    Dim dataValidade As String
    Dim linha As Integer
    
    ' encontrar ultima linha preenchida
    linha = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    Do

        nomeEquipamento = InputBox("Digite o nome do equipamento (ou digite 0 para sair):")
        
        ' para sair do loop
        If nomeEquipamento = "0" Then
            Exit Do
        End If

        quantidade = InputBox("Digite a quantidade de produtos:")

        preco = InputBox("Digite o preço do produto (em R$):")

        dataEntrada = InputBox("Digite a data de entrada (MM/AAAA):")

        Cells(linha, 1).Value = nomeEquipamento
        Cells(linha, 2).Value = quantidade
        Cells(linha, 3).Value = preco
        Cells(linha, 4).Value = dataEntrada
        
        ' avancar pra proxima linha
        linha = linha + 1
        
    Loop
    
    MsgBox "Parabéns! Todos os produtos foram inseridos com sucesso! :)               Agora eles estão no nosso banco de dados.", vbInformation

End Sub

