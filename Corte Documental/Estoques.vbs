Sub CarregarEstoqueDepositos()
    ConectarBanco

    sql = "SELECT Código, Centro, Deposito, [Nome Deposito], Tipo, Processo, Estoque, Status FROM Dados WHERE Status = 'Cadastrado'"
    rs.Open sql, conn

    Dim relatorioHTML
    relatorioHTML = "<table border='1'><tr>" & _
                    "<th>CENTRO</th>" & _
                    "<th>DEPOSITO</th>" & _
                    "<th>DESCRIÇÃO</th>" & _
                    "<th>TIPO</th>" & _
                    "<th>PROCESSO</th>" & _
                    "<th>ESTOQUE</th>" & _
                    "<th>STATUS</th>" & _
                    "<th>AÇÃO</th>" & _
                    "</tr>"

    Do While Not rs.EOF
        Dim acaoHTML
        Dim rowId, linhaClasse
        rowId = rs.Fields("Código").Value

        ' Condição para cores da linha
        If UCase(rs.Fields("Tipo").Value) <> "FISICO" And rs.Fields("Estoque").Value <> 0 Then
            linhaClasse = " style='background-color:#FF9999;'"  ' vermelho
        Else
            linhaClasse = " style='background-color:#CCFFCC;'"  ' verde
        End If

        ' Botões de ação
        acaoHTML = "<button class='btn-editar' onclick='Incluir(" & rowId & ")'>Incluir no Corte</button>" & _
                   "<button class='btn-editar' onclick='Excluir(" & rowId & ")'>Remover do Corte</button>" & _
                   "<button class='btn-editar' onclick='EditarDeposito(" & rowId & ")'>Editar</button>"

        relatorioHTML = relatorioHTML & "<tr" & linhaClasse & ">" & _
                         "<td>" & rs.Fields("Centro").Value & "</td>" & _
                         "<td>" & rs.Fields("Deposito").Value & "</td>" & _
                         "<td>" & rs.Fields("Nome Deposito").Value & "</td>" & _
                         "<td>" & rs.Fields("Tipo").Value & "</td>" & _
                         "<td>" & rs.Fields("Processo").Value & "</td>" & _
                         "<td>" & rs.Fields("Estoque").Value & "</td>" & _
                         "<td>" & rs.Fields("Status").Value & "</td>" & _
                         "<td>" & acaoHTML & "</td>" & _
                         "</tr>"

        rs.MoveNext
    Loop

    relatorioHTML = relatorioHTML & "</table>"

    document.getElementById("CarregarEstoqueDepositos").style.display = "Block"
    document.getElementById("CarregarEstoqueDepositos").innerHTML = relatorioHTML

    rs.Close
    conn.Close
End Sub

Sub AtualizarEstoques()
    ' Conectar ao banco
    ConectarBanco

    ' Seleciona todos os registros que estão cadastrados
    sql = "SELECT Código, Centro, Deposito, Estoque FROM Dados WHERE Status = 'Cadastrado'"
    rs.Open sql, conn

    ' Loop pelos registros
    Do While Not rs.EOF
        Dim codigo, centro, deposito, estoque
        codigo = rs.Fields("Código").Value
        centro = rs.Fields("Centro").Value
        deposito = rs.Fields("Deposito").Value
        estoque = rs.Fields("Estoque").Value

        ' Mostrar MsgBox para identificar
        MsgBox "Centro: " & centro & vbCrLf & "Depósito: " & deposito

        ' Solicitar novo valor de estoque
        Dim novoEstoque
        novoEstoque = InputBox("Informe o novo estoque para o depósito " & deposito & ":", "Atualizar Estoque", estoque)

        ' Verificar se foi preenchido
        If novoEstoque <> "" Then
            ' Atualizar banco
            conn.Execute "UPDATE Dados SET Estoque = " & novoEstoque & " WHERE Código = " & codigo & ";"
        End If

        rs.MoveNext
    Loop

    ' Fechar conexão
    rs.Close
    conn.Close

    MsgBox "Atualização de estoques concluída!"
	
	CarregarEstoqueDepositos
	
End Sub

' Função para "Editar" que limpa Tipo e Processo e atualiza Status
Sub Incluir(idLinha)
    ConectarBanco

    sql = "UPDATE Dados SET Status = 'PROGRAMADO' WHERE Código = " & idLinha & ";"
    conn.Execute sql

    MsgBox "Registro alterado para cadastro!"

    ' Atualizar tabela
    CarregarDepositosCadastrados
	CarregarEstoqueDepositos
End Sub

' Função para "Editar" que limpa Tipo e Processo e atualiza Status
Sub Excluir(idLinha)
    ConectarBanco

    sql = "UPDATE Dados SET Status = 'REMOVER' WHERE Código = " & idLinha & ";"
    conn.Execute sql

    MsgBox "Registro alterado para cadastro!"

    ' Atualizar tabela
    CarregarDepositosCadastrados
	CarregarEstoqueDepositos
End Sub