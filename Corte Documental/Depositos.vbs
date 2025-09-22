Sub CarregarDepositosCadastrados()
    ConectarBanco

    sql = "SELECT Código, Centro, Deposito, [Nome Deposito], Tipo, Processo, Status FROM Dados"
    rs.Open sql, conn

    Dim relatorioHTML
    relatorioHTML = "<table border='1'><tr>" & _
                    "<th>CENTRO</th>" & _
                    "<th>DEPOSITO</th>" & _
                    "<th>DESCRIÇÃO</th>" & _
                    "<th>TIPO</th>" & _
                    "<th>PROCESSO</th>" & _
                    "<th>STATUS</th>" & _
                    "<th>AÇÃO</th>" & _
                    "</tr>"

    Do While Not rs.EOF
        Dim acaoHTML, tipoDepositoHTML, processoDepositoHTML
        Dim linhaClasse, rowId, statusAtual
        rowId = rs.Fields("Código").Value
        statusAtual = rs.Fields("Status").Value

        ' Preencher Status vazio no banco com "Cadastrar"
        If statusAtual = "" Then
            conn.Execute "UPDATE Dados SET Status = 'Cadastrar' WHERE Código = " & rowId & ";"
            statusAtual = "Cadastrar"
        End If

        ' Definir cor da linha: vermelho se Status = "Cadastrar", verde caso contrário
        If UCase(statusAtual) = "CADASTRAR" Then
            linhaClasse = " style='background-color:#FF9999;'"  ' vermelho
        Else
            linhaClasse = " style='background-color:#CCFFCC;'"  ' verde
        End If

        ' Tipo de depósito
        If rs.Fields("Tipo").Value <> "" Then
            tipoDepositoHTML = "<label>" & rs.Fields("Tipo").Value & "</label>"
        Else
            tipoDepositoHTML = "<select id='TipoDeDeposito_" & rowId & "'>" & _
                               "<option value='Selecione'>Selecione</option>" & _
                               "<option value='FISICO'>FISICO</option>" & _
                               "<option value='VIRTUAL'>VIRTUAL</option>" & _
                               "<option value='INATIVO'>INATIVO</option>" & _
                               "</select>"
        End If

        ' Processo de armazenagem
        If rs.Fields("Processo").Value <> "" Then
            processoDepositoHTML = "<label>" & rs.Fields("Processo").Value & "</label>"
        Else
            processoDepositoHTML = "<select id='ProcessoDeArmazenagem_" & rowId & "'>" & _
                                   "<option value='Selecione'>Selecione</option>" & _
                                   "<option value='MM – Material Management'>MM – Material Management</option>" & _
                                   "<option value='WM – Warehouse Management'>WM – Warehouse Management</option>" & _
                                   "<option value='TR - Estoque Terceito'>TR - Estoque Terceito</option>" & _
                                   "<option value='CR - Estoque Graos e Cereais'>CR - Estoque Graos e Cereais</option>" & _
                                   "</select>"
        End If

        ' Botões de ação
        acaoHTML = "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>Salvar Alterações</button>" & _
                   "<button class='btn-editar' onclick='EditarDeposito(" & rowId & ")'>Editar</button>"

        ' Montar a linha com a cor definida
        relatorioHTML = relatorioHTML & "<tr" & linhaClasse & ">" & _
                         "<td>" & rs.Fields("Centro").Value & "</td>" & _
                         "<td>" & rs.Fields("Deposito").Value & "</td>" & _
                         "<td>" & rs.Fields("Nome Deposito").Value & "</td>" & _
                         "<td>" & tipoDepositoHTML & "</td>" & _
                         "<td>" & processoDepositoHTML & "</td>" & _
                         "<td>" & statusAtual & "</td>" & _
                         "<td>" & acaoHTML & "</td>" & _
                         "</tr>"

        rs.MoveNext
    Loop

    relatorioHTML = relatorioHTML & "</table>"

    document.getElementById("CarregarDepositosCadastrados").style.display = "Block"
    document.getElementById("CarregarDepositosCadastrados").innerHTML = relatorioHTML

    rs.Close
    conn.Close
End Sub


' Função para "Editar" que limpa Tipo e Processo e atualiza Status
Sub EditarDeposito(idLinha)
    ConectarBanco

    sql = "UPDATE Dados SET Tipo = '', Processo = '', Status = 'Cadastrar' WHERE Código = " & idLinha & ";"
    conn.Execute sql

    MsgBox "Registro alterado para cadastro!"

    ' Atualizar tabela
    CarregarDepositosCadastrados
	CarregarEstoqueDepositos
End Sub


Sub CadastrarDeposito(idLinha)
    ' Conectar ao banco de dados
    ConectarBanco
    
    ' Recuperar os valores selecionados da linha específica
    Dim tipo, processo
    On Error Resume Next
    tipo = document.getElementById("TipoDeDeposito_" & idLinha).Value
    processo = document.getElementById("ProcessoDeArmazenagem_" & idLinha).Value
    On Error GoTo 0
    
    ' Validar se o usuário selecionou uma opção válida
    If tipo = "Selecione" Or tipo = "" Then
        MsgBox "Por favor, selecione um Tipo de Depósito válido!"
        Exit Sub
    End If
    
    If processo = "Selecione" Or processo = "" Then
        MsgBox "Por favor, selecione um Processo de Armazenagem válido!"
        Exit Sub
    End If

    ' Montar a consulta SQL para atualizar o banco de dados
    sql = "UPDATE Dados SET " & _
          "Tipo = '" & tipo & "', " & _
          "Processo = '" & processo & "', " & _
		  "Status = 'Cadastrado' " & _
          "WHERE Código = " & idLinha & ";"

    ' Executar a consulta
    conn.Execute sql

    ' Exibir mensagem de sucesso
    MsgBox "Registro atualizado com sucesso!"

    ' Atualizar a tabela
    CarregarDepositosCadastrados    
End Sub
