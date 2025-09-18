Sub PreTicket(idLinha)

    Conexao_SAP("ZWM0235")	    

    Dim placa, carreta, motorista, contato, sql, balanca, pesoNF, NNotas
    placa = document.getElementById("placa_" & idLinha).Value
    carreta = document.getElementById("carreta_" & idLinha).Value
    motorista = document.getElementById("motorista_" & idLinha).Value
    contato = document.getElementById("contato_" & idLinha).Value

    ' Conectar ao banco de dados e verificar o tipo de produto
    ConectarBanco


    balanca = InputBox("Informe qual balança de pesagem:")
    NNotas = InputBox("Informe o Numero das Notas")
	pesoNF = InputBox("Informe o peso da nota fiscal:")

    If balanca = "" Then
        MsgBox "O campo balança não pode estar vazio. Por favor, informe o número da balança."
        Exit Sub
    End If
	
	If NNotas = "" Then
        MsgBox "Preecha ao minimo o Numero de uma Nota."
        Exit Sub
    End If

    If pesoNF = "" Then
        MsgBox "O campo peso da nota fiscal não pode estar vazio. Por favor, informe o peso."
        Exit Sub
    End If

    ' Conexão SAP e execução do script
    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If

    With session
        .findById("wnd[0]/usr/ctxtP_CENTRO").text = "376"
        .findById("wnd[0]/usr/txtP_BALAN").text = balanca
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[48]").press
        .findById("wnd[0]/usr/cmbV_CAM_V_C").key = "C"
        .findById("wnd[0]/usr/ctxtV_TIP_OPER").text = "105"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/txtV_MATERIAL").text = "922593"
        .findById("wnd[0]/usr/txtV_OBSERV").text = "NFe: " & NNotas
        .findById("wnd[0]/usr/txtV_MOTORISTA").text = motorista
        .findById("wnd[0]/usr/txtV_TEL").text = contato
        .findById("wnd[0]/usr/txtV_PESO_NF").text = pesoNF
        .findById("wnd[0]").sendVKey 0
		
		On Error Resume Next
		
		.findById("wnd[0]/usr/ctxtV_PLACA").text = placa
		.findById("wnd[0]").sendVKey 0
		.findById("wnd[1]/usr/btnBUTTON_1").press
		.findById("wnd[0]/usr/ctxtV_PLC_1").text = carreta
		.findById("wnd[0]").sendVKey 0
		.findById("wnd[1]/usr/btnBUTTON_1").press

		If Err.Number <> 0 Then
		
			'.findById("wnd[0]/tbar[0]/btn[11]").press		 'salvar dados
		
		End If		
			'.findById("wnd[0]/tbar[0]/btn[11]").press  'salvar dados
			'.findById("wnd[0]/tbar[0]/btn[3]").press
    End With

    ' Atualizar o status no banco de dados


		sql = "UPDATE Cadastro SET Status = 'RECEPCIONADO*', HoraEntrada = #" & FormatDateTime(Now, vbGeneralDate) & "# WHERE Código = " & idLinha



    conn.Execute sql

    MsgBox "Registro atualizado com sucesso!"

    MontarCelula
    
End Sub

Sub Entrada(idLinha)

    Conexao_SAP("ZWM0235")
	
    Dim placa, sql, balanca
    placa = document.getElementById("placa_" & idLinha).innerText
    balanca = InputBox("Informe qual balança de pesagem:")

    If balanca = "" Then
        MsgBox "O campo balança não pode estar vazio. Por favor, informe o número da balança."
        Exit Sub
    End If

    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If

    With session
        .findById("wnd[0]/usr/ctxtP_CENTRO").text = "376"
        .findById("wnd[0]/usr/txtP_BALAN").text = balanca
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[9]").press
        .findById("wnd[1]/usr/txtV_PL_PROC").text = placa
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[2]").press
        .findById("wnd[0]/tbar[1]/btn[6]").press
		
		On Error Resume Next
		
        '.findById("wnd[1]/usr/btnBUTTON_1").press 'para quando nao tiver cadastrado
        '.findById("wnd[1]/usr/btnBUTTON_1").press 'para quando nao tiver cadastrado	

		If Err.Number <> 0 Then
		
			Dim resposta, valorCelula
			valorCelula = .findById("wnd[0]/usr/txtV_PESO_ENTR").text		
			resposta = MsgBox("O volume capturado é: " & valorCelula & " Deseja continuar?", vbYesNo + vbQuestion, "Pesagem")

			If resposta = vbYes Then
				.findById("wnd[0]/tbar[0]/btn[11]").press
				.findById("wnd[0]/tbar[0]/btn[3]").press
				.findById("wnd[0]/tbar[0]/btn[3]").press
				.findById("wnd[0]/tbar[0]/btn[3]").press	
				ConectarBanco
				  
				sql = "UPDATE Cadastro SET " & _
					  "Status = 'EM ANDAMENTO', " & _
					  "HoraEntrada = Now(), " & _
					  "Bloqueio = 'NAO' " & _
					  "WHERE Código = " & idLinha & ";"

				conn.Execute sql

				MsgBox "Registro atualizado com sucesso!"

				MontarCelula	
				
				exit sub

				
			Else	
			MontarCelula
			
			exit sub
			End If	
			
		End If 
		
			valorCelula = .findById("wnd[0]/usr/txtV_PESO_ENTR").text		
			resposta = MsgBox("O volume capturado é: " & valorCelula & " Deseja continuar?", vbYesNo + vbQuestion, "Pesagem")

			If resposta = vbYes Then
				.findById("wnd[0]/tbar[0]/btn[11]").press
				.findById("wnd[0]/tbar[0]/btn[3]").press
				.findById("wnd[0]/tbar[0]/btn[3]").press
				.findById("wnd[0]/tbar[0]/btn[3]").press				
			Else			
				Exit Sub
			End If	

    End With

    ConectarBanco
	  
    sql = "UPDATE Cadastro SET " & _
          "Status = 'EM ANDAMENTO', " & _
          "HoraEntrada = Now(), " & _
          "Bloqueio = 'NAO' " & _
          "WHERE Código = " & idLinha & ";"

    conn.Execute sql

    MsgBox "Registro atualizado com sucesso!"

    MontarCelula	
    
End Sub

Sub Saida(idLinha)

    Conexao_SAP("ZWM0235")
    
    Dim placa, sql, balanca
    placa = document.getElementById("placa_" & idLinha).innerText

    If Not IsObject(application) Then
       Set SapGuiAuto  = GetObject("SAPGUI")
       Set application = SapGuiAuto.GetScriptingEngine
    End If
    If Not IsObject(connection) Then
       Set connection = application.Children(0)
    End If
    If Not IsObject(session) Then
       Set session    = connection.Children(0)
    End If
    If IsObject(WScript) Then
       WScript.ConnectObject session,     "on"
       WScript.ConnectObject application, "on"
    End If

    ' Captura o valor da balança
    balanca = InputBox("Digite o valor da balança:", "Entrada de Dados")
    If balanca = "" Then
        MsgBox "O valor da balança não pode estar vazio.", vbExclamation, "Erro"
        Exit Sub
    End If

    With session
        .findById("wnd[0]/usr/ctxtP_CENTRO").text = "376"
        .findById("wnd[0]/usr/txtP_BALAN").text = balanca
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[9]").press
        .findById("wnd[1]/usr/txtV_PL_PROC").text = placa
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/tbar[1]/btn[6]").press
        .findById("wnd[0]/usr/chkCHK_IMPRI").selected = false 'imprimir titeck    
        
        Dim resposta, valorCelulas
		.findById("wnd[0]/tbar[1]/btn[6]").press
        valorCelulas = .findById("wnd[0]/usr/txtV_PESO_SAI").text        
        resposta = MsgBox("O volume capturado é: " & valorCelulas & " Deseja continuar?", vbYesNo + vbQuestion, "Pesagem")

        If resposta = vbYes Then
            '.findById("wnd[0]/tbar[1]/btn[6]").press 'gravar ao peso
            '.findById("wnd[0]/tbar[0]/btn[11]").press  'gravar ao peso     
        Else             
            Exit Sub
        End If            

    End With

    ConectarBanco
      
    sql = "UPDATE Cadastro SET " & _
          "Status = 'VEICULO LIBERADO', " & _
          "HoraSaida = Now(), " & _
          "Bloqueio = 'NAO' " & _
          "WHERE Código = " & idLinha & ";"

    conn.Execute sql

    MsgBox "Registro atualizado com sucesso!"

    MontarCelula    
    
End Sub
