Sub IniciarcorteCentro()
    ConectarBanco

    sql = "SELECT DISTINCT Centro, Status FROM Dados WHERE Status = 'PROGRAMADO'"
    rs.Open sql, conn

    Dim relatorioHTML
    relatorioHTML = "<table border='1'><tr>" & _
                    "<th>CENTRO</th>" & _
					"<th>PRÉ CORTE</th>" & _
                    "<th>CORTE DOCUMENTAL INICIAL</th>" & _
					"<th>CORTE DOCUMENTAL FINAL</th>" & _
                    "</tr>"

    Do While Not rs.EOF
        Dim CorteDocumentalInicialHTML, PreCorteHTML, CorteDocumentalFinalHTML
        Dim linhaClasse, rowId
        rowId = rs.Fields("Centro").Value
		
		
		' Botões de ação
        PreCorteHTML = "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>CO1P</button>" & _
				   "<button class='btn-editar' onclick='Cogi(" & rowId & ")'>COGI</button>" & _ 
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>COOIS</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>KOC4</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>VF04</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>Z54P</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>Z5A8N</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>Z070007</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>ZWM17</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>LT22</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>LT24</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>1. VL06O</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>2. VL06O</button>" &_
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>MB52</button>"

        ' Botões de ação
        CorteDocumentalInicialHTML = "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>Z678</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>MC.5</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>ZWM24N</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>LX03</button>"
				   
		        ' Botões de ação
        CorteDocumentalFinalHTML = "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>Z678</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>MC.5</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>ZWM24N</button>" & _
				   "<button class='btn-editar' onclick='CadastrarDeposito(" & rowId & ")'>LX03</button>"

        ' Montar a linha com a cor definida
        relatorioHTML = relatorioHTML & "<tr" & linhaClasse & ">" & _
                         "<td>" & rs.Fields("Centro").Value & "</td>" & _
						 "<td>" & PreCorteHTML & "</td>" & _
                         "<td>" & CorteDocumentalInicialHTML & "</td>" & _
						 "<td>" & CorteDocumentalFinalHTML & "</td>" & _
                         "</tr>"

        rs.MoveNext
    Loop

    relatorioHTML = relatorioHTML & "</table>"

    document.getElementById("IniciarcorteCentro").style.display = "Block"
    document.getElementById("IniciarcorteCentro").innerHTML = relatorioHTML

    rs.Close
    conn.Close
End Sub

Sub IniciarcorteDeposito()
    ConectarBanco

    sql = "SELECT Centro, Deposito, [Nome Deposito], Status FROM Dados WHERE Status = 'PROGRAMADO'"
    rs.Open sql, conn

    Dim relatorioHTML
    relatorioHTML = "<table border='1'><tr>" & _
                    "<th>CENTRO</th>" & _
                    "<th>DEPOSITO</th>" & _
					"<th>DESCRIÇÃO</th>" & _
					"<th>CHECK CORTE INICIAL</th>" & _
					"<th>CHECK CORTE FINAL</th>" & _
                    "</tr>"

    Do While Not rs.EOF
        Dim acaoHTML
        Dim linhaClasse, rowId
        rowId = rs.Fields("Centro").Value



        ' Montar a linha com a cor definida
        relatorioHTML = relatorioHTML & "<tr" & linhaClasse & ">" & _
                         "<td>" & rs.Fields("Centro").Value & "</td>" & _
						 "<td>" & rs.Fields("Deposito").Value & "</td>" & _
						 "<td>" & rs.Fields("Nome Deposito").Value & "</td>" & _
						 "<td> 0 </td>" & _
						 "<td> 0 </td>" & _
                         "</tr>"

        rs.MoveNext
    Loop

    relatorioHTML = relatorioHTML & "</table>"

    document.getElementById("IniciarcorteDeposito").style.display = "Block"
    document.getElementById("IniciarcorteDeposito").innerHTML = relatorioHTML

    rs.Close
    conn.Close
End Sub