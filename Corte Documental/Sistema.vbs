Sub AlternarPagina(paginaAtiva)
    Dim paginas
    paginas = Array("PaginaInicial","PaginaDepositos", "PaginaEstoque", "PaginaProgramaçao", "PaginaCorte")
    
    Dim i
    For i = LBound(paginas) To UBound(paginas)
        If paginas(i) = paginaAtiva Then
            Document.getElementById(paginas(i)).style.display = "block"
            
            If paginaAtiva = "PaginaDepositos" Then

			CarregarDepositosCadastrados()

            End If
            
            If paginaAtiva = "PaginaEstoque" Then
			
			CarregarEstoqueDepositos()

            End If		

			if paginaAtiva = "PaginaProgramaçao" Then

			end if		
			
        Else
            Document.getElementById(paginas(i)).style.display = "none"
        End If
    Next
End Sub

