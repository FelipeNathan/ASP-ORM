<!-- #include file = '../ClassEntityManager.asp' -->
<!-- #include file = 'ClassServicoTerceiros.asp' -->
<%
	Dim Operacao, ServicoTerceiros, ServicoTerceirosEmpresas
    Dim strJSON, objJSON, objRS
    Dim Empresas, Emp

    Set ServicoTerceiros = (new classServicoTerceiros)(ObjConn)
    
    Operacao = Request("operacao")
    
    Select Case Operacao
        Case "Save"
            
            With ServicoTerceiros
                .Descricao = Request("descricao")
                .Valor = Request("valor")
                .MesesVigencia = Request("validade")
                .AliqIOF = Request("aliqIof")
                .AliqPIS = Request("aliqPis")
                .AliqCOFINS = Request("aliqCofins")
                .TipoServico = Request("tipoServico")
                .Ativo = Request("ativo")
                If Request("id") <> "" Then
                    .Id = Request("id") 'Caso tenha ID realiza update, senão realiza o save
                End If
                .Salvar()
            End With

        Case "Abrir"
            Call ServicoTerceiros.Abrir(Request("id"))
            
            Response.Write ServicoTerceiros.Id
            Response.Write ServicoTerceiros.Descricao
            Response.Write ServicoTerceiros.Valor
            Response.Write ServicoTerceiros.MesesVigencia)
            Response.Write ServicoTerceiros.AliqIOF
            Response.Write ServicoTerceiros.AliqPIS
            Response.Write ServicoTerceiros.AliqCOFINS
            Response.Write ServicoTerceiros.TipoServico
            Response.Write ServicoTerceiros.Ativo
			
		Case "Deletar"
			ServicoTerceiros.Id = Request("Id")
			Call ServicoTerceiros.Deletar()
			
    End Select
%>
