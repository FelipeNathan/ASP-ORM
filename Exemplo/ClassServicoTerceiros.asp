<%
	Class classServicoTerceiros
    
        Private StrSQL, EntityManager, ObjConexao

        Public Id
        Public MesesVigencia
        Public TipoServico
        Public Descricao
        Public Valor
        Public AliqIOF
        Public AliqPIS
        Public AliqCOFINS
        Public Ativo
       
        Public Default Function Init(Conexao)
            Set ObjConexao = Conexao
            
            'Configuração do EntityManager
            'Instância da classe passando o nome da tabela e a primary key (identity) por parâmetro
            Set EntityManager = (new ClassEntityManager)("servico_terceiros", "codigo_servico_terceiros", me)

            'Registrando os campos do banco referenciando as propriedades da classe Register( banco, classe )
            Call EntityManager.Register("codigo_servico_terceiros", "Id", "int")
            Call EntityManager.Register("descricao", "Descricao", "string")
            Call EntityManager.Register("valor", "Valor", "decimal")
            Call EntityManager.Register("meses_vigencia", "MesesVigencia", "int")
            Call EntityManager.Register("aliq_iof", "AliqIOF", "decimal")
            Call EntityManager.Register("aliq_pis", "AliqPIS", "decimal")
            Call EntityManager.Register("aliq_cofins", "AliqCOFINS", "decimal")
            Call EntityManager.Register("tipo_servico", "TipoServico", "int")
            Call EntityManager.Register("ativo", "Ativo", "bool")

            'Objeto de conexão
            Set EntityManager.ObjConexao = Conexao
    
            Set Init = me
        End Function

        Public Sub Salvar()
            Call EntityManager.Save()
        End Sub

        Public Sub Deletar()
            Call EntityManager.Delete()
        End Sub 

        Public Sub Abrir(pId)
            me.Id = pId
            Call EntityManager.Load()
        End Sub 

		'Para buscar uma coleção de dados
        Public Function BuscarTodos()
            StrSQL = _
                    " SELECT campos " &_
                    " FROM tabela     "
            Set BuscarTodos = ObjConexao.Execute( StrSQL )
        End Function
	End Class
%>
