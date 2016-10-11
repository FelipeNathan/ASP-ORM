<!-- #include file = '../Repository/Base/SQLFormatter.asp' -->
<%
	Class ClassEntityManager
        
        Public ObjConexao

        Private StrSQL, objRs, Identity, ParentClass
        Private Table, Fields, Values, Types
        
        'Iniciações da classe
        'Instância de três dicionários 
        'Fields: Relação entre o Campo do BD com Propriedade da classe "pai"
        'Values: Relação entre o Campo do BD com o VALOR da Propriedade da classe "pai"
        'Types:  Tipagem dos Campos do BD
        Sub Class_Initialize()
            Set Fields = Server.CreateObject("Scripting.Dictionary")
            Set Values = Server.CreateObject("Scripting.Dictionary")
            Set Types  = Server.CreateObject("Scripting.Dictionary")
        End Sub
        
        'Parâmetros passado no "construtor" da classe" para identificar a primary key (identity) e a tabela
        Public Default Function Init(pTable, pIdentity, pParentClass)
            Table = pTable
            Identity = pIdentity
            Set ParentClass = pParentClass
            Set Init = me
        End Function
        
        'Registra em um dicionário o campo do banco referenciando a propriedade da classe
        'e em outro dicionário a tipagem dos dados
        Public Sub Register(pKey, pField, pType)
            If Not Fields.Exists(pKey) Then
                Call Fields.Add(pKey, pField)
                Call Types.Add(pKey, pType)
            End If
        End Sub
        
        'Atualiza o dicionário de Valores, relacionando o campo do banco de dados com o VALOR da propriedade da classe pai
        Private Sub UpdateDictionary()
            Dim Key
            For Each Key In Fields.Keys
                Call SetValue(Key, Eval("ParentClass."& Fields( Key )))
            Next
        End Sub

        'Registra os valores no dicionário Values( campoBD, valorPropriedade )
        Private Sub SetValue(pField, pValue)
            If pValue = "" Then 
                Exit Sub
            End If
    
            If Values.Exists(pField) Then
                Values.Item(pField) = pValue
            Else
                Call Values.Add(pField, pValue)
            End If
        End Sub
        
        'Métodos de persistência
        'Método Save realiza a inserção OU atualização da tabela
        ' - o que define a atualização é o identity, caso a propriedade do identity estiver vazia, a classe realiza uma inserção
        '   caso contrário, irá atualizar a tabela de acordo com o identity (where identity = valorPropriedadeIdentity)
        '   e ao final, atribui o ultimo Id inserido na tabela para o identity da classe "pai"
        Public Sub Save()
            UpdateDictionary()
            
            If Values.Exists( Identity ) Then
                If Values.Item( Identity ) <> "" Then
                    StrSQL = UpdateQuery()
                Else
                    StrSQL = InsertQuery() & InsertValuesQuery()
                End If
            Else
                StrSQL = InsertQuery() & InsertValuesQuery()
            End If
            
            If StrSQL = "" Then Exit Sub

            Set objRs = ObjConexao.Execute( StrSQL )

            If Fields.Exists( Identity ) Then
                Execute("ParentClass." & Fields( Identity ) & " = " & objRs("ID"))
            End If
        End Sub
        
        'Método Delete, realiza o delete da tabela de acordo com o identity, se não passar um identity vai gerar um erro
        Public Sub Delete()
            UpdateDictionary()

            StrSQL = DeleteQuery()
            If StrSQL = "" Then Exit Sub
            ObjConexao.Execute( StrSQL )
        End Sub

        'Método Load: A definir a real execução do método
        Public Sub Load()
            UpdateDictionary()

            StrSQL = SelectQuery()
            If StrSQL = "" Then Exit Sub
            Set objRs = ObjConexao.Execute( StrSQL )

            If Not objRs.Eof Then
                Dim Key
                For Each Key In Fields.Keys
                    Execute("ParentClass." & Fields.Item( Key ) & " = """ & objRs( Key ) & """")
                Next
            End If
        End Sub

        'Métodos para gerar as queries
        Private Function InsertQuery()
            Dim Key, tempArray : tempArray = Array()
            For Each Key In Values.Keys
                If Key <> Identity Then
                    Redim Preserve tempArray( UBound(tempArray) + 1 )
                    tempArray( Ubound(tempArray) ) = Key
                End If
            Next
            InsertQuery = " SET NOCOUNT ON; INSERT INTO " & Table & " (" & Join(tempArray, ", ") & " ) "
        End Function

        Private Function InsertValuesQuery()
            Dim Key, tempArray : tempArray = Array()
            For Each Key In Values.Keys
                If Key <> Identity Then
                    Redim Preserve tempArray( UBound(tempArray) + 1 )
                    tempArray( Ubound(tempArray) ) = FormatType( Values.Item( Key ), Types.Item( Key ) )
                End If
            Next 
            InsertValuesQuery = " VALUES ( " & Join(tempArray, ", ") & " ) SELECT @@IDENTITY AS ID "
        End Function

        Private Function UpdateQuery()
            Dim Key, IdentityQuery, tempArray : tempArray = Array()
            For Each Key in Values.Keys
                If Key = Identity Then
                    IdentityQuery = Identity & " = " & FormatType( Values.Item( Key ), Types.Item( Key ) )
                Else
                    Redim Preserve tempArray( UBound(tempArray) + 1 )
                    tempArray( Ubound(tempArray) ) = Key & " = " & FormatType( Values.Item( Key ), Types.Item( Key ) )
                End If
            Next

            UpdateQuery = " SET NOCOUNT ON; UPDATE " & Table & _
                          " SET " & Join(tempArray, ", ") &_
                          " WHERE " & IdentityQuery &_ 
                          " SELECT " & Values.Item( Identity ) & " AS ID "
        End Function
        
        Private Function DeleteQuery()
            If Not Values.Exists( Identity ) Then Exit Function
            DeleteQuery = "DELETE FROM " & Table & " WHERE " & Identity & " = " & FormatType( Values.Item( Identity ), Types.Item( Identity ) )
        End Function

        Private Function SelectQuery()
            If Not Values.Exists( Identity ) Then Exit Function

            SelectQuery = _
                        " SELECT " & Join(Fields.Keys, ", ") & " FROM " & table &_
                        " WHERE " & Identity & " = " & FormatType( Values.Item( Identity ), Types.Item( Identity ) )
            
        End Function

        'Método responsável por formatar os valores das propriedades para o formato correto aceito pelo banco de dados
        Private Function FormatType(pValue, pType)

            'Permite deixar salvar campos "NULL"
            If UCase(pValue) = "NULL" Then 
                FormatType = pValue
                Exit Function
            End If

            Select Case LCase(pType)
                Case "decimal"
                    FormatType = parseDecimalToSQL(pValue)

                Case "string"
                    FormatType = ToSqlString(pValue, true)
                
                Case "bool"
                    FormatType = ToSqlBit(pValue)
        
                Case Else
                    FormatType = pValue
            End Select
        End Function

	End Class
%>
