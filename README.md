# ASP-ORM - Entity Manager

O Entity Manager é a classe responsável por mapear os campos da tabela do banco de dados;

Métodos a serem utilizados:

<b> Construtor </b> <i>(3 parâmetros)</i>: O construtor faz a ligação da tabela do banco de dados com a classe que corresponde à tabela, mapeando também o PK

<ul>
  <li><b>pTable</b>: Nome da tabela a ser mapeada</li>
  <li><b>pIdentity</b>: Nome do campo identificador da tabela</li>
  <li><b>pParentClass</b>: A classe que está chamando o EntityManager</li>
</ul>

<p>Ex:</p> 
<p>Dim EntityManager</p>
<p>Set EntityManager = (new ClassEntityManager)("NomeTabela", "nome_campo_chave_primaria", me)</p>

<p></p>
<b> Register </b> <i>(3 parâmetros)</i>: Registra as propriedades da classe que correspondem ao campo do banco de dados e seu tipo.
<ul>
  <li><b>pKey</b>: Nome do campo no banco de dados</li>
  <li><b>pField</b>: Nome da propriedade da classe que corresponde ao campo no banco de dados</li>
  <li><b>pType</b>: Tipo do campo no banco de dados: (decimal, string, bool, int) </li>
</ul> 

<p></p>
<b> Save </b><i>(sem parâmetro)</i>: Funciona como o saveOrUpdate do hibernate, caso seja atribuido algum valor ao PK da classe, o EntityManager entende como uma alteração do objeto realizando um update, caso o campo identificador esteja vazio, o EntityManager entende como um novo registro no banco.

<p></p>
<b> Delete </b><i>(sem parâmetro)</i>: Deleta o objeto da classe que esteja com o campo PK da classe atribuido.

<p></p>
<b> Load </b><i>(sem parâmetro)</i>: Procura apenas um registro no banco, o ID que for atribuido ao campo PK da classe. (Este método pode ser revisado)


<h2> Observação 1 </h2>
Não foi criado método para ler uma coleção de dados pois acreditava-se que poderia gerar um carregamento desnecessário à memória do servidor criar várias instâncias do objeto sendo que a busca é feita com as conexões do próprio ASP.

<h2> Observação 2 </h2>
Método para formatar o tipo do campo não foi implementado dentro da classe, tendo que criar o seu próprio para validar o tipo do banco de dados da sua escolha.

<h2> Observação 3 </h2>
Em alguns métodos há o "SET NO COUNT ON;" que é usado no TransactSQL do SQL Server, caso o banco de dados da sua escolha não tenha esta função, deve-se remover.
