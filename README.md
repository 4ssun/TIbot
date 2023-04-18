# *Tibot*
Pick SAP's RBS and put into a Microsoft Sharepoint

## Objetivos
Solução em automação de processos com Python,Selenium e Pandas para o setor administrativo gerenciar novas requisições de compras feitas no SAP. Essa parte da solução consiste no Webscraping para obtenção de dados que serão manipulados em formato de dataframe no Pandas, posteriormente serão enviados em dois Sharepoints onde:
1. O primeiro receberá todas as informações do requerinte da rbs, nome, email, informações, WBS (quando houver), categoria de item, quantidade, valor, peso etc. Servindo como um backup de todos os itens de informação.
2. O segundo, por sua vez, receberá apenas o número da RBS e o e-mail do responsável pelo pedido.
O segundo Sharepoint funcionará como uma base de alimentação para uma aplicação em PowerAutomate Apps que é a responsável por notificar os funcionários do setor administrativo de quando houver novas RBS para serem analisadas e aprovadas.

No final de todo processo, o Tibot irá incrementar os pedidos de RBS em um arquivo Excel que será verificado na primeira etapa de toda nova execução para que uma mesma RBS não seja lançada mais de uma vez no Sharepoint, ainda que *case sensitive* de repetições.
Sua execução será realizada a cada uma hora durante um período de 12 horas diárias, sendo das 07 a.m até as 07 p.m objetivando que requisições abertas mais cedo ou após o horário comum de trabalho (08:00 a.m/05:00 p.m) não passem em branco e sejam lançadas no arquivo excel de processamento e notifique o setor administrativo.

## Composição do Diretório
O tibot encontra-se em sua V1 de funcionamento.
Contando com duas pastas:
1. Main (Código para testes de desenvolvimento da v2 sem interrompimentos do funcionamento da v1)
2. ***SAP*** (Abriga o que importa: O arquivo StartNew2.py)

Dentro do SAP encontramos o core.py, ExcelNew.py, StartNew.py e StartNew2.py como arquivos Python. processadas.xlsx, sap_emails.xlsx e sap_excel.xlsx como arquivos Excel sendo eles, respectivamente, do arquivo de RBS que já foram lançadas no Sharepoint pelo Tibot, relação de usuários SAP, nome e e-mail institucional e sap_excel é a transformação do arquivo .xls para .xlsx como um backup, umas vez que o pandas trabalha em DataFrames de maneira mais prática que diretamente no Excel.. 

>core.py abriga a classe Navigator() que executa toda manipulação do navegados e Iframes da etapa de Webscraping do SAP.
>
>ExcelNew.py possui a classe excel() que possui as funções de manipulação e tratamento de Excel, que é onde o arquivo originado do SAP passa por seu primeiro tratamento, redimensionando a tabela que será trabalhada.
>
>StartNew.py aqui, a classe Sap(Navigator) herda atributos da classe Navigator para realizar o Webscraping no SAP, tendo os botões, manipulação de periféricos, downloads das bases e movimentação de diretório dos arquivos. 
>
>***StartNew2.py*** Aqui é onde realmente acontece a ação, o arquivo batch chama o StartNew2, que é onde o Python chama todas as funções e classes construídas em arquivos separados préviamente, o modelo atual (v1) possui sua versão estruturada, ficando POO na v2 a ser desenvolvida,testada e emplementada em 2023 antes do fim de FY23.
## Composição do Código
Podemos dividi-lo de forma literal a um robô. Imagine que ele tenha
1. Cabeça
2. Tronco
3. Pernas

![Tibot](tibot_img.PNG)
### Cabeça
- A cabeça é onde executa a primeira parte que irá originar o arquivo Excel com as RBS's do dia até aquele horário de execução, prepara as colunas que serão usadas para o tratamento do arquivo e já coloca no path que o Pandas irá em busca para começar a manipulação de *dataframes*.

### Tronco
- O tronco é a definição das variáveis de botões e paths que serão utilizados pelo código desde a limpeza do arquivo Excel até a mudança de path do arquivo processado. A primeira atividade é a leitura do Pandas do arquivo originado do SAP que vem em .xls e transformá-lo em um *Dataframe* que será manipulado.
>pd.read_html(sap_arquivo[1])
>
Tratando colunas com nomes quebrados, comparando com dados processados e gerando o df_share1 que será lançado no primeiro Sharepoint.
>(_Na v1 no entanto, só trata a coluna Purchase para facilitar a busca dos dados que estão dentro da coluna, e posteriormente comparar com as RBS processadas para validar e verificar se há novas linhas de requisição para serem lançadas, ou se não._)
>
Se o df_share1 é lançado no primeiro Sharepoint, devemos questionar o que vai para o segundo Sharepoint, e respondendo essa pergunta: vai o df_share1 sem duplicatas!
>Uma requisição de compra pode ter mais do que um item, porém todos esses objetos estão dentro de um mesmo código de RBS, logo, não teria porquê eu passar um código de RBS repetido ainda que os itens de compra sejam diferentes e para isso usamos o drop_duplicates().
>
Vale acrescenta uma observação, uma vez que no segundo Sharepoint irá a RBS e o responsável pela compra, esse último campo pede o e-mail institucional do usuário e é aí que usamos o sap_emails.xlsx dentro de uma função que busca o nome de usuário na coluna _user name_ e retorna o valor correspondente na coluna _email_. **Todavia** sempre que um novo colaborador da empresa for contratado e ele tiver atividades a ser empenhadas no SAP ele deve ser acrescentado **MANUALMENTE** nesse arquivo **sap_excel.xlsx** com seu nome, email institucional, usuário SAP e nome completo (1º e último nome), caso contrário, no dia que esse colaborador fizer uma requisição e ele não constar no arquivo citado, ele não será citado como responsável por abrir a RBS e o aplicativo ficará sem essa informação que é um campo obrigatório, causando problemas de execução da automação.
 O código de verificação para o df_share2 do usuário e e-mail é:
>sap_emails = pd.read_excel('C:\\Users\\administrator.PLANBRAZIL\\Downloads\\Brazil IT - Tibot\\v1\\SAP\\Projetos\\sap_emails.xlsx')
def tratamento_email(usuario):
      for eml in range(len(sap_emails)):
         if usuario == sap_emails.iat[eml,0]:
          return(sap_emails.iat[eml,3])
>
E é chamado dentro do campo responsávell,realizando a execução da função no momento do preenchimento dos campos automaticamente.

Ainda que ele também seja usado no primeiro Sharepoint, ele é indispensável de fato apenas no segundo, porém feito desde o primeiro por razões de limpeza de código e reaproveitamento de funções.

## Pernas
É a finalização dos processos, pegando aquele excel que foi tratado, processado e lançado nos Sharepoints e mover de diretório, sem antes anexar o que foi lançado na execução para um arquivo que sofre appends constantes ao final de cada execução, criando um histórico para comparação da próxima atividade não relançar RBS que já foram subidas no site.
