# cei-import-irpf

Scripts de importação de extratos do CEI para geração de informações para Imposto de Renda (IRPF)

### Como funciona?

Basta baixar as planilhas de transações direto no site da CEI ([[https://cei.b3.com.br | https://cei.b3.com.br]]). Use o menu 'Extratos e Informativos' > '
Extrato BM&FBOVESPA'.

### Instalar e executar

Após baixar/clonar e acessar o repositório, instale as dependências:

  python3 -m pip install -r requirements.txt

Execute o script de importação:

  python3 bulk_import.py

Pronto! Foi gerado um arquivo com seus preços médios (preco-medio.csv), um arquivo com suas vendas (vendas.csv) contendo lucros/prejuízos e os arquivos de transações no diretório 'negotiations'. Um arquivo para cada ativo.

#### Python 3

Se estiver usando o Windows, instale o Python 3 antes de tudo.

#### Customização

Nos arquivos .py existem algumas constantes que podem ser modificadas para customização. A mais importante, talvez, seja a constante 'COST_OF_OPERATION'. Mude o valor para inserir o custo de cada operação em sua corretora. Por padrão, essa contante está setada com valor 0. Se você tem um custo fixo por operação, apenas defina essa valor para a constante. Se seu custo é proporcional, vai precisar alterar o código após o comentário '# Add the cost of operation on each negotiation'

### IMPORTANTE!!

Na minha conta na CEI, tem um erro nos meses de janeiro/2019 até Abril/2019. Todos os preços de compra estão multiplicados por 10. Abri as planilhas manualmente e editei esses valores.


### Observações

Se quiser criar uma planilha de importação dos preços médios anteriores, use o modelo models/pms-anteriores.xls. Copie o modelo para import-file-xls e renomeie se achar necessário. Sempre que uma célula não tiver valor, preencha com 0. NUNCA DEIXE A CELULA VAZIA!