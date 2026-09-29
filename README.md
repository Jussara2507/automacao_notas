# Automação de Faturamento e ETL de Notas Fiscais (Python)

Ferramenta desenvolvida em Python para automatizar o fluxo financeiro e o controlo de faturamento, cruzando relatórios de emissão de notas em PDF com bases de dados corporativas em Excel.

## Funcionalidades
- **Extração de Dados (ETL):** Leitura e parsing de relatórios em PDF de sistemas de faturação utilizando `pypdf` e expressões regulares (`re`).
- **Normalização e Tratamento de Dados:** Limpeza de strings, remoção de acentos e correspondência inteligente de clientes utilizando `unicodedata` e `pandas`.
- **Validação de Duplicidade:** Verificação de notas fiscais já lançadas para evitar reprocessamentos.
- **Gestão de IDs e Auto-Cadastro:** Identificação de clientes existentes e alocação dinâmica de novos cadastros com IDs sequenciais.
- **Cálculo de Prazos e Vencimentos:** Aplicação de regras específicas de prazos de pagamento (dias) e modalidades (Boleto vs. Transferência) por cliente.
- **Automação de Planilhas (Excel):** Manipulação avançada via `openpyxl` com deteção automática do mês da nota e criação dinâmica de abas mensais.
## Tecnologias Utilizadas
- Python
- Pandas
- OpenPyXL
- PyPDF
