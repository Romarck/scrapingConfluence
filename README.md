# scrapingConfluence
# Exportador de Conteúdo do Confluence

Esta aplicação Streamlit cria uma ferramenta para exportar conteúdos do Confluence (plataforma de colaboração em equipe) para diferentes formatos de arquivo.

## Funcionalidades Principais

### Interface de Usuário via Streamlit
- Formulário para inserir credenciais do Confluence (URL base, chave do espaço, nome de usuário, token de acesso)
- Opções para selecionar o formato de saída (txt, md, pdf, json)
- Checkbox para processar anexos

### Autenticação no Confluence
- Suporta autenticação via token de acesso pessoal ou combinação de usuário/token
- Funciona tanto com Confluence Cloud quanto com instalações locais

### Extração de Conteúdo
- Busca todas as páginas de um espaço específico do Confluence
- Extrai o conteúdo HTML das páginas e o converte para texto
- Preserva blocos de código
- Opcionalmente busca e processa anexos das páginas

### Processamento de Anexos
- Baixa anexos das páginas
- Extrai texto de vários tipos de arquivo:
 - PDFs (usando PyPDF2)
 - Documentos Word (usando python-docx)
 - Planilhas Excel (usando openpyxl)
 - Apresentações PowerPoint (usando python-pptx)
 - Imagens (identificando, mas não extraindo texto)

### Exportação em Múltiplos Formatos
- Texto simples (.txt)
- Markdown (.md)
- PDF (usando FPDF)
- JSON

### Empacotamento e Download
- Cria uma estrutura de diretórios organizada com timestamp
- Compacta todos os arquivos exportados em um arquivo ZIP
- Fornece um botão para baixar o ZIP completo

## Fluxo de Execução

Quando o usuário clica no botão "Carregar Dados do Confluence e Criar ZIP", a aplicação:
1. Valida as entradas do usuário
2. Cria diretórios para armazenar os arquivos exportados
3. Busca e processa todas as páginas do espaço Confluence especificado
4. Salva cada página no formato escolhido
5. Cria um arquivo ZIP contendo todos os arquivos
6. Disponibiliza um botão para baixar o ZIP

Esta ferramenta é útil para organizações que desejam fazer backup ou migrar conteúdo do Confluence, ou para quem quer disponibilizar o conteúdo em formatos mais acessíveis para uso offline.
