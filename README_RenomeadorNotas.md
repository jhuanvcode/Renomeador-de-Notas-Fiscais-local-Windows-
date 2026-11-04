# Renomeador Automático de Notas Fiscais

Automação local desenvolvida em **Python** para identificar o nome do prestador, fornecedor ou emitente em documentos fiscais e renomear os arquivos automaticamente. O sistema processa arquivos PDF, XML e CSV, aplica regras para evitar identificações incorretas e mantém um histórico permanente para não analisar novamente arquivos que já foram tratados.

## Objetivo

Reduzir o trabalho manual de organização de notas fiscais, padronizando os nomes dos arquivos com base no fornecedor identificado no próprio documento.

O programa prioriza a segurança: considera o nome de quem emitiu a nota, bloqueia documentos emitidos pela própria empresa e mantém o nome original quando não encontra uma identificação confiável.

## Funcionalidades

- **Leitura de PDF:** extrai texto e identifica a seção do prestador/emitente, considerando a organização visual do documento.
- **OCR para PDFs digitalizados:** quando a extração de texto não é suficiente, renderiza as páginas e utiliza o mecanismo de OCR nativo do Windows.
- **Leitura de XML:** procura os dados do emitente/prestador em diferentes estruturas de XML, incluindo blocos, tabelas e layouts de notas fiscais.
- **Leitura de CSV:** identifica colunas relacionadas ao prestador, fornecedor ou emitente.
- **Proteção contra autorrenomeação:** bloqueia notas emitidas pela própria empresa, usando o CNPJ configurado e, opcionalmente, termos do nome empresarial.
- **Validação de documentos:** verifica CPF/CNPJ encontrados antes de utilizá-los nas regras de identificação e bloqueio.
- **Padronização dos nomes:** remove acentos e caracteres inválidos, converte o nome para letras maiúsculas e limita o resultado a 20 caracteres.
- **Tratamento de duplicidades:** acrescenta sufixos numéricos (`_2`, `_3`...) quando já existe um arquivo com o mesmo nome.
- **Histórico permanente:** registra arquivos tratados e verifica tamanho e data de modificação para evitar reprocessamentos desnecessários.
- **Modo de simulação:** permite visualizar as renomeações sem alterar os arquivos.
- **Relatórios e logs:** gera um relatório CSV com o resultado de cada análise e um arquivo de log para acompanhamento.
- **Verificação de sincronização:** evita processar arquivos vazios, ilegíveis ou que ainda estejam sendo sincronizados pelo OneDrive.
- **Arquivos ignorados:** permite configurar documentos que não devem ser tratados como notas fiscais.

## Como funciona

1. O programa percorre a pasta configurada e suas subpastas.
2. Filtra os arquivos com extensão PDF, XML ou CSV.
3. Consulta o histórico permanente para verificar se o arquivo já foi processado e permanece inalterado.
4. Verifica se o arquivo está disponível para leitura.
5. Extrai o nome do prestador/emitente conforme o formato do documento.
6. Aplica as regras de validação e bloqueio da própria empresa.
7. Se encontrar um nome confiável, prepara o novo nome e verifica possíveis duplicidades.
8. Executa a renomeação ou apenas registra a simulação, conforme a configuração.
9. Salva o histórico, o relatório e os logs.

```text
Pasta de notas fiscais
        |
        v
Busca recursiva por PDF, XML e CSV
        |
        v
Consulta ao histórico e verificação do arquivo
        |
        v
Extração do prestador / emitente
        |
        +---- PDF: leitura de texto -> OCR do Windows, se necessário
        |
        +---- XML: análise de tags, blocos, tabelas e layout
        |
        +---- CSV: identificação da coluna do prestador
        |
        v
Validação e regras de segurança
        |
        +---- Própria empresa -> bloqueia e mantém o arquivo
        |
        +---- Nome não confiável -> mantém o nome original
        |
        v
Padronização e verificação de duplicidade
        |
        v
Simulação ou renomeação real
        |
        v
Histórico + relatório CSV + log
```

## Tecnologias utilizadas

| Tecnologia | Utilização |
|---|---|
| Python | Linguagem principal e lógica de automação |
| pdfplumber | Extração de texto e palavras de arquivos PDF |
| pypdfium2 | Renderização de páginas PDF para processamento por OCR |
| Windows OCR | Reconhecimento de texto em documentos digitalizados |
| PowerShell | Execução do mecanismo de OCR nativo do Windows |
| XML ElementTree | Leitura e análise de estruturas XML |
| CSV | Leitura de planilhas e relatórios em formato CSV |
| JSON | Armazenamento do histórico de processamento |
| OneDrive | Compatibilidade com pastas sincronizadas localmente |

> `pdfplumber` e `pypdfium2` são dependências opcionais no código. Para utilizar a extração de texto de PDF, instale `pdfplumber`. Para o fluxo de OCR de PDF, instale também `pypdfium2` e execute em um ambiente Windows com o recurso Windows OCR disponível.

## Estrutura do projeto

```text
RenomeadorNotas/
│
├── Notas.py
├── _ocr_engine.ps1
├── .env.example
├── .gitignore
│
├── historico_renomeacao.json       # Gerado durante a execução
├── log_renomeacao_nome.txt        # Gerado durante a execução
├── relatorio_renomeacao_nome.csv  # Gerado durante a execução
│
└── debug_extracao/                # Opcional, quando o debug está ativado
```

Os arquivos de histórico, log, relatório e debug são gerados durante a execução. O script PowerShell `_ocr_engine.ps1` também pode ser recriado automaticamente pelo programa.

## Configuração

### 1. Requisitos

- Windows com PowerShell disponível.
- Python 3.9 ou superior.
- Acesso à pasta local ou sincronizada que contém as notas fiscais.
- Para leitura de PDFs com texto: `pdfplumber`.
- Para OCR de PDFs digitalizados: `pypdfium2` e o mecanismo Windows OCR com idioma adequado instalado.

Instale as dependências:

```bash
pip install pdfplumber pypdfium2
```

Se for utilizar apenas a leitura de XML e CSV, essas bibliotecas de PDF não são necessárias.

### 2. Configurar as variáveis de ambiente

Crie suas variáveis de ambiente com base no arquivo `.env.example`. O programa utiliza as variáveis do ambiente do Windows; ele não carrega automaticamente um arquivo `.env`.

| Variável | Obrigatória | Descrição |
|---|---:|---|
| `NOTAS_PASTA_BASE` | Sim | Caminho da pasta raiz que contém as notas fiscais. |
| `EMPRESA_CNPJ` | Sim | CNPJ da empresa que recebe as notas. Deve ser válido. |
| `EMPRESA_TERMOS_BLOQUEIO` | Não | Termos usados para reconhecer o nome da própria empresa quando não há CNPJ legível. |
| `ARQUIVOS_IGNORADOS` | Não | Nomes de arquivos que devem ser ignorados, separados por vírgula. |
| `DRY_RUN` | Não | `true` simula; `false` realiza a renomeação. |
| `SALVAR_DEBUG` | Não | `true` salva texto extraído para investigação; `false` desativa. |

Exemplo de configuração:

```env
NOTAS_PASTA_BASE=C:\NotasFiscais
EMPRESA_CNPJ=00.000.000/0001-00
EMPRESA_TERMOS_BLOQUEIO=minhaempresa+matriz|minhaempresa+filial
ARQUIVOS_IGNORADOS=comprovante-interno.pdf,manual.pdf
DRY_RUN=true
SALVAR_DEBUG=false
```

**Importante:** substitua os valores fictícios pelos dados corretos do ambiente. O CNPJ do exemplo é apenas ilustrativo e não deve ser utilizado como configuração real.

No Windows, as variáveis podem ser configuradas nas Variáveis de Ambiente do sistema ou do usuário. Depois de configurá-las, abra um novo terminal para executar o programa.

### 3. Executar

No terminal, acesse a pasta do projeto e execute:

```bash
python Notas.py
```

## Modos de execução

### Simulação

Configure:

```env
DRY_RUN=true
```

O programa analisa os documentos e registra os nomes que seriam aplicados, sem renomear os arquivos. É o modo recomendado para validar a configuração antes da primeira execução real.

### Renomeação real

Após conferir o relatório da simulação, configure:

```env
DRY_RUN=false
```

Nessa configuração, os arquivos identificados com segurança são renomeados na própria pasta.

## Regras de segurança e identificação

- O nome utilizado é o do prestador/fornecedor/emitente, não o do tomador ou destinatário.
- Documentos identificados como emitidos pela própria empresa são bloqueados.
- Por padrão, o bloqueio considera também filiais que compartilham a mesma raiz do CNPJ (os oito primeiros dígitos).
- O bloqueio por nome pode ser configurado por grupos em `EMPRESA_TERMOS_BLOQUEIO`. Use `+` para combinar termos de um mesmo grupo e `|` para separar grupos.
- Endereços, bairros, CEPs, e-mails, telefones e rótulos não devem ser utilizados como nome do fornecedor.
- Se houver ambiguidade ou não for possível encontrar um nome confiável, o arquivo permanece com o nome original.
- O nome final é limitado a 20 caracteres, sem espaços por padrão.

## Histórico e reprocessamento

O arquivo `historico_renomeacao.json` armazena informações dos arquivos tratados, incluindo caminho, nome original, nome atual, fornecedor identificado, status, tamanho e data de modificação.

Nas execuções seguintes, o programa compara o tamanho e a data de modificação do arquivo com o registro salvo. Se os dados permanecerem iguais, o arquivo é ignorado.

O histórico é atualizado em execução real. No modo de simulação, as alterações não são gravadas no histórico.

O código possui a configuração `FORCAR_REPROCESSAMENTO`, definida diretamente em `Notas.py`. Quando ativada, ela ignora o histórico e força a análise dos arquivos novamente. Use essa opção com cuidado.

## Arquivos gerados

| Arquivo | Finalidade |
|---|---|
| `log_renomeacao_nome.txt` | Registro de eventos, avisos e erros da execução. |
| `relatorio_renomeacao_nome.csv` | Resultado por arquivo, incluindo nome original, novo nome, status, método de identificação e motivo. |
| `historico_renomeacao.json` | Histórico permanente usado para evitar reprocessamentos. |
| `debug_extracao/` | Arquivos de texto extraído para diagnóstico, quando `SALVAR_DEBUG=true`. |

O relatório CSV utiliza `;` como separador e é salvo em UTF-8 com BOM, facilitando a abertura em ferramentas como Excel.

## Privacidade e cuidados

Notas fiscais podem conter dados pessoais, informações financeiras e dados de fornecedores. Por isso:

- Não envie notas fiscais reais para repositórios públicos.
- Não publique CNPJs, caminhos internos ou nomes de fornecedores em arquivos de configuração.
- Mantenha `.env` e os arquivos de histórico, log, relatório e debug fora do controle de versão.
- Ative `SALVAR_DEBUG` somente quando necessário, pois os arquivos de debug podem conter o texto extraído dos documentos.
- Faça uma cópia de segurança antes de executar a renomeação real em uma pasta de produção.
- Valide os resultados no modo de simulação antes de desativar `DRY_RUN`.

O `.gitignore` do projeto já prevê a exclusão de documentos fiscais e dos principais arquivos gerados pela automação.

## Observações

- O programa percorre subpastas da pasta raiz configurada.
- São considerados arquivos com extensão `.pdf`, `.xml` e `.csv`.
- Arquivos temporários do Office iniciados por `~$` são ignorados.
- O OCR depende do ambiente Windows e do mecanismo de reconhecimento de texto disponível no sistema.
- O histórico identifica arquivos pelo caminho, tamanho e data de modificação; se o conteúdo de um arquivo mudar, ele poderá ser analisado novamente.
- Arquivos cujo nome não foi identificado com segurança permanecem sem renomeação e podem ser analisados novamente em uma execução futura.

---

Desenvolvido para automatizar a organização de documentos fiscais, reduzir tarefas manuais e aumentar a padronização e a rastreabilidade do processo.
