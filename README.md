<div align="center">
  <h1>⚙️ AXL Controller — IWOS</h1>
  <p>Automação desktop para consulta, extração e organização de relatórios operacionais do IWOS.</p>

  <p>
    <img src="https://img.shields.io/badge/Python_3.11-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python 3.11" />
    <img src="https://img.shields.io/badge/Windows-0078D4?style=for-the-badge&logo=windows11&logoColor=white" alt="Windows" />
    <img src="https://img.shields.io/badge/UI_Automation-5C2D91?style=for-the-badge&logo=windows&logoColor=white" alt="Windows UI Automation" />
    <img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white" alt="Pandas" />
    <img src="https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="Microsoft Excel" />
  </p>
</div>

---

## Visão geral

O **AXL Controller — IWOS** automatiza um fluxo completo dentro da aplicação desktop IWOS. O script abre o aplicativo ClickOnce, realiza o login, navega pelos módulos de Recursos Humanos, consulta os dados e gera relatórios Excel prontos para análise.

A automação foi criada para reduzir tarefas manuais repetitivas e padronizar a coleta de funcionários, faltas e atestados. Todo o processo ocorre pela interface gráfica do IWOS, utilizando os identificadores disponibilizados pelo Windows UI Automation.

## Fluxo automatizado

Ao executar o projeto, o script realiza as seguintes etapas:

1. Procura uma janela principal do IWOS já aberta.
2. Se necessário, inicia o aplicativo pelo atalho ClickOnce `..appref-ms`.
3. Lê usuário e senha do arquivo `.env`.
4. Preenche a tela de login e clica em `Entrar (F9)`.
5. Abre `Recursos Humanos → Quadro de Funcionários`.
6. Abre o `Column Chooser` pelo cabeçalho `ID` e adiciona a coluna `Demissão`.
7. Deixa abertas as abas globais:
   - `Recursos Humanos → Painéis → Faltas`
   - `Recursos Humanos → Painéis → Atestados`
8. Consulta o Quadro de Funcionários para:
   - `MG Setel`
   - `Delta`
9. Percorre toda a grade virtualizada, página por página, e exporta os funcionários de cada empresa.
10. Consulta uma única vez os painéis globais de Faltas e Atestados.
11. Aplica o filtro `Earlier this year`:
    - na coluna `Data`, em Faltas;
    - na coluna `Início`, em Atestados.
12. Percorre integralmente as duas grades usando a barra vertical do DevExpress.
13. Exporta todas as colunas dos painéis, exceto `Ações`.
14. Salva os quatro relatórios em uma pasta identificada pela data e hora da execução.
15. Fecha o IWOS automaticamente após a conclusão bem-sucedida.

Se ocorrer um erro antes da conclusão, o IWOS permanece aberto para facilitar o diagnóstico.

## Relatórios gerados

Cada execução cria uma pasta com o padrão:

```text
exports/DD MM AAAA - HH MM/
```

Exemplo de saída:

```text
exports/
└── 07 08 2026 - 10 35/
    ├── IWOS-MGS-2026-08-07.xlsx
    ├── IWOS-DELTA-2026-08-07.xlsx
    ├── IWOS-FALTAS-2026-08-07.xlsx
    └── IWOS-ATESTADOS-2026-08-07.xlsx
```

### Funcionários

Os arquivos de funcionários são separados por empresa e contêm:

- ID
- Nome
- Centro
- Função
- Admissão
- Disponibilidade
- Demissão

### Faltas

O relatório é global, sem divisão por empresa, e contém:

- ID
- Centro
- Funcionário
- Função
- Data
- Motivo
- Status

### Atestados

O relatório também é global e contém:

- ID
- Centro
- Funcionário
- Função
- Patologia
- Início
- Término
- Dias

Os arquivos possuem cabeçalhos formatados, filtros automáticos, linhas alternadas, painel superior congelado e configuração de impressão em paisagem.

## Tecnologias utilizadas

### Automação desktop

- **Pywinauto** para localizar janelas, preencher campos, clicar em menus e controlar grades.
- **Windows UI Automation (UIA)** como backend de acessibilidade e automação da interface.
- **ClickOnce / Application Reference** para iniciar a instalação local do IWOS.

### Dados e relatórios

- **Pandas** para estruturar os registros em tabelas e DataFrames.
- **OpenPyXL** para gerar, estilizar e configurar os arquivos `.xlsx`.
- **Microsoft Excel** como formato final dos relatórios; o Excel não precisa estar aberto durante a execução.

### Plataforma

- **Python 3.11** como runtime recomendado.
- **Windows** como sistema obrigatório, pois a automação depende da UI Automation e da aplicação desktop IWOS.

## Estrutura do projeto

```text
.
├── .env.example       # Modelo das credenciais locais
├── .gitignore         # Arquivos privados e gerados ignorados pelo Git
├── README.md          # Documentação do projeto
├── requirements.txt   # Dependências Python fixadas
├── script.py          # Fluxo completo de automação e exportação
└── exports/           # Relatórios gerados, separados por execução
```

## Pré-requisitos

Antes de executar, confirme que o ambiente possui:

- Windows 10 ou Windows 11;
- Python 3.11 instalado;
- IWOS instalado e acessível pelo atalho ClickOnce;
- usuário com permissão para acessar os módulos de Recursos Humanos;
- sessão do Windows desbloqueada durante toda a automação.

O atalho atualmente utilizado pelo script é:

```text
%APPDATA%\Microsoft\Windows\Start Menu\Programs\..appref-ms
```

Caso o atalho seja movido ou renomeado, atualize a constante `ATALHO_IWOS` em `script.py`.

## Instalação

Abra o PowerShell na pasta do projeto e, de preferência, crie um ambiente virtual:

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Variáveis de ambiente

Copie o arquivo de exemplo:

```powershell
Copy-Item .env.example .env
```

Depois, preencha o `.env` com as credenciais reais:

```env
IWOS_USUARIO=seu_usuario
IWOS_SENHA=sua_senha
```

O `.env` está protegido pelo `.gitignore` e não deve ser enviado ao repositório. O script não imprime a senha no terminal nem a grava nos relatórios.

## Como executar

Com o ambiente virtual ativo, execute:

```powershell
python script.py
```

Durante a execução:

- mantenha a sessão do Windows desbloqueada;
- não minimize ou feche o IWOS manualmente;
- evite usar mouse e teclado sobre o IWOS;
- aguarde a conclusão das buscas e da paginação das grades;
- não abra os arquivos Excel enquanto estiverem sendo gravados.

As grades de Faltas e Atestados podem conter milhares de registros. Por isso, a execução completa pode levar vários minutos.

## Estratégia de paginação

As grades do IWOS são virtualizadas: apenas as linhas visíveis existem na árvore de automação a cada instante. Para não perder registros, o script:

- identifica a barra vertical da grade ativa;
- posiciona o `Thumb` próximo ao topo;
- confirma a primeira faixa de linhas;
- clica no botão real `Page Down`;
- aguarda os índices visíveis mudarem;
- acumula os dados sem duplicar linhas;
- encerra somente quando o final real da barra é alcançado.

Essa abordagem é utilizada principalmente nos painéis de Faltas e Atestados, que podem possuir grandes volumes de dados.

## Segurança

- Nunca versione o arquivo `.env`.
- Utilize uma conta IWOS com apenas as permissões necessárias.
- Evite compartilhar arquivos de exportação sem verificar os dados pessoais contidos neles.
- Os relatórios e credenciais ficam exclusivamente no computador local.
- A pasta `exports/` é ignorada pelo Git.

## Solução de problemas

### O atalho do IWOS não foi encontrado

Confirme se o arquivo `..appref-ms` ainda existe no Menu Iniciar. Se o caminho tiver mudado, atualize `ATALHO_IWOS`.

### A tela de login não foi reconhecida

Verifique se a janela possui o título `Login` e se o IWOS terminou de carregar. Os campos esperados são `editUsuario`, `editSenha` e `buttonEntrar`.

### A busca demorou mais que o esperado

O script aguarda o carregamento e a estabilização das grades. Lentidão de rede, servidor ou banco do IWOS pode ampliar o tempo total.

### A automação não clica no controle correto

Garanta que a sessão esteja desbloqueada e que não existam janelas cobrindo o IWOS. Mudanças futuras na interface ou nos identificadores de automação também podem exigir ajustes no script.

### O PowerShell bloqueou a ativação do ambiente virtual

É possível executar diretamente sem ativá-lo:

```powershell
.\.venv\Scripts\python.exe script.py
```

## Melhorias recomendadas

- tornar o caminho do atalho ClickOnce configurável pelo `.env`;
- adicionar logs estruturados por execução;
- criar testes unitários para transformação e exportação dos dados;
- adicionar captura de tela automática quando ocorrer uma falha de interface;
- permitir seleção de empresas e painéis por configuração;
- gerar um resumo final com duração e quantidade de registros por relatório;
- empacotar a automação em um executável Windows.

## Observações

Este projeto automatiza uma aplicação de terceiros pela interface gráfica. Alterações visuais, atualização do DevExpress ou mudanças nos nomes dos controles do IWOS podem exigir manutenção nos seletores.

Os dados exportados podem conter informações pessoais e devem ser tratados de acordo com as políticas internas de acesso, retenção e privacidade da organização.

---

<div align="center">
  <strong>Automação desenvolvida para tornar a extração de relatórios do IWOS mais rápida, consistente e auditável.</strong>
</div>
