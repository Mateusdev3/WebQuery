
# WebQuery

<img width="1112" height="628" alt="image" src="https://github.com/user-attachments/assets/d663598e-9691-4246-bb34-be1ba483ce08" />

Aplicação de console para execução em lote de buscas (queries) contra APIs usando entradas de planilhas Excel e geração de resultados em arquivos (planilhas ou binários).

Resumo rápido
- Plataforma: .NET 10 (C# 14)
- Propósito: automatizar longas buscas manuais repetitivas e demoradas com parâmetros vindos de planilha e salvar respostas como arquivos ou em planilhas.
- Dependência principal: `ClosedXML` (leitura/escrita de Excel).

-----

Índice
- Visão geral
- Recursos
- Requisitos
- Instalação e build
- Uso (fluxo e exemplos)
- Formato de `config.json`
- Estrutura de pastas e arquivos gerados
- Erros comuns e troubleshooting

-----

Visão geral
------------
`WebQuery` permite definir buscas (queries) que chamam endpoints HTTP em massa usando um Bearer token configurado em `config.json`. As buscas podem ser:
- baseadas em uma planilha de entrada: para cada linha a aplicação realiza a chamada e processa a resposta;
- ou chamadas únicas diretas inserindo URL/método/body no fluxo interativo.

As respostas tratadas atualmente são de dois tipos principais:
- JSON array (agregado para planilha de saída); ou
- conteúdo Base64 (decodificado e salvo como arquivo usando a extensão informada).

Recursos
--------
- Menu interativo via console para gerenciar e executar buscas.
- CRUD básico de buscas salvas persistidas em `config.json`.
- Leitura de planilhas Excel (.xlsx) usando `ClosedXML`.
- Escrita de arquivos de resultado em pasta configurável (padrão `Resultados`).
- Validação de licença simples via JSON remoto (comparação de IP da máquina).

Requisitos
----------
- .NET 10 SDK
- C# 14
- Pacotes NuGet (instalados automaticamente com `dotnet restore`):
  - `ClosedXML`

Instalação e build
------------------
1. Clone o repositório:

```powershell
git clone https://github.com/Mateusdev3/WebQuery.git
cd WebQuery
```

2. Restaurar dependências e compilar:

```powershell
dotnet restore
dotnet build -c Release
```

3. Executar:

```powershell
dotnet run --project WebQuery
```

Ao executar pela primeira vez o programa criará um `config.json` com campo `Tokenid` vazio e a pasta `Resultados` (padrão).

Uso (fluxo e exemplos)
----------------------
Ao iniciar, o programa:
1. Valida licença consultando um JSON remoto com mapeamento de IPs para nomes de usuário;
2. Exibe menu principal com opções para executar buscas salvas, criar uma nova busca, gerenciar buscas e atualizar token.

Menu principal (resumo):
- `[1] Realizar busca salva` — escolhe uma busca salva por índice e a executa;
- `[2] Realizar nova busca` — solicita URL, método (GET/POST), body, planilha de entrada (opcional) e tipo de retorno (extensão); executa imediatamente;
- `[3] Configurações` — editar nome da pasta de resultados, inserir delay entre requisições (placeholder);
- `[4] Gerenciar buscas salvas` — salvar nova busca, listar ou excluir buscas;
- `[5] Atualizar Token` — inserir novo token e testar validade.

Exemplo prático
1. Adicionar token: selecione `Atualizar Token` e insira o token (cuidado para não versionar `config.json`).
2. Salvar nova busca: `Gerenciar buscas salvas` → `Salvar nova busca` → preencha `Name`, `Url`, `Method`, `Body`, `Sheet` (opcional), `ReturnType` (ex.: `kml`, `png`, `xlsx`).
3. Executar busca salva: `Realizar busca salva` → escolha índice.

Quando a busca utiliza uma `Sheet`:
- a aplicação abre a primeira planilha do arquivo e usa o valor da primeira coluna de cada linha substituindo espaços no `Body` (comportamento atual) para formar a requisição;
- se a resposta for um JSON array, os objetos são agregados e exportados para uma planilha `Resultado.{ReturnType}`;
- se a resposta for conteúdo em Base64, cada resultado é decodificado e salvo como `{input}.{ReturnType}` na pasta de resultados.

Formato de `config.json`
----------------------
Exemplo de `config.json` válido:

```json
{
  "Tokenid": "",
  "Queries": [
    {
      "Name": "BuscarKML",
      "Url": "https://api.exemplo/endpoint",
      "Method": "POST",
      "Body": "{\"codigo\": \"\"}",
      "Sheet": "entrada.xlsx",
      "ReturnType": "kml"
    }
  ]
}
```

Campos importantes:
- `Tokenid`: token Bearer usado nas requisições;
- `Queries`: array de objetos com `Name`, `Url`, `Method`, `Body`, `Sheet` e `ReturnType`.

Estrutura de pastas e arquivos gerados
------------------------------------
- `config.json` — arquivo de configuração e armazenamento de buscas salvas.
- `Resultados/` (padrão) — pasta onde os arquivos gerados são salvos. Pode ser alterada nas configurações do app.
- Exemplo de saída:
  - `Resultados/Resultado.kml` (quando respostas JSON são agregadas e `ReturnType` é `kml`);
  - `Resultados/123.kml` (quando resultado por linha é Base64 convertido em arquivo `kml`).

Erros comuns e troubleshooting
------------------------------
- Programa encerra por licença não encontrada
- Token inválido: vá em `Atualizar Token` e use `Testar token salvo` para checar validade.
- Arquivo de planilha não encontrado: informe o caminho correto em `Sheet` ou coloque a planilha no mesmo diretório do executável.



