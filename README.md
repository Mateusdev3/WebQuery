# WebQuery


<img width="1112" height="628" alt="image" src="https://github.com/user-attachments/assets/bf969eaa-ca19-4920-9aa6-aae9c4486fa5" />


Aplicação de console para executar queries web em lote a partir de uma planilha Excel e salvar resultados em diversos formatos de arquivos diferentes
## Descrição

`WebQuery` é uma ferramenta de linha de comando que: 
- Executa requisições HTTP (GET/POST) para APIs com base em entradas de uma planilha Excel.
- Agrega as respostas das requisições em arquivos xlsx, xml, kml, json, txt dentre outros.
- Suporta salvar, listar e excluir buscas pré-configuradas via `config.json`.
- Usa um token Bearer para autenticação nas requisições.

## Requisitos

- .NET 10
- C# 14
- Pacote NuGet: `ClosedXML` (para leitura/escrita de Excel)

## Estrutura de configuração (`config.json`)

O `config.json` contém a configuração da aplicação e queries salvas. Exemplo mínimo gerado automaticamente:

```json
{
  "Tokenid": "",
  "Queries": []
}
```

Propriedades importantes:
- `Tokenid` (string): token Bearer usado nas requisições.
- `Queries` (array): lista de objetos `Query` com as propriedades:
  - `Name` (string): nome da query.
  - `Url` (string): endpoint da API.
  - `Method` (string): `GET` ou `POST`.
  - `Body` (string | nullable): corpo JSON para `POST` (ou parte do caminho para `GET`).
  - `Sheet` (string | nullable): caminho/arquivo da planilha fonte (se aplicável).

## Uso

Ao abrir a aplicação segue um menu com opções:
- `[1] Realizar query salva` — executa uma query gravada.
- `[2] Realizar nova query` — solicita URL/método/body e, opcionalmente, uma planilha fonte.
- `[3] Configurações` — opções simples de configuração (diretório, delay) (placeholder).
- `[4] Gerenciar queries salvas` — salvar, listar ou excluir queries.
- `[5] Atualizar Token` — inserir e testar token no `config.json`.

### Formato da planilha fonte

A aplicação lê a primeira coluna da primeira planilha (`Worksheet(1)`) e para cada linha substitui espaços no `body` com o valor da célula (comportamento atual). Os resultados são agregados em `resultado.xlsx`.

## Observações e limitações

- A aplicação depende de acesso à internet para validar licenças e executar requisições.
- O link de validação de licença atualmente consulta um JSON remoto no GitHub.
- Tratamento de erros, internacionalização e validações são básicos — revisar antes de uso em produção.
- Timeout de requisições: 10 segundos (configurado em `HttpClient`).

## Contribuição

Pull requests são bem-vindos. Para mudanças grandes, abrir uma issue antes para alinhar a implementação.

## Licença

Sem licença explícita no repositório. Adicione um arquivo `LICENSE` se desejar aplicar uma licença específica.

