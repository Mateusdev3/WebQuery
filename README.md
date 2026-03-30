# Documento de Contexto — WebQuery

<img width="1112" height="628" alt="image" src="https://github.com/user-attachments/assets/d663598e-9691-4246-bb34-be1ba483ce08" />


Última atualização: 2026-03-30

Este documento reúne o contexto técnico e operacional do projeto `WebQuery` para facilitar manutenção, evolução e integração. Contém descrição da arquitetura, comportamento esperado, esquema de configuração, fluxos principais, dependências, pontos de atenção e propostas de melhoria.

---

## 1. Visão geral

`WebQuery` é uma aplicação de console (.NET 10, C# 14) para executar requisições HTTP (GET / POST) de forma programática, usando entradas de planilhas Excel como parâmetros, agregando e salvando resultados (JSON agregados em planilhas ou arquivos binários decodificados a partir de Base64). O projeto usa `ClosedXML` para leitura/escrita de arquivos Excel.

Propósito: automatizar buscas em APIs a partir de uma lista de entradas (por exemplo, códigos/IDs) e gerar artefatos de saída (planilhas, arquivos KML, imagens, etc.).

---

## 2. Arquitetura e principais componentes

Arquivos relevantes
- `Program.cs` — único arquivo com toda a lógica da aplicação (menu, I/O, chamadas HTTP, manipulação de Excel, persistência em `config.json`).
- `config.json` — arquivo JSON na raiz usado para armazenar o token de autenticação e as queries (buscas) salvas.
- `Resultados/` — pasta padrão (criada automaticamente) onde saem os arquivos gerados.
- `README.md` — documentação básica do projeto.

Principais classes (em `Program.cs`)
- `JsonConfig` — modelo de configuração com propriedades `Queries[]` e `Tokenid`.
- `Query` — modelo para buscas salvas: `Name`, `Url`, `Method`, `Body`, `Sheet`, `ReturnType`.
- `Licences` e `User` — modelos usados na verificação de licença remota (conteúdo esperado em `licences.json`).

Recursos e dependências
- .NET 10 (target framework)
- C# 14
- NuGet: `ClosedXML` para manipulação de Excel

---

## 3. Fluxo de execução

1. Ao iniciar, cria `config.json` se não existir, com `Tokenid` vazio.
2. Cria pasta `Resultados` por padrão (nome pode ser alterado nas configurações do app e será criada quando necessário).
3. Recupera o IP local da máquina e valida licença consultando um JSON remoto hospedado no GitHub (URL: `https://raw.githubusercontent.com/Mateusdev3/WebQuery/refs/heads/main/licences.json`). Se a licença correspondente ao IP não for encontrada, o programa encerra.
4. Exibe menu interativo com opções para executar buscas salvas, executar novas buscas, gerenciar configurações e tokens, e gerenciar buscas salvas (CRUD básico).
5. Execução de uma busca:
   - Se um arquivo `Sheet` for informado, o programa abre a planilha (primeira worksheet) e para cada linha usa a 1ª coluna para substituir espaços no `Body` e chamar a API.
   - A requisição pode ser `GET` (concatena `/{values}` ao `url`) ou `POST` (envia `Body` como JSON no corpo).
   - Recebe a resposta e identifica dois cenários:
     - Resposta JSON contendo um array (inferido por `response.Contains("[")`) — o código deserializa para `List<Dictionary<string,object>>`, agrega colunas e salva em uma planilha (`Resultado.{format}`) dentro da pasta de resultados.
     - Resposta base64 (ou binária envolvida em string) — o código limpa a string (`ClearResponse64`), converte de Base64 e grava em arquivo com nome derivado da entrada (`{data}.{format}`).
6. As buscas salvas são mantidas em `config.json` e podem ser adicionadas, listadas ou removidas via menu. Cada `Query` armazena também o `ReturnType` para definir a extensão do arquivo de saída.

---

## 4. Esquema de configuração (`config.json`)

Estrutura esperada

{
  "Tokenid": "",
  "Queries": [
    {
      "Name": "exemplo",
      "Url": "https://api.exemplo/endpoint",
      "Method": "POST",
      "Body": "{\"param\": \"valor\"}",
      "Sheet": "entrada.xlsx",
      "ReturnType": "kml"
    }
  ]
}

Observações:
- `Tokenid` é usado como Bearer token em todas as requisições.
- `Queries` é um array opcional; o programa cria `config.json` com `Tokenid` vazio caso não exista.

---

## 5. Exemplo de uso

1. Adicionar token: menu `Atualizar Token` > Inserir novo token (validação simples de tamanho 36 caracteres).
2. Salvar nova busca: menu `Gerenciar buscas salvas` > `Salvar nova busca`.
3. Executar busca salva: menu `Realizar busca salva` > escolher índice.
4. Executar busca direta: menu `Realizar nova busca` > inserir URL, método, body, sheet (opcional), tipo de retorno.

Outputs:
- Planilha com resultados agregados: `Resultados/Resultado.{ReturnType}` (cuando o retorno é JSON array e o formato é definido).
- Arquivos individuais (por linha) quando API retorna Base64: `Resultados/{input}.{ReturnType}`.

---

## 6. Formato do JSON de licença esperado

O repositório remoto deve expor um JSON com esta estrutura:

{
  "Licences": [
    { "Ip": "192.168.0.1", "Name": "Nome do Usuário" }
  ]
}

O programa compara o IP da máquina com os campos `Ip` para identificar o nome do usuário/autorização.

---


## 8. Segurança e privacidade

- `Tokenid` é armazenado em `config.json` sem criptografia. Evitar comitá-lo em repositórios remotos.
- O projeto busca um JSON de licenças em um repositório público; garanta que esse arquivo não exponha informações sensíveis.
- As requisições usam Bearer token em headers; valide e proteja o token conforme políticas de TI.

---

## 9. Testes e depuração

- Não existem testes automatizados no repositório; criar testes unitários para as funções críticas (parsing, ClearResponse64, Init/Read config, GetIp) é recomendado.
- Para depuração local: executar no Visual Studio com ponto de interrupção em `Main` e inspecionar `config.json` e respostas HTTP.

---

## 10. Melhoria sugeridas (prioritárias)

1. Melhorar parsing de respostas (JSON vs Base64) usando tentativa de desserialização com `JsonDocument`/`JsonSerializer`.
2. Padronizar leitura de entradas do usuário (usar sempre `Console.ReadLine` ou um helper) para evitar inconsistências.
3. Melhor tratamento de erros e logs (ex.: adicionar logger, incluir mensagens de erro detalhadas no console e arquivos).
4. Não armazenar tokens em texto puro ou incluir opção de criptografia/KeyVault.
5. Refatorar `Program.cs` em módulos menores (serviço HTTP, serviço de I/O, serviço de licenças, modelos) para facilitar manutenção e testes.
6. Adicionar testes unitários.

---

## 11. Mapa de arquivos e responsabilidades

- `Program.cs` — lógica principal e menus.
- `config.json` — persistência local (token + queries).
- `Resultados/` — saída de arquivos.
- `README.md` — documentação de projeto (visão geral).
- `CONTEXT.md` — este documento de contexto.

---


---


