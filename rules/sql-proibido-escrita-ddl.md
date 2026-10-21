# Bases de dados: NUNCA executar SQL de escrita ou de estrutura

Este repositório é de QA — o agente só **lê** dados para registar defeitos e
fundamentar relatórios. As bases de dados acessíveis (ex.: Oracle `ACCQUA`
via MCP `sqlcl`, ligações do DBeaver) são ambientes partilhados e o utilizador
de base de dados não é só de leitura, por isso esta regra é a única barreira.

## Comandos proibidos

O agente **NUNCA** executa nenhum destes comandos, em nenhuma base de dados:

- `INSERT INTO`
- `UPDATE`
- `DELETE`
- `TRUNCATE`
- `CREATE DATABASE`, `ALTER DATABASE`, `DROP DATABASE`
- `CREATE SCHEMA`, `ALTER SCHEMA`, `DROP SCHEMA`
- `CREATE TABLE`, `ALTER TABLE`, `DROP TABLE`

## Âmbito

- Vale para **qualquer via de execução**: ferramentas MCP (`mcp__sqlcl__run-sql`,
  `mcp__sqlcl__run-sqlcl`, ou outro servidor MCP de base de dados), `sql.exe` /
  `sqlplus` / `psql` / outro cliente via Bash ou PowerShell, scripts `.sql`
  executados com `@ficheiro`, e scripts em qualquer linguagem que abram uma
  ligação.
- Vale também quando o comando vem **embrulhado**: dentro de um bloco PL/SQL
  anónimo (`BEGIN ... END;`), de `EXECUTE IMMEDIATE`, de um script, ou combinado
  com outras instruções na mesma chamada.
- Vale **mesmo que o utilizador peça** a execução na conversa. Nesse caso, o
  agente escreve o SQL como texto (bloco de código) para o utilizador rever e
  correr manualmente, e diz explicitamente que não o executou.

## O que é permitido

Só consultas de leitura: `SELECT` (incluindo `WITH ... SELECT`), `DESCRIBE`,
`EXPLAIN PLAN` e consultas ao dicionário de dados (`ALL_TABLES`,
`ALL_TAB_COLUMNS`, `USER_*`, `V$*`, etc.).

Antes de executar qualquer SQL, confirmar que nenhuma das instruções da chamada
pertence à lista proibida. Em caso de dúvida, não executar e perguntar.
