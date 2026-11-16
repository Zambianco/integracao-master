# Consulta de Batidas - Flask

Aplicacao Flask simples para consultar o AFD do equipamento Control iD e exibir as batidas em uma pagina HTML com Bootstrap.

## Requisitos

- Python com `flask`, `requests` e `urllib3`

## Como executar

```powershell
cd "Consulta de Batidas - Flask"
set CONTROLID_BASE_URL=https://192.168.0.204
set CONTROLID_USERNAME=admin
set CONTROLID_PASSWORD=troque-esta-senha
python app.py
```

Abra no navegador:

```text
http://127.0.0.1:5000
```

## Observacoes

- A consulta ao equipamento e feita no backend Flask, evitando CORS e problemas com o certificado HTTPS do relogio.
- A tela envia apenas filtros de data e NSR. URL, usuario e senha ficam exclusivamente no backend via variaveis de ambiente.
- O retorno mostra a tabela das marcacoes e o AFD bruto.
- A busca por `NSR inicial` tenta `POST /export_afd.fcgi` e, em firmwares REP que nao suportam esse comando, faz fallback para `RepCid.BuscarAFD/LerAFD`.
- As consultas persistem os registros em um unico SQLite local `afd_history.sqlite3`, separado em tabelas de dispositivos, importacoes, registros e conflitos.
- A consulta por data responde a partir do SQLite local depois de sincronizar uma janela de validacao de `-1 dia / +1 dia` no equipamento.

## Configuracao do equipamento

Defina estas variaveis no ambiente do processo Flask:

- `CONTROLID_BASE_URL`
- `CONTROLID_USERNAME`
- `CONTROLID_PASSWORD`

Se alguma delas nao estiver definida, o endpoint `/api/afd` retorna erro de configuracao no servidor.

## Busca por NSR

O modo `Por NSR inicial (API)` faz login no equipamento e tenta primeiro o endpoint:

```text
POST /export_afd.fcgi?session=<session>&mode=595
```

Com body JSON:

```json
{
  "initial_nsr": 12345
}
```

Se o equipamento responder `Invalid command: export_afd`, o backend usa o caminho legado da SDK:

```text
RepCid.BuscarAFD(nsr)
RepCid.LerAFD()
```

Observacoes:

- O parser da tela esta alinhado ao layout AFD 595/1510.
- Em firmwares REP mais antigos, o filtro por `NSR` ainda depende de `RepCid.dll` registrada no Windows e de suporte COM no Python.
- O historico local e segmentado por URL do equipamento e guarda os registros por `NSR`.
- A busca por data usa o banco como fonte principal e marca na tabela o que veio do historico e o que entrou direto na sincronizacao atual.

## Painel de Inconsistencias (`/painel`)

Mostra, empregado por empregado, as ocorrencias de um periodo ja encerrado (padrao: ontem). O dia atual nunca e analisado: `GET /api/painel?start=aaaa-mm-dd&end=aaaa-mm-dd` recusa hoje/futuro e periodos acima de 62 dias.

- **Sem batidas**: nenhuma marcacao no dia.
- **Batidas impares**: quantidade impar de marcacoes.
- **Intervalo intrajornada curto**: diferenca entre a 2a e a 3a batida menor que 60 minutos (`MIN_INTRAJORNADA_MINUTES`), avaliada somente com 4 ou mais batidas.
- Batidas repetidas no mesmo minuto contam uma vez so.

Antes de analisar, o periodo e sincronizado com o relogio (dias ainda nao fechados no SQLite) e a lista de usuarios do relogio (`load_users.fcgi`) e salva na tabela `device_users` para exibir os nomes. Se o relogio estiver indisponivel, o painel usa os dados locais e exibe um aviso.

A lista de empregados vem dos registros AFD tipo 5 (cadastro), desconsiderando quem tem exclusao (`E`) posterior a ultima batida, mais quem bateu ponto nos 30 dias anteriores.

**Impressao**: o botao *Imprimir* gera um relatorio A4 no padrao de impressao do PAP (`static/css/print-standard.css`, `static/js/print-standard.js` e logotipo copiados de la), com o que esta filtrado na tela e apenas o nome do empregado (sem PIS).
