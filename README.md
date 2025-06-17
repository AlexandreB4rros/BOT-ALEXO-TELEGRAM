# Créditos

Joseph A.
Alexandre B. 

## Colaboração

Josimar
Anderson 

# BOT-ALEXO-TELEGRAM

## Visão Geral

Este projeto implementa um bot do Telegram chamado **Alexo**, voltado para automação de rotinas técnicas, administrativas e operacionais, especialmente para equipes que trabalham com redes ópticas (CTO, POP, OLT/SLOT/PON, etc). O bot permite manipulação de arquivos, integração com planilhas, localização geográfica, notificações administrativas e gerenciamento de templates via comandos no chat.

## Estrutura do Projeto

- **BOT_TELEGRAM.py**: Arquivo principal do bot, contendo toda a lógica de comandos, handlers, integração com arquivos, banco de dados e notificações.
- **admins_fallback.json**: Backup da lista de administradores.
- **WebHook.json**: Configuração de webhooks para integração com planilhas/templates.
- **arquv/**: Pasta com arquivos de templates e planilhas de exemplo.
- **Scripts_Alexo/**: Módulos auxiliares, como seleção de token e versionamento.
- **build/**: Arquivos gerados pelo processo de build (PyInstaller).

## Principais Funcionalidades

### 1. Inicialização e Configuração

- Carrega variáveis de ambiente do arquivo `.env`.
- Seleciona o token do bot conforme o modo de debug.
- Configura logging detalhado, inclusive com envio de logs para o Telegram.

### 2. Comandos do Bot

- `/start` ou `/ajuda`: Exibe todos os comandos disponíveis e suas descrições.
- `/atividades <POP>`: Verifica atividades pendentes em um template.
- `/checar <CTO> <FSAN>`: Consulta OLT/SLOT/PON de um cliente.
- `/localizar <CTO>`: Retorna a localização geográfica de uma CTO.
- `/ctos`: Mostra CTOs próximas à localização enviada.
- `/listarIDs <POP> <OLT/SLOT/PON>`: Lista IDs de CTOs disponíveis em uma PON.
- `/input <CTO> <SPLITTER>`: Inputa data e splitter no template.
- `/insert <CTO> <OLT/SLOT/PON>`: Inputa OLT/SLOT/PON na aba 'checar' do template.
- `/novaCTO <POP> <OLT/SLOT/PON> <SPLITTER>`: Adiciona uma nova CTO.
- `/convert`: Converte arquivos KML/KMZ em XLSX.
- `/baixarkmz <POP>`: Baixa arquivos KMZ/KML do drive.
- `/gerarkmzatualizado <POP>`: Gera arquivo KML base a partir do template.
- `/id`: Mostra o ID do usuário e do chat.
- `/info`: Exibe informações do bot, versão e créditos.

### 3. Comandos Administrativos

- `/cadastrar <CARGO>`: Gera link de convite para novo usuário.
- `/exibircidades`: Lista cidades e POPs configurados.
- `/AdicionarTemplate <CIDADE> <POP> <WEBHOOK>`: Adiciona novo template ao WebHook.json.
- `/ExcluirTemplate <POP>`: Remove template do WebHook.json.
- `/configdrive <CAMINHO>`: Define diretório raiz do drive local.
- `/CWH`: Envia o arquivo WebHook.json.
- `/listar_admins`: Exibe lista de administradores.
- `/AjudaAdm`: Lista comandos administrativos.

### 4. Manipulação de Arquivos

- Conversão entre formatos KML/KMZ e XLSX usando `openpyxl` e `simplekml`.
- Geração de mapas com localização de CTOs usando `matplotlib` e `contextily`.
- Manipulação assíncrona de arquivos com `aiofiles`.

### 5. Integração com Banco de Dados

- Conexão assíncrona com banco de dados MySQL via `aiomysql` para autenticação, cadastro e fallback de administradores.

### 6. Sistema de Permissões

- Decorator `@check_permission` para restringir comandos a usuários autorizados.

### 7. Agendamento de Tarefas

- Atualização diária da lista de administradores via `JobQueue` do `python-telegram-bot`.

### 8. Tratamento de Erros

- Handlers centralizados para erros, com notificações para administradores e logs detalhados.

## Fluxo Principal

1. O bot é inicializado e configura todos os handlers de comandos e mensagens.
2. Usuários interagem via comandos no chat do Telegram.
3. O bot executa operações de leitura/escrita em arquivos, manipulação de planilhas, localização geográfica e integração com banco de dados conforme o comando.
4. Logs e notificações são enviados para administradores em caso de erro ou eventos importantes.

## Tecnologias Utilizadas

- Python 3.11+
- [python-telegram-bot](https://python-telegram-bot.org/)
- openpyxl, pandas, matplotlib, contextily, simplekml, aiomysql, aiofiles, dotenv

## Como Executar

1. Configure o arquivo `.env` com as variáveis necessárias.
2. Instale as dependências com `pip install -r requirements.txt`.
3. Execute o bot com `python BOT_TELEGRAM.py`.

---

Para detalhes de cada comando, utilize `/ajuda` ou `/AjudaAdm` no próprio bot.

---