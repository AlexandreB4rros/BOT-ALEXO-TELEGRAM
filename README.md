# BOT TELEGRAM - Alexo

## Introdução

O BOT TELEGRAM Alexo é uma solução automatizada para integração de operações de rede, manipulação de arquivos (KML, KMZ, XLSX), consulta e atualização de dados via Google Apps Script e interação direta com grupos do Telegram. Ele foi desenvolvido para facilitar o gerenciamento de templates, CTOs, POPs e atividades relacionadas a redes ópticas, além de oferecer comandos administrativos e de consulta para equipes técnicas.

---
## Resumo das Funções do Código

### Utilitários e Manipulação de Arquivos

- **ExcluirArquivos(caminho_arquivo):**  
  Exclui todos os arquivos que possuem o mesmo nome base e extensão no diretório informado.

- **ExcluirArquivosporExtensao():**  
  Exclui todos os arquivos com extensão `.xlsx`, `.kml` ou `.kmz` no diretório atual.

- **kml_to_xlsx(kml_file, xlsx_file):**  
  Converte um arquivo KML em uma planilha XLSX, extraindo os placemarks e suas coordenadas.

- **extract_kml_from_kmz(kmz_file, extract_to):**  
  Extrai o arquivo KML de dentro de um arquivo KMZ e renomeia para facilitar o uso.

- **encontrar_arquivo_kml_kmz(DirArquivo):**  
  Procura e retorna o caminho de um arquivo `.kml` ou `.kmz` dentro de um diretório.

---

### Manipulação de Dados e Configuração

- **ListaCidades():**  
  Retorna uma lista formatada de cidades e POPs cadastrados no arquivo `WebHook.json`.

- **buscar_webhook_por_pop(pop):**  
  Busca e retorna o link do webhook associado ao POP informado.

- **buscar_cidade_por_pop(pop):**  
  Busca e retorna o nome da cidade associada ao POP informado.

- **buscar_dir_drive():**  
  Retorna o diretório do drive salvo no arquivo `config_drive.json`.

---

### Integração com Telegram e Google Apps Script

- **send_log_to_telegram(message):**  
  Envia mensagens de log diretamente para um grupo do Telegram.

- **TelegramHandler:**  
  Handler de logging personalizado para enviar logs ao Telegram.

- **fetch_data(webhook_link, payload):**  
  Realiza uma requisição HTTP POST assíncrona para o webhook do Google Apps Script e retorna a resposta.

---

### Comandos do Bot (Handlers)

- **ajuda(update, context):**  
  Envia uma mensagem com a lista de comandos disponíveis e informações do bot.

- **AjudaAdm(update, context):**  
  Envia uma mensagem com comandos administrativos e links úteis.

- **Info(update, context):**  
  Envia informações sobre o bot, versão, criador e créditos.

- **id(update, context):**  
  Exibe o ID do grupo, nome do grupo e ID do usuário.

- **CWH(update, context):**  
  Envia o arquivo `WebHook.json` para o chat.

- **ExibirCidade(update, context):**  
  Exibe a lista de cidades cadastradas.

- **AdcionarTemplate(update, context):**  
  Adiciona um novo template (cidade, POP e webhook) ao arquivo `WebHook.json`.

- **ExcluirTemplate(update, context):**  
  Remove um POP/cidade do arquivo `WebHook.json`.

- **atividades(update, context):**  
  Consulta atividades pendentes para um POP via webhook.

- **checar(update, context):**  
  Consulta informações de OLT/SLOT/PON de uma CTO via webhook.

- **localizar_cto(update, context):**  
  Retorna a localização de uma CTO via webhook.

- **input(update, context):**  
  Inputa informações de splitter para uma CTO no template via webhook.

- **insert(update, context):**  
  Inputa informações de CTO e OLT/SLOT/PON no template via webhook.

- **novaCTO(update, context):**  
  Inicia o processo para adicionar uma nova CTO, solicitando a localização do usuário.

- **handle_location(update, context):**  
  Recebe a localização enviada pelo usuário e finaliza o cadastro da nova CTO.

- **listarIDs(update, context):**  
  Lista os IDs disponíveis para um determinado POP e OLT/SLOT/PON.

- **convert(update, context):**  
  Solicita ao usuário o envio de um arquivo KML/KMZ para conversão.

- **handle_arquivo(update, context):**  
  Processa o arquivo KML/KMZ enviado, converte para XLSX e oferece opções ao usuário.

- **configdrive(update, context):**  
  Salva o diretório do drive informado pelo usuário no arquivo de configuração.

- **baixarkmz(update, context):**  
  Envia para o usuário o arquivo KML/KMZ encontrado no diretório do drive da cidade/POP informado.

- **gerarkmzatualizado(update, context):**  
  Gera um novo arquivo KML base a partir do template da cidade/POP informado.

- **handle_mensagem(update, context):**  
  Gerencia o fluxo de mensagens do usuário, especialmente durante operações de conversão e input de dados.

---

### Auxiliares para Planilhas

- **converter_planilha_template_para_kml(...):**  
  Gera um arquivo KML a partir de uma planilha XLSX, usando os dados da aba KMZ.

- **DE_KMZ_BASE_PARA_TEMPLATE(arquivo_origem, arquivo_destino):**  
  Copia dados das colunas A, B e C de uma planilha de origem para a aba "KMZ" de uma planilha de destino.

- **VerificarTemplatemporPOP(DirTemplate, PopInformado_user, update):**  
  Verifica se existe um template para o POP informado no diretório do drive.

---



