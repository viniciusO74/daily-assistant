# Assistente Executivo com IA (Google Apps Script + Gemini)

Este projeto automatiza a criação de briefings executivos diários, integrando o Google Agenda e o Gmail com a inteligência artificial do Gemini (Google).

## 🚀 Funcionalidades
- **Resumo de Agenda:** Lista compromissos do dia com horários e participantes.
- **Resumo de E-mails:** Analisa e-mails não lidos, agrupando por projeto/tema.
- **Prioridade VIP:** Identifica remetentes importantes e os destaca no topo.
- **Alertas via Google Chat:** Notificações em tempo real no celular via Webhook.
- **Formatação Profissional:** Saída em HTML limpo, sem emojis, focada em produtividade.
- **Execução Inteligente:** O script detecta automaticamente finais de semana e não envia e-mails, mas permite bypass se executado manualmente via editor para facilitar testes.

## 🛠️ Configuração Segura
O código utiliza o `PropertiesService` do Google para garantir que chaves de API e dados sensíveis não fiquem expostos no código-fonte. Você precisará adicionar as variáveis abaixo nas **Propriedades do Script** (Configurações do Projeto ⚙️).

### Variáveis Necessárias:
1. `API_KEY`: Sua chave de API do Google AI Studio (Gemini).
2. `WEBHOOK_CHAT`: URL do Webhook do seu espaço no Google Chat.
3. `EMAIL_DESTINO`: Endereço de e-mail que receberá os resumos.
4. `LISTA_VIP`: Lista de domínios ou e-mails separados por vírgula (ex: `@empresa.com, chefe@gmail.com`).
5. `DIAS_HISTORICO`: Período de busca de e-mails para processamento (ex: `3d`).

## 📅 Instalação
1. Crie um novo projeto no [Google Apps Script](https://script.google.com/).
2. Copie o arquivo `Código.gs` deste repositório e cole no editor.
3. Cadastre as variáveis acima nas Propriedades do Script.
4. Crie os acionadores (Triggers) para os horários que deseja receber o briefing (ex: 07:30, 12:45, 18:30).

## 📄 Licença
Este projeto está sob a licença MIT. Sinta-se à vontade para clonar, compartilhar e adaptar ao seu fluxo de trabalho.
