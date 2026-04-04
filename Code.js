/**
 * ASSISTENTE EXECUTIVO - GMAIL + AGENDA + GEMINI AI
 * Este script realiza a leitura da agenda do dia e dos e-mails não lidos,
 * gerando um resumo estruturado via IA e enviando alertas para o Google Chat.
 */

const props = PropertiesService.getScriptProperties();

const CONFIG = {
  API_KEY: props.getProperty('API_KEY'),
  WEBHOOK_CHAT: props.getProperty('WEBHOOK_CHAT'),
  EMAIL_DESTINO: props.getProperty('EMAIL_DESTINO'),
  LISTA_VIP: props.getProperty('LISTA_VIP') ? props.getProperty('LISTA_VIP').split(',').map(item => item.trim()) : [],
  DIAS_HISTORICO: props.getProperty('DIAS_HISTORICO') || '3d'
};

/**
 * Função Principal (Orquestradora)
 * @param {Object} e - Objeto de evento passado automaticamente pelos acionadores.
 */
function gerarResumoDiario(e) {
  const agora = new Date();
  
  // Verifica se a execução é MANUAL (e será undefined) ou via ACIONADOR (e terá conteúdo)
  const isExecucaoManual = (typeof e === 'undefined');

  // Bloqueia o fim de semana apenas se NÃO for execução manual
  if (!isExecucaoManual && isFinalDeSemana(agora)) {
    Logger.log('Execução via acionador detectada no fim de semana. Abortando para respeitar o descanso.');
    return;
  }

  if (isExecucaoManual) {
    Logger.log('Execução MANUAL detectada. Ignorando trava de fim de semana para testes...');
  } else {
    Logger.log('Execução via ACIONADOR detectada. Iniciando processamento normal...');
  }

  const textoAgenda = getBriefingAgenda(agora);
  const dadosEmails = getDadosEmails(CONFIG.DIAS_HISTORICO);

  if (!textoAgenda && !dadosEmails.textoParaResumir) {
    Logger.log('Nenhuma atividade encontrada.');
    return;
  }

  const prompt = montarPrompt(textoAgenda, dadosEmails.textoParaResumir);
  const resumoHTML = chamarGeminiAPI(prompt);

  if (resumoHTML) {
    const htmlLimpo = limparFormatacaoMarkdown(resumoHTML);
    enviarEmail(htmlLimpo, agora);
    
    if (dadosEmails.vipsEncontrados > 0) {
      notificarVIPNoChat(dadosEmails.vipsEncontrados, dadosEmails.resumoVipChat);
    }
  }
}

function getBriefingAgenda(data) {
  const eventos = CalendarApp.getEventsForDay(data);
  if (eventos.length === 0) return "Nenhum compromisso agendado.";

  return eventos.map(evento => {
    const titulo = evento.getTitle();
    const convidados = evento.getGuestList().map(g => g.getName() || g.getEmail()).join(", ");
    const strConvidados = convidados ? ` | Participantes: ${convidados}` : "";
    const inicio = Utilities.formatDate(evento.getStartTime(), Session.getScriptTimeZone(), "HH:mm");
    const fim = Utilities.formatDate(evento.getEndTime(), Session.getScriptTimeZone(), "HH:mm");
    
    return evento.isAllDayEvent() ? 
      `- O dia todo: ${titulo}${strConvidados}` : 
      `- ${inicio} as ${fim}: ${titulo}${strConvidados}`;
  }).join("\n");
}

function getDadosEmails(periodo) {
  const threads = GmailApp.search(`is:unread in:inbox newer_than:${periodo}`);
  let dados = { textoParaResumir: "", vipsEncontrados: 0, resumoVipChat: "" };

  threads.forEach(thread => {
    const msg = thread.getMessages().pop();
    const remetente = msg.getFrom();
    const assunto = msg.getSubject();
    const corpo = msg.getPlainBody().substring(0, 3500);

    const ehVip = CONFIG.LISTA_VIP.some(vip => remetente.toLowerCase().includes(vip.toLowerCase()));
    
    if (ehVip) {
      dados.vipsEncontrados++;
      dados.resumoVipChat += `- De: ${remetente}\n- Assunto: ${assunto}\n\n`;
    }

    const tagVip = ehVip ? " [ATENCAO: REMETENTE VIP] " : "";
    dados.textoParaResumir += `De:${tagVip} ${remetente}\nAssunto: ${assunto}\nMensagem: ${corpo}\n\n---\n\n`;
  });

  return dados;
}

function montarPrompt(agenda, emails) {
  return `Voce e um assessor executivo de alta performance do Vinicius Oliveira. 
  Analise os dados abaixo com precisao.

  REGRAS CRITICAS:
  1. ZERO ALUCINACAO: Baseie-se apenas nos dados fornecidos.
  2. ZERO EMOJIS: Use formatacao profissional.
  3. NOMES COMPLETOS: Use nome e sobrenome dos envolvidos para evitar ambiguidades.

  ESTRUTURA HTML:
  - <h2> BRIEFING DA AGENDA DE HOJE
  - <h2> PRIORIDADES VIP (Use tag <span style="color:red; font-weight:bold;">[URGENTE]</span> se necessario)
  - <h2> OUTROS ASSUNTOS E PROJETOS (Agrupe por tema, coloque <span style="color:red; font-weight:bold;">[URGENTE]</span> na mesma linha do h3 se critico)

  DADOS:
  AGENDA:\n${agenda}
  EMAILS:\n${emails}`;
}

function chamarGeminiAPI(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.API_KEY}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    return json.candidates[0].content.parts[0].text;
  } catch (e) {
    Logger.log(`Erro API Gemini: ${e}`);
    return null;
  }
}

function enviarEmail(html, data) {
  const hora = data.getHours();
  const periodo = hora < 12 ? "Manha" : (hora < 18 ? "Tarde" : "Noite");
  const dataFmt = Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  GmailApp.sendEmail(CONFIG.EMAIL_DESTINO, `Resumo Executivo - ${periodo} (${dataFmt})`, "", { htmlBody: html });
}

function notificarVIPNoChat(qtd, resumo) {
  if (!CONFIG.WEBHOOK_CHAT) return;
  const payload = { text: `*ATENCAO: ${qtd} E-MAIL(S) VIP IDENTIFICADO(S)*\n\n${resumo}Verifique o e-mail para detalhes.` };
  UrlFetchApp.fetch(CONFIG.WEBHOOK_CHAT, {
    method: "post", contentType: "application/json", payload: JSON.stringify(payload)
  });
}

function isFinalDeSemana(data) {
  const dia = data.getDay();
  return (dia === 0 || dia === 6);
}

function limparFormatacaoMarkdown(texto) {
  return texto.replace(/```html/g, "").replace(/```/g, "").trim();
}
