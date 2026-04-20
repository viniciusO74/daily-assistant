/**
 * @fileoverview EXECUTIVE ASSISTANT - Gmail, Calendar, Drive (Obsidian) and Gemini AI Integration.
 * @description Automation pipeline for data extraction and executive briefing generation.
 * Implements recursive directory scanning and advanced telemetry logs.
 */

const props = PropertiesService.getScriptProperties();

// Environment variables initialization with explicit conditional logic (Senior Clean Code)
let listaVipBruta = props.getProperty('LISTA_VIP');
let listaVipArray = [];

if (listaVipBruta) {
  listaVipArray = listaVipBruta.split(',').map(item => item.trim()).filter(String);
}

let diasHistorico = props.getProperty('DIAS_HISTORICO');
if (!diasHistorico) {
  diasHistorico = '3d';
}

let pastasObsidian = props.getProperty('PASTAS_OBSIDIAN');
if (!pastasObsidian) {
  pastasObsidian = '';
}

let nomeUsuario = props.getProperty('NOME_USUARIO');
if (!nomeUsuario) {
  nomeUsuario = 'Usuário Executivo';
}

const CONFIG = {
  API_KEY: props.getProperty('API_KEY'),
  WEBHOOK_CHAT: props.getProperty('WEBHOOK_CHAT'),
  EMAIL_DESTINO: props.getProperty('EMAIL_DESTINO'),
  NOME_USUARIO: nomeUsuario,
  LISTA_VIP: listaVipArray,
  DIAS_HISTORICO: diasHistorico,
  PASTAS_OBSIDIAN: pastasObsidian
};

/**
 * Script entrypoint. Orchestrates the collection, processing, and delivery flow.
 */
function gerarResumoDiario(e) {
  Logger.log('[SYSTEM] ==========================================');
  Logger.log('[SYSTEM] Starting Executive Assistant execution');
  
  const agora = new Date();
  const hora = agora.getHours();
  const isExecucaoManual = (typeof e === 'undefined');

  let periodo = "Manhã";
  if (hora >= 12 && hora < 18) {
    periodo = "Tarde";
  } else if (hora >= 18) {
    periodo = "Noite";
  }

  Logger.log(`[ROUTING] Period detected: ${periodo}. Manual execution: ${isExecucaoManual}`);

  if (!isExecucaoManual && isFinalDeSemana(agora)) {
    Logger.log('[ROUTING] Execution aborted: Weekend detected.');
    return;
  }

  let textoAgenda = "";
  let dadosObsidian = "";

  if (periodo === "Manhã") {
    Logger.log('[ROUTING] Starting morning data extraction (Calendar and Obsidian)...');
    textoAgenda = getBriefingAgenda(agora);
    dadosObsidian = getDadosObsidian(CONFIG.PASTAS_OBSIDIAN);
  } else {
    Logger.log('[ROUTING] Skipping Calendar and Obsidian (non-morning period).');
  }
  
  const dadosEmails = getDadosEmails(CONFIG.DIAS_HISTORICO);

  if (!textoAgenda && !dadosEmails.textoParaResumir && !dadosObsidian) {
    Logger.log('[SYSTEM] Pipeline terminated: No data captured in any module.');
    return;
  }

  Logger.log('[PROMPT] Building payload for LLM submission...');
  const prompt = montarPrompt(textoAgenda, dadosEmails, dadosObsidian, periodo);
  
  Logger.log('[GEMINI_API] Starting external request...');
  const resumoHTML = chamarGeminiAPI(prompt);

  if (resumoHTML) {
    Logger.log('[SYSTEM] HTML response received successfully. Starting sanitization and dispatch.');
    const htmlLimpo = limparFormatacaoMarkdown(resumoHTML);
    enviarEmail(htmlLimpo, agora);
    
    if (dadosEmails.vipsEncontrados > 0) {
      notificarVIPNoChat(dadosEmails.vipsEncontrados, dadosEmails.resumoVipChat);
    }
    Logger.log('[SYSTEM] Execution completed successfully.');
    Logger.log('[SYSTEM] ==========================================');
  } else {
    Logger.log('[SYSTEM] General failure: LLM did not return valid HTML.');
  }
}

/**
 * Collects calendar events for today and tomorrow.
 */
function getBriefingAgenda(data) {
  Logger.log('[CALENDAR] Collecting appointments...');
  const formatarEvento = (evento) => {
    const titulo = evento.getTitle();
    const convidados = evento.getGuestList().map(g => g.getName() || g.getEmail()).join(", ");
    
    let strConvidados = "";
    if (convidados) {
      strConvidados = ` | Participantes: ${convidados}`;
    }
    
    const inicio = Utilities.formatDate(evento.getStartTime(), Session.getScriptTimeZone(), "HH:mm");
    const fim = Utilities.formatDate(evento.getEndTime(), Session.getScriptTimeZone(), "HH:mm");
    
    if (evento.isAllDayEvent()) {
      return `- Dia todo: ${titulo}${strConvidados}`;
    } else {
      return `- ${inicio} até ${fim}: ${titulo}${strConvidados}`;
    }
  };

  const eventosHoje = CalendarApp.getEventsForDay(data);
  let textoHoje = "Nenhum compromisso agendado para hoje.";
  if (eventosHoje.length > 0) {
    textoHoje = eventosHoje.map(formatarEvento).join("\n");
  }

  const dataAmanha = new Date(data);
  dataAmanha.setDate(dataAmanha.getDate() + 1);
  const eventosAmanha = CalendarApp.getEventsForDay(dataAmanha);
  
  let textoAmanha = "Nenhum compromisso agendado para amanhã.";
  if (eventosAmanha.length > 0) {
    textoAmanha = eventosAmanha.map(formatarEvento).join("\n");
  }

  Logger.log(`[CALENDAR] Found ${eventosHoje.length} appointments for today and ${eventosAmanha.length} for tomorrow.`);
  return `--- HOJE ---\n${textoHoje}\n\n--- AMANHÃ ---\n${textoAmanha}`;
}

/**
 * Scans Inbox for unread emails and identifies VIP senders based on whitelist.
 */
function getDadosEmails(periodo) {
  Logger.log(`[GMAIL] Searching for unread emails from the last ${periodo}...`);
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

    let tagVip = "";
    if (ehVip) {
      tagVip = " [ATENÇÃO: VIP] ";
    }
    
    dados.textoParaResumir += `Remetente:${tagVip} ${remetente}\nAssunto: ${assunto}\nMensagem: ${corpo}\n\n---\n\n`;
  });

  Logger.log(`[GMAIL] Search ended: ${threads.length} threads processed. ${dados.vipsEncontrados} VIPs identified.`);
  return dados;
}

/**
 * Scans configured folders and starts the recursive extraction process of .md files.
 */
function getDadosObsidian(pastasIdsStr) {
  if (!pastasIdsStr) {
    Logger.log('[OBSIDIAN] No folder configured. Skipping module.');
    return "";
  }
  
  Logger.log('[OBSIDIAN] Starting recursive directory scan...');
  const ids = pastasIdsStr.split(',').map(id => id.trim()).filter(String);
  let conteudoAcumulado = "";

  ids.forEach(id => {
    try {
      const pastaRaiz = DriveApp.getFolderById(id);
      Logger.log(`[OBSIDIAN] Accessing root folder: ${pastaRaiz.getName()}`);
      conteudoAcumulado += extrairDadosRecursivamente(pastaRaiz);
    } catch (e) {
      Logger.log(`[OBSIDIAN_ERROR] Failed to access folder ID: ${id}. Detail: ${e}`);
    }
  });

  return conteudoAcumulado;
}

/**
 * Helper Func: Implements Depth-First Search (DFS) to read files in subfolders.
 */
function extrairDadosRecursivamente(pasta) {
  let resultado = `\n--- CONTEXTO: ${pasta.getName()} ---\n`;
  let arquivosLidos = 0;
  
  const arquivos = pasta.getFiles();
  while (arquivos.hasNext()) {
    const arquivo = arquivos.next();
    const nomeArquivo = arquivo.getName().toLowerCase();
    
    if (nomeArquivo.endsWith('.md')) {
      arquivosLidos++;
      const conteudo = arquivo.getBlob().getDataAsString("UTF-8");
      const limiteTexto = conteudo.substring(0, 2500); 
      resultado += `\n[ARQUIVO: ${arquivo.getName()}]\n${limiteTexto}\n`;
    }
  }
  
  if (arquivosLidos > 0) {
    Logger.log(`[OBSIDIAN] Extracted ${arquivosLidos} MD files from folder: ${pasta.getName()}`);
  }
  
  const subpastas = pasta.getFolders();
  while (subpastas.hasNext()) {
    resultado += extrairDadosRecursivamente(subpastas.next());
  }
  
  return resultado;
}

/**
 * Builds the structured prompt for the LLM.
 * Note: The prompt output text remains in Portuguese to instruct the LLM correctly.
 */
function montarPrompt(agenda, dadosEmails, dadosObsidian, periodo) {
  const isManha = (periodo === "Manhã");

  let instrucaoVIP = `STATUS: Não há e-mails de remetentes VIP.`;
  if (dadosEmails.vipsEncontrados > 0) {
    instrucaoVIP = `<h2> 2. PRIORIDADES VIP </h2>
       Analise detalhadamente os e-mails VIP: Contexto, Solicitação e Ação Imediata.`;
  }

  let seccoesExtras = "";
  let dataPayload = "";

  if (isManha) {
    seccoesExtras = `
      <h2> 1. BRIEFING DA AGENDA </h2>
      <p>Apresente os compromissos de hoje e amanhã de forma clara.</p>

      <h2> 4. PLANEJAMENTO E PROJETOS (OBSIDIAN) </h2>
      <p><strong>Visão Estratégica:</strong> Consolide os dados dos cartões MD das pastas e subpastas. Resuma o status dos projetos ativos e planos de gestão.</p>
    `;
    dataPayload = `AGENDA:\n${agenda}\nOBSIDIAN:\n${dadosObsidian}\n`;
  }

  return `Você é o assessor executivo de ${CONFIG.NOME_USUARIO}. Gere o resumo para o período: ${periodo}.
  REGRAS: Retorne apenas HTML válido, sem emojis, use nomes completos.

  ${seccoesExtras}

  ${instrucaoVIP}

  <h2> 3. E-MAILS E ASSUNTOS GERAIS </h2>
  [Resuma por tema. Se for Tarde/Noite, seja ultra direto. Use tag <span style="color:red; font-weight:bold;">[URGENTE]</span> se necessário.]

  DADOS:\n${dataPayload}EMAILS:\n${dadosEmails.textoParaResumir}`;
}

/**
 * Communication with Gemini API. 
 * Implements 5 spaced Retry attempts (15s, 30s, 45s, 60s, 75s) to respect Google's timeout limits.
 */
function chamarGeminiAPI(prompt) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${CONFIG.API_KEY}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxTentativas = 5;
  
  for (let t = 1; t <= maxTentativas; t++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const body = response.getContentText();
      
      if (code === 200) {
        const json = JSON.parse(body);
        if (json.candidates && json.candidates.length > 0) {
          Logger.log(`[GEMINI_API] Request successful. Attempt: ${t}/${maxTentativas}.`);
          return json.candidates[0].content.parts[0].text;
        }
      } else if (code === 503 || code === 429) {
         const delay = t * 15000; 
         Logger.log(`[GEMINI_API] High traffic (HTTP ${code}). Attempt ${t} failed. Waiting ${delay/1000}s...`);
         Utilities.sleep(delay);
      } else {
        Logger.log(`[GEMINI_API_ERROR] Structural or permission error (HTTP ${code}): ${body}`);
        break; 
      }
    } catch (err) {
      Logger.log(`[GEMINI_API_ERROR] Connection failure on attempt ${t}: ${err}`);
      Utilities.sleep(15000); 
    }
  }
  
  Logger.log("[GEMINI_API] Attempts exhausted. Returning Null.");
  return null;
}

/**
 * Sends the formatted briefing via GmailApp.
 */
function enviarEmail(html, data) {
  Logger.log('[GMAIL_OUTBOX] Preparing outgoing email...');
  const hora = data.getHours();
  let p = "Manhã";
  if (hora >= 12 && hora < 18) p = "Tarde";
  if (hora >= 18) p = "Noite";
  
  const dataFmt = Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  const template = `
    <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; line-height: 1.6;">
      ${html}
      <hr style="border: 0; border-top: 1px solid #eee; margin-top: 30px;">
      <p style="font-size: 11px; color: #999;">Automated Executive Briefing System | AI Integration</p>
    </div>
  `;
  
  try {
    GmailApp.sendEmail(CONFIG.EMAIL_DESTINO, `Resumo Executivo - ${p} (${dataFmt})`, "", { htmlBody: template });
    Logger.log('[GMAIL_OUTBOX] Email sent successfully.');
  } catch (e) {
    Logger.log(`[GMAIL_OUTBOX_ERROR] Falha crítica ao enviar o e-mail: ${e.message}`);
  }
}

/**
 * Sends VIP email alerts to Google Chat via Webhook.
 */
function notificarVIPNoChat(qtd, resumo) {
  if (!CONFIG.WEBHOOK_CHAT) {
    Logger.log('[NOTIFICATION] Webhook not configured. Skipping Chat alert.');
    return;
  }
  
  Logger.log(`[NOTIFICATION] Triggering alert for ${qtd} VIPs via Webhook...`);
  const body = { text: `*ALERTA VIP: ${qtd} NOVA(S) MENSAGEM(NS)*\n\n${resumo}` };
  UrlFetchApp.fetch(CONFIG.WEBHOOK_CHAT, {
    method: "post", contentType: "application/json", payload: JSON.stringify(body)
  });
  Logger.log('[NOTIFICATION] Alert triggered in Chat successfully.');
}

/**
 * Helper Func: Checks if a given date falls on a weekend.
 */
function isFinalDeSemana(d) {
  const s = d.getDay();
  if (s === 0 || s === 6) {
    return true;
  }
  return false;
}

/**
 * Helper Func: Removes Markdown code block syntax from the LLM output.
 */
function limparFormatacaoMarkdown(t) {
  return t.replace(/```html/g, "").replace(/```/g, "").trim();
}
