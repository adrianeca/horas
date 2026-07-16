// =============================================================================
// CONFIGURAÇÕES
// =============================================================================

const HORAS_SHEET_ID = '1fbBw4ynJqpIkBwQR0dQyIjnGaw3WBh_KY9bRc-nN-Lg';
const FUNC_SHEET_ID  = '1BDiPjv0FqRJp5EwcvLdYXVvEAWesvwdEgbhYdnTlqPY';
const HUB_SS_ID       = '1eZPbzhzjhjHoPwMhAW5YvOZgYiAvlTYc07dRan6Lyoc';
const MEU_ACESSO      = 'webhoras';
const HUB_URL         = 'https://script.google.com/a/macros/brasas.com/s/AKfycbyF7BArYMYFtcQY7_4RTGGPw89yNohAjR7eGptItP-EsnWhNfiZR2ISRaHdAkwlLSlr/exec';

// E-mail que recebe as solicitações de liberação feitas pelos diretores
const DP_EMAIL = 'dp@brasas.com';

// Motivos que o diretor pode selecionar ao pedir liberação (AJUSTAR: lista provisória)
const MOTIVOS_LIBERACAO = [
  'Esqueci de lançar dentro do prazo',
  'Correção de lançamento com erro',
  'Professor admitido após o fechamento',
  'Outro'
];

// Índices das colunas na planilha de funcionários (base 0), aba "RJ - UNIDADES"
const COL = {
  NOME:         2,   // C
  APELIDO:      6,   // G  (APELIDO (TEACHER))
  FUNCAO:       5,   // F
  ATIVO:        10,  // K
  UNIDADE:      21,  // V  (Unidade Ajustada)
  MATRICULA:    27,  // AB
  UNIDADE_SEC:  30,  // AE
  NIVEL:        33   // AH (NÍVEL PROFESSORES)
};

// Colunas da aba "HORAS" (planilha central) que o app LÊ e ESCREVE — A a Q (0-based 0-16).
// As colunas R a V (17-21: Chave Matrícula, Apelido Ajustado, Nível Ajustado,
// Chave Matrícula Unidade, Data) já têm FÓRMULA na planilha — o app NUNCA escreve nelas.
// "Editado Em"/"Editado Por" ficam depois delas, em W/X (22-23), pra não colidir.
const HORAS_COL = {
  UNIDADE: 0, MES: 1, ANO: 2, MATRICULA: 3, APELIDO: 4, NOME: 5, NIVEL: 6,
  HORAS_TURMAS: 7, HORAS_TURMAS_SABADO: 8, ATIV_EXTRAS: 9, SUBS: 10,
  FALTAS_DESCONTADAS: 11, FALTAS_ABONADAS: 12, TOP_SPECIAL: 13,
  SUBS_OUTRAS_UNIDADE: 14, FALTAS_DESCONTADAS_DIAS: 15, FALTAS_ABONADAS_DIAS: 16,
  EDITADO_EM: 22, EDITADO_POR: 23
};
const HORAS_ABA = 'HORAS';

// Normaliza texto para comparação: minúsculo, sem acento, sem espaços nas bordas
function norm_(s) {
  return String(s || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// Algumas fontes (Hub, lançamentos antigos) chamam a mesma unidade de "NS" (na
// verdade CH) ou "MRI" (na verdade MR). Toda unidade crua deve passar por aqui
// assim que é lida, pra não duplicar a unidade em listas/filtros/lembretes.
const UNIDADE_ALIASES_ = { ns: 'CH', mri: 'MR' };
function canonUnidade_(u) {
  u = String(u || '').trim();
  if (!u) return u;
  const alias = UNIDADE_ALIASES_[norm_(u)];
  return alias || u;
}

// Extrai o número do mês mesmo quando a célula guarda texto como "06 Junho" (em vez de 6)
function parseMes_(v) {
  return parseInt(String(v).trim(), 10) || 0;
}

// Padrão de escrita do mês na planilha: "06 Junho", "07 Julho"...
const MESES_LABEL = ['01 Janeiro', '02 Fevereiro', '03 Março', '04 Abril', '05 Maio', '06 Junho',
                     '07 Julho', '08 Agosto', '09 Setembro', '10 Outubro', '11 Novembro', '12 Dezembro'];

function mesLabel_(m) {
  const n = parseMes_(m);
  return MESES_LABEL[n - 1] || String(m);
}

// Formata data+hora para exibição (colunas "Editado Em")
function fmtDataHora_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  return String(v);
}

// Coluna ATIVO (K) guarda o texto "Ativo"/"Inativo" (às vezes true/false, sim/não)
function isInativo_(v) {
  const n = norm_(v);
  return n === 'inativo' || n === 'inactive' || n === 'false' || n === 'nao' || n === 'no' || n === '0';
}

// =============================================================================
// ENTRY POINT
// =============================================================================

function doGet(e) {
  const token = (e && e.parameter && e.parameter.s) ? e.parameter.s : '';
  const tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.token = token;
  return tmpl.evaluate()
    .setTitle('Horas de Professores — BRASAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =============================================================================
// AUTENTICAÇÃO
// =============================================================================

// Colunas SESSOES: TOKEN(0)|EMAIL(1)|NOME(2)|ROLE(3)|UNIDADE(4)|CRIADO_EM(5)|EXPIRA_EM(6)|ACESSOS(7)
// UNIDADE pode ser pipe-separado (ex: "BG|FG"). Vazio = acesso a todas.

function getUserFromHub(token) {
  if (!token) throw new Error('Token não fornecido.');

  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');

  return user;
}

function getSessionUser_(token) {
  if (!token) return null;
  try {
    const ss       = SpreadsheetApp.openById(HUB_SS_ID);
    const sesSheet = ss.getSheetByName('SESSOES');
    if (!sesSheet) return null;

    const tok   = String(token).trim();
    const found = sesSheet.getRange(1, 1, sesSheet.getLastRow(), 1)
      .createTextFinder(tok).matchEntireCell(true).findNext();
    if (!found) return null;

    // [TOKEN, EMAIL, NOME, ROLE, UNIDADE, CRIADO_EM, EXPIRA_EM, ACESSOS]
    const row = sesSheet.getRange(found.getRow(), 1, 1, 8).getValues()[0];

    if (row[6] && new Date(row[6]) < new Date()) return null; // expirado

    const email = String(row[1] || '').trim().toLowerCase();
    if (!email) return null;

    // Verifica acesso a este dashboard na coluna ACESSOS
    const acessos = String(row[7] || '').toLowerCase()
      .split(',').map(function(a) { return a.trim(); });
    if (!acessos.includes(MEU_ACESSO)) {
      throw new Error('Você não tem permissão para acessar o Controle de Horas. Contacte o administrador.');
    }

    // UNIDADE: vazio = todas; pipe-separado = restringe a essas
    const unidadeRaw = String(row[4] || '').trim();
    const units = unidadeRaw
      ? unidadeRaw.split('|').map(function(u) { return canonUnidade_(u); }).filter(Boolean)
      : [];

    return {
      email:    email,
      nome:     String(row[2] || '').trim(),
      role:     String(row[3] || '').trim().toLowerCase(),
      unidade:  units[0] || '',
      units:    units  // [] = acesso total; preenchido = só essas unidades
    };
  } catch (e) {
    if (e.message && e.message.includes('permissão')) throw e;
    Logger.log('getSessionUser_: ' + e);
    return null;
  }
}

function isUserAllowedUnit_(user, unit) {
  if (!user.units || !user.units.length) return true; // acesso total
  return user.units.some(function(u) {
    return u.toLowerCase().trim() === unit.toLowerCase().trim();
  });
}

// Todas as unidades que o usuário pode ver: as dele (se restrito) ou todas que existem
// (com professor ativo cadastrado OU já com lançamento na planilha de Horas).
function getAllowedUnidades_(user) {
  const set = {};

  if (user.units && user.units.length > 0) {
    user.units.forEach(function(u) { set[u] = true; });
  } else {
    const funcSheet = SpreadsheetApp.openById(FUNC_SHEET_ID).getSheetByName('RJ - UNIDADES');
    if (!funcSheet) throw new Error('Aba "RJ - UNIDADES" não encontrada.');
    const funcRows = funcSheet.getDataRange().getValues();
    for (let i = 1; i < funcRows.length; i++) {
      const nome = String(funcRows[i][COL.NOME] || '').trim();
      if (!nome) continue;
      if (isInativo_(funcRows[i][COL.ATIVO])) continue;
      if (String(funcRows[i][COL.FUNCAO]).trim().toUpperCase() !== 'PROFESSOR') continue;
      const u = canonUnidade_(funcRows[i][COL.UNIDADE]);
      if (u) set[u] = true;
    }

    const sheet = getHorasSheet_();
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const u = canonUnidade_(rows[i][HORAS_COL.UNIDADE]);
      if (u) set[u] = true;
    }
  }

  const result = Object.keys(set).sort();
  if (!result.length) throw new Error('Controle de Horas ainda não está disponível para sua unidade.');
  return result;
}

// Retorna lista de unidades disponíveis para o usuário
function getUnidades(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');
  return getAllowedUnidades_(user);
}

// =============================================================================
// PERÍODO VIGENTE — diferente do VR/VT: aqui o professor lança o PRÓPRIO mês
// corrente (ex.: horas de julho são lançadas em julho), não o mês seguinte.
// O campo ainda se chama "previsto" só para reaproveitar o mesmo formato de
// resposta usado no resto do app (Index.html lê period.previsto.mes/ano).
// =============================================================================

function getCurrentPeriod(token) {
  const now = new Date();
  const previstoMes = now.getMonth() + 1;
  const previstoAno = now.getFullYear();

  // Aberto até o dia 11 do próprio mês; a partir do dia 12 bloqueia automaticamente
  let locked = now.getDate() > 11;

  // Liberação temporária (válida até 23:59 do dia da concessão) ignora o bloqueio para esse usuário
  if (locked) {
    const user = getSessionUser_(token);
    if (user && hasActiveLiberacao_(user.email)) locked = false;
  }

  return {
    previsto: { mes: previstoMes, ano: previstoAno },
    locked:   locked
  };
}

// =============================================================================
// LIBERAÇÕES TEMPORÁRIAS DE EDIÇÃO (até 23:59 do dia) — restrito a admins
// =============================================================================

// Colunas: Email | Liberado Por | Criado Em | Expira Em
function getLiberacoesSheet_() {
  const ss = SpreadsheetApp.openById(HORAS_SHEET_ID);
  let sheet = ss.getSheetByName('LIBERACOES');
  if (!sheet) {
    sheet = ss.insertSheet('LIBERACOES');
    sheet.appendRow(['Email', 'Liberado Por', 'Criado Em', 'Expira Em']);
  }
  return sheet;
}

function requireAdmin_(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');
  if (user.role !== 'admin' && user.role !== 'dp') throw new Error('Acesso restrito a administradores e ao Departamento Pessoal.');
  return user;
}

function hasActiveLiberacao_(email) {
  if (!email) return false;
  const emailNorm = norm_(email);
  const now  = new Date();
  const rows = getLiberacoesSheet_().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (norm_(r[0]) === emailNorm && r[3] && new Date(r[3]) > now) return true;
  }
  return false;
}

// Lista todas as liberações já concedidas (mais recentes primeiro) — só para admins
function getLiberacoes(token) {
  requireAdmin_(token);
  const rows = getLiberacoesSheet_().getDataRange().getValues();

  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    list.push({ email: String(r[0]).trim(), liberadoPor: String(r[1]).trim(), criadoEm: r[2], expiraEm: r[3] });
  }
  list.sort(function(a, b) { return new Date(b.criadoEm) - new Date(a.criadoEm); });
  return list;
}

// Concede edição liberada até 23:59 do dia da concessão — só admins podem chamar
function criarLiberacao(token, email) {
  const admin = requireAdmin_(token);

  email = String(email || '').trim().toLowerCase();
  if (!email || email.indexOf('@') === -1) throw new Error('Informe um e-mail válido.');

  const now    = new Date();
  const expira = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

  getLiberacoesSheet_().appendRow([email, admin.email, now, expira]);
  enviarEmailLiberacao_(email, now, expira);
  return getLiberacoes(token);
}

// Busca o nome cadastrado no Hub (SESSOES) para personalizar o e-mail. Vazio se não achar.
function findNomeByEmail_(email) {
  try {
    const sheet = SpreadsheetApp.openById(HUB_SS_ID).getSheetByName('SESSOES');
    if (!sheet) return '';
    const rows      = sheet.getDataRange().getValues();
    const emailNorm = norm_(email);
    for (let i = rows.length - 1; i >= 1; i--) {
      if (norm_(rows[i][1]) === emailNorm && rows[i][2]) return String(rows[i][2]).trim();
    }
    return '';
  } catch (e) {
    return '';
  }
}

// Avisa por e-mail quem recebeu a liberação. Falha de envio não deve derrubar a liberação em si.
function enviarEmailLiberacao_(email, criadoEm, expira) {
  try {
    const pad = function(n) { return n < 10 ? '0' + n : n; };
    const fmtDT = function(d) {
      return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear() +
        ' às ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
    };
    const criadoStr = fmtDT(criadoEm);
    const expiraStr = fmtDT(expira);
    const nome      = findNomeByEmail_(email);

    const assunto = 'Liberação temporária de edição — Horas de Professores';

    const corpoTexto =
      'Olá' + (nome ? ', ' + nome.toUpperCase() : '') + '.\n\n' +
      'Você recebeu uma liberação temporária para editar o Controle de Horas de Professores fora do prazo normal, ' +
      'concedida pelo Departamento Pessoal.\n\n' +
      'ATENÇÃO: a liberação vale SOMENTE HOJE, até as 23:59. Amanhã a edição já estará bloqueada novamente.\n\n' +
      'Concedida em: ' + criadoStr + '\n' +
      'Válida até: ' + expiraStr + ' (hoje)\n\n' +
      'Acesse pelo Hub BRASAS BI: ' + HUB_URL + '\n\n' +
      'Após esse horário, a edição volta a ficar bloqueada automaticamente.\n\n' +
      'Equipe BRASAS BI';

    const corpoHtml =
      '<div style="background:#f1f5f9;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
        '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e2e8f0">' +

          '<div style="background:#0a1628;padding:26px 32px">' +
            '<h1 style="margin:0;color:#ffffff;font-size:20px;font-weight:700;line-height:1.3">' +
              'Liberação Temporária de Edição — Horas de Professores' +
            '</h1>' +
          '</div>' +

          '<div style="padding:28px 32px">' +

            '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 16px;margin-bottom:22px;font-size:13px;color:#475569">' +
              '&#9888;&#65039; Este é um comunicado automático. <strong>Não responda este e-mail</strong> — em caso de dúvidas, entre em contato com o Departamento Pessoal pelo endereço <a href="mailto:dp@brasas.com" style="color:#2a4d76">dp@brasas.com</a>.' +
            '</div>' +

            '<p style="margin:0 0 14px;font-size:15px;color:#0f2035">Olá' +
              (nome ? ', <strong>' + nome.toUpperCase() + '</strong>' : '') + '.</p>' +

            '<p style="margin:0 0 14px;font-size:15px;color:#0f2035;line-height:1.6">' +
              'Você recebeu uma liberação temporária para editar o ' +
              '<strong>Controle de Horas de Professores</strong> fora do prazo normal (dia 11).' +
            '</p>' +

            '<div style="background:#fef3c7;border:1px solid #fbbf24;border-radius:8px;padding:14px 16px;margin-bottom:22px;font-size:14px;color:#92400e;line-height:1.5">' +
              '&#9200; A liberação vale <strong>somente hoje, até as 23:59</strong>. Amanhã a edição já estará bloqueada novamente.' +
            '</div>' +

            '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:18px 20px;margin-bottom:24px">' +
              '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;margin-bottom:14px">Dados da liberação</div>' +
              '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;width:150px">Liberado por:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">Departamento Pessoal</td>' +
                '</tr>' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b">Concedida em:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + criadoStr + '</td>' +
                '</tr>' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b">Válida até:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + expiraStr + ' (hoje)</td>' +
                '</tr>' +
              '</table>' +
            '</div>' +

            '<div style="text-align:center;margin-bottom:8px">' +
              '<a href="' + HUB_URL + '" style="display:inline-block;background:#0f2035;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 28px;border-radius:8px">' +
                'Acessar o Controle de Horas' +
              '</a>' +
            '</div>' +

            '<p style="margin:20px 0 0;font-size:12.5px;color:#94a3b8;text-align:center">' +
              'Após as 23:59 de hoje, a edição volta a ficar bloqueada automaticamente.' +
            '</p>' +

          '</div>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail(email, assunto, corpoTexto, { htmlBody: corpoHtml, name: 'Horas de Professores — BRASAS BI' });
  } catch (e) {
    Logger.log('enviarEmailLiberacao_: falha ao enviar e-mail para ' + email + ' — ' + e);
  }
}

// =============================================================================
// SOLICITAÇÕES DE LIBERAÇÃO — diretor pede pelo painel, DP aprova ou reprova
// =============================================================================

// Escapa texto livre antes de inseri-lo no HTML dos e-mails
function escHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Colunas: ID | Email | Nome | Motivo | Observações | Status | Criado Em | Respondido Por | Respondido Em | Observação DP
function getSolicitacoesSheet_() {
  const ss = SpreadsheetApp.openById(HORAS_SHEET_ID);
  let sheet = ss.getSheetByName('SOLICITACOES');
  if (!sheet) {
    sheet = ss.insertSheet('SOLICITACOES');
    sheet.appendRow(['ID', 'Email', 'Nome', 'Motivo', 'Observações', 'Status',
                     'Criado Em', 'Respondido Por', 'Respondido Em', 'Observação DP']);
  }
  return sheet;
}

function solicitacaoFromRow_(r) {
  return {
    id: String(r[0]),
    email: String(r[1]).trim(),
    nome: String(r[2]).trim(),
    motivo: String(r[3]).trim(),
    obs: String(r[4]).trim(),
    status: String(r[5]).trim(),          // pendente | aprovada | reprovada
    criadoEm: r[6],
    respondidoPor: String(r[7] || '').trim(),
    respondidoEm: r[8] || '',
    obsDP: String(r[9] || '').trim()
  };
}

// Dados que o painel do diretor precisa: motivos disponíveis + solicitação pendente dele (se houver)
function getSolicitacaoInfo(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');

  const rows      = getSolicitacoesSheet_().getDataRange().getValues();
  const emailNorm = norm_(user.email);
  let pendente    = null;

  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    if (norm_(r[1]) === emailNorm && norm_(r[5]) === 'pendente') { pendente = solicitacaoFromRow_(r); break; }
  }

  return { motivos: MOTIVOS_LIBERACAO, pendente: pendente, liberado: hasActiveLiberacao_(user.email) };
}

// Diretor registra o pedido; o DP é avisado por e-mail
function criarSolicitacaoLiberacao(token, motivo, obs) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');

  motivo = String(motivo || '').trim();
  obs    = String(obs || '').trim();
  if (MOTIVOS_LIBERACAO.indexOf(motivo) === -1) throw new Error('Selecione um motivo válido.');

  const info = getSolicitacaoInfo(token);
  if (info.pendente) throw new Error('Você já tem uma solicitação pendente aguardando resposta do DP.');
  if (info.liberado) throw new Error('Você já está com a edição liberada hoje.');

  const id = Utilities.getUuid();
  getSolicitacoesSheet_().appendRow([id, user.email, user.nome, motivo, obs, 'pendente',
                                     new Date(), '', '', '']);
  enviarEmailSolicitacaoDP_(user, motivo, obs);
  return getSolicitacaoInfo(token);
}

// Lista todas as solicitações (mais recentes primeiro) — só admins/DP
function getSolicitacoes(token) {
  requireAdmin_(token);
  const rows = getSolicitacoesSheet_().getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    list.push(solicitacaoFromRow_(rows[i]));
  }
  list.sort(function(a, b) {
    const pa = a.status === 'pendente' ? 0 : 1;
    const pb = b.status === 'pendente' ? 0 : 1;
    return (pa - pb) || (new Date(b.criadoEm) - new Date(a.criadoEm));
  });
  return list;
}

// DP aprova ou reprova; aprovar cria a liberação (até 23:59 de hoje) e avisa o diretor por e-mail
function responderSolicitacao(token, id, aprovar, obsDP) {
  const admin = requireAdmin_(token);
  obsDP = String(obsDP || '').trim();

  const sheet = getSolicitacoesSheet_();
  const rows  = sheet.getDataRange().getValues();
  let rowIdx  = -1;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) { rowIdx = i; break; }
  }
  if (rowIdx === -1) throw new Error('Solicitação não encontrada.');

  const solic = solicitacaoFromRow_(rows[rowIdx]);
  if (solic.status !== 'pendente') throw new Error('Esta solicitação já foi respondida.');

  const now    = new Date();
  const status = aprovar ? 'aprovada' : 'reprovada';
  sheet.getRange(rowIdx + 1, 6).setValue(status);                                    // Status
  sheet.getRange(rowIdx + 1, 8, 1, 3).setValues([[admin.email, now, obsDP]]);       // Respondido Por | Respondido Em | Observação DP

  let expira = null;
  if (aprovar) {
    expira = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
    getLiberacoesSheet_().appendRow([solic.email, admin.email, now, expira]);
  }

  enviarEmailRespostaSolicitacao_(solic, aprovar, obsDP, expira);
  return { solicitacoes: getSolicitacoes(token), liberacoes: getLiberacoes(token) };
}

// Aviso ao DP de que existe uma nova solicitação para analisar
function enviarEmailSolicitacaoDP_(user, motivo, obs) {
  try {
    const assunto = 'Nova solicitação de liberação de edição — Horas de Professores';

    const corpoTexto =
      'Olá.\n\n' +
      'Uma nova solicitação de liberação de edição do Controle de Horas de Professores foi registrada e aguarda sua análise.\n\n' +
      'Solicitante: ' + (user.nome || user.email) + ' (' + user.email + ')\n' +
      'Motivo: ' + motivo + '\n' +
      (obs ? 'Observações: ' + obs + '\n' : '') +
      '\nPara aprovar ou reprovar, acesse o Controle de Horas pelo Hub BRASAS BI e abra a aba "Liberação":\n' +
      HUB_URL + '\n\n' +
      'Equipe BRASAS BI';

    const corpoHtml =
      '<div style="background:#f1f5f9;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
        '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e2e8f0">' +
          '<div style="background:#0a1628;padding:26px 32px">' +
            '<h1 style="margin:0;color:#ffffff;font-size:20px;font-weight:700;line-height:1.3">' +
              'Nova Solicitação de Liberação — Horas de Professores' +
            '</h1>' +
          '</div>' +
          '<div style="padding:28px 32px">' +
            '<p style="margin:0 0 22px;font-size:15px;color:#0f2035;line-height:1.6">' +
              'Uma nova solicitação de liberação de edição foi registrada e <strong>aguarda sua análise</strong>.' +
            '</p>' +
            '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:18px 20px;margin-bottom:24px">' +
              '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;margin-bottom:14px">Dados da solicitação</div>' +
              '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;width:130px;vertical-align:top">Solicitante:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + escHtml_(user.nome || user.email) + '<br><span style="font-weight:400;color:#64748b">' + escHtml_(user.email) + '</span></td>' +
                '</tr>' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;vertical-align:top">Motivo:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + escHtml_(motivo) + '</td>' +
                '</tr>' +
                (obs ?
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;vertical-align:top">Observações:</td>' +
                  '<td style="padding:6px 0;color:#0f2035">' + escHtml_(obs) + '</td>' +
                '</tr>' : '') +
              '</table>' +
            '</div>' +
            '<div style="text-align:center;margin-bottom:8px">' +
              '<a href="' + HUB_URL + '" style="display:inline-block;background:#0f2035;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 28px;border-radius:8px">' +
                'Analisar solicitação' +
              '</a>' +
            '</div>' +
            '<p style="margin:20px 0 0;font-size:12.5px;color:#94a3b8;text-align:center">' +
              'Abra a aba "Liberação" do Controle de Horas para aprovar ou reprovar.' +
            '</p>' +
          '</div>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail(DP_EMAIL, assunto, corpoTexto, { htmlBody: corpoHtml, name: 'Horas de Professores — BRASAS BI' });
  } catch (e) {
    Logger.log('enviarEmailSolicitacaoDP_: falha ao enviar e-mail — ' + e);
  }
}

// Resposta ao diretor: aprovada (com validade até 23:59 de hoje) ou reprovada, com observação do DP
function enviarEmailRespostaSolicitacao_(solic, aprovada, obsDP, expira) {
  try {
    const pad = function(n) { return n < 10 ? '0' + n : n; };
    const fmtDT = function(d) {
      return pad(d.getDate()) + '/' + pad(d.getMonth() + 1) + '/' + d.getFullYear() +
        ' às ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
    };
    const nome = solic.nome || findNomeByEmail_(solic.email);

    const assunto = aprovada
      ? 'Solicitação de liberação APROVADA — Horas de Professores'
      : 'Solicitação de liberação reprovada — Horas de Professores';

    const corpoTexto =
      'Olá' + (nome ? ', ' + nome.toUpperCase() : '') + '.\n\n' +
      (aprovada
        ? 'Sua solicitação de liberação de edição do Controle de Horas de Professores foi APROVADA pelo Departamento Pessoal.\n\n' +
          'ATENÇÃO: a liberação vale SOMENTE HOJE, até as 23:59 (' + fmtDT(expira) + '). Amanhã a edição já estará bloqueada novamente.\n\n'
        : 'Sua solicitação de liberação de edição do Controle de Horas de Professores foi REPROVADA pelo Departamento Pessoal.\n\n') +
      'Motivo informado por você: ' + solic.motivo + '\n' +
      (obsDP ? 'Observação do DP: ' + obsDP + '\n' : '') +
      (aprovada ? '\nAcesse pelo Hub BRASAS BI: ' + HUB_URL + '\n' : '') +
      '\nEm caso de dúvidas, entre em contato com o Departamento Pessoal (dp@brasas.com).\n\n' +
      'Equipe BRASAS BI';

    const corBarra  = aprovada ? '#15803d' : '#dc2626';
    const titulo    = aprovada ? 'Solicitação Aprovada — Horas de Professores' : 'Solicitação Reprovada — Horas de Professores';

    const corpoHtml =
      '<div style="background:#f1f5f9;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
        '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e2e8f0">' +
          '<div style="background:' + corBarra + ';padding:26px 32px">' +
            '<h1 style="margin:0;color:#ffffff;font-size:20px;font-weight:700;line-height:1.3">' + titulo + '</h1>' +
          '</div>' +
          '<div style="padding:28px 32px">' +
            '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 16px;margin-bottom:22px;font-size:13px;color:#475569">' +
              '&#9888;&#65039; Este é um comunicado automático. <strong>Não responda este e-mail</strong> — em caso de dúvidas, entre em contato com o Departamento Pessoal pelo endereço <a href="mailto:dp@brasas.com" style="color:#2a4d76">dp@brasas.com</a>.' +
            '</div>' +
            '<p style="margin:0 0 14px;font-size:15px;color:#0f2035">Olá' +
              (nome ? ', <strong>' + escHtml_(nome.toUpperCase()) + '</strong>' : '') + '.</p>' +
            '<p style="margin:0 0 14px;font-size:15px;color:#0f2035;line-height:1.6">' +
              (aprovada
                ? 'Sua solicitação de liberação de edição do <strong>Controle de Horas de Professores</strong> foi <strong style="color:#15803d">APROVADA</strong> pelo Departamento Pessoal.'
                : 'Sua solicitação de liberação de edição do <strong>Controle de Horas de Professores</strong> foi <strong style="color:#dc2626">REPROVADA</strong> pelo Departamento Pessoal.') +
            '</p>' +
            (aprovada ?
            '<div style="background:#fef3c7;border:1px solid #fbbf24;border-radius:8px;padding:14px 16px;margin-bottom:22px;font-size:14px;color:#92400e;line-height:1.5">' +
              '&#9200; A liberação vale <strong>somente hoje, até as 23:59</strong>. Amanhã a edição já estará bloqueada novamente.' +
            '</div>' : '') +
            '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:18px 20px;margin-bottom:24px">' +
              '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;margin-bottom:14px">Resumo</div>' +
              '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;width:160px;vertical-align:top">Motivo informado:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + escHtml_(solic.motivo) + '</td>' +
                '</tr>' +
                (solic.obs ?
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;vertical-align:top">Suas observações:</td>' +
                  '<td style="padding:6px 0;color:#0f2035">' + escHtml_(solic.obs) + '</td>' +
                '</tr>' : '') +
                (obsDP ?
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;vertical-align:top">Observação do DP:</td>' +
                  '<td style="padding:6px 0;color:#0f2035">' + escHtml_(obsDP) + '</td>' +
                '</tr>' : '') +
                (aprovada ?
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b">Válida até:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + fmtDT(expira) + ' (hoje)</td>' +
                '</tr>' : '') +
              '</table>' +
            '</div>' +
            (aprovada ?
            '<div style="text-align:center;margin-bottom:8px">' +
              '<a href="' + HUB_URL + '" style="display:inline-block;background:#0f2035;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 28px;border-radius:8px">' +
                'Acessar o Controle de Horas' +
              '</a>' +
            '</div>' +
            '<p style="margin:20px 0 0;font-size:12.5px;color:#94a3b8;text-align:center">' +
              'Após as 23:59 de hoje, a edição volta a ficar bloqueada automaticamente.' +
            '</p>' : '') +
          '</div>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail(solic.email, assunto, corpoTexto, { htmlBody: corpoHtml, name: 'Horas de Professores — BRASAS BI' });
  } catch (e) {
    Logger.log('enviarEmailRespostaSolicitacao_: falha ao enviar e-mail para ' + solic.email + ' — ' + e);
  }
}

// =============================================================================
// FUNCIONÁRIOS (só professores)
// =============================================================================

// Retorna os professores de TODAS as unidades que o usuário pode ver.
// Um professor com unidade principal + secundária aparece uma vez para cada uma
// (desde que esteja entre as unidades permitidas), pois cada uma é um contexto de lançamento distinto.
function getFuncionarios(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');
  const allowedNorm = getAllowedUnidades_(user).map(norm_);

  const ss    = SpreadsheetApp.openById(FUNC_SHEET_ID);
  const sheet = ss.getSheetByName('RJ - UNIDADES');
  if (!sheet) throw new Error('Aba "RJ - UNIDADES" não encontrada na planilha de funcionários.');
  const rows  = sheet.getDataRange().getValues();

  const professores = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const nome = String(row[COL.NOME]).trim();
    if (!nome) continue;
    if (isInativo_(row[COL.ATIVO])) continue;
    if (String(row[COL.FUNCAO]).trim().toUpperCase() !== 'PROFESSOR') continue;

    const matricula = String(row[COL.MATRICULA]).trim();
    if (!matricula) continue;

    const apelido = String(row[COL.APELIDO] || '').trim();
    const nivel   = String(row[COL.NIVEL] || '').trim();

    // Unidade principal + secundária, sem repetir se forem iguais (já normalizadas: NS→CH, MRI→MR)
    const unidades = [canonUnidade_(row[COL.UNIDADE]), canonUnidade_(row[COL.UNIDADE_SEC])]
      .filter(function(u, idx, arr) { return u && arr.indexOf(u) === idx; });

    unidades.forEach(function(u) {
      if (allowedNorm.indexOf(norm_(u)) === -1) return;
      professores.push({ nome: nome, matricula: matricula, apelido: apelido, nivel: nivel, unidade: u });
    });
  }

  professores.sort(function(a, b) { return a.nome.localeCompare(b.nome, 'pt-BR') || a.unidade.localeCompare(b.unidade, 'pt-BR'); });

  return { professores: professores };
}

// =============================================================================
// PLANILHA HORAS
// =============================================================================

function getHorasSheet_() {
  const sheet = SpreadsheetApp.openById(HORAS_SHEET_ID).getSheetByName(HORAS_ABA);
  if (!sheet) throw new Error('Aba "' + HORAS_ABA + '" não encontrada na planilha.');
  // Garante o cabeçalho das colunas de edição (W/X), sem tocar nas colunas de fórmula R-V
  const editCol1 = HORAS_COL.EDITADO_EM + 1; // 1-based
  const header = sheet.getRange(1, editCol1, 1, 2).getValues()[0];
  if (!header[0]) {
    sheet.getRange(1, editCol1, 1, 2).setValues([['Editado Em', 'Editado Por']]);
  }
  return sheet;
}

// =============================================================================
// LEITURA DAS HORAS — todas as linhas das unidades permitidas, estilo planilha
// =============================================================================

function getHorasData(token) {
  const user = getSessionUser_(token);
  if (!user) throw new Error('Sessão inválida.');
  const allowedNorm = getAllowedUnidades_(user).map(norm_);

  const sheet = getHorasSheet_();
  const rows  = sheet.getDataRange().getValues();

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (allowedNorm.indexOf(norm_(canonUnidade_(r[HORAS_COL.UNIDADE]))) === -1) continue;
    const mes = parseMes_(r[HORAS_COL.MES]), ano = Number(r[HORAS_COL.ANO]);
    if (!mes || !ano) continue;

    out.push({
      unidade: canonUnidade_(r[HORAS_COL.UNIDADE]), mes: mes, ano: ano,
      matricula: String(r[HORAS_COL.MATRICULA]).trim(),
      apelido: String(r[HORAS_COL.APELIDO]).trim(),
      nome: String(r[HORAS_COL.NOME]).trim(),
      nivel: String(r[HORAS_COL.NIVEL]).trim(),
      horasTurmas: r[HORAS_COL.HORAS_TURMAS] || 0,
      horasTurmasSabado: r[HORAS_COL.HORAS_TURMAS_SABADO] || 0,
      ativExtras: r[HORAS_COL.ATIV_EXTRAS] || 0,
      subs: r[HORAS_COL.SUBS] || 0,
      faltasDescontadas: r[HORAS_COL.FALTAS_DESCONTADAS] || 0,
      faltasAbonadas: r[HORAS_COL.FALTAS_ABONADAS] || 0,
      topSpecial: r[HORAS_COL.TOP_SPECIAL] || 0,
      subsOutrasUnidade: r[HORAS_COL.SUBS_OUTRAS_UNIDADE] || 0,
      faltasDescontadasDias: r[HORAS_COL.FALTAS_DESCONTADAS_DIAS] || 0,
      faltasAbonadasDias: r[HORAS_COL.FALTAS_ABONADAS_DIAS] || 0,
      editadoEm: fmtDataHora_(r[HORAS_COL.EDITADO_EM]), editadoPor: String(r[HORAS_COL.EDITADO_POR] || '').trim()
    });
  }

  out.sort(function(a, b) {
    return (a.ano - b.ano) || (a.mes - b.mes) || a.unidade.localeCompare(b.unidade, 'pt-BR') || a.nome.localeCompare(b.nome, 'pt-BR');
  });

  return { professores: out };
}

// =============================================================================
// SALVAMENTO — cada item do payload já traz sua própria unidade/mês/ano.
// Escreve só as colunas A-Q (identidade + 10 campos numéricos); NUNCA toca em R-V
// (fórmulas da planilha: Chave Matrícula, Apelido Ajustado, Nível Ajustado,
// Chave Matrícula Unidade, Data), que se autopreenchem ao detectar a linha nova.
// =============================================================================

// Ordem espelha as colunas H-Q da aba HORAS
function valuesFromEntry_(e) {
  return [
    Number(e.horasTurmas) || 0, Number(e.horasTurmasSabado) || 0,
    Number(e.ativExtras) || 0, Number(e.subs) || 0,
    Number(e.faltasDescontadas) || 0, Number(e.faltasAbonadas) || 0,
    Number(e.topSpecial) || 0, Number(e.subsOutrasUnidade) || 0,
    Number(e.faltasDescontadasDias) || 0, Number(e.faltasAbonadasDias) || 0
  ];
}

function saveHorasData(payload) {
  // Passa o token para que uma liberação ativa do usuário destrave o salvamento
  const period = getCurrentPeriod(payload.token);
  if (period.locked) throw new Error('O período está bloqueado. Prazo encerrado no dia 11.');

  const user = getSessionUser_(payload.token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');

  // Nunca confia na unidade vinda do cliente sem checar permissão; normaliza NS→CH, MRI→MR
  // antes de checar/gravar, pra planilha sempre guardar o código canônico.
  const entries = (payload.professores || [])
    .map(function(e) { e.unidade = canonUnidade_(e.unidade); return e; })
    .filter(function(e) { return isUserAllowedUnit_(user, e.unidade); });
  if (!entries.length) return { success: true };

  const sheet   = getHorasSheet_();
  const allRows = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < allRows.length; i++) {
    const r = allRows[i];
    map[norm_(canonUnidade_(r[HORAS_COL.UNIDADE])) + '|' + parseMes_(r[HORAS_COL.MES]) + '|' + Number(r[HORAS_COL.ANO]) + '|' + String(r[HORAS_COL.MATRICULA]).trim()] = i + 1;
  }

  const valCol1  = HORAS_COL.HORAS_TURMAS + 1; // 1-based, primeira coluna de valores (H)
  const nVals    = 10;
  const editCol1 = HORAS_COL.EDITADO_EM + 1;   // 1-based (W)

  entries.forEach(function(e) {
    const mat    = String(e.matricula).trim();
    const key    = norm_(e.unidade) + '|' + Number(e.mes) + '|' + Number(e.ano) + '|' + mat;
    const values = valuesFromEntry_(e);

    if (map[key]) {
      const rowIdx = map[key];
      if (rowIdx <= allRows.length) {
        const oldVals = allRows[rowIdx - 1].slice(HORAS_COL.HORAS_TURMAS, HORAS_COL.HORAS_TURMAS + nVals);
        const editou  = values.some(function(v, j) {
          const antigo = Number(oldVals[j]) || 0;
          return antigo !== 0 && antigo !== v;
        });
        if (editou) {
          sheet.getRange(rowIdx, editCol1, 1, 2).setValues([[new Date(), user.email || '']]);
        }
      }
      sheet.getRange(rowIdx, valCol1, 1, nVals).setValues([values]);
    } else {
      // Colunas A-G (identidade) + H-Q (valores) = 17 células; R em diante fica intocado
      sheet.appendRow([e.unidade, mesLabel_(e.mes), e.ano, mat, e.apelido || '', e.nome, e.nivel || ''].concat(values));
      map[key] = sheet.getLastRow();
    }
  });

  return { success: true };
}

// Exclui um lançamento (unidade+mes+ano+matricula). Só permite excluir do período vigente (Previsto).
function deleteHorasEntry(payload) {
  const user = getSessionUser_(payload.token);
  if (!user) throw new Error('Sessão inválida ou expirada. Acesse novamente pelo Hub.');
  payload.unidade = canonUnidade_(payload.unidade);
  if (!isUserAllowedUnit_(user, payload.unidade)) throw new Error('Você não tem permissão para esta unidade.');

  const period = getCurrentPeriod(payload.token);
  if (period.locked) throw new Error('O período está bloqueado. Prazo encerrado no dia 11.');
  if (Number(payload.mes) !== period.previsto.mes || Number(payload.ano) !== period.previsto.ano) {
    throw new Error('Só é possível excluir lançamentos do período vigente.');
  }

  const sheet = getHorasSheet_();
  const key   = norm_(payload.unidade) + '|' + Number(payload.mes) + '|' + Number(payload.ano) + '|' + String(payload.matricula).trim();
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const k = norm_(canonUnidade_(r[HORAS_COL.UNIDADE])) + '|' + parseMes_(r[HORAS_COL.MES]) + '|' + Number(r[HORAS_COL.ANO]) + '|' + String(r[HORAS_COL.MATRICULA]).trim();
    if (k === key) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  // Linha não encontrada na planilha (provavelmente nunca foi salva) — nada a fazer
  return { success: true };
}

// =============================================================================
// LEMBRETES DE PREENCHIMENTO — gatilhos mensais (dias 1, 5, 9 e 11)
// =============================================================================
// Dia 1 : aviso de abertura do período, para todas as unidades (incondicional)
// Dia 5, 9 e 11: lembrete só para quem ainda não preencheu o mês Previsto;
//                dia 11 avisa que é o último dia antes do bloqueio automático.

// Lista fixa de e-mail do diretor por unidade (mesma usada no VR/VT — passada
// manualmente pela Adriane, mais confiável que derivar da planilha do Hub).
// Observação: o cadastro do Hub usa "NS" e "MRI", mas os lançamentos das
// planilhas usam "CH" e "MR" — por isso o mapeamento já usa os códigos certos.
const DIRETORES_UNIDADE = {
  'bf':     ['dirbf@brasas.com'],
  'bg':     ['dirbg@brasas.com'],
  'cg':     ['dircg@brasas.com'],
  'ch':     ['dirch@brasas.com'],  // cadastrado como "NS" no Hub
  'cp':     ['dircp@brasas.com'],
  'cx':     ['dircx@brasas.com'],
  'dt':     ['dirdt@brasas.com'],
  'fg':     ['dirfg@brasas.com'],
  'ig':     ['marcelo.ig@brasas.com'],
  'ip':     ['dirip@brasas.com'],
  'it':     ['dirit@brasas.com'],
  'lj':     ['dirlj@brasas.com'],
  'mr':     ['dirmr@brasas.com'],  // cadastrado como "MRI" no Hub
  'ni':     ['dirni@brasas.com'],
  'nl':     ['dirnl@brasas.com'],
  'nt':     ['dirnt@brasas.com'],
  'pc':     ['dirpc@brasas.com'],
  'po':     ['dirpo@brasas.com'],
  'rc':     ['dirrc@brasas.com'],
  'tj':     ['dirtj@brasas.com'],
  'tq':     ['dirtq@brasas.com'],
  'vp':     ['dirvp@brasas.com'],
  'vq':     ['dirvq@brasas.com'],
  'pn':     ['dirpn@brasas.com'],
  'online': ['natasha@brasas.com'],
  'bod':    ['pat@brasas.com'],
  'gr':     ['dirgr@brasas.com'],
  'vo':     ['dirvo@brasas.com']
};

// Mês/ano de referência = o mês CORRENTE (aqui o professor lança o próprio mês,
// diferente do VR/VT que lançam o mês seguinte) — mesma regra de getCurrentPeriod,
// mas sem depender de sessão/token.
function getPeriodoPrevistoAtual_() {
  const now = new Date();
  return { mes: now.getMonth() + 1, ano: now.getFullYear() };
}

// Retorna o mapa unidade -> [e-mails de diretor]. Unidades fora de
// DIRETORES_UNIDADE (ex.: EDITORA, EC NEW, MÉTODOS) não têm e-mail cadastrado
// e são puladas pelo restante do fluxo de lembretes.
function getMapaDiretoresPorUnidade_() {
  return DIRETORES_UNIDADE;
}

// Unidades que têm ao menos um PROFESSOR ativo cadastrado (aba "RJ - UNIDADES",
// coluna Unidade Ajustada, filtrado a FUNÇÃO = PROFESSOR — é só isso que este
// controle acompanha).
function getUnidadesAtivas_() {
  const sheet = SpreadsheetApp.openById(FUNC_SHEET_ID).getSheetByName('RJ - UNIDADES');
  if (!sheet) throw new Error('Aba "RJ - UNIDADES" não encontrada.');
  const rows = sheet.getDataRange().getValues();

  const set = {};
  for (let i = 1; i < rows.length; i++) {
    const nome = String(rows[i][COL.NOME] || '').trim();
    if (!nome) continue;
    if (isInativo_(rows[i][COL.ATIVO])) continue;
    if (String(rows[i][COL.FUNCAO]).trim().toUpperCase() !== 'PROFESSOR') continue;
    const u = canonUnidade_(rows[i][COL.UNIDADE]);
    if (u) set[u] = true;
  }
  return Object.keys(set);
}

// Unidades que já têm ao menos um lançamento salvo para o mês/ano informado
// (na aba HORAS) — usado para saber quem ainda falta preencher.
function getUnidadesPreenchidas_(mes, ano) {
  const sheet = getHorasSheet_();
  const rows  = sheet.getDataRange().getValues();
  const preenchidas = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (parseMes_(r[HORAS_COL.MES]) === mes && Number(r[HORAS_COL.ANO]) === ano) {
      const u = canonUnidade_(r[HORAS_COL.UNIDADE]);
      if (u) preenchidas[norm_(u)] = true;
    }
  }
  return preenchidas;
}

// Ponto de entrada dos gatilhos — dia: 1 (abertura), 5, 9 ou 11 (último dia)
function verificarEEnviarLembretesHoras_(dia) {
  const periodo    = getPeriodoPrevistoAtual_();
  const mesLabel    = mesLabel_(periodo.mes);
  const diretores   = getMapaDiretoresPorUnidade_();
  const unidades    = getUnidadesAtivas_();
  const preenchidas = dia === 1 ? {} : getUnidadesPreenchidas_(periodo.mes, periodo.ano);

  unidades.forEach(function(unidade) {
    const key = norm_(unidade);
    if (dia !== 1 && preenchidas[key]) return; // já preencheu — não incomoda mais

    const emails = diretores[key];
    if (!emails || !emails.length) return; // sem e-mail de diretor cadastrado — pula

    enviarEmailLembreteHoras_(emails, unidade, mesLabel, dia);
  });
}

// Monta e envia o e-mail de abertura/lembrete para uma unidade
function enviarEmailLembreteHoras_(emails, unidade, mesLabel, dia) {
  try {
    const destinatarios = emails.join(',');

    let assunto, tituloBarra, corBarra, mensagemPrincipal, avisoDestaque;

    if (dia === 1) {
      assunto = 'Preenchimento das Horas de Professores está aberto — ' + unidade;
      tituloBarra = 'Horas de Professores — Período Aberto';
      corBarra = '#0a1628';
      mensagemPrincipal = 'O período para preenchimento das <strong>Horas de Professores</strong> de <strong>' + escHtml_(mesLabel) + '</strong> está aberto.';
      avisoDestaque = 'Prazo final para preencher: <strong>dia 11</strong>. Após essa data, a edição é bloqueada automaticamente.';
    } else if (dia === 5) {
      assunto = 'Lembrete: Horas de Professores ainda não preenchidas — ' + unidade;
      tituloBarra = 'Lembrete — Horas de Professores Pendente';
      corBarra = '#b45309';
      mensagemPrincipal = 'Identificamos que as <strong>Horas de Professores</strong> de <strong>' + escHtml_(mesLabel) + '</strong> da sua unidade ainda não foram preenchidas.';
      avisoDestaque = 'Prazo final para preencher: <strong>dia 11</strong>. Após essa data, a edição é bloqueada automaticamente.';
    } else if (dia === 9) {
      assunto = '2º lembrete: Horas de Professores ainda não preenchidas — ' + unidade;
      tituloBarra = '2º Lembrete — Horas de Professores Pendente';
      corBarra = '#c2410c';
      mensagemPrincipal = 'As <strong>Horas de Professores</strong> de <strong>' + escHtml_(mesLabel) + '</strong> da sua unidade ainda não foram preenchidas.';
      avisoDestaque = 'Restam poucos dias! Prazo final: <strong>dia 11</strong>. Após essa data, a edição é bloqueada automaticamente.';
    } else { // dia 11
      assunto = 'ÚLTIMO DIA para preencher as Horas de Professores — ' + unidade;
      tituloBarra = 'Último Dia — Horas de Professores Pendente';
      corBarra = '#dc2626';
      mensagemPrincipal = 'Hoje é o <strong>último dia</strong> para preencher as <strong>Horas de Professores</strong> de <strong>' + escHtml_(mesLabel) + '</strong> da sua unidade.';
      avisoDestaque = 'Após hoje (23:59), a edição será bloqueada automaticamente e só poderá ser liberada mediante solicitação ao Departamento Pessoal.';
    }

    const corpoTexto =
      'Olá.\n\n' +
      mensagemPrincipal.replace(/<\/?strong>/g, '') + '\n\n' +
      avisoDestaque.replace(/<\/?strong>/g, '') + '\n\n' +
      'Unidade: ' + unidade + '\n' +
      'Mês de referência: ' + mesLabel + '\n\n' +
      'Acesse pelo Hub BRASAS BI: ' + HUB_URL + '\n\n' +
      'Equipe BRASAS BI';

    const corpoHtml =
      '<div style="background:#f1f5f9;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
        '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e2e8f0">' +
          '<div style="background:' + corBarra + ';padding:26px 32px">' +
            '<h1 style="margin:0;color:#ffffff;font-size:20px;font-weight:700;line-height:1.3">' + tituloBarra + '</h1>' +
          '</div>' +
          '<div style="padding:28px 32px">' +
            '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 16px;margin-bottom:22px;font-size:13px;color:#475569">' +
              '&#9888;&#65039; Este é um comunicado automático. <strong>Não responda este e-mail</strong> — em caso de dúvidas, entre em contato com o Departamento Pessoal pelo endereço <a href="mailto:dp@brasas.com" style="color:#2a4d76">dp@brasas.com</a>.' +
            '</div>' +
            '<p style="margin:0 0 14px;font-size:15px;color:#0f2035;line-height:1.6">' + mensagemPrincipal + '</p>' +
            '<div style="background:#fef3c7;border:1px solid #fbbf24;border-radius:8px;padding:14px 16px;margin-bottom:22px;font-size:14px;color:#92400e;line-height:1.5">' +
              '&#9200; ' + avisoDestaque +
            '</div>' +
            '<div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:18px 20px;margin-bottom:24px">' +
              '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#94a3b8;margin-bottom:14px">Dados do lançamento</div>' +
              '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b;width:150px">Unidade:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + escHtml_(unidade) + '</td>' +
                '</tr>' +
                '<tr>' +
                  '<td style="padding:6px 0;color:#64748b">Mês de referência:</td>' +
                  '<td style="padding:6px 0;color:#0f2035;font-weight:600">' + escHtml_(mesLabel) + '</td>' +
                '</tr>' +
              '</table>' +
            '</div>' +
            '<div style="text-align:center;margin-bottom:8px">' +
              '<a href="' + HUB_URL + '" style="display:inline-block;background:#0f2035;color:#ffffff;text-decoration:none;font-weight:700;font-size:14px;padding:12px 28px;border-radius:8px">' +
                'Preencher as Horas de Professores' +
              '</a>' +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>';

    MailApp.sendEmail(destinatarios, assunto, corpoTexto, { htmlBody: corpoHtml, name: 'Horas de Professores — BRASAS BI' });
  } catch (e) {
    Logger.log('enviarEmailLembreteHoras_: falha ao enviar para ' + emails + ' (unidade ' + unidade + ') — ' + e);
  }
}

// Funções chamadas pelos gatilhos instalados (uma para cada dia)
function lembreteHorasDia1()  { verificarEEnviarLembretesHoras_(1); }
function lembreteHorasDia5()  { verificarEEnviarLembretesHoras_(5); }
function lembreteHorasDia9()  { verificarEEnviarLembretesHoras_(9); }
function lembreteHorasDia11() { verificarEEnviarLembretesHoras_(11); }

// Roda ISSO UMA VEZ no editor do Apps Script (Ctrl+Enter) para instalar os
// 4 gatilhos mensais. Pode rodar de novo com segurança — remove os antigos antes.
function instalarGatilhosLembreteHoras() {
  const handlers = ['lembreteHorasDia1', 'lembreteHorasDia5', 'lembreteHorasDia9', 'lembreteHorasDia11'];
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('lembreteHorasDia1').timeBased().onMonthDay(1).atHour(8).create();
  ScriptApp.newTrigger('lembreteHorasDia5').timeBased().onMonthDay(5).atHour(8).create();
  ScriptApp.newTrigger('lembreteHorasDia9').timeBased().onMonthDay(9).atHour(8).create();
  ScriptApp.newTrigger('lembreteHorasDia11').timeBased().onMonthDay(11).atHour(8).create();

  Logger.log('Gatilhos de lembrete das Horas de Professores instalados com sucesso.');
}

// Simula um lembrete SEM enviar e-mails — só loga quem receberia e por quê.
// Rode no editor (Ctrl+Enter), ajustando "dia" abaixo, para validar antes de
// instalar os gatilhos de verdade.
function diagnosticoLembretesHoras() {
  const dia = 5; // ajuste para 1, 5, 9 ou 11 antes de rodar

  const periodo    = getPeriodoPrevistoAtual_();
  const mesLabel    = mesLabel_(periodo.mes);
  const diretores   = getMapaDiretoresPorUnidade_();
  const unidades    = getUnidadesAtivas_();
  const preenchidas = dia === 1 ? {} : getUnidadesPreenchidas_(periodo.mes, periodo.ano);

  Logger.log('=== SIMULAÇÃO — dia %s | mês previsto: %s/%s ===', dia, mesLabel, periodo.ano);
  unidades.forEach(function(unidade) {
    const key = norm_(unidade);
    const emails = diretores[key] || [];
    if (dia !== 1 && preenchidas[key]) {
      Logger.log('%s → já preenchida, não envia', unidade);
    } else if (!emails.length) {
      Logger.log('%s → SEM e-mail de diretor cadastrado, pula', unidade);
    } else {
      Logger.log('%s → enviaria para: %s', unidade, emails.join(', '));
    }
  });
}

// =============================================================================
// DIAGNÓSTICO — rode no editor do Apps Script e veja os logs (Ctrl+Enter)
// =============================================================================

function diagnosticoHoras() {
  const period = getCurrentPeriod();
  const sheet  = getHorasSheet_();

  Logger.log('=== PERÍODO ATUAL (só Previsto) ===');
  Logger.log('Previsto: mês %s / ano %s — bloqueado: %s', period.previsto.mes, period.previsto.ano, period.locked);

  Logger.log('\n=== LINHAS EM HORAS ===');
  sheet.getDataRange().getValues().slice(1).forEach(function(r, i) {
    Logger.log('Linha %s → unidade="%s" mes=%s ano=%s mat="%s" vals=%s',
      i + 2, r[HORAS_COL.UNIDADE], r[HORAS_COL.MES], r[HORAS_COL.ANO], r[HORAS_COL.MATRICULA],
      JSON.stringify(r.slice(HORAS_COL.HORAS_TURMAS, HORAS_COL.HORAS_TURMAS + 10)));
  });
}
