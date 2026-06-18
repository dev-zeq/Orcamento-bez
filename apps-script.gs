function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || 'precos';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'precos') {
    const sheet = ss.getSheetByName("Precos");
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const result = data.map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
    return json(result);
  }

  if (action === 'buscar') {
    const nome = (params.nome || '').toLowerCase().trim();
    const sheet = ss.getSheetByName("Leads");
    if (!sheet || !nome) return json([]);
    const rows = sheet.getDataRange().getValues();
    const headers = rows.shift();
    const nomeIdx = headers.indexOf('nome');
    const whatsIdx = headers.indexOf('whatsapp');
    const cidadeIdx = headers.indexOf('cidade');
    const endIdx = headers.indexOf('endereco');
    const seen = new Set();
    const results = [];
    rows.forEach(row => {
      const n = String(row[nomeIdx] || '').toLowerCase();
      const wpp = String(row[whatsIdx] || '');
      if (n.includes(nome) && !seen.has(wpp)) {
        seen.add(wpp);
        results.push({ nome: row[nomeIdx] || '', whatsapp: wpp, cidade: row[cidadeIdx] || '', endereco: (endIdx >= 0 ? row[endIdx] : '') || '' });
      }
    });
    return json(results);
  }

  if (action === 'disponivel') {
    const mes = parseInt(params.mes);
    const ano = parseInt(params.ano);
    const sheet = ss.getSheetByName("Agendamentos");
    const counts = {};
    if (sheet && sheet.getLastRow() > 1) {
      const rows = sheet.getDataRange().getValues();
      const headers = rows.shift();
      const dataIdx = headers.indexOf('data');
      const statusIdx = headers.indexOf('status');
      rows.forEach(row => {
        if (String(row[statusIdx] || '') === 'Cancelado') return;
        const d = row[dataIdx] instanceof Date ? row[dataIdx] : new Date(row[dataIdx]);
        if (!isNaN(d) && d.getMonth() + 1 === mes && d.getFullYear() === ano) {
          counts[d.getDate()] = (counts[d.getDate()] || 0) + 1;
        }
      });
    }
    return json(counts);
  }

  if (action === 'agendamentos') {
    const data = params.data;
    const sheet = ss.getSheetByName("Agendamentos");
    if (!sheet || !data) return json([]);
    const rows = sheet.getDataRange().getValues();
    const headers = rows.shift();
    const dataIdx = headers.indexOf('data');
    const statusIdx = headers.indexOf('status');
    const results = rows
      .filter(row => {
        const d = row[dataIdx] instanceof Date
          ? Utilities.formatDate(row[dataIdx], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(row[dataIdx]).slice(0, 10);
        return d === data && String(row[statusIdx] || '') !== 'Cancelado';
      })
      .map(row => { const obj = {}; headers.forEach((h, i) => obj[h] = row[i]); return obj; });
    return json(results);
  }

  return json({ erro: 'Acao invalida' });
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = JSON.parse(e.postData.contents);
  const action = data.action || 'lead';

  if (action === 'agendar') {
    let sheet = ss.getSheetByName("Agendamentos");
    if (!sheet) {
      sheet = ss.insertSheet("Agendamentos");
      sheet.appendRow(["data", "horario", "nome", "whatsapp", "cidade", "endereco", "observacoes", "status"]);
    } else if (sheet.getLastRow() > 1) {
      const rows = sheet.getDataRange().getValues();
      const headers = rows.shift();
      const dIdx = headers.indexOf('data');
      const hIdx = headers.indexOf('horario');
      const sIdx = headers.indexOf('status');
      const taken = rows.some(row => {
        const d = row[dIdx] instanceof Date
          ? Utilities.formatDate(row[dIdx], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(row[dIdx]).slice(0, 10);
        return d === String(data.data) &&
          String(row[hIdx]).trim() === String(data.horario).trim() &&
          String(row[sIdx]) !== 'Cancelado';
      });
      if (taken) return json({ ok: false, erro: 'Horário já reservado' });
    }
    sheet.appendRow([
      data.data || '',
      data.horario || '',
      data.nome || '',
      data.whatsapp || '',
      data.cidade || '',
      data.endereco || '',
      data.observacoes || '',
      'Agendado'
    ]);

    try {
      const [ano, mes, dia] = (data.data || '').split('-').map(Number);
      const [hora, min]     = (data.horario || '00:00').split(':').map(Number);
      const inicio = new Date(ano, mes - 1, dia, hora, min, 0);
      const fim    = new Date(inicio.getTime() + 2 * 60 * 60 * 1000);
      const desc   = [
        `Cliente: ${data.nome || ''}`,
        `WhatsApp: ${data.whatsapp || ''}`,
        `Cidade: ${data.cidade || ''}`,
        data.endereco   ? `Endereço: ${data.endereco}`     : '',
        data.observacoes ? `Obs: ${data.observacoes}`      : ''
      ].filter(Boolean).join('\n');
      const cals = CalendarApp.getCalendarsByName('Agenda Bez Clean');
      const cal  = cals.length ? cals[0] : CalendarApp.getDefaultCalendar();
      cal.createEvent(
        `🧽 ${data.nome} — Bez Clean`,
        inicio, fim,
        { description: desc, location: data.endereco || '' }
      );
    } catch(e) { Logger.log('Calendar error: ' + e); }

    return json({ ok: true });
  }

  let sheet = ss.getSheetByName("Leads");
  if (!sheet) {
    sheet = ss.insertSheet("Leads");
    sheet.appendRow(["data", "nome", "whatsapp", "cidade", "endereco", "itens", "valor", "condicao", "acao", "observacoes", "origem", "status"]);
  }
  sheet.appendRow([
    data.data || new Date(),
    data.nome || "",
    data.whatsapp || "",
    data.cidade || "",
    data.endereco || "",
    data.itens || "",
    data.valor || "",
    data.condicao || "",
    data.acao || "",
    data.observacoes || "",
    data.origem || "site",
    data.status || "Novo"
  ]);
  return json({ ok: true });
}

function enviarLembretes() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Agendamentos");
  if (!sheet || sheet.getLastRow() < 2) return;

  const amanha = new Date();
  amanha.setDate(amanha.getDate() + 1);
  const amanhaStr = Utilities.formatDate(amanha, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const rows    = sheet.getDataRange().getValues();
  const headers = rows.shift();
  const dIdx  = headers.indexOf('data');
  const hIdx  = headers.indexOf('horario');
  const nIdx  = headers.indexOf('nome');
  const wIdx  = headers.indexOf('whatsapp');
  const eIdx  = headers.indexOf('endereco');
  const sIdx  = headers.indexOf('status');

  rows.forEach(row => {
    const rawData = row[dIdx];
    const dataStr = rawData instanceof Date
      ? Utilities.formatDate(rawData, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(rawData).slice(0, 10);
    const status  = String(row[sIdx] || '');
    if (dataStr !== amanhaStr || status === 'Cancelado') return;

    const nome    = row[nIdx] || '';
    const wpp     = '55' + String(row[wIdx] || '').replace(/\D/g, '').replace(/^55/, '');
    const horario = row[hIdx] || '';
    const endereco = row[eIdx] || '';

    const texto = `Olá, ${nome}! 👋\n\nPassando para lembrar que *amanhã* você tem um serviço agendado com a *Bez Clean* 🧽\n\n🕐 Horário: *${horario}*${endereco ? `\n📍 Local: ${endereco}` : ''}\n\nQualquer dúvida é só chamar! 😊`;

    UrlFetchApp.fetch('https://evo.ezstudio.com.br/send/text', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'apikey': '3906814c69adeaccda3d394a87fe0b49' },
      payload: JSON.stringify({ number: wpp, text: texto }),
      muteHttpExceptions: true
    });

    Logger.log('Lembrete enviado para ' + nome + ' (' + wpp + ')');
  });
}

function criarTriggerLembrete() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'enviarLembretes')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('enviarLembretes')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('Trigger criado: enviarLembretes todo dia às 8h');
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
