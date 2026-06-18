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
        const d = new Date(row[dataIdx]);
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
      .filter(row => String(row[dataIdx]).slice(0, 10) === data && String(row[statusIdx] || '') !== 'Cancelado')
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

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
