function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Painel Expedições')
    .addItem('Abrir Painel', 'abrirPainel')
    .addToUi();
}

function abrirPainel() {
  const html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setWidth(1440)
    .setHeight(880);
  SpreadsheetApp.getUi().showModalDialog(html, '8808-Painel Expedições');
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('8808-Painel Expedições')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getExpedicoesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const all = [];
  const abasComFiltro = ['report', 'departed', 'expedicoes'];
  const turnosParaLer = ['T1', 'T2', 'T3'];

  abasComFiltro.forEach(function (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const isDeparted = sheetName.toLowerCase().indexOf('departed') !== -1;
    const statusRegistro = isDeparted ? 'CONCLUÍDO' : 'PLANEJADO';

    const celulaA1 = sheet.getRange("A1");
    let valorOriginal;
    try {
      valorOriginal = celulaA1.getValue();
    } catch (e) {
      valorOriginal = "";
    }

    turnosParaLer.forEach(function (turnoId) {
      celulaA1.setValue(turnoId);
      SpreadsheetApp.flush();
      Utilities.sleep(400);

      const values = sheet.getDataRange().getValues();
      if (!values.length) return;

      let headerRow = -1;
      for (let i = 0; i < Math.min(values.length, 5); i++) {
        if (values[i].join('|').indexOf('DESTINO CODE') !== -1) {
          headerRow = i;
          break;
        }
      }
      if (headerRow === -1) return;

      for (let i = headerRow + 1; i < values.length; i++) {
        const row = values[i];
        const destino = row[1];
        const code = row[2];
        const eta = row[3];
        const cpt = row[4];

        if (!destino || destino.toString().trim() === "" ||
          destino.toString().toUpperCase().indexOf("TOTAL") !== -1) {
          continue;
        }

        const destinoStr = destino.toString().trim();

        let cat = 'OUTRO';
        if (destinoStr.toLowerCase().indexOf('soc') === 0) cat = 'SOC';
        else if (destinoStr.toLowerCase().indexOf('lm hub') === 0) cat = 'LM HUB';
        else if (destinoStr.toLowerCase().indexOf('xpt') === 0) cat = 'XPT';

        let uf = 'OUTRO';
        const parts = destinoStr.split('_');
        for (let p = 0; p < parts.length; p++) {
          if (parts[p].length === 2 && parts[p] === parts[p].toUpperCase() && /^[A-Z]{2}$/.test(parts[p])) {
            uf = parts[p];
            break;
          }
        }

        all.push({
          turno: turnoId,
          status: statusRegistro,
          destino: destinoStr,
          code: code ? code.toString().trim() : '',
          eta: formatTime_(eta),
          cpt: formatTime_(cpt),
          cat: cat,
          uf: uf
        });
      }
    });

    if (valorOriginal) {
      celulaA1.setValue(valorOriginal);
    }
  });

  return all;
}

function formatTime_(value) {
  if (value === null || value === undefined || value === "") return "00:00:00";

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm:ss');
  }

  if (typeof value === 'number') {
    var totalSeconds = Math.round(value * 24 * 3600);
    var hours = Math.floor(totalSeconds / 3600) % 24;
    var minutes = Math.floor((totalSeconds % 3600) / 60);
    var seconds = totalSeconds % 60;
    return [
      String(hours).padStart(2, '0'),
      String(minutes).padStart(2, '0'),
      String(seconds).padStart(2, '0')
    ].join(':');
  }

  var str = value.toString().trim();
  if (/^\d{1,2}:\d{2}(:\d{2})?$/.test(str)) {
    var parts = str.split(':');
    return [
      parts[0].padStart(2, '0'),
      parts[1].padStart(2, '0'),
      parts[2] ? parts[2].padStart(2, '0') : '00'
    ].join(':');
  }

  return "00:00:00";
}
