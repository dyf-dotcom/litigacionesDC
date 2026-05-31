const SHEET_NAME = 'Causas';

function doGet(e) {
  const result = handleRequest(e);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const result = handleRequest(e);
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(e) {
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) return { ok: false, error: 'No se encontró la hoja "Causas"' };

    if (action === 'actualizarEtapa1') {
      return actualizarEtapa1(sheet, e.parameter);
    } else if (action === 'listar') {
      return listarCausas(sheet);
    } else if (action === 'buscar') {
      return buscarCausa(sheet, e.parameter.nro);
    } else if (action === 'crearEtapa1') {
      return crearEtapa1(sheet, e.parameter);
    } else if (action === 'actualizarEtapa2') {
      return actualizarEtapa2(sheet, e.parameter);
    } else if (action === 'actualizarEtapa3') {
      return actualizarEtapa3(sheet, e.parameter);
    } else {
      return { ok: false, error: 'Acción no reconocida: ' + action };
    }
  } catch(err) {
    return { ok: false, error: err.toString() };
  }
}

function listarCausas(sheet) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nroIdx = headers.indexOf('NroCausa');
  const carIdx = headers.indexOf('Caratula');
  const etapaIdx = headers.indexOf('EtapaProceso');
  const deptoIdx = headers.indexOf('DeptoJudicial');
  const matIdx = headers.indexOf('Materia');
  const defIdx = headers.indexOf('Defensoria');

  const nroCasIdx = headers.indexOf('NroCausaCasacion');
  const salaIdx = headers.indexOf('SalaTCP');
  const respIdx = headers.indexOf('ResponsableTCP');
  const tipoSentIdx = headers.indexOf('TipoSentenciaCasacion');
  const sentCasIdx = headers.indexOf('SentenciaCasación');

  var causas = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][nroIdx]) {
      causas.push({
        NroCausa: data[i][nroIdx],
        Caratula: data[i][carIdx],
        EtapaProceso: data[i][etapaIdx] || 'Instancia',
        DeptoJudicial: data[i][deptoIdx],
        Materia: data[i][matIdx],
        Defensoria: data[i][defIdx],
        NroCausaCasacion: data[i][nroCasIdx] || '',
        SalaTCP: data[i][salaIdx] || '',
        ResponsableTCP: data[i][respIdx] || '',
        TipoSentenciaCasacion: data[i][tipoSentIdx] || '',
        SentenciaCasacion: data[i][sentCasIdx] || ''
      });
    }
  }
  return { ok: true, causas: causas };
}

function buscarCausa(sheet, nro) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nroIdx = headers.indexOf('NroCausa');
  const casIdx = headers.indexOf('NroCausaCasacion');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nroIdx]) === String(nro) ||
        String(data[i][casIdx]) === String(nro)) {
      var row = {};
      headers.forEach(function(h, j) {
        // Convertir fechas y números a string para evitar errores
        var val = data[i][j];
        if (val instanceof Date) val = val.toLocaleDateString('es-AR');
        row[h] = val !== null && val !== undefined ? String(val) : '';
      });
      return { ok: true, fila: i + 1, data: row };
    }
  }
  return { ok: false, error: 'Causa no encontrada' };
}

function crearEtapa1(sheet, p) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = new Array(headers.length).fill('');

  const map = {
    'Marcatemporal': new Date(),
    'DeptoJudicial': p.DeptoJudicial,
    'Defensoria': p.Defensoria,
    'CamaraApelacion': p.CamaraApelacion,
    'NroCausa': p.NroCausa,
    'Caratula': p.Caratula,
    'Materia': p.Materia,
    'Delito': p.Delito,
    'TipoCoercion': p.TipoCoercion,
    'MotivoRecursoApelacion': p.MotivoRecursoApelacion,
    'FundamentoSentencia': p.FundamentoSentencia,
    'AgravioRecursoCasacion': p.AgravioRecursoCasacion,
    'EtapaProceso': 'Instancia'
  };

  headers.forEach(function(h, i) {
    if (map[h] !== undefined) row[i] = map[h];
  });

  sheet.appendRow(row);
  return { ok: true, mensaje: 'Causa registrada correctamente' };
}


function actualizarEtapa1(sheet, p) {
  const found = buscarCausa(sheet, p.NroCausa);
  if (!found.ok) return found;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fila = found.fila;

  const campos = {
    'Caratula': p.Caratula,
    'Defensoria': p.Defensoria,
    'DeptoJudicial': p.DeptoJudicial,
    'CamaraApelacion': p.CamaraApelacion,
    'Materia': p.Materia,
    'Delito': p.Delito,
    'TipoCoercion': p.TipoCoercion,
    'MotivoRecursoApelacion': p.MotivoRecursoApelacion,
    'FundamentoSentencia': p.FundamentoSentencia,
    'AgravioRecursoCasacion': p.AgravioRecursoCasacion
  };

  Object.keys(campos).forEach(function(campo) {
    const idx = headers.indexOf(campo);
    if (idx >= 0 && campos[campo] !== undefined) {
      sheet.getRange(fila, idx + 1).setValue(campos[campo]);
    }
  });

  return { ok: true, mensaje: 'Causa actualizada correctamente' };
}

function actualizarEtapa2(sheet, p) {
  const found = buscarCausa(sheet, p.NroCausa);
  if (!found.ok) return found;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fila = found.fila;

  const campos = {
    'NroCausaCasacion': p.NroCausaCasacion,
    'SalaTCP': p.SalaTCP,
    'ResponsableTCP': p.ResponsableTCP,
    'TipoSentenciaCasacion': p.TipoSentenciaCasacion,
    'SentenciaCasación': p.SentenciaCasacion,
    'EtapaProceso': 'Casación'
  };

  Object.keys(campos).forEach(function(campo) {
    const idx = headers.indexOf(campo);
    if (idx >= 0 && campos[campo] !== undefined) {
      sheet.getRange(fila, idx + 1).setValue(campos[campo]);
    }
  });

  return { ok: true, mensaje: 'Causa actualizada correctamente' };
}

function actualizarEtapa3(sheet, p) {
  const found = buscarCausa(sheet, p.NroCausaCasacion);
  if (!found.ok) return found;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fila = found.fila;

  const campos = {
    'NroSCBA': p.NroSCBA,
    'ResponsableSCBA': p.ResponsableSCBA,
    'RecursoExtraordinario': p.RecursoExtraordinario,
    'RecursoFederal': p.RecursoFederal,
    'SentenciaSCBA': p.SentenciaSCBA,
    'ResultadoFinal': p.ResultadoFinal,
    'EtapaProceso': 'Corte/Federal'
  };

  Object.keys(campos).forEach(function(campo) {
    const idx = headers.indexOf(campo);
    if (idx >= 0 && campos[campo] !== undefined) {
      sheet.getRange(fila, idx + 1).setValue(campos[campo]);
    }
  });

  return { ok: true, mensaje: 'Causa cerrada correctamente' };
}
