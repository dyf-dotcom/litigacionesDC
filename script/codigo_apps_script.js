const SHEET_NAME = 'Causas';
const LISTAS_NAME = 'Listas';

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

    if (action === 'listas') {
      return listarListas(ss);
    }

    if (!sheet) return { ok: false, error: 'No se encontró la hoja "Causas"' };

    if (action === 'listar')           return listarCausas(sheet);
    if (action === 'buscar')           return buscarCausa(sheet, e.parameter.nro);
    if (action === 'crearEtapa1')      return crearEtapa1(sheet, e.parameter);
    if (action === 'actualizarEtapa1') return actualizarEtapa1(sheet, e.parameter);
    if (action === 'borrarCausa')      return borrarCausa(sheet, e.parameter.NroCausa);
    if (action === 'actualizarEtapa2') return actualizarEtapa2(sheet, e.parameter);
    if (action === 'actualizarEtapa3') return actualizarEtapa3(sheet, e.parameter);

    return { ok: false, error: 'Acción no reconocida: ' + action };
  } catch(err) {
    return { ok: false, error: err.toString() };
  }
}

// Lee la hoja Listas y devuelve cada columna como array
function listarListas(ss) {
  const sheet = ss.getSheetByName(LISTAS_NAME);
  if (!sheet) return { ok: false, error: 'No se encontró la hoja "Listas"' };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const listas = {};

  headers.forEach(function(h, col) {
    if (!h) return;
    listas[h] = [];
    for (var row = 1; row < data.length; row++) {
      if (data[row][col] !== '' && data[row][col] !== null) {
        listas[h].push(String(data[row][col]));
      }
    }
  });

  return { ok: true, listas: listas };
}

function listarCausas(sheet) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'causas_all';

  // Intentar desde caché (válido por 30 minutos)
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  function idx(name) { return headers.indexOf(name); }

  var causas = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][idx('NroCausa')]) continue;
    if (data[i][idx('Borrado')]) continue; // ignorar filas marcadas como borradas
    causas.push({
      NroCausa:          data[i][idx('NroCausa')],
      EtapaProceso:      data[i][idx('EtapaProceso')] || 'Instancia',
      DeptoJudicial:     data[i][idx('DeptoJudicial')],
      Materia:           data[i][idx('Materia')],
      Defensoria:        data[i][idx('Defensoria')],
      NroCausaCasacion:  data[i][idx('NroCausaCasacion')] || '',
      SalaTCP:           data[i][idx('SalaTCP')] || '',
      ResponsableTCP:    data[i][idx('ResponsableTCP')] || '',
      TipoSentenciaCasacion: data[i][idx('TipoSentenciaCasacion')] || '',
      SentenciaCasacion: data[i][idx('SentenciaCasación')] || ''
    });
  }

  const resultado = { ok: true, causas: causas };

  // Guardar en caché por 30 minutos
  try {
    cache.put(cacheKey, JSON.stringify(resultado), 1800);
  } catch(e) {
    // Ignorar si el payload supera el límite de caché (100KB)
  }

  return resultado;
}

function invalidarCacheCausas() {
  CacheService.getScriptCache().remove('causas_all');
}

// Devuelve todos los campos de una fila como objeto string
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
        var val = data[i][j];
        if (val instanceof Date) val = val.toLocaleDateString('es-AR');
        row[h] = (val !== null && val !== undefined) ? String(val) : '';
      });
      return { ok: true, fila: i + 1, data: row };
    }
  }
  return { ok: false, error: 'Causa no encontrada' };
}

function borrarCausa(sheet, nro) {
  const found = buscarCausa(sheet, nro);
  if (!found.ok) return found;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.indexOf('Borrado');
  if (idx < 0) return { ok: false, error: 'Columna Borrado no encontrada' };

  sheet.getRange(found.fila, idx + 1).setValue(true);
  invalidarCacheCausas();
  return { ok: true, mensaje: 'Causa eliminada correctamente' };
}

function crearEtapa1(sheet, p) {
  // Validar duplicado: mismo NroCausa en el mismo DeptoJudicial
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nroIdx = headers.indexOf('NroCausa');
  const deptoIdx = headers.indexOf('DeptoJudicial');
  const idIdx = headers.indexOf('ID');

  // Calcular próximo ID auto-incremental
  var maxId = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][nroIdx]) === String(p.NroCausa) &&
        String(data[i][deptoIdx]) === String(p.DeptoJudicial)) {
      return { ok: false, error: 'La causa ' + p.NroCausa + ' ya existe en ' + p.DeptoJudicial };
    }
    var idVal = parseInt(data[i][idIdx]);
    if (!isNaN(idVal) && idVal > maxId) maxId = idVal;
  }

  var row = new Array(headers.length).fill('');

  const map = {
    'ID':                   maxId + 1,
    'DeptoJudicial':        p.DeptoJudicial,
    'Defensoria':           p.Defensoria,
    'CamaraApelacion':      p.CamaraApelacion,
    'IPP/PP':               p.IPPPP,
    'NroCausa':             p.NroCausa,
    'Materia':              p.Materia,
    'Delito':               p.Delito,
    'TipoCoercion':         p.TipoCoercion,
    'MotivoApelacion':      p.MotivoApelacion,
    'SentenciaCamara':      p.SentenciaCamara,
    'FundamentoSentencia':  p.FundamentoSentencia,
    'MotivoRecursoApelacion': p.MotivoRecursoApelacion,
    'AgravioRecursoCasacion': p.AgravioRecursoCasacion,
    'Observaciones':        p.Observaciones,
    'EtapaProceso':         'Instancia'
  };

  headers.forEach(function(h, i) {
    if (map[h] !== undefined) row[i] = map[h];
  });

  sheet.appendRow(row);
  invalidarCacheCausas();
  return { ok: true, mensaje: 'Causa registrada correctamente' };
}

function actualizarEtapa1(sheet, p) {
  const found = buscarCausa(sheet, p.NroCausa);
  if (!found.ok) return found;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fila = found.fila;

  const campos = {
    'Defensoria':             p.Defensoria,
    'DeptoJudicial':          p.DeptoJudicial,
    'CamaraApelacion':        p.CamaraApelacion,
    'IPP/PP':                 p.IPPPP,
    'Materia':                p.Materia,
    'Delito':                 p.Delito,
    'TipoCoercion':           p.TipoCoercion,
    'MotivoApelacion':        p.MotivoApelacion,
    'SentenciaCamara':        p.SentenciaCamara,
    'FundamentoSentencia':    p.FundamentoSentencia,
    'MotivoRecursoApelacion': p.MotivoRecursoApelacion,
    'AgravioRecursoCasacion': p.AgravioRecursoCasacion,
    'Observaciones':          p.Observaciones
  };

  Object.keys(campos).forEach(function(campo) {
    const idx = headers.indexOf(campo);
    if (idx >= 0 && campos[campo] !== undefined) {
      sheet.getRange(fila, idx + 1).setValue(campos[campo]);
    }
  });

  invalidarCacheCausas();
  return { ok: true, mensaje: 'Causa actualizada correctamente' };
}

function actualizarEtapa2(sheet, p) {
  const found = buscarCausa(sheet, p.NroCausa);
  if (!found.ok) return found;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fila = found.fila;

  const campos = {
    'NroCausaCasacion':      p.NroCausaCasacion,
    'SalaTCP':               p.SalaTCP,
    'ResponsableTCP':        p.ResponsableTCP,
    'TipoSentenciaCasacion': p.TipoSentenciaCasacion,
    'SentenciaCasación':     p.SentenciaCasacion,
    'EtapaProceso':          'Casación'
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
    'NroSCBA':               p.NroSCBA,
    'ResponsableSCBA':       p.ResponsableSCBA,
    'RecursoExtraordinario': p.RecursoExtraordinario,
    'RecursoFederal':        p.RecursoFederal,
    'SentenciaSCBA':         p.SentenciaSCBA,
    'ResultadoFinal':        p.ResultadoFinal,
    'EtapaProceso':          'Corte/Federal'
  };

  Object.keys(campos).forEach(function(campo) {
    const idx = headers.indexOf(campo);
    if (idx >= 0 && campos[campo] !== undefined) {
      sheet.getRange(fila, idx + 1).setValue(campos[campo]);
    }
  });

  return { ok: true, mensaje: 'Causa cerrada correctamente' };
}
