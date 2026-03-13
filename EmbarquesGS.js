// ══════════════════════════════════════════════════════════
//  URL DEL TABLERO MÉTRICAS
// ══════════════════════════════════════════════════════════
function getTablerMetricasUrl() {
  try {
    var base = ScriptApp.getService().getUrl();
    return base + "?v=TABLERO";   // Ajusta el parámetro v si tu tablero usa otro
  } catch(e) { return ""; }
}


// ══════════════════════════════════════════════════════════
//  GUARDAR PERMISOS E_* (conserva permisos de otras apps)
// ══════════════════════════════════════════════════════════
function guardarPermisosEmbarques(nombre, permisosFinales) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return "Error: pestaña USUARIOS no encontrada";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === String(nombre).trim().toUpperCase()) {
        // permisosFinales ya viene con el merge correcto desde el HTML
        sheet.getRange(i + 1, 6).setValue(permisosFinales);
        return "OK";
      }
    }
    return "Error: usuario no encontrado";
  } catch(e) { return "Error: " + e.toString(); }
}


// ══════════════════════════════════════════════════════════
//  DASHBOARD EMBARQUES
//  Hojas esperadas en ID_HOJA_EMBARQUES:
//    ENVIADO  → A=ID, B=SEMANA, C=FECHA(Date), D=REMISION, E=KG, F=PEDIDO,
//               G=CODIGO, H=DESCRIPCION, I=FAMILIA, J=FOLIO_ENVIO, K=PIEZAS
//    LOTES    → busca col ESTATUS
//    PEDIDOS  → busca col ESTADO
// ══════════════════════════════════════════════════════════
function getDashboardEmbarques(mes, anio) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);

    // ── Hoja ENVIADO ──
    var shEnv = ss.getSheetByName("ENVIADO");
    var kgMes = 0, lotesMes = 0, kgHoy = 0, kgSemana = 0;
    var kgPorSemana = [0, 0, 0, 0, 0];
    var kgPor12Meses = {};
    var ultimosEnvios = [];

    var hoy = new Date();
    var hoyStr = Utilities.formatDate(hoy, Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Calcular lunes de la semana actual
    var diaSemana = hoy.getDay(); // 0=Dom
    var lunes = new Date(hoy);
    lunes.setDate(hoy.getDate() - (diaSemana === 0 ? 6 : diaSemana - 1));
    lunes.setHours(0,0,0,0);

    // Últimos 12 meses para eje X
    var mesesLabels = [];
    for (var mi = 11; mi >= 0; mi--) {
      var dm = new Date(hoy.getFullYear(), hoy.getMonth() - mi, 1);
      var nombresMes = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
      mesesLabels.push(nombresMes[dm.getMonth()] + ' ' + String(dm.getFullYear()).substr(2));
      var key = dm.getFullYear() + '-' + (dm.getMonth() + 1);
      kgPor12Meses[key] = 0;
    }

    if (shEnv && shEnv.getLastRow() > 1) {
      var dataEnv = shEnv.getDataRange().getValues();
      var header  = dataEnv[0];

      // Detectar columnas por nombre
      var iId=0, iSem=1, iFec=2, iRem=3, iKg=4, iPed=5, iCod=6, iDesc=7, iFam=8, iEnv=9, iPzs=10;
      header.forEach(function(h, i) {
        var hn = String(h).toUpperCase().trim();
        if (hn==='ID')          iId=i;
        if (hn==='SEMANA')      iSem=i;
        if (hn==='FECHA')       iFec=i;
        if (hn==='REMISION')    iRem=i;
        if (hn.includes('KG') && !hn.includes('ENT')) iKg=i;
        if (hn==='PEDIDO')      iPed=i;
        if (hn==='CODIGO')      iCod=i;
        if (hn.includes('DESC')) iDesc=i;
        if (hn.includes('FAM')) iFam=i;
        if (hn.includes('FOLIO') || hn==='ENVIO') iEnv=i;
        if (hn.includes('PIE')) iPzs=i;
      });

      for (var r = 1; r < dataEnv.length; r++) {
        var row  = dataEnv[r];
        var fRaw = row[iFec];
        if (!fRaw) continue;
        var fDate = (fRaw instanceof Date) ? fRaw : new Date(fRaw);
        if (isNaN(fDate.getTime())) continue;

        var rMes  = fDate.getMonth() + 1;
        var rAnio = fDate.getFullYear();
        var kg    = Number(row[iKg]) || 0;

        // KG mes seleccionado
        if (rMes === mes && rAnio === anio) {
          kgMes += kg;
          lotesMes++;
          // KG por semana del mes
          var diaMes = fDate.getDate();
          var sem = Math.min(Math.ceil(diaMes / 7), 5) - 1;
          if (sem >= 0 && sem < 5) kgPorSemana[sem] += kg;
          // Últimos envíos
          ultimosEnvios.push({
            fecha:       Utilities.formatDate(fDate, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            remision:    String(row[iRem] || ""),
            pedido:      String(row[iPed] || ""),
            codigo:      String(row[iCod] || ""),
            descripcion: String(row[iDesc] || ""),
            kg:          kg,
            piezas:      Number(row[iPzs]) || 0,
            envio:       String(row[iEnv] || "")
          });
        }

        // KG hoy
        var fStr = Utilities.formatDate(fDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (fStr === hoyStr) kgHoy += kg;

        // KG semana actual
        if (fDate >= lunes && fDate <= hoy) kgSemana += kg;

        // KG últimos 12 meses
        var k12 = rAnio + '-' + rMes;
        if (kgPor12Meses.hasOwnProperty(k12)) kgPor12Meses[k12] += kg;
      }
    }

    // Ordenar últimos envíos por fecha desc, tomar 15
    ultimosEnvios.sort(function(a, b) {
      var da = a.fecha.split('/').reverse().join('-');
      var db = b.fecha.split('/').reverse().join('-');
      return db > da ? 1 : -1;
    });
    ultimosEnvios = ultimosEnvios.slice(0, 15);

    var kgPorMes = Object.keys(kgPor12Meses)
      .sort(function(a, b) {
        var pa = a.split('-').map(Number);
        var pb = b.split('-').map(Number);
        return (pa[0]*12+pa[1]) - (pb[0]*12+pb[1]);
      })
      .map(function(k) { return kgPor12Meses[k]; });

    // ── Pedidos activos ──
    var pedidosActivos = 0;
    try {
      var shPed = ss.getSheetByName("PEDIDOS");
      if (shPed && shPed.getLastRow() > 1) {
        var dataPed = shPed.getDataRange().getValues();
        var hPed = dataPed[0];
        var iEst = -1;
        hPed.forEach(function(h, i) {
          if (String(h).toUpperCase().includes('ESTADO') || String(h).toUpperCase().includes('STATUS')) iEst = i;
        });
        for (var rp = 1; rp < dataPed.length; rp++) {
          if (!dataPed[rp][0]) continue;
          var est = iEst >= 0 ? String(dataPed[rp][iEst]).toUpperCase().trim() : '';
          if (est !== 'ENTREGADO' && est !== 'CANCELADO') pedidosActivos++;
        }
      }
    } catch(ep) {}

    // ── Lotes en cola ──
    var lotesEnCola = 0;
    try {
      var shLot = ss.getSheetByName("LOTES");
      if (shLot && shLot.getLastRow() > 1) {
        var dataLot = shLot.getDataRange().getValues();
        var hLot = dataLot[0];
        var iLotEst = -1;
        hLot.forEach(function(h, i) {
          if (String(h).toUpperCase().includes('ESTATUS') || String(h).toUpperCase().includes('STATUS')) iLotEst = i;
        });
        for (var rl = 1; rl < dataLot.length; rl++) {
          if (!dataLot[rl][0]) continue;
          var lotEst = iLotEst >= 0 ? String(dataLot[rl][iLotEst]).toUpperCase().trim() : '';
          if (lotEst === 'IMPRIMIR') lotesEnCola++;
        }
      }
    } catch(el) {}

    return {
      kgMes:          kgMes,
      lotesMes:       lotesMes,
      kgHoy:          kgHoy,
      kgSemana:       kgSemana,
      pedidosActivos: pedidosActivos,
      lotesEnCola:    lotesEnCola,
      ultimosEnvios:  ultimosEnvios,
      kgPorSemana:    kgPorSemana,
      semanas:        ['S1','S2','S3','S4','S5'],
      mesesLabels:    mesesLabels,
      kgPorMes:       kgPorMes,
      deltaKgMes:     null,
      deltaLotes:     null,
      deltaSemana:    null
    };

  } catch(e) {
    Logger.log("getDashboardEmbarques error: " + e.toString());
    return { error: e.toString() };
  }
}


// ══════════════════════════════════════════════════════════
//  ENVÍOS — getData para loadEnvios() en el HTML
// ══════════════════════════════════════════════════════════
function getEnviosMes() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);
    var sh = ss.getSheetByName("ENVIADO");
    if (!sh || sh.getLastRow() < 2) return [];
    var data   = sh.getDataRange().getValues();
    var header = data[0];
    var tz = Session.getScriptTimeZone();

    return data.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      var obj = {};
      header.forEach(function(h, i) { obj[String(h).trim()] = r[i]; });
      // Normalizar fecha
      if (obj.FECHA instanceof Date) obj.FECHA = Utilities.formatDate(obj.FECHA, tz, "dd/MM/yyyy");
      return obj;
    });
  } catch(e) { return []; }
}

var ID_FOLDER_ENVIOS_PDF = '1MrFvIrPOG7my0pQBpKntXjOIlu5i3n9K';
var ID_HOJA_ENVIADO_PDF  = '1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk';

// ── Genera un ID único tipo "ENV-YYYYMMDD-XXXX" ──
function _generarIdEnvio() {
  var d    = new Date();
  var f    = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyyMMdd');
  var rand = Math.random().toString(36).substr(2, 4).toUpperCase();
  return 'ENV-' + f + '-' + rand;
}

// ── Convierte "dd/MM/yyyy" → semana del año tipo "S52" ──
function _getSemana(fechaStr) {
  var p = String(fechaStr).split('/');
  if (p.length < 3) return '';
  var d = new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 3 - (d.getDay() + 6) % 7);
  var s1  = new Date(d.getFullYear(), 0, 4);
  var sem = 1 + Math.round(((d - s1) / 86400000 - 3 + (s1.getDay() + 6) % 7) / 7);
  return 'S' + sem;
}

// ── Verifica si el número de ENVIO ya existe en col M (posición 13) ──
function verificarEnvioDuplicado(numEnvio) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ENVIADO_PDF);
    var sheet = ss.getSheetByName('ENVIADO');
    if (!sheet || sheet.getLastRow() < 2) return false;
    var data = sheet.getRange(2, 13, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(numEnvio).trim()) return true;
    }
    return false;
  } catch(e) {
    throw new Error('Error verificando duplicado: ' + e.message);
  }
}

// ── Sube el PDF a Drive y devuelve la URL pública ──
function subirPDFaDrive(base64Data, nombreArchivo) {
  try {
    var folder  = DriveApp.getFolderById(ID_FOLDER_ENVIOS_PDF);
    var blob    = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', nombreArchivo);
    var archivo = folder.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch(e) {
    throw new Error('Error subiendo PDF a Drive: ' + e.message);
  }
}

// ── Guarda las filas en la hoja ENVIADO ──
function guardarFilasPDF(filas, urlPdf) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ENVIADO_PDF);
    var sheet = ss.getSheetByName('ENVIADO');
    if (!sheet) throw new Error('No se encontró la hoja ENVIADO');
    filas.forEach(function(f) {
      var id  = _generarIdEnvio();
      var sem = _getSemana(f.fecha);
      sheet.appendRow([
        id,                   // A: ID
        sem,                  // B: SEM
        f.fecha       || '',  // C: FECHA (dd/MM/yyyy)
        f.remision    || '',  // D: REMISION
        f.kgEntregados|| '',  // E: KG_ENTREGADOS
        f.pedido      || '',  // F: PEDIDO
        f.codigo      || '',  // G: CODIGO
        f.descripcion || '',  // H: DESCRIPCION
        f.familia     || '',  // I: FAMILIA
        f.kilos       || 0,   // J: KILOS
        f.piezas      || 0,   // K: PIEZAS
        f.comentario  || '',  // L: COMENTARIO
        f.envio       || '',  // M: ENVIO
        urlPdf                // N: URL
      ]);
    });
    return 'OK';
  } catch(e) {
    throw new Error('Error guardando filas: ' + e.message);
  }
}


// ══════════════════════════════════════════════════════════
//  LOTES DISPONIBLES — para el módulo Despachos
// ══════════════════════════════════════════════════════════
function obtenerLotesDisponibles() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);
    var sh = ss.getSheetByName("LOTES");
    if (!sh || sh.getLastRow() < 2) return { lotes: [], cola: [] };
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var iLote=-1, iPed=-1, iCod=-1, iDesc=-1, iEst=-1, iKg=-1;
    h.forEach(function(col, i) {
      var c = String(col).toUpperCase().trim();
      if (c==='LOTE' || c==='N_LOTE' || c==='FOLIO_LOTE') iLote=i;
      if (c==='PEDIDO') iPed=i;
      if (c==='CODIGO') iCod=i;
      if (c.includes('DESC')) iDesc=i;
      if (c.includes('ESTATUS') || c.includes('STATUS')) iEst=i;
      if (c.includes('KG') || c.includes('PESO')) iKg=i;
    });
    var lotes = [];
    var cola  = [];
    data.slice(1).forEach(function(r) {
      if (!r[Math.max(iLote,0)]) return;
      var obj = {
        lote:    String(r[Math.max(iLote,0)]||''),
        pedido:  String(r[Math.max(iPed,0)]||''),
        codigo:  String(r[Math.max(iCod,0)]||''),
        desc:    String(r[Math.max(iDesc,0)]||''),
        estatus: String(r[Math.max(iEst,0)]||''),
        kg:      Number(r[Math.max(iKg,0)])||0
      };
      lotes.push(obj);
      if (obj.estatus.toUpperCase() === 'IMPRIMIR') cola.push(obj);
    });
    return { lotes: lotes, cola: cola };
  } catch(e) { return { lotes: [], cola: [] }; }
}


// ══════════════════════════════════════════════════════════
//  MARCAR LOTES COMO IMPRESOS
// ══════════════════════════════════════════════════════════
function marcarLotesComoImpresos() {
  // Esta función actúa sobre los lotes que tienen estatus IMPRIMIR
  // Cámbiala según tu lógica de negocio (p.ej. cambiar a IMPRESO, generar registro, etc.)
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);
    var sh = ss.getSheetByName("LOTES");
    if (!sh) return "Error: hoja LOTES no encontrada";
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var iEst = -1;
    h.forEach(function(c, i) {
      if (String(c).toUpperCase().includes('ESTATUS') || String(c).toUpperCase().includes('STATUS')) iEst = i;
    });
    if (iEst < 0) return "Error: columna ESTATUS no encontrada";
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][iEst]).toUpperCase().trim() === 'IMPRIMIR') {
        sh.getRange(i + 1, iEst + 1).setValue('IMPRESO');
        count++;
      }
    }
    return "Se marcaron " + count + " lotes como IMPRESO";
  } catch(e) { return "Error: " + e.toString(); }
}


// ══════════════════════════════════════════════════════════
//  EXISTENCIA ALAMBRÓN
// ══════════════════════════════════════════════════════════
function getExistenciaAlambron() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);
    // Prueba con hojas comunes: EXISTENCIA, ALAMBRON, MP_STOCK
    var nombres = ['EXISTENCIA', 'ALAMBRON', 'MP_STOCK', 'STOCK_MP', 'INVENTARIO_MP'];
    var sh = null;
    for (var n = 0; n < nombres.length; n++) {
      sh = ss.getSheetByName(nombres[n]);
      if (sh) break;
    }
    if (!sh || sh.getLastRow() < 2) return [];
    var data   = sh.getDataRange().getValues();
    var header = data[0];
    return data.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      var obj = {};
      header.forEach(function(h, i) { obj[String(h).trim()] = r[i]; });
      return obj;
    });
  } catch(e) { return []; }
}


// ══════════════════════════════════════════════════════════
//  GESTIÓN MP
// ══════════════════════════════════════════════════════════
function getGestionMP() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);
    var shEnt = ss.getSheetByName("ENTRADAS_MP") || ss.getSheetByName("ENTRADAS") || null;
    var shSal = ss.getSheetByName("SALIDAS_MP")  || ss.getSheetByName("SALIDAS")  || null;
    var tz = Session.getScriptTimeZone();

    var entradas = [];
    if (shEnt && shEnt.getLastRow() > 1) {
      var dEnt = shEnt.getDataRange().getValues();
      var hEnt = dEnt[0];
      entradas = dEnt.slice(1).filter(function(r){return r[0];}).map(function(r){
        var o={};
        hEnt.forEach(function(h,i){ o[String(h).trim()]=r[i]; });
        if (o.FECHA instanceof Date) o.FECHA = Utilities.formatDate(o.FECHA, tz, "dd/MM/yyyy");
        // compatibilidad de nombres de campo
        o.n_rollo = o.N_ROLLO || o.ROLLO || o[hEnt[0]] || '';
        o.sello   = o.SELLO   || o.MARCA || '';
        o.diametro= o.DIAMETRO|| o.DIA   || '';
        o.acero   = o.ACERO   || o.TIPO  || '';
        o.kilos   = Number(o.KG || o.KILOS || o.PESO || 0);
        return o;
      });
    }

    var salidas = [];
    if (shSal && shSal.getLastRow() > 1) {
      var dSal = shSal.getDataRange().getValues();
      var hSal = dSal[0];
      salidas = dSal.slice(1).filter(function(r){return r[0];}).map(function(r){
        var o={};
        hSal.forEach(function(h,i){ o[String(h).trim()]=r[i]; });
        if (o.FECHA instanceof Date) o.FECHA = Utilities.formatDate(o.FECHA, tz, "dd/MM/yyyy");
        o.codigo     = o.CODIGO      || '';
        o.descripcion= o.DESCRIPCION || o.DESC || '';
        o.kg         = Number(o.KG || o.KILOS || 0);
        o.fecha      = o.FECHA || '';
        return o;
      });
    }

    return { entradas: entradas, salidas: salidas };
  } catch(e) { return { entradas: [], salidas: [] }; }
}

function editarEntradaMP(id, nuevosValores, listaCambios, nombreUsuario) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("ENTRADAS_MP");
    var data  = sheet.getDataRange().getValues();

    // ── 1. Buscar la fila por ID (Columna A = índice 0) ──────────────
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(id).trim()) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) {
      throw new Error("No se encontró el registro con ID: " + id);
    }

    var sheetRow = rowIndex + 1; // En Sheets las filas empiezan en 1

    // ── 2. Actualizar los campos en la hoja ──────────────────────────
    // Col D(4)=DIAMETRO, Col E(5)=ACERO, Col F(6)=N_ROLLO
    // Col G(7)=KILOS,    Col H(8)=COLADA
    sheet.getRange(sheetRow, 4).setValue(nuevosValores.diametro);
    sheet.getRange(sheetRow, 5).setValue(nuevosValores.acero);
    sheet.getRange(sheetRow, 6).setValue(nuevosValores.n_rollo);
    sheet.getRange(sheetRow, 7).setValue(parseFloat(nuevosValores.kilos) || 0);
    sheet.getRange(sheetRow, 8).setValue(nuevosValores.colada);

    // ── 3. El nombre del usuario viene directo del login del HTML ─────
    // currentUser.nombre ya tiene el nombre (Col B de USUARIOS)
    var usuario = String(nombreUsuario || "DESCONOCIDO").trim().toUpperCase();

    // ── 4. Construir la línea nueva del historial ─────────────────────
    // Formato: 02/03/2026 20:05:03_ROBERTO DIAZ_Cambió DIAMETRO de 5.50 a 5.60 mm
    var zona       = "GMT-6";
    var ahora      = Utilities.formatDate(new Date(), zona, "dd/MM/yyyy HH:mm:ss");
    var textoNuevo = ahora + "_" + usuario + "_" + listaCambios.join(", ");

    // ── 5. Leer historial anterior en Col R (índice 17 = columna 18) ─
    var historialAnterior = String(data[rowIndex][17] || "").trim();

    // Acumular sin borrar — cada cambio va en línea nueva
    var historialFinal = historialAnterior
      ? historialAnterior + "\n" + textoNuevo
      : textoNuevo;

    sheet.getRange(sheetRow, 18).setValue(historialFinal); // Col R

    Logger.log("editarEntradaMP OK — ID: " + id + " | " + textoNuevo);
    return "OK";

  } catch (e) {
    Logger.log("Error en editarEntradaMP: " + e.toString());
    throw e;
  }
}

// ══════════════════════════════════════════════════════════
//  CIRCULANTE
// ══════════════════════════════════════════════════════════
function getCirculante() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_EMBARQUES);
    var nombres = ['CIRCULANTE', 'INV_CIRCULANTE', 'INVENTARIO_CIRCULANTE'];
    var sh = null;
    for (var n = 0; n < nombres.length; n++) {
      sh = ss.getSheetByName(nombres[n]);
      if (sh) break;
    }
    if (!sh || sh.getLastRow() < 2) return [];
    var data   = sh.getDataRange().getValues();
    var header = data[0];
    var tz = Session.getScriptTimeZone();
    return data.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
      var obj = {};
      header.forEach(function(h, i) {
        var v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, "dd/MM/yyyy");
        obj[String(h).trim()] = v;
      });
      return obj;
    });
  } catch(e) { return []; }
}

var ID_SS_PRODUCCION = "1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk";

// ============================================================
// MÓDULO ENVÍOS
// ============================================================

function obtenerDatosEnvios() {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_PRODUCCION);
    var sheetCodigos = ss.getSheetByName("CODIGOS");
    if (!sheetCodigos) return { catalogo: {}, error: "No existe hoja CODIGOS" };
    var dataCod = sheetCodigos.getDataRange().getValues();
    if (dataCod.length === 0) return { catalogo: {} };
    var hCod = dataCod[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var getIdx = function(name) {
      for(var k=0; k<hCod.length; k++) { if(hCod[k] === name) return k; }
      for(var k=0; k<hCod.length; k++) { if(hCod[k].includes(name)) return k; }
      return -1;
    };
    var iCod = getIdx("CODIGO"); if(iCod<0) iCod=0;
    var iDesc = getIdx("DESCRIPCION");
    var iFam = getIdx("FAMILIA"); if(iFam<0) iFam = getIdx("TIPO");
    var iPeso = getIdx("PESO"); if(iPeso<0) iPeso = getIdx("PESO_UNITARIO");
    var mapCodigos = {};
    for(var i=1; i<dataCod.length; i++) {
      var codigo = String(dataCod[i][iCod]).trim();
      var desc = (iDesc > -1) ? String(dataCod[i][iDesc]) : "";
      var fam  = (iFam  > -1) ? String(dataCod[i][iFam]).trim() : "";
      var peso = (iPeso > -1) ? Number(dataCod[i][iPeso]) : 0;
      if(fam === "") fam = "SIN FAMILIA";
      mapCodigos[codigo] = { d: desc, f: fam, p: peso };
    }
    return { catalogo: mapCodigos };
  } catch (e) {
    return { catalogo: {}, error: e.toString() };
  }
}

function buscarEnvios(fechaIni, fechaFin) {
  var ss = SpreadsheetApp.openById(ID_SS_PRODUCCION);
  var sheet = ss.getSheetByName("ENVIADO");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var getIdx = function(n) { return headers.indexOf(n); };
  var IDX = {
    ID: getIdx("ID"), FECHA: getIdx("FECHA"), CODIGO: getIdx("CODIGO"),
    DESC: getIdx("DESCRIPCION"), FAM: getIdx("FAMILIA"), KILOS: getIdx("KILOS"),
    PIEZAS: getIdx("PIEZAS"), COM: getIdx("COMENTARIOS"), ENVIO: getIdx("ENVIO"),
    REMISION: getIdx("REMISION"), PEDIDO: getIdx("PEDIDO")
  };
  var fIni = new Date(fechaIni); fIni.setHours(0,0,0,0);
  var fFin = new Date(fechaFin); fFin.setHours(23,59,59,999);
  var resultado = [];
  for(var i=1; i<data.length; i++) {
    var fechaRow = data[i][IDX.FECHA];
    if (!(fechaRow instanceof Date)) {
      if(typeof fechaRow === 'string' && fechaRow.includes('/')) {
        var p = fechaRow.split('/');
        fechaRow = new Date(p[2], p[1]-1, p[0]);
      } else {
        fechaRow = new Date(fechaRow);
      }
    }
    if (fechaRow >= fIni && fechaRow <= fFin) {
      resultado.push({
        id:          data[i][IDX.ID],
        fecha:       Utilities.formatDate(fechaRow, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        remision:    data[i][IDX.REMISION],
        pedido:      data[i][IDX.PEDIDO],
        codigo:      data[i][IDX.CODIGO],
        desc:        data[i][IDX.DESC],
        familia:     data[i][IDX.FAM],
        kilos:       Number(data[i][IDX.KILOS]),
        piezas:      Number(data[i][IDX.PIEZAS]),
        comentarios: data[i][IDX.COM],
        envio:       data[i][IDX.ENVIO]
      });
    }
  }
  return resultado;
}

function guardarCambiosEnvios(lista) {
  var ss = SpreadsheetApp.openById(ID_SS_PRODUCCION);
  var sheet = ss.getSheetByName("ENVIADO");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var idxID = headers.indexOf("ID");
  var mapaFilas = {};
  for(var i=1; i<data.length; i++) mapaFilas[String(data[i][idxID])] = i + 1;
  var getIdx = function(n) { return headers.indexOf(n); };
  lista.forEach(function(item) {
    var fila = mapaFilas[item.id];
    if(fila) {
      if(item.fecha      !== undefined) sheet.getRange(fila, getIdx("FECHA")+1).setValue(new Date(item.fecha + "T12:00:00"));
      if(item.remision   !== undefined) sheet.getRange(fila, getIdx("REMISION")+1).setValue(item.remision);
      if(item.pedido     !== undefined) sheet.getRange(fila, getIdx("PEDIDO")+1).setValue(item.pedido);
      if(item.codigo     !== undefined) sheet.getRange(fila, getIdx("CODIGO")+1).setValue(item.codigo);
      if(item.desc       !== undefined) sheet.getRange(fila, getIdx("DESCRIPCION")+1).setValue(item.desc);
      if(item.familia    !== undefined) sheet.getRange(fila, getIdx("FAMILIA")+1).setValue(item.familia);
      if(item.kilos      !== undefined) sheet.getRange(fila, getIdx("KILOS")+1).setValue(item.kilos);
      if(item.piezas     !== undefined) sheet.getRange(fila, getIdx("PIEZAS")+1).setValue(item.piezas);
      if(item.comentarios!== undefined) sheet.getRange(fila, getIdx("COMENTARIOS")+1).setValue(item.comentarios);
    }
  });
  return "✅ Cambios guardados.";
}

function guardarNuevosEnvios(listaNuevos) {
  try {
    var ss = SpreadsheetApp.openById(ID_SS_PRODUCCION);
    var sheet = ss.getSheetByName("ENVIADO");
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                    .map(function(h){ return String(h).toUpperCase().trim(); });
    var filas = listaNuevos.map(function(item) {
      var arr = new Array(headers.length).fill("");
      var set = function(k, v) { var i = headers.indexOf(k); if(i > -1) arr[i] = v; };
      set("ID", Utilities.getUuid());
      var d = new Date();
      if(item.fecha) {
        var partes = String(item.fecha).split("-");
        if(partes.length === 3) d = new Date(partes[0], partes[1]-1, partes[2], 12, 0, 0);
        else d = new Date(item.fecha);
      }
      set("FECHA", d);
      set("REMISION",    item.remision);
      set("PEDIDO",      item.pedido);
      set("CODIGO",      "'" + item.codigo);
      set("DESCRIPCION", item.desc);
      set("FAMILIA",     item.familia);
      set("KILOS",       item.kilos);
      set("PIEZAS",      item.piezas);
      set("COMENTARIOS", item.comentarios);
      set("ENVIO",       item.envio || ("IMP-" + new Date().getTime()));
      return arr;
    });
    if(filas.length > 0) {
      sheet.getRange(sheet.getLastRow()+1, 1, filas.length, filas[0].length).setValues(filas);
    }
    return "✅ Se importaron " + filas.length + " registros correctamente.";
  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}

function generarReporteEnvios(mesesSeleccionados) {
  var ss = SpreadsheetApp.openById(ID_SS_PRODUCCION);
  var sheet = ss.getSheetByName("ENVIADO");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX = {
    FECHA:   headers.indexOf("FECHA"),
    CODIGO:  headers.indexOf("CODIGO"),
    KILOS:   headers.indexOf("KILOS"),
    PIEZAS:  headers.indexOf("PIEZAS"),
    FAMILIA: headers.indexOf("FAMILIA"),
    DESC:    headers.indexOf("DESCRIPCION"),
    REM:     headers.indexOf("REMISION"),
    PED:     headers.indexOf("PEDIDO"),
    ENV:     headers.indexOf("ENVIO")
  };
  var reporteFam = {}, reporteCat = {}, totalKilos = 0, rawRows = [];
  var ordenCats = ["CLAVO CONCRETO","CLAVO MADERA","BARRA","VARILLA ROSCADA","TORNILLO","COLATADO","ALAMBRES","OTROS"];
  ordenCats.forEach(function(c){ reporteCat[c] = { kilos:0, piezas:0 }; });
  var yearActual = new Date().getFullYear();
  var mesesNombres = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"];
  var tituloMeses = mesesSeleccionados.map(function(m){ return mesesNombres[m]; }).join(", ");
  for(var i=1; i<data.length; i++) {
    var fecha = data[i][IDX.FECHA];
    if(!(fecha instanceof Date)) fecha = new Date(fecha);
    if(fecha.getFullYear() == yearActual && mesesSeleccionados.indexOf(fecha.getMonth()) > -1) {
      var kilos = Number(data[i][IDX.KILOS]) || 0;
      var pzas  = Number(data[i][IDX.PIEZAS]) || 0;
      var fam   = IDX.FAMILIA > -1 ? String(data[i][IDX.FAMILIA]).toUpperCase().trim() : "SIN FAMILIA";
      if(!fam) fam = "SIN FAMILIA";
      if(!reporteFam[fam]) reporteFam[fam] = { kilos:0, piezas:0 };
      reporteFam[fam].kilos  += kilos;
      reporteFam[fam].piezas += pzas;
      var cat = "OTROS";
      if(!reporteCat[cat]) reporteCat[cat] = { kilos:0, piezas:0 };
      reporteCat[cat].kilos  += kilos;
      reporteCat[cat].piezas += pzas;
      totalKilos += kilos;
      rawRows.push({
        FECHA:       Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        REMISION:    data[i][IDX.REM],
        PEDIDO:      data[i][IDX.PED],
        CODIGO:      String(data[i][IDX.CODIGO]),
        DESCRIPCION: data[i][IDX.DESC],
        FAMILIA:     data[i][IDX.FAMILIA],
        KILOS:       kilos,
        PIEZAS:      pzas,
        ENVIO:       data[i][IDX.ENV]
      });
    }
  }
  return {
    reporteFam: reporteFam, reporteCat: reporteCat, ordenCats: ordenCats,
    total: totalKilos, tituloMeses: tituloMeses, year: yearActual,
    diasOperativos: 1, promedio: totalKilos, rawRows: rawRows
  };
}

function eliminarRegistroEnvio(id) {
  var ss = SpreadsheetApp.openById(ID_SS_PRODUCCION);
  var sheet = ss.getSheetByName("ENVIADO");
  var data = sheet.getDataRange().getValues();
  var idxID = data[0].map(function(h){ return String(h).toUpperCase().trim(); }).indexOf("ID");
  for(var i=1; i<data.length; i++) {
    if(String(data[i][idxID]) === String(id)) {
      sheet.deleteRow(i + 1);
      return "✅ Eliminado";
    }
  }
  return "⚠️ No encontrado";
}

// ============================================================
// MÓDULO DESPACHOS
// Lee hoja LOTES de la hoja de Producción.
// Muestra lotes de los últimos 150 días (FECHA_REG col I = índice 8).
// Excluye ENVIADO y CANCELADO.
// ============================================================

function obtenerLotesDisponibles(dias) {
  try {
    // dias: cuántos días hacia atrás buscar (default 60)
    var numDias  = (typeof dias === 'number' && dias > 0) ? dias : 60;

    var ss     = SpreadsheetApp.openById(ID_SS_PRODUCCION);
    var sheetL = ss.getSheetByName("LOTES");
    var sheetO = ss.getSheetByName("ORDENES");
    if (!sheetL) return [{ lote:"ERROR", desc:"No existe hoja LOTES", estatus:"CRITICO" }];

    var lastRowL = sheetL.getLastRow();
    if (lastRowL < 2) return [];

    // Cabeceras (solo fila 1, cols A..P)
    var hRow = sheetL.getRange(1, 1, 1, 16).getValues()[0];
    var hL   = hRow.map(function(h){ return String(h).toUpperCase().trim(); });
    var iLote   = hL.indexOf("LOTE");         if(iLote<0)   iLote=4;
    var iOrden  = hL.indexOf("ORDEN");        if(iOrden<0)  iOrden=2;
    var iEst    = hL.indexOf("ESTATUS");      if(iEst<0)    iEst=14;
    var iKg     = hL.indexOf("KG_EMBARQUES"); if(iKg<0)     iKg=13;
    var iTina   = hL.indexOf("TINA");         if(iTina<0)   iTina=11;
    var iSello  = hL.indexOf("SELLO_EMB");   if(iSello<0)  iSello=15;
    var iFecha  = hL.indexOf("FECHA_REG");   if(iFecha<0)  iFecha=8;

    // Leer datos (cols A..P)
    var dataL = sheetL.getRange(2, 1, lastRowL - 1, 16).getValues();

    var limiteMs = new Date().getTime() - numDias * 24 * 60 * 60 * 1000;

    // Pre-filtrar lotes
    var candidatos     = [];
    var idsNecesarios  = {};
    for (var j = 0; j < dataL.length; j++) {
      var row    = dataL[j];
      var loteID = String(row[iLote]).trim();
      if (!loteID || loteID === "undefined") continue;

      var fReg = row[iFecha];
      if (fReg) {
        var fMs = (fReg instanceof Date) ? fReg.getTime() : new Date(fReg).getTime();
        if (!isNaN(fMs) && fMs < limiteMs) continue;
      }

      var est = String(row[iEst]).toUpperCase().trim();
      if (est === "ENVIADO" || est === "CANCELADO") continue;

      var idOrden = String(row[iOrden]).trim();
      idsNecesarios[idOrden] = true;
      candidatos.push({
        lote:    loteID,
        idOrden: idOrden,
        estatus: (est === "NADA" || est === "") ? "PENDIENTE" : est,
        kg:      Number(row[iKg])   || 0,
        tina:    String(row[iTina]  || ""),
        sello:   String(row[iSello] || "")
      });
    }

    if (candidatos.length === 0) return [];

    // Leer ORDENES: solo cols A-B y G-H
    var mapOrd = {};
    if (sheetO) {
      var lastRowO = sheetO.getLastRow();
      if (lastRowO > 1) {
        var colsAB = sheetO.getRange(2, 1, lastRowO - 1, 2).getValues();
        var colsGH = sheetO.getRange(2, 7, lastRowO - 1, 2).getValues();
        for (var i = 0; i < colsAB.length; i++) {
          var idO = String(colsAB[i][0]).trim();
          if (idO && idsNecesarios[idO]) {
            mapOrd[idO] = {
              pedido: String(colsAB[i][1] || "-"),
              codigo: String(colsGH[i][0] || "-"),
              desc:   String(colsGH[i][1] || "Sin Desc")
            };
          }
        }
      }
    }

    return candidatos.map(function(c) {
      var info = mapOrd[c.idOrden] || { pedido:"-", codigo:"-", desc:"Orden "+c.idOrden };
      return {
        lote:    c.lote,
        pedido:  info.pedido,
        codigo:  info.codigo,
        desc:    info.desc,
        estatus: c.estatus,
        kg:      c.kg,
        tina:    c.tina,
        sello:   c.sello
      };
    });

  } catch (e) {
    return [{ lote:"ERROR", desc:e.toString(), estatus:"CRITICO" }];
  }
}

function procesarYRenderizarEtiquetas(listaAImprimir) {
  var ss       = SpreadsheetApp.openById(ID_SS_PRODUCCION);
  var sheetL   = ss.getSheetByName("LOTES");
  var sheetO   = ss.getSheetByName("ORDENES");
  var dataL    = sheetL.getDataRange().getValues();
  var dataO    = sheetO.getDataRange().getValues();

  // Mapeo Ordenes
  var mapOrd = {};
  for(var i=1; i<dataO.length; i++) {
    var idO = String(dataO[i][0]).trim();
    if(!idO) continue;
    mapOrd[idO] = {
      pedido: dataO[i][1],
      codigo: dataO[i][6],
      desc:   dataO[i][7],
      pesoU:  Number(dataO[i][18]) || 0,
      tipo:   String(dataO[i][19] || "").toUpperCase(),
      dia:    dataO[i][20],
      long:   dataO[i][21],
      cuerda: dataO[i][22] || "",
      cuerpo: dataO[i][23] || "",
      acero:  dataO[i][24] || ""
    };
  }

  // Mapeo filas de LOTES
  var mapFilasLotes = {};
  for(var j=1; j<dataL.length; j++) {
    var idL = String(dataL[j][4]).trim();
    if(idL) mapFilasLotes[idL] = j + 1;
  }

  var datosParaTemplate = [];
  var fechaNow   = new Date();
  var fechaImpStr = Utilities.formatDate(fechaNow, "GMT-6", "dd/MM/yyyy HH:mm");

  listaAImprimir.forEach(function(item) {
    var idLoteStr = String(item.lote).trim();
    var fila      = mapFilasLotes[idLoteStr];
    var infoO     = mapOrd[String(item.ordenRef || "").trim()];
    if(!infoO && fila) {
      infoO = mapOrd[String(dataL[fila-1][2]).trim()];
    }
    if(fila && infoO) {
      sheetL.getRange(fila, 12).setValue(item.tina);
      sheetL.getRange(fila, 13).setValue(fechaNow);
      sheetL.getRange(fila, 14).setValue(item.kg);        // Col N: KG_EMBARQUES (Peso Neto)
      sheetL.getRange(fila, 15).setValue("IMPRESO");
      sheetL.getRange(fila, 16).setValue(item.sello);
      sheetL.getRange(fila, 18).setValue(Number(item.pesoBruto) || 0);  // Col R: Peso Bruto
      sheetL.getRange(fila, 19).setValue(Number(item.pesoTina)  || 0);  // Col S: Peso Tina
      var pNeto = Number(item.kg) || 0;
      var pUnit = Number(infoO.pesoU) || 0;
      var conversionTxt = "-";
      if(pUnit > 0) {
        if(infoO.tipo.includes("VARILLA")) {
          var longitudNum = parseFloat(String(infoO.long).replace(/[^0-9.]/g,"")) || 1;
          conversionTxt = "( " + Math.round(pNeto/(pUnit*longitudNum)).toLocaleString() + " PZA )";
        } else if(infoO.tipo.includes("COLATADO")) {
          conversionTxt = "( " + Math.round(pNeto/pUnit).toLocaleString() + " ROLLOS )";
        } else {
          conversionTxt = "( " + Math.round(pNeto/pUnit).toLocaleString() + " PZA )";
        }
      }
      datosParaTemplate.push({
        lote:       idLoteStr,
        pedido:     infoO.pedido || "-",
        codigo:     infoO.codigo || "-",
        desc:       infoO.desc   || "-",
        tipo:       infoO.tipo   || "-",
        dia:        infoO.dia    || "-",
        long:       infoO.long   || "-",
        det:        infoO.cuerda + " " + infoO.cuerpo,
        acero:      infoO.acero  || "-",
        pesoU:      pUnit,
        pesoNeto:   pNeto,
        conversion: conversionTxt,
        sello:      item.sello        || "-",
        tina:       item.tina         || "-",
        pesoBruto:  Number(item.pesoBruto) || 0,
        pesoTina:   Number(item.pesoTina)  || 0,
        fechaImp:   fechaImpStr,
        iconoTipo:  obtenerIconoSVG(infoO.tipo || "")
      });
    }
  });

  if(datosParaTemplate.length === 0) throw "No se encontraron datos para las etiquetas seleccionadas.";

  // Generar HTML directamente (sin depender de archivo EtiquetasEmbarqueHTML)
  // Usar template HTML del proyecto (EtiquetasEmbarqueHTML debe existir en este GS)
  var template = HtmlService.createTemplateFromFile('EtiquetasEmbarqueHTML');
  template.etiquetas = datosParaTemplate;
  return template.evaluate().getContent();
}

function obtenerIconoSVG(tipo) {
  // Limpiamos el texto para evitar errores por espacios o nulos
  var t = String(tipo || "").toUpperCase().trim();
  var svgStart = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" style="width:100%; height:100%;">';
  var path = "";

  // --- 1. BLOQUE DE CLAVOS ---
  if (t.includes("CLAVO")) {
      if (t.includes("COLATADO") || t.includes("ROOFING")) {
          path = '<rect x="6" y="4" width="2" height="16" fill="#9e9e9e"/><rect x="5" y="3" width="4" height="1" fill="#757575"/><path d="M6 20 L7 22 L8 20 Z" fill="#9e9e9e"/><rect x="16" y="4" width="2" height="16" fill="#9e9e9e"/><rect x="15" y="3" width="4" height="1" fill="#757575"/><path d="M16 20 L17 22 L18 20 Z" fill="#9e9e9e"/><line x1="6" y1="8" x2="18" y2="8" stroke="#ff8f00" stroke-width="1"/><line x1="6" y1="14" x2="18" y2="14" stroke="#ff8f00" stroke-width="1"/>';
      }
      else if (t.includes("BOYA")) {
          path = '<rect x="8" y="4" width="8" height="16" fill="#212121"/><rect x="6" y="2" width="12" height="2" fill="#000"/><path d="M8 20 L12 24 L16 20 Z" fill="#212121"/><line x1="9" y1="6" x2="15" y2="8" stroke="#616161" stroke-width="0.5"/><line x1="9" y1="9" x2="15" y2="11" stroke="#616161" stroke-width="0.5"/><line x1="9" y1="12" x2="15" y2="14" stroke="#616161" stroke-width="0.5"/><line x1="9" y1="15" x2="15" y2="17" stroke="#616161" stroke-width="0.5"/>';
      }
      else if (t.includes("CONCRETO") && (t.includes("LISO") || t.includes("LISA"))) {
          path = '<rect x="10" y="4" width="4" height="16" fill="#212121"/><rect x="8" y="2" width="8" height="2" fill="#000"/><path d="M10 20 L12 24 L14 20 Z" fill="#212121"/>';
      }
      else if (t.includes("CONCRETO")) {
          path = '<rect x="10" y="4" width="4" height="16" fill="#212121"/><rect x="8" y="2" width="8" height="2" fill="#000"/><path d="M10 20 L12 24 L14 20 Z" fill="#212121"/><line x1="10" y1="6" x2="14" y2="7" stroke="#757575" stroke-width="0.5"/><line x1="10" y1="9" x2="14" y2="10" stroke="#757575" stroke-width="0.5"/><line x1="10" y1="12" x2="14" y2="13" stroke="#757575" stroke-width="0.5"/><line x1="10" y1="15" x2="14" y2="16" stroke="#757575" stroke-width="0.5"/>';
      }
      else {
          path = '<rect x="11" y="4" width="2" height="17" fill="#9e9e9e"/><rect x="9" y="3" width="6" height="1" fill="#616161"/><path d="M11 21 L12 24 L13 21 Z" fill="#9e9e9e"/>';
      }
  }

  // --- 2. BLOQUE DE TORNILLOS Y OTROS (ELSE IF PARA EVITAR SOBREESCRITURA) ---
  else if (t.includes("CUA") && t.includes("CN")) {
    path = '<rect x="4" y="2" width="16" height="8" rx="1" fill="#78909c" stroke="#37474f"/> <rect x="9" y="10" width="6" height="12" fill="#b0bec5"/> <line x1="9" y1="14" x2="15" y2="14" stroke="#78909c"/> <line x1="9" y1="18" x2="15" y2="18" stroke="#78909c"/> <text x="12" y="8" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle" fill="black">CN</text>';
  }
  else if (t.includes("CUA") && t.includes("CV")) {
    path = '<rect x="4" y="2" width="16" height="8" rx="1" fill="#78909c" stroke="#37474f"/> <rect x="9" y="10" width="6" height="12" fill="#b0bec5"/> <line x1="9" y1="14" x2="15" y2="14" stroke="#78909c"/> <line x1="9" y1="18" x2="15" y2="18" stroke="#78909c"/> <text x="12" y="8" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle" fill="black">CV</text>';
  }
  else if (t.includes("CAR")) {
    path = '<path d="M4 9 Q12 1 20 9" fill="#90a4ae" stroke="#546e7a"/> <rect x="8" y="9" width="8" height="4" fill="#607d8b"/> <rect x="9" y="13" width="6" height="10" fill="#cfd8dc"/> <line x1="9" y1="16" x2="15" y2="16" stroke="#90a4ae"/> <line x1="9" y1="19" x2="15" y2="19" stroke="#90a4ae"/>';
  }
  else if (t.includes("B7") || t.includes("B-7")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle" fill="#333">B-7</text>';
  }
  else if (t.includes("A394") && (t.includes("T0") || t.includes("T-0"))) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle" fill="#333">T-0</text>';
  }
  else if (t.includes("RED") && t.includes("RAN")) {
    path = '<path d="M16 6a4 4 0 0 0-8 0" fill="#78909c"/> <rect x="11" y="2" width="2" height="4" fill="#333"/> <rect x="10" y="6" width="4" height="16" fill="#b0bec5"/> <line x1="10" y1="10" x2="14" y2="10" stroke="#78909c"/> <line x1="10" y1="14" x2="14" y2="14" stroke="#78909c"/> <line x1="10" y1="18" x2="14" y2="18" stroke="#78909c"/>';
  }
  else if (t.includes("PIJ") || t.includes("PIJA") || t.includes("PHI")) {
    path = '<path d="M17 5H7V3h10v2z M10 5l2 14 2-14" fill="#757575"/> <path d="M11 9h2 M11 12h2 M11 15h2" stroke="#424242"/>';
  }
  else if (t.includes("MAC")) {
    path = '<rect x="6" y="8" width="12" height="14" fill="#90a4ae" stroke="#546e7a"/> <rect x="10" y="2" width="4" height="6" fill="#cfd8dc" stroke="#546e7a"/>';
  }
  else if (t.includes("HEM")) {
    path = '<path d="M6 2h12v20H6z M10 2v6h4V2" fill="#90a4ae" fill-rule="evenodd" stroke="#546e7a"/>';
  }
  else if (t.includes("HEX") && t.includes("G2")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle" fill="#333">G2</text>';
  }
  else if (t.includes("G5")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <line x1="12" y1="12" x2="12" y2="4" stroke="black" stroke-width="1.5"/> <line x1="12" y1="12" x2="5" y2="16" stroke="black" stroke-width="1.5"/> <line x1="12" y1="12" x2="19" y2="16" stroke="black" stroke-width="1.5"/>';
  }
  else if (t.includes("G8")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <line x1="12" y1="6" x2="12" y2="4" stroke="black" stroke-width="2"/> <line x1="12" y1="18" x2="12" y2="20" stroke="black" stroke-width="2"/> <line x1="6" y1="9" x2="4" y2="8" stroke="black" stroke-width="2"/> <line x1="18" y1="9" x2="20" y2="8" stroke="black" stroke-width="2"/> <line x1="6" y1="15" x2="4" y2="16" stroke="black" stroke-width="2"/> <line x1="18" y1="15" x2="20" y2="16" stroke="black" stroke-width="2"/>';
  }
  else if (t.includes("A325") || t.includes("A-325")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle">A325</text>';
  }
  else if (t.includes("A490") || t.includes("A-490")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="5" text-anchor="middle">A490</text>';
  }
  else if (t.includes("T-1") || t.includes("T1")) {
    path = '<polygon points="12,2 20.6,7 20.6,17 12,22 3.4,17 3.4,7" fill="#bdbdbd" stroke="#616161"/> <text x="12" y="13" font-family="Arial" font-weight="bold" font-size="6" text-anchor="middle">T-1</text>';
  }
  else if (t.includes("BIS HEM")) {
    path = '<rect x="4" y="4" width="10" height="16" fill="#90a4ae"/> <circle cx="18" cy="12" r="3" fill="none" stroke="#546e7a" stroke-width="2"/>';
  }
  else if (t.includes("BIS MAC")) {
    path = '<rect x="10" y="4" width="10" height="16" fill="#90a4ae"/> <rect x="4" y="10" width="6" height="4" fill="#546e7a"/>';
  }
  else if (t.includes("ARM")) {
    path = '<path d="M12 2a6 6 0 1 0 0 12 6 6 0 0 0 0-12zm0 2a4 4 0 1 1 0 8 4 4 0 0 1 0-8z M12 14v8" stroke="#616161" stroke-width="2" fill="none"/>';
  }
  else if (t.includes("BIR") || t.includes("BIRLO")) {
    path = '<rect x="8" y="2" width="8" height="20" fill="#bdbdbd"/> <line x1="8" y1="6" x2="16" y2="6" stroke="#424242"/> <line x1="8" y1="10" x2="16" y2="10" stroke="#424242"/> <line x1="8" y1="14" x2="16" y2="14" stroke="#424242"/> <line x1="8" y1="18" x2="16" y2="18" stroke="#424242"/>';
  }
  else if (t.includes("GRA") || t.includes("GRAPA")) {
    path = '<path d="M6 18V8a6 6 0 0 1 12 0v10" fill="none" stroke="#616161" stroke-width="3"/> <path d="M6 18l-2-2 M18 18l2-2" stroke="#616161" stroke-width="2"/>';
  }
  else if (t.includes("PER") || t.includes("PERNO")) {
  path = '<rect x="5" y="2" width="14" height="4" rx="1" fill="#424242"/>' + // Cabeza plana y ancha
             '<rect x="9" y="6" width="6" height="15" fill="#757575"/>' +      // Cuerpo
             '<path d="M9 21 Q12 23 15 21 Z" fill="#757575"/>';               // Base levemente redondeada
  }
  else if (t.includes("REM") || t.includes("REMACHE")) {
    path = '<path d="M12 4a4 4 0 0 0-4 4v2h8V8a4 4 0 0 0-4-4z" fill="#78909c"/> <rect x="10" y="10" width="4" height="10" fill="#b0bec5"/>';
  }
  else if (t.includes("VARILLA")) {
     path = '<rect x="7" y="2" width="10" height="20" fill="#bdbdbd" stroke="#424242" stroke-width="0.5"/>' +
             '<line x1="7" y1="4" x2="17" y2="6" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="6" x2="17" y2="8" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="8" x2="17" y2="10" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="10" x2="17" y2="12" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="12" x2="17" y2="14" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="14" x2="17" y2="16" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="16" x2="17" y2="18" stroke="#616161" stroke-width="0.8"/>' +
             '<line x1="7" y1="18" x2="17" y2="20" stroke="#616161" stroke-width="0.8"/>';
  }
  else if (t.includes("BARRA") || t.includes("COLD") || t.includes("REDONDO")) {
    var fill = "#bdbdbd";
    if (t.includes("COLD")) fill = "#a5d6a7";
    else if (t.includes("REDONDO")) fill = "#fff59d";
    path = '<rect x="4" y="2" width="16" height="20" fill="' + fill + '" stroke="black" stroke-width="1.0"/>';
  }
  else if (t.includes("ALA") || t.includes("ALAMBRE") || t.includes("ROLLO")) {
      path = '<g fill="none" stroke="#f57c00" stroke-width="2"><circle cx="12" cy="12" r="11"/><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="7"/></g>';
  }
  
  // --- 3. FALLBACK FINAL: SI NO COINCIDIÓ NADA ---
  if (path === "") {
      path = '<rect x="2" y="6" width="20" height="12" fill="#cfd8dc" stroke="#78909c" stroke-width="2"/> <text x="12" y="15" font-family="Arial" font-weight="bold" font-size="7" text-anchor="middle" fill="#455a64">ESP</text>';
  }

  return svgStart + path + "</svg>";
}
