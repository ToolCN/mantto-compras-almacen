////////////////////////////////////////////////////////////////////////////////////////////
////////////              ManttoGS.gs — FUNCIONES DE MANTENIMIENTO         ////////////////
////////////////////////////////////////////////////////////////////////////////////////////
// Contiene TODAS las funciones usadas por MainAppHTML para el módulo de Mantenimiento.
// Las variables globales (ID_HOJA_OM, ID_HOJA_ESTANDARES, TOKEN, etc.)
// están declaradas en Codigo_GS.gs — ambos archivos deben ir en el mismo proyecto GAS.
////////////////////////////////////////////////////////////////////////////////////////////


// ══════════════════════════════════════════════════════════════════════
//  TAREAS BASE
// ══════════════════════════════════════════════════════════════════════

function getSoloTareas() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("TAREAS_BASE");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    return data.slice(1).filter(function(r){ return r[0]; }).map(function(r){
      var acts = [];
      try { acts = JSON.parse(r[3] || "[]"); } catch(e) {}
      if (!acts.length) acts = [String(r[1] || "")];
      return {
        clave:        String(r[0]),
        descripcion:  String(r[1] || acts[0] || ""),
        frecuencia:   Number(r[2]),
        actividades:  acts,
        horas_hombre: Number(r[4] || 0)   // columna E
      };
    });
  } catch(e) { return []; }
}

function guardarTareaBase(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss, "TAREAS_BASE",
      ["CLAVE","DESCRIPCION","FRECUENCIA","ACTIVIDADES_JSON","HORAS_HOMBRE"]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(obj.clave))
        return "EXISTE"; // nueva tarea duplicada — no tocar
    }
    sheet.appendRow([
      String(obj.clave),
      String(obj.descripcion || (obj.actividades||[])[0] || ""),
      Number(obj.frecuencia),
      JSON.stringify(obj.actividades || [obj.descripcion || ""]),
      Number(obj.horas_hombre || 0)
    ]);
    return "OK";
  } catch(e) { return "Error: " + e; }
}

function editarTareaBase(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("TAREAS_BASE");
    if (!sheet) return "NO_SHEET";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(obj.clave)) {
        sheet.getRange(i+1, 2).setValue(String(obj.descripcion || (obj.actividades||[])[0] || ""));
        sheet.getRange(i+1, 3).setValue(Number(obj.frecuencia));
        sheet.getRange(i+1, 4).setValue(JSON.stringify(obj.actividades || [obj.descripcion || ""]));
        sheet.getRange(i+1, 5).setValue(Number(obj.horas_hombre || 0));  // columna E
        return "OK";
      }
    }
    return "NOT_FOUND";
  } catch(e) { return "Error: " + e; }
}

function eliminarTareaBase(clave) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("TAREAS_BASE");
    if (!sheet) return "Error: hoja TAREAS_BASE no encontrada";
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).toUpperCase() === String(clave).toUpperCase()) {
        sheet.deleteRow(i + 1);
        return "OK";
      }
    }
    return "Error: clave no encontrada";
  } catch(e) { return "Error: " + e; }
}


// ══════════════════════════════════════════════════════════════════════
//  PLANES DE MANTTO V2
// ══════════════════════════════════════════════════════════════════════

function getPlanesMPV2() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("PLANES_MP_V2");
    if (!sheet||sheet.getLastRow()<2) return [];

    var shMP = ss.getSheetByName("MAQUINAS_PLANES");
    var maqMap = {};
    if (shMP && shMP.getLastRow() > 1) {
      shMP.getDataRange().getValues().slice(1).forEach(function(r){
        var maq  = String(r[0]).trim();
        var plan = String(r[1]).trim().toUpperCase();
        if (!maq || !plan) return;
        if (!maqMap[plan]) maqMap[plan] = [];
        if (!maqMap[plan].includes(maq)) maqMap[plan].push(maq);
      });
    }

    return sheet.getDataRange().getValues().slice(1).map(function(r){
      var tareas; try{tareas=JSON.parse(r[4]);}catch(e){tareas=[];}
      var clave = String(r[1]);
      return {
        clave:      clave,
        nombre:     String(r[2]),
        frecuencia: parseInt(r[3])||0,
        tareas:     tareas,
        refacciones: (function(){ try{ return JSON.parse(r[5]||'[]'); }catch(e){ return []; } })(),
        maquinas:   maqMap[clave.toUpperCase()] || []
      };
    });
  } catch(e) { return []; }
}

function guardarPlanMP(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss,"PLANES_MP_V2",["ID","CLAVE","NOMBRE","FRECUENCIA","TAREAS_JSON"]);
    var data = sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][1]).toUpperCase()===String(obj.clave).toUpperCase()) {
        sheet.getRange(i+1,3).setValue(obj.nombre.toUpperCase());
        sheet.getRange(i+1,4).setValue(parseInt(obj.frecuencia)||0);
        sheet.getRange(i+1,5).setValue(JSON.stringify(obj.tareas||[]));
        return "OK-update";
      }
    }
    sheet.appendRow([Utilities.getUuid(),obj.clave.toUpperCase(),obj.nombre.toUpperCase(),parseInt(obj.frecuencia)||0,JSON.stringify(obj.tareas||[])]);
    return "OK-create";
  } catch(e) { return "Error: "+e; }
}

function eliminarPlanMP(clave) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("PLANES_MP_V2");
    if (!sheet) return "Error";
    var data = sheet.getDataRange().getValues();
    for (var i=data.length-1;i>=1;i--) {
      if (String(data[i][1]).toUpperCase()===clave.toUpperCase()) { sheet.deleteRow(i+1); return "OK"; }
    }
    return "Error: no encontrado";
  } catch(e) { return "Error: "+e; }
}

function guardarRefaccionesPlan(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("PLANES_MP_V2");
    if(!sheet) return "Error: hoja no encontrada";
    var data = sheet.getDataRange().getValues();
    for(var i=1; i<data.length; i++){
      if(String(data[i][1]).toUpperCase() === String(obj.clavePlan).toUpperCase()){
        sheet.getRange(i+1, 6).setValue(JSON.stringify(obj.refacciones||[]));
        return "OK";
      }
    }
    return "Error: plan no encontrado";
  } catch(e){ return "Error: "+e; }
}


// ══════════════════════════════════════════════════════════════════════
//  MANTENIMIENTO AUTÓNOMO
// ══════════════════════════════════════════════════════════════════════

function getPlanesAutonomo() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("PLANES_AUTO");
    if (!sheet||sheet.getLastRow()<2) return [];
    return sheet.getDataRange().getValues().slice(1)
      .filter(function(r){return r[1];})
      .map(function(r){
        var tareas; try{tareas=JSON.parse(r[3]);}catch(e){tareas=[];}
        return { clave:String(r[1]), maquina:String(r[2]), tareas:tareas };
      });
  } catch(e) { return []; }
}

function guardarPlanAutonomo(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss,"PLANES_AUTO",["ID","CLAVE","MAQUINA","TAREAS_JSON"]);
    var data = sheet.getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][1]).toUpperCase()===String(obj.clave).toUpperCase()) {
        sheet.getRange(i+1,3).setValue(obj.maquina||"");
        sheet.getRange(i+1,4).setValue(JSON.stringify(obj.tareas||[]));
        return "OK-update";
      }
    }
    sheet.appendRow([Utilities.getUuid(),obj.clave.toUpperCase(),obj.maquina||"",JSON.stringify(obj.tareas||[])]);
    return "OK-create";
  } catch(e) { return "Error: "+e; }
}


// ══════════════════════════════════════════════════════════════════════
//  HELPERS INTERNOS
// ══════════════════════════════════════════════════════════════════════

// Normaliza cualquier valor de celda fecha a "yyyy-MM-dd" o ""
function _normDateVal(v) {
  if (!v) return "";
  if (v instanceof Date && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, "GMT-6", "yyyy-MM-dd");
  }
  if (typeof v === "string") {
    var s = v.trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
    var mf = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (mf) return mf[3]+"-"+mf[2]+"-"+mf[1];
    return "";
  }
  if (typeof v === "number" && v > 40000) {
    return Utilities.formatDate(new Date(Math.round((v-25569)*86400000)), "GMT-6", "yyyy-MM-dd");
  }
  return "";
}

// Obtener o crear hoja con encabezados
function _getOrCreate(ss, nombre, headers) {
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(headers);
  }
  return sheet;
}

function _emptyKPIs(){
  return {om:0,mp:0,ot:0,criticas:0,tiempoMuerto:0,mpPlaneadas:0,mpRealizadas:0,semanas:[],porArea:[]};
}

function _calcularSemanas(mes, anio, data){
  var diasEnMes = new Date(anio, mes, 0).getDate();
  var semanas = [];
  var d = 1;
  while(d <= diasEnMes){
    var fin = Math.min(d + 6, diasEnMes);
    var ini = new Date(anio, mes-1, d);
    var finD = new Date(anio, mes-1, fin);
    if(fin < diasEnMes){
      var dayOfWeek = finD.getDay();
      if(dayOfWeek !== 6){
        fin = Math.min(d + (6 - ini.getDay()), diasEnMes);
        finD = new Date(anio, mes-1, fin);
      }
    }
    var generadas=0, terminadas=0;
    data.forEach(function(r){
      var f = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if(isNaN(f.getTime())) return;
      if(f >= ini && f <= finD){
        var folio=String(r[2]||"");
        if(folio) generadas++;
        if(String(r[17]||"").toUpperCase()==='CERRADA') terminadas++;
      }
    });
    var pad = function(n){ return String(n).padStart(2,'0'); };
    semanas.push({
      inicio: pad(d)+'/'+pad(mes)+'/'+anio,
      fin:    pad(fin)+'/'+pad(mes)+'/'+anio,
      generadas: generadas,
      terminadas: terminadas
    });
    d = fin + 1;
  }
  return semanas;
}


// ══════════════════════════════════════════════════════════════════════
//  PROGRAMADOR MP
// ══════════════════════════════════════════════════════════════════════

function getProgramadorData(mes, anio) {
  try {
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);

    var dataAct = ssEst.getSheetByName("ESTANDARES").getDataRange().getValues();

    var shProg = _getOrCreate(ss, "PROG_ORDENES",
      ["ID","FOLIO","MAQUINA","PLAN","FECHA_PLANEADA","FECHA_REAL","ESTADO","RETRASADA","MES","ANIO","PROG_FIJA"]);

    var dataProg = shProg.getLastRow() > 1 ? shProg.getDataRange().getValues().slice(1) : [];

    var ordenesMes = dataProg
      .filter(function(r) { return parseInt(r[8]) === mes && parseInt(r[9]) === anio; })
      .map(function(r) {
        return {
          folio:          String(r[1]),
          maquina:        String(r[2]).trim(),
          plan:           String(r[3]),
          fecha_planeada: _normDateVal(r[4]),
          fecha_real:     _normDateVal(r[5]),
          estado:         String(r[6] || "ABIERTA"),
          retrasada:      r[7] === true || r[7] === "TRUE" || r[7] === "true",
          prog_fija:      _normDateVal(r[10])
        };
      });

    var gruposMap = {};
    for (var i = 1; i < dataAct.length; i++) {
      var maq   = String(dataAct[i][3] || "").trim();
      var grupo = String(dataAct[i][9] || "SIN GRUPO").trim().toUpperCase();
      if (!maq) continue;
      if (!gruposMap[grupo]) gruposMap[grupo] = [];
      if (!gruposMap[grupo].find(function(m) { return m.nombre === maq; })) {
        gruposMap[grupo].push({ nombre: maq });
      }
    }

    var grupos = Object.keys(gruposMap).map(function(g) {
      return { nombre: g, maquinas: gruposMap[g] };
    });

    return { grupos: grupos, ordenes: ordenesMes };

  } catch(e) {
    Logger.log("getProgramadorData ERROR: " + e);
    return { grupos: [], ordenes: [] };
  }
}

function generarOrdenesPreventivas(mes, anio) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);

    var shOM   = ss.getSheetByName("OM");
    var shProg = _getOrCreate(ss,"PROG_ORDENES",["ID","FOLIO","MAQUINA","PLAN","FECHA_PLANEADA","FECHA_REAL","ESTADO","RETRASADA","MES","ANIO"]);
    var shPlanes = ss.getSheetByName("PLANES_MP_V2");
    if (!shPlanes) return {creadas:0};

    var planesData = shPlanes.getDataRange().getValues().slice(1);
    var dataAct = ssEst.getSheetByName("ESTANDARES").getDataRange().getValues();

    var progExist = shProg.getDataRange().getValues().slice(1);
    var creadas=0;
    var primerDia = new Date(anio,mes-1,1);

    planesData.forEach(function(plan){
      var clavePlan = String(plan[1]);
      var frecuencia= parseInt(plan[3])||30;

      var maquinas = getMaquinasDelPlan(clavePlan, ss, dataAct);

      maquinas.forEach(function(maq){
        var yaExiste = progExist.some(function(r){
          return String(r[2])===maq && String(r[3])===clavePlan && parseInt(r[8])===mes && parseInt(r[9])===anio;
        });
        if(yaExiste) return;

        var hayAbierta = progExist.some(function(r){
          var estadoR = String(r[6]).toUpperCase();
          return String(r[2])===maq && String(r[3])===clavePlan
            && (estadoR==='ABIERTA' || estadoR==='INICIADO');
        });
        if(hayAbierta) { Logger.log("SKIP: " + maq + " / " + clavePlan + " — hay orden sin cerrar"); return; }

        var fechaPlaneada = calcularFechaPlaneada(maq, clavePlan, frecuencia, primerDia, ss);
        var folio = getSiguienteFolio("MP");

        var rowOM = new Array(18).fill("");
        rowOM[0]=Utilities.getUuid(); rowOM[1]=new Date(); rowOM[2]=folio;
        rowOM[3]=maq; rowOM[4]="MANTTO PREVENTIVO"; rowOM[5]="PLAN: "+clavePlan;
        rowOM[6]="PREVENTIVO"; rowOM[7]="NORMAL"; rowOM[8]=new Date(); rowOM[17]="ABIERTA";
        shOM.appendRow(rowOM);

        shProg.appendRow([Utilities.getUuid(),folio,maq,clavePlan,fechaPlaneada,"","ABIERTA",false,mes,anio,""]);
        creadas++;
      });
    });

    return {creadas:creadas};
  } catch(e) {
    Logger.log("generarOrdenesPreventivas error: "+e);
    throw e;
  }
}

function calcularFechaPlaneada(maq, clavePlan, frecuencia, primerDiaMes, ss) {
  try {
    var shOM = ss.getSheetByName("OM");
    var data = shOM ? shOM.getDataRange().getValues() : [];
    var ultimaFecha = null;

    for (var i = 1; i < data.length; i++) {
      var esMismaMAQ  = String(data[i][3]).trim() === String(maq).trim();
      var esMismoPlan = String(data[i][5]).indexOf(clavePlan) !== -1;
      var esCerrada   = String(data[i][17]) === "CERRADA";
      if (esMismaMAQ && esMismoPlan && esCerrada) {
        var fc = data[i][10] ? new Date(data[i][10]) : null;
        if (fc && !isNaN(fc.getTime()) && (!ultimaFecha || fc > ultimaFecha)) {
          ultimaFecha = fc;
        }
      }
    }

    var primerDia = (primerDiaMes instanceof Date && !isNaN(primerDiaMes.getTime()))
      ? primerDiaMes
      : new Date();

    var fechaBase;
    if (ultimaFecha) {
      fechaBase = new Date(ultimaFecha.getTime() + (frecuencia * 86400000));
    } else {
      fechaBase = new Date(primerDia.getFullYear(), primerDia.getMonth(), 1);
    }

    if (fechaBase < primerDia) {
      fechaBase = new Date(primerDia.getFullYear(), primerDia.getMonth(), 1);
    }

    return Utilities.formatDate(fechaBase, "GMT-6", "yyyy-MM-dd");

  } catch(e) {
    Logger.log("calcularFechaPlaneada error: " + e);
    return Utilities.formatDate(
      new Date(primerDiaMes instanceof Date ? primerDiaMes : new Date()),
      "GMT-6", "yyyy-MM-dd"
    );
  }
}

function reprogramarOrdenMP(obj) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("PROG_ORDENES");
    if (!sheet) return "Error: no existe PROG_ORDENES";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(obj.folio)) {
        sheet.getRange(i+1, 5).setValue(String(obj.nuevaFecha));
        sheet.getRange(i+1, 8).setValue(obj.retrasada ? true : false);
        var partesFecha = String(obj.nuevaFecha).split('-');
        if (partesFecha.length >= 2) {
          sheet.getRange(i+1, 9).setValue(parseInt(partesFecha[1]));
          sheet.getRange(i+1, 10).setValue(parseInt(partesFecha[0]));
        }
        if (obj.progFija) {
          var actual = String(data[i][10] || "").trim();
          if (!actual) {
            sheet.getRange(i+1, 11).setValue(String(obj.progFija));
          }
        }
        return "OK";
      }
    }
    return "Error: folio no encontrado";
  } catch(e) { return "Error: " + e; }
}

function fijarOrdenMP(obj) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("PROG_ORDENES");
    if (!sheet) return "Error: no existe PROG_ORDENES";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(obj.folio)) {
        var actual = String(data[i][10] || "").trim();
        if (!actual) {
          sheet.getRange(i+1, 11).setValue(String(obj.fechaFija));
        }
        return "OK";
      }
    }
    return "Error: folio no encontrado";
  } catch(e) { return "Error: " + e; }
}

function previsualizarOrdenesPreventivas(mes, anio) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);

    var shProg   = _getOrCreate(ss, "PROG_ORDENES",
      ["ID","FOLIO","MAQUINA","PLAN","FECHA_PLANEADA","FECHA_REAL","ESTADO","RETRASADA","MES","ANIO"]);
    var shPlanes = ss.getSheetByName("PLANES_MP_V2");
    if (!shPlanes) return [];

    var planesData  = shPlanes.getDataRange().getValues().slice(1);
    var dataAct     = ssEst.getSheetByName("ESTANDARES").getDataRange().getValues();
    var progExist   = shProg.getLastRow() > 1 ? shProg.getDataRange().getValues().slice(1) : [];
    var primerDia   = new Date(anio, mes - 1, 1);
    var resultado   = [];

    planesData.forEach(function(plan) {
      var clavePlan  = String(plan[1]).trim();
      var planNombre = String(plan[2] || clavePlan);
      var frecuencia = parseInt(plan[3]) || 30;
      if (!clavePlan) return;

      var maquinas = getMaquinasDelPlan(clavePlan, ss, dataAct);

      maquinas.forEach(function(maq) {
        var yaExiste = progExist.some(function(r) {
          return String(r[2]).trim() === String(maq).trim()
            && String(r[3]).trim() === clavePlan
            && parseInt(r[8]) === mes
            && parseInt(r[9]) === anio;
        });
        if (yaExiste) return;

        var hayAbierta = progExist.some(function(r) {
          var est = String(r[6]).toUpperCase();
          return String(r[2]).trim()===String(maq).trim() && String(r[3]).trim()===clavePlan
            && (est==='ABIERTA' || est==='INICIADO');
        });
        if (hayAbierta) return;

        var fechaPlaneada = calcularFechaPlaneada(maq, clavePlan, frecuencia, primerDia, ss);
        resultado.push({
          maquina:      maq,
          plan:         clavePlan,
          planNombre:   planNombre,
          fechaPlaneada: fechaPlaneada,
          frecuencia:   frecuencia
        });
      });
    });

    return resultado;
  } catch(e) {
    Logger.log("previsualizarOrdenesPreventivas error: " + e);
    return [];
  }
}

function generarOrdenesSeleccionadas(mes, anio, seleccionadas) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var shOM  = ss.getSheetByName("OM");
    var shProg = _getOrCreate(ss, "PROG_ORDENES",
      ["ID","FOLIO","MAQUINA","PLAN","FECHA_PLANEADA","FECHA_REAL","ESTADO","RETRASADA","MES","ANIO"]);

    var progExist = shProg.getLastRow() > 1 ? shProg.getDataRange().getValues().slice(1) : [];
    var creadas = 0;

    (seleccionadas || []).forEach(function(item) {
      var maq       = String(item.maquina).trim();
      var clavePlan = String(item.plan).trim();

      var yaExiste = progExist.some(function(r) {
        return String(r[2]).trim() === maq
          && String(r[3]).trim() === clavePlan
          && parseInt(r[8]) === mes
          && parseInt(r[9]) === anio;
      });
      if (yaExiste) return;

      var hayAbierta = progExist.some(function(r) {
        var est = String(r[6]).toUpperCase();
        return String(r[2]).trim()===maq && String(r[3]).trim()===clavePlan
          && (est==='ABIERTA' || est==='INICIADO');
      });
      if (hayAbierta) { Logger.log("SKIP seleccionadas: "+maq+"/"+clavePlan+" — orden sin cerrar"); return; }

      var fechaStr = String(item.fechaPlaneada || "").trim();
      if (!fechaStr || fechaStr.length < 8) {
        var primerDia = new Date(anio, mes - 1, 1);
        fechaStr = calcularFechaPlaneada(maq, clavePlan, item.frecuencia || 30, primerDia, ss);
      }

      var folio = getSiguienteFolio("MP");

      var rowOM = new Array(18).fill("");
      rowOM[0]  = Utilities.getUuid();
      rowOM[1]  = new Date();
      rowOM[2]  = folio;
      rowOM[3]  = maq;
      rowOM[4]  = "MANTTO PREVENTIVO";
      rowOM[5]  = "PLAN: " + clavePlan;
      rowOM[6]  = "PREVENTIVO";
      rowOM[7]  = "NORMAL";
      rowOM[8]  = new Date();
      rowOM[17] = "ABIERTA";
      shOM.appendRow(rowOM);

      shProg.appendRow([
        Utilities.getUuid(), folio, maq, clavePlan, fechaStr,
        "", "ABIERTA", false, mes, anio, ""
      ]);

      creadas++;
    });

    return { creadas: creadas };
  } catch(e) {
    Logger.log("generarOrdenesSeleccionadas error: " + e);
    throw e;
  }
}

function actualizarGestionOrdenMP(obj) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var shOM  = ss.getSheetByName("OM");
    var shProg = ss.getSheetByName("PROG_ORDENES");
    var shHist = _getOrCreate(ss, "HIST_TAREAS_MP",
      ["ID","FOLIO","MAQUINA","PLAN","FECHA_EJECUCION","TAREA_CLAVE","RESULTADO","COMENTARIO"]);

    var folio      = String(obj.folio);
    var estado     = String(obj.estado || "");
    var tecnico    = String(obj.tecnico || "");
    var comentario = String(obj.comentario || "");
    var tareasCheck = obj.tareasCheck || {};

    var fechaInicio = null;
    var fechaCierre = null;
    if (obj.fechaInicio) { try { fechaInicio = new Date(obj.fechaInicio); } catch(e){} }
    if (obj.fechaCierre) { try { fechaCierre = new Date(obj.fechaCierre); } catch(e){} }

    if (shOM) {
      var dataOM = shOM.getDataRange().getValues();
      for (var i = 1; i < dataOM.length; i++) {
        if (String(dataOM[i][2]) === folio) {
          if (estado)    shOM.getRange(i+1, 18).setValue(estado);
          if (tecnico)   shOM.getRange(i+1, 15).setValue(tecnico);
          if (comentario !== undefined) shOM.getRange(i+1, 14).setValue(comentario);
          if (fechaInicio && !isNaN(fechaInicio.getTime()) && !dataOM[i][9]) {
            shOM.getRange(i+1, 10).setValue(fechaInicio);
          }
          if (estado === "CERRADA") {
            var dtCierre = (fechaCierre && !isNaN(fechaCierre.getTime())) ? fechaCierre : new Date();
            shOM.getRange(i+1, 11).setValue(dtCierre);
            if (!dataOM[i][9]) {
              var dtInicio = (fechaInicio && !isNaN(fechaInicio.getTime())) ? fechaInicio : dtCierre;
              shOM.getRange(i+1, 10).setValue(dtInicio);
            }
          }
          break;
        }
      }
    }

    var maquina = "";
    var plan    = "";
    if (shProg && shProg.getLastRow() > 1) {
      var dataProg = shProg.getDataRange().getValues();
      for (var j = 1; j < dataProg.length; j++) {
        if (String(dataProg[j][1]) === folio) {
          if (estado) shProg.getRange(j+1, 7).setValue(estado);
          if (estado === "CERRADA") {
            var solofecha = fechaCierre && !isNaN(fechaCierre.getTime())
              ? Utilities.formatDate(fechaCierre, "GMT-6", "yyyy-MM-dd")
              : Utilities.formatDate(new Date(), "GMT-6", "yyyy-MM-dd");
            shProg.getRange(j+1, 6).setValue(solofecha);
            shProg.getRange(j+1, 5).setValue(solofecha);
            var partesCierre = solofecha.split('-');
            shProg.getRange(j+1, 9).setValue(parseInt(partesCierre[1]));
            shProg.getRange(j+1, 10).setValue(parseInt(partesCierre[0]));
          }
          maquina = String(dataProg[j][2]);
          plan    = String(dataProg[j][3]);
          break;
        }
      }
    }

    var claves = Object.keys(tareasCheck);
    if (claves.length && maquina) {
      var hoy = new Date();
      var histData = shHist.getDataRange().getValues();
      var rowsABorrar = [];
      for (var k = histData.length - 1; k >= 1; k--) {
        if (String(histData[k][1]) === folio) rowsABorrar.push(k+1);
      }
      rowsABorrar.forEach(function(row){ shHist.deleteRow(row); });
      claves.forEach(function(clave){
        var resultado = tareasCheck[clave] || "PENDIENTE";
        shHist.appendRow([Utilities.getUuid(),folio,maquina,plan,hoy,clave,resultado,comentario]);
      });
    }

    return "OK";
  } catch(e) {
    Logger.log("actualizarGestionOrdenMP error: " + e);
    return "Error: " + e;
  }
}


// ══════════════════════════════════════════════════════════════════════
//  VISTAS ANUALES
// ══════════════════════════════════════════════════════════════════════

function getVistaAnualData(anio) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var shProg   = ss.getSheetByName("PROG_ORDENES");
    var shPlanes = ss.getSheetByName("PLANES_MP_V2");
    var dataAct  = ssEst.getSheetByName("ESTANDARES").getDataRange().getValues();

    var ordenes = [];
    if (shProg && shProg.getLastRow() > 1) {
      shProg.getDataRange().getValues().slice(1).forEach(function(r) {
        if (parseInt(r[9]) === parseInt(anio)) {
          ordenes.push({ maquina:String(r[2]), plan:String(r[3]),
            folio:String(r[1]), mes:parseInt(r[8]), estado:String(r[6]),
            fecha_planeada:String(r[4]||'') });
        }
      });
    }

    var proyecciones = [];
    var planesData = shPlanes ? shPlanes.getDataRange().getValues().slice(1) : [];
    planesData.forEach(function(plan) {
      var clavePlan  = String(plan[1]);
      var frecuencia = parseInt(plan[3]) || 30;
      var maquinas   = getMaquinasDelPlan(clavePlan, ss, dataAct);
      maquinas.forEach(function(maq) {
        for (var m = 1; m <= 12; m++) {
          var yaExiste = ordenes.some(function(o) {
            return o.maquina === maq && o.plan === clavePlan && o.mes === m;
          });
          if (!yaExiste) {
            var primerDia = new Date(anio, m-1, 1);
            var fp = calcularFechaPlaneada(maq, clavePlan, frecuencia, primerDia, ss);
            if (fp instanceof Date && fp.getMonth() === m-1 && fp.getFullYear() === parseInt(anio)) {
              proyecciones.push({ maquina:maq, plan:clavePlan, mes:m });
            }
          }
        }
      });
    });

    var gruposMap = {};
    for (var i = 1; i < dataAct.length; i++) {
      var maq   = String(dataAct[i][3]||'').trim();
      var grupo = String(dataAct[i][9]||'SIN GRUPO').trim().toUpperCase();
      if (!maq) continue;
      if (!gruposMap[grupo]) gruposMap[grupo] = [];
      if (!gruposMap[grupo].find(function(m){ return m.nombre===maq; }))
        gruposMap[grupo].push({ nombre: maq });
    }
    var grupos = Object.keys(gruposMap).map(function(g){ return { nombre:g, maquinas:gruposMap[g] }; });
    return { ordenes:ordenes, proyecciones:proyecciones, grupos:grupos };
  } catch(e) { Logger.log("getVistaAnualData: "+e); return { ordenes:[], proyecciones:[], grupos:[] }; }
}

function getOrdenesAnio(anio) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var shProg = _getOrCreate(ss, "PROG_ORDENES",
      ["ID","FOLIO","MAQUINA","PLAN","FECHA_PLANEADA","FECHA_REAL","ESTADO","RETRASADA","MES","ANIO"]);

    var dataProg = shProg.getLastRow() > 1 ? shProg.getDataRange().getValues().slice(1) : [];

    var ordenes = dataProg
      .filter(function(r){ return parseInt(r[9]) === parseInt(anio); })
      .map(function(r){
        return {
          folio:          String(r[1]),
          maquina:        String(r[2]).trim(),
          plan:           String(r[3]),
          fecha_planeada: _normDateVal(r[4]),
          fecha_real:     _normDateVal(r[5]),
          estado:         String(r[6] || "ABIERTA"),
          retrasada:      r[7] === true || r[7] === "TRUE" || r[7] === "true",
          mes:            parseInt(r[8]),
          anio:           parseInt(r[9]),
          prog_fija:      _normDateVal(r[10])
        };
      });

    var dataAct  = ssEst.getSheetByName("ESTANDARES").getDataRange().getValues();
    var gruposMap = {};
    for (var i = 1; i < dataAct.length; i++) {
      var maq   = String(dataAct[i][3]||"").trim();
      var grupo = String(dataAct[i][9]||"SIN GRUPO").trim().toUpperCase();
      if (!maq) continue;
      if (!gruposMap[grupo]) gruposMap[grupo] = [];
      if (!gruposMap[grupo].find(function(m){ return m.nombre===maq; }))
        gruposMap[grupo].push({ nombre: maq });
    }
    var grupos = Object.keys(gruposMap).map(function(g){ return { nombre:g, maquinas:gruposMap[g] }; });

    return { ordenes: ordenes, grupos: grupos };
  } catch(e) {
    Logger.log("getOrdenesAnio error: " + e);
    return { ordenes: [], grupos: [] };
  }
}


// ══════════════════════════════════════════════════════════════════════
//  TÉCNICOS
// ══════════════════════════════════════════════════════════════════════

// Versión unificada — filtra por Col E (AREA) = "MANTENIMIENTO"
function getTecnicos() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var data = ss.getSheetByName("OPERADORES").getDataRange().getValues();
    return data.filter(function(r){ return String(r[4]).toUpperCase().includes("MANTENIMIENTO"); })
               .map(function(r){ return r[1]; });
  } catch(e) { return []; }
}

function getTecnicosDeArea(area) {
  try {
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var shOp  = ssEst.getSheetByName("OPERADORES");
    if (!shOp || shOp.getLastRow() < 2) return [];
    var data = shOp.getDataRange().getValues().slice(1);
    return data
      .filter(function(r){ return r[0] && String(r[4]).toUpperCase().trim() === String(area).toUpperCase().trim(); })
      .map(function(r){ return String(r[1]); });
  } catch(e) { Logger.log("getTecnicosDeArea: "+e); return []; }
}

function getTecnicosPorArea(area) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sh = ss.getSheetByName("OPERADORES");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r){ return String(r[4]||'').trim().toUpperCase()===String(area||'').trim().toUpperCase() && r[1]; })
      .map(function(r){ return String(r[1]).trim(); });
  } catch(e) { Logger.log("getTecnicosPorArea: "+e); return []; }
}


// ══════════════════════════════════════════════════════════════════════
//  DASHBOARD KPIs y METRICS
// ══════════════════════════════════════════════════════════════════════

function getDashboardKPIs(mes, anio) {
  try {
    var ss   = SpreadsheetApp.openById(ID_HOJA_OM);
    var shOM = ss.getSheetByName("OM");
    if (!shOM || shOM.getLastRow() < 2) return _emptyKPIs();
    var data = shOM.getDataRange().getValues().slice(1);

    var now  = new Date();
    var m    = mes  ? parseInt(mes)  : (now.getMonth() + 1);
    var y    = anio ? parseInt(anio) : now.getFullYear();

    var om=0, mp=0, ot=0, criticas=0;
    var tiempoMuerto=0, mpPlaneadas=0, mpRealizadas=0;
    var porArea={};

    data.forEach(function(r){
      var folio  = String(r[2]||"");
      var estado = String(r[17]||"").toUpperCase();
      var fRow   = r[1] instanceof Date ? r[1] : new Date(r[1]);
      var fRowM  = fRow.getMonth()+1;
      var fRowY  = fRow.getFullYear();
      var enMes  = (fRowM===m && fRowY===y);

      if(!folio) return;

      if(estado!=='CERRADA'){
        if(folio.startsWith('OM-')) om++;
        else if(folio.startsWith('MP-')) mp++;
        else if(folio.startsWith('OT-')) ot++;
        if(String(r[7]||"").toUpperCase()==='CRÍTICA'||String(r[7]||"").toUpperCase()==='CRITICA') criticas++;
      }

      if(!enMes) return;

      if((folio.startsWith('OM-')||folio.startsWith('OT-')) && estado==='CERRADA'){
        tiempoMuerto += parseFloat(r[16]||0);
      }

      if(folio.startsWith('MP-')){
        mpPlaneadas++;
        if(estado==='CERRADA') mpRealizadas++;
      }

      var maq = String(r[3]||"Sin área");
      porArea[maq] = (porArea[maq]||0) + 1;
    });

    var semanas = _calcularSemanas(m, y, data);

    var topArea = Object.keys(porArea)
      .map(function(k){ return {maquina:k,total:porArea[k]}; })
      .sort(function(a,b){ return b.total-a.total; });

    return {
      om: om, mp: mp, ot: ot, criticas: criticas,
      tiempoMuerto: Math.round(tiempoMuerto*10)/10,
      mpPlaneadas: mpPlaneadas, mpRealizadas: mpRealizadas,
      semanas: semanas, porArea: topArea
    };
  } catch(e) {
    Logger.log("getDashboardKPIs: "+e);
    return _emptyKPIs();
  }
}

function getDashboardMetrics(mes, anio) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("OM");
    if (!sh || sh.getLastRow() < 2) return {om:0,mp:0,ot:0,criticas:0,backlog:0,tiempoMuerto:0,semanas:[],prevPlaneadas:0,prevHechas:0};

    var rows = sh.getDataRange().getValues().slice(1);
    var hoy = new Date(); hoy.setHours(0,0,0,0);
    var diasEnMes = new Date(anio, mes, 0).getDate();

    var om=0, mp=0, ot=0, criticas=0, backlog=0, tiempoMuerto=0;
    var semStats = [{sem:1,gen:0,ter:0},{sem:2,gen:0,ter:0},{sem:3,gen:0,ter:0},{sem:4,gen:0,ter:0}];

    rows.forEach(function(r) {
      var folio = String(r[2]||'');
      var estado = String(r[17]||'').trim().toUpperCase();
      var fechaCreacion = r[1] ? new Date(r[1]) : null;
      var fechaCierre = r[14] ? new Date(r[14]) : null;
      var prioridad = String(r[7]||'').trim().toUpperCase();
      var horas = parseFloat(r[15]||0);

      if(estado!=='CERRADA'){
        if(folio.startsWith('OM-')) om++;
        else if(folio.startsWith('MP-')) mp++;
        else if(folio.startsWith('OT-')) ot++;
        if(prioridad==='CRITICA'||prioridad==='CRÍTICA') criticas++;
        if(fechaCreacion && (fechaCreacion.getMonth()+1 < mes || fechaCreacion.getFullYear() < anio)) backlog++;
      }

      var enMes = fechaCreacion && (fechaCreacion.getMonth()+1===parseInt(mes)) && (fechaCreacion.getFullYear()===parseInt(anio));
      var cierreEnMes = fechaCierre && (fechaCierre.getMonth()+1===parseInt(mes)) && (fechaCierre.getFullYear()===parseInt(anio));

      if(enMes){
        var dia = fechaCreacion.getDate();
        var semIdx = dia<=7?0:dia<=14?1:dia<=21?2:3;
        semStats[semIdx].gen++;
      }
      if(cierreEnMes || (estado==='CERRADA' && enMes)){
        var diaRef = fechaCierre||fechaCreacion;
        if(diaRef){ var dia2=diaRef.getDate(); var si2=dia2<=7?0:dia2<=14?1:dia2<=21?2:3; semStats[si2].ter++; }
        if(!isNaN(horas)) tiempoMuerto+=horas;
      }
    });

    var shProg = ss.getSheetByName("PROG_ORDENES");
    var prevPlaneadas=0, prevHechas=0;
    if(shProg && shProg.getLastRow()>1){
      shProg.getDataRange().getValues().slice(1).forEach(function(r){
        if(parseInt(r[8])===parseInt(mes) && parseInt(r[9])===parseInt(anio)){
          prevPlaneadas++;
          if(String(r[6]||'').trim().toUpperCase()==='CERRADA') prevHechas++;
        }
      });
    }

    return {om,mp,ot,criticas,backlog,tiempoMuerto:Math.round(tiempoMuerto*10)/10,semanas:semStats,prevPlaneadas,prevHechas};
  } catch(e) {
    Logger.log("getDashboardMetrics: "+e);
    return {om:0,mp:0,ot:0,criticas:0,backlog:0,tiempoMuerto:0,semanas:[{sem:1,gen:0,ter:0},{sem:2,gen:0,ter:0},{sem:3,gen:0,ter:0},{sem:4,gen:0,ter:0}],prevPlaneadas:0,prevHechas:0};
  }
}


// ══════════════════════════════════════════════════════════════════════
//  BITÁCORAS — DEFINICIÓN Y LLENADO
// ══════════════════════════════════════════════════════════════════════

function getBitacoras() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("BITACORAS_DEF");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[1]) continue;
      var columnas = [];
      var detalles = [];
      try { columnas = JSON.parse(row[4] || "[]"); } catch(e) { columnas = []; }
      try { detalles = JSON.parse(row[5] || "[]"); } catch(e) { detalles = []; }
      result.push({
        titulo:      String(row[1] || ""),
        encabezado:  String(row[2] || ""),
        pie:         String(row[3] || ""),
        columnas:    columnas,
        detalles:    detalles,
        orientacion: String(row[6] || "portrait"),
        splitDias:   String(row[7]) === "true"
      });
    }
    return result;
  } catch(e) { return []; }
}

function guardarBitacora(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss, "BITACORAS_DEF", ["ID","TITULO","ENCABEZADO","PIE","COLUMNAS_JSON","DETALLES_JSON","ORIENTACION","SPLIT_DIAS"]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(obj.titulo)) {
        sheet.getRange(i+1, 3).setValue(obj.encabezado || "");
        sheet.getRange(i+1, 4).setValue(obj.pie || "");
        sheet.getRange(i+1, 5).setValue(JSON.stringify(obj.columnas || []));
        sheet.getRange(i+1, 6).setValue(JSON.stringify(obj.detalles || []));
        sheet.getRange(i+1, 7).setValue(obj.orientacion || "portrait");
        sheet.getRange(i+1, 8).setValue(obj.splitDias ? "true" : "false");
        return "OK-update";
      }
    }
    sheet.appendRow([
      Utilities.getUuid(),
      String(obj.titulo),
      String(obj.encabezado || ""),
      String(obj.pie || ""),
      JSON.stringify(obj.columnas || []),
      JSON.stringify(obj.detalles || []),
      String(obj.orientacion || "portrait"),
      obj.splitDias ? "true" : "false"
    ]);
    return "OK-create";
  } catch(e) { return "Error: " + e; }
}

function eliminarBitacora(titulo) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("BITACORAS_DEF");
    if (!sheet) return "Error";
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][1]) === String(titulo)) { sheet.deleteRow(i+1); return "OK"; }
    }
    return "Error: no encontrado";
  } catch(e) { return "Error: "+e; }
}

function guardarLlenadoBitacora(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss, "BITACORAS_DATA", ["ID","BITACORA_TITULO","MES","ANIO","DATOS_JSON"]);
    var data = sheet.getDataRange().getValues();
    var mes = parseInt(obj.mes), anio = parseInt(obj.anio);
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1])===String(obj.titulo) && parseInt(data[i][2])===mes && parseInt(data[i][3])===anio) {
        sheet.getRange(i+1,5).setValue(JSON.stringify(obj.datos||{}));
        return "OK-update";
      }
    }
    sheet.appendRow([Utilities.getUuid(), String(obj.titulo), mes, anio, JSON.stringify(obj.datos||{})]);
    return "OK-create";
  } catch(e) { return "Error: "+e; }
}

function getLlenadoBitacora(titulo, mes, anio) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("BITACORAS_DATA");
    if (!sheet || sheet.getLastRow() < 2) return {};
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1])===String(titulo) && parseInt(data[i][2])===parseInt(mes) && parseInt(data[i][3])===parseInt(anio)) {
        try { return JSON.parse(data[i][4]); } catch(e) { return {}; }
      }
    }
    return {};
  } catch(e) { return {}; }
}


// ══════════════════════════════════════════════════════════════════════
//  MÁQUINAS ↔ PLANES (asociación, modal, historial)
// ══════════════════════════════════════════════════════════════════════

function getMaquinasDelPlanInverso(maquina, ss, dataAct) {
  var shMP = ss.getSheetByName("MAQUINAS_PLANES");
  if (shMP && shMP.getLastRow() > 1) {
    var data = shMP.getDataRange().getValues().slice(1);
    var planes = [];
    data.forEach(function(r){ if(String(r[0])===maquina) planes.push(String(r[1])); });
    if (planes.length > 0) return planes;
  }
  var shPlanes = ss.getSheetByName("PLANES_MP_V2");
  if (shPlanes && shPlanes.getLastRow() > 1) {
    return [String(shPlanes.getDataRange().getValues()[1][1])];
  }
  return ["SIN_PLAN"];
}

function getMaquinasDelPlan(clavePlan, ss, dataAct) {
  var shMP = ss.getSheetByName("MAQUINAS_PLANES");
  if (shMP && shMP.getLastRow() > 1) {
    var data = shMP.getDataRange().getValues().slice(1);
    var maquinas = [];
    data.forEach(function(r){
      if(String(r[1]).toUpperCase() === String(clavePlan).toUpperCase()) {
        var m = String(r[0]).trim();
        if(m && !maquinas.includes(m)) maquinas.push(m);
      }
    });
    if(maquinas.length > 0) return maquinas;
  }
  return [];
}

function getMaquinasDelPlanModal(clavePlan) {
  try {
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var dataAct = ssEst.getSheetByName("ESTANDARES").getDataRange().getValues();
    var todasMaquinas = [];
    for(var i=1;i<dataAct.length;i++){
      var m = String(dataAct[i][3]||"").trim();
      if(m && !todasMaquinas.includes(m)) todasMaquinas.push(m);
    }
    var asociadas = getMaquinasDelPlan(clavePlan, ss, dataAct);
    var disponibles = todasMaquinas.filter(function(m){ return !asociadas.includes(m); });
    return { asociadas: asociadas, disponibles: disponibles };
  } catch(e) { return { asociadas: [], disponibles: [] }; }
}

function getPlanesDeActivo(maquina) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var shPlanes = ss.getSheetByName("PLANES_MP_V2");
    var shMP = _getOrCreate(ss, "MAQUINAS_PLANES", ["MAQUINA","CLAVE_PLAN"]);
    var todosPlanes = [];
    if(shPlanes && shPlanes.getLastRow()>1){
      shPlanes.getDataRange().getValues().slice(1).forEach(function(r){
        todosPlanes.push({clave:String(r[1]),nombre:String(r[2]),frecuencia:parseInt(r[3])||0});
      });
    }
    var clavesPlanesAsoc = [];
    if(shMP.getLastRow()>1){
      shMP.getDataRange().getValues().slice(1).forEach(function(r){
        if(String(r[0]).trim()===maquina) clavesPlanesAsoc.push(String(r[1]).toUpperCase());
      });
    }
    var asociados = todosPlanes.filter(function(p){ return clavesPlanesAsoc.includes(p.clave.toUpperCase()); });
    var disponibles = todosPlanes.filter(function(p){ return !clavesPlanesAsoc.includes(p.clave.toUpperCase()); });
    return { asociados: asociados, disponibles: disponibles };
  } catch(e) { return { asociados: [], disponibles: [] }; }
}

function asociarMaquinaAPlan(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = _getOrCreate(ss, "MAQUINAS_PLANES", ["MAQUINA","CLAVE_PLAN"]);
    var data = sh.getLastRow()>1 ? sh.getDataRange().getValues().slice(1) : [];
    for(var i=0;i<data.length;i++){
      if(String(data[i][0])===obj.maquina && String(data[i][1]).toUpperCase()===obj.clavePlan.toUpperCase()) return "OK-existe";
    }
    sh.appendRow([obj.maquina, obj.clavePlan.toUpperCase()]);
    return "OK";
  } catch(e) { return "Error: "+e; }
}

function desasociarMaquinaDelPlan(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("MAQUINAS_PLANES");
    if(!sh || sh.getLastRow()<2) return "OK";
    var data = sh.getDataRange().getValues();
    for(var i=data.length-1;i>=1;i--){
      if(String(data[i][0])===obj.maquina && String(data[i][1]).toUpperCase()===obj.clavePlan.toUpperCase()){
        sh.deleteRow(i+1);
      }
    }
    return "OK";
  } catch(e) { return "Error: "+e; }
}

function getHistorialActivoMP(nombreMaq) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var shOM = ss.getSheetByName("OM");
    if(!shOM || shOM.getLastRow()<2) return [];
    var data = shOM.getDataRange().getValues();
    var fmt = function(d){ return d instanceof Date ? Utilities.formatDate(d,"GMT-6","dd/MM/yyyy") : String(d||""); };
    var results = [];
    for(var i=1;i<data.length;i++){
      var folio = String(data[i][2]||"");
      var maq = String(data[i][3]||"");
      var estado = String(data[i][17]||"").toUpperCase();
      if(maq !== nombreMaq || estado !== "CERRADA") continue;
      if(!folio.startsWith("MP-")) continue;
      results.push({
        folio: folio,
        fechaInicio: fmt(data[i][9]),
        fechaCierre: fmt(data[i][10]),
        trabajo: String(data[i][13]||""),
        tec: String(data[i][14]||"")
      });
    }
    results.reverse();
    return results.slice(0,10);
  } catch(e) { return []; }
}

function getHistorialAutoActivo(nombreMaq) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var shOM = ss.getSheetByName("OM");
    if(!shOM || shOM.getLastRow()<2) return [];
    var data = shOM.getDataRange().getValues();
    var fmt = function(d){ return d instanceof Date ? Utilities.formatDate(d,"GMT-6","dd/MM/yyyy") : String(d||""); };
    var results = [];
    for(var i=1;i<data.length;i++){
      var folio = String(data[i][2]||"");
      var maq = String(data[i][3]||"");
      var estado = String(data[i][17]||"").toUpperCase();
      var tipo = String(data[i][6]||"").toUpperCase();
      if(maq !== nombreMaq || estado !== "CERRADA") continue;
      if(!folio.startsWith("AU-") && tipo !== "AUTÓNOMO" && tipo !== "AUTONOMO") continue;
      results.push({
        folio: folio,
        fechaInicio: fmt(data[i][9]),
        fechaCierre: fmt(data[i][10]),
        trabajo: String(data[i][13]||""),
        tec: String(data[i][14]||"")
      });
    }
    results.reverse();
    return results.slice(0,10);
  } catch(e) { return []; }
}


// ══════════════════════════════════════════════════════════════════════
//  REFACCIONES / CATÁLOGO DE INSUMOS
// ══════════════════════════════════════════════════════════════════════

function getCatInsumos() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("CAT_INSUMOS");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    return data.slice(1).filter(function(r){ return r[1]; });
  } catch(e) { return []; }
}

function guardarInsumo(obj) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss, "CAT_INSUMOS",
      ["ID","CODIGO","CATEGORIA","DESCRIPCION","UNIDAD","UBICACION","STOCK_ACTUAL",
       "MINIMO","MAXIMO","PUNTO_REORDEN","ESPECIFICACIONES","ESTADO","PROVEEDOR",
       "REFERENCIA","TIPO","URL"]);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toUpperCase() === String(obj.codigo).toUpperCase()) {
        return { ok: false, msg: "El código " + obj.codigo + " ya existe." };
      }
    }
    var lastId = 0;
    for (var j = 1; j < data.length; j++) {
      var id = parseInt(data[j][0]) || 0;
      if (id > lastId) lastId = id;
    }
    sheet.appendRow([
      lastId + 1,
      String(obj.codigo || "").toUpperCase(),
      String(obj.categoria || ""),
      String(obj.descripcion || ""),
      String(obj.unidad || ""),
      String(obj.ubicacion || ""),
      Number(obj.stock || 0),
      Number(obj.minimo || 0),
      Number(obj.maximo || 0),
      Number(obj.punto_reorden || 0),
      String(obj.especificaciones || ""),
      String(obj.estado || "ACTIVO"),
      String(obj.proveedor || ""),
      String(obj.referencia || ""),
      String(obj.tipo || ""),
      String(obj.url || "")
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, msg: String(e) }; }
}

function guardarInsumosLote(lista) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = _getOrCreate(ss, "CAT_INSUMOS",
      ["ID","CODIGO","CATEGORIA","DESCRIPCION","UNIDAD","UBICACION","STOCK_ACTUAL",
       "MINIMO","MAXIMO","PUNTO_REORDEN","ESPECIFICACIONES","ESTADO","PROVEEDOR",
       "REFERENCIA","TIPO","URL"]);
    var data = sheet.getDataRange().getValues();
    var existentes = {};
    var lastId = 0;
    for (var i = 1; i < data.length; i++) {
      existentes[String(data[i][1]).toUpperCase()] = true;
      var id = parseInt(data[i][0]) || 0;
      if (id > lastId) lastId = id;
    }
    var guardados = 0;
    var rows = [];
    for (var k = 0; k < lista.length; k++) {
      var obj = lista[k];
      var cod = String(obj.codigo || "").toUpperCase();
      if (!cod || existentes[cod]) continue;
      existentes[cod] = true;
      lastId++;
      rows.push([
        lastId, cod,
        String(obj.categoria || ""), String(obj.descripcion || ""),
        String(obj.unidad || ""), String(obj.ubicacion || ""),
        Number(obj.stock || 0), Number(obj.minimo || 0),
        Number(obj.maximo || 0), Number(obj.punto_reorden || 0),
        String(obj.especificaciones || ""), String(obj.estado || "ACTIVO"),
        String(obj.proveedor || ""), String(obj.referencia || ""),
        String(obj.tipo || ""), String(obj.url || "")
      ]);
      guardados++;
    }
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 16).setValues(rows);
    }
    return { guardados: guardados };
  } catch(e) { return { guardados: 0, error: String(e) }; }
}


// ══════════════════════════════════════════════════════════════════════
//  CATÁLOGOS, ACTIVOS Y FOLIOS
// ══════════════════════════════════════════════════════════════════════

function getCatalogos() {
  var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
  var sheet = ss.getSheetByName("ESTANDARES");
  var data = sheet.getDataRange().getValues();
  var procesos = [], maquinas = [];
  for (var i = 1; i < data.length; i++) {
    var proc = String(data[i][2] || "").trim();
    var maq = String(data[i][3] || "").trim();
    var grupo = String(data[i][9] || "OTROS").trim();
    if (proc) { procesos.push(proc); if (maq) maquinas.push({proceso: proc, nombre: maq, grupo: grupo}); }
  }
  return { procesos: [...new Set(procesos)].sort(), maquinas: maquinas };
}

function getActivos() {
  var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
  var data = ss.getSheetByName("ESTANDARES").getDataRange().getValues();
  var activos = [];
  for (var i = 1; i < data.length; i++) {
    activos.push({
      id: data[i][0],
      maquina: data[i][3], // Col D
      proceso: data[i][2], // Col C
      foto: data[i][8]     // Col I
    });
  }
  return activos;
}

function getSystemData() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var data = ss.getSheetByName("ESTANDARES").getDataRange().getValues();
    var activos=data.slice(1).filter(function(r){return r[3];}).map(function(r){return[r[0],r[1],r[2],r[3],r[8]];});
    return{activos:activos};
  } catch(e){return{activos:[]}}
}

function getSiguienteFolio(prefijo) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("OM");
  var data = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < data.length; i++) {
    var folioStr = String(data[i][2]);
    if (folioStr.startsWith(prefijo)) {
      var partes = folioStr.split("-");
      if(partes[1]) {
        var num = parseInt(partes[1]);
        if (!isNaN(num) && num > max) max = num;
      }
    }
  }
  return prefijo + "-" + ("0000" + (max + 1)).slice(-4);
}


// ══════════════════════════════════════════════════════════════════════
//  ÓRDENES DE TRABAJO (OT / OM / MP)
// ══════════════════════════════════════════════════════════════════════

function getOrdenPorFolio(folio) {
  if (!folio) return null;
 
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("OM");
    var data  = sheet.getDataRange().getValues();
    var buscar = String(folio).trim().toUpperCase();
 
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim().toUpperCase() === buscar) {
 
        var fechaRaw = data[i][1];
        var fechaStr = "";
        if (fechaRaw instanceof Date && !isNaN(fechaRaw.getTime())) {
          fechaStr = Utilities.formatDate(fechaRaw, "GMT-6", "dd/MM/yyyy HH:mm");
        }
 
        // Intentar convertir la foto de Drive a base64
        var urlFoto   = String(data[i][21] || "");
        var fotoBase64 = "";
 
        if (urlFoto.trim() !== "") {
          try {
            // Extraer el FILE_ID de la URL de Drive
            // Formato: https://drive.google.com/file/d/FILE_ID/view...
            var match = urlFoto.match(/\/d\/([a-zA-Z0-9_-]+)/);
            if (match) {
              var fileId = match[1];
              var file   = DriveApp.getFileById(fileId);
              var blob   = file.getBlob();
              var bytes  = blob.getBytes();
              var mime   = blob.getContentType() || "image/jpeg";
              fotoBase64 = "data:" + mime + ";base64," + Utilities.base64Encode(bytes);
            }
          } catch(eFoto) {
            Logger.log("No se pudo convertir foto a base64: " + eFoto);
            // Si falla, continúa sin foto — no interrumpir el flujo
            fotoBase64 = "";
          }
        }
 
        return {
          fila:        i + 1,
          folio:       String(data[i][2]  || ""),
          maquina:     String(data[i][3]  || ""),
          area:        String(data[i][4]  || ""),
          falla:       String(data[i][5]  || ""),
          tipo:        String(data[i][6]  || ""),
          prioridad:   String(data[i][7]  || ""),
          tm:          String(data[i][11] || "NO").toUpperCase(),
          fecha:       fechaStr,
          msgId:       String(data[i][16] || ""),
          estado:      String(data[i][17] || ""),
          solicitante: String(data[i][18] || ""),
          urlFoto:     urlFoto,      // URL original (por si se necesita)
          fotoBase64:  fotoBase64    // ← imagen lista para <img src="...">
        };
      }
    }
  } catch(e) {
    Logger.log("getOrdenPorFolio error: " + e);
  }
 
  return null;
}

function getOrdenPorFolioCompleta(folio) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var data = ss.getSheetByName("OM").getDataRange().getValues();
    for (var i=1;i<data.length;i++) {
      if (String(data[i][2])===folio) {
        var tareas=[];
        if(folio.startsWith("MP-")){
          var shProg=ss.getSheetByName("PROG_ORDENES");
          if(shProg){
            var dp=shProg.getDataRange().getValues();
            for(var j=1;j<dp.length;j++){
              if(String(dp[j][1])===folio){
                var clavePlan=String(dp[j][3]);
                var shPlanes=ss.getSheetByName("PLANES_MP_V2");
                if(shPlanes){
                  var dpl=shPlanes.getDataRange().getValues();
                  for(var k=1;k<dpl.length;k++){
                    if(String(dpl[k][1])===clavePlan){
                      try{
                        var claveTareas=JSON.parse(dpl[k][4]);
                        var shTareas=ss.getSheetByName("TAREAS_BASE");
                        if(shTareas){
                          var dt=shTareas.getDataRange().getValues();
                          claveTareas.forEach(function(ct){
                            var tf=dt.find(function(r){ return String(r[0])===ct; });
                            if(tf) tareas.push({ clave: String(tf[0]), descripcion: String(tf[1]), frecuencia: parseInt(tf[2])||0 });
                            else tareas.push({clave:ct, descripcion:ct, frecuencia:0});
                          });
                        }
                      }catch(e){}
                      break;
                    }
                  }
                }
                break;
              }
            }
          }
        }
        var fec=data[i][1]; var fecStr=fec instanceof Date?Utilities.formatDate(fec,"GMT-6","dd/MM/yyyy HH:mm"):String(fec||"");
        var fci=data[i][10]; var fciStr=fci instanceof Date?Utilities.formatDate(fci,"GMT-6","dd/MM/yyyy"):String(fci||"");
        return {
          folio:data[i][2], maquina:data[i][3], area:data[i][4], falla:data[i][5],
          tipo:data[i][6], prioridad:data[i][7], fecha:fecStr, tecnico:data[i][14],
          estado:data[i][17], horas:data[i][12], trabajo:data[i][13], firma:data[i][15],
          cierre:fciStr, tareas:tareas
        };
      }
    }
    return null;
  } catch(e) { Logger.log(e); return null; }
}

function guardarOrden(obj) {
  try {
    Logger.log("guardarOrden START — folio: " + obj.folio);
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("OM");
 
    // Calculamos cuántas columnas necesitamos
    // Si Col V (índice 21) no existe aún en la hoja, appendRow la crea automáticamente
    var row = new Array(22).fill("");
    row[0]  = Utilities.getUuid();
    row[1]  = new Date();
    row[2]  = String(obj.folio    || "").toUpperCase();
    row[3]  = String(obj.maquina  || "").toUpperCase();
    row[4]  = String(obj.area     || "").toUpperCase();
    row[5]  = String(obj.falla    || "").toUpperCase();
    row[6]  = String(obj.tipo     || "").toUpperCase();
    row[7]  = String(obj.prioridad|| "").toUpperCase();
    row[8]  = new Date();
    row[11] = String(obj.tm       || "NO").toUpperCase();
    row[17] = "ABIERTA";
    row[18] = String(obj.usuario  || "").toUpperCase();
    row[21] = String(obj.urlFoto  || "");  // Col V
 
    sheet.appendRow(row);
    Logger.log("guardarOrden OK — urlFoto: " + (obj.urlFoto || "vacío"));
    return "OK";
  } catch(e) {
    Logger.log("guardarOrden ERROR: " + e.toString());
    // Lanzar la excepción para que el withFailureHandler del HTML la capture
    throw new Error("guardarOrden falló: " + e.toString());
  }
}

function finalizarGestion(obj, base64Image) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("OM");
 
  var checkEstado = sheet.getRange(obj.fila, 18).getValue();
  if (checkEstado !== "ABIERTA") return "YA_PROCESADA";
 
  var ahora      = new Date();
  var fechaTxt   = Utilities.formatDate(ahora, "GMT-6", "dd/MM/yyyy HH:mm");
  var estadoFinal = (obj.accion === "ACEPTAR") ? "INICIADO" : "CANCELADO";
 
  sheet.getRange(obj.fila, 18).setValue(estadoFinal); // Col R
  sheet.getRange(obj.fila, 10).setValue(ahora);        // Col J = fecha inicio
  sheet.getRange(obj.fila, 15).setValue(obj.tecnico);  // Col O = técnico
 
  // ── Leer chatId y threadId desde Col Q ──
  var msgIdRaw = String(obj.msgId || "").trim();
  var chatIdDest, threadIdDest, msgIdLimpio;
 
  if (msgIdRaw.indexOf("|") !== -1) {
    var partes   = msgIdRaw.split("|");
    chatIdDest   = partes[0];
    threadIdDest = partes[1];
    msgIdLimpio  = partes[2];
  } else {
    var serie    = String(obj.folio).split("-")[0].toUpperCase();
    var destino  = DESTINOS_TELEGRAM[serie] || DESTINOS_TELEGRAM["OM"];
    chatIdDest   = destino.chat_id;
    threadIdDest = String(destino.thread_id);
    msgIdLimpio  = msgIdRaw;
    Logger.log("finalizarGestion: fallback DESTINOS_TELEGRAM para " + obj.folio);
  }
 
  // ── Borrar mensaje original ──
  if (msgIdLimpio && msgIdLimpio !== "0" && msgIdLimpio !== "") {
    try {
      UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/deleteMessage", {
        method: "post", contentType: "application/json",
        payload: JSON.stringify({ chat_id: String(chatIdDest), message_id: parseInt(msgIdLimpio) }),
        muteHttpExceptions: true
      });
    } catch(e) { Logger.log("finalizarGestion: no se pudo borrar — " + e); }
  }
 
  // ── Caption ──
  var emoji    = (estadoFinal === "INICIADO") ? "▶️" : "❌";
  var etiqueta = (estadoFinal === "INICIADO") ? "INICIADA" : "CANCELADA";
 
  var caption = emoji + " ORDEN " + etiqueta + " — " + obj.folio + "\n" +
                "📅 " + fechaTxt + "\n" +
                "👤 Gestionó: " + String(obj.user || "").toUpperCase() + "\n" +
                "🛠 Técnico: " + (obj.tecnico || "Sin asignar");
 
  // ── Reenviar foto en la pestaña correcta — SIN botón gestionar ──
  var blob = Utilities.newBlob(
    Utilities.base64Decode(base64Image.split(",")[1]),
    "image/png", "orden_" + obj.folio + ".png"
  );
 
  var nuevoMsgId = _enviarFotoGrupo(chatIdDest, threadIdDest, blob, caption, null); // null = sin botón
 
  if (nuevoMsgId) {
    sheet.getRange(obj.fila, 17).setValue(chatIdDest + "|" + threadIdDest + "|" + nuevoMsgId);
  }
 
  return "OK";
}

// ── Variables por tipo de orden ──
var THREAD_ID_OM  = "5980";
var CHAT_ID_OM    = "-1003025160907";   // tu CHAT_ID actual

var THREAD_ID_OT  = "596";
var CHAT_ID_OT    = "-1003149278847";

var THREAD_ID_OH  = "436";
var CHAT_ID_OH    = "-1003511639155";

function enviarTelegramConBotones(base64Image, folio) {
  try {
    Logger.log("enviarTelegramConBotones START — folio: " + folio);
 
    var tipo = folio.startsWith("OT") ? "OT" : folio.startsWith("OH") ? "OH" : "OM";
    var chatDest   = tipo === "OT" ? CHAT_ID_OT   : tipo === "OH" ? CHAT_ID_OH   : CHAT_ID_OM;
    var threadDest = tipo === "OT" ? THREAD_ID_OT : tipo === "OH" ? THREAD_ID_OH : THREAD_ID_OM;
 
    Logger.log("Destino: chat=" + chatDest + " thread=" + threadDest);
 
    var url    = "https://api.telegram.org/bot" + TOKEN + "/sendPhoto";
    var urlApp = ScriptApp.getService().getUrl() + "?v=gestion&f=" + encodeURIComponent(folio);
    var keyboard = { inline_keyboard: [[ { text: "⚙️ GESTIONAR ORDEN", url: urlApp } ]] };
 
    // Verificar que la imagen no esté vacía
    if (!base64Image || base64Image.indexOf(",") === -1) {
      throw new Error("base64Image inválida o vacía");
    }
 
    var imageBytes = Utilities.base64Decode(base64Image.split(",")[1]);
    Logger.log("Imagen decodificada — bytes: " + imageBytes.length);
 
    var payload = {
      'chat_id':           chatDest,
      'message_thread_id': parseInt(threadDest),
      'photo':             Utilities.newBlob(imageBytes, "image/png", "orden.png"),
      'caption':           "🚨 NUEVA SOLICITUD: " + folio,
      'reply_markup':      JSON.stringify(keyboard)
    };
 
    var response = UrlFetchApp.fetch(url, { 'method': 'post', 'payload': payload });
    var resText  = response.getContentText();
    Logger.log("Telegram respuesta: " + resText);
 
    var resData = JSON.parse(resText);
 
    if (!resData.ok) {
      throw new Error("Telegram rechazó el mensaje: " + resText);
    }
 
    var messageId = resData.result.message_id;
    Logger.log("messageId recibido: " + messageId);
 
    // Guardar chatId|threadId|messageId en Col Q para poder borrar/reenviar después
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("OM");
    var data  = sheet.getRange("C:C").getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == folio) {
        var valorQ = chatDest + "|" + threadDest + "|" + messageId;
        sheet.getRange(i + 1, 17).setValue(valorQ);
        Logger.log("Col Q guardada: " + valorQ + " en fila " + (i + 1));
        break;
      }
    }
 
    Logger.log("enviarTelegramConBotones OK");
 
  } catch(e) {
    Logger.log("enviarTelegramConBotones ERROR: " + e.toString());
    throw new Error("Telegram falló: " + e.toString());
  }
}

function finalizarTrabajo(obj) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("OM");
  var data = sheet.getRange("C:C").getValues();
  var fila = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == obj.folio) { fila = i + 1; break; }
  }
  if (fila == -1) return "Error: Folio no encontrado";

  var folder = DriveApp.getFolderById(ID_CARPETA_ADJUNTOS);
  var blob = Utilities.newBlob(Utilities.base64Decode(obj.firma.split(",")[1]), "image/png", "Firma_" + obj.folio + ".png");
  var urlFirma = folder.createFile(blob).getUrl();

  var ahora = new Date();
  sheet.getRange(fila, 11).setValue(ahora);
  sheet.getRange(fila, 13).setValue(obj.horas);
  sheet.getRange(fila, 14).setValue(obj.trabajo.toUpperCase());
  sheet.getRange(fila, 15).setValue(obj.tecnico);
  sheet.getRange(fila, 16).setValue(urlFirma);
  sheet.getRange(fila, 18).setValue("CERRADA");

  return "OK";
}

function generarOTPreventiva(planObj) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheetOM = ss.getSheetByName("OM");
  var nuevoFolio = getSiguienteFolio("MP");
  var row = new Array(18).fill("");
  row[0] = Utilities.getUuid();
  row[1] = new Date();
  row[2] = nuevoFolio;
  row[3] = planObj.maquina;
  row[4] = "PLANIFICADO";
  row[5] = "MANTTO. PREVENTIVO: " + planObj.tarea;
  row[6] = "PREVENTIVO";
  row[7] = "NORMAL";
  row[8] = new Date();
  row[17] = "ABIERTA";
  sheetOM.appendRow(row);
  return "OK - Folio: " + nuevoFolio;
}

function generarOrdenIndividual(obj) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var ssEst = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var shOM   = ss.getSheetByName("OM");
    var shProg = _getOrCreate(ss,"PROG_ORDENES",["ID","FOLIO","MAQUINA","PLAN","FECHA_PLANEADA","FECHA_REAL","ESTADO","RETRASADA","MES","ANIO"]);
    var shPlanes = ss.getSheetByName("PLANES_MP_V2");

    var clavePlan = "SIN_PLAN";
    if (shPlanes) {
      var planes = getMaquinasDelPlanInverso(obj.maquina, ss, ssEst.getSheetByName("ESTANDARES").getDataRange().getValues());
      if (planes.length > 0) clavePlan = planes[0];
    }

    var folio = getSiguienteFolio("MP");
    var rowOM = new Array(18).fill("");
    rowOM[0]=Utilities.getUuid(); rowOM[1]=new Date(); rowOM[2]=folio;
    rowOM[3]=obj.maquina; rowOM[4]="MANTTO PREVENTIVO"; rowOM[5]="PLAN: "+clavePlan;
    rowOM[6]="PREVENTIVO"; rowOM[7]="NORMAL"; rowOM[8]=new Date(); rowOM[17]="ABIERTA";
    shOM.appendRow(rowOM);

    shProg.appendRow([Utilities.getUuid(),folio,obj.maquina,clavePlan,obj.fecha||"","","ABIERTA",false,parseInt(obj.mes),parseInt(obj.anio),""]);
    return { folio: folio };
  } catch(e) {
    Logger.log("generarOrdenIndividual: "+e);
    throw e;
  }
}

function getOrdenesGestion() {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const data = ss.getSheetByName("OM").getDataRange().getValues();
  return data.slice(1).filter(r => r[17] === "ABIERTA" || r[17] === "INICIADO").map(r => {
    return {
      fecha: r[1] ? Utilities.formatDate(new Date(r[1]), "GMT-6", "dd/MM HH:mm") : "",
      folio: r[2],
      maquina: r[3],
      falla: r[5],
      tipo: r[6],
      prioridad: r[7],
      tecnico: r[14],
      estado: r[17]
    };
  }).reverse();
}

function asignarOrden(obj) {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const sheet = ss.getSheetByName("OM");
  const data = sheet.getRange("C:C").getValues();
  let fila = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == obj.folio) { fila = i + 1; break; }
  }
  if (fila === -1) return "Error: No se encontró el folio";
  sheet.getRange(fila, 10).setValue(new Date());
  sheet.getRange(fila, 15).setValue(obj.tecnico.toUpperCase());
  sheet.getRange(fila, 18).setValue("INICIADO");
  return "OK";
}

function getHistorialOrdenes(diasExtra) {
  const limiteDias = 90 + (diasExtra || 0);
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const data = ss.getSheetByName("OM").getDataRange().getValues();
  const hoy = new Date();
  const msPorDia = 24 * 60 * 60 * 1000;
  return data.slice(1).filter(r => {
    let fechaCierre = r[10] ? new Date(r[10]) : null;
    let estado = r[17];
    if (!fechaCierre) return false;
    let dif = (hoy - fechaCierre) / msPorDia;
    return dif <= limiteDias && (estado === "CERRADA" || estado === "TERMINADA" || estado === "CANCELADO");
  }).map(r => ({
    folio: r[2], maq: r[3], falla: r[5], tipo: r[6],
    cierre: Utilities.formatDate(new Date(r[10]), "GMT-6", "dd/MM/yy"), tec: r[14], estado: r[17]
  })).reverse();
}

function guardarNuevoPlan(obj) {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const sheet = ss.getSheetByName("PLANES_MP");
  sheet.appendRow([Utilities.getUuid(), obj.maquina, obj.tarea, obj.frecuencia, "", "ACTIVO"]);
  return "Plan creado correctamente";
}

function getHistorialActivo(nombreMaq) {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const data = ss.getSheetByName("OM").getDataRange().getValues();
  const fmt = (d) => d instanceof Date ? Utilities.formatDate(d, "GMT-6", "dd/MM/yyyy") : String(d||"");
  return data.slice(1)
    .filter(r => r[3] === nombreMaq && r[17] === "CERRADA")
    .map(r => ({ folio: r[2], fechaInicio: fmt(r[9]), fechaCierre: fmt(r[10]), trabajo: r[13], tec: r[14] }))
    .reverse().slice(0,10);
}

function getOrdenTrabajoHTML() {
  return HtmlService.createHtmlOutputFromFile('OrdenTrabajoHTML').getContent();
}

function subirFotoOrden(base64Data, folio, idCarpeta) {
  try {
    var folder = DriveApp.getFolderById(idCarpeta);
    var blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data.split(",")[1]),
      "image/jpeg",
      "Foto_" + folio + "_" + new Date().getTime() + ".jpg"
    );
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    Logger.log("Error subiendo foto: " + e);
    throw e;
  }
}

// ══════════════════════════════════════════════════════════════════════
//  REPORTE DIARIO DE ÓRDENES ABIERTAS → TELEGRAM
//  Triggers: Lun-Sáb a las 06:30, 14:30 y 22:00
//  Destinos: OM → grupo OM | OT → grupo OT | OH → grupo OH
// ══════════════════════════════════════════════════════════════════════

var ESTADOS_CERRADOS = ["CERRADA", "CERRADO", "TERMINADA", "TERMINADO", "CANCELADA", "CANCELADO"];

var GRUPOS_REPORTE = {
  OM: { chat_id: "-1003025160907", thread_id: 5980 },
  OT: { chat_id: "-1003149278847", thread_id: 596  },
  OH: { chat_id: "-1003511639155", thread_id: 436  }
};

var PRIO_CONFIG = {
  "CRITICA":  { emoji: "🚨", label: "CRÍTICA",  orden: 1 },
  "CRÍTICA":  { emoji: "🚨", label: "CRÍTICA",  orden: 1 },
  "ALTA":     { emoji: "🔴", label: "ALTA",     orden: 2 },
  "MEDIA":    { emoji: "🟡", label: "MEDIA",    orden: 3 },
  "BAJA":     { emoji: "🔵", label: "BAJA",     orden: 4 }
};

function reporteDiarioOrdenes() {
  var hoy = new Date();
  // No ejecutar domingos (0 = domingo)
  if (hoy.getDay() === 0) return;

  var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("OM");
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  var ahora = new Date();

  // ── 1. Leer filas activas y calcular HRS_TM ──────────────────────
  // Agrupar por tipo de orden: OM / OT / OH
  var grupos = { OM: [], OT: [], OH: [] };

  for (var i = 1; i < data.length; i++) {
    var estado = String(data[i][17] || "").trim().toUpperCase();
    if (ESTADOS_CERRADOS.indexOf(estado) >= 0) continue;
    if (!estado) continue; // filas vacías

    var folio    = String(data[i][2]  || "").trim();
    var maquina  = String(data[i][3]  || "").trim();
    var area     = String(data[i][4]  || "").trim();
    var falla    = String(data[i][5]  || "").trim();
    var prioRaw  = String(data[i][7]  || "MEDIA").trim().toUpperCase();
    var fechaRep = data[i][8];   // Col I: fecha reporte (Date)
    var tm       = String(data[i][11] || "NO").trim().toUpperCase();

    // Calcular horas TM solo si máquina está parada
    var hrsTm = "";
    if (tm === "SI" && fechaRep instanceof Date) {
      var diffMs  = ahora - fechaRep;
      var diffHrs = diffMs / (1000 * 60 * 60);
      hrsTm = diffHrs.toFixed(1);
      // Guardar en Col M (índice 12 → columna 13)
      sheet.getRange(i + 1, 13).setValue(parseFloat(hrsTm));
    }

    // Detectar tipo por prefijo del folio
    var tipo = folio.startsWith("OT") ? "OT" : folio.startsWith("OH") ? "OH" : "OM";

    var prio = PRIO_CONFIG[prioRaw] || { emoji: "⚪", label: prioRaw, orden: 5 };

    grupos[tipo].push({
      folio:   folio,
      maquina: maquina,
      area:    area,
      falla:   falla,
      estado:  estado,
      prio:    prio,
      prioRaw: prioRaw,
      tm:      tm,
      hrsTm:   hrsTm
    });
  }

  // ── 2. Enviar un mensaje por cada tipo que tenga órdenes ──────────
  var hora = Utilities.formatDate(ahora, "GMT-6", "HH:mm");
  var fecha = Utilities.formatDate(ahora, "GMT-6", "dd/MM/yyyy");

  Object.keys(grupos).forEach(function(tipo) {
    var lista = grupos[tipo];
    if (lista.length === 0) return; // sin órdenes activas, no mandar nada

    var destino = GRUPOS_REPORTE[tipo];
    var msj = construirMensajeReporte(tipo, lista, fecha, hora);
    enviarMensajeReporteTelegram(destino.chat_id, destino.thread_id, msj);
  });
}

// ── Construir el texto del mensaje ────────────────────────────────────
function construirMensajeReporte(tipo, ordenes, fecha, hora) {
  var titulos = {
    OM: "🔧 ÓRDENES DE MANTENIMIENTO",
    OT: "🏭 ÓRDENES DE TALLER MECANIZADO",
    OH: "🔩 ÓRDENES DE HERRAMENTALES"
  };

  // Ordenar: primero por prioridad, luego por área
  ordenes.sort(function(a, b) {
    if (a.prio.orden !== b.prio.orden) return a.prio.orden - b.prio.orden;
    return a.area.localeCompare(b.area);
  });

  // Agrupar por área
  var porArea = {};
  ordenes.forEach(function(o) {
    if (!porArea[o.area]) porArea[o.area] = [];
    porArea[o.area].push(o);
  });

  var total = ordenes.length;
  var criticas = ordenes.filter(function(o){ return o.prio.orden === 1; }).length;

  // ── Cabecera ──
  var msj = "";
  msj += "━━━━━━━━━━━━━━━━━━━━━━\n";
  msj += titulos[tipo] + "\n";
  msj += "━━━━━━━━━━━━━━━━━━━━━━\n";
  msj += "📅 *" + fecha + "  🕐 " + hora + "*\n";
  msj += "📊 Total activas: *" + total + "*";
  if (criticas > 0) msj += "   🚨 Críticas: *" + criticas + "*";
  msj += "\n";

  // ── Cuerpo por área ──
  var areas = Object.keys(porArea).sort();
  areas.forEach(function(area) {
    var ords = porArea[area];
    // Ordenar dentro del área por prioridad
    ords.sort(function(a, b){ return a.prio.orden - b.prio.orden; });

    msj += "\n📍 *" + area + "*\n";
    msj += "┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄\n";

    ords.forEach(function(o) {
      msj += o.prio.emoji + " *" + o.prio.label + "* — `" + o.folio + "`\n";
      msj += "   ⚙️ " + o.maquina + "\n";
      msj += "   📝 " + o.falla + "\n";
      // Solo mostrar TM si máquina está parada
      if (o.tm === "SI" && o.hrsTm !== "") {
        msj += "   ⏱ *MÁQUINA PARADA: " + o.hrsTm + " hrs*\n";
      }
      msj += "   📌 Estado: _" + o.estado + "_\n";
    });
  });

  // ── Pie ──
  msj += "\n━━━━━━━━━━━━━━━━━━━━━━\n";
  msj += "_CLAVOS NACIONALES CN SA DE CV_";

  return msj;
}

// ── Enviar al grupo/pestaña correcto ──────────────────────────────────
function enviarMensajeReporteTelegram(chatId, threadId, texto) {
  var url = "https://api.telegram.org/bot" + TOKEN + "/sendMessage";
  var payload = {
    "chat_id":            String(chatId),
    "message_thread_id":  parseInt(threadId),
    "text":               texto,
    "parse_mode":         "Markdown"
  };
  try {
    var options = {
      "method":           "post",
      "contentType":      "application/json",
      "payload":          JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    Logger.log("Error reporte Telegram: " + e.toString());
  }
}

function getMisOrdenes(usuario) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("OM");
  var data  = sheet.getDataRange().getValues();

  // Límite: 60 días atrás
  var limite = new Date();
  limite.setDate(limite.getDate() - 60);

  var resultado = [];

  for (var i = 1; i < data.length; i++) {
    var usu   = String(data[i][18] || '').trim().toUpperCase(); // Col S
    var fecha = data[i][1];                                      // Col B: fecha creación

    // Filtrar por usuario
    if (usu !== usuario.trim().toUpperCase()) continue;
    // Filtrar por fecha
    if (!(fecha instanceof Date) || fecha < limite) continue;

    // Calcular hrsTm si TM = SI
    var tm    = String(data[i][11] || 'NO').trim().toUpperCase(); // Col L
    var hrsTm = '';
    if (tm === 'SI') {
      var fechaRep = data[i][8]; // Col I: fecha reporte
      if (fechaRep instanceof Date) {
        var diff = (new Date() - fechaRep) / (1000 * 60 * 60);
        hrsTm = diff.toFixed(1);
      }
    }

    resultado.push({
      folio:    String(data[i][2]  || ''),   // Col C
      maquina:  String(data[i][3]  || ''),   // Col D
      area:     String(data[i][4]  || ''),   // Col E
      falla:    String(data[i][5]  || ''),   // Col F
      tipo:     String(data[i][6]  || ''),   // Col G
      prioridad:String(data[i][7]  || ''),   // Col H
      estado:   String(data[i][17] || ''),   // Col R
      tm:       tm,
      hrsTm:    hrsTm,
      fecha:    Utilities.formatDate(fecha, 'GMT-6', 'dd/MM/yyyy HH:mm')
    });
  }

  // Más reciente primero
  resultado.reverse();
  return resultado;
}

// ──────────────────────────────────────────────────────────────────────
//  validarLogin  — usada por GestionMantoHTML para autenticar
//  Reutiliza la misma hoja USUARIOS de ID_HOJA_ESTANDARES
//  (Col B = nombre, Col C = password)
// ──────────────────────────────────────────────────────────────────────
function validarLogin(usuario, password) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return { ok: false, msg: "Hoja USUARIOS no encontrada" };
 
    var data = sheet.getDataRange().getValues();
 
    for (var i = 1; i < data.length; i++) {
      var nombreHoja = String(data[i][1] || "").trim().toUpperCase();
      var passHoja   = String(data[i][2] || "").trim();
 
      if (nombreHoja === String(usuario).trim().toUpperCase()
          && passHoja === String(password).trim()) {
        return { ok: true };
      }
    }
 
    // Usuario no encontrado o contraseña incorrecta
    return { ok: false, msg: "Credenciales incorrectas" };
 
  } catch(e) {
    Logger.log("validarLogin error: " + e);
    return { ok: false, msg: "Error: " + e.toString() };
  }
}

// ── Destinos Telegram por SERIE del folio ─────────────────────────────
// La serie se detecta con el prefijo del folio: "OM-", "OT-", "OH-", "MP-"
var DESTINOS_TELEGRAM = {
  "OM": { chat_id: "-1003025160907", thread_id: 5980 },
  "OT": { chat_id: "-1003149278847", thread_id: 596  },
  "OH": { chat_id: "-1003511639155", thread_id: 436  },
  "MP": { chat_id: "-1003025160907", thread_id: 5980 }
};
// Si el solicitante no tiene chat_id en Col G, se usa este:
var CHAT_SOLICITANTE_DEFAULT = "625827165";
 
 
// ══════════════════════════════════════════════════════════════════════
//  getOrdenesVivas
//  Lee la hoja OM y devuelve todas las órdenes cuyo estado NO sea
//  CERRADA, TERMINADA ni CANCELADA. Las ordena de más antigua a más
//  reciente por fecha de creación (Col B).
// ══════════════════════════════════════════════════════════════════════
function getOrdenesVivas() {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("OM");
    if (!sheet || sheet.getLastRow() < 2) return [];
 
    var data      = sheet.getDataRange().getValues();
    var EXCLUIDOS = ["CERRADA","CERRADO","TERMINADA","TERMINADO","CANCELADA","CANCELADO"];
 
    var fmt = function(v) {
      if (v instanceof Date && !isNaN(v.getTime()))
        return Utilities.formatDate(v, "GMT-6", "dd/MM/yyyy HH:mm");
      return String(v || "");
    };
    var fmtD = function(v) {
      if (v instanceof Date && !isNaN(v.getTime()))
        return Utilities.formatDate(v, "GMT-6", "dd/MM/yyyy");
      return String(v || "");
    };
 
    var resultado = [];
    for (var i = 1; i < data.length; i++) {
      var folio  = String(data[i][2]  || "").trim();
      var estado = String(data[i][17] || "").trim().toUpperCase();
      if (!folio) continue;
      if (EXCLUIDOS.indexOf(estado) >= 0) continue;
 
      resultado.push({
        fila:         i + 1,
        folio:        folio,
        fecha:        fmt(data[i][1]),
        fechaReporte: fmt(data[i][8]),
        fechaInicio:  fmtD(data[i][9]),
        maquina:      String(data[i][3]  || ""),
        area:         String(data[i][4]  || ""),
        falla:        String(data[i][5]  || ""),
        tipo:         String(data[i][6]  || ""),
        prioridad:    String(data[i][7]  || ""),
        tm:           String(data[i][11] || "NO").toUpperCase(),
        trabajo:      String(data[i][13] || ""),
        tecnico:      String(data[i][14] || ""),
        msgId:        String(data[i][16] || ""),
        estado:       String(data[i][17] || ""),
        solicitante:  String(data[i][18] || ""),
        cuadro:       String(data[i][19] || ""),   // Col T: CUADRO_CAMBIOS (índice 19)
        validado:     String(data[i][20] || ""),   // Col U: VALIDADO (índice 20)
        urlFoto:      String(data[i][21] || "")    // Col V: URL foto Drive (índice 21)
      });
    }
 
    resultado.sort(function(a, b) {
      return _parseFmtFecha(a.fecha) - _parseFmtFecha(b.fecha);
    });
 
    return resultado;
  } catch(e) {
    Logger.log("getOrdenesVivas error: " + e);
    return [];
  }
}
 
// Convierte "dd/MM/yyyy HH:mm" a timestamp para poder ordenar
function _parseFmtFecha(str) {
  try {
    var p = String(str).split(" ");
    var d = p[0].split("/");
    var t = (p[1] || "00:00").split(":");
    return new Date(
      parseInt(d[2]), parseInt(d[1]) - 1, parseInt(d[0]),
      parseInt(t[0]), parseInt(t[1])
    ).getTime();
  } catch(e) { return 0; }
}
 
 
// ══════════════════════════════════════════════════════════════════════
//  actualizarOrdenViva
//  Guarda técnico, trabajo realizado, estado y fechas en la hoja OM.
//  Si hay imagen (base64Image != null), dispara notificaciones Telegram.
//
//  Parámetros que llegan del HTML:
//    obj = {
//      fila          : número de fila en Sheets (base 1)
//      folio         : ej. "OM-0042"
//      tecnico       : nombre del técnico (puede ser "")
//      trabajo       : texto de trabajo realizado
//      nuevoEstado   : "INICIADO" | "TERMINADA" | "CANCELADA" | ""
//      msgId         : id del mensaje original en Telegram (Col Q)
//      solicitante   : nombre del usuario que creó la orden (Col S)
//      usuarioGestion: nombre del usuario que está gestionando ahora
//    }
//    base64Image: dataURL PNG del ticket capturado con html2canvas, o null
// ══════════════════════════════════════════════════════════════════════
function actualizarOrdenViva(obj, base64Image) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("OM");
    var fila  = parseInt(obj.fila);
    var ahora = new Date();
 
    // Verificar que el folio coincide (seguridad)
    var folioHoja = String(sheet.getRange(fila, 3).getValue()).trim();
    if (folioHoja !== String(obj.folio).trim()) {
      return "Error: el folio " + obj.folio + " no está en la fila " + fila;
    }
 
    // ── 1. Técnico (Col O = columna 15) ──
    sheet.getRange(fila, 15).setValue(String(obj.tecnico || ""));
 
    // ── 2. Trabajo realizado (Col N = columna 14) ──
    if (obj.trabajo) {
      sheet.getRange(fila, 14).setValue(String(obj.trabajo).toUpperCase());
    }
 
    // ── 3. Estado y fechas ──
    var nuevo = String(obj.nuevoEstado || "").trim().toUpperCase();
    if (nuevo) {
      var fechaInicioActual = sheet.getRange(fila, 10).getValue();
      var tieneInicio = (fechaInicioActual instanceof Date && !isNaN(fechaInicioActual.getTime()));
 
      if (nuevo === "INICIADO") {
        if (!tieneInicio) sheet.getRange(fila, 10).setValue(ahora);
        sheet.getRange(fila, 18).setValue("INICIADO");
      }
      else if (nuevo === "TERMINADA") {
        if (!tieneInicio) sheet.getRange(fila, 10).setValue(ahora);
        sheet.getRange(fila, 11).setValue(ahora);
        sheet.getRange(fila, 18).setValue("TERMINADA");
      }
      else if (nuevo === "CANCELADA") {
        sheet.getRange(fila, 11).setValue(ahora);
        sheet.getRange(fila, 18).setValue("CANCELADA");
      }
      else if (nuevo === "REVISAR") {
        sheet.getRange(fila, 18).setValue("REVISAR");
      }
      else if (nuevo === "CERRADA") {
        // Leer estado actual ANTES de cambiarlo (para saber si venia de REVISAR)
        var estadoActualEnHoja = String(sheet.getRange(fila, 18).getValue() || "").toUpperCase();
        var vieneDe = estadoActualEnHoja;
 
        var validado = String(sheet.getRange(fila, 21).getValue() || "").toUpperCase();
 
        // Limpiar Col U (VALIDADO) - el ciclo vuelve a empezar
        sheet.getRange(fila, 21).setValue("");
 
        // Registrar en Col T (CUADRO_CAMBIOS)
        var fechaTxtC    = Utilities.formatDate(ahora, "GMT-6", "dd/MM/yyyy HH:mm");
        var cuadroActual = String(sheet.getRange(fila, 20).getValue() || "");
        var entradaLog;
        if (vieneDe === "REVISAR") {
          entradaLog = "> " + fechaTxtC + " \u2014 Re-cerrada por " + String(obj.usuarioGestion || "") + " (pendiente validacion solicitante)";
        } else if (validado === "RECHAZADO") {
          entradaLog = "> " + fechaTxtC + " \u2014 Se paso de RECHAZADO a CERRADA";
        } else {
          entradaLog = "> " + fechaTxtC + " \u2014 Cerrada por " + String(obj.usuarioGestion || "");
        }
        var cuadroNuevo = cuadroActual.trim() !== ""
          ? cuadroActual.trim() + "\n" + entradaLog
          : entradaLog;
        sheet.getRange(fila, 20).setValue(cuadroNuevo);

        // Pasar estadoAnterior al objeto para _notificarCambioEstado
        obj.estadoAnterior = vieneDe;
 
        // Fecha de cierre
        sheet.getRange(fila, 11).setValue(ahora);
        sheet.getRange(fila, 18).setValue("CERRADA");
      }
    }
 
    // ── 4. Notificaciones Telegram ──
    if (base64Image) {
      _notificarCambioEstado(sheet, fila, obj, base64Image, nuevo, ahora);
    }
 
    Logger.log("actualizarOrdenViva OK — folio:" + obj.folio + " nuevoEstado:" + nuevo);
    return "OK";
 
  } catch(e) {
    Logger.log("actualizarOrdenViva error: " + e);
    return "Error: " + e.toString();
  }
}
 
function _notificarCambioEstado(sheet, fila, obj, base64Image, nuevo, ahora) {
  var folio    = String(obj.folio).trim();
  var fechaTxt = Utilities.formatDate(ahora, "GMT-6", "dd/MM/yyyy HH:mm");
 
  // Leer chatId y threadId desde Col Q
  var msgIdRaw = String(obj.msgId || "").trim();
  var chatIdDest, threadIdDest, msgIdLimpio;
 
  if (msgIdRaw.indexOf("|") !== -1) {
    var partes   = msgIdRaw.split("|");
    chatIdDest   = partes[0];
    threadIdDest = partes[1];
    msgIdLimpio  = partes[2];
  } else {
    var serie   = folio.split("-")[0].toUpperCase();
    var destino = DESTINOS_TELEGRAM[serie] || DESTINOS_TELEGRAM["OM"];
    chatIdDest   = destino.chat_id;
    threadIdDest = String(destino.thread_id);
    msgIdLimpio  = msgIdRaw;
    Logger.log("_notificarCambioEstado: fallback DESTINOS_TELEGRAM para " + folio);
  }
 
  var emojis    = { "INICIADO":"▶️", "TERMINADA":"✅", "CANCELADA":"❌" };
  var etiquetas = { "INICIADO":"INICIADA", "TERMINADA":"TERMINADA", "CANCELADA":"CANCELADA" };
  var emoji     = emojis[nuevo]    || "🔄";
  var etiqueta  = etiquetas[nuevo] || nuevo;
 
  var caption = emoji + " ORDEN " + etiqueta + " — " + folio + "\n" +
                "📅 " + fechaTxt + "\n" +
                "👤 Gestionó: " + String(obj.usuarioGestion || "").toUpperCase() + "\n" +
                "🛠 Técnico: "  + (obj.tecnico || "Sin asignar");
  if (obj.trabajo) {
    caption += "\n📝 Trabajo: " + obj.trabajo;
  }
 
  var blob = Utilities.newBlob(
    Utilities.base64Decode(base64Image.split(",")[1]),
    "image/png", "orden_" + folio + ".png"
  );
 
  // Borrar mensaje original
  if (msgIdLimpio && msgIdLimpio !== "0" && msgIdLimpio !== "") {
    _borrarMensajeTelegram(chatIdDest, msgIdLimpio);
  }
 
  // Reenviar SIN botón gestionar en ningún estado
  var nuevoMsgId = _enviarFotoGrupo(chatIdDest, threadIdDest, blob, caption, null);
 
  if (nuevoMsgId) {
    sheet.getRange(fila, 17).setValue(chatIdDest + "|" + threadIdDest + "|" + nuevoMsgId);
  }
 
  var esReCierre = (nuevo === "CERRADA" && String(obj.estadoAnterior || "").toUpperCase() === "REVISAR");
  if (nuevo === "TERMINADA" || nuevo === "CANCELADA" || esReCierre) {
    var chatIdSolicitante = _getChatIdSolicitante(String(obj.solicitante || ""));
    var mensajeBase = esReCierre
      ? "El equipo de mantenimiento revisó el trabajo y lo ha cerrado nuevamente.\nPor favor confirma si el trabajo fue realizado correctamente:"
      : "Revisa el trabajo realizado y confirma:";
    var etiquetaSol = esReCierre ? "RE-CERRADA" : etiqueta;
    var captionSol = "📋 *Orden " + etiquetaSol + ": " + folio + "*\n" +
                     mensajeBase + "\n\n" +
                     (obj.trabajo ? "📝 " + obj.trabajo : "");
    var keyboardSol = {
      inline_keyboard: [[
        { text: "✅ Acepto Trabajo",  callback_data: "VALIDA|ACEPTADO|"  + folio },
        { text: "❌ Rechazo Trabajo", callback_data: "VALIDA|RECHAZADO|" + folio }
      ]]
    };
    _enviarFotoDM(chatIdSolicitante, blob, captionSol, keyboardSol);
  }
}
 
// ══════════════════════════════════════════════════════════════════════
//  CAMBIO en ManttoGS.gs
//  Reemplaza procesarCallbackValidacion() completa
//
//  Al RECHAZAR ahora:
//  1. Cambia estado a REVISAR en Sheets
//  2. Responde al solicitante editando su mensaje
//  3. Borra el mensaje del grupo (el que quedó en TERMINADO)
//  4. Genera un ticket con marca de agua REVISAR (rojo) y lo reenvía
//  5. Manda el texto de aviso al grupo
//  6. Actualiza Col Q con el nuevo msgId
// ══════════════════════════════════════════════════════════════════════
 
function procesarCallbackValidacion(callbackQuery) {
  try {
    var partes    = String(callbackQuery.data || "").split("|");
    if (partes.length < 3) return;
 
    var respuesta = partes[1].toUpperCase();  // "ACEPTADO" o "RECHAZADO"
    var folio     = partes[2].trim();
    var cbId      = callbackQuery.id;
    var chatId    = String(callbackQuery.from.id);
    var msgId     = callbackQuery.message ? callbackQuery.message.message_id : null;
 
    // Buscar la fila y todos los datos de la orden
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("OM");
    var data  = sheet.getDataRange().getValues();
    var fila  = -1;
    var orden = null;
 
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === folio) {
        fila  = i + 1;
        orden = data[i];
        break;
      }
    }
 
    if (fila === -1) {
      _responderCallback(cbId, "❌ No se encontró la orden " + folio);
      return;
    }
 
    // Guardar respuesta en Col U (índice 20)
    sheet.getRange(fila, 21).setValue(respuesta);
 
    if (respuesta === "ACEPTADO") {
      // ── Trabajo aceptado: ciclo completado ────────────────────────
      _responderCallback(cbId, "✅ Trabajo aceptado. ¡Gracias!");
 
      if (msgId) {
        _editarCaption(chatId, msgId,
          "✅ *Trabajo ACEPTADO* — " + folio + "\n¡Gracias por confirmar!");
      }
 
    } else {
      // ── Trabajo rechazado ─────────────────────────────────────────
 
      // 1. Cambiar estado a REVISAR en Col R
      sheet.getRange(fila, 18).setValue("REVISAR");
 
      // 2. Responder al solicitante editando su mensaje DM
      _responderCallback(cbId, "❌ Trabajo rechazado. El equipo será notificado.");
      if (msgId) {
        _editarCaption(chatId, msgId,
          "❌ *Trabajo RECHAZADO* — " + folio + "\nEl equipo de mantenimiento revisará.");
      }
 
      // 3. Determinar destino (grupo + thread) según serie del folio
      var serie   = folio.split("-")[0].toUpperCase();
      var destino = DESTINOS_TELEGRAM[serie] || DESTINOS_TELEGRAM["OM"];
 
      // 4. Leer msgId del grupo desde Col Q para borrar el mensaje TERMINADO
      var msgIdGrupo = String(orden[16] || "").trim();  // Col Q = índice 16
 
      // Col Q puede tener formato "chatId|threadId|msgId" o solo el número
      var msgIdNumero = msgIdGrupo;
      if (msgIdGrupo.indexOf("|") !== -1) {
        var partesMsgId = msgIdGrupo.split("|");
        msgIdNumero = partesMsgId[partesMsgId.length - 1]; // último elemento = msgId
      }
 
      if (msgIdNumero && msgIdNumero !== "0") {
        _borrarMensajeTelegram(destino.chat_id, msgIdNumero);
      }
 
      // 5. Generar ticket con marca de agua REVISAR y reenviarlo
      var ticketBlob = _generarTicketRevisar(orden, folio);
 
      var caption = "⚠️ ORDEN RECHAZADA — REQUIERE REVISIÓN\n" +
                    "Folio: " + folio + "\n" +
                    "El solicitante rechazó el trabajo realizado.\n" +
                    "Favor de verificar y gestionar nuevamente.";
 
      var nuevoMsgId = _enviarFotoGrupo(destino.chat_id, destino.thread_id, ticketBlob, caption, null);
 
      // 6. Actualizar Col Q con el nuevo msgId
      if (nuevoMsgId) {
        var nuevoValorQ = destino.chat_id + "|" + destino.thread_id + "|" + nuevoMsgId;
        sheet.getRange(fila, 17).setValue(nuevoValorQ);
        Logger.log("Col Q actualizada tras rechazo: " + nuevoValorQ);
      }
    }
 
  } catch(e) {
    Logger.log("procesarCallbackValidacion error: " + e);
  }
}


// ══════════════════════════════════════════════════════════════════════
//  NUEVA función auxiliar — genera el ticket de REVISAR como imagen
//  directamente en GAS (sin html2canvas, usando Charts)
//
//  Agrega esta función en ManttoGS.gs junto a los otros helpers
// ══════════════════════════════════════════════════════════════════════
 
function _generarTicketRevisar(ordenRow, folio) {
  // Leer datos de la fila
  var area     = String(ordenRow[4]  || "").toUpperCase();
  var maquina  = String(ordenRow[3]  || "").toUpperCase();
  var falla    = String(ordenRow[5]  || "").toUpperCase();
  var tipo     = String(ordenRow[6]  || "");
  var prioridad= String(ordenRow[7]  || "MEDIA").toUpperCase();
  var tecnico  = String(ordenRow[14] || "Sin asignar").toUpperCase();
  var trabajo  = String(ordenRow[13] || "").toUpperCase();
 
  // Colores según prioridad
  var coloresPrio = {
    "CRITICA": "#b71c1c", "CRÍTICA": "#b71c1c",
    "ALTA":    "#ff9800",
    "MEDIA":   "#fbc02d",
    "BAJA":    "#03a9f4"
  };
  var leyendasPrio = {
    "CRITICA": "PRIORIDAD CRÍTICA — ATENCIÓN INMEDIATA",
    "CRÍTICA": "PRIORIDAD CRÍTICA — ATENCIÓN INMEDIATA",
    "ALTA":    "PRIORIDAD ALTA — DAR PRIORIDAD",
    "MEDIA":   "PRIORIDAD MEDIA — HACERLE UN ESPACIO",
    "BAJA":    "PRIORIDAD BAJA — PROGRAMAR TRABAJO"
  };
  var colorBanner = coloresPrio[prioridad]  || "#1a237e";
  var leyenda     = leyendasPrio[prioridad] || prioridad;
  var txtBanner   = (prioridad === "MEDIA" || prioridad === "BAJA") ? "#000000" : "#ffffff";
 
  var fecha = Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm");
 
  // Construir HTML del ticket como string
  // GAS puede convertir HTML a imagen usando Charts.newDataTable approach,
  // pero la forma más confiable es usar HtmlService + UrlFetch a una API de imagen.
  // Sin embargo, la forma más simple en GAS puro es construir el ticket
  // como texto estructurado con Slides o simplemente como imagen con Charts.
  //
  // ALTERNATIVA SIMPLE Y CONFIABLE: construir el ticket como texto plano formateado
  // y enviarlo como mensaje de texto (no foto) cuando se rechaza.
  // Si quieres foto real, necesitaría un servicio externo.
  //
  // Por eso usamos Charts API de GAS para generar una imagen básica:
 
  var chart = Charts.newBarChart()
    .setTitle("⚠️ REVISAR — " + folio)
    .setDimensions(450, 50)
    .build();
 
  // La forma más práctica en GAS sin servicios externos es crear
  // la imagen usando Google Slides como canvas:
  return _generarTicketConSlides(folio, area, maquina, falla, tipo, prioridad,
                                  tecnico, trabajo, colorBanner, leyenda, txtBanner, fecha);
}
 
 
function _generarTicketConSlides(folio, area, maquina, falla, tipo, prioridad,
                                  tecnico, trabajo, colorBanner, leyenda, txtBanner, fecha) {
  var pres  = SlidesApp.create("TEMP_TICKET_" + folio);
  var slide = pres.getSlides()[0];
 
  // Limpiar slide
  slide.getPageElements().forEach(function(el) { el.remove(); });
 
  var W = 450, H = 600; // dimensiones base en puntos
  pres.setPageWidth(W).setPageHeight(H);
 
  // ── Banner superior ──
  var bannerTop = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, 0, W, 50);
  bannerTop.getFill().setSolidFill(colorBanner);
  bannerTop.getBorder().setTransparent();
  var txtTop = bannerTop.getText();
  txtTop.setText(leyenda);
  txtTop.getTextStyle().setFontSize(11).setBold(true)
        .setForegroundColor(txtBanner);
  txtTop.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
 
  // ── Folio ──
  var folioBox = slide.insertTextBox(folio, 270, 60, 170, 40);
  folioBox.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#cc0000");
  folioBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.RIGHT);
  folioBox.getFill().setTransparent();
  folioBox.getBorder().setTransparent();
 
  // ── Empresa ──
  var empBox = slide.insertTextBox("CLAVOS NACIONALES CN SA DE CV\n" + fecha, 10, 60, 250, 40);
  empBox.getText().getTextStyle().setFontSize(9).setBold(true).setForegroundColor("#333333");
  empBox.getFill().setTransparent();
  empBox.getBorder().setTransparent();
 
  // ── Datos ──
  var datos = "📍 ÁREA: " + area + "\n" +
              "⚙️ MÁQUINA: " + maquina + "\n" +
              "👤 TÉCNICO: " + tecnico + "\n" +
              "🛠 TIPO: " + tipo;
  var datosBox = slide.insertTextBox(datos, 10, 110, W - 20, 100);
  datosBox.getText().getTextStyle().setFontSize(12).setForegroundColor("#000000");
  datosBox.getFill().setTransparent();
  datosBox.getBorder().setTransparent();
 
  // ── Falla ──
  var fallaBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 10, 220, W - 20, 80);
  fallaBox.getFill().setSolidFill("#f9f9f9");
  fallaBox.getBorder().setWeight(1.5).setDashStyle(SlidesApp.DashStyle.SOLID);
  var fallaText = fallaBox.getText();
  fallaText.setText("FALLA:\n" + falla);
  fallaText.getTextStyle().setFontSize(12).setForegroundColor("#000000");
 
  // ── Trabajo realizado ──
  if (trabajo) {
    var trabajoBox = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 10, 310, W - 20, 70);
    trabajoBox.getFill().setSolidFill("#f0fff0");
    trabajoBox.getBorder().setWeight(1.5).setDashStyle(SlidesApp.DashStyle.SOLID);
    var trabajoText = trabajoBox.getText();
    trabajoText.setText("TRABAJO REALIZADO:\n" + trabajo);
    trabajoText.getTextStyle().setFontSize(11).setForegroundColor("#000000");
  }
 
  // ── Marca de agua REVISAR ──
  var wmY = trabajo ? 395 : 320;
  var wm = slide.insertTextBox("REVISAR", 80, wmY, 300, 100);
  wm.getText().getTextStyle()
    .setFontSize(70).setBold(true)
    .setForegroundColor("#cc0000");
  wm.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  wm.setRotation(-30);
  wm.setOpacity ? wm.setOpacity(0.25) : null; // no todos los métodos están disponibles
  wm.getFill().setTransparent();
  wm.getBorder().setTransparent();
 
  // ── Banner inferior ──
  var bannerBot = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 0, H - 50, W, 50);
  bannerBot.getFill().setSolidFill(colorBanner);
  bannerBot.getBorder().setTransparent();
  var txtBot = bannerBot.getText();
  txtBot.setText(leyenda);
  txtBot.getTextStyle().setFontSize(11).setBold(true).setForegroundColor(txtBanner);
  txtBot.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
 
  // ── Exportar como PNG ──
  var slideUrl = "https://docs.google.com/presentation/d/" + pres.getId() +
                 "/export/png?pageid=" + slide.getObjectId();
  var token    = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(slideUrl, {
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
 
  var blob = response.getBlob().setName("ticket_revisar_" + folio + ".png");
 
  // Borrar la presentación temporal
  try { DriveApp.getFileById(pres.getId()).setTrashed(true); } catch(e) {}
 
  return blob;
}
 
// ══════════════════════════════════════════════════════════════════════
//  HELPERS INTERNOS DE TELEGRAM
// ══════════════════════════════════════════════════════════════════════
 
// Borra un mensaje del grupo (para reemplazarlo con la foto nueva)
function _borrarMensajeTelegram(chatId, msgId) {
  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/deleteMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id:    String(chatId),
        message_id: parseInt(msgId)
      }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log("_borrarMensajeTelegram: " + e); }
}
 
// Envía una foto a un grupo/hilo. Devuelve el message_id del mensaje enviado.
function _enviarFotoGrupo(chatId, threadId, blob, caption, keyboard) {
  try {
    var payload = {
      chat_id:  String(chatId),
      photo:    blob,
      caption:  caption
    };
    if (threadId) payload.message_thread_id = parseInt(threadId);
    if (keyboard) payload.reply_markup = JSON.stringify(keyboard);
 
    var res  = UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + TOKEN + "/sendPhoto",
      { method: "post", payload: payload, muteHttpExceptions: true }
    );
    var data = JSON.parse(res.getContentText());
    if (data.ok) return data.result.message_id;
  } catch(e) { Logger.log("_enviarFotoGrupo: " + e); }
  return null;
}
 
// Envía una foto a un chat personal (DM) con teclado inline
function _enviarFotoDM(chatId, blob, caption, keyboard) {
  try {
    var payload = {
      chat_id:      String(chatId),
      photo:        blob,
      caption:      caption,
      parse_mode:   "Markdown",
      reply_markup: JSON.stringify(keyboard)
    };
    UrlFetchApp.fetch(
      "https://api.telegram.org/bot" + TOKEN + "/sendPhoto",
      { method: "post", payload: payload, muteHttpExceptions: true }
    );
  } catch(e) { Logger.log("_enviarFotoDM: " + e); }
}
 
// Responde a un callback_query (quita el "reloj" del botón en Telegram)
function _responderCallback(cbId, texto) {
  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/answerCallbackQuery", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        callback_query_id: String(cbId),
        text: texto
      }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log("_responderCallback: " + e); }
}
 
// Edita el caption de un mensaje (para quitar botones después de responder)
function _editarCaption(chatId, msgId, nuevoCaption) {
  try {
    UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/editMessageCaption", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id:    String(chatId),
        message_id: parseInt(msgId),
        caption:    nuevoCaption,
        parse_mode: "Markdown"
      }),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log("_editarCaption: " + e); }
}
 
// Envía un mensaje de texto a un grupo/hilo
function _enviarMensajeGrupo(chatId, threadId, texto) {
  try {
    var payload = {
      chat_id:    String(chatId),
      text:       texto,
      parse_mode: "Markdown"
    };
    if (threadId) payload.message_thread_id = parseInt(threadId);
    UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch(e) { Logger.log("_enviarMensajeGrupo: " + e); }
}
 
// Busca en la hoja USUARIOS el chat_id personal del solicitante (Col G)
function _getChatIdSolicitante(nombreSolicitante) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return CHAT_SOLICITANTE_DEFAULT;
    var data  = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var nombre = String(data[i][1] || "").trim().toUpperCase(); // Col B
      if (nombre === nombreSolicitante.trim().toUpperCase()) {
        var chatId = String(data[i][6] || "").trim();             // Col G
        return chatId || CHAT_SOLICITANTE_DEFAULT;
      }
    }
  } catch(e) { Logger.log("_getChatIdSolicitante: " + e); }
  return CHAT_SOLICITANTE_DEFAULT;
}