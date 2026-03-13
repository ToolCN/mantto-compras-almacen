////////////////////////////////////////////////////////////////////////////////////////////
////////////                        GENERALES                             //////////////////
////////////////////////////////////////////////////////////////////////////////////////////

var ID_HOJA_ESTANDARES = "1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk";
var ID_HOJA_OM = "1v21_Glgvk3ZV4SYpsMGbqNc97I7MO98BqkwmiJvYvnI";
var TOKEN = "7947767393:AAFmZUcSTnV5gvP6u_UsBcSHlz-0s9x1kSQ";
var CHAT_ID = "-1003025160907";
var ID_CARPETA_ADJUNTOS = "1_GEq--PEfyK1P_X-MJx11H5XDDqpxS58"; //Carpeta para guardar adjuntos en requisiciones
var ID_CARPETA_INSUMOS = "1Lz3LZWGkp0jAhJuPCYeSvjqphSQMGAh7";  //Carpeta para guardar adjuntos en refacciones de almacen
var CHAT_ID_COMPRAS = "-1001321274993";
var CHAT_ID_CALIDAD = "-1003608646187";
var THREAD_ID_SELLOS = "1250";
var ID_HOJA_CALCULO = "1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk";

////////////////////////////////////////////////////////////////////////////////////////////
///////////////// MENSAJES AL GRUPO DE COMPRAS PLANTAS TELEGRAM /////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

function buscarEstatusPedidoBot_Integrado(key) {
  var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
  var sheetPed = ss.getSheetByName("PEDIDOS");
  var sheetOrd = ss.getSheetByName("ORDENES");
  var sheetEnv = ss.getSheetByName("ENVIADO");

  var dataPed = sheetPed.getDataRange().getValues();
  var abiertos = [];
  var todosLosMatches = [];

  // 1. RECORRER PEDIDOS (B=Pedido[1], C=Fecha[2], D=Codigo[3], E=Desc[4], G=Cant[6], H=Unidad[7], I=Estado[8], L=FechaEntrega[11])
  for (var i = 1; i < dataPed.length; i++) {
    var pedID = String(dataPed[i][1]).replace(/\s+/g, "").toUpperCase();
    var codProd = String(dataPed[i][3]).toUpperCase();
    var estado = String(dataPed[i][8]).toUpperCase();

    if (pedID.includes(key) || codProd.includes(key)) {
      
      // Formatear Fechas de Pedidos
      var fReg = dataPed[i][2];
      var fEnt = dataPed[i][11];
      var fRegStr = (fReg instanceof Date) ? Utilities.formatDate(fReg, "GMT-6", "dd/MM/yy") : String(fReg);
      var fEntStr = (fEnt instanceof Date) ? Utilities.formatDate(fEnt, "GMT-6", "dd/MM/yy") : "PENDIENTE";

      var info = {
        id: dataPed[i][1],
        fecha: fRegStr,
        codigo: dataPed[i][3],
        desc: dataPed[i][4],
        cant: dataPed[i][6],
        un: dataPed[i][7],
        estado: dataPed[i][8],
        fechaEntrega: fEntStr
      };
      
      todosLosMatches.push(info);
      
      if (estado !== "CANCELADO" && estado !== "TERMINADO" && estado !== "CERRADO") {
        abiertos.push(info);
      }
    }
  }

  if (abiertos.length > 0) {
    var mensaje = "🔎 *PEDIDOS ABIERTOS ENCONTRADOS:*\n\n";
    return construirCuerpoMensaje(mensaje, abiertos, sheetOrd, sheetEnv);
  } 

  if (todosLosMatches.length > 0) {
    var ultimo = todosLosMatches[todosLosMatches.length - 1];
    var mensajeTerminado = "⚠️ *SIN ÓRDENES ABIERTAS*\n\n";
    mensajeTerminado += "Actualmente no tenemos órdenes abiertas del código `" + ultimo.codigo + " - " + ultimo.desc + "`\n";
    mensajeTerminado += "_(Material correspondiente al pedido " + ultimo.id + " que se encuentra " + ultimo.estado + ")_\n\n";
    mensajeTerminado += "📍 *Último pedido registrado:* \n";
    return construirCuerpoMensaje(mensajeTerminado, [ultimo], sheetOrd, sheetEnv);
  }

  return "❌ No se encontró ningún registro para: `" + key + "`";
}

function construirCuerpoMensaje(msjInicial, listaPedidos, sheetOrd, sheetEnv) {
  var mensaje = msjInicial;
  var dataOrd = sheetOrd.getDataRange().getValues();
  var dataEnv = sheetEnv.getDataRange().getValues();

  listaPedidos.forEach(function(p) {
    // Cabecera del Pedido
    mensaje += "🔹 " + p.fecha + " - " + p.id + "\n";
    mensaje += "📣 *" + p.estado + "*\n";
    mensaje += "`" + p.codigo + "`\n";
    mensaje += "*" + p.desc + "*\n\n";

    // --- ÓRDENES Y PROCESOS ---
    var ordenesMap = {}; 
    for (var j = 1; j < dataOrd.length; j++) {
      if (String(dataOrd[j][1]) === p.id) { 
        var s = dataOrd[j][4]; var n = dataOrd[j][5];
        var oName = s + "." + ("0000" + n).slice(-4);
        if (!ordenesMap[oName]) ordenesMap[oName] = [];
        ordenesMap[oName].push({
          proc: dataOrd[j][11],
          sol: dataOrd[j][13],
          prod: dataOrd[j][14]
        });
      }
    }

    var oKeys = Object.keys(ordenesMap);
    if (oKeys.length > 0) {
      mensaje += "🛠 *Órdenes asociadas:* \n";
      oKeys.forEach(function(name) {
        mensaje += "📌ORD: " + name + "\n";
        ordenesMap[name].forEach(function(r) {
          mensaje += "    🧩 " + r.proc + ": " + Math.round(r.prod).toLocaleString() + " / " + Math.round(r.sol).toLocaleString() + "\n";
        });
      });
    }

    // --- ENVÍOS ---
    mensaje += "\n🚚 *Registro de envíos:* \n";
    var envios = [];
    for (var k = 1; k < dataEnv.length; k++) {
      if (String(dataEnv[k][5]) === p.id) {
        var fEnv = dataEnv[k][2];
        var fStr = (fEnv instanceof Date) ? Utilities.formatDate(fEnv, "GMT-6", "dd/MM/yy") : String(fEnv);
        // Formato: 🔻[FECHA] [ENVIO] [KILOS] [PIEZAS]
        envios.push(" 🔻 " + fStr + " [" + dataEnv[k][12] + "] " + (dataEnv[k][9] || 0).toLocaleString() + " kg - " + (dataEnv[k][10] || 0).toLocaleString() + " pzs");
      }
    }
    
    if (envios.length > 0) {
      mensaje += envios.join("\n") + "\n";
    } else {
      mensaje += " 🚫 No se tiene información de los envíos.\n";
    }

    // Entrega Estimada (Columna L)
    mensaje += "\n🗓 *ENTREGA ESTIMADA:* " + p.fechaEntrega + "\n";
    mensaje += "───────────────────\n\n";
  });

  return mensaje;
}

////////////////////////////////////////////////////////////////////////////////////////////
/////// MENSAJES A CHAT DE TORNILLO N2 PARA DETECTAR ERRORES DE ACERO //////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

// =================================================================================
// MOTOR COMPLETO PARA EL SCRIPT DEL BOT (SCRIPT B)
// =================================================================================

function actualizarSelloDesdeTelegram(chatId, msgIdOrig, nSello) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetProd = ss.getSheetByName("PRODUCCION");
  var data = sheetProd.getDataRange().getValues();
  var f = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][23]) == String(msgIdOrig)) { f = i + 1; break; }
  }

  if (f > -1) {
    var row = data[f-1];
    sheetProd.getRange(f, 15).setValue(nSello); 
    borrarMensajeTelegram(chatId, msgIdOrig);  
    sheetProd.getRange(f, 24).setValue("");    
    SpreadsheetApp.flush();

    var reg = {
      idProdInterno: row[0],
      ordenID: row[2],
      loteFull: row[3],
      maquina: row[4],
      fecha: Utilities.formatDate(new Date(row[5]), "GMT-6", "dd/MM/yy"),
      turno: row[6],
      producido: row[10],
      sello: nSello
    };
    
    // Ejecutar validación
    validarAcerosProduccion([reg], true); 
    
    // ESPERAR a que termine
    SpreadsheetApp.flush();
    Utilities.sleep(1000); // Esperar 1 segundo
    
    // Leer DESPUÉS de la validación
    var dataActualizada = sheetProd.getDataRange().getValues();
    var nuevoIdMsg = "";
    
    for (var i = 1; i < dataActualizada.length; i++) {
      if (String(dataActualizada[i][0]) == String(reg.idProdInterno)) {
        nuevoIdMsg = dataActualizada[i][23]; // Columna 24 (índice 23)
        break;
      }
    }
    
    // SOLO si NO hay mensaje de error nuevo, está correcto
    if (!nuevoIdMsg || nuevoIdMsg === "") {
       enviarMensajeTelegram(CHAT_ID_CALIDAD, "✅ *¡Cambio Exitoso!*\nEl nuevo sello `" + nSello + "` cumple con el acero correcto.", THREAD_ID_SELLOS);
    }
  } else {
    enviarMensajeTelegram(CHAT_ID_CALIDAD, "🙅🏻‍♂️ Ya no se puede modificar este registro.", THREAD_ID_SELLOS);
  }
}

function validarAcerosProduccion(registrosNuevos, esRevalidacion) {
  var ssProd = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var ssMP = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheetProd = ssProd.getSheetByName("PRODUCCION");
  var sheetOrd = ssProd.getSheetByName("ORDENES");
  var sheetMP = ssMP.getSheetByName("ENTRADAS_MP");
  
  var dMP = sheetMP.getDataRange().getValues();
  var dOrd = sheetOrd.getDataRange().getValues();

  registrosNuevos.forEach(reg => {
    var infoO = dOrd.find(o => String(o[0]) == reg.ordenID);
    if (!infoO) return;
    
    var proceso = String(infoO[11]).toUpperCase();
    if (!["FORJA", "PUNTEADO", "ROLADO TORN"].includes(proceso)) return;
    
    var selloIngresado = String(reg.sello).trim();
    var infoM = dMP.find(m => String(m[8]).trim() == selloIngresado);
    var aceroOrden = String(infoO[24]); 
    var productoDesc = infoO[19] + " " + infoO[20] + " X " + infoO[21];
    
    var msjTecnico = "";
    var msjAmigable = "";
    var lanzarAlerta = false;

    if (!infoM) {
      lanzarAlerta = true;
      msjTecnico = "🚫🙅‍♂️ *SELLO NO EXISTE* 🙅‍♂️🚫\n" +
            "🎯 *Lote:* " + reg.loteFull + "\n" +
            "🎰 *Máq:* " + reg.maquina + "\n" +
            "🔩 *Prod:* " + productoDesc + "\n" +
            "💠 *Producción:* " + reg.producido + " kg\n" +
            "#️⃣ *Sello reg:* `" + selloIngresado + "`\n" +
            "❌ El sello no existe en la base de datos de MP.\n" +
            "La orden requiere Acero: *" + aceroOrden + "*";
      
      if (esRevalidacion) {
        msjAmigable = "😰 *El sello nuevo que registraste sigue estando mal*\n\n" +
                      "❌ El sello `" + selloIngresado + "` *NO EXISTE* en la base de datos de MP.\n\n" +
                      "Por favor corrige nuevamente 😬";
      }
    } else {
      var aceroMP = String(infoM[4]);
      var selloMP = String(infoM[8]);
      
      if (!fuzzySteelMatch(aceroOrden, aceroMP)) {
        lanzarAlerta = true;
        msjTecnico = "⚠️📛 *ACERO INCORRECTO* 📛⚠️\n" +
              "🎯 *Lote:* " + reg.loteFull + "\n" +
              "🎰 *Máq:* " + reg.maquina + "\n" +
              "🔩 *Prod:* " + productoDesc + "\n" +
              "💠 *Producción:* " + reg.producido + " kg\n" +
              "#️⃣ *Sello reg:* `" + selloIngresado + "` (" + aceroOrden + ")\n" +
              "🈴 *Base MP:* Sello `" + selloMP + "` (" + aceroMP + ")";
        
        if (esRevalidacion) {
          msjAmigable = "😰 *El sello nuevo que registraste sigue estando mal*\n\n" +
                        "El sello `" + selloIngresado + "` corresponde a un acero *" + aceroMP + "*\n" +
                        "pero la orden requiere un acero *" + aceroOrden + "*\n\n" +
                        "Por favor corrige nuevamente 😬";
        }
      }
    }

    if (lanzarAlerta) {
      // Enviar mensaje técnico AL TOPIC
      var res = enviarMensajeTelegram(CHAT_ID_CALIDAD, msjTecnico, THREAD_ID_SELLOS);
      
      if (res && res.ok) {
        var dP = sheetProd.getDataRange().getValues();
        for(var i = dP.length - 1; i >= 1; i--) {
          if (String(dP[i][0]) == String(reg.idProdInterno)) {
            sheetProd.getRange(i + 1, 24).setValue(res.result.message_id);
            break;
          }
        }
        
        // Si es re-validación, enviar mensaje amigable AL MISMO TOPIC
        if (esRevalidacion && msjAmigable) {
          enviarMensajeTelegram(CHAT_ID_CALIDAD, msjAmigable, THREAD_ID_SELLOS);
        }
      }
    }
  });
}

function enviarMensajeTelegram(chatId, texto, threadId) {
  var url = "https://api.telegram.org/bot" + TOKEN + "/sendMessage";
  var payload = { 
    "chat_id": String(chatId), 
    "text": texto, 
    "parse_mode": "Markdown" 
  };
  
  if (threadId) {
    payload["message_thread_id"] = parseInt(threadId);
  }
  
  try {
    var options = { 
      "method": "post", 
      "contentType": "application/json", 
      "payload": JSON.stringify(payload), 
      "muteHttpExceptions": true 
    };
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("Error enviando mensaje: " + e);
    return { ok: false };
  }
}

function borrarMensajeTelegram(chatId, msgId) {
  var url = "https://api.telegram.org/bot" + TOKEN + "/deleteMessage?chat_id=" + chatId + "&message_id=" + msgId;
  try {
    UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
  } catch(e) {
    Logger.log("Error borrando mensaje: " + e);
  }
}

function fuzzySteelMatch(steelA, steelB) {
  // Extraer todos los dígitos consecutivos más largos
  var extraerNumeros = function(s) {
    var texto = String(s).toUpperCase().trim();
    
    // Buscar secuencias de 4 dígitos (con o sin letras intercaladas)
    // Ejemplo: "10B21" → "1021", "TER \ 10B21 CHQ AK" → "1021"
    var matches = texto.match(/\d+/g);
    
    if (!matches) return null;
    
    // Buscar el número más largo o el primero de 4 dígitos
    for (var i = 0; i < matches.length; i++) {
      if (matches[i].length >= 4) {
        return matches[i].substring(0, 4);
      }
    }
    
    // Si no hay de 4 dígitos, concatenar los que haya
    var concatenado = matches.join('');
    return concatenado.length >= 4 ? concatenado.substring(0, 4) : null;
  };
  
  var numA = extraerNumeros(steelA);
  var numB = extraerNumeros(steelB);
  
  // Si no se pudieron extraer números, no hay coincidencia
  if (!numA || !numB) return false;
  
  // Si los 4 dígitos son iguales, es correcto
  if (numA === numB) return true;
  
  // Grupo especial: 1006, 1008, 1010 son intercambiables
  var grupoEspecial = ["1006", "1008", "1010"];
  if (grupoEspecial.includes(numA) && grupoEspecial.includes(numB)) {
    return true;
  }
  
  return false;
}

////////////////////////////////////////////////////////////////////////////////////////////
///////////// PRUEBAS PARA BOT DE TELEGRAM >>> FUNCION DOPOST <<< //////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

function doPost(e) {
  try {
    var contents = JSON.parse(e.postData.contents);
 
    // ── Callback query (botones inline de Telegram) ──
    if (contents.callback_query) {
      var cbData = String(contents.callback_query.data || "");
      var cbId   = contents.callback_query.id;
 
      // CRÍTICO: responder a Telegram INMEDIATAMENTE antes de cualquier
      // operación lenta. Esto evita que Telegram reintente el callback.
      try {
        UrlFetchApp.fetch("https://api.telegram.org/bot" + TOKEN + "/answerCallbackQuery", {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({ callback_query_id: String(cbId) }),
          muteHttpExceptions: true
        });
      } catch(eAns) { Logger.log("answerCallbackQuery error: " + eAns); }
 
      // Ahora sí procesar la acción (ya sin riesgo de loop)
      if (cbData.startsWith("VALIDA|")) {
        procesarCallbackValidacion(contents.callback_query);
      }
 
      return ContentService.createTextOutput("OK");
    }
 
    // ── Mensaje normal ──
    if (!contents.message) return ContentService.createTextOutput("OK");
 
    var chatId    = contents.message.chat.id;
    var text      = contents.message.text || "";
    var textLower = text.toLowerCase().trim();
 
    registrarUsuarioTelegram(contents);
 
    if (textLower === "test") {
      enviarTextoTelegram_Interno(chatId, "✅ Conexión recuperada.", THREAD_ID_SELLOS);
      return ContentService.createTextOutput("OK");
    }
 
    if (textLower.includes("estatus")) {
      var matchCodigo = text.match(/\d{9}/);
      var matchPedido = text.match(/(QSQ|ZEQ|XER|TEM)\s*(-?)\s*\d+/i);
      var key = matchCodigo ? matchCodigo[0] : (matchPedido ? matchPedido[0].replace(/\s+/g, "").toUpperCase() : "");
      if (key.length > 2) {
        enviarTextoTelegram_Interno(chatId, buscarEstatusPedidoBot_Integrado(key), null);
      }
      return ContentService.createTextOutput("OK");
    }
 
    if (contents.message.reply_to_message) {
      var original = contents.message.reply_to_message;
      if (original.text && (original.text.includes("ACERO INCORRECTO") || original.text.includes("SELLO NO EXISTE"))) {
        actualizarSelloDesdeTelegram(chatId, original.message_id, text.trim());
        return ContentService.createTextOutput("OK");
      }
    }
 
  } catch (err) {
    enviarTextoTelegram_Interno(CHAT_ID_CALIDAD, "❌ Error en doPost: " + err.toString(), THREAD_ID_SELLOS);
  }
 
  return ContentService.createTextOutput("OK");
}

// NUEVA FUNCIÓN COMPLEMENTARIA (No afecta el flujo principal)
function registrarUsuarioTelegram(contents) {
  try {
    var ss = SpreadsheetApp.openById("1v21_Glgvk3ZV4SYpsMGbqNc97I7MO98BqkwmiJvYvnI");
    var sheet = ss.getSheetByName("USUARIOS_TELEGRAM");
    if (!sheet) return; // Si no creaste la pestaña, no hace nada

    var chatId = String(contents.message.chat.id);
    var nombre = contents.message.from.first_name || "Sin nombre";
    var username = contents.message.from.username || "Sin username";
    var fecha = new Date();

    // Revisar si el usuario ya existe en la lista para no duplicar
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]) === chatId) return; // Si ya está el ID, salimos
    }

    // Si es nuevo, lo agregamos al final
    sheet.appendRow([fecha, nombre, "@" + username, chatId]);
  } catch (e) {
    console.error("Error al registrar usuario: " + e.toString());
  }
}

// Función que busca en el Excel (Aislada para que no rompa el doPost)
function buscarDetalleEnHojas(msgIdOriginal) {
  var ss = SpreadsheetApp.openById("1v21_Glgvk3ZV4SYpsMGbqNc97I7MO98BqkwmiJvYvnI");
  var sheetCab = ss.getSheetByName("REQUISICIONES");
  var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
  
  var dataCab = sheetCab.getDataRange().getValues();
  var folio = "";

  for (var i = 1; i < dataCab.length; i++) {
    if (String(dataCab[i][10]).trim() === String(msgIdOriginal).trim()) {
      folio = dataCab[i][2]; 
      break;
    }
  }

  if (folio === "") return "❌ No encontré el folio para este mensaje (ID: " + msgIdOriginal + ")";

  var dataDet = sheetDet.getDataRange().getValues();
  var msj = "📊 *Detalles Folio " + folio + "*:\n\n";
  var item = 1;

  for (var j = 1; j < dataDet.length; j++) {
    if (dataDet[j][1] == folio) {
      msj += item + ". *" + dataDet[j][6] + "*\n";
      msj += "   Cant: " + dataDet[j][3] + " " + dataDet[j][4];
      msj += " | Recibido: " + (dataDet[j][14] || 0);
      msj += "\n   Estatus: `" + dataDet[j][13] + "`\n\n";
      item++;
    }
  }
  return msj;
}

// Función de envío interna (No depende de variables externas)
function enviarTextoTelegram_Interno(chatId, texto) {
  var token = "7947767393:AAFmZUcSTnV5gvP6u_UsBcSHlz-0s9x1kSQ";
  var url = "https://api.telegram.org/bot" + token + "/sendMessage";
  var payload = {
    "chat_id": String(chatId),
    "text": texto,
    "parse_mode": "Markdown"
  };
  UrlFetchApp.fetch(url, { "method": "post", "payload": payload, "muteHttpExceptions": true });
}

function setWebhook() {
  // Pega aquí la URL que copiaste en el paso anterior (Versión 88)
  var urlWebApp = "https://script.google.com/macros/s/AKfycbzZcStBdIBWL6Zmot6u8uBD3X-tST7uCe6hJ1Iqqrv1FLZsH7HtyFjGz1J3GvfCiO8CBQ/exec"; 
  var urlTelegram = "https://api.telegram.org/bot" + TOKEN + "/setWebhook?url=" + urlWebApp;
  var response = UrlFetchApp.fetch(urlTelegram);
  Logger.log(response.getContentText());
}

////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////////   FUNCION DOGET PARA LLAMAR HTML's   /////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

function doGet(e) {
  var v = (e.parameter.v || "").toLowerCase().trim();
  var user = e.parameter.user || "USUARIO";
  var folio = e.parameter.folio || e.parameter.f || "";

  try {
    if (v == "circulante")    return renderPage('CirculanteHTML', user, "");
    if (v == "existencia") return renderPage('ExistenciaHTML', user, "");
    if (v == "refacciones")   return renderPage('RefaccionesHTML', user, "");
    if (v == "mantenimiento") return renderPage('OrdenTrabajoHTML', user, folio);
    if (v == "gestion")       return renderPage('GestionMantoHTML', user, folio);
    ///*****if (v == "requisicion")   return renderPage('RequisicionHTML', user, "");
    ///*****if (v == "compras")       return renderPage('GestionComprasHTML', user, "");
    if (v == "imprimiroc")    return renderPage('OrdenCompraHTML', user, folio);
    ///*****if (e.parameter.v == "cotizar") return renderPage('CotizarHTML', user, "");
    if (v == "gestionmp") return renderPage('GestionMP_HTML', user, "");
    if (v == "cierre")     return renderPage('CierreOrdenHTML', user, folio);
    if (v == "adminmp")    return renderPage('AdminMP_HTML', user, "");
    if (v == "calendariomp") return renderPage('CalendarioMP_HTML', user, "");
    if (v == "main_app") return renderPage('MainAppHTML', user, "");
    if (v == "gestionordenes") return renderPage('GestionOrdenesHTML', user, "");

if (v == "main_embarques") {
    return HtmlService.createHtmlOutputFromFile('MainApp_EmbarquesHTML')
      .addMetaTag('viewport','width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no')
      .setTitle('EMBARQUES')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

    if (v == "main_compras") {
  return HtmlService.createHtmlOutputFromFile('MainApp_ComprasHTML')
    .addMetaTag('viewport','width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no')
    .setTitle('COMPRAS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

    if (v == "recepcion")     return renderPage('RecepcionOCHTML', user, "");

    return ContentService.createTextOutput("⚠️ Vista no encontrada.");
  } catch (error) {
    return ContentService.createTextOutput("❌ Error: " + error.toString());
  }
}

/////////////////////////   FUNCIONES RENDERIZADORAS   //////////////////////////////////////

function renderPage(filename, user, folio) {
  var tmp = HtmlService.createTemplateFromFile(filename);
  tmp.usuario_inyectado = user;
  tmp.folio_inyectado = folio;
  return tmp.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setTitle("TOOL CN | " + filename.replace("HTML",""))
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////  (MainAppHTML)) ///////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

// ══════════════════════════════════════════════════════════════════════
//  USUARIOS
// ══════════════════════════════════════════════════════════════════════
function getUsuariosLogin() {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[1] && r[2]; })
      .map(function(r){
        return {
          nombre:   String(r[1] || "").trim().toUpperCase(),
          password: String(r[2] || "").trim(),
          rol:      String(r[3] || "").trim(),        // Col D: ROL  ← NUEVO
          foto:     String(r[4] || "").trim(),        // Col E: FOTO ← NUEVO
          permisos: String(r[5] || "").trim(),        // Col F: todos los permisos
          telegram: String(r[6] || "").trim(),        // Col G: TELEGRAM_USER
          ult_req:  String(r[7] || "").trim()         // Col H: ULT_REQ
        };
      });
  } catch(e) { Logger.log("getUsuariosLogin: " + e); return []; }
}

function getUltReqId() {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("REQUISICIONES_DETALLE");
    if (!sheet || sheet.getLastRow() < 2) return { ult_req_id: 0 };
    var data = sheet.getDataRange().getValues();
    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      var idVal = parseInt(data[i][0]);
      if (!isNaN(idVal) && idVal > maxId) maxId = idVal;
    }
    if (maxId === 0) maxId = data.length - 1;
    return { ult_req_id: maxId };
  } catch(e) { return { ult_req_id: 0 }; }
}

function actualizarUltReq(nombreUsuario, ultReqId) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return "Error";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === String(nombreUsuario).trim().toUpperCase()) {
        sheet.getRange(i + 1, 8).setValue(String(ultReqId));
        return "OK";
      }
    }
    return "No encontrado";
  } catch(e) { return "Error: " + e.toString(); }
}

function getUsuariosAdmin() {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[1]; })
      .map(function(r){
        return {
          email:    String(r[0] || "").trim(),
          nombre:   String(r[1] || "").trim().toUpperCase(),
          password: String(r[2] || "").trim(),
          rol:      String(r[3] || "").trim(),
          foto:     String(r[4] || "").trim(),        // Col E: FOTO ← NUEVO
          permisos: String(r[5] || "").trim(),
          telegram: String(r[6] || "").trim(),
          ult_req:  String(r[7] || "").trim()
        };
      });
  } catch(e) { return []; }
}

function cambiarPasswordUsuario(nombre, nuevaPassword) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return "Error: hoja USUARIOS no encontrada";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === String(nombre).trim().toUpperCase()) {
        sheet.getRange(i + 1, 3).setValue(String(nuevaPassword).trim()); // Col C = PASSWORD
        return "OK";
      }
    }
    return "Error: usuario no encontrado";
  } catch(e) {
    return "Error: " + e.toString();
  }
}

function guardarUsuario(obj) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return "Error: pestaña USUARIOS no encontrada";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === String(obj.nombre).trim().toUpperCase()) {
        return "Error: el usuario ya existe";
      }
    }
    // Al crear: los permisos son solo los C_ seleccionados (obj.permisos ya viene filtrado)
    sheet.appendRow([
      obj.email    || "",
      obj.nombre.trim().toUpperCase(),
      obj.password || "",
      obj.rol      || "",
      "",
      obj.permisos || "",  // solo C_ tokens
      "",                  // telegram (vacío al crear)
      ""                   // ult_req (vacío al crear)
    ]);
    return "OK";
  } catch(e) { return "Error: " + e.toString(); }
}

function actualizarUsuario(obj) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    if (!sheet) return "Error: pestaña USUARIOS no encontrada";
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === String(obj.nombre).trim().toUpperCase()) {

        // 1. Leer permisos actuales de la celda
        var permActuales = String(data[i][5] || "")
          .split(",")
          .map(function(p){ return p.trim(); })
          .filter(function(p){ return p.length > 0; });

        // 2. Quitar todos los C_ que ya había (de esta app)
        var otrosPerms = permActuales.filter(function(p){ return !p.startsWith("C_"); });

        // 3. Los nuevos C_ vienen del HTML (obj.permisos, ya filtrado)
        var nuevosC = String(obj.permisos || "")
          .split(",")
          .map(function(p){ return p.trim(); })
          .filter(function(p){ return p.startsWith("C_"); });

        // 4. Combinar: otros primero, luego los C_ nuevos
        var permFinal = otrosPerms.concat(nuevosC).join(",");

        // 5. Guardar contraseña y permisos
        sheet.getRange(i + 1, 3).setValue(obj.password || "");
        sheet.getRange(i + 1, 6).setValue(permFinal);
        return "OK";
      }
    }
    return "Error: usuario no encontrado";
  } catch(e) { return "Error: " + e.toString(); }
}

function eliminarUsuario(nombre) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheet = ss.getSheetByName("USUARIOS");
    var data = sheet.getDataRange().getValues();
    for (var i=data.length-1;i>=1;i--) {
      if (String(data[i][1]).toUpperCase()===nombre.toUpperCase()) { sheet.deleteRow(i+1); return "OK"; }
    }
    return "Error: no encontrado";
  } catch(e) { return "Error: "+e; }
}


// ══════════════════════════════════════════════════════════════════════
//  TAREAS DE MANTTO  (reemplaza getSoloTareas y guardarTareaBase)
// ══════════════════════════════════════════════════════════════════════

function getDashboardCompras(mes, anio) {
  try {
    var ss    = SpreadsheetApp.openById(ID_HOJA_OM);
    var shCab = ss.getSheetByName("REQUISICIONES");
    var shDet = ss.getSheetByName("REQUISICIONES_DETALLE");

    if (!shCab || !shDet) return { error: "Hojas no encontradas" };

    var now  = new Date();
    var m    = mes  ? parseInt(mes)  : (now.getMonth() + 1);
    var y    = anio ? parseInt(anio) : now.getFullYear();
    var hoy  = new Date(now.getFullYear(), now.getMonth(), now.getDate());

    var dataCab = shCab.getDataRange().getValues();
    var dataDet = shDet.getDataRange().getValues();

    // Mes anterior
    var mAntDate = new Date(y, m - 2, 1); // m-2 porque m es 1-based
    var mAnt     = mAntDate.getMonth() + 1;
    var yAnt     = mAntDate.getFullYear();

    // ── KPIs: altas hoy y esta semana ──────────────────────────────
    var lunesActual = new Date(hoy);
    lunesActual.setDate(hoy.getDate() - ((hoy.getDay() + 6) % 7));
    var hoyAltas = 0, semAltas = 0;

    for (var i = 1; i < dataCab.length; i++) {
      var fReg = dataCab[i][1];
      if (!(fReg instanceof Date)) continue;
      var fDay = new Date(fReg.getFullYear(), fReg.getMonth(), fReg.getDate());
      if (fDay.getTime() === hoy.getTime()) hoyAltas++;
      if (fDay >= lunesActual) semAltas++;
    }

    // ── Críticas del mes ───────────────────────────────────────────
    var criticas = 0;
    for (var ci = 1; ci < dataCab.length; ci++) {
      var fCrit = dataCab[ci][1];
      if (!(fCrit instanceof Date)) continue;
      if (fCrit.getMonth() + 1 !== m || fCrit.getFullYear() !== y) continue;
      var prio = String(dataCab[ci][4] || "").toUpperCase();
      if (prio === "CRITICA" || prio === "CRÍTICA" || prio === "URGENTE") criticas++;
    }

    // ── Mapa folio → datos cabecera ────────────────────────────────
    var folioMap = {};
    for (var fc = 1; fc < dataCab.length; fc++) {
      var fol = String(dataCab[fc][2] || "").trim();
      var fch = dataCab[fc][1];
      if (!fol || !(fch instanceof Date)) continue;
      folioMap[fol] = {
        fecha:      fch,
        solicitante:String(dataCab[fc][3] || "").trim(),
        prioridad:  String(dataCab[fc][4] || "MEDIA"),
        estado:     String(dataCab[fc][9] || "ABIERTA").trim().toUpperCase() // Col J
      };
    }

    // ── Contadores de estados (partidas) + semanas + aging cotizar ─
    var est = { ABIERTA:0, "COT EN PROCESO":0, COTIZADO:0, AUTORIZADO:0, PARCIAL:0, TERMINADO:0, CANCELADO:0 };
    var aging = { total:0, mas2:0, mas3:0, mas5:0 };

    var diasEnMes = new Date(y, m, 0).getDate();
    var segs = [
      {sem:1, ini:1,  fin:7},
      {sem:2, ini:8,  fin:14},
      {sem:3, ini:15, fin:21},
      {sem:4, ini:22, fin:diasEnMes}
    ];
    var semanas = segs.map(function(s){ return {sem:s.sem, gen:0, cot:0, ini:s.ini, fin:s.fin}; });

    for (var d = 1; d < dataDet.length; d++) {
      var folioDet = String(dataDet[d][1] || "").trim();
      var estadoDet = String(dataDet[d][13] || "").trim().toUpperCase();
      var cab = folioMap[folioDet];
      if (!cab) continue;
      var fCab = cab.fecha;
      if (fCab.getMonth() + 1 !== m || fCab.getFullYear() !== y) continue;

      if (est[estadoDet] !== undefined) est[estadoDet]++;

      var dia = fCab.getDate();
      for (var s = 0; s < semanas.length; s++) {
        if (dia >= semanas[s].ini && dia <= semanas[s].fin) {
          semanas[s].gen++;
          if (["COTIZADO","AUTORIZADO","PARCIAL","TERMINADO"].indexOf(estadoDet) >= 0) semanas[s].cot++;
          break;
        }
      }

      if (estadoDet === "ABIERTA" || estadoDet === "COT EN PROCESO") {
        var fSolo = new Date(fCab.getFullYear(), fCab.getMonth(), fCab.getDate());
        var dp = Math.floor((hoy - fSolo) / 86400000);
        aging.total++;
        if      (dp > 5) aging.mas5++;
        else if (dp > 3) aging.mas3++;
        else if (dp > 2) aging.mas2++;
      }
    }

    // ── reqTracking: seguimiento por requisición (Col J de REQUISICIONES) ──
    // Incluye:
    //   - Mes actual: req no TERMINADO no CANCELADO
    //   - Mes anterior: req no TERMINADO no CANCELADO (marcadas esMesPasado=true)
    var terminadasMes = 0, generadasMes = 0;
    var listaTracking = [];

    for (var r = 1; r < dataCab.length; r++) {
      var fReq = dataCab[r][1];
      if (!(fReq instanceof Date)) continue;
      var rMes = fReq.getMonth() + 1;
      var rAnio = fReq.getFullYear();
      var esMesActual  = (rMes === m && rAnio === y);
      var esMesPasado  = (rMes === mAnt && rAnio === yAnt);
      if (!esMesActual && !esMesPasado) continue;

      var estadoReq = String(dataCab[r][9] || "ABIERTA").trim().toUpperCase(); // Col J
      var folio2    = String(dataCab[r][2] || "").trim();
      var sol2      = String(dataCab[r][3] || "").trim();
      var prio2     = String(dataCab[r][4] || "MEDIA").trim();

      if (esMesActual) {
        generadasMes++;
        if (estadoReq === "TERMINADO") terminadasMes++;
      }

      // Lista de pendientes (no terminadas, no canceladas)
      if (estadoReq !== "TERMINADO" && estadoReq !== "CANCELADO") {
        var fSoloReq = new Date(fReq.getFullYear(), fReq.getMonth(), fReq.getDate());
        var diasReq  = Math.floor((hoy - fSoloReq) / 86400000);
        var fechaStr = Utilities.formatDate(fReq, Session.getScriptTimeZone(), "dd/MM/yyyy");
        listaTracking.push({
          folio:       folio2,
          fecha:       fechaStr,
          solicitante: sol2,
          prioridad:   prio2,
          estado:      estadoReq,
          dias:        diasReq,
          esMesPasado: esMesPasado
        });
      }
    }

    // listaAll incluye TODAS las requisiciones del periodo (terminadas y pendientes)
    // para que el frontend agrupe por rango de edad
    var listaAll = [];

    for (var ra = 1; ra < dataCab.length; ra++) {
      var fReqA = dataCab[ra][1];
      if (!(fReqA instanceof Date)) continue;
      var rMesA  = fReqA.getMonth() + 1;
      var rAnioA = fReqA.getFullYear();
      var esMesActualA  = (rMesA === m  && rAnioA === y);
      var esMesPasadoA  = (rMesA === mAnt && rAnioA === yAnt);
      if (!esMesActualA && !esMesPasadoA) continue;

      var estadoA = String(dataCab[ra][9] || "ABIERTA").trim().toUpperCase();
      if (estadoA === "CANCELADO") continue; // omitir canceladas

      var fSoloA  = new Date(fReqA.getFullYear(), fReqA.getMonth(), fReqA.getDate());
      var diasA   = Math.floor((hoy - fSoloA) / 86400000);
      var fechaA  = Utilities.formatDate(fReqA, Session.getScriptTimeZone(), "dd/MM/yyyy");

      listaAll.push({
        folio:       String(dataCab[ra][2] || "").trim(),
        fecha:       fechaA,
        solicitante: String(dataCab[ra][3] || "").trim(),
        prioridad:   String(dataCab[ra][4] || "MEDIA").trim(),
        estado:      estadoA,
        dias:        diasA,
        esMesPasado: esMesPasadoA
      });
    }

    return {
      hoy:     { total: hoyAltas, critica: criticas, alta: hoyAltas },
      semana:  { total: semAltas },
      estados: est,
      semanas: semanas,
      aging:   aging,
      reqTracking: {
        terminadas: terminadasMes,
        generadas:  generadasMes,
        listaAll:   listaAll  // todas (pendientes + terminadas, sin canceladas)
      }
    };

  } catch(e) {
    Logger.log("getDashboardCompras error: " + e.toString());
    return { error: e.toString() };
  }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  getWebAppUrl()
////////////////////////////////////////////////////////////////////////////////////////////
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getProveedoresAdmin() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("PROVEEDORES");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[0] && String(r[0]).trim(); })
      .map(function(r){
        return {
          nombre:    String(r[0]||"").toUpperCase().trim(),
          material:  String(r[1]||""),
          rfc:       String(r[2]||"").toUpperCase().trim(),
          direccion: String(r[3]||""),
          tel:       String(r[4]||""),
          email:     String(r[5]||""),
          contacto:  String(r[6]||""),
          condicion: String(r[7]||"CONTADO").toUpperCase(),
          credito:   parseInt(r[8]||0)||0,
          activo:    String(r[9]||"SI").toUpperCase()==="NO"?"NO":"SI",
          id:        parseInt(r[10]||0)||0
        };
      });
  } catch(e) { Logger.log("getProveedoresAdmin: "+e); return []; }
}

function guardarLoteProveedores(loteEditar, loteNuevos) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("PROVEEDORES");
    if (!sh) return "Hoja PROVEEDORES no encontrada";
    var data = sh.getDataRange().getValues();

    // 1. EDITAR existentes
    (loteEditar||[]).forEach(function(obj) {
      var idB = parseInt(obj.id);
      for (var i=1; i<data.length; i++) {
        if (parseInt(data[i][10]) === idB) {
          sh.getRange(i+1,1,1,10).setValues([[
            String(obj.nombre   ||"").toUpperCase().trim(),
            obj.material  ||"",
            String(obj.rfc||"").toUpperCase().trim(),
            obj.direccion ||"",
            obj.tel       ||"",
            obj.email     ||"",
            obj.contacto  ||"",
            obj.condicion ||"CONTADO",
            parseInt(obj.credito)||0,
            String(obj.activo||"SI").toUpperCase()==="NO"?"NO":"SI"
          ]]);
          break;
        }
      }
    });

    // 2. AGREGAR nuevos
    var maxId = 0;
    for (var j=1; j<data.length; j++) { var cid=parseInt(data[j][10]||0); if(cid>maxId)maxId=cid; }
    (loteNuevos||[]).forEach(function(p) {
      maxId++;
      sh.appendRow([
        String(p.nombre   ||"").toUpperCase().trim(),
        p.material  ||"",
        String(p.rfc||"").toUpperCase().trim(),
        p.direccion ||"",
        p.tel       ||"",
        p.email     ||"",
        p.contacto  ||"",
        p.condicion ||"CONTADO",
        parseInt(p.credito)||0,
        String(p.activo||"SI").toUpperCase()==="NO"?"NO":"SI",
        maxId
      ]);
    });
    return "OK";
  } catch(e) { Logger.log("guardarLoteProveedores: "+e); return "Error: "+e.toString(); }
}

function eliminarProveedor(id) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("PROVEEDORES");
    if (!sh) return "Hoja PROVEEDORES no encontrada";
    var data = sh.getDataRange().getValues();
    var idB = parseInt(id);
    for (var i=data.length-1; i>=1; i--) {
      if (parseInt(data[i][10])===idB) { sh.deleteRow(i+1); return "OK"; }
    }
    return "No encontrado (ID:"+id+")";
  } catch(e) { return "Error: "+e.toString(); }
}

function getMisRequisiciones(nombreUsuario) {
  try {
    var ss      = SpreadsheetApp.openById(ID_HOJA_OM);
    var shCab   = ss.getSheetByName("REQUISICIONES");
    var shDet   = ss.getSheetByName("REQUISICIONES_DETALLE");

    if (!shCab || !shDet) return [];

    var dataCab = shCab.getDataRange().getValues();   // row 0 = header
    var dataDet = shDet.getDataRange().getValues();   // row 0 = header

    var nombre = String(nombreUsuario || "").trim().toUpperCase();

    // ── 1. Construir mapa folio → { fecha, prioridad, obs, solicita } ──
    // REQUISICIONES: A=ID B=FECHA C=FOLIO D=SOLICITANTE E=PRIORIDAD G=IMPORTACION H=OBS
    var folioMap = {};
    for (var c = 1; c < dataCab.length; c++) {
      var sol  = String(dataCab[c][3] || "").trim().toUpperCase();
      if (sol !== nombre) continue;
      var folio = String(dataCab[c][2] || "");
      var fRaw  = dataCab[c][1];
      folioMap[folio] = {
        fecha:     (fRaw instanceof Date) ? Utilities.formatDate(fRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(fRaw||""),
        prioridad: String(dataCab[c][4] || "MEDIA").toUpperCase(),
        obs:       String(dataCab[c][7] || ""),
        fechaRaw:  fRaw instanceof Date ? fRaw : new Date(0)
      };
    }

    if (!Object.keys(folioMap).length) return [];

    // ── 2. Recopilar partidas de ese solicitante ──────────────────────
    // REQUISICIONES_DETALLE:
    //   A=ID  B=FOLIO  C=PA  D=CANT  E=UNIDAD  F=(libre)  G=DESC
    //   H=PRECIO  I=IVA  J=TOTAL  K=MONEDA  L=PROVEEDOR
    //   M=FOLIO_OC  N=ESTADO  O=RECIBIDO
    var lista = [];
    for (var d = 1; d < dataDet.length; d++) {
      var folioDet = String(dataDet[d][1] || "");
      var cab = folioMap[folioDet];
      if (!cab) continue;

      lista.push({
        folio:     folioDet,
        fecha:     cab.fecha,
        fechaRaw:  cab.fechaRaw,
        prioridad: cab.prioridad,
        obs:       cab.obs,
        pa:        dataDet[d][2] || 1,
        cant:      dataDet[d][3] || 0,
        unidad:    String(dataDet[d][4] || ""),
        desc:      String(dataDet[d][6] || ""),
        precio:    parseFloat(dataDet[d][7]  || 0) || 0,
        iva:       parseFloat(dataDet[d][8]  || 0) || 0,
        total:     parseFloat(dataDet[d][9]  || 0) || 0,
        moneda:    String(dataDet[d][10] || "MXN"),
        proveedor: String(dataDet[d][11] || ""),
        folio_oc:  String(dataDet[d][12] || ""),
        estado:    String(dataDet[d][13] || "ABIERTA").toUpperCase()
      });
    }

    // ── 3. Ordenar: más reciente primero, luego por PA dentro del folio ──
    lista.sort(function(a, b) {
      var td = b.fechaRaw.getTime() - a.fechaRaw.getTime();
      if (td !== 0) return td;
      if (a.folio < b.folio) return 1;
      if (a.folio > b.folio) return -1;
      return parseInt(a.pa) - parseInt(b.pa);
    });

    // Limpiar fechaRaw (no serializable limpiamente)
    lista.forEach(function(r){ delete r.fechaRaw; });

    return lista;

  } catch(e) {
    Logger.log("getMisRequisiciones: " + e.toString());
    return [];
  }
}

////////////////////////////////////////////////////////////////////////////////////////////
////////////////////  PAGOS — MainApp_ComprasHTML  /////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

// ── UTILIDAD: formatear fecha a dd/MM/yyyy ──────────────────────────────────────────────
function _fmtFechaPag(d) {
  if (!d) return "";
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(d);
}

// ── UTILIDAD: generar uniqueid para filas nuevas ────────────────────────────────────────
function _pagGenId(prefix) {
  return (prefix || "PAG") + "-" + new Date().getTime() + "-" + Math.floor(Math.random() * 9999);
}

////////////////////////////////////////////////////////////////////////////////////////////
//  PAGO_PROVEEDORES
////////////////////////////////////////////////////////////////////////////////////////////

function getPagosProveedores() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("PAGO_PROVEEDORES");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[0] && String(r[0]).trim(); })
      .map(function(r){
        return [
          String(r[0]||""),           // A: ID
          _fmtFechaPag(r[1]),         // B: FECHA_FAC
          String(r[2]||""),           // C: DOCTO
          String(r[3]||""),           // D: PROVEEDOR
          String(r[4]||""),           // E: TIPO
          String(r[5]||""),           // F: CONCEPTO
          String(r[6]||""),           // G: TIPO_MOV
          String(r[7]||""),           // H: AREA
          String(r[8]||""),           // I: PRODUCTO
          String(r[9]||""),           // J: CANTIDAD
          String(r[10]||""),          // K: UNIDAD
          String(r[11]||""),          // L: COSTO_UNIT
          String(r[12]||""),          // M: COSTO_TOTAL
          String(r[13]||""),          // N: TOTAL
          String(r[14]||""),          // O: PAGO_PARCIAL
          String(r[15]||""),          // P: SALDO
          String(r[16]||""),          // Q: CONDICION
          String(r[17]||""),          // R: ATRASO
          _fmtFechaPag(r[18]),        // S: F_FACT_ENV
          String(r[19]||"")           // T: AUTORIZO
        ];
      });
  } catch(e) { Logger.log("getPagosProveedores: "+e); return []; }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  CAJA_CHICA
////////////////////////////////////////////////////////////////////////////////////////////

function getCajaChica() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("CAJA_CHICA");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[0] && String(r[0]).trim(); })
      .map(function(r){
        return [
          String(r[0]||""),           // A: ID
          _fmtFechaPag(r[1]),         // B: FECHA
          String(r[2]||""),           // C: CONCEPTO
          String(r[3]||""),           // D: TIPO
          String(r[4]||""),           // E: DIRECCION
          String(r[5]||""),           // F: AREA
          String(r[6]||""),           // G: SUCURSAL
          String(r[7]||""),           // H: DESCRIPCION_GASTO
          String(r[8]||""),           // I: BANCO
          String(r[9]||""),           // J: EFECTIVO
          String(r[10]||""),          // K: ANTICIPO
          String(r[11]||""),          // L: OC
          String(r[12]||""),          // M: T_EGRESO
          String(r[13]||""),          // N: MES
          String(r[14]||"")           // N: ASIGNACION_GASTO
        ];
      });
  } catch(e) { Logger.log("getCajaChica: "+e); return []; }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  GASTOS_CORPO
////////////////////////////////////////////////////////////////////////////////////////////

function getGastosCorpo() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("GASTOS_CORPO");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[0] && String(r[0]).trim(); })
      .map(function(r){
        return [
          String(r[0]||""),           // A: ID
          _fmtFechaPag(r[1]),         // B: FECHA
          String(r[2]||""),           // C: CONCEPTO
          String(r[3]||""),           // D: TIPO
          String(r[4]||""),           // E: DIRECCION
          String(r[5]||""),           // F: AREA
          String(r[6]||""),           // G: SUCURSAL
          String(r[7]||""),           // H: DESCRIPCION_GASTO
          String(r[8]||""),           // I: BANCO
          String(r[9]||""),           // J: EFECTIVO
          String(r[10]||""),          // K: ANTICIPO
          String(r[11]||"")           // L: COMENTARIOS
        ];
      });
  } catch(e) { Logger.log("getGastosCorpo: "+e); return []; }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  CONCEPTOS  (combobox compartido)
////////////////////////////////////////////////////////////////////////////////////////////

function getConceptosPagos() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName("CONCEPTOS");
    if (!sh || sh.getLastRow() < 2) return [];
    return sh.getDataRange().getValues().slice(1)
      .filter(function(r){ return r[1] && String(r[1]).trim(); })
      .map(function(r){
        return {
          ID:        String(r[0]||""),
          CONCEPTO:  String(r[1]||"").toUpperCase().trim(),
          CATEGORIA: String(r[2]||"").toUpperCase().trim()
        };
      });
  } catch(e) { Logger.log("getConceptosPagos: "+e); return []; }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  GUARDAR LOTE (editar + nuevos) — compartido para los 3 módulos
////////////////////////////////////////////////////////////////////////////////////////////

function guardarLotePagos(modulo, loteEditar, loteNuevos) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheetName = modulo; // 'PAGO_PROV' → se busca hoja 'PAGO_PROVEEDORES'
    var mapNombre = {
      'PAGO_PROV':    'PAGO_PROVEEDORES',
      'CAJA_CHICA':   'CAJA_CHICA',
      'GASTOS_CORPO': 'GASTOS_CORPO'
    };
    var nombre = mapNombre[modulo];
    if (!nombre) return "Módulo desconocido: " + modulo;

    var sh = ss.getSheetByName(nombre);
    if (!sh) return "Hoja no encontrada: " + nombre;

    var cols = _pagCols(modulo);
    var data = sh.getDataRange().getValues();

    // 1. EDITAR existentes — buscar por columna A (ID) y sobreescribir fila
    (loteEditar || []).forEach(function(obj) {
      var idBuscar = String(obj.ID || "").trim();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0] || "").trim() === idBuscar) {
          var rowVals = cols.map(function(c){ return String(obj[c]||"").toUpperCase().trim(); });
          sh.getRange(i + 1, 1, 1, rowVals.length).setValues([rowVals]);
          break;
        }
      }
    });

    // 2. AGREGAR nuevos
    (loteNuevos || []).forEach(function(obj) {
      var rowVals = cols.map(function(c){ return String(obj[c]||"").toUpperCase().trim(); });
      sh.appendRow(rowVals);
    });

    return "OK";
  } catch(e) { Logger.log("guardarLotePagos: "+e); return "Error: " + e.toString(); }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  ELIMINAR FILA — busca por ID (col A) y elimina la fila
////////////////////////////////////////////////////////////////////////////////////////////

function eliminarPagosRow(modulo, id) {
  try {
    var mapNombre = {
      'PAGO_PROV':    'PAGO_PROVEEDORES',
      'CAJA_CHICA':   'CAJA_CHICA',
      'GASTOS_CORPO': 'GASTOS_CORPO'
    };
    var nombre = mapNombre[modulo];
    if (!nombre) return "Módulo desconocido";
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sh = ss.getSheetByName(nombre);
    if (!sh) return "Hoja no encontrada: " + nombre;
    var data = sh.getDataRange().getValues();
    var idStr = String(id || "").trim();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0] || "").trim() === idStr) {
        sh.deleteRow(i + 1);
        return "OK";
      }
    }
    return "No encontrado (ID: " + id + ")";
  } catch(e) { return "Error: " + e.toString(); }
}

////////////////////////////////////////////////////////////////////////////////////////////
//  HELPER: orden de columnas por módulo
////////////////////////////////////////////////////////////////////////////////////////////

function _pagCols(modulo) {
  if (modulo === 'PAGO_PROV') return [
    'ID','FECHA_FAC','DOCTO','PROVEEDOR','TIPO','CONCEPTO','TIPO_MOV',
    'AREA','PRODUCTO','CANTIDAD','UNIDAD','COSTO_UNIT','COSTO_TOTAL',
    'TOTAL','PAGO_PARCIAL','SALDO','CONDICION','ATRASO','F_FACT_ENV','AUTORIZO'
  ];
  if (modulo === 'CAJA_CHICA') return [
    'ID','FECHA','CONCEPTO','TIPO','DIRECCION','AREA','SUCURSAL',
    'DESCRIPCION_GASTO','BANCO','EFECTIVO','ANTICIPO','OC','T_EGRESO',
    'MES','ASIGNACION_GASTO'
  ];
  if (modulo === 'GASTOS_CORPO') return [
    'ID','FECHA','CONCEPTO','TIPO','DIRECCION','AREA','SUCURSAL',
    'DESCRIPCION_GASTO','BANCO','EFECTIVO','ANTICIPO','COMENTARIOS'
  ];
  return [];
}

////////////////////////////////////////////////////////////////////////////////////////////
///////////  FORMULARIO PARA GENERAR REQUISICION DE COMPRA (RequisicionHTML) ///////////////
////////////////////////////////////////////////////////////////////////////////////////////

function getSiguienteFolioReq() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheet = ss.getSheetByName("REQUISICIONES");
    var data = sheet.getDataRange().getValues();
    var max = 0;
    // Empezamos desde i=1 para saltar encabezados
    for (var i = 1; i < data.length; i++) {
      var folioStr = String(data[i][2]); // Columna C
      if (folioStr && folioStr.indexOf("-") !== -1) {
        var num = parseInt(folioStr.split("-")[1]);
        if (!isNaN(num) && num > max) max = num;
      }
    }
    return "R-" + ("00000" + (max + 1)).slice(-5);
  } catch(e) {
    return "R-00001"; // Folio de emergencia si falla la lectura
  }
}

function guardarRequisicion(obj, archivoData) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheetCab = ss.getSheetByName("REQUISICIONES");
    var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
    
    var urlArchivo = "";
    
    // Si hay archivo, intentar guardarlo
    if (archivoData && archivoData.base64 && archivoData.base64 !== "") {
      try {
        var folder = DriveApp.getFolderById(ID_CARPETA_ADJUNTOS);
        var blob = Utilities.newBlob(Utilities.base64Decode(archivoData.base64), archivoData.type, archivoData.name);
        var file = folder.createFile(blob);
        // Cambiar permisos para que cualquiera con el link lo vea (opcional)
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        urlArchivo = file.getUrl();
      } catch(errDrive) {
        return "Error al subir archivo a Drive: " + errDrive.toString();
      }
    }

    // Cabecera (Columnas A a K)
    sheetCab.appendRow([
      Utilities.getUuid(), 
      new Date(), 
      obj.folio.toUpperCase(), 
      obj.solicitante.toUpperCase(), 
      obj.prioridad.toUpperCase(), 
      "PLANTA", 
      obj.importacion.toUpperCase(), 
      obj.observaciones.toUpperCase(), 
      urlArchivo, 
      "ABIERTA",
      obj.msg_id || "" 
    ]);

    // Partidas (Columnas A a N)
    obj.partidas.forEach(function(p, index) {
      sheetDet.appendRow([
        Utilities.getUuid(),           
        obj.folio.toUpperCase(),        
        index + 1,                      
        p.cantidad,                     
        p.unidad.toUpperCase(),         
        "-",                            
        p.descripcion.toUpperCase(),    
        "",                             
        "",                             
        "",                             
        "",                             
        "",                             
        "",                             
        "ABIERTA"                      
      ]);
    });

    return "OK";
  } catch(e) {
    return "Error general en servidor: " + e.toString();
  }
}

// Envío a Grupo con gestión de errores
function enviarTelegramReq(base64Image, folio) {
  try {
    var url = "https://api.telegram.org/bot" + TOKEN + "/sendPhoto";
    var payload = {
      'chat_id': CHAT_ID_COMPRAS,
      'message_thread_id': 27831, 
      'photo': Utilities.newBlob(Utilities.base64Decode(base64Image.split(",")[1]), "image/png", "req.png"),
      'caption': "📦 NUEVA REQUISICIÓN: " + folio
    };
    UrlFetchApp.fetch(url, { 'method': 'post', 'payload': payload, 'muteHttpExceptions': true });
  } catch(e) { console.log("Fallo Telegram Grupo: " + e); }
}

// Envío a Privado con gestión de errores
function enviarTelegramPrivadoReq(base64Image, folio, chatId) {
  try {
    var url = "https://api.telegram.org/bot" + TOKEN + "/sendPhoto";
    var payload = {
      'chat_id': chatId,
      'photo': Utilities.newBlob(Utilities.base64Decode(base64Image.split(",")[1]), "image/png", "req_priv.png"),
      'caption': "🔔 *ESTADO DE TU REQUISICIÓN:* " + folio + "\n\n📍 *Estado actual:* SOLICITADO\nTe notificaré por aquí cada avance.",
      'parse_mode': 'Markdown'
    };
    var response = UrlFetchApp.fetch(url, { 'method': 'post', 'payload': payload, 'muteHttpExceptions': true });
    var resObj = JSON.parse(response.getContentText());
    return (resObj.ok) ? resObj.result.message_id : ""; 
  } catch(e) { 
    return ""; 
  }
}

////////////////////////////////////////////////////////////////////////////////////////////
///////////////////  GENERAR ORDENES DE COMPRAS (GestionComprasHTML) ///////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

// --- OBTENER PARTIDAS COTIZADAS (MAPEO SEGÚN TU CAPTURA) ---
function getPartidasCotizadas() {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("REQUISICIONES_DETALLE");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var lista = [];
  
  for (var i = 1; i < data.length; i++) {
    // Columna N (índice 13) es ESTADO
    var estado = String(data[i][13] || "").trim().toUpperCase(); 
    
    if (estado === "COTIZADO" || estado === "COTIZADA") {
      lista.push({
        id: String(data[i][0]),      // A: ID_DETALLE
        folioReq: data[i][1],        // B: FOLIO
        cant: data[i][3],            // D: CANTIDAD
        uni: data[i][4],             // E: UNIDAD
        desc: data[i][6],            // G: DESCRIPCION
        precio: data[i][7],          // H: PRECIO_UNITARIO
        iva: data[i][8],             // I: IVA
        total: data[i][9],           // J: TOTAL
        moneda: data[i][10],         // K: MONEDA
        proveedor: data[i][11],      // L: PROVEEDOR
        fila: i + 1
      });
    }
  }
  return lista;
}

function crearOrdenCompra(partidasEditadas, proveedor, moneda, usuario) {
  try {
    var ss     = SpreadsheetApp.openById(ID_HOJA_OM);
    var ssEst  = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
    var sheetOC  = ss.getSheetByName("ORDENES_COMPRA");
    var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
    var sheetCab = ss.getSheetByName("REQUISICIONES");
    var sheetUsr = ssEst.getSheetByName("USUARIOS");

    // ── Mapa nombre→chatId desde USUARIOS col G ──
    var chatMap = {};
    if (sheetUsr) {
      var dataUsr = sheetUsr.getDataRange().getValues();
      for (var u = 1; u < dataUsr.length; u++) {
        var nm = String(dataUsr[u][1] || "").trim().toUpperCase();
        var ch = String(dataUsr[u][6] || "").trim();
        if (nm && ch) chatMap[nm] = ch;
      }
    }

    // 1. Generar Folio OC
    var ultimoFolioVal = sheetOC.getLastRow() > 1
      ? sheetOC.getRange(sheetOC.getLastRow(), 1).getValue() : "OC-0000";
    var num = parseInt(String(ultimoFolioVal).split("-")[1]) || 0;
    var nuevoFolio = "OC-" + ("0000" + (num + 1)).slice(-4);

    var subtotal = 0, ivaTotal = 0, totalGral = 0;
    var foliosRequiAfectados = {};

    // 2. Actualizar partidas → estado AUTORIZADO
    partidasEditadas.forEach(function(p) {
      var fila = p.fila;
      subtotal  += (Number(p.cant) * Number(p.precio));
      ivaTotal  += Number(p.iva);
      totalGral += Number(p.total);

      sheetDet.getRange(fila, 4).setValue(p.cant);
      sheetDet.getRange(fila, 5).setValue(p.uni.toUpperCase());
      sheetDet.getRange(fila, 7).setValue(p.desc.toUpperCase());
      sheetDet.getRange(fila, 8).setValue(p.precio);
      sheetDet.getRange(fila, 9).setValue(p.iva);
      sheetDet.getRange(fila, 10).setValue(p.total);
      sheetDet.getRange(fila, 11).setValue(p.moneda);
      sheetDet.getRange(fila, 12).setValue(p.proveedor.toUpperCase());
      sheetDet.getRange(fila, 13).setValue(nuevoFolio);
      sheetDet.getRange(fila, 14).setValue("AUTORIZADO");

      var folioR = sheetDet.getRange(fila, 2).getValue();
      if (!foliosRequiAfectados[folioR]) foliosRequiAfectados[folioR] = {};
    });

    // 3. Registrar cabecera OC
    sheetOC.appendRow([
      nuevoFolio, new Date(), proveedor.toUpperCase(),
      moneda, subtotal, ivaTotal, totalGral, usuario.toUpperCase(), "EMITIDA"
    ]);

    // 4. Recopilar metadata completa con chatId
    var dataCab    = sheetCab.getDataRange().getValues();
    var dataDetTodo = sheetDet.getDataRange().getValues();

    for (var fR in foliosRequiAfectados) {
      for (var i = 1; i < dataCab.length; i++) {
        if (String(dataCab[i][2]).trim() === String(fR).trim()) {
          var sol = String(dataCab[i][3] || "").trim().toUpperCase();
          foliosRequiAfectados[fR] = {
            solicitante:   sol,
            prioridad:     String(dataCab[i][4] || "Media"),
            observaciones: String(dataCab[i][7] || ""),
            msg_id:        String(dataCab[i][10] || ""),
            chatId:        chatMap[sol] || "",   // ← chatId del solicitante
            partidas:      []
          };
          break;
        }
      }
      // Todas las partidas del folio para ticket completo con colores
      for (var j = 1; j < dataDetTodo.length; j++) {
        if (String(dataDetTodo[j][1]).trim() === String(fR).trim()) {
          var info = foliosRequiAfectados[fR];
          if (info && info.partidas) {
            info.partidas.push({
              cantidad:    dataDetTodo[j][3],
              unidad:      dataDetTodo[j][4],
              descripcion: dataDetTodo[j][6],
              estado:      String(dataDetTodo[j][13] || "ABIERTA")
            });
          }
        }
      }
    }

    return { success: true, folio: nuevoFolio, metaData: foliosRequiAfectados };

  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// FUNCIÓN DE SERVIDOR PARA EDITAR MENSAJE (AUTORIZADO)
function notificarAutorizacionTelegram(base64Image, folio, chatId, msgId) {
  var botUrl = "https://api.telegram.org/bot" + TOKEN + "/";
  var media = {
    'type': 'photo',
    'media': 'attach://photo',
    'caption': "🟠 *ACTUALIZACIÓN:* " + folio + "\n📍 *Estado:* AUTORIZADO",
    'parse_mode': 'Markdown'
  };
  // 1. Editar Imagen
  var payloadMedia = {
    'chat_id': chatId,
    'message_id': msgId,
    'media': JSON.stringify(media),
    'photo': Utilities.newBlob(Utilities.base64Decode(base64Image.split(",")[1]), "image/png", "autorizado.png")
  };
  UrlFetchApp.fetch(botUrl + "editMessageMedia", { 'method': 'post', 'payload': payloadMedia });

  // 2. Enviar Notificación con Sonido
  var payloadAviso = {
    'chat_id': chatId,
    'text': "🚀 *¡Orden de Compra Generada!* \nTu requisición *" + folio + "* ha sido *AUTORIZADA*.\n\n_Ya estamos en proceso de adquisición._",
    'parse_mode': 'Markdown',
    'reply_to_message_id': msgId
  };
  UrlFetchApp.fetch(botUrl + "sendMessage", { 'method': 'post', 'payload': payloadAviso });
}

function getDatosImpresionOC(folioOC) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var dataOC = ss.getSheetByName("ORDENES_COMPRA").getDataRange().getValues();
    var dataDet = ss.getSheetByName("REQUISICIONES_DETALLE").getDataRange().getValues();
    var dataProv = ss.getSheetByName("PROVEEDORES").getDataRange().getValues();
    
    var buscar = String(folioOC).trim().toUpperCase();
    var cabecera = null;

    // 1. Buscar en ORDENES_COMPRA
    for (var i = 1; i < dataOC.length; i++) {
      if (String(dataOC[i][0]).trim().toUpperCase() === buscar) {
        // CORRECCIÓN: Convertir Fecha a String aquí
        var fechaRaw = dataOC[i][1];
        var fechaStr = (fechaRaw instanceof Date) 
          ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") 
          : String(fechaRaw);

        cabecera = {
          folio: String(dataOC[i][0]),
          fecha: fechaStr, 
          proveedorNom: dataOC[i][2],
          moneda: dataOC[i][3],
          subtotal: dataOC[i][4],
          iva: dataOC[i][5],
          total: dataOC[i][6],
          solicita: dataOC[i][7],
          estado: String(dataOC[i][8] || "").toUpperCase() 
        };
        break;
      }
    }
    
    if (!cabecera) return null;

    // 2. Buscar Datos del Proveedor
    var pEx = { tel: "", email: "", contacto: "", condicion: "", credito: "" };
    for (var k = 1; k < dataProv.length; k++) {
      if (String(dataProv[k][0]).trim().toUpperCase() === String(cabecera.proveedorNom).toUpperCase()) {
        pEx.tel = dataProv[k][4];
        pEx.email = dataProv[k][5];
        pEx.contacto = dataProv[k][6];
        pEx.condicion = dataProv[k][7];
        pEx.credito = dataProv[k][8];
        break;
      }
    }
    cabecera.prov = pEx;

    // 3. Buscar Partidas
    var partidas = [];
    for (var j = 1; j < dataDet.length; j++) {
      if (String(dataDet[j][12]).trim().toUpperCase() === buscar) {
        partidas.push({
          desc: dataDet[j][6],
          cant: dataDet[j][3],
          uni: dataDet[j][4],
          precio: dataDet[j][7],
          total: dataDet[j][9]
        });
      }
    }

    return { cabecera: cabecera, partidas: partidas };
  } catch (e) {
    // IMPORTANTE: Devolver el error como string para que el cliente lo vea
    throw new Error(e.toString());
  }
}

// OBTENER ÓRDENES DE LOS ÚLTIMOS 60 DÍAS
function getOrdenesRecientes() {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("ORDENES_COMPRA");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var hoy = new Date();
  var limite = new Date();
  limite.setDate(hoy.getDate() - 60); // CAMBIADO A 60 DÍAS
  
  var lista = [];
  for (var i = data.length - 1; i >= 1; i--) {
    var fechaValue = data[i][1];
    var fechaOC = (fechaValue instanceof Date) ? fechaValue : new Date(fechaValue);
    
    if (!isNaN(fechaOC.getTime()) && fechaOC >= limite) {
      lista.push({
        folio: data[i][0],
        fecha: Utilities.formatDate(fechaOC, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        proveedor: data[i][2],
        moneda: data[i][3],
        total: data[i][6],
        estado: data[i][8],
        fila: i + 1
      });
    }
  }
  return lista;
}

// CANCELAR ORDEN Y REVERTIR PARTIDAS A "COTIZADO"
function cancelarOC(folioOC, filaOC) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheetOC = ss.getSheetByName("ORDENES_COMPRA");
    var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
    
    // 1. Cambiar estado de la OC a CANCELADA
    sheetOC.getRange(filaOC, 9).setValue("CANCELADA"); // Col I
    
    // 2. Buscar partidas vinculadas en DETALLE y revertirlas
    var dataDet = sheetDet.getDataRange().getValues();
    var buscar = String(folioOC).trim().toUpperCase();
    
    for (var j = 1; j < dataDet.length; j++) {
      // Columna M (índice 12) es Folio_OC
      if (String(dataDet[j][12]).trim().toUpperCase() === buscar) {
        sheetDet.getRange(j + 1, 13).setValue(""); // Limpiar Folio_OC (M)
        sheetDet.getRange(j + 1, 14).setValue("COTIZADO"); // Revertir Estado (N)
      }
    }
    return "OK";
  } catch(e) {
    return "Error: " + e.toString();
  }
}


////////////////////////////////////////////////////////////////////////////////////////////
///////////  VISUALIZAR Y EDITAR LOS REQUERIMIENTOS PENDIENTES (CotizarHTML) ///////////////
////////////////////////////////////////////////////////////////////////////////////////////

function obtenerDatosCotizacion() {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
  var sheetProv = ss.getSheetByName("PROVEEDORES");
  var sheetCab = ss.getSheetByName("REQUISICIONES");

  var dataDet = sheetDet.getDataRange().getValues();
  var dataCab = sheetCab.getDataRange().getValues();
  var dataProv = sheetProv.getDataRange().getValues();

  // Mapeo de Cabecera (REQUISICIONES) usando Folio como llave
  var cabMap = {};
  for(var i=1; i<dataCab.length; i++) {
    var fRaw = dataCab[i][1];
    var fechaFormateada = (fRaw instanceof Date) ? Utilities.formatDate(fRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") : fRaw;
    cabMap[dataCab[i][2]] = { 
      fecha: fechaFormateada, 
      solicita: dataCab[i][3], 
      prioridad: dataCab[i][4],
      url: dataCab[i][8] // <-- LÍNEA AGREGADA: Columna I (ARCHIVO_URL)
    };
  }

  var provs = [];
  for(var p=1; p<dataProv.length; p++){
    if(String(dataProv[p][9]).toUpperCase() === "SI"){
      provs.push({ nombre: dataProv[p][0] });
    }
  }

  var lista = [];
  for (var j = dataDet.length - 1; j >= 1; j--) {
    var estado = String(dataDet[j][13]).trim().toUpperCase();
    if (estado === "ABIERTA" || estado === "COT EN PROCESO") {
      var folio = dataDet[j][1];
      var cab = cabMap[folio] || { fecha: "N/A", solicita: "N/A", prioridad: "MEDIA", url: "" };
      
      lista.push({
        fila: j + 1, 
        id: dataDet[j][0], 
        folio: folio, 
        fecha: cab.fecha,
        solicita: cab.solicita,
        pa: dataDet[j][2], 
        cant: dataDet[j][3], 
        uni: dataDet[j][4], 
        desc: dataDet[j][6],
        prioridad: cab.prioridad, 
        url: cab.url, // <-- LÍNEA AGREGADA: Enviamos la URL al HTML
        precio: dataDet[j][7] || 0, 
        iva: dataDet[j][8] || 0,
        total: dataDet[j][9] || 0, 
        moneda: dataDet[j][10] || "MXN", 
        proveedor: dataDet[j][11] || ""
      });
    }
  }
  return { lista: lista, proveedores: provs };
}

function guardarLoteCotizacion(lote) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_OM);
  var ssEst    = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
  var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
  var sheetCab = ss.getSheetByName("REQUISICIONES");
  var sheetUsr = ssEst.getSheetByName("USUARIOS");

  var dataCab  = sheetCab.getDataRange().getValues();
  var dataUsr  = sheetUsr ? sheetUsr.getDataRange().getValues() : [];

  // ── Mapa nombre→chatId desde USUARIOS col G ───────────────────────
  var chatMap = {};
  for (var u = 1; u < dataUsr.length; u++) {
    var nombre = String(dataUsr[u][1] || "").trim().toUpperCase();
    var chat   = String(dataUsr[u][6] || "").trim(); // Col G = TELEGRAM_USER
    if (nombre && chat) chatMap[nombre] = chat;
  }

  var foliosAfectados = {};

  lote.forEach(function(item) {
    var p     = Number(item.precio);
    var prov  = String(item.proveedor || "").trim();
    var estado = (p > 0 && prov !== "") ? "COTIZADO" : "COT EN PROCESO";

    sheetDet.getRange(item.fila, 4, 1, 9).setValues([[
      item.cant, item.uni, "", item.desc, p, item.iva, item.total, item.moneda, prov.toUpperCase()
    ]]);
    sheetDet.getRange(item.fila, 14).setValue(estado);

    if (!foliosAfectados[item.folio]) {
      for (var i = 1; i < dataCab.length; i++) {
        if (String(dataCab[i][2]).trim() === String(item.folio).trim()) {
          var sol = String(dataCab[i][3] || "").trim().toUpperCase();
          foliosAfectados[item.folio] = {
            solicitante:   sol,
            prioridad:     String(dataCab[i][4] || "Media"),
            observaciones: String(dataCab[i][7] || ""),
            msg_id:        String(dataCab[i][10] || ""), // Col K
            chatId:        chatMap[sol] || ""            // ← NUEVO: chatId del solicitante
          };
          break;
        }
      }
    }
  });

  // Recopilar todas las partidas actuales para ticket completo
  var todosLosDetalles = sheetDet.getDataRange().getValues();
  for (var f in foliosAfectados) {
    var partidas = [];
    for (var d = 1; d < todosLosDetalles.length; d++) {
      if (String(todosLosDetalles[d][1]).trim() === String(f).trim()) {
        partidas.push({
          cantidad:    todosLosDetalles[d][3],
          unidad:      todosLosDetalles[d][4],
          descripcion: todosLosDetalles[d][6],
          estado:      String(todosLosDetalles[d][13] || "")
        });
      }
    }
    foliosAfectados[f].partidas = partidas;
  }

  return foliosAfectados;
}

function obtenerInformacionParaImprimirOC(folioOC) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var sheetOC = ss.getSheetByName("ORDENES_COMPRA");
    var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
    var sheetProv = ss.getSheetByName("PROVEEDORES");
    
    // Usamos getDisplayValues para obtener los textos tal cual se ven en las celdas
    var dataOC = sheetOC.getDataRange().getDisplayValues();
    var dataDet = sheetDet.getDataRange().getDisplayValues();
    var dataProv = sheetProv.getDataRange().getDisplayValues();
    
    var buscar = String(folioOC).trim().toUpperCase();
    var cabecera = null;

    // 1. BUSCAR EN ORDENES_COMPRA (Captura 2)
    // Columnas: A:Folio, B:Fecha, C:Proveedor, D:Moneda, E:Subtotal, F:IVA, G:Total, H:Solicita, I:Estado
    for (var i = 1; i < dataOC.length; i++) {
      if (dataOC[i][0].trim().toUpperCase() === buscar) {
        cabecera = {
          folio: dataOC[i][0],
          fecha: dataOC[i][1],
          proveedorNom: dataOC[i][2], 
          moneda: dataOC[i][3] || "MXN",
          subtotal: dataOC[i][4],
          iva: dataOC[i][5],
          total: dataOC[i][6],
          solicita: dataOC[i][7],
          estado: (dataOC[i][8] || "").toUpperCase()
        };
        break;
      }
    }
    
    if (!cabecera) return { error: "No se encontró la Orden de Compra: " + buscar };

    // 2. BUSCAR DATOS DEL PROVEEDOR (Captura 3)
    // Columnas: A:Nombre, ..., E:Tel, F:Email, G:Contacto, H:Condicion, I:Credito
    var pEx = { tel: "N/A", email: "N/A", contacto: "N/A", condicion: "CONTADO", credito: "0" };
    var provNombre = cabecera.proveedorNom.trim().toUpperCase();
    
    for (var k = 1; k < dataProv.length; k++) {
      if (dataProv[k][0].trim().toUpperCase() === provNombre) {
        pEx.tel = dataProv[k][4];
        pEx.email = dataProv[k][5];
        pEx.contacto = dataProv[k][6];
        pEx.condicion = dataProv[k][7];
        pEx.credito = dataProv[k][8];
        break;
      }
    }
    cabecera.prov = pEx; 

    // 3. BUSCAR PARTIDAS EN REQUISICIONES_DETALLE (Captura 1)
    // El Folio_OC está en la columna M (índice 12)
    var partidas = [];
    for (var j = 1; j < dataDet.length; j++) {
      if (dataDet[j][12].trim().toUpperCase() === buscar) {
        partidas.push({
          desc: dataDet[j][6],   // Col G: Descripcion
          cant: dataDet[j][3],   // Col D: Cantidad
          uni: dataDet[j][4],    // Col E: Unidad
          precio: dataDet[j][7], // Col H: Precio Unitario
          total: (parseFloat(String(dataDet[j][3]).replace(/[$,]/g,''))||0) * (parseFloat(String(dataDet[j][7]).replace(/[$,]/g,''))||0)   // Sin IVA: Cant x Precio
        });
      }
    }

    return { cabecera: cabecera, partidas: partidas };

  } catch (e) {
    return { error: "Error en servidor: " + e.toString() };
  }
}

// FUNCIÓN PARA CANCELAR UNA PARTIDA INDIVIDUAL
function cancelarPartidaRequi(idDetalle) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
  var sheetCab = ss.getSheetByName("REQUISICIONES");
  
  var dataDet = sheetDet.getDataRange().getValues();
  var filaReal = -1;
  var folio = "";

  // 1. Buscar la fila y el folio
  for (var i = 1; i < dataDet.length; i++) {
    if (dataDet[i][0] == idDetalle) {
      filaReal = i + 1;
      folio = dataDet[i][1];
      break;
    }
  }

  if (filaReal != -1) {
    // 2. Cambiar estado a CANCELADO
    sheetDet.getRange(filaReal, 14).setValue("CANCELADO");

    // 3. Recopilar Metadata (igual que en guardarLote)
    var dataCab = sheetCab.getDataRange().getValues();
    var metaData = {};

    for (var i = 1; i < dataCab.length; i++) {
      if (dataCab[i][2] == folio) {
        metaData[folio] = {
          solicitante: dataCab[i][3],
          prioridad: dataCab[i][4],
          observaciones: dataCab[i][7],
          msg_id: dataCab[i][10],
          partidas: []
        };
        break;
      }
    }

    // Buscar todas las partidas del folio para el ticket actualizado
    var dataDetActualizado = sheetDet.getDataRange().getValues();
    for (var j = 1; j < dataDetActualizado.length; j++) {
      if (dataDetActualizado[j][1] == folio) {
        metaData[folio].partidas.push({
          cantidad: dataDetActualizado[j][3],
          unidad: dataDetActualizado[j][4],
          descripcion: dataDetActualizado[j][6],
          estado: dataDetActualizado[j][13]
        });
      }
    }
    return metaData;
  }
  return null;
}

// MODIFICAR LA FUNCIÓN DE NOTIFICACIÓN PARA ACEPTAR TEXTOS DINÁMICOS
function editarTelegramPrivado(base64Image, folio, chatId, msgId, tipoAccion) {
  var botUrl = "https://api.telegram.org/bot" + TOKEN + "/";
  
  var statusText = (tipoAccion === "CANCELADO") ? "❌ PARTIDA CANCELADA" : "📍 Estado: COTIZADO";
  var notifText = (tipoAccion === "CANCELADO") 
    ? "⚠️ *Atención:* Se ha cancelado una o más partidas de tu folio *" + folio + "*."
    : "📊 *¡Avance en tu Requisición!* El folio *" + folio + "* ha cambiado al estado: *COTIZADO*.";

  var media = {
    'type': 'photo',
    'media': 'attach://photo',
    'caption': "🔄 *ACTUALIZACIÓN:* " + folio + "\n" + statusText,
    'parse_mode': 'Markdown'
  };

  var payloadMedia = {
    'chat_id': chatId,
    'message_id': msgId,
    'media': JSON.stringify(media),
    'photo': Utilities.newBlob(Utilities.base64Decode(base64Image.split(",")[1]), "image/png", "update.png")
  };
  UrlFetchApp.fetch(botUrl + "editMessageMedia", { 'method': 'post', 'payload': payloadMedia });

  var payloadAviso = {
    'chat_id': chatId,
    'text': notifText + "\n\n_Revisa la imagen de arriba para ver el detalle._",
    'parse_mode': 'Markdown',
    'reply_to_message_id': msgId
  };
  UrlFetchApp.fetch(botUrl + "sendMessage", { 'method': 'post', 'payload': payloadAviso });
}

////////////////////////////////////////////////////////////////////////////////////////////
///////////////////  SURTIMIENTO DE ORDEN DE COMPRA (RecepcionOCHTML) //////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

// OBTENER DATOS PARA RECEPCIÓN
function obtenerDatosRecepcion() {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheetDet = ss.getSheetByName("REQUISICIONES_DETALLE");
  var sheetOC = ss.getSheetByName("ORDENES_COMPRA");
  
  var dataDet = sheetDet.getDataRange().getDisplayValues();
  var dataOC = sheetOC.getDataRange().getDisplayValues();
  
  // Mapeo de OC para traer Fecha y Proveedor
  var ocMap = {};
  for(var i=1; i<dataOC.length; i++){
    ocMap[dataOC[i][0]] = { fecha: dataOC[i][1], proveedor: dataOC[i][2] };
  }
  
  var lista = [];
  for (var j = 1; j < dataDet.length; j++) {
    var estado = dataDet[j][13].toUpperCase(); // Col N
    // Solo mostramos lo que ya tiene OC y no se ha terminado
    if (estado === "AUTORIZADO" || estado === "PARCIAL") {
      var folioOC = dataDet[j][12]; // Col M
      var infoOC = ocMap[folioOC] || { fecha: "S/N", proveedor: "S/N" };
      
      lista.push({
        fila: j + 1,
        id: dataDet[j][0],
        folio: dataDet[j][1], // Folio Requisición
        pa: dataDet[j][2],
        cant: Number(dataDet[j][3]),
        uni: dataDet[j][4],
        desc: dataDet[j][6],
        folioOC: folioOC,
        fechaOC: infoOC.fecha,
        proveedor: infoOC.proveedor,
        recibido: dataDet[j][14] || 0, // Col O
        estado: estado
      });
    }
  }
  return lista;
}

// GUARDAR RECEPCIÓN Y NOTIFICAR
function guardarRecepcionLote(lote) {
  var ss     = SpreadsheetApp.openById(ID_HOJA_OM);
  var ssEst  = SpreadsheetApp.openById(ID_HOJA_ESTANDARES);
  var shDet  = ss.getSheetByName("REQUISICIONES_DETALLE");
  var shCab  = ss.getSheetByName("REQUISICIONES");
  var shUsr  = ssEst.getSheetByName("USUARIOS");

  // ── Mapa nombre→chatId ──
  var chatMap = {};
  if (shUsr) {
    var dataUsr = shUsr.getDataRange().getValues();
    for (var u = 1; u < dataUsr.length; u++) {
      var nm = String(dataUsr[u][1] || "").trim().toUpperCase();
      var ch = String(dataUsr[u][6] || "").trim();
      if (nm && ch) chatMap[nm] = ch;
    }
  }

  var foliosAfectados = {};

  // ── 1. Actualizar cada partida recibida ──
  lote.forEach(function(item) {
    var cantPedida   = Number(item.cant);
    var cantRecibida = Number(item.recibido);
    var nuevoEstado  = (cantRecibida >= cantPedida) ? "TERMINADO" : "PARCIAL";

    shDet.getRange(item.fila, 15).setValue(cantRecibida); // Col O = RECIBIDO
    shDet.getRange(item.fila, 14).setValue(nuevoEstado);  // Col N = ESTADO

    if (!foliosAfectados[item.folio]) foliosAfectados[item.folio] = true;
  });

  // ── 2. Leer metadata completa para ticket y notificación ──
  var resMeta    = {};
  var dataCab    = shCab.getDataRange().getValues();
  var dataDetAll = shDet.getDataRange().getValues();

  // Índice de fila por folio en cabecera (para actualizar Col J si termina todo)
  var cabFilaMap = {};
  for (var c = 1; c < dataCab.length; c++) {
    var fol = String(dataCab[c][2] || "").trim();
    if (fol) cabFilaMap[fol] = c + 1; // fila GAS (1-based)
  }

  for (var fR in foliosAfectados) {
    // Buscar datos de cabecera
    for (var i = 1; i < dataCab.length; i++) {
      if (String(dataCab[i][2]).trim() === String(fR).trim()) {
        var sol = String(dataCab[i][3] || "").trim().toUpperCase();
        resMeta[fR] = {
          solicitante:    sol,
          prioridad:      String(dataCab[i][4] || "Media"),
          observaciones:  String(dataCab[i][7] || ""),
          msg_id:         String(dataCab[i][10] || ""),
          chatId:         chatMap[sol] || "",
          partidas:       [],
          todasTerminadas: true  // se falsifica si alguna no es TERMINADO
        };
        break;
      }
    }

    if (!resMeta[fR]) continue;

    // Recopilar partidas actuales (ya actualizadas en hoja)
    // Releer para tener el estado más reciente
    var dataDetFresh = shDet.getDataRange().getValues();
    for (var j = 1; j < dataDetFresh.length; j++) {
      if (String(dataDetFresh[j][1]).trim() !== String(fR).trim()) continue;
      var estPartida = String(dataDetFresh[j][13] || "").trim().toUpperCase();
      resMeta[fR].partidas.push({
        cantidad:    dataDetFresh[j][3],
        unidad:      dataDetFresh[j][4],
        descripcion: dataDetFresh[j][6],
        estado:      estPartida,
        recibido:    dataDetFresh[j][14] || 0  // Col O = cantidad recibida
      });
      // Si alguna partida no es TERMINADO (excluyendo CANCELADO), no terminó todo
      if (estPartida !== "TERMINADO" && estPartida !== "CANCELADO") {
        resMeta[fR].todasTerminadas = false;
      }
    }

    // ── 3. Si TODAS las partidas (no canceladas) son TERMINADO → actualizar Col J cabecera ──
    if (resMeta[fR].todasTerminadas && cabFilaMap[fR]) {
      shCab.getRange(cabFilaMap[fR], 10).setValue("TERMINADO"); // Col J = ESTADO
      Logger.log("guardarRecepcionLote: REQUISICION " + fR + " marcada TERMINADO en Col J");
    }
  }

  return resMeta;
}

// MODIFICACIÓN: Envía al privado (edit) y al grupo (nuevo mensaje en tópico)
function notificarRecepcionTelegram(base64Image, folio, chatId, msgId, hayTerminados) {
  var botUrl = "https://api.telegram.org/bot" + TOKEN + "/";
  var textoEstado = hayTerminados ? "✅ TERMINADO" : "🟡 PARCIAL";
  
  // 1. EDITAR PRIVADO DEL USUARIO (Para que su historial esté actualizado)
  var media = {
    'type': 'photo', 'media': 'attach://photo',
    'caption': "📦 *ACTUALIZACIÓN DE ENTREGA:* " + folio + "\n📍 *Estado:* " + textoEstado,
    'parse_mode': 'Markdown'
  };
  UrlFetchApp.fetch(botUrl + "editMessageMedia", {
    'method': 'post',
    'payload': {
      'chat_id': chatId, 'message_id': msgId, 'media': JSON.stringify(media),
      'photo': Utilities.newBlob(Utilities.base64Decode(base64Image.split(",")[1]), "image/png", "recepcion.png")
    }
  });

  // 2. ENVIAR NUEVO MENSAJE AL TÓPICO DE RECEPCIÓN (ID: 27833)
  var payloadGrupo = {
    'chat_id': CHAT_ID_COMPRAS, // El ID del grupo principal
    'message_thread_id': 27833, // ID del subgrupo RECEPCIÓN
    'photo': Utilities.newBlob(Utilities.base64Decode(base64Image.split(",")[1]), "image/png", "recepcion_grupo.png"),
    'caption': "📥 *NUEVO INGRESO A ALMACÉN*\nFolio Requisición: *" + folio + "*\nEstado de entrega: " + textoEstado,
    'parse_mode': 'Markdown'
  };
  UrlFetchApp.fetch(botUrl + "sendPhoto", { 'method': 'post', 'payload': payloadGrupo });

  // 3. AVISO DE TEXTO AL USUARIO (Notificación Push)
  UrlFetchApp.fetch(botUrl + "sendMessage", {
    'method': 'post',
    'payload': {
      'chat_id': chatId, 'reply_to_message_id': msgId, 'parse_mode': 'Markdown',
      'text': "📥 *¡Artículos recibidos!* Se ha registrado movimiento en el folio *" + folio + "*."
    }
  });
}

////////////////////////////////////////////////////////////////////////////////////////////
/////////////////////      MODULO CIRCULANTE (CircuñanteHTML)     //////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

function getDatosIniciales() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var shCat = ss.getSheetByName("CATALOGO_CIRC");
    var shCirc = ss.getSheetByName("CIRCULANTE");
    if (!shCat || !shCirc) throw new Error("No se encontraron las pestañas.");
    var catalogo = shCat.getDataRange().getDisplayValues();
    var circulanteRaw = shCirc.getDataRange().getDisplayValues();
    return { catalogo: catalogo, headers: circulanteRaw[0] || [] };
  } catch (e) { return { error: e.toString() }; }
}

function getRegistrosCirculante() {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var data = ss.getSheetByName("CIRCULANTE").getDataRange().getValues();
  var zonaHoraria = ss.getSpreadsheetTimeZone(); // Detecta la zona de la hoja (Mexico)

  return data.map(row => row.map(cell => {
    if (cell instanceof Date) {
      // Forzamos el formato DD/MM/YYYY usando la zona horaria real de la hoja
      return Utilities.formatDate(cell, zonaHoraria, "dd/MM/yyyy");
    }
    return cell;
  }));
}

function guardarRegistrosCirculante(registros) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sheet = ss.getSheetByName("CIRCULANTE");
  var ahora = new Date();
  
  var lastRow = sheet.getLastRow();
  var nextId = 1;
  if (lastRow > 1) {
    var valorId = sheet.getRange(lastRow, 1).getValue();
    nextId = isNaN(parseInt(valorId)) ? 1 : parseInt(valorId) + 1;
  }

  var filas = registros.map(function(r) {
    // 1. Convertir fecha DD/MM/YYYY a Objeto Date real
    var partes = r.fecha.split("/");
    var fechaObjeto = new Date(partes[2], partes[1] - 1, partes[0]);

    // 2. EXTRAER MES Y AÑO DE LA FECHA DEL REGISTRO (No del sistema)
    var mesFila = fechaObjeto.getMonth() + 1;
    var anioFila = fechaObjeto.getFullYear();

    var k = parseFloat(r.kilos || 0);
    var kr = parseFloat(r.kilos_rec || 0);
    var dif = Math.round(kr - k);

    return [
      nextId++, 
      fechaObjeto, // Columna B: Fecha real
      r.envio, 
      r.codigo, 
      r.descripcion, 
      r.diametro, 
      r.largo, 
      r.acero,
      r.pedido, 
      r.origen, 
      r.cantidad, 
      r.unidad, 
      k, 
      kr, 
      mesFila,   // MES basado en la fecha ingresada
      anioFila,  // AÑO basado en la fecha ingresada
      dif, 
      r.comentarios, 
      r.movimiento, 
      "", 
      ahora      // Fecha de auditoría (esta sí es la actual)
    ];
  });

  var startRow = lastRow + 1;
  sheet.getRange(startRow, 1, filas.length, 21).setValues(filas);
  
  // Formatear columna de fecha para que Sheets no la cambie
  sheet.getRange(startRow, 2, filas.length, 1).setNumberFormat("dd/mm/yyyy");
  // Formatear columna de diferencia como entero
  sheet.getRange(startRow, 17, filas.length, 1).setNumberFormat("0");

  return "OK";
}

function crudCatalogo(item) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sh = ss.getSheetByName("CATALOGO_CIRC");
  var data = sh.getDataRange().getValues();
  
  if (item.id === "NUEVO") {
    // Generar ID automático (último ID + 1)
    var maxId = 0;
    for (var i = 1; i < data.length; i++) {
      var currentId = parseInt(data[i][0]);
      if (!isNaN(currentId) && currentId > maxId) maxId = currentId;
    }
    var nextId = maxId + 1;
    sh.appendRow([nextId, item.codigo, item.descripcion, item.diametro, item.largo, item.acero, item.unidad]);
  } else {
    // Editar registro existente
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() == item.id.toString()) {
        sh.getRange(i + 1, 2, 1, 6).setValues([[
          item.codigo, item.descripcion, item.diametro, item.largo, item.acero, item.unidad
        ]]);
        break;
      }
    }
  }
  return "OK";
}

////////////////////////////////////////////////////////////////////////////////////////////
/////////////      MODULO CONTROL DE REFACCIONES ALMACEN (CircuñanteHTML)     //////////////
////////////////////////////////////////////////////////////////////////////////////////////

function getDatosInsumos() {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var shCat = ss.getSheetByName("CAT_INSUMOS");
    var shMov = ss.getSheetByName("MOV_INSUMOS");
    
    var catalogo = shCat.getDataRange().getDisplayValues();
    
    // Obtener los últimos 10 movimientos
    var lastRow = shMov.getLastRow();
    var ultimosMov = [];
    if (lastRow > 1) {
      var numRows = Math.min(lastRow - 1, 10);
      ultimosMov = shMov.getRange(lastRow - numRows + 1, 1, numRows, 10).getDisplayValues().reverse();
    }

    return {
      catalogo: catalogo,
      ultimosMov: ultimosMov
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

function crudCatalogoInsumos(item) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var sh = ss.getSheetByName("CAT_INSUMOS");
  var data = sh.getDataRange().getValues();
  const clean = (val) => val ? val.toString().toUpperCase().trim() : "";
  
  var urlArchivo = item.url_existente || "";

  // GESTIÓN DE DRIVE
  if (item.archivo && item.archivo.base64) {
    var folder = DriveApp.getFolderById(ID_CARPETA_INSUMOS);
    
    // Borrar anterior si existe
    if (urlArchivo !== "") {
      try {
        var oldId = urlArchivo.match(/[-\w]{25,}/);
        if (oldId) DriveApp.getFileById(oldId[0]).setTrashed(true);
      } catch(e) {}
    }
    
    // Subir nuevo
    var blob = Utilities.newBlob(Utilities.base64Decode(item.archivo.base64), item.archivo.type, item.archivo.name);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urlArchivo = file.getUrl();
  }

  if (item.id === "NUEVO") {
    // ... tu logica de maxId ...
    sh.appendRow([
      "INS-" + (maxId + 1), clean(item.codigo), clean(item.categoria), clean(item.descripcion), 
      clean(item.unidad), clean(item.ubicacion), 0, item.minimo, item.maximo, 
      item.reorden, clean(item.especificaciones), "OK", 
      clean(item.proveedor), clean(item.referencia), clean(item.tipo_insumo), urlArchivo // Columna P
    ]);
  } else {
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString() === item.id.toString()) {
        var row = i + 1;
        // Escribir datos en mayúsculas (B a F)
        sh.getRange(row, 2, 1, 5).setValues([[clean(item.codigo), clean(item.categoria), clean(item.descripcion), clean(item.unidad), clean(item.ubicacion)]]);
        // Escribir mínimos y especificaciones (H a K)
        sh.getRange(row, 8, 1, 4).setValues([[item.minimo, item.maximo, item.reorden, clean(item.especificaciones)]]);
        // Escribir extras (L a O en mayúsculas, P es la URL original)
        sh.getRange(row, 12, 1, 4).setValues([[clean(item.estado), clean(item.proveedor), clean(item.referencia), clean(item.tipo_insumo)]]);
        sh.getRange(row, 16).setValue(urlArchivo); // Columna P fija
        break;
      }
    }
  }
  return "OK";
}

function guardarMovimientosInsumos(listaMovs) {
  var ss = SpreadsheetApp.openById(ID_HOJA_OM);
  var shCat = ss.getSheetByName("CAT_INSUMOS");
  var shMov = ss.getSheetByName("MOV_INS_KARDEX"); // Cambié el nombre si es historial, o MOV_INSUMOS
  if(!shMov) shMov = ss.getSheetByName("MOV_INSUMOS");
  
  var ahora = new Date();
  var catData = shCat.getDataRange().getValues();
  const clean = (val) => val ? val.toString().toUpperCase().trim() : "";

  listaMovs.forEach(function(m) {
    // 1. Historial en MAYÚSCULAS
    shMov.appendRow([
      "REG-" + ahora.getTime() + Math.floor(Math.random()*100),
      m.fecha, clean(m.tipo), clean(m.codigo), m.cantidad, clean(m.usuario), 
      clean(m.destino), clean(m.referencia), clean(m.comentarios), ahora
    ]);

    // 2. Actualizar Stock
    for (var i = 1; i < catData.length; i++) {
      if (catData[i][1] == m.codigo) {
        var currentStock = parseFloat(catData[i][6] || 0);
        var diff = parseFloat(m.cantidad);
        var newStock = (m.tipo === "ENTRADA") ? currentStock + diff : currentStock - diff;
        shCat.getRange(i + 1, 7).setValue(newStock);
        
        // Semáforo
        var min = parseFloat(catData[i][7] || 0);
        var reo = parseFloat(catData[i][9] || 0);
        var status = "OK";
        if (newStock <= min) status = "CRITICO";
        else if (newStock <= reo) status = "REORDEN";
        shCat.getRange(i + 1, 12).setValue(status);
        break;
      }
    }
  });
  return "OK";
}

function getHistorialEspecificoInsumo(codigo) {
  try {
    var ss = SpreadsheetApp.openById(ID_HOJA_OM);
    var shMov = ss.getSheetByName("MOV_INSUMOS");
    var data = shMov.getDataRange().getDisplayValues();
    
    // Filtrar los registros donde la columna D (índice 3) sea igual al código
    // Luego invertimos el orden para que el más reciente aparezca arriba
    var filtrados = data.filter(function(row) {
      return row[3] === codigo;
    }).reverse();

    // Devolvemos solo los primeros 10 del historial de ese producto
    return filtrados.slice(0, 10);
  } catch (e) {
    return [];
  }
}

////////////////////////////////////////////////////////////////////////////////////////////
/////////////////  VER LA EXISTENCIA DE MATERIA PRIMA (ExistenciaHTML) /////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

function getMateriaPrimaData() {
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA_OM);
    const sheet = ss.getSheetByName("ENTRADAS_MP");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    // Mapeo de columnas según tu estructura
    const col = {
      ID: 0, PROVEEDOR: 1, FECHA: 2, DIAMETRO: 3, ACERO: 4, N_ROLLO: 5, KILOS: 6,
      COLADA: 7, SELLO: 8, N_REMISION: 9, CONSECUTIVO: 10, ESTADO: 11,
      FECHA_INSP: 18, C_DEC_PARC: 19, C_DEC_LIBRE: 20, P_DEC_PARCIAL: 21,
      P_DEC_LIBRE: 22, RESIST_C: 23, RESIST_P: 24
    };

    const cleanData = rows.filter(r => r[col.ESTADO] === "ACTIVO").map(r => {
      return {
        id: r[col.ID],
        proveedor: r[col.PROVEEDOR],
        fecha: r[col.FECHA] instanceof Date ? Utilities.formatDate(r[col.FECHA], "GMT-6", "dd/MM/yyyy") : r[col.FECHA],
        diametro: parseFloat(r[col.DIAMETRO]) || 0,
        aceroFull: String(r[col.ACERO]),
        aceroLimpio: String(r[col.ACERO]).replace(/\D/g, ""),
        n_rollo: r[col.N_ROLLO],
        kilos: parseFloat(r[col.KILOS]) || 0,
        colada: r[col.COLADA],
        sello: r[col.SELLO],
        n_remision: r[col.N_REMISION],
        consecutivo: r[col.CONSECUTIVO],
        // Inspección
        f_insp: r[col.FECHA_INSP] instanceof Date ? Utilities.formatDate(r[col.FECHA_INSP], "GMT-6", "dd/MM/yyyy") : r[col.FECHA_INSP],
        c_parc: r[col.C_DEC_PARC], c_total: r[col.C_DEC_LIBRE],
        p_parc: r[col.P_DEC_PARCIAL], p_total: r[col.P_DEC_LIBRE],
        res_c: r[col.RESIST_C], res_p: r[col.RESIST_P],
        tieneInsp: (r[col.FECHA_INSP] || r[col.RESIST_C]) ? true : false
      };
    });

    return cleanData;
  } catch (e) { return { error: e.toString() }; }
}

////////////////////////////////////////////////////////////////////////////////////////////
///////////////  ENTRADAS Y SALIDAS DE MATERIA PRIMA (GestionMP_HTML) /////////////////////
////////////////////////////////////////////////////////////////////////////////////////////

function verificarSellosDuplicados(sellosNuevos) {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const sheet = ss.getSheetByName("ENTRADAS_MP");
  // Obtenemos todos los sellos de la columna I (desde fila 2 hasta el final)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; 
  
  const data = sheet.getRange(2, 9, lastRow - 1, 1).getValues().flat();
  const baseDatosSellos = data.map(s => String(s).trim().toUpperCase());
  
  // Filtramos: ¿Qué sellos del usuario ya existen en la base de datos?
  const duplicadosEncontrados = sellosNuevos.filter(sello => {
    const sLimpio = String(sello).trim().toUpperCase();
    return sLimpio !== "" && baseDatosSellos.includes(sLimpio);
  });

  return duplicadosEncontrados;
}

// Guardar nuevas entradas desde el Excel procesado
function guardarEntradasMP(registros) {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const sheet = ss.getSheetByName("ENTRADAS_MP");
  const lastRow = sheet.getLastRow();
  let nextId = lastRow > 0 ? parseInt(sheet.getRange(lastRow, 1).getValue()) + 1 : 1;
  if(isNaN(nextId)) nextId = 1;

  // Modifica esta parte dentro de guardarEntradasMP
const rowsToAppend = registros.map((r, index) => {
    // Convertimos la fecha de YYYY-MM-DD a DD/MM/YYYY para que Sheets lo reconozca bien
    const partes = r.fecha.split("-");
    const fechaFormateada = `${partes[2]}/${partes[1]}/${partes[0]}`;

    return [
      nextId + index, r.proveedor, fechaFormateada, r.diametro, r.acero, r.n_rollo, r.kilos, 
      r.colada, r.sello, r.n_remision, r.consecutivo, "ACTIVO"
    ];
});

  sheet.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  return true;
}

// Procesar Salidas
function procesarSalidasMP(cambios) {
  const ss = SpreadsheetApp.openById(ID_HOJA_OM);
  const sheet = ss.getSheetByName("ENTRADAS_MP");
  const data = sheet.getDataRange().getValues();
  
  cambios.forEach(cambio => {
    const rowIndex = data.findIndex(row => row[0] == cambio.id); // Buscar por ID (Col A)
    if (rowIndex !== -1) {
      const sheetRow = rowIndex + 1;
      // Actualizar: ESTADO=BAJA (Col L/11), FECHA_SALIDA (Col M/12), KILOS_SALIDA=KILOS (Col O/14), COMENTARIOS_M (Col Q/16)
      sheet.getRange(sheetRow, 12).setValue("BAJA");
      sheet.getRange(sheetRow, 13).setValue(cambio.fechaSalida);
      sheet.getRange(sheetRow, 15).setValue(data[rowIndex][6]); // Kilos originales a Kilos Salida
      sheet.getRange(sheetRow, 17).setValue(cambio.comentarios.toUpperCase());
    }
  });
  return true;
}

function getDatosReporteMensual(mes, anio) {
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA_OM);
    const sheet = ss.getSheetByName("ENTRADAS_MP");
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, 15).getValues(); // Leemos hasta Col O (15)

    return data.map(r => {
      // Función para normalizar fechas a string DD/MM/YYYY
      const formatFecha = (f) => {
        if (!f || !(f instanceof Date)) return String(f || "");
        return Utilities.formatDate(f, "GMT-6", "dd/MM/yyyy");
      };

      return {
        proveedor: String(r[1]),
        fechaEntrada: formatFecha(r[2]),  // Col C
        diametro: String(r[3]),
        acero: String(r[4]),
        kilos: parseFloat(r[6]) || 0,     // Col G
        sello: String(r[8]),
        estado: String(r[11]),            // Col L
        fechaSalida: formatFecha(r[12]),  // Col M
        kilosSalida: parseFloat(r[14]) || 0 // Col O
      };
    });
  } catch (e) {
    console.log("Error en reporte: " + e.toString());
    return [];
  }
}

// Obtener datos para la pestaña de Consultas
function getConsultasMP(tipo, fInicio, fFin) {
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA_OM);
    const sheet = ss.getSheetByName("ENTRADAS_MP");
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    const rows = data.slice(1);

    // Convertimos fechas de los inputs (YYYY-MM-DD) a milisegundos a medianoche
    const pIni = fInicio.split("-");
    const pFin = fFin.split("-");
    const tInicio = new Date(pIni[0], pIni[1]-1, pIni[2]).getTime();
    const tFin    = new Date(pFin[0], pFin[1]-1, pFin[2]).getTime();

    // Col C es indice 2 (Entradas), Col M es indice 12 (Salidas)
    const colBusqueda = (tipo === 'entradas') ? 2 : 12;

    const filtered = rows.filter(r => {
      let valor = r[colBusqueda];
      if (!valor || valor === "") return false;
      
      let fechaRow;
      if (valor instanceof Date) {
        // Normalizar fecha de la celda a medianoche
        fechaRow = new Date(valor.getFullYear(), valor.getMonth(), valor.getDate());
      } else {
        // Parsear DD/MM/YYYY
        const p = String(valor).split("/");
        if (p.length !== 3) return false;
        fechaRow = new Date(p[2], p[1]-1, p[0]);
      }
      
      const tRow = fechaRow.getTime();
      return tRow >= tInicio && tRow <= tFin;
    });

    return filtered.map(r => ({
      id: r[0], 
      proveedor: r[1], 
      fecha: r[2] instanceof Date ? Utilities.formatDate(r[2], "GMT-6", "dd/MM/yyyy") : r[2], 
      diametro: r[3], 
      acero: r[4],
      n_rollo: r[5], 
      kilos: r[6], 
      colada: r[7], 
      sello: r[8], 
      n_remision: r[9],
      consecutivo: r[10], 
      estado: r[11], 
      fecha_salida: r[12] instanceof Date ? Utilities.formatDate(r[12], "GMT-6", "dd/MM/yyyy") : r[12], 
      kilos_salida: r[14],
      comentarios: r[16]
    }));
  } catch (e) {
    return [];
  }
}

// Actualizar estado y campos relacionados
function actualizarEstadoMP(id, nuevoEstado, infoSalida, nombreUsuario) {
  const ss    = SpreadsheetApp.openById(ID_HOJA_OM);
  const sheet = ss.getSheetByName("ENTRADAS_MP");
  const data  = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[0] == id);

  if (rowIndex === -1) return false;
  const sheetRow = rowIndex + 1;

  if (nuevoEstado === "ACTIVO") {
    // Limpiar campos de salida
    sheet.getRange(sheetRow, 12).setValue("ACTIVO"); // Col L
    sheet.getRange(sheetRow, 13).clearContent();     // Col M (Fecha Salida)
    sheet.getRange(sheetRow, 15).clearContent();     // Col O (Kilos Salida)
    sheet.getRange(sheetRow, 17).clearContent();     // Col Q (Comentarios)
  } else {
    // Pasar a BAJA
    sheet.getRange(sheetRow, 12).setValue("BAJA");
    sheet.getRange(sheetRow, 13).setValue(infoSalida.fecha);
    sheet.getRange(sheetRow, 15).setValue(infoSalida.kilos);
    sheet.getRange(sheetRow, 17).setValue(infoSalida.comentarios.toUpperCase());
  }

  // ── Guardar historial de cambio de estado en Col R ────────────────
  var usuario    = String(nombreUsuario || "DESCONOCIDO").trim().toUpperCase();
  var zona       = "GMT-6";
  var ahora      = Utilities.formatDate(new Date(), zona, "dd/MM/yyyy HH:mm:ss");

  var descripcion = nuevoEstado === "ACTIVO"
    ? "Cambió ESTADO de BAJA a ACTIVO"
    : "Cambió ESTADO de ACTIVO a BAJA (Fecha salida: " + infoSalida.fecha + ")";

  var textoNuevo = ahora + "_" + usuario + "_" + descripcion;

  // Leer historial anterior en Col R (índice 17 = columna 18)
  var historialAnterior = String(data[rowIndex][17] || "").trim();

  var historialFinal = historialAnterior
    ? historialAnterior + "\n" + textoNuevo
    : textoNuevo;

  sheet.getRange(sheetRow, 18).setValue(historialFinal); // Col R

  Logger.log("actualizarEstadoMP OK — ID: " + id + " | " + textoNuevo);
  return true;
}

/////////////////////////////////////////////////////////////////////////////////////////////////////
//////////// esto no me acuerdo para que era creo apra enlazar el DRIVE /////////////////////////////
/////////////////////////////////////////////////////////////////////////////////////////////////////
function test2(){ DriveApp.getRootFolder(); }