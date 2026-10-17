function doGet(e) {
  var sheetName = e.parameter.sheet || 'Itinerario';
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(sheetName);

  if (!sheet) {
    return replyJSON([]);
  }

  var values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return replyJSON([]);
  }

  // [CAMBIO 2] Zona horaria del propio Sheet (Argentina, Japón, etc.)
  var tz = doc.getSpreadsheetTimeZone();

  var headers = values[0];
  var data = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var obj = {};
    var hasContent = false;

    for (var j = 0; j < headers.length; j++) {
      var headerKey = String(headers[j]).trim();
      if (headerKey !== '') {
        var val = row[j];
        // [CAMBIO 2] Si la celda es una fecha, la mandamos como texto dd/mm/aaaa
        if (esFecha(val)) {
          val = Utilities.formatDate(val, tz, 'dd/MM/yyyy');
        }
        obj[headerKey] = val;
        if (val !== '' && val !== null && val !== undefined) hasContent = true;
      }
    }
    if (hasContent) {
      data.push(obj);
    }
  }

  return replyJSON(data);
}

function doPost(e) {
  try {
    var item = JSON.parse(e.postData.contents);
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = item.sheet || 'Itinerario';
    var sheet = doc.getSheetByName(sheetName);

    if (!sheet) {
      return ContentService.createTextOutput("Error: No se encontró la pestaña " + sheetName);
    }

    // RESOLUCIÓN DE COORDENADAS AUTOMÁTICA
    if (sheetName === 'Itinerario') {
      if (item.UbicacionInicio || item.location) {
        item.CoordsInicio = resolveGoogleMapsCoords(item.UbicacionInicio || item.location);
      }
      if (item.UbicacionFin || item.destLink) {
        item.CoordsFin = resolveGoogleMapsCoords(item.UbicacionFin || item.destLink);
      }
    } else if (sheetName === 'Alojamiento') {
      if (item.Direccion || item.location) {
        item.Coords = resolveGoogleMapsCoords(item.Direccion || item.location);
      }
    } else if (sheetName === 'Recomendaciones') {
      if (item.LinkGoogleMaps || item.Ubicacion) {
        item.Coords = resolveGoogleMapsCoords(item.LinkGoogleMaps || item.Ubicacion);
      }
    }

    var data = sheet.getDataRange().getValues();
    var headers = data.length > 0 ? data[0] : [];

    // ACCIÓN: ELIMINAR
    if (item.action === 'delete') {
      var targetTitle = item.title || item.Titulo || item.Concepto || item.Nombre || '';
      var targetDate = item.date || item.FechaInicio || item.Fecha || item.start || '';

      for (var i = 1; i < data.length; i++) {
        var rowMatch = false;
        for (var c = 0; c < headers.length; c++) {
          var cellVal = String(data[i][c]).trim();
          if (targetTitle !== '' && cellVal === String(targetTitle).trim()) {
            rowMatch = true;
            break;
          }
        }
        if (rowMatch) {
          sheet.deleteRow(i + 1);
          return ContentService.createTextOutput("Eliminado correctamente.");
        }
      }
      return ContentService.createTextOutput("Registro no encontrado para eliminar.");
    }

    // ACCIÓN: EDITAR / GUARDAR
    var rowToUpdate = -1;
    var oldTitle = item.oldTitle || item.title || item.Titulo || item.Concepto || item.Nombre || '';

    if (oldTitle !== '' && data.length > 1) {
      for (var r = 1; r < data.length; r++) {
        var matchFound = false;
        for (var col = 0; col < headers.length; col++) {
          if (String(data[r][col]).trim() === String(oldTitle).trim()) {
            matchFound = true;
            break;
          }
        }
        if (matchFound) {
          rowToUpdate = r + 1; // Fila en Google Sheets (base 1)
          break;
        }
      }
    }

    // Mapeo dinámico contra la Fila 1 de encabezados
    var newRow = [];
    for (var h = 0; h < headers.length; h++) {
      var key = String(headers[h]).trim();
      var value = item[key] !== undefined ? item[key] : (item[getFallbackKey(key)] !== undefined ? item[getFallbackKey(key)] : '');
      newRow.push(value);
    }

    if (rowToUpdate > 0) {
      // Reemplaza la fila existente sin borrarla
      var range = sheet.getRange(rowToUpdate, 1, 1, newRow.length);
      range.setValues([newRow]);
      return ContentService.createTextOutput("Actualizado correctamente.");
    } else {
      // Agrega una fila nueva al final
      sheet.appendRow(newRow);
      return ContentService.createTextOutput("Guardado con éxito.");
    }

  } catch (err) {
    return ContentService.createTextOutput("Error en backend: " + err.message);
  }
}

// Helpers útiles integrados
function replyJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// [CAMBIO 2] Devuelve true si el valor es una fecha válida (forma segura de detectarlo)
function esFecha(val) {
  return Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime());
}

function formatDateStandard(d) {
  var day = ("0" + d.getDate()).slice(-2);
  var month = ("0" + (d.getMonth() + 1)).slice(-2);
  var year = d.getFullYear();
  return day + "/" + month + "/" + year;
}

function getFallbackKey(key) {
  var map = {
    'Titulo': 'title', 'Tipo': 'type', 'FechaInicio': 'date', 'HoraInicio': 'startHour',
    'UbicacionInicio': 'location', 'CoordsInicio': 'CoordsInicio',
    'FechaFin': 'endDate', 'HoraFin': 'endHour', 'UbicacionFin': 'destLink', 'CoordsFin': 'CoordsFin',
    'Costo': 'cost', 'Detalles': 'details', 'Duracion': 'durationHours',
    'CodigoReserva': 'bookingCode', 'Nombre': 'title', 'Ciudad': 'location',
    'CheckIn': 'startHour', 'CheckOut': 'endHour', 'LinkReserva': 'bookingCode',
    'Direccion': 'location', 'Coords': 'Coords',
    'Concepto': 'title', 'Monto_ARS': 'cost', 'Monto_Dlr': 'cost'
  };
  return map[key] || key;
}

// Expande links cortos de Google Maps leyendo el encabezado Location (sin seguir redirecciones solas)
// [CAMBIO 3] Si el link no trae coordenadas (links compartidos desde la app del celular),
// usa el nombre/dirección del lugar y lo consulta en el geocodificador de Apps Script.
function resolveGoogleMapsCoords(urlOrText) {
  if (!urlOrText || typeof urlOrText !== 'string') return "";

  // 1. Si ya viene como "lat,lng"
  var directMatch = urlOrText.match(/^(-?\d+\.\d+),\s*(-?\d+\.\d+)$/);
  if (directMatch) return directMatch[1] + "," + directMatch[2];

  // 2. Primer link del texto
  var urlMatch = urlOrText.match(/https?:\/\/[^\s]+/);
  if (!urlMatch) return "";

  var currentUrl = urlMatch[0];

  // 3. Si el link ya es largo y trae coordenadas, no hace falta pedir nada
  var coordsDirectas = extraerCoordsDeTexto(currentUrl);
  if (coordsDirectas) return coordsDirectas;

  var nombreLugar = extraerNombreLugar(currentUrl);

  var options = {
    followRedirects: false,
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
  };

  try {
    // 4. Hasta 5 saltos de redirección, leyendo Location en cada uno
    for (var hop = 0; hop < 5; hop++) {
      var resp = UrlFetchApp.fetch(currentUrl, options);
      var headers = resp.getHeaders();
      var location = headers['Location'] || headers['location'] || '';

      var coords = extraerCoordsDeTexto(location);
      if (coords) return coords;

      // Guardamos el nombre del lugar si la redirección lo trae (plan B)
      var nombre = extraerNombreLugar(location);
      if (nombre) nombreLugar = nombre;

      if (!location) {
        // Sin redirección: buscamos en el contenido de la página
        var coordsPagina = extraerCoordsDeTexto(resp.getContentText());
        if (coordsPagina) return coordsPagina;
        break;
      }
      currentUrl = location;
    }
  } catch (e) {
    Logger.log("Error resolviendo coordenadas: " + e.toString());
  }

  // 5. Plan B: geocodificar por nombre/dirección del lugar
  if (nombreLugar) {
    return geocodificarNombre(nombreLugar);
  }

  return "";
}

// Busca coordenadas dentro de un texto o URL de Google Maps.
// [CAMBIO 5] Prioriza !3d!4d (punto exacto del lugar) sobre @ (centro de la pantalla).
function extraerCoordsDeTexto(texto) {
  if (!texto) return "";
  var t = String(texto);
  var m = t.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) ||
          t.match(/query=(-?\d+\.\d+)(?:%2C|,)(-?\d+\.\d+)/) ||
          t.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/);
  return m ? m[1] + "," + m[2] : "";
}

// [CAMBIO 3] Saca el nombre/dirección del lugar de una URL de Google Maps.
// Ej: .../maps/place/Abasto+Shopping,+Av.+Corrientes+3247,.../data=...  ->  "Abasto Shopping, Av. Corrientes 3247, ..."
// También entiende links del tipo ...?q=Nombre+del+lugar
function extraerNombreLugar(url) {
  if (!url) return "";
  var u = String(url);
  var crudo = "";

  var mPlace = u.match(/\/maps\/place\/([^\/?]+)/);
  if (mPlace) {
    crudo = mPlace[1];
  } else {
    var mQ = u.match(/[?&]q=([^&]+)/);
    if (mQ) crudo = mQ[1];
  }
  if (!crudo) return "";

  var limpio = crudo.replace(/\+/g, ' ');
  try {
    limpio = decodeURIComponent(limpio);
  } catch (e) {
    // si no se puede decodificar, usamos el texto tal cual
  }
  return limpio.trim();
}

// [CAMBIO 3] Pide coordenadas al geocodificador incorporado de Apps Script (gratis, sin clave)
function geocodificarNombre(nombre) {
  try {
    var resp = Maps.newGeocoder().geocode(nombre);
    if (resp && resp.status === 'OK' && resp.results && resp.results.length > 0) {
      var loc = resp.results[0].geometry.location;
      return loc.lat + "," + loc.lng;
    }
    Logger.log('Geocoder sin resultados para: "' + nombre + '" (estado: ' + (resp ? resp.status : 'sin respuesta') + ')');
  } catch (e) {
    Logger.log('Error en geocoder: ' + e.toString());
  }
  return "";
}

// Prueba de diagnóstico: NO escribe en el Sheet
function probarCoords() {
  var link = 'https://maps.app.goo.gl/Pf9BvmsjKartN9SR7';
  Logger.log('Coordenadas detectadas: ' + resolveGoogleMapsCoords(link));
}

// Diagnóstico: muestra cada salto de redirección de un link. NO escribe en el Sheet.
function diagnosticarLink() {
  var link = 'https://maps.app.goo.gl/3npcg51vywo9iYdY9?g_st=ac';
  var options = {
    followRedirects: false,
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
  };
  var url = link;
  for (var hop = 0; hop < 6; hop++) {
    var resp = UrlFetchApp.fetch(url, options);
    var headers = resp.getHeaders();
    var location = headers['Location'] || headers['location'] || '';
    Logger.log('Salto ' + hop + ' | código ' + resp.getResponseCode() + ' | Location: ' + location);
    if (!location) {
      var texto = resp.getContentText();
      Logger.log('Sin más redirecciones. Coords encontradas en la página: "' + extraerCoordsDeTexto(texto) + '"');
      Logger.log('Primeros 300 caracteres de la página: ' + texto.substring(0, 300));
      break;
    }
    url = location;
  }
  Logger.log('Resultado de resolveGoogleMapsCoords: "' + resolveGoogleMapsCoords(link) + '"');
}
// Diagnóstico: muestra cada salto de redirección de un link y busca coordenadas en varios formatos.
// NO escribe en el Sheet.
function diagnosticarLink() {
  var link = 'https://maps.app.goo.gl/phY1r4YLacR9Jeiq5';
  var options = {
    followRedirects: false,
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
  };
  var url = link;
  var nombreLugar = extraerNombreLugar(link);

  buscarPatron(link, 'En el link original @', /@(-?\d+\.\d+),(-?\d+\.\d+)/);
  buscarPatron(link, 'En el link original !3d!4d', /!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);

  for (var hop = 0; hop < 6; hop++) {
    var resp = UrlFetchApp.fetch(url, options);
    var headers = resp.getHeaders();
    var location = headers['Location'] || headers['location'] || '';
    Logger.log('Salto ' + hop + ' | código ' + resp.getResponseCode() + ' | Location: ' + location);

    if (location) {
      buscarPatron(location, 'Salto ' + hop + ' @', /@(-?\d+\.\d+),(-?\d+\.\d+)/);
      buscarPatron(location, 'Salto ' + hop + ' !3d!4d', /!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
      var nombre = extraerNombreLugar(location);
      if (nombre) nombreLugar = nombre;
    } else {
      var texto = resp.getContentText();
      Logger.log('--- Buscando dentro de la página ---');
      buscarPatron(texto, 'Página center=', /center=(-?\d+\.\d+)(?:%2C|,)(-?\d+\.\d+)/);
      buscarPatron(texto, 'Página markers=', /markers=(-?\d+\.\d+)(?:%2C|,)(-?\d+\.\d+)/);
      buscarPatron(texto, 'Página !3d!4d', /!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/);
      buscarPatron(texto, 'Página [null,null,lat,lng]', /\[null,null,(-?\d+\.\d+),(-?\d+\.\d+)\]/);
      buscarPatron(texto, 'Página @lat,lng', /@(-?\d+\.\d+),(-?\d+\.\d+)/);
      break;
    }
    url = location;
  }

  Logger.log('Nombre del lugar detectado: "' + nombreLugar + '"');
  if (nombreLugar) {
    Logger.log('Resultado del geocodificador (plan B): "' + geocodificarNombre(nombreLugar) + '"');
  }
  Logger.log('Resultado final de resolveGoogleMapsCoords: "' + resolveGoogleMapsCoords(link) + '"');
}

// Ayudante del diagnóstico: busca un patrón y muestra lo que encontró
function buscarPatron(texto, etiqueta, regex) {
  var m = String(texto).match(regex);
  Logger.log(etiqueta + ': ' + (m ? m[1] + ',' + m[2] : '(no encontrado)'));
}
