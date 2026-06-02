/**
 * ============================================================================
 *  ENVÍO MASIVO DE CORREOS (Mail Merge)  —  Google Sheets + Apps Script
 *  Versión con soporte de CC (copia) y CCO (copia oculta)
 * ============================================================================
 *
 *  CÓMO FUNCIONA
 *  -------------
 *  1. Tienes una hoja con tus contactos. La PRIMERA fila son los encabezados.
 *  2. Redactas el correo como un BORRADOR en Gmail (con formato, imágenes,
 *     adjuntos si quieres) y en el cuerpo usas marcadores como {{Nombre}}.
 *  3. El script toma cada fila, reemplaza {{Encabezado}} por el valor de esa
 *     columna y envía el correo. Marca cada fila como "ENVIADO" para no
 *     duplicar si lo vuelves a correr.
 *
 *  ESTRUCTURA DE LA HOJA "Contactos" (los nombres deben coincidir con CONFIG)
 *  --------------------------------------------------------------------------
 *  | Correo | Nombre | Dependencia | CC | CCO | Estado | FechaEnvio |
 *  |--------|--------|-------------|----|-----|--------|------------|
 *  | a@x.com| Juan   | C5i         |    |     |        |            |
 *
 *  - "Correo" es obligatorio (destinatario principal).
 *  - "CC" y "CCO" son opcionales: puedes dejarlas vacías o poner uno o varios
 *    correos separados por COMA, por ejemplo:  jefe@x.com, archivo@x.com
 *  - Puedes agregar las columnas que quieras y usarlas como {{Columna}}.
 *  - "Estado" y "FechaEnvio" las llena el script (déjalas vacías).
 *
 *  PRIMERA VEZ
 *  -----------
 *  1. Pega este código en Extensiones > Apps Script.
 *  2. Ajusta el objeto CONFIG de abajo (sobre todo ASUNTO_BORRADOR).
 *  3. Guarda, recarga la hoja y usa el menú "📧 Envío masivo".
 *  4. La primera ejecución te pedirá autorizar permisos de Gmail.
 *
 *  CUOTA (importante)
 *  ------------------
 *  - Gmail gratuito: ~100 correos/día.
 *  - Google Workspace: ~1,500 correos/día.
 *  OJO: cada destinatario cuenta, incluidos los de CC y CCO. Si copias siempre
 *  a las mismas direcciones, esos correos también consumen tu cuota diaria.
 * ============================================================================
 */

// ===================== CONFIGURACIÓN =====================
const CONFIG = {
  HOJA_CONTACTOS: 'Contactos',   // Nombre de la pestaña con los datos
  COL_CORREO:     'Correo',      // Encabezado de la columna con el email principal
  COL_CC:         'CC',          // Columna de copia (opcional)
  COL_CCO:        'CCO',         // Columna de copia oculta (opcional)
  COL_ESTADO:     'Estado',      // El script escribe aquí ENVIADO / ERROR
  COL_FECHA:      'FechaEnvio',  // El script escribe aquí la fecha de envío

  ASUNTO_BORRADOR: 'PON AQUÍ EL ASUNTO EXACTO DE TU BORRADOR',

  NOMBRE_REMITENTE: '',          // Opcional: nombre visible. Ej. 'SESESP Informática'
  RESPONDER_A:      '',          // Opcional: email de respuesta (reply-to)
  CORREO_PRUEBA:    '',          // Tu correo para "Enviar prueba". Si vacío, usa el tuyo.

  PAUSA_MS:         1200,        // Pausa entre correos (evita ráfagas)
  MARCADOR_INICIO:  '{{',        // Delimitadores de los campos de combinación
  MARCADOR_FIN:     '}}'
};
// =========================================================


/** Crea el menú al abrir la hoja. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 Envío masivo')
    .addItem('Enviar correos pendientes', 'enviarCorreos')
    .addItem('Enviar correo de prueba',   'enviarPrueba')
    .addSeparator()
    .addItem('Reiniciar estados (volver a enviar a todos)', 'reiniciarEstados')
    .addItem('Ver cuota disponible',       'verCuota')
    .addToUi();
}


/** Acción principal: recorre las filas y envía a los pendientes. */
function enviarCorreos() {
  const ui = SpreadsheetApp.getUi();

  if (!CONFIG.ASUNTO_BORRADOR || CONFIG.ASUNTO_BORRADOR.indexOf('PON AQUÍ') === 0) {
    ui.alert('Falta configurar ASUNTO_BORRADOR en el código (debe ser el asunto exacto de tu borrador de Gmail).');
    return;
  }

  const plantilla = obtenerPlantillaDesdeBorrador_(CONFIG.ASUNTO_BORRADOR);
  if (!plantilla) {
    ui.alert('No encontré un borrador en Gmail con el asunto:\n\n"' + CONFIG.ASUNTO_BORRADOR + '"\n\nRevisa que el asunto coincida exactamente.');
    return;
  }

  const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_CONTACTOS);
  if (!hoja) {
    ui.alert('No existe la pestaña "' + CONFIG.HOJA_CONTACTOS + '".');
    return;
  }

  const datos        = hoja.getDataRange().getDisplayValues();
  const encabezados  = datos[0];
  const filas        = datos.slice(1);

  const iCorreo = encabezados.indexOf(CONFIG.COL_CORREO);
  const iCC     = encabezados.indexOf(CONFIG.COL_CC);
  const iCCO    = encabezados.indexOf(CONFIG.COL_CCO);
  const iEstado = encabezados.indexOf(CONFIG.COL_ESTADO);
  const iFecha  = encabezados.indexOf(CONFIG.COL_FECHA);

  if (iCorreo === -1) {
    ui.alert('No encontré la columna "' + CONFIG.COL_CORREO + '" en la primera fila.');
    return;
  }

  // Contar pendientes para confirmar.
  const pendientes = filas.filter(f =>
    f[iCorreo] && (iEstado === -1 || f[iEstado] !== 'ENVIADO')
  ).length;

  if (pendientes === 0) {
    ui.alert('No hay correos pendientes por enviar.');
    return;
  }

  const cuota = MailApp.getRemainingDailyQuota();
  const r = ui.alert(
    'Confirmar envío',
    'Se enviarán hasta ' + pendientes + ' correos.\n' +
    'Cuota disponible hoy: ' + cuota + '.\n' +
    '(Recuerda: cada CC y CCO también consume cuota.)\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r !== ui.Button.YES) return;

  let enviados = 0, errores = 0, omitidos = 0;

  for (let r = 0; r < filas.length; r++) {
    const fila = filas[r];
    const filaHoja = r + 2; // +1 por encabezado, +1 porque las filas inician en 1
    const correo = (fila[iCorreo] || '').toString().trim();

    if (!correo) { omitidos++; continue; }
    if (iEstado !== -1 && fila[iEstado] === 'ENVIADO') { omitidos++; continue; }

    // Detener si la cuota se agota.
    if (MailApp.getRemainingDailyQuota() <= 0) {
      ui.alert('Se alcanzó la cuota diaria de Gmail. Vuelve a ejecutar mañana; los enviados ya quedaron marcados.');
      break;
    }

    try {
      const personalizado = combinarPlantilla_(plantilla, encabezados, fila);

      const opciones = {
        htmlBody:     personalizado.html,
        attachments:  plantilla.adjuntos,
        inlineImages: plantilla.imagenesEnLinea
      };
      if (CONFIG.NOMBRE_REMITENTE) opciones.name    = CONFIG.NOMBRE_REMITENTE;
      if (CONFIG.RESPONDER_A)      opciones.replyTo = CONFIG.RESPONDER_A;

      // CC y CCO (admiten varios correos separados por coma).
      if (iCC  !== -1 && fila[iCC])  opciones.cc  = fila[iCC].toString().trim();
      if (iCCO !== -1 && fila[iCCO]) opciones.bcc = fila[iCCO].toString().trim();

      GmailApp.sendEmail(correo, personalizado.asunto, personalizado.texto, opciones);

      if (iEstado !== -1) hoja.getRange(filaHoja, iEstado + 1).setValue('ENVIADO');
      if (iFecha  !== -1) hoja.getRange(filaHoja, iFecha  + 1).setValue(new Date());
      enviados++;

    } catch (e) {
      if (iEstado !== -1) hoja.getRange(filaHoja, iEstado + 1).setValue('ERROR: ' + e.message);
      errores++;
    }

    SpreadsheetApp.flush();
    Utilities.sleep(CONFIG.PAUSA_MS);
  }

  ui.alert('Proceso terminado.\n\nEnviados: ' + enviados +
           '\nErrores: ' + errores +
           '\nOmitidos: ' + omitidos);
}


/** Envía un solo correo de prueba a tu propia dirección usando la primera fila con datos. */
function enviarPrueba() {
  const ui = SpreadsheetApp.getUi();
  const plantilla = obtenerPlantillaDesdeBorrador_(CONFIG.ASUNTO_BORRADOR);
  if (!plantilla) {
    ui.alert('No encontré el borrador con asunto: "' + CONFIG.ASUNTO_BORRADOR + '"');
    return;
  }

  const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_CONTACTOS);
  const datos = hoja.getDataRange().getDisplayValues();
  const encabezados = datos[0];
  const primeraFila = datos[1] || encabezados.map(() => 'EJEMPLO');

  const destino = CONFIG.CORREO_PRUEBA || Session.getActiveUser().getEmail();
  const personalizado = combinarPlantilla_(plantilla, encabezados, primeraFila);

  const opciones = {
    htmlBody:     personalizado.html,
    attachments:  plantilla.adjuntos,
    inlineImages: plantilla.imagenesEnLinea
  };
  if (CONFIG.NOMBRE_REMITENTE) opciones.name    = CONFIG.NOMBRE_REMITENTE;
  if (CONFIG.RESPONDER_A)      opciones.replyTo = CONFIG.RESPONDER_A;
  // En la prueba NO se copian CC/CCO para no molestar a terceros.

  GmailApp.sendEmail(destino, '[PRUEBA] ' + personalizado.asunto, personalizado.texto, opciones);
  ui.alert('Correo de prueba enviado a: ' + destino + '\n(En la prueba no se envían copias CC/CCO.)');
}


/** Borra la columna Estado y FechaEnvio para reenviar a todos. */
function reiniciarEstados() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert('Esto borrará Estado y FechaEnvio de todas las filas. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  const hoja = SpreadsheetApp.getActive().getSheetByName(CONFIG.HOJA_CONTACTOS);
  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0];
  const iEstado = encabezados.indexOf(CONFIG.COL_ESTADO);
  const iFecha  = encabezados.indexOf(CONFIG.COL_FECHA);
  const nFilas  = datos.length - 1;
  if (nFilas <= 0) return;

  if (iEstado !== -1) hoja.getRange(2, iEstado + 1, nFilas, 1).clearContent();
  if (iFecha  !== -1) hoja.getRange(2, iFecha  + 1, nFilas, 1).clearContent();
  ui.alert('Estados reiniciados.');
}


/** Muestra la cuota de correos restante para hoy. */
function verCuota() {
  SpreadsheetApp.getUi().alert('Correos que puedes enviar hoy: ' + MailApp.getRemainingDailyQuota());
}


// ===================== FUNCIONES INTERNAS =====================

/**
 * Busca un borrador de Gmail por su asunto y extrae asunto, cuerpo (texto/HTML),
 * adjuntos e imágenes en línea.
 */
function obtenerPlantillaDesdeBorrador_(asunto) {
  try {
    const borradores = GmailApp.getDrafts();
    const borrador = borradores.find(d => d.getMessage().getSubject() === asunto);
    if (!borrador) return null;

    const msg = borrador.getMessage();

    // Adjuntos "reales" (excluye imágenes en línea).
    const adjuntos = msg.getAttachments({ includeInlineImages: false });

    // Imágenes en línea (las que van dentro del HTML con cid:).
    const imagenesInline = msg.getAttachments({ includeInlineImages: true, includeAttachments: false });
    const porNombre = {};
    imagenesInline.forEach(img => { porNombre[img.getName()] = img; });

    const html = msg.getBody();
    const imagenesEnLinea = {};
    // Mapea cada src="cid:xxx" con su imagen, usando el atributo alt como nombre.
    const regex = /<img[^>]*src="cid:([^"]+)"[^>]*alt="([^"]*)"[^>]*>/g;
    let m;
    while ((m = regex.exec(html)) !== null) {
      const cid = m[1];
      const alt = m[2];
      if (porNombre[alt]) imagenesEnLinea[cid] = porNombre[alt];
    }

    return {
      asunto: asunto,
      texto:  msg.getPlainBody(),
      html:   html,
      adjuntos: adjuntos,
      imagenesEnLinea: imagenesEnLinea
    };
  } catch (e) {
    return null;
  }
}


/**
 * Reemplaza los marcadores {{Encabezado}} con los valores de la fila.
 */
function combinarPlantilla_(plantilla, encabezados, fila) {
  let asunto = plantilla.asunto;
  let texto  = plantilla.texto;
  let html   = plantilla.html;

  for (let c = 0; c < encabezados.length; c++) {
    const clave = CONFIG.MARCADOR_INICIO + encabezados[c] + CONFIG.MARCADOR_FIN;
    const valor = (fila[c] != null ? fila[c] : '').toString();
    asunto = reemplazarTodo_(asunto, clave, valor);
    texto  = reemplazarTodo_(texto,  clave, valor);
    html   = reemplazarTodo_(html,   clave, escaparHtml_(valor));
  }
  return { asunto: asunto, texto: texto, html: html };
}


/** Reemplazo global sin usar RegExp (evita problemas con caracteres especiales). */
function reemplazarTodo_(cadena, busca, reemplazo) {
  return cadena.split(busca).join(reemplazo);
}


/** Escapa caracteres HTML para los valores que van dentro del cuerpo HTML. */
function escaparHtml_(s) {
  return s.replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');
}
