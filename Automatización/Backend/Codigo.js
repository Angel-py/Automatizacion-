const NOMBRES_HOJAS_FUENTES = ['Scorecard Data', 'No records'];
const NOMBRE_HOJA_PRINCIPAL = 'LSC Emails';

const CORREO_REMITENTE = ''; 
const CORREO_DESTINO = ''; 
const CORREOS_CC = ', ,';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Formulario')
      .addItem('LSC Call Verification', 'abrirModal')
      .addToUi();
}

function abrirModal() {
  const html = HtmlService.createHtmlOutputFromFile('Frontend/Formulario')
      .setWidth(450)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Control de Calidad (QA)');
}

function enviarCorreoNotificacion(headers, values, ccDinamico) {
  const usuarioActivo = Session.getActiveUser().getEmail();
  const idImagenLogo = '13Z5uiLNH_hy1Jv3yAwGvitpjvcIak_Em'; 
  
  
  const indiceScheduler = headers.indexOf("Scheduler Name");
  const nombreAgente = (indiceScheduler !== -1 && values[indiceScheduler]) ? values[indiceScheduler] : "Unknown Agent";
  
  
  const asunto = `SMS Survey Call Verification - ${nombreAgente}`;

  
  let tablaHtml = `
    <table border="0" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; border: 1px solid #c0c0c0;">
      <tr>
        <th colspan="2" style="background-color: #d8d8c0; color: #333; text-align: center; border: 1px solid #c0c0c0; text-transform: uppercase; font-size: 14px;">AUDIT DETAILS:</th>
      </tr>`;
  
  headers.forEach((header, i) => {
    let valor = values[i] || '---';
    let estiloCeldaValor = 'color: #333;'; 
    let valorLimpio = String(valor).trim().toUpperCase();

   
    if (header === "Red Flag") {
      if (valorLimpio === "NO") {
        estiloCeldaValor = 'color: #008000; font-weight: bold;';
      } else if (valorLimpio === "YES") {
        estiloCeldaValor = 'color: #FF0000; font-weight: bold;';
      }
    }

  
    if (header === "Result") {
      if (valorLimpio === "VALID") {
        estiloCeldaValor = 'color: #008000; font-weight: bold;';
      } else if (valorLimpio === "NOT VALID") {
        estiloCeldaValor = 'color: #FF0000; font-weight: bold;';
      }
    }

    
    if (header === "Agent Recording" || header === "Recording Link (URL)") {
      if (!valor || valor === "---" || valor === "Sin audio proporcionado" || valorLimpio === "") {
        valor = "Audio is not provided";
      } else if (String(valor).startsWith("http")) {
        valor = `<a href="${valor}" style="color: #0b57d0; text-decoration: none; font-weight: bold;">View Recording</a>`;
      }
    }

    tablaHtml += `
      <tr>
        <td style="padding: 10px; background-color: #8dbadb; width: 35%; border: 1px solid #c0c0c0; font-weight: bold; color: #000;">${header}:</td>
        <td style="padding: 10px; border: 1px solid #c0c0c0; text-align: center; ${estiloCeldaValor}">${valor}</td>
      </tr>`;
  });
  tablaHtml += '</table>';

  
  const cuerpoHtml = `
    <div style="max-width: 700px; margin: auto; font-family: Arial, sans-serif; color: #333;">
      
      <table border="0" cellpadding="0" cellspacing="0" style="width: 100%; background-color: #3b5998; margin-bottom: 20px;">
        <tr>
          <td style="padding: 20px; text-align: center;">
            <img src="cid:logoOclinicals" style="display: block; width: 100%; max-width: 700px; margin: 0 auto;">
          </td>
        </tr>
      </table>

      <p style="font-size: 14px;">Hi team,</p>
      <p style="font-size: 14px;">Attached to this message, you will find the final results...</p>
      
      <br>
      ${tablaHtml}
      <br>

      <p style="font-size: 14px;">Thank you for your concern.</p>
      
      <div style="color: #777; font-size: 13px; line-height: 1.6;">
        --<br>
        <strong>Best regards,</strong><br>
        <span style="color: #3b5998; font-weight: bold;">SMS Survey & QA Team</span><br>
        Oclinicals • 19495 Biscanyne Blvd. #608 • Aventura, FL 33180<br>
        <a href="http://www.oclinicals.com" style="color: #0b57d0; text-decoration: none;">www.oclinicals.com</a>
      </div>
    </div>
  `;

  
  try {
    const logoBlob = DriveApp.getFileById(idImagenLogo).getBlob().setName("logoOclinicals");

    const opciones = {
      cc: ccDinamico,
      htmlBody: cuerpoHtml,
      replyTo: usuarioActivo,
      inlineImages: {
        logoOclinicals: logoBlob
      }
    };
    
    GmailApp.sendEmail(CORREO_DESTINO, asunto, "", opciones);
  } catch (e) {
    Logger.log("Error enviando el correo: " + e.toString());
  }
}

function procesarFormulario(obj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaDestino = ss.getSheetByName(NOMBRE_HOJA_PRINCIPAL);
  
  if (!hojaDestino) throw new Error("No se encontró la pestaña '" + NOMBRE_HOJA_PRINCIPAL + "'");

  let datosDestino = hojaDestino.getDataRange().getValues();
  let encabezadosDestino = datosDestino[0];

  const infoExterna = buscarEnPestanasInternas(ss, obj.mrn);
  if (!infoExterna) throw new Error("El MRN '" + obj.mrn + "' no existe en las bases de datos.");

  const colMrn = encabezadosDestino.indexOf("MRN");
  let filaIndice = -1;
  for (let i = 1; i < datosDestino.length; i++) {
    if (datosDestino[i][colMrn] == obj.mrn) {
      filaIndice = i + 1;
      break;
    }
  }

  if (filaIndice === -1) {
    hojaDestino.appendRow([obj.mrn]); 
    filaIndice = hojaDestino.getLastRow();
  }

  Object.keys(infoExterna).forEach(nombreCol => {
    const idx = encabezadosDestino.indexOf(nombreCol);
    if (idx !== -1 && nombreCol !== "MRN") {
      hojaDestino.getRange(filaIndice, idx + 1).setValue(infoExterna[nombreCol]);
    }
  });

  let urlAudio = "Sin audio proporcionado";
  if (!obj.sinAudio && obj.audioBytes) {
    const carpeta = obtenerOCrearCarpeta("Grabaciones_QA");
    const blob = Utilities.newBlob(Utilities.base64Decode(obj.audioBytes), obj.audioContentType, obj.audioNombre);
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urlAudio = archivo.getUrl();
  }

  const colRes = encabezadosDestino.indexOf("Result");
  const colSum = encabezadosDestino.indexOf("Summary");
  const colRec = encabezadosDestino.indexOf("Agent Recording");

  if(colRes !== -1) hojaDestino.getRange(filaIndice, colRes + 1).setValue(obj.result);
  if(colSum !== -1) hojaDestino.getRange(filaIndice, colSum + 1).setValue(obj.summary);
  if(colRec !== -1) hojaDestino.getRange(filaIndice, colRec + 1).setValue(urlAudio);
  
  let filaCompleta = hojaDestino.getRange(filaIndice, 1, 1, encabezadosDestino.length).getValues()[0];
  
  const colFecha = encabezadosDestino.indexOf("Call Date");
  if (colFecha !== -1 && filaCompleta[colFecha] instanceof Date) {
    filaCompleta[colFecha] = Utilities.formatDate(filaCompleta[colFecha], Session.getScriptTimeZone(), "MM/dd/yy");
  }
  
  const correoScheduler = filaCompleta[encabezadosDestino.indexOf("Agent Email")];
  const correoSupervisor = filaCompleta[encabezadosDestino.indexOf("Supervisor Email")];

  let listaCC = CORREOS_CC;
  if (correoScheduler && correoScheduler !== "Sin Scheduler Email proporcionado") {
    listaCC += ", " + correoScheduler;
  }
  if (correoSupervisor && correoSupervisor !== "Sin Supervisor Email proporcionado") {
    listaCC += ", " + correoSupervisor;
  }

  enviarCorreoNotificacion(encabezadosDestino, filaCompleta, listaCC);

  return "¡Éxito! Registro guardado y notificado a " + correoScheduler + " y " + correoSupervisor;
}

function buscarEnPestanasInternas(ss, mrn) {
  let resultado = null;
  const mrnBuscado = String(mrn).trim();

  const columnasExcluidas = ["Result", "Summary", "Agent Recording"];

  NOMBRES_HOJAS_FUENTES.forEach(nombreHoja => {
    if (resultado) return;
    
    const hojaFuente = ss.getSheetByName(nombreHoja);
    if (!hojaFuente) return;

    const dataFuente = hojaFuente.getDataRange().getValues();
    if (dataFuente.length < 1) return;
    
    const headersFuente = dataFuente[0];
    const idxMrnFuente = headersFuente.indexOf("MRN");

    if (idxMrnFuente === -1) return;

    for (let i = 1; i < dataFuente.length; i++) {
      let valorCelda = String(dataFuente[i][idxMrnFuente]).trim();
      
      if (valorCelda === mrnBuscado) {
        resultado = {};
        headersFuente.forEach((h, index) => {
          
          if (columnasExcluidas.indexOf(h) !== -1) return;

          let valor = dataFuente[i][index];
          
          if (valor === "" || valor === null || valor === undefined) {
            resultado[h] = "Sin " + h + " proporcionado";
          } else {
            resultado[h] = valor;
          }
        });
        break;
      }
    }
  });
  
  return resultado;
}

function obtenerOCrearCarpeta(nombre) {
  const carpetas = DriveApp.getFoldersByName(nombre);
  return carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(nombre);
}