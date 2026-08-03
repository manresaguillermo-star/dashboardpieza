// Google Apps Script — Empretienda → Supabase webhook
// Instalación: Ejecutar > configurar trigger → onNewEmail (cada 5 min, Time-driven)

const WEBHOOK_URL    = "https://vvcthfxbksbmzjztkymo.supabase.co/functions/v1/nueva-venta";
const WEBHOOK_SECRET = "Guille46070043"; // Pegar el mismo valor que WEBHOOK_SECRET en la Edge Function
const GMAIL_LABEL    = "Empretienda";  // Etiqueta opcional para filtrar; dejar "" para buscar por asunto
const PROCESSED_PROP = "lastProcessedDate";

function onNewEmail() {
  const props        = PropertiesService.getScriptProperties();
  const lastDate     = new Date(props.getProperty(PROCESSED_PROP) || "2000-01-01");
  const query        = 'subject:"Tienes una nueva orden" newer_than:1d';
  const threads      = GmailApp.search(query, 0, 50);
  let   newest       = lastDate;

  for (const thread of threads) {
    for (const msg of thread.getMessages()) {
      const msgDate = msg.getDate();
      if (msgDate <= lastDate) continue;
      if (msgDate > newest)   newest = msgDate;

      const subject = msg.getSubject();
      const match   = subject.match(/orden #(\d+)/i);
      if (!match) continue;

      const orderNumber = parseInt(match[1], 10);
      sendToWebhook(orderNumber);
    }
  }

  props.setProperty(PROCESSED_PROP, newest.toISOString());
}

function sendToWebhook(orderNumber) {
  const payload = JSON.stringify({ order_number: orderNumber });
  const options = {
    method: "post",
    contentType: "application/json",
    payload: payload,
    headers: {
      "x-webhook-secret": WEBHOOK_SECRET,
    },
    muteHttpExceptions: true,
  };

  try {
    const response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    const code     = response.getResponseCode();
    const body     = response.getContentText();
    Logger.log("orden #%s → HTTP %s: %s", orderNumber, code, body);
  } catch (e) {
    Logger.log("Error enviando orden #%s: %s", orderNumber, e.message);
  }
}

// Util: crear el trigger automáticamente (ejecutar una sola vez)
function crearTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("onNewEmail")
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log("Trigger creado: onNewEmail cada 5 minutos");
}
