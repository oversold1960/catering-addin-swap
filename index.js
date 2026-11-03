/*
 * Backend-Option B: Azure Function (Node.js) statt Power Automate.
 * ------------------------------------------------------------------
 * Nimmt die strukturierte Catering-Anfrage per HTTP POST entgegen und
 * versendet die E-Mail über Microsoft Graph mit einer App-Berechtigung
 * (Client-Credentials-Flow) aus einem dedizierten Service-/Shared-Postfach.
 *
 * WICHTIG (Hybrid-Umgebung):
 * Diese Variante sendet über ein Postfach, das in Exchange Online liegt
 * (das Service-/Shared-Postfach für den Versand). Es ist unabhängig davon,
 * wo das Postfach der anfragenden Person liegt, da der Add-in-Client nur
 * JSON an diese Function schickt und NICHT selbst über Graph sendet.
 *
 * Voraussetzungen in Azure AD / Entra ID:
 *  1. App-Registrierung anlegen.
 *  2. API-Berechtigung: Microsoft Graph → Application permissions → Mail.Send
 *     (KEIN "Mail.Send" als Delegated Permission - das würde einen
 *     angemeldeten Benutzer voraussetzen, den es hier nicht gibt).
 *  3. Admin-Zustimmung (Admin consent) durch einen Global-/Exchange-Admin.
 *  4. Client Secret oder Zertifikat erzeugen.
 *  5. Empfohlen: Mail.Send über eine "Application Access Policy" in Exchange
 *     Online auf genau das Absender-Postfach einschränken, damit die App
 *     nicht im Namen jedes beliebigen Postfachs senden kann:
 *       New-ApplicationAccessPolicy -AppId <AppId> `
 *         -PolicyScopeGroupId catering-relay@meinefirma.de `
 *         -AccessRight RestrictAccess -Description "Nur Catering-Relay"
 *
 * Umgebungsvariablen (Azure Function Application Settings):
 *   AAD_TENANT_ID, AAD_CLIENT_ID, AAD_CLIENT_SECRET, SENDER_MAILBOX
 */

const { ConfidentialClientApplication } = require("@azure/msal-node");

const TENANT_ID = process.env.AAD_TENANT_ID;
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;
const SENDER_MAILBOX = process.env.SENDER_MAILBOX; // z. B. catering-relay@meinefirma.de
const TARGET_MAILBOX = "empfang@bruening-group.de";

const msalClient = new ConfidentialClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET
  }
});

module.exports = async function (context, req) {
  const payload = req.body;

  const validationError = validatePayload(payload);
  if (validationError) {
    context.res = { status: 400, body: { error: validationError } };
    return;
  }

  try {
    const token = await getGraphToken();
    await sendCateringMail(token, payload);
    context.res = { status: 200, body: { status: "sent" } };
  } catch (err) {
    context.log.error("Catering-Mail-Versand fehlgeschlagen:", err);
    context.res = { status: 502, body: { error: "Mailversand fehlgeschlagen" } };
  }
};

function validatePayload(p) {
  if (!p || !p.meeting || !p.catering || !p.requestedBy) return "Unvollständige Anfrage.";
  if (!p.catering.attendeeCount || p.catering.attendeeCount < 1) return "Ungültige Personenanzahl.";
  if (!p.catering.costCenter) return "Kostenstelle fehlt.";
  return null;
}

async function getGraphToken() {
  const result = await msalClient.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"]
  });
  return result.accessToken;
}

async function sendCateringMail(token, payload) {
  const { meeting, catering, requestedBy, requestId } = payload;

  const htmlSummary = `
    <h3>Neue Catering-Anfrage</h3>
    <table>
      <tr><td><b>Termin</b></td><td>${escapeHtml(meeting.subject)}</td></tr>
      <tr><td><b>Datum/Zeit</b></td><td>${escapeHtml(meeting.startUtc)} – ${escapeHtml(meeting.endUtc)}</td></tr>
      <tr><td><b>Ort</b></td><td>${escapeHtml(meeting.location)}</td></tr>
      <tr><td><b>Personen</b></td><td>${catering.attendeeCount}</td></tr>
      <tr><td><b>Art</b></td><td>${escapeHtml(catering.type)}</td></tr>
      <tr><td><b>Lieferzeit</b></td><td>${escapeHtml(catering.deliveryTime)}</td></tr>
      <tr><td><b>Kostenstelle</b></td><td>${escapeHtml(catering.costCenter)}</td></tr>
      <tr><td><b>Zu bewirtende Firma</b></td><td>${escapeHtml(catering.guestCompany || "–")}</td></tr>
      <tr><td><b>Zu bewirtende Personen</b></td><td>${escapeHtml(catering.guestNames || "–").replace(/\n/g, "<br>")}</td></tr>
      <tr><td><b>Besonderheiten</b></td><td>${escapeHtml(catering.dietaryNotes || "–")}</td></tr>
      <tr><td><b>Angefragt von</b></td><td>${escapeHtml(requestedBy.name)} (${escapeHtml(requestedBy.email)}, ${escapeHtml(requestedBy.phone || "–")})</td></tr>
    </table>
    <pre>--- CATERING-REQUEST (maschinenlesbar) ---
${escapeHtml(JSON.stringify(payload))}
--- ENDE ---</pre>
  `;

  const message = {
    message: {
      subject: `Catering-Anfrage: ${meeting.subject} am ${meeting.startUtc}`,
      body: { contentType: "HTML", content: htmlSummary },
      toRecipients: [{ emailAddress: { address: TARGET_MAILBOX } }],
      internetMessageHeaders: [
        { name: "X-Catering-Request-Id", value: requestId || "" }
      ]
    },
    saveToSentItems: true
  };

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${SENDER_MAILBOX}/sendMail`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(message)
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph sendMail Status ${response.status}: ${text}`);
  }
}

function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
