/*
 * taskpane.js
 * -----------
 * Liest die Termindaten aus dem gerade erstellten/bearbeiteten Termin,
 * zeigt das Catering-Formular an und sendet die Anfrage in einem festen,
 * maschinenlesbaren Format an ein Backend (Power Automate Flow oder
 * Azure Function – siehe /backend), das die eigentliche E-Mail an
 * empfang@bruening-group.de verschickt.
 *
 * WICHTIG: Der Add-in-Code selbst versendet KEINE E-Mail direkt.
 * Das hält das Add-in schlank und funktioniert unabhängig davon, ob das
 * Postfach des Anfragenden in Exchange Online oder (im Hybrid-Setup) noch
 * on-premises liegt. Siehe README.md, Abschnitt "Architekturentscheidung".
 */

// TODO: Durch die echte URL des Power-Automate-Flows / der Azure Function ersetzen.
const CATERING_ENDPOINT = "https://default536d21304a39421d8c59bc1dcba1f1.29.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/30/workflows/c818a16fd72145168ccb3cbe086c6815/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=lZh3j93BntMmsnDKVkK_uRmAftjPk69nakdMixPVypE";

let itemContext = {
  subject: "",
  start: null,
  end: null,
  location: ""
};

Office.onReady(() => {
  loadMeetingContext();
});

function loadMeetingContext() {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((subjectResult) => {
    itemContext.subject = subjectResult.value || "(ohne Betreff)";
    document.getElementById("mSubject").textContent = itemContext.subject;
  });

  item.start.getAsync((startResult) => {
    itemContext.start = startResult.value;
    item.end.getAsync((endResult) => {
      itemContext.end = endResult.value;
      document.getElementById("mWhen").textContent = formatRange(itemContext.start, itemContext.end);
      prefillDeliveryTime(itemContext.start);
    });
  });

  item.location.getAsync((locationResult) => {
    itemContext.location = locationResult.value || "–";
    document.getElementById("mLocation").textContent = itemContext.location;
  });

  // Sinnvolle Vorbelegung der Personenanzahl aus Pflicht- + optionalen Teilnehmern.
  if (item.requiredAttendees && item.optionalAttendees) {
    Promise.all([
      new Promise((res) => item.requiredAttendees.getAsync((r) => res(r.value || []))),
      new Promise((res) => item.optionalAttendees.getAsync((r) => res(r.value || [])))
    ]).then(([required, optional]) => {
      const count = required.length + optional.length + 1; // +1 für Organisator/in
      const input = document.getElementById("attendeeCount");
      if (!input.value) input.value = count;
    });
  }
}

function prefillDeliveryTime(start) {
  if (!start) return;
  const d = new Date(start);
  d.setMinutes(d.getMinutes() - 15); // Standard: 15 Min. vor Terminbeginn liefern
  const hh = String(d.getHours()).padStart(2, "0");
  const mm = String(d.getMinutes()).padStart(2, "0");
  document.getElementById("deliveryTime").value = `${hh}:${mm}`;
}

function formatRange(start, end) {
  if (!start || !end) return "–";
  const s = new Date(start);
  const e = new Date(end);
  const dateFmt = new Intl.DateTimeFormat("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
  const timeFmt = new Intl.DateTimeFormat("de-DE", { hour: "2-digit", minute: "2-digit" });
  return `${dateFmt.format(s)}, ${timeFmt.format(s)}–${timeFmt.format(e)} Uhr`;
}

document.getElementById("cateringForm").addEventListener("submit", async (event) => {
  event.preventDefault();
  await handleSubmit();
});

async function handleSubmit() {
  const statusEl = document.getElementById("statusMsg");
  const submitBtn = document.getElementById("submitBtn");

  const attendeeCount = document.getElementById("attendeeCount").value;
  const cateringType = document.getElementById("cateringType").value;
  const deliveryTime = document.getElementById("deliveryTime").value;
  const costCenter = document.getElementById("costCenter").value.trim();
  const guestCompany = document.getElementById("guestCompany").value.trim().slice(0, 200);
  const guestNames = document.getElementById("guestNames").value.trim().slice(0, 1000);
  const dietaryNotes = document.getElementById("dietaryNotes").value.trim();
  const phone = document.getElementById("phone").value.trim();

  if (!attendeeCount || !cateringType || !deliveryTime || !costCenter) {
    setStatus(statusEl, "Bitte alle Pflichtfelder (*) ausfüllen.", "error");
    return;
  }

  const profile = Office.context.mailbox.userProfile;

  // Festes, maschinenlesbares Schema für die automatische Verarbeitung
  // durch das Catering-Team / dessen Backend.
  const payload = {
    requestId: generateRequestId(),
    submittedAtUtc: new Date().toISOString(),
    meeting: {
      subject: itemContext.subject,
      startUtc: itemContext.start ? new Date(itemContext.start).toISOString() : null,
      endUtc: itemContext.end ? new Date(itemContext.end).toISOString() : null,
      location: itemContext.location
    },
    catering: {
      attendeeCount: parseInt(attendeeCount, 10),
      type: cateringType,
      deliveryTime: deliveryTime,
      dietaryNotes: dietaryNotes || null,
      costCenter: costCenter,
      guestCompany: guestCompany || null,
      guestNames: guestNames || null
    },
    requestedBy: {
      name: profile.displayName,
      email: profile.emailAddress,
      phone: phone || null
    },
    targetMailbox: "empfang@bruening-group.de"
  };

  submitBtn.disabled = true;
  setStatus(statusEl, "Anfrage wird gesendet…", "");

  // Bewirtungsbeleg (PDF) für den Empfang erzeugen und mitsenden (siehe beleg.js).
  // Schlägt das fehl, wird die Anfrage trotzdem ohne Anhang gesendet.
  payload.bewirtungsbeleg = null;
  payload.bewirtungsbelegError = null;
  try {
    payload.bewirtungsbeleg = await buildBelegAttachment(payload.requestId);
  } catch (belegErr) {
    console.error(belegErr);
    payload.bewirtungsbelegError = "PDF konnte im Add-in nicht erzeugt werden";
  }

  try {
    const response = await fetch(CATERING_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error(`Backend antwortete mit Status ${response.status}`);
    }

    await markItemAsRequested(payload.requestId);
    setStatus(statusEl, "Catering-Anfrage wurde an empfang@bruening-group.de gesendet.", "ok");
  } catch (err) {
    console.error(err);
    setStatus(statusEl, "Senden fehlgeschlagen. Bitte später erneut versuchen oder IT-Support kontaktieren.", "error");
    submitBtn.disabled = false;
  }
}

function setStatus(el, text, cls) {
  el.textContent = text;
  el.className = "status" + (cls ? " " + cls : "");
}

function generateRequestId() {
  return "cat-" + Date.now() + "-" + Math.random().toString(36).slice(2, 8);
}

// Vermerkt am Termin selbst, dass bereits eine Anfrage gestellt wurde
// (überlebt Schließen/erneutes Öffnen des Termins vor dem Senden der Einladung).
function markItemAsRequested(requestId) {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    if (!item.loadCustomPropertiesAsync) {
      resolve();
      return;
    }
    item.loadCustomPropertiesAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const props = result.value;
        props.set("CateringRequested", "true");
        props.set("CateringRequestId", requestId);
        props.saveAsync(() => resolve());
      } else {
        resolve();
      }
    });
  });
}
