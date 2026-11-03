/*
 * beleg.js
 * --------
 * Erzeugt beim Absenden der Catering-Anfrage einen vorausgefüllten
 * Bewirtungsbeleg (Formularvorlage BH-H-BM0063) als PDF. Das PDF wird NICHT
 * dem Besteller angezeigt oder heruntergeladen, sondern als Base64 im JSON an
 * den Power-Automate-Flow übergeben, der es an die Mail an den Empfang anhängt.
 * Der Empfang ist für die Bewirtungsbelege allein verantwortlich.
 *
 * Erzeugt wird mit pdf-lib (MIT-Lizenz, lokal eingebunden als
 * src/lib/pdf-lib.min.js, Version 1.17.1). Vorausgefüllt werden nur Angaben aus
 * Termin und Formular (Tag, Ort, bewirtete Personen, Anlass, Kästchen
 * „In anderen Fällen“). Beträge, Ort/Datum und Unterschrift bleiben leer.
 *
 * Abhängigkeiten: globales PDFLib, itemContext aus taskpane.js, Office.
 */

/* global PDFLib, itemContext, Office */

// ---------------------------------------------------------------------------
// Konfiguration – an die Vorlage angepasst (Feldnamen bzw. Koordinaten).
// ---------------------------------------------------------------------------
const BELEG_CONFIG = {
  templateUrl: "../../assets/bewirtungsbeleg-vorlage.pdf",
  fileNamePrefix: "Bewirtungsbeleg",
  // Feld „Ort, Datum“ vorausfüllen? Standard: nein (füllt der Empfang beim
  // Abschluss aus). Bei true: signPlace + Erstellungsdatum.
  fillPlaceDate: false,
  signPlace: "",
  // Bei Fehlern in der PDF-Erzeugung wird die Anfrage trotzdem gesendet
  // (ohne Anhang, mit Hinweis im JSON-Feld bewirtungsbelegError).
  fontSize: 10,
  // Feldnamen der Vorlage BH-H-BM0063 (ausgelesen am 2026-11-02).
  // Bei einer neuen Formularversion hier anpassen.
  fields: {
    date: "Textfeld 67",                       // Tag der Bewirtung
    place: ["Textfeld 68", "Textfeld 69"],     // Ort der Bewirtung (2 Zeilen)
    persons: [                                 // Bewirtete Personen (8 Zeilen)
      "Textfeld 70", "Textfeld 71", "Textfeld 72", "Textfeld 73",
      "Textfeld 74", "Textfeld 75", "Textfeld 76", "Textfeld 77"
    ],
    occasion: ["Textfeld 78", "Textfeld 79"],  // Anlass der Bewirtung (2 Zeilen)
    amount: "Textfeld 80",                     // Höhe der Aufwendungen (leer)
    tip: "Textfeld 81",                        // Trinkgeld (leer)
    total: "Textfeld 82",                      // Summe (leer)
    placeDate: "Textfeld 83",                  // Ort, Datum
    checkRestaurant: "Kontrollkästchen 44",    // Bei Bewirtung in Gaststätte
    checkOther: "Kontrollkästchen 45"          // In anderen Fällen
  }
};

/**
 * Erzeugt den Bewirtungsbeleg für die Anfrage.
 * @param {string} requestId  ID der Catering-Anfrage (für den Dateinamen)
 * @returns {Promise<{fileName: string, contentType: string, contentBase64: string, personsTruncated: boolean}>}
 */
async function buildBelegAttachment(requestId) {
  if (typeof PDFLib === "undefined") throw new Error("pdf-lib nicht geladen");

  const data = collectBelegData();
  const response = await fetch(BELEG_CONFIG.templateUrl, { cache: "no-cache" });
  if (!response.ok) throw new Error(`Vorlage nicht abrufbar (Status ${response.status})`);
  const templateBytes = await response.arrayBuffer();

  const contentBase64 = await fillBeleg(templateBytes, data);
  return {
    fileName: buildFileName(data.dateIso, requestId),
    contentType: "application/pdf",
    contentBase64,
    personsTruncated: Boolean(data.personsTruncated)
  };
}

// ---------------------------------------------------------------------------
// Daten aus Termin und Formular zusammenstellen
// ---------------------------------------------------------------------------
function collectBelegData() {
  const start = itemContext.start ? new Date(itemContext.start) : new Date();
  const dateFmt = new Intl.DateTimeFormat("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });

  const company = document.getElementById("guestCompany").value.trim().slice(0, 200);
  const names = document.getElementById("guestNames").value
    .split(/\r?\n/)
    .map((n) => n.trim())
    .filter(Boolean)
    .slice(0, 50);

  const host = Office.context.mailbox.userProfile.displayName || "";
  return {
    date: dateFmt.format(start),
    dateIso: toIsoDate(start),
    place: itemContext.location && itemContext.location !== "–" ? itemContext.location : "",
    host,
    company,
    names,
    occasion: itemContext.subject && itemContext.subject !== "(ohne Betreff)" ? itemContext.subject : "",
    placeDate: BELEG_CONFIG.fillPlaceDate
      ? [BELEG_CONFIG.signPlace, dateFmt.format(new Date())].filter(Boolean).join(", ")
      : ""
  };
}

function toIsoDate(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// Dateiname nur aus Datum und Anfrage-ID (keine Personennamen im Dateinamen).
function buildFileName(dateIso, requestId) {
  const id = String(requestId || "").replace(/[^A-Za-z0-9-]/g, "").slice(0, 40);
  return `${BELEG_CONFIG.fileNamePrefix}_${dateIso}${id ? "_" + id : ""}.pdf`;
}

// ---------------------------------------------------------------------------
// PDF befüllen
// ---------------------------------------------------------------------------
async function fillBeleg(templateBytes, data) {
  const { PDFDocument, StandardFonts } = PDFLib;
  const pdfDoc = await PDFDocument.load(templateBytes);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  const form = pdfDoc.getForm();
  const f = BELEG_CONFIG.fields;
  const size = BELEG_CONFIG.fontSize;

  // Alle Textfelder zuerst leeren (die Vorlage kann Restwerte enthalten).
  for (const field of form.getFields()) {
    if (field instanceof PDFLib.PDFTextField) field.setText("");
  }

  const fieldWidth = (name) => form.getTextField(name).acroField.getWidgets()[0].getRectangle().width - 4;

  // Setzt Text; zu lange Zeilen werden verkleinert (bis 7 pt) und notfalls gekürzt.
  const setText = (name, value) => {
    const tf = form.getTextField(name);
    let text = sanitizeForFont(value, font);
    const maxW = fieldWidth(name);
    let fs = size;
    const w = font.widthOfTextAtSize(text, size);
    if (w > maxW) fs = Math.max(7, Math.floor((size * maxW / w) * 10) / 10);
    while (text.length > 1 && font.widthOfTextAtSize(text, fs) > maxW) {
      text = text.slice(0, -2) + "…";
    }
    tf.setText(text);
    tf.setFontSize(fs);
  };

  // Mehrzeilige Angaben über mehrere einzeilige Felder verteilen.
  const setLines = (names, lines) => {
    distributeLines(lines, names.length).forEach((line, i) => setText(names[i], line));
  };
  const wrapFor = (names, text) => {
    const clean = sanitizeForFont(text, font);
    return wrapText(clean, font, size, Math.min(...names.map(fieldWidth)));
  };

  setText(f.date, data.date);
  setLines(f.place, wrapFor(f.place, data.place));
  const personLines = buildPersonLines(data, f.persons, font, size, fieldWidth);
  data.personsTruncated = personLines.truncated;
  setLines(f.persons, personLines.lines);
  setLines(f.occasion, wrapFor(f.occasion, data.occasion));
  if (data.placeDate) setText(f.placeDate, data.placeDate);

  // Catering im Haus = „In anderen Fällen“.
  form.getCheckBox(f.checkRestaurant).uncheck();
  form.getCheckBox(f.checkOther).check();

  form.updateFieldAppearances(font);
  // Felder bleiben ausfüllbar, damit der Empfang Beträge später eintragen kann.
  return pdfDoc.saveAsBase64();
}

// Bewirtete Personen: erste Zeile Gastgeber/in, danach Gäste.
// Passen alle Gäste einzeln (je Zeile „Name (Firma)“), wird so verteilt;
// sonst kompakt als „Firma: Name, Name, …“ über die restlichen Zeilen umbrochen.
function buildPersonLines(data, names, font, size, fieldWidth) {
  const lines = [];
  if (data.host) lines.push(`${data.host} (Gastgeber/in, Brüning Group)`);
  const slots = names.length - lines.length;
  const guestWidth = Math.min(...names.slice(lines.length).map(fieldWidth));
  const guests = data.names.map((n) => sanitizeForFont(n, font));
  const company = sanitizeForFont(data.company, font);

  if (guests.length === 0) {
    if (company) lines.push(`Gäste der ${company} (Namen bitte ergänzen)`);
    return { lines, truncated: false };
  }

  const single = guests.map((n) => (company ? `${n} (${company})` : n));
  if (single.length <= slots && single.every((l) => font.widthOfTextAtSize(l, size) <= guestWidth)) {
    return { lines: lines.concat(single), truncated: false };
  }

  const items = guests.map((n, i) => (i < guests.length - 1 ? `${n},` : n));
  if (company) items.unshift(`${company}:`);
  const wrapped = wrapItems(items, font, size, guestWidth);
  return { lines: lines.concat(wrapped), truncated: wrapped.length > slots };
}

// Häufige Zeichen ohne WinAnsi-Entsprechung, die NFD nicht zerlegt.
const CHAR_FALLBACK = { "ł": "l", "Ł": "L", "đ": "d", "Đ": "D", "ı": "i", "ħ": "h", "Ħ": "H" };

// Ersetzt Zeichen, die die Standardschrift (WinAnsi) nicht darstellen kann,
// z. B. „ł“ → „l“; nicht abbildbare Zeichen werden zu „?“.
function sanitizeForFont(text, font) {
  let out = "";
  for (const ch of String(text || "")) {
    if (ch === "\n") { out += ch; continue; }
    if (canEncode(font, ch)) { out += ch; continue; }
    const base = CHAR_FALLBACK[ch] || ch.normalize("NFD").replace(/[̀-ͯ]/g, "");
    out += base && canEncode(font, base) ? base : "?";
  }
  return out;
}

function canEncode(font, ch) {
  try {
    font.widthOfTextAtSize(ch, 10);
    return true;
  } catch (e) {
    return false;
  }
}

// Bricht Text so um, dass jede Zeile in maxWidth passt.
function wrapText(text, font, size, maxWidth) {
  const words = String(text || "").split(/\s+/).filter(Boolean);
  const lines = [];
  let current = "";
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth || !current) {
      current = candidate;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines;
}

// Wie wrapText, bricht aber nur zwischen ganzen Einträgen (z. B. Namen) um;
// ein einzelner zu langer Eintrag wird wortweise umbrochen.
function wrapItems(items, font, size, maxWidth) {
  const lines = [];
  let current = "";
  for (const item of items) {
    const candidate = current ? `${current} ${item}` : item;
    if (font.widthOfTextAtSize(candidate, size) <= maxWidth) {
      current = candidate;
      continue;
    }
    if (current) lines.push(current);
    if (font.widthOfTextAtSize(item, size) <= maxWidth) {
      current = item;
    } else {
      const parts = wrapText(item, font, size, maxWidth);
      current = parts.pop() || "";
      lines.push(...parts);
    }
  }
  if (current) lines.push(current);
  return lines;
}

// Verteilt Zeilen auf eine feste Anzahl Felder; Überhang wird in der letzten
// Zeile angehängt (und dort ggf. verkleinert/gekürzt).
function distributeLines(lines, slots) {
  if (lines.length <= slots) return lines.concat(Array(slots - lines.length).fill(""));
  const head = lines.slice(0, slots - 1);
  head.push(lines.slice(slots - 1).join(" "));
  return head;
}
