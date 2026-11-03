# Backend-Option A (empfohlen): Power Automate Flow

Diese Variante braucht keinen eigenen App-Registrierungs-/Graph-Berechtigungsaufwand
und funktioniert unabhängig davon, ob das Postfach der anfragenden Person in
Exchange Online oder (im Hybrid-Setup) noch on-premises liegt – das Add-in
selbst versendet nie eine E-Mail, es ruft nur diesen Flow per HTTP auf.

## 1. Flow anlegen

1. [make.powerautomate.com](https://make.powerautomate.com) → **Erstellen** → **Instant cloud flow**
2. Trigger: **"Wenn eine HTTP-Anfrage empfangen wird"** (When a HTTP request is received)
3. Bei "Anforderungstext-JSON-Schema" folgendes Schema einfügen:

```json
{
  "type": "object",
  "properties": {
    "requestId": { "type": "string" },
    "submittedAtUtc": { "type": "string" },
    "meeting": {
      "type": "object",
      "properties": {
        "subject": { "type": "string" },
        "startUtc": { "type": "string" },
        "endUtc": { "type": "string" },
        "location": { "type": "string" }
      }
    },
    "catering": {
      "type": "object",
      "properties": {
        "attendeeCount": { "type": "integer" },
        "type": { "type": "string" },
        "deliveryTime": { "type": "string" },
        "dietaryNotes": { "type": ["string", "null"] },
        "costCenter": { "type": "string" },
        "guestCompany": { "type": ["string", "null"] },
        "guestNames": { "type": ["string", "null"] }
      }
    },
    "requestedBy": {
      "type": "object",
      "properties": {
        "name": { "type": "string" },
        "email": { "type": "string" },
        "phone": { "type": ["string", "null"] }
      }
    },
    "targetMailbox": { "type": "string" },
    "bewirtungsbeleg": {
      "type": ["object", "null"],
      "properties": {
        "fileName": { "type": "string", "maxLength": 120, "pattern": "^[A-Za-z0-9_.-]+\\.pdf$" },
        "contentType": { "type": "string", "enum": ["application/pdf"] },
        "contentBase64": { "type": "string", "maxLength": 3000000 },
        "personsTruncated": { "type": "boolean" }
      }
    },
    "bewirtungsbelegError": { "type": ["string", "null"] }
  }
}
```

4. Schritt hinzufügen: **"Send an email (V2)"** (Office 365 Outlook-Connector).
   Am besten über ein **dediziertes Service-/Flow-Konto oder Shared Mailbox**
   ausführen (nicht das persönliche Postfach eines Admins), damit der Flow
   auch bei Mitarbeiterwechsel weiterläuft.

   - **An:** `empfang@bruening-group.de`
   - **Betreff:** `Catering-Anfrage: [meeting_subject] am [meeting_startUtc]`
   - **Text (HTML):** Kombination aus lesbarer Übersicht + maschinenlesbarem
     Block, z. B.:

```html
<h3>Neue Catering-Anfrage</h3>
<table>
  <tr><td><b>Termin</b></td><td>@{triggerBody()?['meeting']?['subject']}</td></tr>
  <tr><td><b>Datum/Zeit</b></td><td>@{triggerBody()?['meeting']?['startUtc']} – @{triggerBody()?['meeting']?['endUtc']}</td></tr>
  <tr><td><b>Ort</b></td><td>@{triggerBody()?['meeting']?['location']}</td></tr>
  <tr><td><b>Personen</b></td><td>@{triggerBody()?['catering']?['attendeeCount']}</td></tr>
  <tr><td><b>Art</b></td><td>@{triggerBody()?['catering']?['type']}</td></tr>
  <tr><td><b>Lieferzeit</b></td><td>@{triggerBody()?['catering']?['deliveryTime']}</td></tr>
  <tr><td><b>Kostenstelle</b></td><td>@{triggerBody()?['catering']?['costCenter']}</td></tr>
  <tr><td><b>Zu bewirtende Firma</b></td><td>@{coalesce(triggerBody()?['catering']?['guestCompany'], '–')}</td></tr>
  <tr><td><b>Zu bewirtende Personen</b></td><td>@{replace(coalesce(triggerBody()?['catering']?['guestNames'], '–'), decodeUriComponent('%0A'), '<br>')}</td></tr>
  <tr><td><b>Besonderheiten</b></td><td>@{triggerBody()?['catering']?['dietaryNotes']}</td></tr>
  <tr><td><b>Angefragt von</b></td><td>@{triggerBody()?['requestedBy']?['name']} (@{triggerBody()?['requestedBy']?['email']}, @{triggerBody()?['requestedBy']?['phone']})</td></tr>
</table>

<p><b>Bewirtungsbeleg:</b> @{if(equals(triggerBody()?['bewirtungsbeleg'], null), concat('kein Anhang – ', coalesce(triggerBody()?['bewirtungsbelegError'], 'nicht erzeugt')), if(equals(triggerBody()?['bewirtungsbeleg']?['personsTruncated'], true), 'angehängt (vorausgefüllt) – ACHTUNG: nicht alle Gäste passten aufs Formular, vollständige Liste beifügen', 'angehängt (vorausgefüllt)'))}</p>

<pre>
--- CATERING-REQUEST (maschinenlesbar, bitte nicht ändern) ---
@{string(removeProperty(triggerBody(), 'bewirtungsbeleg'))}
--- ENDE ---
</pre>
```

   **Wichtig (ab 1.0.0.8):** Im maschinenlesbaren Block `removeProperty(…,
   'bewirtungsbeleg')` verwenden, sonst landet das komplette PDF als
   Base64-Text in der Mail.

   - **Anlagen (Bewirtungsbeleg-PDF):** In „Send an email (V2)“ unter
     *Erweiterte Parameter* → **Attachments** auf
     „Switch to input entire array“ umschalten und diesen Ausdruck eintragen:

```
if(equals(triggerBody()?['bewirtungsbeleg'], null), json('[]'), createArray(addProperty(addProperty(json('{}'), 'Name', triggerBody()?['bewirtungsbeleg']?['fileName']), 'ContentBytes', triggerBody()?['bewirtungsbeleg']?['contentBase64'])))
```

     Damit wird das PDF angehängt, wenn das Add-in eines erzeugt hat; sonst
     geht die Mail ohne Anhang raus. `ContentBytes` erwartet im Array-Modus
     den Base64-Text – kein `base64ToBinary()` verwenden.
     *Nicht im Tenant getestet – bitte mit einer Testanfrage prüfen, ob sich
     der Anhang öffnen lässt.*
   - **Schema-Validierung einschalten:** Trigger → *Einstellungen* →
     „Schema Validation“ = Ein. Dann lehnt der Flow zu große oder falsch
     benannte Anhänge mit HTTP 400 ab.

5. Speichern → die HTTP-Trigger-URL kopieren.
6. Diese URL in `src/taskpane/taskpane.js` bei `CATERING_ENDPOINT` eintragen.

## 2. Absicherung

- Die generierte HTTP-Trigger-URL enthält bereits ein Zugriffs-Token
  (SAS-artig) – trotzdem **nicht öffentlich dokumentieren**.
- Optional zusätzlich absichern: Azure API Management oder eine Function
  davor schalten, die z. B. das Absender-Mailbox-Postfach gegen eine erlaubte
  Domain prüft.
- HTML-Einschleusung: Freitextfelder (Personen, Firma, Wünsche) werden vom
  Flow unverändert ins HTML übernommen. Für den Empfang ungefährlich, aber
  unschön; bei Bedarf mit `replace(..., '<', '&lt;')` entschärfen.
- Empfehlenswert: Im Flow eine einfache Plausibilitätsprüfung einbauen
  (z. B. `attendeeCount` > 0, `costCenter` nicht leer) und bei Fehlern mit
  HTTP 400 antworten, damit die Task Pane eine sinnvolle Fehlermeldung zeigt.

- Datenschutz: Der **Ausführungsverlauf** des Flows speichert die
  Eingabedaten inkl. PDF (Gästenamen) – Standard 28 Tage. Optional in den
  Einstellungen von Trigger und Mail-Aktion „Secure inputs“/„Secure outputs“
  aktivieren, dann werden die Inhalte im Verlauf nicht angezeigt.

## 3. Verantwortlichkeit im Fehlerfall

Lege fest, wer den Flow im "Meine Flows"-Bereich betreut (Eigentümerschaft,
Freigabe an ein Team, Benachrichtigung bei Flow-Fehlern aktivieren:
Einstellungen → "Meine Flows" → Fehlerbenachrichtigungen).
