# Catering-Add-in für Outlook

## Zweck
Outlook-Add-in (Termin bearbeiten/erstellen), mit dem Catering für einen Termin
bestellt wird. Die Task Pane sammelt die Angaben und schickt sie als JSON an
einen Power-Automate-Flow, der die E-Mail an `empfang@bruening-group.de`
versendet. Der Besteller steht in der Mail (Name, E-Mail, Telefon); Absender
ist das Konto des Flows.

Mit jeder Anfrage erzeugt das Add-in automatisch einen **vorausgefüllten
Bewirtungsbeleg** (Formular BH-H-BM0063) als PDF und übergibt ihn an den Flow,
der ihn an die Mail an den Empfang anhängt. Der Besteller sieht das PDF nicht;
der Empfang ist für die Bewirtungsbelege allein verantwortlich.

## Aufbau
| Pfad | Inhalt |
|---|---|
| `manifest.xml` | Add-in-Manifest (Version, URLs, Berechtigung `ReadWriteItem`) |
| `src/taskpane/` | Formular (HTML/CSS/JS) |
| `src/taskpane/beleg.js` | Erzeugung des Bewirtungsbelegs (PDF, Base64) inkl. Feldzuordnung |
| `src/lib/pdf-lib.min.js` | pdf-lib 1.17.1 (MIT, Lizenz in `pdf-lib.LICENSE.md`), lokal eingebunden |
| `assets/bewirtungsbeleg-vorlage.pdf` | Ausfüllbare Vorlage BH-H-BM0063 |
| `backend/README-Power-Automate.md` | Flow-Einrichtung (aktiv genutztes Backend) |
| `backend/azure-function-graph/` | Alternative Azure Function (nicht aktiv) |

## Installation / Bereitstellung
1. Dateien aus `src/` und `assets/` auf dem HTTPS-Hosting aktualisieren
   (derzeit GitHub Pages, siehe URLs in `manifest.xml`).
2. Flow gemäß `backend/README-Power-Automate.md` anpassen (Schema und
   Mail-Tabelle um die zwei neuen Felder ergänzen). Ohne Anpassung stehen die
   neuen Felder nur im maschinenlesbaren Block der Mail.
3. `manifest.xml` im M365 Admin Center (Integrierte Apps) aktualisieren.

## Bedienung
Termin öffnen → Menüband **Catering bestellen** → Formular ausfüllen →
**Anfrage senden**. Pflichtfelder sind mit * markiert.

### Bewirtungsbeleg (PDF) – für den Empfang
Beim Klick auf **Anfrage senden** wird das PDF im Hintergrund erzeugt und als
Anhang `Bewirtungsbeleg_JJJJ-MM-TT_<Anfrage-ID>.pdf` an die Mail an den Empfang
gehängt (Flow-Anpassung nötig, siehe `backend/README-Power-Automate.md`).
Schlägt die PDF-Erzeugung fehl, wird die Anfrage trotzdem gesendet; die Mail
enthält dann den Hinweis „kein Anhang“.

| Feld im Beleg | Quelle |
|---|---|
| Tag der Bewirtung | Datum des Terminbeginns |
| Ort der Bewirtung | Terminort (2 Zeilen, umbrochen) |
| Bewirtete Personen | Zeile 1: Besteller/in als Gastgeber/in; danach Gäste aus „Zu bewirtende Personen“ mit Firma. Passen nicht alle einzeln, kompakt als „Firma: Name, Name, …“ |
| Anlass der Bewirtung | Terminbetreff (2 Zeilen, umbrochen) |
| „In anderen Fällen“ | angekreuzt (Catering im Haus) |
| Höhe, Trinkgeld, Summe, Ort/Datum, Unterschrift | bleiben leer (Empfang) |

Die Formularfelder bleiben ausfüllbar, damit der Empfang Beträge und Datum
später eintragen kann. Passen nicht alle Gäste aufs Formular, steht ein
Hinweis in der Mail. Zeichen, die die Standardschrift nicht kennt, werden
ersetzt (z. B. „ł“ → „l“).

## Konfiguration
- `CATERING_ENDPOINT` in `src/taskpane/taskpane.js`: HTTP-Trigger-URL des Flows.
- Empfängeradresse: im Flow („An“).
- `BELEG_CONFIG` in `src/taskpane/beleg.js`: Pfad der Vorlage, Dateiname,
  ob „Ort, Datum“ vorausgefüllt wird (`fillPlaceDate`, Standard aus;
  `signPlace`), Schriftgröße und die **Feldnamen der Vorlage**
  („Textfeld 67“ … „Kontrollkästchen 45“). Bei einer neuen Formularversion
  müssen die Feldnamen geprüft und ggf. angepasst werden.

## Datenformat (Auszug)
```json
"catering": {
  "costCenter": "BGG",
  "guestCompany": "Beispiel AG",
  "guestNames": "Erika Beispiel\nMax Muster"
},
"bewirtungsbeleg": {
  "fileName": "Bewirtungsbeleg_2026-11-05_cat-1793620000000-ab12cd.pdf",
  "contentType": "application/pdf",
  "contentBase64": "JVBERi0xLjcK…",
  "personsTruncated": false
},
"bewirtungsbelegError": null
```

## Sicherheitshinweise
- Die Flow-URL enthält ein Zugriffstoken und steht im gehosteten JavaScript –
  wie ein Passwort behandeln, nicht weitergeben.
- Das Flow-Konto sollte ein dediziertes Dienstkonto mit minimalen Rechten sein.
- Keine Zugangsdaten im Code; `local.settings.json` nur mit Platzhaltern.
- Eingaben werden clientseitig begrenzt (Firma 200, Personen 1000 Zeichen);
  serverseitige Prüfung im Flow empfohlen.
- pdf-lib wird lokal ausgeliefert (kein CDN). Version 1.17.1 ist die aktuelle,
  wird aber seit 2021 nicht mehr weiterentwickelt; `npm audit` meldete am
  2026-11-02 keine bekannten Schwachstellen. Regelmäßig prüfen.
- Die Belegvorlage liegt wie der übrige Code auf dem öffentlichen Hosting
  (derzeit GitHub Pages) und ist damit für jeden abrufbar. Ist das nicht
  gewünscht, Hosting auf eine interne Quelle (SharePoint/Azure) umstellen.

## Datenschutzhinweise (DSGVO)
- Neu verarbeitet werden **Namen externer Gäste** und deren Firma. Felder sind
  optional; nur ausfüllen, wenn für die Bewirtung nötig (Datenminimierung).
- Rechtsgrundlage klären (vermutlich Art. 6 Abs. 1 lit. f DSGVO – berechtigtes
  Interesse an der Organisation der Bewirtung; ggf. steuerliche Bewirtungs-
  nachweise nach § 4 Abs. 5 Nr. 2 EStG → lit. c). Bitte mit dem
  Datenschutzbeauftragten abstimmen.
- **Löschfrist** festlegen für die Mails im Postfach empfang@ und in
  „Gesendete Elemente“ sowie für den Power-Automate-Ausführungsverlauf
  (Standard-Aufbewahrung des Verlaufs: 28 Tage).
- Power Automate / Exchange Online sind Microsoft-Dienste: Abdeckung durch den
  bestehenden Microsoft-AV-Vertrag (DPA) prüfen.
- Der Bewirtungsbeleg enthält dieselben Daten wie die Mail (Gästenamen,
  Firma) und wird über den Flow an empfang@ übertragen; er liegt danach im
  Postfach empfang@, in „Gesendete Elemente“ des Flow-Kontos und im
  Flow-Ausführungsverlauf (Standard 28 Tage, „Secure inputs“ empfohlen).
  Der fertige Beleg ist ein Buchungsbeleg: Aufbewahrung nach § 147 AO
  (Buchungsbelege nach derzeitigem Stand 8 Jahre – mit Buchhaltung
  bestätigen). Die Mail-Kopien davon unabhängig nach Löschfrist löschen.
- Allergieangaben können Gesundheitsdaten (Art. 9 DSGVO) sein – nur anonym
  („2x glutenfrei“), ohne Namen erfassen.

Dies ist keine Rechtsberatung.

## Änderungsprotokoll
| Version | Datum | Änderung |
|---|---|---|
| 1.0.0.8 | 2026-11-03 | Bewirtungsbeleg wird nicht mehr beim Besteller heruntergeladen, sondern beim Senden automatisch erzeugt und per Flow an die Mail an den Empfang angehängt (JSON-Feld `bewirtungsbeleg`, Base64); Button entfernt; „Ort, Datum“ bleibt leer; Flow-Anleitung (Schema, Anlage, removeProperty, Secure inputs) ergänzt |
| 1.0.0.7 | 2026-11-02 | Neu: Button „Bewirtungsbeleg (PDF) erstellen“ – füllt Formular BH-H-BM0063 lokal mit pdf-lib vor (Tag, Ort, Personen, Anlass, Datum, „In anderen Fällen“); Vorlage unter `assets/`, pdf-lib unter `src/lib/`; Dokumentation ergänzt |
| 1.0.0.6 | 2026-11-02 | Neue optionale Felder „Zu bewirtende Firma“ und „Zu bewirtende Personen“ (Formular, JSON, Flow-Vorlage, Azure-Function-Vorlage); Dokumentation ergänzt |
| 1.0.0.5 | – | Ausgangsstand (vor dieser Änderung) |
