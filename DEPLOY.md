# Trigger selbst erstellen: Azure Function Schritt für Schritt

Diese Anleitung baut auf dem Code in diesem Ordner auf (`CateringRequest/index.js`
+ `CateringRequest/function.json`) und macht daraus einen echten, öffentlich per
HTTPS erreichbaren Trigger, den du in `src/taskpane/taskpane.js` bei
`CATERING_ENDPOINT` eintragen kannst.

Ordnerstruktur (bereits vorhanden):
```
azure-function-graph/
├── host.json                  Function-App-weite Konfiguration
├── local.settings.json        Nur für lokales Testen, wird NICHT deployt
├── package.json
└── CateringRequest/
    ├── function.json          Definiert den HTTP-Trigger (POST)
    └── index.js                Die eigentliche Logik
```

## 0. Voraussetzungen

- Azure-Abo mit Rechten, Ressourcen anzulegen
- Node.js 20 LTS
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local) (`npm install -g azure-functions-core-tools@4`)
- [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) (`az login` vorher ausführen)
- Für Schritt 4: Rechte zur App-Registrierung in Entra ID (mind. "Application
  Developer"); für den Admin-Consent-Schritt braucht ihr zusätzlich einen
  Global Admin / Privileged Role Administrator

## 1. Lokal testen

```bash
cd backend/azure-function-graph
npm install
func start
```

In einem zweiten Terminal:

```bash
curl -X POST http://localhost:7071/api/CateringRequest \
  -H "Content-Type: application/json" \
  -d '{
    "requestId": "test-1",
    "meeting": {"subject": "Testtermin", "startUtc": "2026-08-10T08:00:00Z", "endUtc": "2026-08-10T09:00:00Z", "location": "Raum 1"},
    "catering": {"attendeeCount": 5, "type": "Kaffee & Gebäck", "deliveryTime": "07:45", "dietaryNotes": null, "costCenter": "4711-IT"},
    "requestedBy": {"name": "Max Mustermann", "email": "max@meinefirma.de", "phone": null},
    "targetMailbox": "empfang@bruening-group.de"
  }'
```

Ohne gültige `AAD_*`-Werte in `local.settings.json` schlägt der eigentliche
Mailversand fehl (Status 502) – das ist an dieser Stelle normal. Die
Validierung (Status 400 bei fehlenden Feldern) kannst du schon jetzt testen.

## 2. Azure-Ressourcen anlegen

Namen (`stcateringaddin`, `func-catering-addin`) sind global eindeutig –
bei Bedarf anpassen.

```bash
az group create --name rg-catering-addin --location westeurope

az storage account create \
  --name stcateringaddin \
  --resource-group rg-catering-addin \
  --location westeurope \
  --sku Standard_LRS

az functionapp create \
  --resource-group rg-catering-addin \
  --consumption-plan-location westeurope \
  --runtime node \
  --runtime-version 20 \
  --functions-version 4 \
  --name func-catering-addin \
  --storage-account stcateringaddin \
  --os-type Linux
```

## 3. App-Einstellungen (Umgebungsvariablen) setzen

Erst nach Schritt 4 (App-Registrierung) habt ihr die echten Werte für
`AAD_CLIENT_ID`/`AAD_CLIENT_SECRET` – diesen Schritt könnt ihr auch danach
nachholen.

```bash
az functionapp config appsettings set \
  --name func-catering-addin \
  --resource-group rg-catering-addin \
  --settings \
    AAD_TENANT_ID="<eure-tenant-id>" \
    AAD_CLIENT_ID="<appId-aus-schritt-4>" \
    AAD_CLIENT_SECRET="<secret-aus-schritt-4>" \
    SENDER_MAILBOX="catering-relay@meinefirma.de"
```

## 4. App-Registrierung in Entra ID (für Mail.Send)

```bash
# App anlegen
az ad app create --display-name "Catering-Relay-Mailsender"
# -> Notiere die "appId" aus der Ausgabe

# Client Secret erzeugen
az ad app credential reset --id <appId> --years 2
# -> Notiere den Wert "password" aus der Ausgabe = AAD_CLIENT_SECRET

# Die genaue Rollen-ID für "Mail.Send" (Application) bei Microsoft Graph
# nachschlagen, statt sie zu erraten:
az ad sp show --id 00000003-0000-0000-c000-000000000000 \
  --query "appRoles[?value=='Mail.Send']"
# -> Notiere die zurückgegebene "id" = <mailSendRoleId>

# Berechtigung zuweisen und Admin-Consent erteilen (braucht Admin-Rechte)
az ad app permission add --id <appId> \
  --api 00000003-0000-0000-c000-000000000000 \
  --api-permissions <mailSendRoleId>=Role

az ad app permission admin-consent --id <appId>
```

**Wichtig:** `Mail.Send` als Application Permission erlaubt der App
grundsätzlich, **als jedes beliebige Postfach** zu senden. Das solltet ihr
mit Schritt 6 auf euer Relay-Postfach einschränken.

## 5. Code deployen

```bash
cd backend/azure-function-graph
func azure functionapp publish func-catering-addin
```

## 6. Zugriff der App auf genau ein Postfach einschränken

Voraussetzung: `catering-relay@meinefirma.de` existiert bereits als
(Shared-)Postfach.

```powershell
Connect-ExchangeOnline
New-ApplicationAccessPolicy -AppId <appId> `
  -PolicyScopeGroupId catering-relay@meinefirma.de `
  -AccessRight RestrictAccess `
  -Description "Nur Catering-Relay darf über diese App senden"

Test-ApplicationAccessPolicy -AppId <appId> -Identity catering-relay@meinefirma.de
```

## 7. Trigger-URL (inkl. Schlüssel) abrufen

```bash
az functionapp function keys list \
  --name func-catering-addin \
  --resource-group rg-catering-addin \
  --function-name CateringRequest
```

Alternativ im Azure-Portal: Function App → Functions → `CateringRequest` →
**„Get Function URL"** → kopieren. Das Ergebnis sieht so aus:

```
https://func-catering-addin.azurewebsites.net/api/CateringRequest?code=AbCdEf...
```

Diese **komplette** URL (inklusive `?code=…`) in
`src/taskpane/taskpane.js` bei `CATERING_ENDPOINT` eintragen.

## 8. End-to-End-Test

```bash
curl -X POST "https://func-catering-addin.azurewebsites.net/api/CateringRequest?code=<key>" \
  -H "Content-Type: application/json" \
  -d '{ ... gleiches Beispiel wie oben ... }'
```

Prüfen, ob die E-Mail bei `empfang@bruening-group.de` ankommt. Danach das
Add-in mit der aktualisierten `taskpane.js` neu hochladen bzw. bei
zentraler Bereitstellung im M365 Admin Center aktualisieren.

## Fehlersuche

- **Status 502 vom Trigger:** Meist fehlender/fehlerhafter Admin-Consent
  oder falsche `AAD_*`-Werte. Logs prüfen: `az functionapp log tail --name func-catering-addin --resource-group rg-catering-addin`
- **"Insufficient privileges" von Graph:** Admin-Consent (Schritt 4, letzter
  Befehl) wurde nicht erteilt, oder es wurde versehentlich eine Delegated-
  statt Application-Permission hinzugefügt.
- **E-Mail kommt nicht an, aber Status 200:** Application Access Policy
  (Schritt 6) prüfen – ggf. blockiert sie das Senden vom falschen Postfach.
