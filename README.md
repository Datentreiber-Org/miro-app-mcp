# Miro → Lovable MCP Button Only

Reduzierte Version des ursprünglichen Repos. Diese Version enthält nur noch den Miro-Board-Button, der Lovable mit einem Initial-Prompt öffnet. Der Prompt weist Lovable an, den in Lovable verbundenen Miro-MCP-Server zu verwenden und nur den Inhalt eines bestimmten Miro-Frames auszuwerten.

## Was entfernt wurde

- Keine PDF-Analyse.
- Kein OpenAI/GPT-Backend.
- Keine Vercel-Serverless-API-Routes.
- Keine Miro-REST-API-Tokens und keine Environment Variables.
- Kein `app.html`, kein `/api/analyze-selected-pdf`, kein `/api/table-to-stickies-mcp`.

## Was übrig bleibt

- `index.html`: eine statische Miro-Web-SDK-App.
- `manifest.example.yml`: Beispielwerte für die Miro-App-Konfiguration.
- `package.json`: nur für lokalen statischen Testserver.

Die App sucht auf dem aktuellen Board einen Frame mit dem Titel `lovable`, legt darin oben rechts einen grünen Shape-Button an und hängt an diesen Shape einen Lovable-Link. Wird der Shape ausgewählt, öffnet die App Lovable mit einem Initial-Prompt. Falls ein Browser Popup blockiert, kann der Link über das Pfeil-Icon am Shape geöffnet werden.

## Wird weiterhin Vercel benötigt?

Für die Funktion selbst: nein. Es gibt kein Backend mehr.

Für eine produktive Miro-App brauchst du aber weiterhin eine öffentlich erreichbare HTTPS-URL, unter der `index.html` liegt. Vercel ist dafür weiterhin eine einfache Option, aber nicht zwingend. Alternativen sind Netlify, Cloudflare Pages, GitHub Pages oder jeder andere statische HTTPS-Host.

Für lokale Entwicklung kann Miro auch `http://localhost:3000/` verwenden. Für produktive App-URLs sollte die URL HTTPS sein.

## Voraussetzungen

1. Miro Developer App / private Miro-App mit Web SDK v2.
2. Scopes in Miro: `boards:read` und `boards:write`.
3. Ein Miro-Board mit einem Frame namens `lovable`.
4. Lovable-Workspace, in dem Miro als MCP-Integration verbunden ist.
5. Das Miro-Board muss in dem Miro-Team liegen, das du bei der Lovable/Miro-MCP-OAuth-Verbindung autorisiert hast.

## Lokaler Test

```bash
npm install
npm run dev
```

Danach läuft die App lokal unter:

```text
http://localhost:3000/
```

Diese URL kannst du in der Miro-App-Konfiguration als SDK/App-URL für lokale Tests eintragen.

## Deployment mit Vercel

1. Dieses Repo zu GitHub pushen.
2. In Vercel ein neues Projekt aus dem Repo importieren.
3. Framework Preset: `Other` oder automatische statische Erkennung verwenden.
4. Es sind keine Environment Variables nötig.
5. Es ist kein Backend und kein Build-Step nötig.
6. Nach dem Deployment die URL notieren, z. B.:

```text
https://dein-projekt.vercel.app/
```

Diese URL trägst du in Miro als SDK/App-URL ein.

## Miro-App einrichten

Nutze `manifest.example.yml` als Vorlage oder trage die Werte in der Miro Developer UI ein:

```yaml
appName: Miro Lovable MCP Button
sdkVersion: SDK_V2
sdkUri: https://dein-projekt.vercel.app/
scopes:
  - boards:read
  - boards:write
boardPicker:
  allowedDomains:
    - dein-projekt.vercel.app
    - localhost
```

Danach die App in deinem Miro-Team installieren und auf dem Board öffnen.

## Lovable/Miro MCP einrichten

In Lovable:

1. Profil / Settings öffnen.
2. Integrations öffnen.
3. Unter MCP / Chat Connectors `Miro` suchen.
4. `Set up` bzw. `Connect` klicken.
5. Miro OAuth abschließen und das richtige Miro-Team autorisieren.

Erst danach kann Lovable den Board-Link aus dem Prompt über Miro MCP auslesen.

## Nutzung im Board

1. Im Miro-Board einen Frame mit exakt diesem Titel anlegen oder umbenennen:

```text
lovable
```

Groß-/Kleinschreibung ist egal.

2. In diesen Frame die Inhalte legen, die Lovable verwenden soll.

Mit dem aktuellen Default-Prompt erwartet Lovable insbesondere:

- genau ein Miro Doc als `MASTER SPEC`,
- eine Miro Table mit dem Namen `Voted Criterias`,
- optional ein zweites Miro Doc namens `Scoring Rulebook`.

3. Die Miro-App öffnen bzw. das App-Icon klicken. Der Button `🚀 Lovable: Generate App` wird im Frame erzeugt oder aktualisiert.
4. Den Button-Shape auswählen. Lovable öffnet sich mit dem vorbereiteten Prompt.
5. Falls kein Tab geöffnet wird: das Pfeil-Icon oben rechts am Shape klicken.

## Prompt oder Frame ändern

Alle wichtigen Einstellungen stehen oben in `index.html`:

```js
const TARGET_FRAME_TITLE = "lovable";
const BUTTON_LABEL_TEXT = "🚀 Lovable: Generate App";
const LOVABLE_BASE_URL = "https://lovable.dev/?autosubmit=true#prompt=";
```

Den eigentlichen Initial-Prompt änderst du in:

```js
function buildLovablePrompt(boardId) {
  // ...
}
```

Nach Änderungen neu deployen und die Miro-App-URL unverändert lassen, sofern der Host gleich bleibt.

## Troubleshooting

### Button wird nicht erstellt

Prüfe, ob der Frame wirklich `lovable` heißt. Dann das App-Icon in Miro erneut anklicken. Außerdem muss die App `boards:read` und `boards:write` haben.

### Lovable öffnet, kann aber Miro nicht lesen

Prüfe in Lovable, ob Miro MCP verbunden ist und ob beim OAuth-Flow das richtige Miro-Team autorisiert wurde. Das Board muss für den autorisierten Account und das autorisierte Team zugänglich sein.

### Popup wird geblockt

Das ist erwartbar möglich, weil der Launch über ein Board-Selection-Event passiert. Der Link wird zusätzlich am Shape über `linkedTo` gesetzt. Nutze in diesem Fall das Pfeil-Icon am Shape.

### Der falsche Inhalt wird verwendet

Der Default-Prompt beschränkt Lovable auf den Frame `lovable`. Stelle sicher, dass die relevanten Docs/Tables wirklich innerhalb dieses Frames liegen. Passe bei Bedarf `TARGET_FRAME_TITLE` oder `buildLovablePrompt()` an.
