# Vertretungsanalyse Add-in

Ein Outlook Add-in zur Analyse von Lehrermails und Erstellung strukturierter Regelungsvorschläge für UNTIS.

## Funktionen

- Analysiert E-Mails von Lehrkräften
- Extrahiert relevante Informationen wie Lehrkraft-Kürzel, Typ der Anfrage, Datum und Zeitraum
- Erstellt strukturierte Regelungsvorschläge für UNTIS
- Fügt Wochentage zu Datumsangaben hinzu
- Prüft auf Berechtigungen für Küchenräume

## Entwicklung

### Voraussetzungen

- Node.js (empfohlen: neueste LTS-Version)
- npm (wird mit Node.js installiert)
- Git

### Installation

1. Repository klonen:
   ```
   git clone https://github.com/dclausen01/VertretungsAnalyse.git
   cd VertretungsAnalyse
   ```

2. Abhängigkeiten installieren:
   ```
   npm install
   ```

3. Entwicklungsserver starten:
   ```
   npm start
   ```

### Entwicklungsmodus

Im Entwicklungsmodus wird das Add-in lokal auf https://localhost:3000 gehostet. Der Befehl `npm start` startet den Entwicklungsserver und öffnet das Add-in in Outlook.

## Deployment auf GitHub Pages

### Vorbereitung

1. Stellen Sie sicher, dass GitHub Pages für Ihr Repository aktiviert ist:
   - Gehen Sie zu Ihrem Repository auf GitHub
   - Klicken Sie auf "Settings" > "Pages"
   - Wählen Sie unter "Source" die Option "Deploy from a branch"
   - Wählen Sie unter "Branch" den Branch "main" und den Ordner "/docs"
   - Klicken Sie auf "Save"

### Deployment

1. Führen Sie den Deployment-Befehl aus:
   ```
   npm run deploy
   ```

2. Commit und Push der generierten Dateien:
   ```
   git add docs
   git commit -m "Deploy to GitHub Pages"
   git push origin main
   ```

3. Überprüfen Sie die Deployment-URL:
   - Ihre Add-in-Dateien sind nun unter `https://dclausen01.github.io/VertretungsAnalyse/` verfügbar
   - Das Manifest für die Produktion ist unter `https://dclausen01.github.io/VertretungsAnalyse/manifest.xml` verfügbar

## Installation des Add-ins in Outlook

### Entwicklungsversion (lokal)

1. Starten Sie den Entwicklungsserver mit `npm start`
2. Outlook wird automatisch geöffnet und das Add-in wird geladen

### Produktionsversion (GitHub Pages)

1. Öffnen Sie Outlook im Web oder Desktop
2. Klicken Sie auf "..." oder "Get Add-ins" in der Menüleiste
3. Wählen Sie "My Add-ins" > "Add a custom add-in" > "Add from URL"
4. Geben Sie die URL zum Manifest ein: `https://dclausen01.github.io/VertretungsAnalyse/manifest.xml`
5. Klicken Sie auf "OK" oder "Add"

## Verwendung

1. Öffnen Sie eine E-Mail in Outlook
2. Klicken Sie auf die Schaltfläche "Vertretungsanalyse" in der Menüleiste
3. Klicken Sie im Taskpane auf "E-Mail analysieren"
4. Die Analyse wird angezeigt und kann für UNTIS verwendet werden

## Fehlerbehebung

- **API-Schlüssel-Fehler**: Stellen Sie sicher, dass Sie einen gültigen OpenAI API-Schlüssel eingegeben haben
- **Netzwerkfehler**: Überprüfen Sie Ihre Internetverbindung
- **Add-in wird nicht geladen**: Stellen Sie sicher, dass die URL zum Manifest korrekt ist

## Lizenz

MIT
