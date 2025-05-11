# Vertretungsanalyse Outlook Add-in

Ein Outlook Add-in zur Analyse von Lehrermails und Erstellung strukturierter Regelungsvorschläge für die Schulsoftware UNTIS.

## Funktionen

1. **Analyse-Button in jeder E-Mail**: Das Add-in fügt einen "Vertretungsanalyse"-Button in jede E-Mail ein.
2. **KI-gestützte Analyse**: Bei Klick auf den Button wird die gesamte E-Mail (inklusive Header) mit einem OpenAI-Modell analysiert.
3. **Formatierte Ausgabe**: Das Analyseergebnis wird am Anfang der E-Mail in dunkelblauer, fetter Arial-Schrift angezeigt.

## Anforderungen

- Outlook 2019 oder neuer
- Internetverbindung für die OpenAI API-Anfragen
- OpenAI API-Schlüssel

## Installation

### Entwicklungsumgebung

1. Klonen Sie das Repository:
   ```
   git clone https://github.com/yourusername/vertretungsanalyse-addin.git
   cd vertretungsanalyse-addin
   ```

2. Installieren Sie die Abhängigkeiten:
   ```
   npm install
   ```

3. Starten Sie den Entwicklungsserver:
   ```
   npm start
   ```

### Installation in Outlook

#### Während der Entwicklung

1. Öffnen Sie Outlook 2019.
2. Gehen Sie zu "Datei" > "Optionen" > "Add-Ins".
3. Klicken Sie auf "Verwalten" (COM-Add-Ins) und dann auf "Durchsuchen".
4. Navigieren Sie zum Manifest-Datei (manifest.xml) im Projektverzeichnis und wählen Sie diese aus.
5. Folgen Sie den Anweisungen auf dem Bildschirm, um die Installation abzuschließen.

#### Für Produktionsumgebungen

Für die Bereitstellung in einer Produktionsumgebung sollten Sie das Add-in paketieren und über einen der folgenden Wege verteilen:

- Direktes Teilen der Manifest-Datei
- Bereitstellung über einen Webserver
- Veröffentlichung im Microsoft AppSource

## Verwendung

1. Öffnen Sie eine E-Mail in Outlook.
2. Klicken Sie auf den "Vertretungsanalyse"-Button in der Menüleiste.
3. Bei der ersten Verwendung werden Sie aufgefordert, Ihren OpenAI API-Schlüssel einzugeben.
4. Das Add-in analysiert die E-Mail und zeigt das Ergebnis im Aufgabenbereich an.
5. In einer vollständigen Implementierung würde das Analyseergebnis auch am Anfang der E-Mail eingefügt werden.

## Entwicklung

### Projektstruktur

```
vertretungsanalyse-addin/
├── assets/                  # Bilder und Icons
├── src/
│   ├── commands/            # Befehlsimplementierungen
│   ├── helpers/             # Hilfsfunktionen
│   ├── services/            # Dienste (z.B. OpenAI API)
│   └── taskpane/            # Hauptlogik und UI
├── manifest.xml             # Add-in-Manifest
├── package.json             # Projektkonfiguration
└── webpack.config.js        # Build-Konfiguration
```

### Anpassung

- Die OpenAI-Prompt kann in der Datei `src/services/openai-service.js` angepasst werden.
- Das UI-Design kann in der Datei `src/taskpane/taskpane.html` und den zugehörigen CSS-Stilen angepasst werden.
- Die Logik für die Analyse und Anzeige der Ergebnisse kann in der Datei `src/taskpane/taskpane.js` angepasst werden.

## Sicherheitshinweise

- Der OpenAI API-Schlüssel wird lokal im Browser-Speicher gespeichert. In einer Produktionsumgebung sollte eine sicherere Methode zur Speicherung verwendet werden.
- Die API-Anfragen sollten über einen sicheren Backend-Dienst erfolgen, um den API-Schlüssel zu schützen und Rate-Limiting, Fehlerbehandlung usw. zu handhaben.

## Lizenz

MIT
