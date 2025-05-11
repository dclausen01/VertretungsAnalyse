# Installation Guide für das Vertretungsanalyse Outlook Add-in

Diese Anleitung führt Sie durch die Installation und Einrichtung des Vertretungsanalyse Add-ins für Outlook 2019.

## Voraussetzungen

Bevor Sie beginnen, stellen Sie sicher, dass Sie Folgendes haben:

1. Microsoft Outlook 2019 installiert
2. Einen OpenAI API-Schlüssel (erhältlich unter https://platform.openai.com/)
3. Administratorrechte auf Ihrem Computer (für die Installation)

## Installation für Entwickler

### 1. Projekt klonen und Abhängigkeiten installieren

```bash
# Repository klonen
git clone https://github.com/yourusername/vertretungsanalyse-addin.git

# In das Projektverzeichnis wechseln
cd vertretungsanalyse-addin

# Abhängigkeiten installieren
npm install
```

### 2. Entwicklungsserver starten

```bash
npm start
```

Dies startet einen lokalen Webserver auf https://localhost:3000 und öffnet die Manifest-Datei in Outlook.

### 3. Add-in in Outlook sideloaden

1. Öffnen Sie Outlook 2019
2. Gehen Sie zu **Datei** > **Optionen** > **Add-Ins**
3. Unten auf der Seite, wählen Sie **COM-Add-Ins** im Dropdown-Menü und klicken Sie auf **Gehe zu...**
4. Klicken Sie auf **Add-In hinzufügen** und navigieren Sie zur `manifest.xml` Datei im Projektverzeichnis
5. Wählen Sie die Datei aus und klicken Sie auf **Öffnen**
6. Stellen Sie sicher, dass das Add-In in der Liste erscheint und aktiviert ist

## Installation für Endbenutzer

### Option 1: Direkte Installation aus einer Datei

1. Laden Sie die Manifest-Datei (`manifest.xml`) herunter
2. Öffnen Sie Outlook 2019
3. Gehen Sie zu **Datei** > **Optionen** > **Add-Ins**
4. Unten auf der Seite, wählen Sie **COM-Add-Ins** im Dropdown-Menü und klicken Sie auf **Gehe zu...**
5. Klicken Sie auf **Add-In hinzufügen** und navigieren Sie zur heruntergeladenen `manifest.xml` Datei
6. Wählen Sie die Datei aus und klicken Sie auf **Öffnen**
7. Stellen Sie sicher, dass das Add-In in der Liste erscheint und aktiviert ist

### Option 2: Installation über einen Administrator (für Organisationen)

In einer Organisationsumgebung kann ein Administrator das Add-in zentral bereitstellen:

1. Der Administrator lädt die Manifest-Datei auf einen SharePoint-Dokumentenbibliothek oder einen anderen zentralen Speicherort hoch
2. Der Administrator konfiguriert die Gruppenrichtlinie, um das Add-in für alle Benutzer bereitzustellen
3. Das Add-in wird automatisch in Outlook für alle Benutzer installiert

## Erste Verwendung

1. Öffnen Sie eine E-Mail in Outlook
2. Sie sollten einen neuen Button "Vertretungsanalyse" in der Menüleiste sehen
3. Klicken Sie auf diesen Button
4. Bei der ersten Verwendung werden Sie aufgefordert, Ihren OpenAI API-Schlüssel einzugeben
5. Geben Sie Ihren API-Schlüssel ein und klicken Sie auf OK
6. Das Add-in analysiert nun die E-Mail und zeigt das Ergebnis an

## Fehlerbehebung

### Das Add-in erscheint nicht in Outlook

- Stellen Sie sicher, dass der Entwicklungsserver läuft (wenn Sie die Entwicklerinstallation verwenden)
- Überprüfen Sie, ob das Add-in in der Liste der installierten Add-ins erscheint
- Starten Sie Outlook neu

### Fehler bei der API-Anfrage

- Überprüfen Sie, ob Ihr OpenAI API-Schlüssel gültig ist
- Stellen Sie sicher, dass Sie über eine aktive Internetverbindung verfügen
- Überprüfen Sie, ob Ihr API-Kontingent nicht überschritten wurde

### Sonstige Probleme

Bei weiteren Problemen wenden Sie sich bitte an den Support oder erstellen Sie ein Issue im GitHub-Repository.

## Deinstallation

1. Öffnen Sie Outlook 2019
2. Gehen Sie zu **Datei** > **Optionen** > **Add-Ins**
3. Finden Sie das Vertretungsanalyse Add-in in der Liste
4. Wählen Sie es aus und klicken Sie auf **Entfernen**
5. Starten Sie Outlook neu
