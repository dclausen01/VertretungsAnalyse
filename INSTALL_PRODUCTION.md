# Installation des Vertretungsanalyse Add-ins in Outlook Desktop

Diese Anleitung beschreibt, wie Sie das Vertretungsanalyse Add-in direkt in Outlook Desktop installieren können.

## Voraussetzungen

- Microsoft Outlook 2019 oder neuer (Version 16.0.13709 oder höher für optimale Kompatibilität)
- Ein OpenAI API-Schlüssel (erhältlich unter https://platform.openai.com/)

## Installation

### Methode 1: Direkte Installation aus der ZIP-Datei

1. Entpacken Sie die Datei `VertretungsAnalyse-Addin.zip` in einen Ordner Ihrer Wahl, z.B. `C:\VertretungsAnalyse-Addin`

2. Öffnen Sie Outlook

3. Gehen Sie zu **Datei** > **Optionen** > **Add-Ins**

4. Unten auf der Seite, wählen Sie **COM-Add-Ins** im Dropdown-Menü "Verwalten" und klicken Sie auf **Gehe zu...**

5. Klicken Sie auf **Add-In hinzufügen** und navigieren Sie zum entpackten Ordner

6. Wählen Sie die Datei `manifest.xml` aus und klicken Sie auf **Öffnen**

7. Stellen Sie sicher, dass das Add-In in der Liste erscheint und aktiviert ist

8. Starten Sie Outlook neu

### Methode 2: Installation über einen Netzwerkfreigabe (für Organisationen)

1. Kopieren Sie den Inhalt der ZIP-Datei auf einen Netzwerkfreigabe, die für alle Benutzer zugänglich ist

2. Stellen Sie sicher, dass die Benutzer Leserechte für alle Dateien haben

3. Folgen Sie den Schritten 2-8 aus Methode 1, aber wählen Sie die `manifest.xml` Datei von der Netzwerkfreigabe aus

### Methode 3: Installation über einen Webserver

1. Laden Sie den Inhalt der ZIP-Datei auf einen Webserver hoch

2. Bearbeiten Sie die `manifest.xml` Datei und ändern Sie alle URLs von `https://localhost:3000/` zu Ihrer Webserver-URL, z.B. `https://ihr-server.de/vertretungsanalyse/`

3. Speichern Sie die bearbeitete `manifest.xml` Datei lokal

4. Folgen Sie den Schritten 2-8 aus Methode 1, aber wählen Sie die bearbeitete `manifest.xml` Datei aus

## Erste Verwendung

1. Öffnen Sie eine E-Mail in Outlook

2. Sie sollten einen neuen Button "Vertretungsanalyse" in der Menüleiste sehen

3. Klicken Sie auf diesen Button

4. Bei der ersten Verwendung werden Sie aufgefordert, Ihren OpenAI API-Schlüssel einzugeben

5. Geben Sie Ihren API-Schlüssel ein und klicken Sie auf OK

6. Das Add-in analysiert nun die E-Mail und zeigt das Ergebnis an

## Fehlerbehebung

### Das Add-in erscheint nicht in Outlook

- Stellen Sie sicher, dass Sie Outlook nach der Installation neu gestartet haben
- Überprüfen Sie, ob das Add-in in der Liste der installierten Add-ins erscheint
- Versuchen Sie, das Add-in manuell zu aktivieren unter **Datei** > **Optionen** > **Add-Ins**

### Fehler bei der API-Anfrage

- Überprüfen Sie, ob Ihr OpenAI API-Schlüssel gültig ist
- Stellen Sie sicher, dass Sie über eine aktive Internetverbindung verfügen
- Überprüfen Sie, ob Ihr API-Kontingent nicht überschritten wurde

### Sonstige Probleme

Bei weiteren Problemen wenden Sie sich bitte an den Support oder erstellen Sie ein Issue im GitHub-Repository.

## Deinstallation

1. Öffnen Sie Outlook
2. Gehen Sie zu **Datei** > **Optionen** > **Add-Ins**
3. Finden Sie das Vertretungsanalyse Add-in in der Liste
4. Wählen Sie es aus und klicken Sie auf **Entfernen**
5. Starten Sie Outlook neu
