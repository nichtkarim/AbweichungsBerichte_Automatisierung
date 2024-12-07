# README für TuPack Version 0.1 - Abweichungsberichte extrahieren

## Übersicht
**TuPack Version 0.1** ist eine Python-basierte Anwendung, die es ermöglicht, Daten aus mehreren Excel-Dateien zu extrahieren und in eine CSV-Datei zu exportieren. Die Anwendung konzentriert sich auf die Extraktion von Abweichungsberichten, die spezifische Informationen zu Bestellungen, Aufträgen und Fehlern enthalten. Diese Informationen werden aus zwei Arbeitsblättern in den Excel-Dateien abgerufen und anschließend in einer strukturierten CSV-Datei gespeichert.

## Funktionen
- **Auswahl von Excel-Dateien**: Ermöglicht die Auswahl mehrerer Excel-Dateien zur Verarbeitung.
- **CSV-Dateiausgabe**: Exportiert extrahierte Daten in eine CSV-Datei mit einem automatisch generierten Dateinamen basierend auf dem aktuellen Datum.
- **Logging**: Alle Verarbeitungsschritte werden in einer Log-Datei dokumentiert, um Fehler und Fortschritte nachzuvollziehen.
- **Fortschrittsanzeige**: Eine grafische Fortschrittsanzeige zeigt an, wie weit der Exportprozess fortgeschritten ist.
- **Benutzerfreundliche GUI**: Die grafische Benutzeroberfläche (GUI) macht die Anwendung einfach zu bedienen und ermöglicht die Auswahl von Dateien und Speicherorten.

## Voraussetzungen
Die Anwendung erfordert die folgenden Python-Bibliotheken:
- `csv`: Zum Schreiben der CSV-Datei.
- `openpyxl`: Zum Arbeiten mit Excel-Dateien (.xlsx und .xlsm).
- `tkinter`: Für die GUI.
- `warnings`: Zum Ignorieren von Warnmeldungen.
- `logging`: Zum Erstellen von Log-Dateien.
- `os`: Für Dateisystemoperationen.
- `datetime`: Zum Arbeiten mit Datums- und Uhrzeitangaben.

Stellen Sie sicher, dass alle erforderlichen Bibliotheken installiert sind. Sie können sie mit `pip` installieren:
```bash
pip install openpyxl
```

## Funktionsweise
1. **Excel-Dateien auswählen**:
   - Klicken Sie auf "Excel-Dateien auswählen", um die gewünschten Excel-Dateien auszuwählen. Diese Dateien müssen im `.xlsx`- oder `.xlsm`-Format vorliegen.
   - Die Anwendung ermöglicht die Auswahl mehrerer Dateien gleichzeitig.

2. **Speicherort für die CSV-Datei auswählen**:
   - Klicken Sie auf "Speicherort der CSV-Datei auswählen", um den Ordner auszuwählen, in dem die CSV-Datei gespeichert werden soll.
   - Die CSV-Datei wird automatisch mit dem aktuellen Datum als Dateinamen gespeichert.

3. **Daten verarbeiten**:
   - Klicken Sie auf "Verarbeitung starten", um die Extraktion der Daten zu starten. Der Fortschritt wird auf einem Balken angezeigt, und das Ergebnis wird als CSV-Datei gespeichert.
   - Die Anwendung extrahiert Daten aus zwei Arbeitsblättern und speichert diese in der CSV-Datei.

4. **Logging**:
   - Während der Verarbeitung werden alle Schritte in einer Log-Datei dokumentiert. Diese Datei wird im selben Verzeichnis wie die CSV-Datei gespeichert und enthält alle Fehler und Verarbeitungsdetails.

## Dateistruktur
- **CSV-Datei**: Enthält die extrahierten Daten im CSV-Format. Jede Zeile repräsentiert einen Datensatz, und jede Spalte enthält eine spezifische Information aus der Excel-Datei.
- **Log-Datei**: Eine Log-Datei, die alle Schritte der Verarbeitung dokumentiert. Diese Datei hilft, den Verlauf und eventuelle Fehler während des Prozesses nachzuvollziehen.

## Beispiel für die CSV-Datei
Die CSV-Datei enthält folgende Spalten:
- Datum
- MeldungsNr
- UHPT-Bestellnummer und Pos.
- UHPT-Auftragsnummer
- Kaufm. Sachbearbeiter UHPT
- Techn. Sachbearbeiter UHPT
- UHPT-Zeichnungsnummer und Revisionstand
- Benennung
- Name und Anschrift Lieferant
- Kontaktdaten Lieferant
- Auftragsnummer Lieferant
- Bestellmenge
- Betroffene Menge (n.i.O.)
- Fehlerbeschreibung
- Anlagenverzeichnis
- Vorgeschlagene Maßnahme / Sofortmaßnahme
- Fehlerursache
- Korrekturmaßnahme zur Fehlerursache
- Sonderfreigabe / Tolerierung
- Nacharbeit (Lieferant)
- Nacharbeit (UHDE)
- Ausschuss (Verschrottung Lieferant)
- Ausschuss (Verschrottung Uhde)
- Auftrag reduzieren
- Neufertigung UHPT
- Nacharbeitskosten zu Lasten UHPT
- Neufertigung Lieferant (Materialbestellung nicht erforderlich)
- Neufertigung Lieferant (Materialbestellung erforderlich)
- Nacharbeitskosten zu Lasten Lieferant

## Fehlerbehandlung
- **Keine Excel-Dateien ausgewählt**: Es wird eine Fehlermeldung angezeigt, wenn keine Dateien ausgewählt wurden.
- **Kein Speicherort für CSV ausgewählt**: Es wird eine Fehlermeldung angezeigt, wenn der Speicherort für die CSV-Datei nicht angegeben wurde.
- **Verarbeitungsfehler**: Tritt ein Fehler bei der Verarbeitung einer Excel-Datei auf, wird der Fehler in der Log-Datei dokumentiert und der Prozess für diese Datei übersprungen.

## Installation und Ausführung
1. **Python installieren**: Stellen Sie sicher, dass Python auf Ihrem Computer installiert ist. Sie können es von der offiziellen Website herunterladen: [python.org](https://www.python.org/).
2. **Code ausführen**: Führen Sie das Python-Skript aus, indem Sie den folgenden Befehl im Terminal verwenden:
   ```bash
   python tuPack.py
   ```
   Dies startet die GUI, und Sie können mit der Auswahl von Dateien und dem Export der Daten beginnen.

## Credits
- Entwickelt von **Karim Abdulhadi**.
- Version 0.1 Beta.

## Lizenz
Dieses Projekt ist unter der MIT-Lizenz lizenziert.

---

Falls es weitere Fragen oder Probleme gibt, wenden Sie sich bitte an den Entwickler oder erstellen Sie ein Issue im GitHub-Repository (falls vorhanden).