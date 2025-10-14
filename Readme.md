# Excel_Manager.py

## Übersicht

`Excel_Manager.py` enthält die Klasse **ExcelManager**, die verschiedene Funktionen zur Auswertung und Verarbeitung von Excel-Dateien bereitstellt. Das Hauptziel ist die Analyse und Aggregation von Integrationsszenarien aus einer Excel-Tabelle.v

## Hauptfunktionen

### 1. Daten einlesen und sortieren

- Liest alle Zeilen aus dem Arbeitsblatt `Full Evaluation Results`.
- Sortiert die Zeilen nach der ersten Spalte ("Integration Scenario"), sodass alle Einträge zu einem Szenario gruppiert sind.

### 2. Szenario-basierte Auswertung

- Für jedes Integrationsszenario werden verschiedene Kennzahlen berechnet und in Tabellen gespeichert:
  - Anzahl Empfänger
  - Mapping-Typ
  - UDF-Nutzung
  - Quality of Service
  - Anzahl Schnittstellen (FTP, SFTP, FTPS, UDF)
  - TShirt Size und Aufwandsschätzungen

### 3. Ergebnisse schreiben

- Die berechneten Werte werden in ein neues Arbeitsblatt geschrieben, sodass die Auswertung übersichtlich dargestellt wird.

## Beispiel: Nutzung der Klasse

```python
from Excel_Manager import ExcelManager

manager = ExcelManager(workbook)
manager.fill_sheet(sheet123)
```

## Wichtige Tabellen und Begriffe

- **table_mapping**: Zuordnung von Szenario zu Mapping-Typ.
- **table_udf**: Nutzung von UDF pro Szenario.
- **table_receivers_count**: Anzahl Empfänger pro Szenario.
- **table_ftp_count**: Anzahl FTP-Schnittstellen pro Szenario.
- **table_tshirt_size**: Aufwandsschätzung (TShirt Size) pro Szenario.

## Hinweise

- Die Sortierung nach "Integration Scenario" stellt sicher, dass alle Einträge zu einem Szenario korrekt zusammengefasst werden.
- Die Klasse ist für die Verarbeitung großer Excel-Dateien mit vielen Szenarien ausgelegt.

## Lizenz

Dieses Projekt ist nur für interne Zwecke bestimmt.
