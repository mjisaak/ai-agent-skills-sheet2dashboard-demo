# Sheet-to-Dashboard Agent Demo

Agent Skills Demo: Eine Excel-Tabelle bereinigen („sanitize“) und daraus einen sauberen, palette-basierten, interaktiven Report erzeugen.

## Sample Data

Die Excel-Datei `test-data.xlsx` enthält **100 Testdatensätze** mit synthetisch generierten Personendaten:

- **Name** (Vor- und Nachname)
- **Stadt**
- **Beruf**
- **Abteilung**
- **Teilzeit** (Ja/Nein)
- **Alter**
- **Monatlicher Umsatz** für die **letzten drei vollen Kalenderjahre** (je Monat eine Spalte)

Die Daten sind **zufällig**, aber **plausibel und konsistent** aufgebaut, z. B.:

### Prompt zur Datengenerierung

Die Datei wurde mit folgendem Prompt generiert:

```prompt
Erstelle eine Excel-Datei (.xlsx) mit 100 Datensätzen (eine Zeile pro Person) und folgenden Spalten:

Stammdaten:
- Name (Vor- und Nachname)
- Stadt (realistische Städte in Deutschland/Österreich/Schweiz)
- Beruf (realistisch, passend zur Abteilung)
- Abteilung (z. B. Vertrieb, Marketing, IT, HR, Finance, Operations, Customer Support, Produkt, Einkauf, Recht)
- Teilzeit (Ja/Nein)
- Alter (Ganzzahl, z. B. 18–65)

Umsatz (monatlich):
- Für jeden Monat der letzten 3 vollen Kalenderjahre jeweils eine Spalte im Format: Umsatz_YYYY-MM
  Beispiel: Umsatz_2023-01, Umsatz_2023-02, …, Umsatz_2025-12
- Wertebereich je Monat: 0 bis 12.000 EUR (Ganzzahl oder mit 2 Dezimalstellen)

Datenanforderungen:
- Namen sollen zufällig generiert sein, aber realistisch klingen (keine Fantasie-Strings).
- Alle Werte sollen zufällig, aber plausibel sein:
  - Bei Teilzeit = Ja ist der Umsatz tendenziell niedriger als bei Vollzeit.
  - Beruf und Abteilung müssen sinnvoll zusammenpassen (z. B. „Softwareentwickler“ → IT, „Account Manager“ → Vertrieb).
  - Optional: leichte Saisonalität im Umsatz (z. B. Q4 etwas höher), aber keine Pflicht.
- Keine leeren Zellen (Umsatz darf 0 sein).
- Erste Zeile enthält Spaltenüberschriften, Datei ist sofort nutzbar.

Output:
- Stelle die Datei als Download bereit und nenne sie: test-data.xlsx

```

## "Sheet-to-Dashboard" Skill

Der "Sheet-to-Dashboard" Skill bereinigt die Daten, analysiert sie und erstellt daraus einen interaktiven Report mit Diagrammen, KPIs und Filtermöglichkeiten.

Bereinigungsschritte:

- Vor- und Nachnamen in getrennte Spalten aufteilen
- Spalte Bundesland einführen (basierend auf Stadt)
- Sortierung nach Abteilung und Beruf

### Report-Elemente (Dashboard)

- **KPIs (Executive Row):**
  - Durchschnittsalter (optional: Median)
  - Gesamtumsatz (für ausgewählten Zeitraum)
  - Umsatz nach Abteilung (inkl. Top-Abteilung + Anteil)
  - Teilzeitquote + Umsatzvergleich Teilzeit vs. Vollzeit
- **Diagramme (Storyline: Überblick → Treiber → Zeit):**
  - Umsatzentwicklung über die Zeit (mit Zeitraum-Selector)
  - Umsatz nach Abteilung (palette-basiert, konsistente Farben je Abteilung)
  - Top 10 Berufe nach Umsatz (Tooltip: Headcount + Ø Umsatz)
  - Umsatzverteilung (Histogramm/Boxplot für Streuung & Ausreißer)
- **Filter & Interaktion:**
  - Filter: Stadt, Bundesland, Abteilung, Beruf, Teilzeit/Vollzeit
  - Altersgruppen: Schieberegler (z. B. 25–45) mit Live-Update
  - Interaktive Elemente: Dropdowns, Zeitraum-Picker, Cross-Filtering (Klick auf Chart-Element filtert den Rest)
  - Optional: Detailtabelle mit Suche/Sortierung als „Drill-down“

Der Skill wurde mit folgenden Prompt generiert:

```prompt
Erstelle einen „Sheet-to-Dashboard“-Skill. Deine Aufgabe: Nimm eine Excel-Datei als Input und liefere (1) eine bereinigte/standardisierte Excel-Datei und (2) einen polierten, palette-basierten, interaktiven HTML-Report mit KPIs, Charts und Filtern.

## Input
- Eine Excel-Datei (.xlsx) mit einer Tabelle (ein Sheet reicht).
- Erwartete Spalten (mindestens):
  - Name (Vor- und Nachname in einer Zelle ODER bereits Vorname/Nachname)
  - Stadt
  - Beruf
  - Abteilung
  - Teilzeit (Ja/Nein)
  - Alter
  - Monatsumsatz-Spalten im Pattern: Umsatz_YYYY-MM (z. B. Umsatz_2023-01 …)

## Output (Artefakte)
1) `sanitized-data.xlsx`
   - Bereinigte, validierte und standardisierte Version der Daten
2) `dashboard.html`
   - Interaktiver Report (ein File), modern, „executive-ready“, mit konsistenter Farbpalette (Farben pro Abteilung)

## Ziele
- Maximale Plausibilität, konsistentes Schema, keine kaputten Datentypen.
- Dashboard soll in einer Demo beeindrucken: klare KPIs, gute Visual Storyline, flüssige Interaktion.

## Verarbeitungsschritte (Sanitization & Enrichment)
1. Schema-Check:
   - Prüfe erwartete Spalten. Wenn etwas fehlt, liefere eine klare Fehlermeldung + Vorschlag zur Behebung.
2. Datentypen normalisieren:
   - Alter → Integer
   - Teilzeit → Enum {Ja, Nein}
   - Umsatz_* → Numeric (float oder int), Werte < 0 auf 0 setzen, leere Zellen als 0 interpretieren.
3. Name normalisieren:
   - Falls nur `Name` existiert: splitte in `Vorname` und `Nachname` (Trim, Mehrfachspaces, Doppelnamen robust behandeln).
   - Falls `Vorname`/`Nachname` existieren: stelle sicher, dass beide sauber getrimmt sind.
4. Stadt → Bundesland/Region:
   - Ergänze Spalte `Bundesland` (oder `Region`) basierend auf Stadt (DE/AT/CH). Nutze ein Mapping; wenn unbekannt, setze „Unbekannt“.
5. Umsatz-Spalten harmonisieren:
   - Erkenne alle Spalten `Umsatz_YYYY-MM`, sortiere sie chronologisch.
   - Erzeuge zusätzlich folgende Spalten/Features:
     - `Umsatz_Gesamt` (Summe über alle Monate)
     - Optional: `Umsatz_Ø_Monat` (Durchschnitt pro Monat)
6. Sortierung & Layout:
   - Standard-Spaltenreihenfolge: Vorname, Nachname, Stadt, Bundesland, Abteilung, Beruf, Teilzeit, Alter, Umsatz_* (chronologisch), Umsatz_Gesamt, Umsatz_Ø_Monat
   - Sortiere Zeilen nach Abteilung, dann Beruf, dann Nachname.

## Optional (für bessere Dashboard-Fähigkeit)
- Erzeuge zusätzlich ein „Long Format“ (tidy):
  - Spalten: Vorname, Nachname, Stadt, Bundesland, Abteilung, Beruf, Teilzeit, Alter, Datum (YYYY-MM), Umsatz
  - Nutze dieses Format primär für Zeitreihencharts und Filter.
  - (Du darfst dieses Long-Format als zweites Sheet in `sanitized-data.xlsx` speichern, z. B. „facts_long“.)

## Dashboard-Anforderungen (dashboard.html)
- Layout: modern, clean, responsive (Desktop first).
- Farbpalette: konsistent; jede Abteilung hat eine feste Farbe (Palette im Code definieren).
- Interaktivität:
  - Filterleiste: Stadt, Bundesland, Abteilung, Beruf, Teilzeit
  - Altersbereich via Slider (min/max)
  - Zeitraum-Filter (Start/Ende Monat) + Quick-Button „Letzte 12 Monate“
  - Cross-Filtering: Klick auf Abteilung/Beruf im Chart filtert alle anderen Visuals.
  - Tooltips: zeigen Details (Umsatz, Anteil, Headcount, Ø Umsatz etc.)

### KPIs (Executive Row)
- Gesamtumsatz (im aktuellen Filter/Zeitraum)
- Ø Monatsumsatz pro Person
- Umsatz nach Abteilung (Top-Abteilung + Anteil)
- Ø Alter (optional zusätzlich Median)
- Teilzeitquote + Umsatzvergleich Teilzeit vs. Vollzeit

### Charts (Storyline: Überblick → Treiber → Zeit)
1. Umsatzentwicklung über die Zeit (Line/Area)
2. Umsatz nach Abteilung (Bar oder Stacked)
3. Top 10 Berufe nach Umsatz (Bar, Tooltip mit Headcount + Ø Umsatz)
4. Umsatzverteilung (Histogramm oder Boxplot; sichtbar machen: Streuung & Ausreißer)
Optional:
5. Heatmap Monat × Abteilung (zeigt Peaks/Saisonalität)

## Demo-Qualität / Polishing
- Saubere Achsenbeschriftungen, Tausendertrennzeichen, EUR-Format.
- Sinnvolle Default-Ansicht (z. B. letztes volles Jahr + alle Abteilungen).
- Keine leeren Panels; wenn Filter alles rausfiltert: freundliche „No data“-State.
- Am Ende eine kurze Zusammenfassung im Report (1–2 Sätze), was aktuell gefiltert wird.

## Ergebnis liefern
- Speichere beide Artefakte in den Outputs:
  - sanitized-data.xlsx
  - dashboard.html
- Gib zusätzlich eine kurze Text-Zusammenfassung:
  - Anzahl Datensätze, erkannte Umsatzmonate (von/bis), Anzahl Warnungen (z. B. unbekannte Städte für Bundesland), und welche Default-Filter gesetzt sind.
```

## Skill installieren und testen

1. **Skill installieren**: Lade den "Sheet-to-Dashboard" Skill herunter und installiere ihn in deiner Agenten-Umgebung. Hier ein Beispiel für GitHub Copilot CLI:

```bash
/skills add sheet2dashboard
```

### Testen

Sobald der Skill installiert ist, kannst du ihn mit der bereitgestellten Excel-Datei testen:

```prompt
Nimm die Datei test-data.xlsx und erstelle daraus eine bereinigte Excel-Datei sowie ein interaktives HTML-Dashboard
   mit KPIs, Charts und Filtern.
```
