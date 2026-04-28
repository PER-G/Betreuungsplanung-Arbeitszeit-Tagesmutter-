# Betreuungsplanung Niklas Tom-Hardy

Tools zur Planung der Wochenarbeitszeit von Paul und Dominique inklusive Tagesmutter-Kosten und Niklas-Betreuung. Enthält 12 vorberechnete Varianten als Excel-Datei sowie eine interaktive HTML-Seite mit Live-Berechnung.

## Inhalt

| Datei | Beschreibung |
|---|---|
| `Betreuungsplanung_Niklas.html` | **Interaktive Web-App** mit Slidern für Gehalt, Stunden, Modus und Tagesmutter-Zeiten. Berechnet Wochenplan, Brutto/Netto-Einkommen und TM-Kosten live. Offline im Browser nutzbar. |
| `Betreuungsplanung_Niklas_V0.04.xlsx` | Aktuelle Excel-Datei mit 12 Varianten + Vergleichs-Sheet. |
| `build_betreuung.py` | Python-Skript zur Generierung der Excel-Datei (openpyxl). |
| `Betreuungsplanung_Niklas_V0.0X.xlsx` | Frühere Versionen (V0.01–V0.03). |

## Varianten in der Excel-Datei

- **V1–V5**: ohne Tagesmutter, verschiedene Stundenmodelle
- **V6–V9**: mit Tagesmutter (5/3/2 Tage), Standard 07:30–13:00
- **V10–V12**: Paul 35h auf 4 Tage, Tagesmutter mit individuellen Zeitfenstern

Jede Variante zeigt:
- Wochenplan im 30-Min-Raster mit Niklas-Kontrollspalte (grün = betreut, rot = Lücke)
- Brutto/Netto-Einkommen pro Person und Haushalt
- TM-Kosten gem. Anlage 1a Landkreis Tübingen ('unter 3 Jahren', gültig 01.01.2026 – 31.08.2026)

## Annahmen

- Paul: Vollzeit 96.000 €/J, 30 Min Anfahrt mit Auto
- Dominique: Vollzeit 84.000 €/J, 15 Min Anfahrt zu Fuß
- Niklas: unter 3 Jahre
- Steuerklasse IV/IV, verheiratet, 1 Kind, 2026 (Netto-Schätzung)
- Pause: 0,5h unbezahlt ab >6h Tagesarbeitszeit
- Homeoffice: ca. 50 % Eigen-Betreuung möglich

## HTML-App nutzen

Datei `Betreuungsplanung_Niklas.html` mit Doppelklick im Browser öffnen. Keine Installation nötig.

## Excel neu generieren

```bash
python build_betreuung.py
```

Erzeugt `Betreuungsplanung_Niklas_V0.04.xlsx`. Benötigt `openpyxl`.
