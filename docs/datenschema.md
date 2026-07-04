# Datenschema

## Termine-CSV

Erwartete Spalten:

- `project_match_id`: numerische Projekt-ID fuer den Merge
- `start`: Datum oder Zeitstempel des Termins
- `localized_type`: Terminart, relevant sind DC-Montage mit oder ohne Geruest
- `deleted`: Wahrheitswert, geloeschte Termine werden ignoriert

Weitere Spalten duerfen vorhanden sein und werden ignoriert.

## Projekte-CSV

Erwartete Spalten:

- `id`: Projekt-ID, passend zu `termine.project_match_id`
- `project_nr`: Projektnummer fuer die Excel-Tabelle
- `contact_id`: Kontakt-ID fuer den Merge mit der Kontakte-CSV

## Kontakte-CSV

Erwartete Spalten:

- `id`: Kontakt-ID, passend zu `projekte.contact_id`
- `first_name`: Vorname oder synthetischer Demo-Wert
- `last_name`: Nachname oder synthetischer Demo-Wert

## Excel-Tabelle

Standardmaessig wird die Tabelle `Tabelle1` im aktiven Arbeitsblatt genutzt. Die
Sync-Logik schreibt in die ersten vier Spalten:

- `project_nr`
- `first_name`
- `last_name`
- `start_datum`

Die sechste Tabellenspalte wird als `Einspeiseranfrage`-Status interpretiert.
Wenn eine Projektnummer mehrfach vorhanden ist und mindestens eine Zeile dort
einen Wert hat, werden leere Dubletten dieser Projektnummer entfernt.
