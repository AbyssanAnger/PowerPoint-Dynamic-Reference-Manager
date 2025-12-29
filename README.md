# PowerPoint VBA Reference Manager - Anleitung

## Installation

1. **PowerPoint öffnen** und Ihre Präsentation laden
2. **VBA-Editor öffnen**: Drücken Sie `Alt + F11`
3. **Modul importieren**:
   - `Datei` → `Datei importieren...`
   - Wählen Sie `ReferenceManager.bas`
   - Klicken Sie auf `Öffnen`

## Verwendung

### Schritt 1: Quellen laden

1. Wenn Sie das Makro starten, öffnet sich automatisch ein Datei-Dialog.
2. Wählen Sie Ihre `.bib` Datei aus (z.B. aus Zotero/Citavi exportiert).
3. Alternativ können Sie die mitgelieferte `test_references.bib` zum Testen verwenden.

> **Hinweis**: Falls Sie keine Datei auswählen, bleibt die Literaturliste leer.

### Schritt 2: Zitationen in Folien einfügen

Fügen Sie in Ihren PowerPoint-Folien Zitationen im Format `[Key]` ein, wobei `Key` dem Schlüssel in Ihrer BibTeX-Datei entspricht:

**Beispiel:**
```
"Laut [Smith2020] ist Machine Learning..."
```

### Schritt 3: Makro ausführen

1. Drücken Sie `Alt + F8`
2. Wählen Sie `UpdateReferences`
3. Klicken Sie auf `Ausführen`

### Ergebnis

- Alle Zitationen werden durch Nummern ersetzt: `[1]`, `[2]`, etc.
- Eine Literaturverzeichnis-Folie wird am Ende erstellt
- Gleiche Quellen erhalten die gleiche Nummer

## Beispiel

**Vorher:**
```
Folie 1: "Machine Learning [Smith2020] ist wichtig"
Folie 2: "Weitere Studien [Jones2021] und [Smith2020] zeigen..."
```

**Nachher:**
```
Folie 1: "Machine Learning [1] ist wichtig"
Folie 2: "Weitere Studien [2] und [1] zeigen..."

Literaturverzeichnis-Folie:
[1] Smith, J. et al. (2020). Machine Learning Fundamentals. IEEE Transactions on Neural Networks.
[2] Jones, A. & Brown, B. (2021). Advanced AI Techniques in Modern Computing. Nature Machine Intelligence.
```

## Anpassungen

### Zitationsformat ändern

In der Funktion `ProcessTextContent()` können Sie das Regex-Pattern anpassen:

```vba
' Aktuell: [SourceID]
regex.Pattern = "\[([A-Za-z0-9]+)\]"

' Alternative: (SourceID)
regex.Pattern = "\(([A-Za-z0-9]+)\)"
```

### Externe Quelldatei verwenden

Sie können die Quellen auch aus einer CSV-Datei laden. Beispiel-Code:

```vba
Private Sub LoadSourceRegistryFromFile()
    Dim filePath As String
    Dim fileNum As Integer
    Dim line As String
    Dim parts() As String
    
    filePath = "C:\Pfad\zu\sources.csv"
    fileNum = FreeFile
    
    Open filePath For Input As fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, line
        parts = Split(line, "|")
        
        If UBound(parts) >= 4 Then
            Call AddSource(parts(0), parts(1), parts(2), parts(3), parts(4))
        End If
    Loop
    Close fileNum
End Sub
```

## Fehlerbehebung

**Problem**: "Unknown citation" Warnung im Debug-Fenster  
**Lösung**: Fügen Sie die fehlende Quelle in `LoadSourceRegistry()` hinzu

**Problem**: Makro funktioniert nicht  
**Lösung**: Stellen Sie sicher, dass Makros aktiviert sind (`Datei` → `Optionen` → `Trust Center`)

**Problem**: Formatierung geht verloren  
**Lösung**: Das Makro ersetzt nur den Text, Formatierung sollte erhalten bleiben. Falls nicht, kontaktieren Sie den Support.
